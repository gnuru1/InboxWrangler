"""
Analyze Reports Script for Hybrid Outlook Organizer

This script analyzes the output of the organizer's reports to check if unread status
is properly affecting scores and folder assignments.
"""

import os
import sys
import argparse
import logging
import re
import json
import datetime
import pandas as pd
import win32com.client
import pywintypes
from bs4 import BeautifulSoup
from pathlib import Path
from collections import defaultdict

from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox
from config import load_config

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

def load_report_from_html(html_path):
    """Load email data from an HTML report file."""
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    
    # Find the table
    table = soup.find('table')
    if not table:
        raise ValueError("Could not find table in HTML report")
    
    # Extract rows
    rows = table.find_all('tr')
    if len(rows) <= 1:
        raise ValueError("Table has no data rows")
    
    # Extract headers
    headers = [header.text.strip() for header in rows[0].find_all('th')]
    
    # Extract data
    data = []
    for row in rows[1:]:  # Skip header row
        cells = row.find_all('td')
        if len(cells) == len(headers):
            row_data = {}
            for i, cell in enumerate(cells):
                row_data[headers[i]] = cell.text.strip()
            data.append(row_data)
    
    return data

def load_report_from_csv(csv_path):
    """Load email data from a CSV report file."""
    try:
        df = pd.read_csv(csv_path)
        return df.to_dict('records')
    except Exception as e:
        logger.error(f"Error loading CSV report: {e}")
        return []

def get_outlook_item_details(inbox, subject, sender=None):
    """Try to find an Outlook item by subject and return its details."""
    try:
        # Get all items in inbox
        items = inbox.Items
        
        # Clean the subject for more reliable matching by removing trailing spaces
        subject = subject.strip()
        
        # Create a filter by subject - use a more relaxed filter to find potential matches
        # Use a more relaxed filter that might catch more potential matches
        subject_filter = f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{subject}'"
        filtered_items = items.Restrict(subject_filter)
        
        result = {}
        for item in filtered_items:
            # Only check mail items
            if getattr(item, 'Class', 0) != 43:  # 43 = olMail
                continue
                
            item_subject = getattr(item, 'Subject', '').strip()
            item_sender = getattr(item, 'SenderName', 'Unknown')
            item_sender_email = getattr(item, 'SenderEmailAddress', '').lower()
            
            # Log potential matches for debugging
            logger.debug(f"Checking match: '{item_subject}' vs '{subject}' from {item_sender}")
            
            # More flexible matching:
            # 1. Subject comparison - ignore trailing spaces, case insensitive
            subject_matches = (
                item_subject.lower() == subject.lower() or
                item_subject.lower().startswith(subject.lower()) or
                subject.lower().startswith(item_subject.lower())
            )
            
            # 2. Sender matching - be more flexible
            sender_matches = (
                sender is None or 
                sender.lower() in item_sender.lower() or 
                item_sender.lower() in sender.lower() or
                (item_sender_email and sender.lower() in item_sender_email)
            )
            
            if subject_matches and sender_matches:
                # We found a likely match
                result = {
                    'EntryID': getattr(item, 'EntryID', ''),
                    'Subject': item_subject,
                    'SenderName': item_sender,
                    'SenderEmailAddress': getattr(item, 'SenderEmailAddress', ''),
                    'Unread': getattr(item, 'Unread', False),  # Critical for our analysis
                    'ReceivedTime': getattr(item, 'ReceivedTime', None),
                    'Importance': getattr(item, 'Importance', 1),
                    'FlagStatus': getattr(item, 'FlagStatus', 0),
                    'CheckCount': 0  # We'll set this later from email_tracking
                }
                
                # Print debug info
                logger.debug(f"Found item: '{item_subject}' from {item_sender}, Unread: {result['Unread']}")
                break
                
        return result
    except pywintypes.com_error as ce:
        logger.debug(f"COM error while searching for item: {ce}")
        return {}
    except Exception as e:
        logger.debug(f"Error retrieving Outlook item: {e}")
        return {}

def scan_inbox_status(inbox, limit=100):
    """Scan the inbox directly to count read vs unread messages."""
    read_count = 0
    unread_count = 0
    
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # Newest first
        total = items.Count
        count_limit = min(total, limit)
        
        logger.info(f"Scanning inbox directly: {count_limit} of {total} items...")
        
        # Track by sender for additional analysis
        sender_stats = defaultdict(lambda: {'read': 0, 'unread': 0})
        
        for i in range(1, count_limit + 1):
            try:
                item = items.Item(i)
                if getattr(item, 'Class', 0) != 43:  # Only count mail items
                    continue
                    
                is_unread = getattr(item, 'Unread', False)
                sender = getattr(item, 'SenderName', 'Unknown')
                subject = getattr(item, 'Subject', 'Unknown')
                
                if is_unread:
                    unread_count += 1
                    sender_stats[sender]['unread'] += 1
                else:
                    read_count += 1
                    sender_stats[sender]['read'] += 1
                    logger.debug(f"Found READ email: {subject} from {sender}")
                    
            except Exception as e:
                logger.debug(f"Error processing inbox item {i}: {e}")
                
        # Print senders with both read and unread
        print("\n--- SENDERS WITH MIXED READ/UNREAD STATUS ---")
        mixed_senders = [(s, stats) for s, stats in sender_stats.items() 
                          if stats['read'] > 0 and stats['unread'] > 0]
        
        for sender, stats in sorted(mixed_senders, key=lambda x: -(x[1]['read'] + x[1]['unread'])):
            print(f"{sender}: {stats['read']} read, {stats['unread']} unread")
                
        return {'read': read_count, 'unread': unread_count, 'sender_stats': dict(sender_stats)}
            
    except Exception as e:
        logger.error(f"Error scanning inbox: {e}")
        return {'read': 0, 'unread': 0, 'sender_stats': {}}

def analyze_report(report_data, inbox, config, email_tracking=None):
    """Analyze the report data for proper unread status handling."""
    if not report_data:
        logger.error("No report data to analyze")
        return {}
        
    # First scan inbox directly to get read/unread counts
    inbox_scan = scan_inbox_status(inbox, limit=200)
    
    results = {
        'total_emails': len(report_data),
        'unread_count': 0,
        'read_count': 0,
        'inbox_scan': inbox_scan,  # Include direct inbox scan results
        'unread_scores': [],
        'read_scores': [],
        'unread_avg_score': 0,
        'read_avg_score': 0,
        'unread_by_folder': defaultdict(int),
        'read_by_folder': defaultdict(int),
        'score_differential': 0,
        'highest_unread': {'score': 0, 'subject': '', 'sender': ''},
        'lowest_unread': {'score': 1.0, 'subject': '', 'sender': ''},
        'items_checked': 0,
        'items_found': 0,
        'problems': [],
        'not_found': []  # Track items not found in Outlook
    }
    
    # Extract the message state weight and unread penalty from config
    message_state_weight = config.get('message_state_weight', 0.1)
    unread_penalty = config.get('unread_penalty', 0.2)
    read_kept_bonus = config.get('read_kept_bonus', 0.3)
    ignore_penalty = config.get('ignore_penalty', 0.15)
    expected_unread_impact = message_state_weight * unread_penalty  # Should be negative now
    expected_read_impact = message_state_weight * read_kept_bonus  # Positive
    
    logger.info(f"Analyzing {len(report_data)} emails from report...")
    
    # Loop through each email in the report
    for i, email in enumerate(report_data):
        if i % 10 == 0 and i > 0:
            logger.info(f"Processed {i}/{len(report_data)} emails...")
        
        # Extract data from report
        try:
            subject = email.get('Subject', '')
            sender = email.get('Sender', '')
            score = float(email.get('Score', 0))
            folder = email.get('Recommended Folder', '')
            
            # Log what we're checking
            logger.debug(f"Checking email: '{subject}' from {sender}")
            
            # Connect to Outlook and get the actual item's unread status
            outlook_details = get_outlook_item_details(inbox, subject, sender)
            results['items_checked'] += 1
            
            if outlook_details:
                results['items_found'] += 1
                entry_id = outlook_details.get('EntryID', '')
                unread = outlook_details.get('Unread', False)
                
                # Add more verbose debug info
                logger.debug(f"Found match in Outlook - Subject: '{outlook_details.get('Subject', '')}', " +
                          f"Sender: {outlook_details.get('SenderName', '')}, Unread: {unread}")
                
                # Get check count from email_tracking if available
                check_count = 1
                if email_tracking and entry_id and entry_id in email_tracking:
                    track_record = email_tracking[entry_id]
                    check_count = track_record.get('check_count', 1)
                    outlook_details['CheckCount'] = check_count
                
                # Track by read status
                if unread:
                    results['unread_count'] += 1
                    results['unread_scores'].append(score)
                    results['unread_by_folder'][folder] += 1
                    
                    # Track highest/lowest unread
                    if score > results['highest_unread']['score']:
                        results['highest_unread'] = {
                            'score': score,
                            'subject': subject,
                            'sender': sender,
                            'check_count': check_count
                        }
                    if score < results['lowest_unread']['score']:
                        results['lowest_unread'] = {
                            'score': score,
                            'subject': subject,
                            'sender': sender,
                            'check_count': check_count
                        }
                else:
                    results['read_count'] += 1
                    results['read_scores'].append(score)
                    results['read_by_folder'][folder] += 1
                    
                    # Track this read email for later analysis
                    logger.debug(f"Counted as READ: '{subject}' from {sender}, Score: {score}, Folder: {folder}")
                    
                    # Problem detection: Read email kept in inbox but still has low score
                    if score < 0.5:  # Read emails should have higher scores with the new logic
                        results['problems'].append({
                            'type': 'read_kept_low_score',
                            'subject': subject,
                            'sender': sender,
                            'score': score,
                            'folder': folder
                        })
                
                # Problem detection: Unread item with high score (with new logic, unread should have lower scores)
                if unread and score > 0.5:  # Unread items should have lower scores now
                    results['problems'].append({
                        'type': 'high_unread_score',
                        'subject': subject,
                        'sender': sender,
                        'score': score,
                        'folder': folder,
                        'check_count': check_count
                    })
                
                # Problem detection: Repeatedly ignored (high check_count) but still high score
                if unread and check_count > 2 and score > 0.4:
                    # Calculate expected penalty: message_state_weight * ignore_penalty * (check_count-1)
                    # For example: 0.1 * 0.3 * 2 = 0.06 for check_count=3
                    expected_penalty = message_state_weight * ignore_penalty * (check_count-1)
                    max_expected_score = 0.6 - expected_penalty  # 0.6 is a reasonable baseline for an important email
                    
                    if score > max_expected_score:
                        results['problems'].append({
                            'type': 'ignored_email_high_score',
                            'subject': subject,
                            'sender': sender,
                            'score': score,
                            'check_count': check_count,
                            'expected_penalty': expected_penalty,
                            'max_expected': max_expected_score,
                            'folder': folder
                        })
                    
                # Problem detection: Very old unread item with high score
                if unread and outlook_details.get('ReceivedTime'):
                    rec_time = outlook_details['ReceivedTime']
                    if isinstance(rec_time, (datetime.datetime, pywintypes.TimeType)):
                        now = datetime.datetime.now()
                        try:
                            days_old = (now - rec_time).days
                            if days_old > 30 and score > 0.6:  # Old email with high score
                                results['problems'].append({
                                    'type': 'old_unread_high_score',
                                    'subject': subject,
                                    'sender': sender,
                                    'score': score,
                                    'days_old': days_old,
                                    'check_count': check_count,
                                    'folder': folder
                                })
                        except TypeError:
                            pass  # Skip if datetime comparison fails
            else:
                # Track items not found in Outlook
                results['not_found'].append({
                    'subject': subject,
                    'sender': sender,
                    'score': score,
                    'folder': folder
                })
                logger.debug(f"Could not find email in Outlook: '{subject}' from {sender}")
                
        except Exception as e:
            logger.debug(f"Error processing email {subject}: {e}")
            continue
    
    # Calculate averages and differentials
    if results['unread_scores']:
        results['unread_avg_score'] = sum(results['unread_scores']) / len(results['unread_scores'])
    if results['read_scores']:
        results['read_avg_score'] = sum(results['read_scores']) / len(results['read_scores'])
    
    # Calculate the average score differential between unread and read
    results['score_differential'] = results['unread_avg_score'] - results['read_avg_score']
    
    # Analyze if the unread penalty is having the expected effect
    if results['score_differential'] > expected_unread_impact:
        results['problems'].append({
            'type': 'insufficient_unread_impact',
            'details': f"Expected unread impact: {expected_unread_impact:.4f}, Actual: {results['score_differential']:.4f}"
        })
    
    # Generate a summary 
    results['status'] = 'ok' if not results['problems'] else 'issues_detected'
    
    # Log summary of read/unread count
    logger.info(f"Analysis found {results['read_count']} read and {results['unread_count']} unread emails out of {results['items_found']} items found in Outlook")
    logger.info(f"Direct inbox scan found {inbox_scan['read']} read and {inbox_scan['unread']} unread emails")
    
    return results

def print_report_analysis(analysis):
    """Print the analysis results in a readable format."""
    print("\n========== EMAIL REPORT ANALYSIS ==========\n")
    
    # Direct inbox scan results
    if 'inbox_scan' in analysis:
        print("--- DIRECT INBOX SCAN ---")
        scan = analysis['inbox_scan']
        print(f"Direct inbox read count: {scan['read']}")
        print(f"Direct inbox unread count: {scan['unread']}")
        print(f"Total scanned directly: {scan['read'] + scan['unread']}")
        if scan['read'] + scan['unread'] > 0:
            read_pct = (scan['read'] / (scan['read'] + scan['unread'])) * 100
            print(f"Read percentage in inbox: {read_pct:.1f}%")
        print()
    
    print(f"Total Emails in Report: {analysis['total_emails']}")
    print(f"Items checked in Outlook: {analysis['items_checked']}")
    print(f"Items found in Outlook: {analysis['items_found']}")
    print(f"Items not found in Outlook: {len(analysis.get('not_found', []))}")
    print(f"Found unread count: {analysis['unread_count']}")
    print(f"Found read count: {analysis['read_count']}")
    
    if analysis['read_count'] > 0:
        print("\n--- READ EMAILS FOUND ---")
        for folder, count in sorted(analysis['read_by_folder'].items(), key=lambda x: -x[1]):
            print(f"  {folder}: {count}")
    
    print("\n--- SCORING ANALYSIS ---")
    print(f"Average score for unread emails: {analysis['unread_avg_score']:.4f}")
    print(f"Average score for read emails: {analysis['read_avg_score']:.4f}")
    print(f"Score differential (unread - read): {analysis['score_differential']:.4f}")
    
    print("\n--- UNREAD BY FOLDER ---")
    for folder, count in sorted(analysis['unread_by_folder'].items(), key=lambda x: -x[1]):
        print(f"{folder}: {count}")
    
    print("\n--- READ BY FOLDER ---")
    if analysis['read_by_folder']:
        for folder, count in sorted(analysis['read_by_folder'].items(), key=lambda x: -x[1]):
            print(f"{folder}: {count}")
    else:
        print("No read emails found in report")
    
    print("\n--- HIGHEST SCORED UNREAD EMAIL ---")
    highest = analysis['highest_unread']
    if highest['score'] > 0:
        print(f"Score: {highest['score']:.4f}")
        print(f"Subject: {highest['subject']}")
        print(f"Sender: {highest['sender']}")
        print(f"Check count: {highest.get('check_count', 1)}")
    else:
        print("No unread emails found")
    
    print("\n--- LOWEST SCORED UNREAD EMAIL ---")
    lowest = analysis['lowest_unread']
    if lowest['score'] < 1.0:
        print(f"Score: {lowest['score']:.4f}")
        print(f"Subject: {lowest['subject']}")
        print(f"Sender: {lowest['sender']}")
        print(f"Check count: {lowest.get('check_count', 1)}")
    else:
        print("No unread emails found")
    
    # Show items not found in Outlook
    not_found = analysis.get('not_found', [])
    if not_found:
        print(f"\n--- EMAILS NOT FOUND IN OUTLOOK ({len(not_found)}) ---")
        for i, item in enumerate(not_found[:5], 1):  # Show first 5
            print(f"{i}. '{item['subject']}' from {item['sender']}")
        if len(not_found) > 5:
            print(f"   ... and {len(not_found) - 5} more emails not found")
    
    # Group problems by type
    problem_types = defaultdict(list)
    for problem in analysis.get('problems', []):
        problem_types[problem['type']].append(problem)
    
    print("\n--- PROBLEMS DETECTED ---")
    if problem_types:
        print(f"Total problems: {len(analysis['problems'])}")
        
        # Print summary by type
        print("\nProblem summary by type:")
        for ptype, problems in problem_types.items():
            print(f"- {ptype}: {len(problems)} issues")
        
        # Print details of each problem
        print("\nDetailed problems:")
        i = 1
        for ptype, problems in problem_types.items():
            print(f"\n{ptype.upper().replace('_', ' ')} ({len(problems)} issues):")
            
            for problem in problems:
                print(f"{i}. Subject: {problem.get('subject', 'N/A')}")
                print(f"   Sender: {problem.get('sender', 'N/A')}")
                print(f"   Score: {problem.get('score', 'N/A')}")
                
                if 'check_count' in problem:
                    print(f"   Check count: {problem['check_count']}")
                if 'days_old' in problem:
                    print(f"   Days old: {problem['days_old']}")
                if 'expected_penalty' in problem:
                    print(f"   Expected penalty: {problem['expected_penalty']:.4f}")
                    print(f"   Maximum expected score: {problem['max_expected']:.4f}")
                if 'folder' in problem:
                    print(f"   Folder: {problem['folder']}")
                    
                print()
                i += 1
                
        # Print recommendations based on problems
        print("\n--- RECOMMENDATIONS ---")
        
        if 'insufficient_unread_impact' in problem_types:
            print("1. INCREASE UNREAD IMPACT:")
            print("   - Change 'message_state_weight' from 0.1 to 0.25 in config.json")
            print("   - Change 'unread_penalty' from 0.2 to 0.4 in config.json")
            
        if 'ignored_email_high_score' in problem_types:
            print("2. INCREASE IGNORE PENALTY:")
            print("   - Change 'ignore_penalty' from 0.15 to 0.3 in config.json")
            print("   - This will more severely penalize repeatedly ignored emails")
            
        if 'high_unread_score' in problem_types:
            print("3. ADJUST BASE SCORES:")
            print("   - Reduce content-based weights in the config")
            print("   - This will prevent unread emails from scoring too high")
            print("   - Consider adding an 'age_penalty_factor' for old unread emails")
            
        if 'read_kept_low_score' in problem_types:
            print("4. INCREASE READ KEPT BONUS:")
            print("   - Change 'read_kept_bonus' from 0.3 to 0.5 in config.json")
            print("   - This will boost emails you've read but kept in your inbox")
            print("   - Consider implementing a 'days_kept_factor' that increases score based on how many days a read email has been kept")
            
        if 'old_unread_high_score' in problem_types:
            print("5. ADD AGE PENALTY FOR OLD UNREAD EMAILS:")
            print("   - Add 'age_penalty_factor: 0.005' to config.json")
            print("   - This will reduce scores of old unread emails by 0.005 per day old")

        # General advice about read kept emails
        print("\n6. ADVICE FOR PRIORITIZING READ KEPT EMAILS:")
        print("   - Emails that you've read but kept in the inbox are likely important")
        print("   - If these emails are getting low scores, increase the 'read_kept_bonus' parameter") 
        print("   - You might also want to add a special folder named 'Important' or 'Follow-up'")
        print("   - Consider adding a rule that automatically flags emails you've replied to")
            
    else:
        print("No problems detected! Unread handling appears to be working correctly.")
    
    print("\n============================================\n")

def save_analysis_to_json(analysis, output_path):
    """Save the analysis results to a JSON file."""
    # Convert defaultdicts to regular dicts for JSON serialization
    analysis_copy = {**analysis}
    analysis_copy['unread_by_folder'] = dict(analysis['unread_by_folder'])
    analysis_copy['read_by_folder'] = dict(analysis['read_by_folder'])
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(analysis_copy, f, indent=2, default=str)
    
    logger.info(f"Analysis saved to {output_path}")

def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(description="Analyze report output for unread handling.")
    parser.add_argument("report_file", type=str, help="Path to the HTML or CSV report file")
    parser.add_argument("--output", type=str, default=None, help="Path to save the analysis JSON (optional)")
    parser.add_argument("--config", type=str, default="./config.json", help="Path to configuration file")
    parser.add_argument("--data-dir", type=str, default="./email_data", help="Path to the data directory containing tracking files")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    
    args = parser.parse_args()
    
    # Set up debug logging if requested
    if args.debug:
        logger.setLevel(logging.DEBUG)
        for handler in logger.handlers:
            handler.setLevel(logging.DEBUG)
    
    # Check if report file exists
    report_path = Path(args.report_file)
    if not report_path.exists():
        logger.error(f"Report file not found: {report_path}")
        return 1
    
    # Load configuration
    try:
        config = load_config(args.config)
        logger.info(f"Loaded configuration from {args.config}")
    except Exception as e:
        logger.error(f"Error loading configuration: {e}")
        return 1
    
    # Load email tracking data if available
    email_tracking = None
    data_dir = Path(args.data_dir)
    tracking_file = data_dir / 'email_tracking.pkl'
    if tracking_file.exists():
        try:
            import pickle
            with open(tracking_file, 'rb') as f:
                email_tracking = pickle.load(f)
            logger.info(f"Loaded email tracking data from {tracking_file}")
        except Exception as e:
            logger.warning(f"Could not load email tracking data: {e}")
    
    # Connect to Outlook
    try:
        outlook, namespace = connect_to_outlook()
        if not namespace:
            logger.error("Failed to connect to Outlook. Exiting.")
            return 1
        
        inbox = get_default_folder(namespace, olFolderInbox)
        if not inbox:
            logger.error("Could not retrieve Inbox folder. Exiting.")
            return 1
    except Exception as e:
        logger.error(f"Error connecting to Outlook: {e}")
        return 1
    
    # Load report data based on file extension
    try:
        file_ext = report_path.suffix.lower()
        if file_ext == '.html':
            report_data = load_report_from_html(report_path)
        elif file_ext == '.csv':
            report_data = load_report_from_csv(report_path)
        else:
            logger.error(f"Unsupported report file format: {file_ext}. Use .html or .csv")
            return 1
            
        logger.info(f"Loaded report with {len(report_data)} emails")
    except Exception as e:
        logger.error(f"Error loading report: {e}")
        return 1
    
    # Analyze the report
    analysis_results = analyze_report(report_data, inbox, config, email_tracking)
    
    # Print results
    print_report_analysis(analysis_results)
    
    # Save to JSON if output path provided
    if args.output:
        output_path = Path(args.output)
        save_analysis_to_json(analysis_results, output_path)
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 