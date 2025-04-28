"""
Inbox Sender Statistics Utility

Analyzes the current state of the Outlook Inbox to provide statistics
about senders, focusing on volume, read/unread ratios, and subject line
repetitiveness (using fuzzy matching) to help identify potential cleanup targets.

Usage:
    python inbox_sender_stats.py --limit 1000 --top-senders 20 [--fuzzy-threshold 85] [--data-dir path/to/data] [--debug]
"""

import argparse
import sys
import logging
import datetime
import pywintypes
from pathlib import Path
from collections import Counter, defaultdict
import re

try:
    from thefuzz import fuzz
except ImportError:
    print("Error: 'thefuzz' library not found. Please install it using:")
    print("pip install thefuzz python-Levenshtein")
    sys.exit(1)

# Project modules
from config import load_config # Although config isn't directly used, needed for data_dir logic maybe
from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox, safe_get_property
from analyzer import EmailAnalyzer # Needed for contact normalization

# Basic logger setup
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('sender_stats')

def setup_logging(level=logging.INFO):
    """Configure logging settings."""
    logging.getLogger().setLevel(level)
    logger.setLevel(level)
    if not any(isinstance(h, logging.StreamHandler) for h in logging.getLogger().handlers):
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logging.getLogger().addHandler(handler)
    logger.info(f"Logging setup complete. Level set to: {logging.getLevelName(level)}")

def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description="Analyze Inbox sender statistics.")
    parser.add_argument("--limit", type=int, required=True,
                        help="Number of the most recent emails from the Inbox to analyze.")
    parser.add_argument("--top-senders", type=int, default=20,
                        help="Number of top senders to display in the report (default: 20).")
    parser.add_argument("--fuzzy-threshold", type=int, default=85,
                        help="Similarity threshold (0-100) for grouping subjects (default: 85).")
    parser.add_argument("--data-dir", type=str, default="./email_data",
                        help="Directory containing analysis data (needed for contact map) (default: ./email_data).")
    parser.add_argument("--config", type=str, default="./config.json",
                        help="Path to configuration file (default: ./config.json). Used mainly if data_dir isn't specified.")
    parser.add_argument("--debug", action="store_true",
                        help="Enable debug logging.")
    args = parser.parse_args()
    if not 0 <= args.fuzzy_threshold <= 100:
        parser.error("Fuzzy threshold must be between 0 and 100.")
    return args

def clean_subject(subject):
    """Rudimentary cleaning of subject line for comparison."""
    if not subject:
        return ""
    # Remove RE:, FWD:, etc.
    cleaned = re.sub(r'^(re|fw|fwd|aw|wg):\s*', '', subject.strip(), flags=re.IGNORECASE)
    # Optional: Lowercase and remove extra whitespace
    cleaned = cleaned.lower()
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    # Remove potential variations in dates/numbers that might throw off fuzzy match
    cleaned = re.sub(r'\b\d+[/-]\d+([/-]\d+)?\b', ' ', cleaned) # Remove dates like MM/DD/YY
    cleaned = re.sub(r'\b\d+\b', ' ', cleaned) # Remove standalone numbers
    cleaned = re.sub(r'\s+', ' ', cleaned).strip() # Clean whitespace again
    return cleaned

def calculate_fuzzy_subject_similarity(subjects, threshold=85):
    """Calculates subject similarity based on fuzzy matching clusters.
       Returns the percentage of emails belonging to the largest cluster.
    """
    if not subjects:
        return 0.0

    total_count = len(subjects)
    if total_count == 0:
         return 0.0

    cleaned_subjects = [clean_subject(s) for s in subjects if clean_subject(s)] # Clean and remove empty
    if not cleaned_subjects:
         return 0.0 # Return 0 if all subjects became empty after cleaning

    clusters = [] # Stores tuples: (representative_subject, count)

    for subject in cleaned_subjects:
        matched = False
        best_match_idx = -1
        highest_ratio = -1

        # Find the best matching existing cluster above threshold
        for i in range(len(clusters)):
            representative, _ = clusters[i]
            # Consider using partial_ratio if subjects contain garbage, but ratio is generally better
            ratio = fuzz.ratio(subject, representative)
            if ratio >= threshold and ratio > highest_ratio:
                highest_ratio = ratio
                best_match_idx = i
                matched = True # Found a potential match

        if matched:
            # Add to the best matching cluster
            rep, count = clusters[best_match_idx]
            clusters[best_match_idx] = (rep, count + 1)
        else:
            # Start a new cluster
            clusters.append((subject, 1))

    if not clusters:
        return 0.0

    # Find the size of the largest cluster
    max_count = 0
    if clusters:
        max_count = max(count for _, count in clusters)

    similarity_score = (max_count / total_count) * 100 # Use original total_count including potentially empty subjects
    return similarity_score

def main():
    """Main execution function"""
    args = parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(level=log_level)

    # Load config primarily to potentially find data_dir if not specified
    try:
        config = load_config(args.config)
        # Determine data_dir: command line overrides config
        data_dir = Path(args.data_dir if args.data_dir else config.get('data_dir', './email_data'))
        logger.info(f"Using data directory: {data_dir}")
    except Exception as e:
        logger.error(f"Error loading configuration: {e}")
        print(f"Error: Could not load config file {args.config}. Exiting.", file=sys.stderr)
        return 1

    # --- Connect to Outlook & Get Inbox --- 
    outlook, namespace = connect_to_outlook()
    if not namespace:
        print("Error: Failed to connect to Outlook. Exiting.", file=sys.stderr)
        return 1

    inbox = get_default_folder(namespace, olFolderInbox)
    if not inbox:
        print("Error: Could not retrieve Inbox folder. Exiting.", file=sys.stderr)
        return 1

    # --- Load Contact Map via Analyzer --- 
    analyzer = None # Initialize
    contact_map = {}
    try:
        # Initialize Analyzer minimally to load the contact map
        # Suppress its internal logging unless we are debugging this script
        logging.getLogger('analyzer').setLevel(logging.WARNING if not args.debug else logging.DEBUG)
        analyzer = EmailAnalyzer(namespace, inbox, None, config, data_dir) # Pass None for sent_items
        contact_map = analyzer.contact_map # Use the loaded map
        logger.info(f"Loaded contact map with {len(contact_map)} mappings.")
    except Exception as e:
        logger.error(f"Could not initialize EmailAnalyzer to load contact map: {e}")
        logger.warning("Proceeding without contact normalization. Sender stats may be fragmented.")
        # Don't exit, just proceed without normalization

    # --- Process Inbox Items --- 
    logger.info(f"Fetching the latest {args.limit} emails from Inbox '{inbox.Name}'...")
    sender_data = defaultdict(lambda: {'subjects': [], 'read': 0, 'unread': 0, 'total': 0})
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True) # Sort newest first
        num_items_to_process = min(items.Count, args.limit)
        if num_items_to_process == 0:
             print("Inbox is empty or no items found.")
             return 0
             
        logger.info(f"Analyzing {num_items_to_process} items...")
    except pywintypes.com_error as ce:
        logger.error(f"COM Error accessing/sorting Inbox items: {ce}")
        print(f"Error: Could not access Inbox items. {ce}", file=sys.stderr)
        return 1
    except Exception as e:
        logger.error(f"Error accessing/sorting Inbox items: {e}", exc_info=True)
        print(f"Error: Could not access Inbox items. {e}", file=sys.stderr)
        return 1

    processed_count = 0
    errors_count = 0
    for i in range(num_items_to_process):
        item = None
        try:
            if (i + 1) > items.Count:
                logger.warning("Inbox item count changed during processing. Stopping.")
                break
                
            item = items.Item(i + 1)

            if safe_get_property(item, 'Class', default=-1) != 43: # olMail
                continue

            processed_count += 1
            
            # Get Sender and Normalize
            raw_sender = "Unknown Sender"
            sender_addr = safe_get_property(item, 'SenderEmailAddress')
            sender_name = safe_get_property(item, 'SenderName')
            if sender_addr and '@' in sender_addr:
                raw_sender = sender_addr
            elif sender_name:
                raw_sender = sender_name
            
            # Use analyzer's normalize method if available
            normalized_sender = raw_sender.lower()
            if analyzer:
                normalized_sender = analyzer._normalize_contact(raw_sender)
            else: # Basic normalization if analyzer failed
                 if '@' not in normalized_sender:
                      normalized_sender = normalized_sender.lower() # Lowercase display names
                      
            # Get Subject & Read Status
            subject = safe_get_property(item, 'Subject', default="")
            is_unread = safe_get_property(item, 'UnRead', default=True)
            
            # Store data
            stats = sender_data[normalized_sender]
            stats['subjects'].append(subject)
            stats['total'] += 1
            if is_unread:
                stats['unread'] += 1
            else:
                stats['read'] += 1
                
            if processed_count % 250 == 0:
                 logger.info(f"Processed {processed_count}/{num_items_to_process} emails...")

        except pywintypes.com_error as ce:
            errors_count += 1
            logger.debug(f"COM Error processing item index {i+1}: {ce}")
        except Exception as e:
            errors_count += 1
            subject = safe_get_property(item, 'Subject', 'N/A') if item else 'N/A'
            logger.debug(f"Error processing item {i+1} ('{subject}'): {e}", exc_info=True)
        finally:
            if item is not None:
                del item
                item = None
                
    logger.info(f"Finished processing {processed_count} emails with {errors_count} errors.")

    # --- Calculate Stats & Prepare Report --- 
    report_data = []
    for sender, data in sender_data.items():
        total = data['total']
        unread = data['unread']
        subjects = data['subjects']
        
        unread_percent = (unread / total * 100) if total > 0 else 0
        # Use the new fuzzy calculation
        subject_similarity = calculate_fuzzy_subject_similarity(subjects, args.fuzzy_threshold)
        
        report_data.append({
            'sender': sender,
            'total_emails': total,
            'unread_percent': unread_percent,
            'subject_similarity_percent': subject_similarity
        })
        
    # Sort by total emails descending
    report_data.sort(key=lambda x: x['total_emails'], reverse=True)
    
    # --- Print Report --- 
    # Simplified print statements to avoid potential f-string escape issues
    print(f"\n--- Top {min(args.top_senders, len(report_data))} Senders in Inbox (Analyzed {processed_count} emails, Fuzzy Threshold: {args.fuzzy_threshold}) ---")
    header_format = "{:<50} | {:>12} | {:>10} | {:>18}"
    header = header_format.format("Sender", "Total Emails", "% Unread", "% Fuzzy Subject")
    separator = "-" * 50 + "-|-" + "-" * 12 + "-|-" + "-" * 10 + "-|-" + "-" * 18
    print(header)
    print(separator)
    
    row_format = "{:<50} | {:>12} | {:>9.1f}% | {:>17.1f}%"
    for i, entry in enumerate(report_data):
        if i >= args.top_senders:
            break
        row = row_format.format(
            entry['sender'],
            entry['total_emails'],
            entry['unread_percent'],
            entry['subject_similarity_percent']
        )
        print(row)
        
    if not report_data:
         print("No sender data collected.")

    return 0

if __name__ == "__main__":
    sys.exit(main())
 