"""
Standalone HTML Sender Statistics Report Utility

Connects to Outlook, analyzes the Inbox for sender statistics (volume,
read/unread ratios, fuzzy subject similarity), and generates an HTML report.
This script is standalone and does not depend on other project modules like
analyzer.py or config.py, meaning it uses RAW sender names/emails without
contact normalization.

Usage:
    python html_sender_report.py --limit 1000 --output sender_report.html [--top-senders 20] [--fuzzy-threshold 85] [--debug]
"""

import argparse
import sys
import logging
import datetime
import pywintypes
import win32com.client # Direct import for standalone utility functions
from pathlib import Path
from collections import Counter, defaultdict
import re

try:
    from thefuzz import fuzz
except ImportError:
    print("Error: 'thefuzz' library not found. Please install it using:")
    print("pip install thefuzz python-Levenshtein")
    sys.exit(1)

# Basic logger setup
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('html_sender_report')

# --- Standalone Outlook Utilities (Copied from outlook_utils.py) --- 

olFolderInbox = 6

def connect_to_outlook():
    """
    Establish connection to Outlook application and return namespace.
    Tries to connect to an active instance first.
    Returns: tuple: (outlook_app, namespace) or (None, None) on failure.
    """
    try:
        logger.debug("Connecting to Outlook...")
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            logger.debug("Connected to active Outlook instance.")
        except pywintypes.com_error:
            logger.debug("No active Outlook instance found, launching new one...")
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as get_active_err:
             logger.warning(f"Error trying GetActiveObject: {get_active_err}. Attempting Dispatch...")
             outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except pywintypes.com_error as ce:
        logger.error(f"Failed to connect to Outlook (COM Error): {ce}")
        return None, None
    except Exception as e:
        logger.error(f"Failed to connect to Outlook: {e}", exc_info=True)
        return None, None

def get_default_folder(namespace, folder_id):
    """
    Safely get a default folder from the namespace.
    Args: namespace: Outlook MAPI namespace object.
          folder_id (int): The constant representing the folder.
    Returns: MAPIFolder object or None if not found.
    """
    if not namespace:
        logger.error("Cannot get default folder: Namespace is None.")
        return None
    try:
        return namespace.GetDefaultFolder(folder_id)
    except pywintypes.com_error as ce:
        logger.error(f"Could not get default folder ID {folder_id} (COM Error): {ce}")
        return None
    except Exception as e:
        logger.error(f"Error getting default folder ID {folder_id}: {e}")
        return None

def safe_get_property(item, property_name, default=None):
    """
    Safely gets a property from an Outlook item, handling COM errors.
    Args: item: The Outlook item.
          property_name (str): The name of the property to access.
          default: Value to return if property is inaccessible or doesn't exist.
    Returns: The property value or the default.
    """
    try:
        return getattr(item, property_name, default)
    except (pywintypes.com_error, AttributeError) as e:
        # Common error: Property is invalid for this object type or restricted
        logger.debug(f"Error accessing property '{property_name}' on item: {e}")
        return default
    except Exception as e:
        logger.warning(f"Unexpected error accessing property '{property_name}': {e}")
        return default

# --- End Standalone Outlook Utilities --- 

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
    parser = argparse.ArgumentParser(description="Generate HTML report of Inbox sender statistics.")
    parser.add_argument("--limit", type=int, required=True,
                        help="Number of the most recent emails from the Inbox to analyze.")
    parser.add_argument("--output", type=str, required=True,
                        help="Path to save the HTML report file (e.g., sender_report.html).")
    parser.add_argument("--top-senders", type=int, default=20,
                        help="Number of top senders to display in the report (default: 20).")
    parser.add_argument("--fuzzy-threshold", type=int, default=85,
                        help="Similarity threshold (0-100) for grouping subjects (default: 85).")
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
    cleaned = re.sub(r'^(re|fw|fwd|aw|wg):\s*', '', subject.strip(), flags=re.IGNORECASE)
    cleaned = cleaned.lower()
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    cleaned = re.sub(r'\b\d+[/-]\d+([/-]\d+)?\b', ' ', cleaned) # Remove dates
    cleaned = re.sub(r'\b\d+\b', ' ', cleaned) # Remove numbers
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
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
    cleaned_subjects = [clean_subject(s) for s in subjects if clean_subject(s)]
    if not cleaned_subjects:
         return 0.0
    clusters = []
    for subject in cleaned_subjects:
        matched = False
        best_match_idx = -1
        highest_ratio = -1
        for i in range(len(clusters)):
            representative, _ = clusters[i]
            ratio = fuzz.ratio(subject, representative)
            if ratio >= threshold and ratio > highest_ratio:
                highest_ratio = ratio
                best_match_idx = i
                matched = True
        if matched:
            rep, count = clusters[best_match_idx]
            clusters[best_match_idx] = (rep, count + 1)
        else:
            clusters.append((subject, 1))
    if not clusters:
        return 0.0
    max_count = 0
    if clusters:
        max_count = max(count for _, count in clusters)
    similarity_score = (max_count / total_count) * 100
    return similarity_score

def generate_html_report(report_data, num_analyzed, top_n, fuzzy_threshold):
    """Generates an HTML string for the report."""
    html = ["""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inbox Sender Statistics Report</title>
    <style>
        body { font-family: sans-serif; margin: 20px; }
        h1, h2 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        td:nth-child(1) { word-break: break-all; } /* Allow long senders to wrap */
        td:nth-child(2), td:nth-child(3), td:nth-child(4) { text-align: right; }
        .summary { margin-bottom: 20px; padding: 10px; background-color: #eee; border: 1px solid #ccc; }
    </style>
</head>
<body>
    <h1>Inbox Sender Statistics Report</h1>
"""]
    html.append(f'''
    <div class="summary">
        Generated: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}<br>
        Emails Analyzed: {num_analyzed}<br>
        Fuzzy Subject Threshold: {fuzzy_threshold}
    </div>
''')
    html.append(f"<h2>Top {min(top_n, len(report_data))} Senders</h2>")
    html.append("<table>")
    html.append("<tr><th>Sender</th><th>Total Emails</th><th>% Unread</th><th>% Fuzzy Subject</th></tr>")

    for i, entry in enumerate(report_data):
        if i >= top_n:
            break
        sender_escaped = entry['sender'].replace('<', '&lt;').replace('>', '&gt;') # Basic HTML escape
        html.append(f"""
        <tr>
            <td>{sender_escaped}</td>
            <td>{entry['total_emails']}</td>
            <td>{entry['unread_percent']:.1f}%</td>
            <td>{entry['subject_similarity_percent']:.1f}%</td>
        </tr>
""")

    html.append("</table>")
    if not report_data:
        html.append("<p>No sender data collected.</p>")
    html.append("</body></html>")
    return "\n".join(html)

def main():
    """Main execution function"""
    args = parse_args()
    log_level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(level=log_level)

    # --- Connect to Outlook & Get Inbox --- 
    outlook, namespace = connect_to_outlook()
    if not namespace:
        print("Error: Failed to connect to Outlook. Exiting.", file=sys.stderr)
        return 1
    inbox = get_default_folder(namespace, olFolderInbox)
    if not inbox:
        print("Error: Could not retrieve Inbox folder. Exiting.", file=sys.stderr)
        return 1

    # --- Process Inbox Items --- 
    logger.info(f"Fetching the latest {args.limit} emails from Inbox '{safe_get_property(inbox, 'Name', 'Inbox')}'...")
    sender_data = defaultdict(lambda: {'subjects': [], 'read': 0, 'unread': 0, 'total': 0})
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
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
            
            # Get Raw Sender (No Normalization)
            raw_sender = "Unknown Sender"
            sender_addr = safe_get_property(item, 'SenderEmailAddress')
            sender_name = safe_get_property(item, 'SenderName')
            if sender_addr and '@' in sender_addr:
                raw_sender = sender_addr # Prefer email address if valid
            elif sender_name:
                raw_sender = sender_name # Fallback to display name
            
            # Get Subject & Read Status
            subject = safe_get_property(item, 'Subject', default="")
            is_unread = safe_get_property(item, 'UnRead', default=True)
            
            # Store data against raw sender
            stats = sender_data[raw_sender] # Use raw_sender as key
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
            subject_err = safe_get_property(item, 'Subject', 'N/A') if item else 'N/A'
            logger.debug(f"Error processing item {i+1} ('{subject_err}'): {e}", exc_info=True)
        finally:
            if item is not None:
                del item
                item = None
                
    logger.info(f"Finished processing {processed_count} emails with {errors_count} errors.")

    # --- Calculate Stats & Prepare Report Data --- 
    report_data = []
    for sender, data in sender_data.items():
        total = data['total']
        unread = data['unread']
        subjects = data['subjects']
        unread_percent = (unread / total * 100) if total > 0 else 0
        subject_similarity = calculate_fuzzy_subject_similarity(subjects, args.fuzzy_threshold)
        report_data.append({
            'sender': sender,
            'total_emails': total,
            'unread_percent': unread_percent,
            'subject_similarity_percent': subject_similarity
        })
        
    report_data.sort(key=lambda x: x['total_emails'], reverse=True)
    
    # --- Generate and Save HTML Report --- 
    html_content = generate_html_report(report_data, processed_count, args.top_senders, args.fuzzy_threshold)
    output_path = Path(args.output)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        logger.info(f"HTML report saved to {output_path.resolve()}")
        print(f"\nHTML report saved to: {output_path.resolve()}")
    except Exception as e:
        logger.error(f"Failed to save HTML report to {output_path}: {e}")
        print(f"\nError: Failed to save HTML report to {output_path}: {e}", file=sys.stderr)
        return 1

    return 0

if __name__ == "__main__":
    sys.exit(main()) 