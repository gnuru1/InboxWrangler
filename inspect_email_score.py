"""
Inspect Email Score Utility

This script connects to Outlook, loads analysis data, retrieves recent emails
from the Inbox, scores them using the main scoring logic, and prints a
detailed breakdown of how the final score was calculated based on the current
configuration weights.

Usage:
    python inspect_email_score.py --limit 5 [--config path/to/config.json] [--data-dir path/to/data] [--debug]
"""

import argparse
import sys
import logging
import datetime
import pywintypes
from pathlib import Path
import textwrap

# Project modules - assuming they are in the same directory or python path
from config import load_config
from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox, safe_get_property
from analyzer import EmailAnalyzer
from scorer import score_email

# Basic logger setup
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('inspect_score') # Specific logger for this script

def setup_logging(level=logging.INFO):
    """Configure logging settings for this script."""
    # Remove existing handlers to avoid duplicate logs if called again
    root_logger = logging.getLogger()
    # Only remove handlers associated with *this* script's logger if needed,
    # or adjust basicConfig if running standalone. For simplicity, let's just set the level.
    logging.getLogger().setLevel(level) # Set root logger level
    logger.setLevel(level) # Set specific logger level
    
    # Ensure console output for this script
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
         handler = logging.StreamHandler(sys.stdout)
         formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
         handler.setFormatter(formatter)
         # Add handler to this script's logger, not root, to avoid conflicts if run via main
         # logger.addHandler(handler) 
         # Actually, adding to root might be better for seeing logs from other modules too
         logging.getLogger().addHandler(handler)

    logger.info(f"Logging setup complete. Level set to: {logging.getLevelName(level)}")


def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description="Inspect email score calculation details.")
    parser.add_argument("--limit", type=int, required=True,
                        help="Number of the most recent emails from the Inbox to inspect.")
    parser.add_argument("--config", type=str, default="./config.json",
                        help="Path to configuration file (default: ./config.json).")
    parser.add_argument("--data-dir", type=str, default="./email_data",
                        help="Directory containing analysis data (default: ./email_data).")
    parser.add_argument("--debug", action="store_true",
                        help="Enable debug logging.")
    return parser.parse_args()

def print_score_breakdown(score_data, config):
    """Prints a formatted breakdown of the score components."""
    if not score_data or 'components' not in score_data or 'metadata' not in score_data:
        print("  Error: Invalid score data received.")
        return

    metadata = score_data['metadata']
    components = score_data['components']
    final_score = score_data['final_score']

    # --- Email Info ---
    subject = metadata.get('subject', 'N/A')
    sender = metadata.get('sender', 'N/A')
    received = metadata.get('received_time', 'N/A')
    if isinstance(received, datetime.datetime):
        received_str = received.strftime('%Y-%m-%d %H:%M:%S')
    else:
        received_str = str(received)

    print(f"Subject : {textwrap.shorten(subject, width=70)}")
    print(f"Sender  : {sender}")
    print(f"Received: {received_str}")
    print(f"----------------------------------------------------------------------")
    print(f"FINAL SCORE: {final_score:.4f}")
    print(f"----------------------------------------------------------------------")
    print(f"{'Component':<20} | {'Raw Score':>10} | {'Weight':>8} | {'Contribution':>15}")
    print(f"---------------------|------------|----------|----------------")

    # --- Get Weights (handle potential missing keys) ---
    # Ensure weights sum to 1.0 for accurate contribution display
    weights = {
        'sender': config.get('sender_weight', 0),
        'topic': config.get('topic_weight', 0),
        'temporal': config.get('temporal_weight', 0),
        'message_state': config.get('message_state_weight', 0),
        'recipient': config.get('recipient_weight', 0),
    }
    total_weight = sum(weights.values())
    if total_weight == 0:
        logger.warning("All scoring weights are zero in the config!")
        factor = 0
    elif abs(total_weight - 1.0) > 0.001:
        logger.debug(f"Weights do not sum to 1.0 (sum={total_weight}). Normalizing for display.")
        factor = 1.0 / total_weight
    else:
        factor = 1.0

    normalized_weights = {k: v * factor for k, v in weights.items()}

    # --- Print Component Breakdown ---
    total_contribution = 0
    component_map = {
        'sender_score': 'Sender',
        'topic_score': 'Topic',
        'temporal_score': 'Temporal',
        'message_state_score': 'Message State',
        'recipient_score': 'Recipient'
    }

    for key, name in component_map.items():
        raw_score = components.get(key, 0.0) # Default to 0.0 if missing
        weight_key = name.lower().replace(' ', '_')
        weight = normalized_weights.get(weight_key, 0.0)
        contribution = raw_score * weight
        total_contribution += contribution
        print(f"{name:<20} | {raw_score:>10.4f} | {weight:>8.3f} | {contribution:>15.4f}")

    # Display a sanity check for the sum vs final score
    print(f"---------------------|------------|----------|----------------")
    print(f"{'Sum of Contributions':<43} | {total_contribution:>15.4f}")
    # Note: Final score might differ slightly if it was capped at 0 or 1 in score_email
    print(f"{'Reported Final Score':<43} | {final_score:>15.4f}")
    print(f"======================================================================
")


def main():
    """Main execution function"""
    args = parse_args()

    # Setup logging based on --debug flag
    log_level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(level=log_level)

    # Load configuration
    try:
        config = load_config(args.config)
        logger.info(f"Loaded configuration from {args.config}")
    except Exception as e:
        logger.error(f"Error loading configuration: {e}")
        print(f"Error: Could not load config file {args.config}. Exiting.", file=sys.stderr)
        return 1

    # Set up data directory Path object
    data_dir = Path(args.data_dir)

    # Connect to Outlook
    outlook, namespace = connect_to_outlook()
    if not namespace:
        print("Error: Failed to connect to Outlook. Exiting.", file=sys.stderr)
        return 1

    # Get Inbox
    inbox = get_default_folder(namespace, olFolderInbox)
    if not inbox:
        print("Error: Could not retrieve Inbox folder. Exiting.", file=sys.stderr)
        return 1

    # Initialize Analyzer to load data (don't run analysis, just load)
    # Suppress analyzer's own logging below INFO unless debugging this script
    logging.getLogger('analyzer').setLevel(logging.INFO if not args.debug else logging.DEBUG)
    try:
        # We pass None for sent_items as we only need inbox for inspection
        analyzer = EmailAnalyzer(namespace, inbox, None, config, data_dir)
        logger.info("Loaded analysis data via EmailAnalyzer.")
    except Exception as e:
         logger.error(f"Error initializing EmailAnalyzer (loading data): {e}", exc_info=True)
         print(f"Error: Failed to load analysis data from {data_dir}. Exiting.", file=sys.stderr)
         return 1
         
    # Check if essential data was loaded
    if not analyzer.sender_scores:
         logger.warning("Sender scores data is empty or failed to load.")
    if not analyzer.email_tracking:
         logger.warning("Email tracking data is empty or failed to load.")
    if not analyzer.contact_map:
         logger.warning("Contact map is empty or failed to load.")
         
    # Don't need LLM for inspection usually, pass None
    llm_service = None
    
    # Get latest emails from Inbox
    logger.info(f"Fetching the latest {args.limit} emails from Inbox '{inbox.Name}'...")
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True) # Sort newest first
        num_items_to_process = min(items.Count, args.limit)
        if num_items_to_process == 0:
             print("Inbox is empty or no items found.")
             return 0
    except pywintypes.com_error as ce:
        logger.error(f"COM Error accessing/sorting Inbox items: {ce}")
        print(f"Error: Could not access Inbox items. {ce}", file=sys.stderr)
        return 1
    except Exception as e:
        logger.error(f"Error accessing/sorting Inbox items: {e}", exc_info=True)
        print(f"Error: Could not access Inbox items. {e}", file=sys.stderr)
        return 1

    print(f"
Inspecting scores for the {num_items_to_process} most recent emails...
")

    processed_count = 0
    errors_count = 0
    # Process items
    for i in range(num_items_to_process):
        item = None
        try:
             # Check index validity
             if (i + 1) > items.Count:
                 logger.warning(f"Inbox item count changed during processing. Stopping.")
                 break
                 
             item = items.Item(i + 1)

             # Check if it's a mail item
             if safe_get_property(item, 'Class', default=-1) != 43: # olMail
                 logger.debug(f"Skipping item {i+1} as it's not a MailItem (Class={safe_get_property(item, 'Class', default='?')}).")
                 continue

             processed_count += 1
             logger.debug(f"Scoring email {i+1}/{num_items_to_process} - Subject: {safe_get_property(item, 'Subject', 'N/A')}")

             # Score the email using the loaded data
             score_data = score_email(
                 email_item=item,
                 sender_scores=analyzer.sender_scores,
                 email_patterns=analyzer.email_patterns, # Contains read/kept stats etc.
                 config=config,
                 llm=llm_service, # Pass None if LLM analysis isn't needed for score inspection
                 email_tracking=analyzer.email_tracking
             )

             # Print the breakdown
             if score_data:
                 print_score_breakdown(score_data, config)
             else:
                 errors_count += 1
                 subject = safe_get_property(item, 'Subject', 'N/A')
                 print(f"Failed to score email: Subject='{subject}'")
                 print(f"======================================================================
")


        except pywintypes.com_error as ce:
            errors_count += 1
            logger.error(f"COM Error processing item index {i+1}: {ce}")
            print(f"Error processing item {i+1}: {ce}
")
        except Exception as e:
            errors_count += 1
            subject = safe_get_property(item, 'Subject', 'N/A') if item else 'N/A'
            logger.error(f"Error processing item {i+1} ('{subject}'): {e}", exc_info=True)
            print(f"Error processing item {i+1} ('{subject}'): {e}
")
        finally:
            # Explicitly release COM object
            if item is not None:
                del item
                item = None

    logger.info(f"Inspection complete. Processed {processed_count} emails, encountered {errors_count} errors.")
    return 0

if __name__ == "__main__":
    sys.exit(main()) 