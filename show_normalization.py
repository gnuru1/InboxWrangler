#!/usr/bin/env python3
"""
Script to inspect contact normalization mappings in the analyzer.
"""

import sys
import logging
from config import load_config
from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox, olFolderSentMail
from analyzer import EmailAnalyzer

def main():
    """Main entry point to display contact normalization mappings."""
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    logger = logging.getLogger('contact_normalization')
    
    # Load configuration
    config = load_config('./config.json')
    logger.info("Loaded configuration")
    
    # Connect to Outlook and get folders
    try:
        outlook, namespace = connect_to_outlook()
        logger.info("Connected to Outlook")
        
        # Get default folders
        inbox = get_default_folder(namespace, olFolderInbox)
        sent_items = get_default_folder(namespace, olFolderSentMail)
        if not inbox or not sent_items:
            logger.error("Could not retrieve required Outlook folders")
            return 1
        logger.info("Retrieved Outlook folders")
        
        # Create analyzer and load data
        analyzer = EmailAnalyzer(namespace, inbox, sent_items, config, 'email_data')
        logger.info("Loaded analyzer data")
        
        # Print the contact normalization map
        print("\nContact Normalization Mappings:")
        if analyzer.contact_map:
            for display_name, email in sorted(analyzer.contact_map.items()):
                print(f"  '{display_name}' -> '{email}'")
            print(f"\nTotal mappings: {len(analyzer.contact_map)}")
        else:
            print("  No contact normalization mappings found")
        
        # Also check email_tracking for entries with both raw_sender and sender
        print("\nContact Pairs in Email Tracking:")
        pairs_found = 0
        for entry_id, data in analyzer.email_tracking.items():
            if 'sender' in data and 'raw_sender' in data:
                sender = data.get('sender', '')
                raw_sender = data.get('raw_sender', '')
                if sender != raw_sender and sender and raw_sender:
                    print(f"  '{raw_sender}' -> '{sender}'")
                    pairs_found += 1
        
        print(f"\nTotal email_tracking pairs: {pairs_found}")
        return 0
        
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1

if __name__ == '__main__':
    sys.exit(main()) 