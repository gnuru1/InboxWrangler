"""
Hybrid Outlook Organizer - Main Entry Point

This script serves as the main entry point for the Hybrid Outlook Organizer application.
It parses command-line arguments, sets up logging, initializes components from other modules,
and orchestrates the overall workflow.

Usage:
    python main.py analyze      # Analyze email patterns
    python main.py organize     # Organize inbox based on patterns (dry run)
    python main.py organize --apply  # Organize inbox and apply changes
    python main.py report       # Generate recommendations report
"""

import argparse
import sys
import logging
import subprocess
from pathlib import Path

from config import load_config
from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox, olFolderSentMail
from analyzer import EmailAnalyzer
from llm_service import create_llm_service
from organizer import get_recommendations_report, organize_inbox

logger = logging.getLogger(__name__)

def setup_logging(level=logging.INFO):
    """Configure logging settings."""
    # Remove existing handlers to avoid duplicate logs if called again
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
        
    # Configure logging
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("outlook_organizer.log"), # Log to file
            logging.StreamHandler() # Log to console
        ]
    )
    # Suppress noisy logs from http libraries if needed
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("httpx").setLevel(logging.WARNING)
    
    logger = logging.getLogger(__name__) # Get logger for this module
    logger.info(f"Logging setup complete. Level set to: {logging.getLevelName(level)}")


def parse_args():
    """Parse command-line arguments"""
    # Create the main parser
    parser = argparse.ArgumentParser(description="Hybrid LLM-Enhanced Outlook Email Organizer")
    
    # Add global options to the main parser
    parser.add_argument("--config", type=str, default="./config.json", 
                      help="Path to configuration file")
    parser.add_argument("--data-dir", type=str, default="./email_data", 
                      help="Directory for data storage")
    parser.add_argument("--debug", action="store_true", 
                      help="Enable debug logging")
    
    # Create subparsers with shared parent for global options
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # Create parent with shared arguments
    parent_parser = argparse.ArgumentParser(add_help=False)
    
    # Add shared arguments to parent parser (repeating global arguments)
    parent_parser.add_argument("--config", type=str, default="./config.json", 
                             help=argparse.SUPPRESS)  # Hide in help
    parent_parser.add_argument("--data-dir", type=str, default="./email_data", 
                             help=argparse.SUPPRESS)  # Hide in help
    parent_parser.add_argument("--debug", action="store_true", 
                             help=argparse.SUPPRESS)  # Hide in help
    
    # Analyze command with parent
    analyze_parser = subparsers.add_parser("analyze", help="Analyze email patterns", parents=[parent_parser])
    analyze_parser.add_argument("--limit", type=int, default=5000, 
                              help="Maximum number of emails to analyze per folder")
    
    # Organize command with parent
    organize_parser = subparsers.add_parser("organize", help="Organize inbox based on patterns", parents=[parent_parser])
    organize_parser.add_argument("--limit", type=int, default=None, 
                                help="Maximum number of emails to process")
    organize_parser.add_argument("--apply", action="store_true", 
                               help="Apply changes (default is dry run)")
    
    # Report command with parent
    report_parser = subparsers.add_parser("report", help="Generate recommendations report", parents=[parent_parser])
    report_parser.add_argument("--limit", type=int, default=100, 
                             help="Maximum number of emails to include in report")
    report_parser.add_argument("--output", type=str, default="./reports", 
                             help="Directory for report output")
    
    # Recommend contacts command
    recommend_parser = subparsers.add_parser("recommend", help="Recommend contacts", parents=[parent_parser])
    recommend_parser.add_argument("--limit", type=int, default=10, 
                                 help="Limit number of recommendations")
    recommend_parser.add_argument("--threshold", type=float, default=0.5, 
                                 help="Minimum score threshold")
    
    # Process emails command
    return parser.parse_args()


def main():
    """Main entry point"""
    # Parse arguments
    args = parse_args()
    
    # Set up logging - ENSURE THIS HAPPENS BEFORE ANY OTHER LOGGING
    log_level = logging.DEBUG if args.debug else logging.INFO
    setup_logging(level=log_level)
    
    # Get logger after setup_logging is called
    logger = logging.getLogger(__name__)
    
    # Load configuration
    try:
        config = load_config(args.config)
        logger.info(f"Loaded configuration from {args.config}")
    except Exception as e:
        logger.error(f"Error loading configuration: {e}")
        return 1
    
    # Set up data directory
    data_dir = Path(args.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    
    # Initialize LLM Service based on configuration
    llm_service = None
    if config.get('use_llm_for_content', False):
        logger.info("LLM content analysis is enabled.")
        effective_llm_config = config.get('llm_config', {})
        
        # Check if Copilot Proxy is enabled
        if effective_llm_config.get('use_copilot_proxy', False):
            logger.info("Using Copilot Chat Bridge for LLM service.")
            try:
                # Import the new bridge
                from copilot_chat_bridge import create_copilot_llm_service
                llm_service = create_copilot_llm_service(effective_llm_config.get('copilot_proxy', {}))
                logger.info("Copilot Chat Bridge initialized successfully.")
            except ImportError:
                logger.error("Copilot Chat Bridge (copilot_chat_bridge.py) not found. Please ensure it exists.")
                llm_service = None # Ensure fallback if import fails
            except Exception as e:
                logger.error(f"Error initializing Copilot Chat Bridge: {e}")
                llm_service = None
        else:
            logger.info("Using standard LLMService.")
            try:
                llm_service = create_llm_service(effective_llm_config)
                logger.info("Standard LLMService initialized successfully.")
            except Exception as e:
                logger.error(f"Error initializing standard LLMService: {e}")
                llm_service = None
    
    # Connect to Outlook
    try:
        outlook, namespace = connect_to_outlook()
        if not namespace:
            logger.error("Failed to connect to Outlook. Exiting.")
            return 1
        
        inbox = get_default_folder(namespace, olFolderInbox)
        sent_items = get_default_folder(namespace, olFolderSentMail)
        
        if not inbox:
            logger.error("Could not retrieve Inbox folder. Exiting.")
            return 1
        if not sent_items:
            logger.warning("Could not retrieve Sent Items folder. Some analysis may be skipped.")
    except Exception as e:
        logger.error(f"Error connecting to Outlook: {e}")
        return 1
    
    # Initialize analyzer
    analyzer = EmailAnalyzer(namespace, inbox, sent_items, config, data_dir)
    
    # Process command
    if args.command == "analyze":
        logger.info(f"Starting email analysis (limit: {args.limit} emails per folder)")
        success = analyzer.run_all_analyses(max_items_per_folder=args.limit)
        logger.info(f"Analysis {'completed successfully' if success else 'completed with errors'}")
    
    elif args.command == "organize":
        dry_run = not args.apply
        if dry_run:
            logger.info("Running in DRY RUN mode - no changes will be made")
        else:
            logger.info("Running in APPLY mode - changes will be made to your inbox")
            
        # Make sure we have analysis data
        if not analyzer.sender_scores:
            logger.warning("No sender scores available. Running analysis first...")
            analyzer.run_all_analyses()
        
        stats = organize_inbox(
            inbox=inbox,
            sender_scores=analyzer.sender_scores,
            email_patterns=analyzer.email_patterns,
            config=config,
            llm=llm_service,
            email_tracking=analyzer.email_tracking,
            limit=args.limit,
            dry_run=dry_run
        )
        
        if stats:
            if dry_run:
                logger.info(f"Dry run complete. Processed {stats['processed']} emails with {stats['errors']} errors.")
                logger.info("Run with --apply to apply these changes.")
            else:
                logger.info(f"Organization complete. Moved {stats['moved']} emails with {stats['errors']} errors.")
    
    elif args.command == "report":
        # Make sure we have analysis data
        if not analyzer.sender_scores:
            logger.warning("No sender scores available. Running analysis first...")
            analyzer.run_all_analyses()
        
        output_dir = Path(args.output)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        report_result = get_recommendations_report(
            inbox=inbox,
            sender_scores=analyzer.sender_scores,
            email_patterns=analyzer.email_patterns,
            config=config,
            llm=llm_service,
            email_tracking=analyzer.email_tracking,
            limit=args.limit,
            output_dir=str(output_dir)
        )
        
        if report_result:
            logger.info(f"Report generated with {report_result['num_recommendations']} recommendations")
            logger.info(f"CSV report: {report_result['csv_path']}")
            if report_result.get('html_path'):
                logger.info(f"HTML report: {report_result['html_path']}")
    
    else:
        logger.info("No command specified. Use analyze, organize, or report.")
        logger.info("Run with --help for more information.")
    
    logger.info("Execution completed.")
    return 0


if __name__ == "__main__":
    sys.exit(main()) 