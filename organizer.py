"""
Organizer Module for Hybrid Outlook Organizer

This module handles the actual organization of emails in the inbox,
moving emails to appropriate folders based on scoring and recommendations.
It also generates recommendation reports.
"""

import logging
import datetime
import pywintypes
from pathlib import Path
import pandas as pd

from scorer import score_email, recommend_action
from outlook_utils import get_or_create_folder, create_task_from_email

logger = logging.getLogger(__name__)

def organize_inbox(inbox, sender_scores, email_patterns, config, llm=None, email_tracking=None, limit=None, dry_run=True):
    """
    Organize inbox based on behavior patterns

    Args:
        inbox: Outlook inbox folder object
        sender_scores: Dictionary of sender importance scores
        email_patterns: Dictionary of email reading behavior patterns
        config: Configuration dictionary with scoring weights and thresholds
        llm: Optional LLM service for content analysis
        email_tracking: Dictionary of email tracking history
        limit: Maximum number of emails to process (None = all)
        dry_run: If True, only show recommendations without moving emails

    Returns:
        Dictionary containing statistics about the organization process
    """
    if not inbox:
        logger.error("Inbox folder not found.")
        return None

    logger.info(f"Starting inbox organization (dry_run={dry_run})...")

    # Get inbox items
    try:
        inbox_items = inbox.Items
        inbox_items.Sort("[ReceivedTime]", True)  # Sort by received time, descending
    except pywintypes.com_error as ce:
        logger.error(f"COM error accessing inbox items: {ce}")
        return None
    except Exception as e:
        logger.error(f"Error accessing inbox items: {e}")
        return None

    # Limit number of items to process
    total_inbox_count = inbox_items.Count
    num_items_to_process = total_inbox_count if limit is None else min(total_inbox_count, limit)
    logger.info(f"Found {total_inbox_count} items in Inbox. Processing up to {num_items_to_process}.")

    # Create counters for statistics
    stats = {
        'processed': 0, 'moved': 0, 'flagged': 0, 'task_created': 0,
        'high_priority': 0, 'medium_priority': 0, 'action_required': 0,
        'category_folders': 0, 'archived': 0, 'errors': 0, 'skipped_subfolder': 0
    }

    # Cache for created folders to reduce COM calls
    folder_cache = {}
    processed_ids = set()  # Keep track of processed items to avoid potential duplicates if list changes

    # Process inbox items
    for i in range(num_items_to_process):
        item = None # Initialize for finally block
        try:
            # Check if index is still valid after potential moves
            if (i+1) > inbox_items.Count:
                logger.warning(f"Index {i+1} out of bounds after item moves. Stopping processing.")
                break

            item = inbox_items.Item(i+1)  # 1-based indexing

            # Basic check for valid item and EntryID
            if item is None or not hasattr(item, 'EntryID') or not item.EntryID:
                logger.debug(f"Skipping invalid item at index {i+1}.")
                continue

            entry_id = item.EntryID
            # Skip if already processed
            if entry_id in processed_ids:
                continue

            # Skip non-email items
            if not hasattr(item, 'Class') or item.Class != 43:  # 43 = olMail
                continue

            # Skip if already in a subfolder of the Inbox
            try:
                parent_folder = item.Parent
                if parent_folder is None or parent_folder.EntryID != inbox.EntryID:
                    stats['skipped_subfolder'] += 1
                    processed_ids.add(entry_id)  # Mark as processed even if skipped
                    continue
            except pywintypes.com_error:
                logger.debug(f"COM error checking parent folder for item {getattr(item, 'Subject', 'N/A')}")
                stats['errors'] += 1
                continue  # Skip if parent check fails

            # Get recommendation
            recommendation = recommend_action(item, sender_scores, email_patterns, config, llm, email_tracking)

            if not recommendation:
                stats['errors'] += 1
                logger.warning(f"Could not get recommendation for item: {getattr(item, 'Subject', 'N/A')}")
                continue

            stats['processed'] += 1
            folder_name = recommendation['folder']
            subject_display = getattr(item, 'Subject', 'No Subject')[:50]  # For logging

            # Update statistics
            if folder_name.startswith("High Priority"):
                stats['high_priority'] += 1
            elif folder_name.startswith("Medium Priority"):
                stats['medium_priority'] += 1
            elif folder_name.startswith("Action Required"):
                stats['action_required'] += 1
            elif folder_name.startswith("Archive/"):
                stats['archived'] += 1
            else:
                stats['category_folders'] += 1

            # Log recommendation
            if dry_run:
                logger.info(f"RECOMMENDATION: Move '{subject_display}' to '{folder_name}', Flag: {recommendation['flag']}, Task: {recommendation['create_task']}")
                continue

            # --- PERFORM ACTUAL ACTIONS ---
            # Only executes if dry_run=False

            # 1. Create folder if needed
            try:
                target_folder = None
                if folder_name in folder_cache:
                    target_folder = folder_cache[folder_name]
                else:
                    # Handle potential subfolder path
                    if "/" in folder_name:
                        parent_name, subfolder_name = folder_name.split("/", 1)
                        parent_folder = get_or_create_folder(inbox, parent_name)
                        if parent_folder:
                            target_folder = get_or_create_folder(parent_folder, subfolder_name)
                    else:
                        target_folder = get_or_create_folder(inbox, folder_name)

                    # Add to cache
                    if target_folder:
                        folder_cache[folder_name] = target_folder
                    else:
                        logger.error(f"Failed to create folder '{folder_name}'")
                        stats['errors'] += 1
                        continue

                # 2. Move the item
                if target_folder:
                    item.Move(target_folder)
                    stats['moved'] += 1
                    logger.info(f"Moved: '{subject_display}' to '{folder_name}'")
                else:
                    logger.error(f"Target folder '{folder_name}' not available")
                    stats['errors'] += 1

                # 3. Flag if needed (have to do this before move)
                if recommendation['flag']:
                    try:
                        item.FlagStatus = 2  # 2 = olFlagMarked
                        stats['flagged'] += 1
                    except pywintypes.com_error as flag_err:
                        logger.debug(f"COM error flagging item: {flag_err}")
                    except Exception as flag_err:
                        logger.debug(f"Error flagging item: {flag_err}")

                # 4. Create task if needed
                if recommendation['create_task']:
                    try:
                        task_created = create_task_from_email(item, config.get('task_reminder_days', 2))
                        if task_created:
                            stats['task_created'] += 1
                    except pywintypes.com_error as task_err:
                        logger.debug(f"COM error creating task: {task_err}")
                    except Exception as task_err:
                        logger.debug(f"Error creating task: {task_err}")

            except pywintypes.com_error as move_err:
                logger.error(f"COM error moving item to folder: {move_err}")
                stats['errors'] += 1
            except Exception as act_err:
                logger.error(f"Error during folder/move operations: {act_err}")
                stats['errors'] += 1

            # Mark as processed
            processed_ids.add(entry_id)

            # Log progress for large inboxes
            if stats['processed'] % 50 == 0:
                logger.info(f"Processed {stats['processed']} of {num_items_to_process} items...")

        except pywintypes.com_error as ce:
            logger.error(f"COM error processing inbox item {i+1}: {ce}")
            stats['errors'] += 1
        except Exception as e:
            logger.error(f"Error processing inbox item {i+1}: {e}")
            stats['errors'] += 1
        finally:
            # Explicitly release item COM object
            if item is not None:
                del item
                item = None

    # Log final statistics
    logger.info(f"""
    Inbox organization complete.
    Processed: {stats['processed']}
    Moved: {stats['moved']}
    Flagged: {stats['flagged']}
    Tasks created: {stats['task_created']}
    Errors: {stats['errors']}
    
    By category:
    High Priority: {stats['high_priority']}
    Medium Priority: {stats['medium_priority']}
    Action Required: {stats['action_required']}
    Category folders: {stats['category_folders']}
    Archived: {stats['archived']}
    """)

    return stats


def get_recommendations_report(inbox, sender_scores, email_patterns, config, llm=None, email_tracking=None, limit=100, output_dir="./reports"):
    """
    Generate a recommendations report for inbox organization without making changes

    Args:
        inbox: Outlook inbox folder object
        sender_scores: Dictionary of sender importance scores
        email_patterns: Dictionary of email reading behavior patterns
        config: Configuration dictionary
        llm: Optional LLM service for content analysis
        email_tracking: Dictionary of email tracking history
        limit: Maximum number of emails to analyze
        output_dir: Directory to save the report

    Returns:
        Path to the generated report file
    """
    if not inbox:
        logger.error("Inbox folder not found.")
        return None

    logger.info(f"Generating inbox recommendations report (limit={limit})...")

    # Ensure output directory exists
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # Get inbox items
    try:
        inbox_items = inbox.Items
        inbox_items.Sort("[ReceivedTime]", True)  # Sort by received time, descending
    except Exception as e:
        logger.error(f"Error accessing inbox items: {e}")
        return None

    # Limit number of items to process
    total_inbox_count = inbox_items.Count
    num_items_to_process = min(total_inbox_count, limit)
    logger.info(f"Found {total_inbox_count} items in Inbox. Processing up to {num_items_to_process}.")

    # Prepare data collection
    recommendations_data = []

    # Process inbox items
    for i in range(num_items_to_process):
        item = None # Initialize for finally block
        try:
            item = inbox_items.Item(i+1)  # 1-based indexing

            # Skip non-email items
            if not hasattr(item, 'Class') or item.Class != 43:  # 43 = olMail
                continue

            # Get recommendation
            recommendation = recommend_action(item, sender_scores, email_patterns, config, llm, email_tracking)

            if not recommendation:
                logger.warning(f"Could not get recommendation for item: {getattr(item, 'Subject', 'N/A')}")
                continue

            # Extract essential data
            sender = recommendation['score_data']['metadata'].get('sender', 'Unknown')
            subject = recommendation['score_data']['metadata'].get('subject', 'No Subject')
            received_time = recommendation['score_data']['metadata'].get('received_time', datetime.datetime.now())
            score = recommendation['score']
            folder = recommendation['folder']
            category = recommendation['score_data']['metadata'].get('content_analysis', {}).get('category', 'unknown')
            urgency = recommendation['score_data']['metadata'].get('content_analysis', {}).get('urgency', 'medium')
            flag = recommendation['flag']
            create_task = recommendation['create_task']

            # Append to data collection
            recommendations_data.append({
                'sender': sender,
                'subject': subject,
                'received_time': received_time,
                'score': score,
                'folder': folder,
                'category': category,
                'urgency': urgency,
                'flag': flag,
                'create_task': create_task
            })

            # Log progress for large inboxes
            if len(recommendations_data) % 50 == 0:
                logger.info(f"Processed {len(recommendations_data)} of {num_items_to_process} items...")

        except pywintypes.com_error as ce:
            logger.debug(f"COM error processing inbox item {i+1}: {ce}")
        except Exception as e:
            logger.debug(f"Error processing inbox item {i+1}: {e}")
        finally:
            # Explicitly release item COM object
            if item is not None:
                del item
                item = None

    # Create DataFrame
    if not recommendations_data:
        logger.warning("No recommendations data collected")
        return None

    df = pd.DataFrame(recommendations_data)

    # Generate timestamp for filename
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = output_path / f"email_recommendations_{timestamp}.csv"
    html_path = output_path / f"email_recommendations_{timestamp}.html"

    # Save as CSV
    df.to_csv(csv_path, index=False)
    logger.info(f"Saved recommendations to CSV: {csv_path}")

    # Generate HTML report with basic styling
    try:
        # Sort by score descending for HTML report
        df_sorted = df.sort_values(by='score', ascending=False)

        # Create styled HTML
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Email Organization Recommendations</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1 {{ color: #2c3e50; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th {{ background-color: #2c3e50; color: white; text-align: left; padding: 8px; }}
                td {{ border: 1px solid #ddd; padding: 8px; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .high {{ background-color: #ffdddd; }}
                .medium {{ background-color: #ffffcc; }}
                .low {{ background-color: #ddffdd; }}
                .summary {{ margin-bottom: 20px; padding: 10px; background-color: #eee; }}
            </style>
        </head>
        <body>
            <h1>Email Organization Recommendations</h1>
            <div class="summary">
                <p><strong>Total emails analyzed:</strong> {len(df)}</p>
                <p><strong>Generated:</strong> {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><strong>Breakdown by folder:</strong></p>
                <ul>
        """

        # Add folder breakdown
        folder_counts = df['folder'].value_counts()
        for folder, count in folder_counts.items():
            html_content += f"<li>{folder}: {count}</li>\n"

        html_content += """
                </ul>
            </div>
            <table>
                <tr>
                    <th>Score</th>
                    <th>Sender</th>
                    <th>Subject</th>
                    <th>Received</th>
                    <th>Category</th>
                    <th>Urgency</th>
                    <th>Recommended Folder</th>
                    <th>Actions</th>
                </tr>
        """

        # Add rows for each email
        for _, row in df_sorted.iterrows():
            # Determine row class based on score
            row_class = ""
            if row['score'] >= 0.8:
                row_class = "high"
            elif row['score'] >= 0.5:
                row_class = "medium"
            else:
                row_class = "low"

            # Format date
            date_str = row['received_time'].strftime("%Y-%m-%d %H:%M") if isinstance(row['received_time'], datetime.datetime) else str(row['received_time'])

            # Determine actions
            actions = []
            if row['flag']:
                actions.append("Flag")
            if row['create_task']:
                actions.append("Create Task")
            actions_str = ", ".join(actions) or "None"

            # Add row to HTML
            html_content += f"""
                <tr class="{row_class}">
                    <td>{row['score']:.2f}</td>
                    <td>{row['sender']}</td>
                    <td>{row['subject']}</td>
                    <td>{date_str}</td>
                    <td>{row['category']}</td>
                    <td>{row['urgency']}</td>
                    <td>{row['folder']}</td>
                    <td>{actions_str}</td>
                </tr>
            """

        html_content += """
            </table>
        </body>
        </html>
        """

        # Write HTML file
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        logger.info(f"Saved HTML report to: {html_path}")

    except Exception as html_err:
        logger.error(f"Error generating HTML report: {html_err}")

    return {
        'csv_path': str(csv_path),
        'html_path': str(html_path) if 'html_path' in locals() else None,
        'num_recommendations': len(recommendations_data)
    } 