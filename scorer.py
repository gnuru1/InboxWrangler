"""
Scorer Module for Hybrid Outlook Organizer

This module contains functions for scoring emails based on sender importance,
content analysis, and temporal factors, as well as recommending appropriate actions.
"""

import re
import logging
import datetime
import numpy as np
import pywintypes
from content_processor import process_email_content

logger = logging.getLogger(__name__)

def score_email(email_item, sender_scores, email_patterns, config, llm=None, email_tracking=None):
    """
    Score an individual email based on behavior patterns

    Args:
        email_item: The Outlook email item to score
        sender_scores: Dictionary of sender importance scores
        email_patterns: Dictionary of email reading behavior patterns
        config: Configuration dictionary with scoring weights and thresholds
        llm: Optional LLM service for content analysis
        email_tracking: Optional dictionary of email tracking data to check for repeatedly ignored emails

    Returns:
        Dictionary containing the final score, component scores, and metadata
    """
    # Check if analysis data exists
    if not sender_scores:
        logger.warning("Sender scores not available for scoring. Results may be inaccurate.")
        sender_scores = {}  # Initialize to prevent errors, but scoring will use defaults

    # Extract email metadata safely
    try:
        # Get sender
        sender = "Unknown Sender"
        try:
            sender = getattr(email_item, 'SenderEmailAddress', getattr(email_item, 'SenderName', "Unknown Sender"))
            # Basic email format validation
            if not re.match(r"[^@]+@[^@]+\.[^@]+", sender) and '@' not in sender:  # Simple check if it looks like an address
                sender = getattr(email_item, 'SenderName', "Unknown Sender")  # Prefer Name if Address is odd
        except pywintypes.com_error:
            sender = getattr(email_item, 'SenderName', "Unknown Sender")

        # Get subject and body
        subject = getattr(email_item, 'Subject', "")
        body = getattr(email_item, 'Body', "")

        # Get received time
        received_time = getattr(email_item, 'ReceivedTime', None)
        # Convert to naive datetime for consistent calculation
        if received_time:
            if received_time.tzinfo is None or received_time.tzinfo.utcoffset(received_time) is None:
                received_time_naive = received_time.replace(tzinfo=None)
            else:
                received_time_naive = received_time.replace(tzinfo=None)  # Use naive for simplicity
        else:
            received_time_naive = datetime.datetime.now()  # Fallback to now (naive)

        # Is this a reply or forward?
        is_reply = subject.lower().startswith("re:")
        is_forward = subject.lower().startswith("fw:") or subject.lower().startswith("fwd:")

        # Get conversation ID
        conversation_id = getattr(email_item, 'ConversationID', None)

        # Analyze content using hybrid approach
        content_analysis = process_email_content(email_item, llm, config)

        # --- Calculate Score Components ---

        # 1. Sender Score
        sender_score = 0.5  # Default score

        if sender in sender_scores:
            sender_score = sender_scores[sender].get('normalized_score', 0.5)
        elif '@' in sender:  # Attempt domain matching only if sender looks like an email
            sender_domain = sender.split('@')[-1].split('>')[0].lower()  # Normalize domain
            domain_scores = []
            for known_sender, data in sender_scores.items():
                if '@' in known_sender:
                    known_domain = known_sender.split('@')[-1].split('>')[0].lower()
                    if known_domain and sender_domain and known_domain == sender_domain:
                        domain_scores.append(data.get('normalized_score', 0.5))

            # If we have domain matches, use average domain score
            if domain_scores:
                sender_score = np.mean(domain_scores)
        # Else: keep default 0.5 for non-email-like senders or no matches

        # 2. Re-open Bonus (based on sender read ratios)
        reopen_bonus = 0
        sender_ratios = email_patterns.get('sender_read_ratios', {}).get(sender)
        if sender_ratios:
            # read_ratio = sender_ratios.get('read_ratio', 0.5) # Not used directly in bonus
            reopen_ratio = sender_ratios.get('reopen_ratio', 0)

            # Boost score for senders whose emails are frequently re-opened
            if reopen_ratio > 0.5:  # More than half of emails re-opened
                reopen_bonus = 0.2
            elif reopen_ratio > 0.2:  # More than 20% of emails re-opened
                reopen_bonus = 0.1

        # Apply bonus to sender score (capped at 1.0)
        sender_score = min(1.0, sender_score + reopen_bonus)

        # 3. Topic Score (based on content analysis category and urgency)
        topic_score = 0.5  # Default score
        category = content_analysis.get('category', 'general')
        urgency_level = content_analysis.get('urgency', 'medium')

        # Adjust topic score based on category priority
        category_priority = {  # Define priorities
            'personal': 0.8, 'professional': 0.7, 'transactional': 0.6,
            'general': 0.5,  # Default/General
            'newsletter': 0.3, 'promotional': 0.2
        }
        topic_score = category_priority.get(category, 0.5)  # Use category priority

        # Consider urgency from content analysis
        urgency_map = {'high': 0.9, 'urgent': 0.95, 'medium': 0.6, 'low': 0.3}  # Added 'urgent'
        urgency_score = urgency_map.get(urgency_level, 0.5)

        # Boost urgency score slightly if action items are present
        if content_analysis.get('action_items'):
            urgency_score = min(1.0, urgency_score + 0.1)  # Smaller boost for action items

        # 4. Temporal Score (based on recency and urgency)
        temporal_score = 0.5  # Default score
        now_naive = datetime.datetime.now()  # Use naive for comparison
        days_old = (now_naive - received_time_naive).days if received_time_naive else 365  # Default old if no time

        recency_score = 0.5  # Default
        if days_old <= 1: recency_score = 0.95
        elif days_old <= 3: recency_score = 0.85
        elif days_old <= 7: recency_score = 0.75
        elif days_old <= 14: recency_score = 0.65
        elif days_old <= 30: recency_score = 0.5
        else: recency_score = max(0.1, 0.5 - (days_old / 365))  # Decay for older

        # Adjust for conversation activity (slight boost if part of a thread)
        if conversation_id:
            recency_score = min(1.0, recency_score + 0.05)  # Smaller boost for thread

        # Check if received outside business hours (8 AM - 6 PM)
        if received_time_naive:
            hour = received_time_naive.hour
            if hour < 8 or hour >= 18:
                # Off-hours email bonus (configurable)
                recency_score = min(1.0, recency_score + config.get('off_hours_bonus', 0.05))

        # Combine recency and urgency for the temporal component
        # Weight urgency slightly higher than pure recency
        temporal_score = (0.4 * recency_score + 0.6 * urgency_score)

        # 5. NEW: Message State Score
        message_state_score = 0.5  # Default
        entry_id = getattr(email_item, 'EntryID', None) # Get EntryID earlier for logging
        logger.debug(f"--- Scoring Message State for EntryID: {entry_id} Subject: {subject} ---")

        # Check Unread status - REVERSE LOGIC to penalize unread
        is_unread = getattr(email_item, 'Unread', True)  # Default to TRUE if property not found
        logger.debug(f"  Initial is_unread status: {is_unread}")

        unread_penalty_applied = 0.0
        ignore_penalty_applied = 0.0

        if is_unread:
            # Apply penalty for being unread (instead of bonus)
            base_unread_penalty = config.get('unread_penalty', 0.2)
            message_state_score -= base_unread_penalty
            unread_penalty_applied = base_unread_penalty
            logger.debug(f"  Applied base unread penalty: -{base_unread_penalty:.3f}. Score now: {message_state_score:.3f}")

            # Get entry_id to check for repeated ignoring
            # Moved entry_id retrieval up
            if email_tracking and entry_id and entry_id in email_tracking:
                track_record = email_tracking[entry_id]
                logger.debug(f"  Found tracking record: {track_record}")
                # If email has been seen before and is still unread after multiple checks
                # Use the state *from the previous run* stored in the record
                was_unread_last = not track_record.get('is_currently_read', True) # If is_currently_read was False last time -> was_unread_last = True
                check_count = track_record.get('check_count', 0) # Get the count stored previously
                logger.debug(f"  Tracking - was_unread_last_run: {was_unread_last}, check_count from record: {check_count}")

                # Apply penalty only if it was unread last run AND the counter shows it's been seen unread before (count > 1)
                if was_unread_last and check_count > 1:
                    # Each time we see it again unread, apply additional penalty (scaled)
                    # Use check_count - 1 because the *first* time it's seen unread, count becomes 1, but no penalty applied yet.
                    ignore_penalty_val = config.get('ignore_penalty', 0.15) * (check_count - 1)
                    capped_ignore_penalty = min(0.4, ignore_penalty_val) # Cap at -0.4 to prevent extreme penalties
                    message_state_score -= capped_ignore_penalty
                    ignore_penalty_applied = capped_ignore_penalty
                    logger.debug(f"  Applied ignore penalty: -{capped_ignore_penalty:.3f} (raw: {ignore_penalty_val:.3f}, count: {check_count}). Score now: {message_state_score:.3f}")
                else:
                    logger.debug(f"  Ignore penalty not applied (was_unread_last: {was_unread_last}, check_count: {check_count})")
            else:
                logger.debug(f"  No tracking record found for {entry_id} or email_tracking is None.")
        else:
            # Apply bonus for being read but still in inbox (intentionally kept)
            read_kept_bonus_val = config.get('read_kept_bonus', 0.3)
            message_state_score += read_kept_bonus_val
            logger.debug(f"  Applied read_kept_bonus: +{read_kept_bonus_val:.3f}. Score now: {message_state_score:.3f}")

        # Check Flag status
        flag_status = getattr(email_item, 'FlagStatus', 0)
        flagged_bonus_applied = 0.0
        due_today_bonus_applied = 0.0
        due_soon_bonus_applied = 0.0
        if flag_status == 2:  # olFlagMarked = 2
            flagged_bonus_val = config.get('flagged_bonus', 0.15)
            message_state_score += flagged_bonus_val
            flagged_bonus_applied = flagged_bonus_val
            logger.debug(f"  Applied flagged_bonus: +{flagged_bonus_val:.3f}. Score now: {message_state_score:.3f}")

            # Check for due date
            try:
                flag_due_by = getattr(email_item, 'FlagDueBy', None)
                logger.debug(f"  Flag due date found: {flag_due_by}")
                if flag_due_by and isinstance(flag_due_by, (datetime.datetime, pywintypes.TimeType)):
                     # Convert pywintypes if necessary
                     if isinstance(flag_due_by, pywintypes.TimeType):
                         # Attempt conversion, fallback if error
                         try:
                             flag_due_by = datetime.datetime(flag_due_by.year, flag_due_by.month, flag_due_by.day,
                                                          flag_due_by.hour, flag_due_by.minute, flag_due_by.second)
                         except Exception:
                             logger.warning(f"Could not convert pywintypes due date {flag_due_by} to datetime.")
                             flag_due_by = None

                     if flag_due_by:
                         # Calculate days until due
                         today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                         due_date = flag_due_by.replace(hour=0, minute=0, second=0, microsecond=0)
                         days_until_due = (due_date - today).days
                         logger.debug(f"  Days until due: {days_until_due}")

                         # Add bonus for items due soon
                         if days_until_due <= 0:  # Due today or overdue
                             due_bonus = config.get('due_today_bonus', 0.25)
                             message_state_score += due_bonus
                             due_today_bonus_applied = due_bonus
                             logger.debug(f"  Applied due_today_bonus: +{due_bonus:.3f}. Score now: {message_state_score:.3f}")
                         elif days_until_due <= 2:  # Due in next 2 days
                             due_bonus = config.get('due_soon_bonus', 0.15)
                             message_state_score += due_bonus
                             due_soon_bonus_applied = due_bonus
                             logger.debug(f"  Applied due_soon_bonus: +{due_bonus:.3f}. Score now: {message_state_score:.3f}")
            except Exception as e:
                logger.debug(f"  Error checking flag due date: {e}")

        # Check importance level
        importance = getattr(email_item, 'Importance', 1)  # 2 = olImportanceHigh
        high_importance_bonus_applied = 0.0
        if importance == 2:
            high_importance_bonus_val = config.get('high_importance_bonus', 0.2)
            message_state_score += high_importance_bonus_val
            high_importance_bonus_applied = high_importance_bonus_val
            logger.debug(f"  Applied high_importance_bonus: +{high_importance_bonus_val:.3f}. Score now: {message_state_score:.3f}")

        logger.debug(f"  Final Raw Message State Score (before weighting): {message_state_score:.3f}")

        # 6. NEW: Recipient Information Score
        recipient_score = 0.5  # Default
        logger.debug(f"--- Calculating Recipient Score ---")
        try:
            # Try to get session to determine current user
            session = getattr(email_item, 'Session', None)
            if session:
                try:
                    current_user = session.CurrentUser
                    me_address = getattr(current_user, 'Address', '').lower()
                    logger.debug(f"  Current user address: {me_address}")
                    if me_address:
                        # Check if directly addressed to me
                        to_field = getattr(email_item, 'To', '').lower()
                        cc_field = getattr(email_item, 'CC', '').lower()
                        logger.debug(f"  To field: {to_field}")
                        logger.debug(f"  CC field: {cc_field}")

                        if me_address in to_field:
                            # Direct to me - higher priority
                            to_me_bonus_val = config.get('to_me_bonus', 0.15)
                            recipient_score += to_me_bonus_val
                            logger.debug(f"  Applied to_me_bonus: +{to_me_bonus_val:.3f}. Score now: {recipient_score:.3f}")

                            # Check number of recipients
                            try:
                                recipients = getattr(email_item, 'Recipients', None)
                                if recipients:
                                    recipient_count = recipients.Count
                                    logger.debug(f"  Recipient count: {recipient_count}")
                                    if recipient_count <= 3:
                                        # Few recipients - likely personally addressed
                                        direct_bonus = config.get('direct_to_me_bonus', 0.1)
                                        recipient_score += direct_bonus
                                        logger.debug(f"  Applied direct_to_me_bonus: +{direct_bonus:.3f}. Score now: {recipient_score:.3f}")
                                    elif recipient_count > 10:
                                        # Mass email - lower priority
                                        many_penalty = config.get('many_recipients_penalty', 0.1)
                                        recipient_score -= many_penalty
                                        logger.debug(f"  Applied many_recipients_penalty: -{many_penalty:.3f}. Score now: {recipient_score:.3f}")
                            except Exception as rec_err:
                                logger.debug(f"  Error checking recipient count: {rec_err}")

                        elif me_address in cc_field:
                            # CC'd to me - likely FYI only
                            cc_penalty = config.get('cc_me_penalty', 0.05)
                            recipient_score -= cc_penalty
                            logger.debug(f"  Applied cc_me_penalty: -{cc_penalty:.3f}. Score now: {recipient_score:.3f}")
                except Exception as user_err:
                    logger.debug(f"  Error getting current user: {user_err}")
            else:
                logger.debug("  Session not available for recipient analysis.")
        except Exception as session_err:
            logger.debug(f"  Error getting session for recipient analysis: {session_err}")

        logger.debug(f"  Final Raw Recipient Score (before weighting): {recipient_score:.3f}")

        # --- Calculate Final Score (adjusted for new components) ---
        sender_weight = config.get('sender_weight', 0.4)
        topic_weight = config.get('topic_weight', 0.25)
        temporal_weight = config.get('temporal_weight', 0.15)
        message_state_weight = config.get('message_state_weight', 0.1)  # New weight
        recipient_weight = config.get('recipient_weight', 0.1)  # New weight
        
        # Ensure weights sum to 1.0
        total_weight = sender_weight + topic_weight + temporal_weight + message_state_weight + recipient_weight
        if total_weight != 1.0:
            # Normalize weights proportionally
            factor = 1.0 / total_weight
            sender_weight *= factor
            topic_weight *= factor
            temporal_weight *= factor
            message_state_weight *= factor
            recipient_weight *= factor
        
        final_score = (
            sender_weight * sender_score +
            topic_weight * topic_score +
            temporal_weight * temporal_score +
            message_state_weight * message_state_score +
            recipient_weight * recipient_score
        )
        
        # Ensure score is within 0-1 range
        final_score = max(0.0, min(1.0, final_score))

        # Create score record with new components
        score_data = {
            'final_score': final_score,
            'components': {
                'sender_score': sender_score,
                'topic_score': topic_score,
                'temporal_score': temporal_score,
                'message_state_score': message_state_score,  # New component
                'recipient_score': recipient_score,  # New component
                'urgency_score': urgency_score,
                'reopen_bonus': reopen_bonus,
                'category_priority': category_priority.get(category, 0.5),
                'recency_score': recency_score
            },
            'metadata': {
                'sender': sender,
                'subject': subject,
                'received_time': received_time_naive,
                'is_reply': is_reply,
                'is_forward': is_forward,
                'conversation_id': conversation_id,
                'content_analysis': content_analysis,
                # New metadata
                'is_unread': is_unread,
                'is_flagged': flag_status == 2,
                'importance': importance,
                'flag_due_by': getattr(email_item, 'FlagDueBy', None)
            }
        }

        return score_data

    except pywintypes.com_error as ce:
        logger.error(f"COM error scoring email: {ce}")
        return None
    except Exception as e:
        logger.error(f"Error scoring email: {e}", exc_info=True)  # Log traceback
        return None


def recommend_action(email_item, sender_scores, email_patterns, config, llm=None, email_tracking=None):
    """
    Recommend action for an email based on score and patterns

    Args:
        email_item: The Outlook email item to score
        sender_scores: Dictionary of sender importance scores
        email_patterns: Dictionary of email reading behavior patterns
        config: Configuration dictionary with scoring weights and thresholds
        llm: Optional LLM service for content analysis and folder suggestions
        email_tracking: Dictionary of email tracking history

    Returns:
        Dictionary containing the recommended actions and score data
    """
    # Score the email
    score_data = score_email(email_item, sender_scores, email_patterns, config, llm, email_tracking)

    if not score_data:
        return None  # Cannot recommend if scoring failed

    final_score = score_data['final_score']
    content_analysis = score_data['metadata'].get('content_analysis', {})
    category = content_analysis.get('category', 'general')
    action_items = content_analysis.get('action_items', [])
    sender = score_data['metadata'].get('sender', 'Unknown Sender')
    subject = score_data['metadata'].get('subject', 'No Subject')
    topics = content_analysis.get('topics', [])
    
    # New: extract message state metadata
    is_flagged = score_data['metadata'].get('is_flagged', False)
    flag_due_by = score_data['metadata'].get('flag_due_by', None)
    importance = score_data['metadata'].get('importance', 1)

    # --- Determine Recommended Action ---
    folder = "General"  # Default folder
    flag = False
    create_task = False
    auto_archive = False

    # Check for due-soon flags first (highest priority)
    due_today = False
    if flag_due_by:
        try:
            today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            due_date = flag_due_by.replace(hour=0, minute=0, second=0, microsecond=0) if isinstance(flag_due_by, datetime.datetime) else None
            if due_date:
                days_until_due = (due_date - today).days
                if days_until_due <= 0:  # Due today or overdue
                    folder = "Due Today"
                    flag = True
                    create_task = True
                    due_today = True
        except Exception as e:
            logger.debug(f"Error calculating due date: {e}")
    
    # High importance override (if not already due today)
    if not due_today and importance == 2:  # olImportanceHigh = 2
        if final_score >= config.get('high_priority_threshold', 0.7):
            folder = "High Priority"
        else:
            folder = "Important"
        flag = True
        
    # Priority Folders (if not already handled)
    if not due_today and folder not in ["High Priority", "Important", "Due Today"]:
        if final_score >= config.get('high_priority_threshold', 0.8):
            folder = "High Priority"
            flag = True
            create_task = True  # High priority often implies action needed
        elif final_score >= config.get('medium_priority_threshold', 0.5):
            folder = "Medium Priority"
            flag = True
            # Don't automatically create task for medium, unless action items found

    # Topic/Category Folders for Lower Priority
    if folder in ["General"]:  # Only change folder if still at default
        # Specific handling for common low-priority categories
        if category in ['promotional', 'newsletter']:
            folder = category.capitalize()
            # Check for auto-archive condition
            if final_score < 0.3:  # Threshold for auto-archiving low-prio categories
                auto_archive = True
                folder = f"Archive/{folder}"  # Suggest subfolder in Archive

        # Use LLM to generate folder names for 'personal' or 'professional' if enabled and not high/medium
        elif category in ['personal', 'professional'] and config.get('use_llm_for_content', True) and llm:
            try:
                # Prepare context for LLM folder generation
                email_context = [{'subject': subject, 'topics': topics, 'sender': sender}]
                folder_suggestion = llm.generate_folder_name(email_context)

                if folder_suggestion and isinstance(folder_suggestion, str) and len(folder_suggestion) > 0 and folder_suggestion != "Miscellaneous":
                    # Sanitize folder name (basic)
                    folder_suggestion = re.sub(r'[<>:"/\\|?*]', '_', folder_suggestion)  # Remove invalid chars
                    folder_suggestion = folder_suggestion[:50]  # Limit length
                    folder = f"{category.capitalize()}/{folder_suggestion}"  # Create subfolder under category
                elif topics:  # Fallback to first topic if LLM fails
                    topic_name = re.sub(r'[<>:"/\\|?*]', '_', topics[0].capitalize())[:50]
                    folder = f"{category.capitalize()}/{topic_name}"
                else:  # Fallback to just category
                    folder = category.capitalize()

            except Exception as e:
                logger.error(f"Error generating LLM folder name: {e}")
                folder = category.capitalize()  # Fallback on error

        # Default to category name for other types (transactional, general, etc.)
        else:
            folder = category.capitalize()

    # Adjust actions based on identified action items
    if action_items:
        create_task = True  # Always create task if action items found
        # Optionally elevate folder if not already high/medium priority
        if folder not in ["High Priority", "Medium Priority", "Due Today", "Important"]:
            # Don't override specific category folders like "Promotional" if action items are generic
            # Consider if the *content* suggests it should be moved despite category
            # Simple approach: If action items exist, ensure it's not just archived or in a generic folder
            if auto_archive:  # Don't auto-archive if action needed
                auto_archive = False
                folder = "Action Required"  # Move out of archive path
            elif folder in ["General", "Newsletter", "Promotional", "Transactional"]:
                folder = "Action Required"  # Create a specific folder for actions

    return {
        'folder': folder,
        'flag': flag,
        'create_task': create_task,
        'auto_archive': auto_archive,  # Note: If True, folder might be like "Archive/Promotional"
        'score': final_score,
        'score_data': score_data  # Include for context
    } 