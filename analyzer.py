import logging
import pickle
import datetime
import re
import numpy as np
import pandas as pd
import pywintypes
from collections import defaultdict, Counter
from pathlib import Path
from datetime import timezone

# Assuming outlook_utils provides safe_get_property and folder constants if needed
from outlook_utils import safe_get_property #, olFolderInbox, olFolderSentMail

logger = logging.getLogger(__name__)

class EmailAnalyzer:
    """
    Analyzes Outlook email data including sent items, inbox behavior,
    and folder structure to derive patterns and scores.
    """
    def __init__(self, namespace, inbox, sent_items, config, data_dir):
        """
        Initialize the analyzer.

        Args:
            namespace: Outlook MAPI namespace object.
            inbox: MAPIFolder object for the Inbox.
            sent_items: MAPIFolder object for Sent Items.
            config (dict): Application configuration dictionary.
            data_dir (str or Path): Path to the directory for saving/loading analysis data.
        """
        self.namespace = namespace
        self.inbox = inbox
        self.sent_items = sent_items
        self.config = config
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

        # Internal state for analysis results
        self.sender_scores = {}
        self.email_patterns = {} # Will store inbox behavior results
        self.folder_structure = {}
        self.email_tracking = {} # For tracking read status across runs
        self.conversation_history = defaultdict(list) # Stores {conv_id: [ {received_time: dt, sender: str, entry_id: str}, ... ]}
        
        # Contact normalization map to resolve duplicate identities
        self.contact_map = {} # Maps display names to email addresses
        
        # Load existing data if available
        self._load_analysis_data()
        
        # Initialize contact map now that we have loaded the email_tracking data
        self._initialize_contact_map()

    def _load_analysis_data(self):
        """ Load previously saved analysis data from pickle files, including post-load validation. """
        # --- Load Email Tracking Data ---
        tracking_file = self.data_dir / 'email_tracking.pkl'
        if tracking_file.exists():
            try:
                with open(tracking_file, 'rb') as f:
                    loaded_tracking = pickle.load(f)

                # Post-load validation and conversion for tracking data
                validated_tracking = {}
                conversion_needed = False
                for entry_id, record in loaded_tracking.items():
                    validated_record = record.copy()
                    for key in ['received_time', 'last_modified', 'last_checked', 'first_opened_time', 'sent_on_time']:
                        if key in validated_record and validated_record[key] is not None:
                            original_value = validated_record[key]
                            converted_value = self._convert_to_naive_datetime(original_value)
                            if converted_value is None:
                                logger.warning(f"Post-load conversion failed for '{key}' in {entry_id}. Removing timestamp.")
                                validated_record[key] = None
                                conversion_needed = True
                            elif original_value is not converted_value:
                                validated_record[key] = converted_value
                                conversion_needed = True # Mark if actual conversion happened

                    validated_tracking[entry_id] = validated_record

                self.email_tracking = validated_tracking
                if conversion_needed:
                    logger.info(f"Performed post-load timestamp conversion/validation for {tracking_file}")
                else:
                    logger.info(f"Loaded email tracking data from {tracking_file}")

            except (EOFError, pickle.UnpicklingError) as load_err:
                 logger.warning(f"Could not load email tracking file '{tracking_file}': {load_err}. Starting fresh.")
                 self.email_tracking = {}
            except Exception as e:
                 logger.warning(f"Error loading email tracking file '{tracking_file}': {e}. Starting fresh.")
                 self.email_tracking = {}
        else:
            self.email_tracking = {}
            logger.info("Email tracking file not found. Starting fresh.")

        # --- Load Sender Scores Data ---
        sender_scores_file = self.data_dir / 'sender_scores.pkl'
        if sender_scores_file.exists():
            try:
                 with open(sender_scores_file, 'rb') as f:
                     loaded_scores = pickle.load(f)

                 # Post-load validation and conversion for sender scores
                 validated_scores = {}
                 conversion_needed_scores = False
                 for sender, score_data in loaded_scores.items():
                     validated_score_data = score_data.copy()
                     key = 'last_interaction'
                     if key in validated_score_data and validated_score_data[key] is not None:
                          original_value = validated_score_data[key]
                          converted_value = self._convert_to_naive_datetime(original_value)
                          if converted_value is None:
                               logger.warning(f"Post-load conversion failed for '{key}' for sender {sender}. Removing timestamp.")
                               validated_score_data[key] = None
                               conversion_needed_scores = True
                          elif original_value is not converted_value:
                               validated_score_data[key] = converted_value
                               conversion_needed_scores = True
                     validated_scores[sender] = validated_score_data

                 self.sender_scores = validated_scores
                 if conversion_needed_scores:
                     logger.info(f"Performed post-load timestamp conversion/validation for {sender_scores_file}")
                 else:
                     logger.info(f"Loaded sender scores data from {sender_scores_file}")

            except (EOFError, pickle.UnpicklingError) as load_err:
                logger.warning(f"Could not load sender scores file '{sender_scores_file}': {load_err}. Starting fresh.")
                self.sender_scores = {}
            except Exception as e:
                logger.warning(f"Error loading sender scores file '{sender_scores_file}': {e}. Starting fresh.")
                self.sender_scores = {}
        else:
             self.sender_scores = {}
             logger.info("Sender scores file not found. Starting fresh.")

        # --- Load Contact Map Data ---
        contact_map_file = self.data_dir / 'contact_map.pkl'
        if contact_map_file.exists():
            try:
                with open(contact_map_file, 'rb') as f:
                    self.contact_map = pickle.load(f)
                logger.info(f"Loaded contact map from {contact_map_file} with {len(self.contact_map)} mappings")
            except Exception as e:
                logger.warning(f"Error loading contact map file '{contact_map_file}': {e}. Starting fresh.")
                self.contact_map = {}
        else:
             self.contact_map = {}
             logger.info("Contact map file not found. Starting fresh.")

        # --- Load Inbox Behavior Data --- (No timestamps stored directly)
        inbox_behavior_file = self.data_dir / 'inbox_behavior.pkl'
        if inbox_behavior_file.exists():
             try:
                with open(inbox_behavior_file, 'rb') as f:
                    loaded_patterns = pickle.load(f)
                    self.email_patterns = loaded_patterns
                logger.info(f"Loaded inbox behavior data from {inbox_behavior_file}")
             except Exception as e:
                logger.warning(f"Could not load inbox behavior file '{inbox_behavior_file}': {e}")
                self.email_patterns = {}
                # Ensure the nested dict for read/kept stats exists if loading old data
                if 'sender_read_kept_stats' not in self.email_patterns:
                    self.email_patterns['sender_read_kept_stats'] = {}
        else:
             self.email_patterns = {}
             self.email_patterns['sender_read_kept_stats'] = {} # Initialize if file doesn't exist
             logger.info("Inbox behavior file not found. Starting fresh.")

        # --- Load Folder Structure Data --- (No timestamps stored directly)
        folder_structure_file = self.data_dir / 'folder_structure.pkl'
        if folder_structure_file.exists():
             try:
                 with open(folder_structure_file, 'rb') as f:
                     self.folder_structure = pickle.load(f)
                 logger.info(f"Loaded folder structure data from {folder_structure_file}")
             except Exception as e:
                 logger.warning(f"Could not load folder structure file '{folder_structure_file}': {e}")
                 self.folder_structure = {}
        else:
            self.folder_structure = {}
            logger.info("Folder structure file not found. Starting fresh.")

    def _convert_to_naive_datetime(self, time_obj):
        """Converts pywintypes.TimeType or datetime.datetime to a naive Python datetime."""
        if isinstance(time_obj, datetime.datetime):
            # If already datetime, just make naive
            return time_obj.replace(tzinfo=None)
        elif isinstance(time_obj, pywintypes.TimeType):
            # Explicitly convert pywintypes.TimeType to datetime.datetime
            try:
                # Convert via formatting, which might be more robust
                # Format: YYYY-MM-DD HH:MM:SS (ISO compatible subset)
                time_str = time_obj.Format('%Y-%m-%d %H:%M:%S')
                dt = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                # Microseconds might be lost, but guarantees standard datetime
                return dt
            except Exception as format_err:
                 logger.debug(f"Could not convert pywintypes time {time_obj} using Format method: {format_err}. Trying property access...")
                 # Fallback to property access if Format fails
                 try:
                     return datetime.datetime(
                         time_obj.year, time_obj.month, time_obj.day,
                         time_obj.hour, time_obj.minute, time_obj.second,
                         getattr(time_obj, 'microsecond', 0), # Include microseconds if possible
                         tzinfo=None # Ensure naive
                     )
                 except Exception as prop_err:
                      logger.error(f"Failed to convert pywintypes time {time_obj} using properties: {prop_err}")
                      return None # Give up if both methods fail

        # Try parsing LAST if it's some other type (e.g., string) and not None
        elif time_obj:
             try:
                # Use pandas for robust parsing of various formats
                dt_obj = pd.to_datetime(str(time_obj))
                # Ensure it's timezone-naive before returning
                if dt_obj.tzinfo is not None:
                    # Convert to UTC then remove tz for consistency if timezone aware
                    dt_obj = dt_obj.tz_convert('UTC').tz_localize(None)
                return dt_obj.replace(tzinfo=None) # Final explicit make naive
             except Exception as parse_err:
                 logger.debug(f"Could not convert time object {time_obj} of type {type(time_obj)} to datetime using pandas: {parse_err}")
                 return None

        return None # Return None if input is None or conversion failed

    def _initialize_contact_map(self):
        """Initialize the contact map from existing data to normalize contact identities without manual mapping."""
        logger.info("Initializing contact normalization map...")
        
        # Only rebuild the map if it's empty or we're specifically asked to
        if self.contact_map:
            logger.info(f"Using existing contact map with {len(self.contact_map)} mappings")
            return
            
        # Maps to store display name to email and email to display name relationships
        email_to_names = defaultdict(set)
        name_to_emails = defaultdict(set)
        name_parts = {}  # Store first/last name parts
        
        try:
            # Step 0: First pass for raw sender preservation
            logger.debug("Step 0: Looking for raw_sender fields for initial mapping")
            raw_sender_pairs = []
            
            for entry_id, data in self.email_tracking.items():
                sender = data.get('sender', '').strip()
                raw_sender = data.get('raw_sender', '').strip()
                sender_name = data.get('sender_name', '').strip()
                sender_email = data.get('sender_email', '').strip()
                
                # If we have both raw_sender and sender fields, and they're different formats
                if sender and raw_sender and sender != raw_sender:
                    if '@' in sender and '@' not in raw_sender:
                        # sender is email, raw_sender is display name
                        email, display_name = sender.lower(), raw_sender.lower()
                        raw_sender_pairs.append((display_name, email))
                        logger.debug(f"  Found raw sender pair: '{display_name}' -> '{email}'")
                    elif '@' in raw_sender and '@' not in sender:
                        # raw_sender is email, sender is display name
                        email, display_name = raw_sender.lower(), sender.lower()
                        raw_sender_pairs.append((display_name, email))
                        logger.debug(f"  Found raw sender pair: '{display_name}' -> '{email}'")
                
                # If we have explicit sender_name and sender_email fields (added in newer versions)
                if sender_name and sender_email and '@' in sender_email:
                    email, display_name = sender_email.lower(), sender_name.lower()
                    raw_sender_pairs.append((display_name, email))
                    logger.debug(f"  Found explicit name/email pair: '{display_name}' -> '{email}'")
            
            # Step 1: Build initial relationships from email tracking data
            logger.debug("Step 1: Analyzing email data to extract name/email components")
            for entry_id, data in self.email_tracking.items():
                sender = data.get('sender', '').strip()
                if not sender:
                    continue
                
                # Skip system email addresses and likely non-human senders
                if any(pattern in sender.lower() for pattern in ['noreply', 'no-reply', 'donotreply', 'system', 'notification', 'alert']):
                    continue
                    
                if '@' in sender:  # It's an email address
                    # Extract potential name parts from email (for later matching)
                    email_local = sender.split('@')[0].lower()
                    
                    # Process email local part to find potential name components
                    # e.g., "john.doe" -> "john" and "doe"
                    name_parts_from_email = re.split(r'[._-]', email_local)
                    if len(name_parts_from_email) > 1:
                        for part in name_parts_from_email:
                            if len(part) > 2:  # Avoid tiny fragments
                                name_parts[part] = sender
                                logger.debug(f"  Extracted name component '{part}' from email '{sender}'")
                else:  # It's a display name
                    # Store name words for potential matching
                    name_words = [w.lower() for w in re.split(r'\W+', sender) if len(w) > 2]
                    for word in name_words:
                        name_parts[word] = sender
                        logger.debug(f"  Extracted name component '{word}' from display name '{sender}'")
            
            # Step 2: Find correspondences between names and emails in the data
            logger.debug("Step 2: Finding correspondences between names and emails")
            
            # First add the raw sender pairs
            for display_name, email in raw_sender_pairs:
                email_to_names[email].add(display_name)
                name_to_emails[display_name].add(email)
                logger.debug(f"  Added correspondence from raw data: '{display_name}' <-> '{email}'")
            
            # Then find other correspondences through the email data
            for entry_id, data in self.email_tracking.items():
                sender = data.get('sender', '').strip()
                if not sender:
                    continue
                    
                if '@' in sender:  # It's an email
                    email = sender.lower()
                    # Look for display names in other messages from this sender
                    for other_id, other_data in self.email_tracking.items():
                        other_sender = other_data.get('sender', '').strip()
                        if other_sender and '@' not in other_sender:
                            # Check if email contains parts of the name or vice versa
                            name_lower = other_sender.lower()
                            name_words = [w for w in re.split(r'\W+', name_lower) if len(w) > 2]
                            
                            # Match if a substantial name part is in the email or vice versa
                            email_local = email.split('@')[0]
                            if (any(word in email_local for word in name_words) or
                                any(part in name_lower for part in re.split(r'[._-]', email_local))):
                                email_to_names[email].add(name_lower)
                                name_to_emails[name_lower].add(email)
                                logger.debug(f"  Found correspondence: '{other_sender}' <-> '{email}'")
                else:  # It's a display name
                    name = sender.lower()
                    name_words = [w for w in re.split(r'\W+', name) if len(w) > 2]
                    
                    # Match with emails containing parts of this name
                    for other_id, other_data in self.email_tracking.items():
                        other_sender = other_data.get('sender', '').strip()
                        if other_sender and '@' in other_sender:
                            email = other_sender.lower()
                            email_local = email.split('@')[0]
                            
                            # Match if a substantial name part is in the email or vice versa
                            if (any(word in email_local for word in name_words) or
                                any(part in name for part in re.split(r'[._-]', email_local))):
                                email_to_names[email].add(name)
                                name_to_emails[name].add(email)
                                logger.debug(f"  Found correspondence: '{name}' <-> '{email}'")
            
            # Add specific matching for names based on name components
            for name in list(name_to_emails.keys()):
                name_words = [w for w in re.split(r'\W+', name) if len(w) > 2]
                
                for email_data in self.email_tracking.values():
                    sender = email_data.get('sender', '').lower()
                    if '@' in sender and any(word in sender for word in name_words):
                        # Strong match found
                        email_to_names[sender].add(name)
                        name_to_emails[name].add(sender)
                        logger.debug(f"  Added correspondence using name components: '{name}' <-> '{sender}'")
            
            # Log found relationships for debugging
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug("Name to emails relationships:")
                for name, emails in name_to_emails.items():
                    logger.debug(f"  '{name}' -> {emails}")
                    
                logger.debug("Email to names relationships:")
                for email, names in email_to_names.items():
                    logger.debug(f"  '{email}' -> {names}")
            
            # Step 3: Resolve the best email address for each display name
            logger.debug("Step 3: Resolving best email for each display name")
            for name, emails in name_to_emails.items():
                if emails:
                    # Prefer the email that matches the name pattern better
                    best_email = None
                    best_score = 0
                    
                    name_words = [w for w in re.split(r'\W+', name) if len(w) > 2]
                    
                    for email in emails:
                        email_local = email.split('@')[0]
                        
                        # Calculate match score based on name parts found in email
                        score = 0
                        
                        # Check each word in name against email
                        for word in name_words:
                            if word in email_local:
                                score += 2  # Direct match is strongest
                                logger.debug(f"  Word '{word}' direct match in '{email_local}' +2 points")
                            elif any(word in part for part in re.split(r'[._-]', email_local)):
                                score += 1  # Partial match
                                logger.debug(f"  Word '{word}' partial match in '{email_local}' +1 point")
                        
                        # Check each part of email against name
                        for part in re.split(r'[._-]', email_local):
                            if len(part) > 2 and part in name:
                                score += 2  # Direct match
                                logger.debug(f"  Email part '{part}' direct match in '{name}' +2 points")
                            elif any(part in word for word in name_words):
                                score += 1  # Partial match
                                logger.debug(f"  Email part '{part}' partial match in '{name}' +1 point")
                        
                        # Names exactly matching email local parts get highest score
                        # e.g., "john.doe@example.com" with "John Doe"
                        if ''.join(name_words) == email_local.replace('.', '').replace('-', '').replace('_', ''):
                            score += 5
                            logger.debug(f"  Full name match pattern for '{name}' and '{email}' +5 points")
                            
                        # First initial + last name pattern (j.smith@example.com with "John Smith")
                        if (len(name_words) > 1 and 
                            email_local.startswith(name_words[0][0]) and 
                            email_local[1:].startswith(name_words[-1])):
                            score += 4
                            logger.debug(f"  First initial + last name pattern for '{name}' and '{email}' +4 points")
                        
                        logger.debug(f"  Score for '{name}' -> '{email}': {score}")
                        
                        if score > best_score:
                            best_score = score
                            best_email = email
                            
                    # Lower the threshold to 1 to increase matches
                    if best_email and best_score >= 1:
                        self.contact_map[name] = best_email
                        logger.debug(f"  MAPPED: '{name}' -> '{best_email}' (score: {best_score})")
            
            # Step 4: Fast track for direct mappings from SenderName -> SenderEmailAddress
            direct_mappings_count = 0
            for display_name, email in raw_sender_pairs:
                if display_name and email and '@' in email:
                    # Don't overwrite higher-confidence mappings
                    if display_name not in self.contact_map:
                        self.contact_map[display_name] = email
                        direct_mappings_count += 1
                        logger.debug(f"  DIRECT MAPPING: '{display_name}' -> '{email}'")
            
            if direct_mappings_count > 0:
                logger.info(f"Added {direct_mappings_count} direct mappings from sender_name/sender_email fields")
                    
            logger.info(f"Contact normalization map initialized with {len(self.contact_map)} mappings")
            if self.contact_map:
                logger.debug("Contact map entries:")
                for name, email in self.contact_map.items():
                    logger.debug(f"  '{name}' -> '{email}'")
                    
                # Save the newly built contact map
                self._save_contact_map()
        except Exception as e:
            logger.error(f"Error initializing contact map: {e}")
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug("Traceback:", exc_info=True)
    
    def _normalize_contact(self, contact):
        """
        Normalize a contact identifier to a consistent form.
        Converts known display names to email addresses.
        
        Args:
            contact (str): Display name or email address
            
        Returns:
            str: Normalized contact identifier (preferably email address)
        """
        if not contact:
            return "unknown"
            
        contact_lower = contact.lower()
        
        # If already an email address, return as is
        if '@' in contact_lower:
            return contact_lower
            
        # Check if we know this display name
        if contact_lower in self.contact_map:
            normalized = self.contact_map[contact_lower]
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"Normalized contact: '{contact}' -> '{normalized}'")
            return normalized
            
        # No mapping found, return as is
        return contact_lower
        
    def _get_sender_address(self, item):
        """Safely retrieve the sender's email address or name."""
        sender_addr = safe_get_property(item, 'SenderEmailAddress')
        # Check for Exchange addresses (no '@')
        if sender_addr and '@' not in sender_addr and sender_addr.startswith('/'):
            try:
                # Attempt to resolve Exchange address
                sender_entry = self.namespace.CreateRecipient(sender_addr)
                sender_entry.Resolve()
                if sender_entry.Resolved:
                    exchange_user = sender_entry.AddressEntry.GetExchangeUser()
                    if exchange_user:
                        return exchange_user.PrimarySmtpAddress
                # Fallback if resolution fails
                sender_name = safe_get_property(item, 'SenderName')
                return sender_name if sender_name else "Unknown Sender"
            except Exception as e:
                logger.debug(f"Error resolving Exchange sender address {sender_addr}: {e}")
                # Fallback to name if Exchange resolution fails
                sender_name = safe_get_property(item, 'SenderName')
                return sender_name if sender_name else "Unknown Sender"
        elif sender_addr and '@' in sender_addr:
             return sender_addr # Use SMTP address if available
        else:
             # Fallback to SenderName if address is invalid or missing
             sender_name = safe_get_property(item, 'SenderName')
             return sender_name if sender_name else "Unknown Sender"

    def run_all_analyses(self, max_items_per_folder=None):
        """ Runs all analysis components and saves the results. """
        if not self.namespace:
            logger.error("Outlook namespace not available. Cannot run analyses.")
            return False

        logger.info("Starting comprehensive email analysis...")
        success = True
        max_items = max_items_per_folder or self.config.get('max_analysis_emails', 5000)

        try:
            if self.sent_items:
                logger.info("--- Analyzing Sent Items ---")
                self.analyze_sent_items(max_items=max_items)
                # Sender scores are calculated and saved within analyze_sent_items
            else:
                logger.warning("Sent Items folder not available, skipping analysis.")
        except Exception as e:
            logger.error(f"Error during analyze_sent_items: {e}", exc_info=True)
            success = False

        try:
            logger.info("--- Analyzing Folder Structure ---")
            self.analyze_folder_structure() # Analysis result stored in self.folder_structure
            self._save_folder_structure() # Save the result
        except Exception as e:
            logger.error(f"Error during analyze_folder_structure: {e}", exc_info=True)
            success = False

        try:
            if self.inbox:
                logger.info("--- Analyzing Inbox Behavior ---")
                self.analyze_inbox_behavior(max_items=max_items)
                # Inbox behavior and tracking data are saved within analyze_inbox_behavior
            else:
                logger.warning("Inbox folder not available, skipping inbox behavior analysis.")
        except Exception as e:
            logger.error(f"Error during analyze_inbox_behavior: {e}", exc_info=True)
            success = False

        if success:
            logger.info("Comprehensive email analysis complete.")
        else:
            logger.warning("Comprehensive email analysis completed with errors.")

        return success


    # ---------------- Sent Items Analysis ----------------
    def analyze_sent_items(self, max_items=None):
        """
        Analyze sent mail to identify response patterns and contact importance.
        Updates self.sender_scores based on replies AND initiations.
        """
        if not self.sent_items:
            logger.warning("Sent Items folder not found. Skipping analysis.")
            return None

        logger.info("Analyzing sent items to identify response patterns AND initiations...")
        max_items = max_items or self.config.get('max_analysis_emails', 5000)
        sent_folder = self.sent_items

        try:
            sent_items_collection = sent_folder.Items
            sent_items_collection.Sort("[ReceivedTime]", True) # Sort by received time, descending
            total_items = sent_items_collection.Count
        except pywintypes.com_error as ce:
            logger.error(f"COM Error accessing Sent Items: {ce}")
            return None
        except Exception as e:
            logger.error(f"Error accessing Sent Items: {e}")
            return None

        num_to_process = min(total_items, max_items)
        logger.info(f"Processing up to {num_to_process} items from Sent Items.")

        # Data structures for this analysis run
        response_data = defaultdict(list) # For replies to incoming emails
        sender_initiations = defaultdict(lambda: {'count': 0, 'dates': []}) # For emails sent TO contacts
        conversation_threads = defaultdict(list) # For general conversation analysis

        processed_count = 0
        for i in range(num_to_process):
            item = None # Initialize for finally block
            try:
                item = sent_items_collection.Item(i + 1)

                if not hasattr(item, 'Class') or item.Class != 43: continue # olMail

                # --- Get Essential Info ---
                recipients_to = [] # Primary recipients (To field)
                all_recipients = [] # All recipients (To, CC, BCC)
                try:
                    for recipient in item.Recipients:
                        # olTo = 1, olCC = 2, olBCC = 3
                        recipient_type = getattr(recipient, 'Type', 1)
                        # Normalize address - prefer Address if available, else Name
                        address = getattr(recipient, 'Address', None)
                        name = getattr(recipient, 'Name', 'Unknown')
                        raw_recipient = address if address and '@' in address else name
                        normalized_recipient = self._normalize_contact(raw_recipient)
                        
                        all_recipients.append(normalized_recipient)
                        if recipient_type == 1: # Only track 'To' recipients for initiation score
                            recipients_to.append(normalized_recipient)

                except pywintypes.com_error as r_ce:
                    logger.debug(f"COM Error accessing recipients for item {i+1}: {r_ce}")

                conversation_id = getattr(item, 'ConversationID', None)
                sent_on_time = getattr(item, 'SentOn', None)
                subject = getattr(item, 'Subject', '')
                body_len = len(getattr(item, 'Body', ''))
                sent_on_time_naive = self._convert_to_naive_datetime(sent_on_time)
                
                # Skip if critical info is missing
                if not sent_on_time_naive: continue

                # --- Track Initiations --- 
                if not subject.lower().startswith("re:") and not subject.lower().startswith("fw:"): # Not a reply/forward
                    for recipient in recipients_to: # Score based on To: field
                        sender_initiations[recipient]['count'] += 1
                        sender_initiations[recipient]['dates'].append(sent_on_time_naive)

                # --- Process REPLIES using Conversation History (Existing Logic) --- 
                if subject.lower().startswith("re:") and conversation_id:
                    # Find the message being replied to in the conversation history
                    thread_history = self.conversation_history.get(conversation_id)
                    parent_message = None
                    if thread_history:
                        for msg_in_thread in reversed(thread_history):
                            msg_received_time = msg_in_thread.get('received_time')
                            if msg_received_time and msg_received_time < sent_on_time_naive:
                                parent_message = msg_in_thread
                                break # Found the most recent preceding message
                                
                    if parent_message:
                        parent_received = parent_message.get('received_time')
                        parent_sender = parent_message.get('sender', "Unknown Sender")
                        # Normalize the parent sender
                        parent_sender = self._normalize_contact(parent_sender)
                        # parent_entry_id = parent_message.get('entry_id') # Not currently used
                        
                        logger.debug(f"Reply '{subject}' matched to parent email from '{parent_sender}' received at {parent_received}")
                        try:
                            if isinstance(parent_received, datetime.datetime):
                                response_time_td = sent_on_time_naive - parent_received
                                response_time_hours = response_time_td.total_seconds() / 3600
                                
                                response_data[parent_sender].append({
                                    'response_time': response_time_hours,
                                    'response_length': body_len,
                                    'subject': subject,
                                    'sent_date': sent_on_time_naive 
                                })
                            else:
                                logger.debug(f"Skipping response time calc due to invalid parent time type: {type(parent_received)}")
                        except (TypeError, OverflowError) as dt_err:
                            logger.debug(f"Time calculation error for reply {subject}: {dt_err}")
                    else:
                        logger.debug(f"Could not find preceding message in history for reply: {subject} (ConvID: {conversation_id})")

                # --- Store sent item details for conversation threading analysis (remains the same) ---
                if conversation_id: # Track all sent items in threads
                    conversation_threads[conversation_id].append({
                        'sent_time': sent_on_time_naive,
                        'recipients': all_recipients, # Use all recipients here
                        'subject': subject,
                        'body_length': body_len
                    })

                processed_count += 1
                if processed_count % 500 == 0:
                    logger.info(f"Processed {processed_count}/{num_to_process} sent items...")

            except pywintypes.com_error as ce:
                logger.debug(f"COM error processing sent item index {i+1}: {ce}")
            except Exception as e:
                logger.debug(f"General error processing sent item index {i+1}: {e}", exc_info=True)
            finally:
                # Explicitly release item COM object
                if item is not None:
                    del item
                    item = None

        logger.info(f"Completed initial scan of {processed_count} sent items.")
        # Pass the new initiations data to the calculation function
        self._calculate_contact_importance(response_data, conversation_threads, dict(sender_initiations))
        self._save_sender_scores()

        # Return the collected data (optional)
        return {
            'response_data': dict(response_data),
            'conversation_threads': dict(conversation_threads),
            'sender_initiations': dict(sender_initiations)
        }

    def _calculate_contact_importance(self, response_data, conversation_threads, sender_initiations):
        """
        Calculate sender importance scores based on multiple interaction types:
        - Replies (response time, rate, length) derived from `response_data`.
        - Initiations (emails sent TO sender) derived from `sender_initiations`.
        - Relevance (emails FROM sender read & kept) derived from `self.email_patterns`.
        Updates self.sender_scores.
        """
        logger.info("Calculating contact importance scores based on multiple factors...")
        
        # Retrieve read/kept stats from email_patterns
        sender_read_kept_stats = self.email_patterns.get('sender_read_kept_stats', {})
        
        # Get factors from config
        reply_factor = self.config.get('reply_pattern_score_factor', 1.0)
        initiation_factor = self.config.get('initiation_score_factor', 0.5)
        read_kept_factor = self.config.get('read_kept_score_factor', 0.3)
        min_emails_reply = self.config.get('min_emails_for_pattern', 1) # Min emails for reply analysis component
        # min_total_interactions = self.config.get('min_interactions_for_score', 2) # Optional: Threshold for total interactions
        
        # --- Combine all known senders/contacts --- 
        all_contacts = set(response_data.keys()) | set(sender_initiations.keys()) | set(sender_read_kept_stats.keys())
        
        # Normalize all contacts in the combined set
        normalized_contacts = set(self._normalize_contact(contact) for contact in all_contacts)
        
        new_sender_scores = {}
        raw_scores = [] # To collect raw scores for normalization
        
        # --- Calculate raw score for each contact --- 
        for normalized_contact in normalized_contacts:
            raw_score = 0
            last_interaction_date = None
            interaction_dates = []
            
            # For each normalized contact, find all possible identities
            possible_identities = [normalized_contact]
            # Add any display names that map to this email
            for display_name, email in self.contact_map.items():
                if email == normalized_contact:
                    possible_identities.append(display_name)
            
            # 1. Contribution from Reply Patterns 
            reply_component_score = 0
            reply_count = 0
            
            # Gather all responses across possible identities
            all_responses = []
            for identity in possible_identities:
                if identity in response_data:
                    all_responses.extend(response_data[identity])
            
            if len(all_responses) >= min_emails_reply:
                valid_responses = [r for r in all_responses if r.get('response_time') is not None and r['response_time'] >= 0]
                if valid_responses:
                    reply_count = len(valid_responses)
                    avg_response_time = np.mean([r['response_time'] for r in valid_responses])
                    avg_response_length = np.mean([r['response_length'] for r in valid_responses])
                    response_dates = [r['sent_date'] for r in valid_responses if r.get('sent_date')]
                    interaction_dates.extend(response_dates)
                    
                    response_time_score = 1.0 / (1.0 + max(0, avg_response_time) / 24.0)
                    min_date, max_date = min(response_dates), max(response_dates)
                    date_range_days = max(1, (max_date - min_date).days)
                    response_freq = reply_count / date_range_days

                    reply_component_score = (
                        self.config.get('reply_time_weight', 0.4) * response_time_score +
                        self.config.get('reply_rate_weight', 0.4) * min(1.0, response_freq * 10) +
                        self.config.get('reply_length_weight', 0.2) * min(1.0, max(0, avg_response_length) / 500)
                    )
            raw_score += reply_component_score * reply_factor # Add weighted reply component

            # 2. Contribution from Initiations (Sent TO)
            initiation_count = 0
            initiation_dates = []
            for identity in possible_identities:
                if identity in sender_initiations:
                    initiation_data = sender_initiations[identity]
                    initiation_count += initiation_data['count']
                    initiation_dates.extend(initiation_data['dates'])
            
            raw_score += initiation_count * initiation_factor
            interaction_dates.extend(initiation_dates)
            
            # 3. Contribution from Read & Kept (Received FROM)
            read_kept_count = 0
            read_kept_dates = []
            for identity in possible_identities:
                if identity in sender_read_kept_stats:
                    read_kept_data = sender_read_kept_stats[identity]
                    read_kept_count += read_kept_data['count']
                    if 'dates' in read_kept_data:
                        read_kept_dates.extend(read_kept_data['dates'])
            
            raw_score += read_kept_count * read_kept_factor
            interaction_dates.extend(read_kept_dates)
            
            # Find the latest interaction date across all types
            if interaction_dates:
                # Filter out any non-datetime objects that might have slipped in
                valid_dates = [d for d in interaction_dates if isinstance(d, datetime.datetime)]
                if valid_dates:
                    last_interaction_date = max(valid_dates)
                
            # --- Debug Log --- 
            total_interactions = reply_count + initiation_count + read_kept_count
            logger.debug(f"Contact: '{normalized_contact}' | Interactions: Reply={reply_count}, Initiation={initiation_count}, ReadKept={read_kept_count} | Total={total_interactions} | RawScore={raw_score:.2f}")
            
            # Store results if there was any interaction
            if total_interactions > 0:
                new_sender_scores[normalized_contact] = {
                    'raw_score': raw_score,
                    'reply_count': reply_count,
                    'initiation_count': initiation_count,
                    'read_kept_count': read_kept_count,
                    'total_interactions': total_interactions,
                    'last_interaction': last_interaction_date, # Store the latest date
                    'normalized_score': 0.5, # Default normalized score
                    # Back-compat: some downstream diagnostics expect 'score'.
                    # We will populate it after normalization to maintain compatibility.
                    'score': 0.0,
                    'last_updated': datetime.datetime.now()
                }
                raw_scores.append(raw_score)

        # --- Normalize Scores --- 
        if new_sender_scores and raw_scores:
            max_raw_score = max(raw_scores) if raw_scores else 0
            min_raw_score = min(raw_scores) if raw_scores else 0
            raw_score_range = max_raw_score - min_raw_score

            if raw_score_range > 0:
                for contact, data in new_sender_scores.items():
                    norm_score = (data['raw_score'] - min_raw_score) / raw_score_range
                    # Apply a gentle sigmoid-like curve to spread scores more towards ends? Optional.
                    # norm_score = 1 / (1 + np.exp(- (norm_score - 0.5) * 5)) # Example sigmoid scaling
                    data['normalized_score'] = max(0.0, min(1.0, norm_score))
                    data['score'] = data['normalized_score']  # Alias
            elif len(new_sender_scores) == 1: # Only one sender with a score
                 # Give the single sender a moderate score if range is zero
                 single_sender = list(new_sender_scores.keys())[0]
                 new_sender_scores[single_sender]['normalized_score'] = 0.6 # Or some other default

        # Ensure 'score' present for each sender and refresh timestamp
        for sender, data in new_sender_scores.items():
            data['score'] = data.get('normalized_score', 0.0)
            data['last_updated'] = datetime.datetime.now()

        self.sender_scores = new_sender_scores # Update the instance attribute
        logger.info(f"Calculated importance scores for {len(self.sender_scores)} contacts based on combined factors.")

    def _save_sender_scores(self):
        """ Saves the current sender scores to a pickle file, with guaranteed serializable types. """
        scores_file = self.data_dir / 'sender_scores.pkl'
        try:
            # Create a sanitized copy with guaranteed picklable types
            sanitized_scores = {}
            
            for sender, data in self.sender_scores.items():
                clean_data = {}
                for key, value in data.items():
                    # Clean last_interaction timestamp if present
                    if key == 'last_interaction' and value is not None:
                        if hasattr(value, 'strftime'):
                            try:
                                # Format to string then parse back to ensure standard datetime
                                dt_str = value.strftime('%Y-%m-%d %H:%M:%S')
                                clean_data[key] = datetime.datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
                            except Exception as dt_err:
                                logger.warning(f"Sanitization error for {key} in sender {sender}: {dt_err}. Setting to None.")
                                clean_data[key] = None
                        else:
                            logger.warning(f"Non-datetime value for {key} in sender {sender}: {type(value)}. Setting to None.")
                            clean_data[key] = None
                    else:
                        # Copy other values directly
                        clean_data[key] = value
                        
                sanitized_scores[sender] = clean_data
                
            with open(scores_file, 'wb') as f:
                pickle.dump(sanitized_scores, f)
            logger.info(f"Saved sender scores to {scores_file}")
        except Exception as e:
            logger.error(f"Failed to save sender scores to {scores_file}: {e}")

    # ---------------- Inbox Behavior Analysis ----------------
    def _save_contact_map(self):
        """Save the contact normalization map to a pickle file."""
        contact_map_file = self.data_dir / 'contact_map.pkl'
        try:
            with open(contact_map_file, 'wb') as f:
                pickle.dump(self.contact_map, f)
            logger.info(f"Saved contact map to {contact_map_file} with {len(self.contact_map)} mappings")
        except Exception as e:
            logger.error(f"Error saving contact map to {contact_map_file}: {e}")
            
    def analyze_inbox_behavior(self, max_items=1000):
        """Analyze user interactions with inbox emails (read, delete, move)."""
        logger.info("Analyzing inbox behavior patterns...")
        if not self.inbox:
            logger.warning("Inbox folder not available, skipping inbox behavior analysis.")
            return False

        try:
            # Ensure items are sorted by received time, newest first
            items = self.inbox.Items
            items.Sort("[ReceivedTime]", True)
        except Exception as e:
            logger.warning(f"Could not sort inbox items: {e}. Proceeding without sorting.")
            items = self.inbox.Items

        processed_count = 0
        sender_behavior = defaultdict(lambda: {
            'read_kept': 0, 
            'unread_kept': 0, 
            'deleted': 0, 
            'moved': 0, 
            'total': 0
        })

        # Map for direct sender-to-display-name associations
        name_email_pairs = []

        num_items_to_process = min(items.Count, max_items) if max_items else items.Count
        logger.info(f"Processing up to {num_items_to_process} items from Inbox.")

        for i in range(num_items_to_process):
            try:
                item = items[i + 1]  # 1-based index for Outlook collections
                if item.Class == 43:  # olMail
                    entry_id = safe_get_property(item, 'EntryID')
                    
                    # --- Load previous tracking state for this email --- 
                    previous_tracking_data = self.email_tracking.get(entry_id, {})
                    previous_check_count = previous_tracking_data.get('check_count', 0)
                    # Determine if it was unread *last* time we checked
                    # If 'is_currently_read' was False last time, then it was unread.
                    # Default to True (read) if key is missing, meaning it wasn't unread last time.
                    was_unread_last_run = not previous_tracking_data.get('is_currently_read', True)

                    # Try to get both display name and email address for contact normalization
                    sender_name = safe_get_property(item, 'SenderName')
                    sender_email = safe_get_property(item, 'SenderEmailAddress')
                    
                    # Capture direct name-email mappings as we process emails
                    if sender_name and sender_email and '@' in sender_email:
                        name_email_pairs.append((sender_name.lower(), sender_email.lower()))
                        logger.debug(f"Captured direct name-email pair: '{sender_name}' -> '{sender_email}'")
                    
                    # Use our regular get_sender_address method
                    raw_sender = self._get_sender_address(item)
                    sender = self._normalize_contact(raw_sender)  # Normalize the contact
                    
                    if not sender or not entry_id:
                        continue

                    current_folder = safe_get_property(item, 'Parent.FolderPath')
                    is_read = safe_get_property(item, 'UnRead') == False
                    received_time = self._convert_to_naive_datetime(safe_get_property(item, 'ReceivedTime'))
                    last_modified = self._convert_to_naive_datetime(safe_get_property(item, 'LastModificationTime'))

                    # --- Calculate new check_count based on current read status --- 
                    if not is_read: # Currently unread
                        new_check_count = previous_check_count + 1
                    else: # Currently read
                        new_check_count = 0 # Reset the counter

                    logger.debug(f"  Tracking update for {entry_id}: was_unread_last={was_unread_last_run}, is_read_now={is_read}, prev_count={previous_check_count}, new_count={new_check_count}")

                    # Update tracking data
                    self.email_tracking[entry_id] = {
                        'subject': safe_get_property(item, 'Subject', default="<No Subject>"),
                        'sender': sender,  # Store normalized sender
                        'raw_sender': raw_sender,  # Keep the original sender for reference
                        'sender_name': sender_name,  # Store actual Outlook fields for reference
                        'sender_email': sender_email,
                        'received_time': received_time, # Already converted
                        'last_modified': last_modified, # Already converted
                        'folder_path': current_folder,
                        'is_currently_read': is_read, # Store the *current* read status for the *next* run
                        'check_count': new_check_count, # Store the updated count
                        'last_checked': datetime.datetime.now(timezone.utc) # Use datetime.datetime.now()
                    }

                    # Basic behavior analysis (can be refined)
                    sb = sender_behavior[sender]  # Use normalized sender
                    sb['total'] += 1
                    if current_folder == self.inbox.FolderPath: # Still in Inbox
                        if is_read:
                            sb['read_kept'] += 1
                        else:
                            sb['unread_kept'] += 1
                    elif safe_get_property(item, 'Parent.Name') == 'Deleted Items':
                         sb['deleted'] += 1
                    else: # Moved to another folder
                         sb['moved'] += 1

                    processed_count += 1

            except Exception as e:
                subject_preview = safe_get_property(item, 'Subject', default="<Unknown>")[:50]
                logger.error(f"Error processing inbox item {i+1} ('{subject_preview}...'): {e}", exc_info=False)
                # Optional: log full traceback in debug mode
                if logger.isEnabledFor(logging.DEBUG):
                    logger.debug(f"Traceback for error processing item {i+1}:", exc_info=True)

            if max_items and processed_count >= max_items:
                logger.info(f"Reached processing limit of {max_items} inbox items.")
                break

        logger.info(f"Completed initial scan of {processed_count} inbox items.")
        
        # Update contact_map with direct sender-display pairs
        if name_email_pairs:
            direct_mappings = 0
            for name, email in name_email_pairs:
                if name and email and '@' in email:
                    self.contact_map[name] = email
                    direct_mappings += 1
            
            if direct_mappings > 0:
                logger.info(f"Added {direct_mappings} direct name->email mappings to contact map")
                # Save the updated contact map
                self._save_contact_map()
        
        # Update email_patterns correctly - store the behavior under a specific key
        self.email_patterns['sender_inbox_behavior'] = dict(sender_behavior)
        # Also update the read/kept stats used by scoring
        self.email_patterns['sender_read_kept_stats'] = { 
            sender: {'count': data['read_kept'], 'dates': []} # Store count, dates might be added later if needed
            for sender, data in sender_behavior.items()
        }
        
        # Save the updated data
        self._save_email_tracking()
        self._save_inbox_behavior()
        
        logger.info(f"Inbox behavior analysis complete. Analyzed patterns for {len(self.email_patterns['sender_inbox_behavior'])} senders.")
        return True

    def _save_email_tracking(self):
        """ Saves the current email tracking data to a pickle file, with guaranteed serializable types. """
        tracking_file = self.data_dir / 'email_tracking.pkl'
        try:
            # Create a sanitized copy for safety before pickling
            sanitized_tracking = {}
            
            for entry_id, record in self.email_tracking.items():
                # Create a clean record with simple, picklable types
                clean_record = {}
                for key, value in record.items():
                    # Special handling for datetime fields
                    # Include new fields in sanitization
                    if key in ['received_time', 'last_modified', 'last_checked', 'first_opened_time', 'sent_on_time'] and value is not None:
                        # Convert any datetime-like object to string then back to datetime for absolute safety
                        if hasattr(value, 'strftime'):
                            try:
                                # Format to string then parse back to ensure standard datetime
                                dt_str = value.strftime('%Y-%m-%d %H:%M:%S')
                                clean_record[key] = datetime.datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
                            except Exception as dt_err:
                                logger.warning(f"Sanitization error for {key} in {entry_id}: {dt_err}. Setting to None.")
                                clean_record[key] = None
                        else:
                            logger.warning(f"Non-datetime value for {key} in {entry_id}: {type(value)}. Setting to None.")
                            clean_record[key] = None
                    else:
                        # Copy other values directly
                        clean_record[key] = value
                
                sanitized_tracking[entry_id] = clean_record

            # Pickle using the sanitized copy
            with open(tracking_file, 'wb') as f:
                pickle.dump(sanitized_tracking, f)
            logger.info(f"Saved email tracking data to {tracking_file}")
        
        except Exception as e:
            logger.error(f"Failed to save email tracking data to {tracking_file}: {e}")

    def _save_inbox_behavior(self):
        """ Saves the current inbox behavior patterns to a pickle file. """
        behavior_file = self.data_dir / 'inbox_behavior.pkl'
        try:
            # Ensure nested defaultdicts are converted for pickling if necessary
            # For sender_read_kept_stats, we already converted to dict() when storing
            with open(behavior_file, 'wb') as f:
                pickle.dump(self.email_patterns, f) # Save the whole patterns dict
            logger.info(f"Saved inbox behavior data to {behavior_file}")
        except Exception as e:
             logger.error(f"Failed to save inbox behavior data to {behavior_file}: {e}")


    # ---------------- Folder Structure Analysis ----------------
    def analyze_folder_structure(self):
        """
        Analyze existing folder structure to understand user organization patterns.
        Updates self.folder_structure.
        """
        if not self.namespace:
            logger.error("Namespace not available for folder analysis.")
            return None

        logger.info("Analyzing folder structure...")
        folder_data = {}

        # Use a local import for re if only used here, or ensure it's imported at the top
        import re # Needed for cleaning subject in stats

        # Recursive helper function
        def process_folders_recursive(parent_folder, parent_path=""):
            try:
                folder_name = getattr(parent_folder, 'Name', 'ErrorGettingName')
                # Skip problematic system folders explicitly by name
                # Add other known system/problematic folder names if necessary
                if folder_name in ['Conversation Action Settings', 'Quick Step Settings', 'RSS Feeds', 'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures']:
                     logger.debug(f"Skipping system/problematic folder: {folder_name}")
                     return None

                folder_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
                item_count = parent_folder.Items.Count # This can be slow/error-prone
            except pywintypes.com_error as ce:
                # Log the parent path where the error occurred
                logger.warning(f"COM error accessing basic folder properties under '{parent_path}': {ce}. Skipping.")
                return None
            except Exception as e:
                logger.error(f"Error accessing basic folder info under '{parent_path}': {e}. Skipping.")
                return None

            folder_info = {
                'name': folder_name,
                'path': folder_path,
                'item_count': item_count,
                'unread_count': 0,
                'subfolders': [],
                'stats': {}
            }

            # Get unread count safely
            try:
                filter_unread = "[Unread]=True"
                unread_items = parent_folder.Items.Restrict(filter_unread)
                folder_info['unread_count'] = unread_items.Count
            except Exception as e_uc:
                logger.debug(f"Error getting unread count for '{folder_path}': {e_uc}")

            # Analyze contents (simplified sample)
            if item_count > 0:
                try:
                    items = parent_folder.Items
                    sender_counts = Counter()
                    topic_counts = Counter()
                    date_counts = Counter()
                    read_status = {'read': 0, 'unread': 0}
                    sample_size = min(item_count, 50) # Smaller sample for speed

                    # Sorting might be slow/error-prone, consider removing or sampling differently
                    try:
                         items.Sort("[ReceivedTime]", True)
                    except Exception as sort_err:
                         logger.debug(f"Could not sort items in folder '{folder_path}'. Proceeding without sorting. Error: {sort_err}")


                    for i in range(sample_size):
                        try:
                            # Access item by index (1-based) - needs error handling
                            if (i + 1) > items.Count: break # Stop if index goes out of bounds
                            item = items.Item(i + 1)

                            if not hasattr(item, 'Class') or item.Class != 43: continue

                            # Get Sender Safely
                            sender = "Unknown Sender"
                            try:
                                sender_addr = getattr(item, 'SenderEmailAddress', None)
                                sender_name = getattr(item, 'SenderName', None)
                                if sender_addr and '@' in sender_addr: sender = sender_addr
                                elif sender_name: sender = sender_name
                            except pywintypes.com_error:
                                sender = getattr(item, 'SenderName', "Unknown Sender")
                            except Exception: pass
                            sender_counts[sender] += 1

                            # Get Date Safely
                            rec_time = getattr(item, 'ReceivedTime', None)
                            if rec_time:
                                try:
                                     if isinstance(rec_time, (datetime.datetime, pywintypes.TimeType)):
                                         month_year = rec_time.strftime("%Y-%m")
                                         date_counts[month_year] += 1
                                except (AttributeError, OverflowError, ValueError) as fmt_err:
                                     logger.debug(f"Error formatting date {rec_time}: {fmt_err}")

                            # Get Read Status Safely
                            is_unread = getattr(item, 'Unread', True)
                            read_status['read' if not is_unread else 'unread'] += 1

                            # Get Subject Safely
                            subject = getattr(item, 'Subject', '')
                            if subject:
                                try:
                                     clean_subject = re.sub(r'^(RE|FWD|FW):\s*', '', subject, flags=re.IGNORECASE).strip()
                                     if clean_subject: # Avoid counting empty subjects as topics
                                          topic_counts[clean_subject] += 1
                                except Exception as subj_e:
                                     logger.debug(f"Error cleaning subject '{subject}': {subj_e}")


                        except pywintypes.com_error as item_ce:
                             logger.debug(f"COM error accessing item {i+1} in '{folder_path}': {item_ce}")
                        except IndexError:
                             logger.warning(f"Index error accessing item {i+1} in '{folder_path}'. Collection may have changed.")
                             break # Stop sampling this folder
                        except Exception as item_e:
                            logger.debug(f"Error analyzing item {i+1} in '{folder_path}': {item_e}")

                    folder_info['stats'] = {
                        'top_senders': dict(sender_counts.most_common(5)),
                        'top_topics': dict(topic_counts.most_common(5)),
                        'date_distribution': dict(sorted(date_counts.items(), reverse=True)),
                        'read_status': read_status
                    }
                except pywintypes.com_error as items_ce:
                     logger.debug(f"COM error accessing Items collection in '{folder_path}': {items_ce}")
                except Exception as items_e:
                    logger.debug(f"Error accessing/analyzing items in '{folder_path}': {items_e}")

            # Process subfolders
            try:
                subfolders_collection = parent_folder.Folders
                if subfolders_collection.Count > 0:
                    # Iterate safely by index
                    for i in range(1, subfolders_collection.Count + 1):
                        subfolder_obj = None # Initialize
                        try:
                            subfolder_obj = subfolders_collection.Item(i)
                            # Recursive call
                            subfolder_info = process_folders_recursive(subfolder_obj, folder_path)
                            if subfolder_info: # Append only if successfully processed
                                folder_info['subfolders'].append(subfolder_info)
                        except pywintypes.com_error as ce_sf:
                            sf_name = f"subfolder index {i}"
                            try: sf_name = getattr(subfolder_obj, 'Name', sf_name) # Try to get name for logging
                            except: pass
                            logger.warning(f"COM error accessing {sf_name} of '{folder_path}': {ce_sf}. Skipping.")
                        except Exception as e_sf:
                            sf_name = f"subfolder index {i}"
                            try: sf_name = getattr(subfolder_obj, 'Name', sf_name)
                            except: pass
                            logger.error(f"Error processing {sf_name} of '{folder_path}': {e_sf}. Skipping.")
            except pywintypes.com_error as sf_coll_ce:
                 logger.debug(f"COM error accessing Folders collection of '{folder_path}': {sf_coll_ce}")
            except Exception as sf_e:
                logger.debug(f"Error accessing subfolders collection of '{folder_path}': {sf_e}")

            return folder_info

        # --- Start Processing ---
        root_folders_to_process = []
        try:
            # Prefer the default store's root folder
            default_store = self.namespace.DefaultStore
            root_folder = default_store.GetRootFolder()
            root_folders_to_process.append(('default_store_root', root_folder))
            logger.info(f"Starting folder analysis from root: {getattr(root_folder, 'Name', 'N/A')}")
        except Exception as e_root:
            logger.error(f"Could not get default store root folder: {e_root}. Analysis may be incomplete.")
            # Fallback: Try processing Inbox and Sent Items folders directly if root fails
            if self.inbox:
                root_folders_to_process.append(('inbox', self.inbox))
                logger.info("Adding Inbox to folder analysis roots as fallback.")
            if self.sent_items:
                 root_folders_to_process.append(('sent_items', self.sent_items))
                 logger.info("Adding Sent Items to folder analysis roots as fallback.")

        if not root_folders_to_process:
            logger.error("Could not determine any root folders to start structure analysis.")
            return None

        # Process identified roots
        processed_structures = {}
        for key, folder_obj in root_folders_to_process:
             logger.info(f"Processing folder structure starting from: {getattr(folder_obj, 'Name', 'N/A')} ({key})")
             structure = process_folders_recursive(folder_obj)
             if structure:
                 processed_structures[key] = structure # Store under the root key

        self.folder_structure = processed_structures # Update the main structure
        logger.info(f"Folder structure analysis complete.")
        return self.folder_structure # Return the structure just analyzed

    def _save_folder_structure(self):
        """ Saves the current folder structure to a pickle file. """
        structure_file = self.data_dir / 'folder_structure.pkl'
        try:
            with open(structure_file, 'wb') as f:
                pickle.dump(self.folder_structure, f)
            logger.info(f"Saved folder structure to {structure_file}")
        except Exception as e:
            logger.error(f"Failed to save folder structure to {structure_file}: {e}") 