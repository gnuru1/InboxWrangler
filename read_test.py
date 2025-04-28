import win32com.client
import logging
import time
import pywintypes
import gc # Garbage Collector

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def connect_to_outlook():
    """Connect to Outlook and return the application and namespace objects."""
    try:
        # Try getting active instance first
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            logger.info("Connected to active Outlook instance.")
        except:
            logger.info("No active Outlook instance found, launching new one...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            
        namespace = outlook.GetNamespace("MAPI")
        logger.info(f"Connected to Outlook namespace successfully")
        return outlook, namespace
    except Exception as e:
        logger.error(f"Failed to connect to Outlook: {e}")
        return None, None

def check_read_status():
    """Check read/unread status of ALL emails in inbox directly with more refinements."""
    outlook, namespace = connect_to_outlook()
    if not outlook or not namespace:
        logger.error("Failed to connect to Outlook")
        return
    
    inbox = None
    items = None
    item = None # Ensure item is defined in outer scope for finally block
    initial_item_count = 0 # Track initial count
    
    try:
        # Get inbox folder
        try:
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            initial_item_count = inbox.Items.Count # Store initial count
            logger.info(f"Retrieved inbox '{inbox.Name}' with {initial_item_count} items (Initial Count)")
        except Exception as e:
            logger.error(f"Failed to get Inbox folder: {e}")
            return

        # Track read/unread counts
        read_count = 0
        unread_count = 0
        read_emails = []
        problem_items = [] # Track items with unexpected status
        item_types_found = {} # Track item classes found
        
        # Get a new sorted collection
        try:
            items = inbox.Items
            # Attempt sorting, but don't fail if it doesn't work for mixed types
            try: 
                items.Sort("[ReceivedTime]", True)  # Sort newest first
                logger.info("Sorted items collection retrieved.")
            except Exception as sort_err:
                 logger.warning(f"Could not sort items collection (might contain non-sortable types): {sort_err}")
            
            # Add a small delay - might help with COM object state updates
            time.sleep(0.5) 
        except Exception as e:
            logger.error(f"Failed to get items collection: {e}")
            return

        # Print detailed info for ALL items
        actual_items_in_collection = items.Count # Get count after potential sort attempt
        logger.info(f"Examining all {actual_items_in_collection} items in the collection:")
        processed_count = 0
        for i in range(actual_items_in_collection):
            item = None # Reset item for each loop
            try:
                # Check index validity again just in case count changes mid-loop
                if (i + 1) > items.Count:
                     logger.warning(f"Item count changed during loop, stopping at index {i}.")
                     break
                     
                item = items.Item(i + 1)  # 1-based index
                
                # --- Log Item Type BEFORE Filtering ---
                item_class = getattr(item, 'Class', 0)
                message_class = getattr(item, 'MessageClass', 'N/A')
                item_types_found[item_class] = item_types_found.get(item_class, 0) + 1
                logger.debug(f"Processing Item {i+1}: Class={item_class}, MessageClass='{message_class}'")
                
                # --- Removed MailItem Class Filter --- 
                # if item_class != 43:  # 43 = olMail
                #    logger.debug(f"Skipping item {i+1}, not a MailItem (Class={item_class})")
                #    continue
                    
                # --- Refined Property Access (Only for Mail Items) ---
                subject = "(N/A for non-mail)"
                sender = "(N/A for non-mail)"
                entry_id = "(Error getting EntryID)"
                is_unread = None # Default to None
                
                # 1. Try accessing EntryID (usually available)
                try: entry_id = item.EntryID
                except: pass # Ignore errors for this test
                
                # Only check Unread, Subject, Sender for MailItems (Class 43)
                if item_class == 43:
                    # 2. Try accessing LastModificationTime 
                    try: item.LastModificationTime
                    except: pass # Ignore errors for this test

                    # 3. Now try accessing Unread
                    try: 
                        is_unread = item.Unread 
                    except Exception as e_unread:
                         logger.warning(f"Item {i+1} ({entry_id}): Error accessing Unread property: {e_unread}. Assuming unread.")
                         is_unread = True # Default to True on error
                    
                    # Get other details for logging 
                    try: subject = item.Subject if item.Subject else "(No Subject)"
                    except: subject = "(Error Subject)" # Ignore errors for logging details
                    try: sender = item.SenderName if item.SenderName else "(No Sender)"
                    except: sender = "(Error Sender)" # Ignore errors for logging details
                else:
                    # For non-mail items, try getting a subject if it exists, otherwise mark N/A
                    try: subject = getattr(item, 'Subject', f"(Non-Mail Item: {message_class})")
                    except: subject = f"(Non-Mail Item: {message_class})"
                    is_unread = None # Cannot determine read status reliably for non-mail items this way

                # Print details
                status = "UNREAD" if is_unread else ("READ" if is_unread is False else "N/A")
                logger.info(f"Item {i+1}/{actual_items_in_collection}: [{status}] Class={item_class} | From: {sender} | Subject: {subject}") 
                
                # Count status ONLY for MailItems where Unread was determined
                if item_class == 43 and is_unread is not None:
                    if is_unread:
                        unread_count += 1
                    else:
                        read_count += 1
                        read_emails.append(f"{sender}: {subject}")
                
                processed_count += 1
                
            except pywintypes.com_error as ce_loop:
                 logger.error(f"Outer COM error processing item index {i+1}: {ce_loop}")
                 problem_items.append({'index': i+1, 'error': str(ce_loop)}) 
            except Exception as e_loop:
                logger.error(f"Outer general error processing item {i+1}: {e_loop}")
                problem_items.append({'index': i+1, 'error': str(e_loop), 'subject': getattr(item, 'Subject', 'N/A') if item else 'N/A'})
            finally:
                # Explicitly release the item object in each loop iteration
                if item is not None:
                    del item
                    item = None
                
    finally:
        # --- Final Summary & Cleanup ---
        logger.info(f"\n--- SUMMARY ({processed_count} items processed, Initial Count was {initial_item_count}) ---")
        
        # Log Item Types Found
        logger.info("Item Class Distribution:")
        for item_class, count in item_types_found.items():
             logger.info(f"  - Class {item_class}: {count} items")
        
        logger.info(f"Found {read_count} read and {unread_count} unread MailItems (Class 43)")
        
        if read_count < 8:
            logger.warning(f"WARNING: Expected at least 8 read emails, but only found {read_count}.")
        
        # Print all read emails found
        if read_emails:
            logger.info(f"\n--- READ EMAILS FOUND --- ({read_count}):")
            for i, email in enumerate(read_emails, 1):
                logger.info(f"{i}. {email}")
        else:
            logger.info("No read emails found")
            
        # Report problems
        if problem_items:
            logger.error(f"\n--- PROBLEMS ENCOUNTERED DURING PROCESSING --- ({len(problem_items)}):")
            for problem in problem_items:
                logger.error(f"  - Index {problem['index']}: {problem['error']}")
                if 'subject' in problem: logger.error(f"    Subject context: {problem['subject']}")
                
        # Explicitly release COM objects
        logger.info("Releasing COM objects...")
        if items is not None: del items
        if inbox is not None: del inbox
        if namespace is not None: del namespace
        if outlook is not None: del outlook
        # Trigger garbage collection
        gc.collect()
        logger.info("COM objects released.")

if __name__ == "__main__":
    check_read_status() 