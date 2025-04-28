import logging
import win32com.client
import pywintypes
import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

# Outlook folder constants
olFolderInbox = 6
olFolderSentMail = 5
olFolderArchive = 23

def connect_to_outlook():
    """
    Establish connection to Outlook application and return namespace.
    Tries to connect to an active instance first.

    Returns:
        tuple: (outlook_app, namespace) or (None, None) on failure.
    """
    try:
        logger.info("Connecting to Outlook...")
        # Try getting active instance first
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            logger.info("Connected to active Outlook instance.")
        except pywintypes.com_error: # More specific exception check
            logger.info("No active Outlook instance found, launching new one...")
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as get_active_err: # Catch other potential errors
             logger.warning(f"Error trying GetActiveObject: {get_active_err}. Attempting Dispatch...")
             outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")

        # Log account information
        try:
            accounts = namespace.Accounts
            account_info = []
            for i in range(1, accounts.Count + 1):
                account = accounts.Item(i)
                account_info.append(f"{getattr(account, 'DisplayName', 'N/A')} ({getattr(account, 'SmtpAddress', 'N/A')})")
            logger.info(f"Connected to Outlook with {len(account_info)} accounts: {', '.join(account_info)}")
        except Exception as acc_e:
            logger.warning(f"Could not retrieve Outlook account details: {acc_e}")

        return outlook, namespace

    except pywintypes.com_error as ce:
        logger.error(f"Failed to connect to Outlook (COM Error): {ce}")
        logger.error("Ensure Outlook is running and not in a modal dialog.")
        return None, None
    except Exception as e:
        logger.error(f"Failed to connect to Outlook: {e}", exc_info=True)
        return None, None

def get_default_folder(namespace, folder_id):
    """
    Safely get a default folder from the namespace.

    Args:
        namespace: Outlook MAPI namespace object.
        folder_id (int): The constant representing the folder (e.g., olFolderInbox).

    Returns:
        MAPIFolder object or None if not found.
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
        logger.error(f"Error getting default folder ID {folder_id}: {e}", exc_info=True)
        return None

def get_folder_by_path(namespace, folder_path_str, store_name=None):
    """
    Finds an Outlook folder by its full path string (e.g., "Inbox/Subfolder").

    Args:
        namespace: Outlook MAPI namespace object.
        folder_path_str (str): The folder path separated by '/'.
        store_name (str, optional): Specify the store name if not the default store.
                                     Defaults to None (uses default store).

    Returns:
        MAPIFolder object or None if not found.
    """
    if not namespace:
        logger.error("Cannot get folder by path: Namespace is None.")
        return None

    try:
        if store_name:
            # Find specific store
            target_store = None
            for store in namespace.Stores:
                if store.DisplayName == store_name:
                    target_store = store
                    break
            if not target_store:
                logger.error(f"Store '{store_name}' not found.")
                return None
            current_folder = target_store.GetRootFolder()
        else:
            # Use default store's root folder (usually parent of Inbox)
            current_folder = namespace.GetDefaultFolder(olFolderInbox).Parent

        path_parts = folder_path_str.strip('/').split('/')

        for part in path_parts:
            if not part: continue
            try:
                current_folder = current_folder.Folders(part)
            except pywintypes.com_error:
                logger.warning(f"Folder part '{part}' not found in path '{folder_path_str}'")
                return None # Folder path does not exist

        logger.debug(f"Successfully found folder: {current_folder.FolderPath}")
        return current_folder

    except pywintypes.com_error as ce:
        logger.error(f"COM Error getting folder by path '{folder_path_str}': {ce}")
        return None
    except Exception as e:
        logger.error(f"Error getting folder by path '{folder_path_str}': {e}", exc_info=True)
        return None

def safe_get_property(item, property_name, default=None):
    """
    Safely gets a property from an Outlook item, handling COM errors.

    Args:
        item: The Outlook item (e.g., MailItem).
        property_name (str): The name of the property to access.
        default: The value to return if the property is inaccessible or doesn't exist.

    Returns:
        The property value or the default.
    """
    try:
        return getattr(item, property_name, default)
    except pywintypes.com_error as ce:
        # Common error: Property is invalid for this object type or restricted
        logger.debug(f"COM error accessing property '{property_name}' on item {getattr(item, 'Subject', '?')}: {ce}")
        return default
    except Exception as e:
        logger.warning(f"Unexpected error accessing property '{property_name}': {e}")
        return default

def get_or_create_folder(parent_folder, folder_name):
    """
    Gets a subfolder by name. If it doesn't exist, it creates it.

    Args:
        parent_folder: The parent MAPIFolder object.
        folder_name (str): The name of the subfolder to get or create.

    Returns:
        MAPIFolder object or None if creation/retrieval fails.
    """
    try:
        # Check if the folder already exists
        subfolder = parent_folder.Folders(folder_name)
        logger.debug(f"Folder '{folder_name}' already exists under '{parent_folder.Name}'.")
        return subfolder
    except pywintypes.com_error as ce:
        # Specific check if error is due to folder not found
        if "0x80070003" in str(ce): # Error code for "Path not found" might vary
            logger.info(f"Folder '{folder_name}' not found under '{parent_folder.Name}'. Creating it...")
            try:
                subfolder = parent_folder.Folders.Add(folder_name)
                logger.info(f"Successfully created folder: {subfolder.FolderPath}")
                return subfolder
            except pywintypes.com_error as ce_add:
                logger.error(f"COM Error creating folder '{folder_name}' under '{parent_folder.Name}': {ce_add}")
                return None
            except Exception as e_add:
                logger.error(f"Error creating folder '{folder_name}' under '{parent_folder.Name}': {e_add}")
                return None
        else:
            # Log other COM errors when trying to access the folder
            logger.error(f"COM Error accessing folder '{folder_name}' under '{parent_folder.Name}': {ce}")
            return None
    except Exception as e:
        logger.error(f"Error getting or creating folder '{folder_name}' under '{parent_folder.Name}': {e}")
        return None

def create_task_from_email(email_item, reminder_days=None):
    """
    Creates an Outlook Task based on an email item.

    Args:
        email_item: The Outlook MailItem object.
        reminder_days (int, optional): Days from now to set a reminder. Defaults to None (no reminder).

    Returns:
        True if the task was created successfully, False otherwise.
    """
    try:
        if not email_item:
            logger.warning("Cannot create task: email_item is None.")
            return False

        outlook = win32com.client.Dispatch("Outlook.Application")
        task = outlook.CreateItem(3) # 3 = olTaskItem

        # Set task properties based on email
        task.Subject = f"Task from Email: {safe_get_property(email_item, 'Subject', 'No Subject')}"
        task.Body = f"Original Email Subject: {safe_get_property(email_item, 'Subject', 'No Subject')}\n" \
                    f"Received: {safe_get_property(email_item, 'ReceivedTime', 'N/A')}\n" \
                    f"From: {safe_get_property(email_item, 'SenderName', 'N/A')} ({safe_get_property(email_item, 'SenderEmailAddress', 'N/A')})\n\n" \
                    f"---\n{safe_get_property(email_item, 'Body', '')}"

        task.StartDate = datetime.datetime.now()
        task.DueDate = datetime.datetime.now() + datetime.timedelta(days=1) # Default due tomorrow

        if reminder_days is not None and isinstance(reminder_days, (int, float)) and reminder_days > 0:
            task.ReminderSet = True
            reminder_time = datetime.datetime.now() + datetime.timedelta(days=reminder_days)
            # Outlook typically expects time in local timezone format
            # Ensure it's a datetime object before assigning
            if isinstance(reminder_time, datetime.datetime):
                 task.ReminderTime = reminder_time # May need timezone adjustment depending on system config
            else:
                 logger.warning(f"Could not set reminder time; invalid type: {type(reminder_time)}")
                 task.ReminderSet = False


        # Attach the original email to the task
        # 0 = olByValue, 5 = olEmbeddeditem
        attachment_type = 5
        attachment_name = f"Original Email - {safe_get_property(email_item, 'Subject', 'email')[:50]}.msg"
        try:
            task.Attachments.Add(email_item, attachment_type, 1, attachment_name)
        except pywintypes.com_error as attach_err:
             logger.warning(f"Could not attach original email to task (COM Error): {attach_err}")
             # Fallback: Try adding as plain text
             try:
                 temp_file = Path(f"./temp_email_{email_item.EntryID}.txt")
                 with open(temp_file, "w", encoding="utf-8") as f:
                     f.write(task.Body) # Write the body we already constructed
                 task.Attachments.Add(str(temp_file.resolve()))
                 temp_file.unlink() # Clean up temp file
             except Exception as txt_attach_err:
                 logger.error(f"Could not attach email body as text file either: {txt_attach_err}")
        except Exception as attach_err_gen:
             logger.error(f"Could not attach original email to task: {attach_err_gen}")


        task.Save()
        logger.info(f"Created task '{task.Subject}' from email.")
        return True

    except pywintypes.com_error as ce:
        logger.error(f"COM Error creating task from email: {ce}")
        return False
    except Exception as e:
        logger.error(f"Error creating task from email: {e}", exc_info=True)
        return False 