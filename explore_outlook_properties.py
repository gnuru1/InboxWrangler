import logging
import json
import datetime
import win32com.client
import pywintypes
from collections import defaultdict
from pathlib import Path

# Assuming outlook_utils is in the same directory or accessible via PYTHONPATH
try:
    from outlook_utils import connect_to_outlook, get_default_folder, olFolderInbox
except ImportError:
    logging.error("Could not import outlook_utils. Make sure it's in the same directory or accessible.")
    exit()

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("outlook_explorer")

# --- Global dictionary to store property stats ---
property_stats = defaultdict(lambda: {"count": 0, "types": set(), "values": set()})
MAX_SAMPLE_VALUES = 5 # Limit the number of unique sample values stored

def safe_get_attr(obj, attr_name, path_prefix=""):
    """Safely get an attribute, logging errors."""
    full_path = f"{path_prefix}.{attr_name}" if path_prefix else attr_name
    try:
        # Handle potential issues with accessing properties that are methods without calling them
        attr = getattr(obj, attr_name)
        if callable(attr) and not isinstance(attr, (str, int, float, bool, datetime.datetime, pywintypes.TimeType)):
             # For methods, we just note its existence, don't call it. Except for known simple getters maybe? No, keep it simple.
             return f"<method {attr_name}>"
        return attr
    except (pywintypes.com_error, AttributeError) as ce:
        # COM errors or attributes not existing are common, log lightly
        logger.debug(f"Could not access property: {full_path} - Error: {ce}")
        return None
    except Exception as e:
        # Log other unexpected errors more prominently
        logger.warning(f"Unexpected error accessing property: {full_path} - Error: {e}")
        return None

def format_sample_value(value):
    """Format value for JSON output."""
    if isinstance(value, (datetime.datetime, pywintypes.TimeType)):
        # Consistent string format for dates
        try:
            # Handle potential timezone info if present
            if hasattr(value, 'tzinfo') and value.tzinfo:
                 # Convert to UTC for consistency before formatting
                 utc_dt = value.astimezone(datetime.timezone.utc)
                 return utc_dt.strftime('%Y-%m-%d %H:%M:%S %Z')
            else:
                 # Naive datetime formatting
                 return value.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
             return str(value) # Fallback
    elif isinstance(value, str):
        # Truncate long strings
        return value[:100] + "..." if len(value) > 100 else value
    elif isinstance(value, memoryview):
        return f"<memoryview size {value.nbytes}>"
    elif isinstance(value, win32com.client.CDispatch):
        # Represent COM objects simply
        # Try to get a Name or Subject property if available for context
        name = safe_get_attr(value, 'Name')
        subject = safe_get_attr(value, 'Subject')
        if name: return f"<COMObject Name='{format_sample_value(name)}'>"
        if subject: return f"<COMObject Subject='{format_sample_value(subject)}'>"
        return "<COMObject>"
    elif isinstance(value, bytes):
         try:
            return value.decode('utf-8', errors='replace')[:100] + "..." if len(value) > 100 else value.decode('utf-8', errors='replace')
         except:
             return f"<bytes length {len(value)}>"
    # Ensure other types are stringified
    return str(value)

def explore_object(obj, path_prefix, max_depth=2, current_depth=0):
    """Recursively explore object properties."""
    if current_depth >= max_depth:
        logger.debug(f"Max depth reached exploring {path_prefix}")
        return

    # Check if obj itself is None or not an object we can dir()
    if obj is None or not hasattr(obj, '__dict__') and not isinstance(obj, win32com.client.CDispatch):
        return

    attributes = []
    try:
        attributes = dir(obj)
    except Exception as e:
        logger.warning(f"Could not dir() object at {path_prefix}: {e}")
        return

    for attr_name in attributes:
        if attr_name.startswith("_"):
            continue # Skip private/internal attributes

        val = safe_get_attr(obj, attr_name, path_prefix)
        # Don't record if access failed or returned None/empty string often
        if val is None or val == "":
            # Still explore if it's a COM object even if primary access returns None/empty
             if not isinstance(val, win32com.client.CDispatch):
                  continue


        full_path = f"{path_prefix}.{attr_name}"
        val_type = type(val)

        # --- Update Stats ---
        stats = property_stats[full_path]
        stats["count"] += 1
        stats["types"].add(str(val_type))
        if len(stats["values"]) < MAX_SAMPLE_VALUES:
            stats["values"].add(format_sample_value(val))

        # --- Handle Collections ---
        # Simple heuristic: ends with 's', is COM object, might have Count/Item
        if attr_name.lower().endswith('s') and isinstance(val, win32com.client.CDispatch):
            count = safe_get_attr(val, 'Count', full_path)
            if isinstance(count, int):
                count_path = f"{full_path}.Count"
                count_stats = property_stats[count_path]
                count_stats["count"] += 1
                count_stats["types"].add(str(type(count)))
                if len(count_stats["values"]) < MAX_SAMPLE_VALUES:
                     count_stats["values"].add(format_sample_value(count))

                # If collection has items, explore the first item
                if count > 0:
                    try:
                        first_item = val.Item(1) # 1-based index for COM collections
                        explore_object(first_item, f"{full_path}[0]", max_depth, current_depth + 1)
                    except Exception as item_e:
                        logger.debug(f"Could not access Item(1) for {full_path}: {item_e}")

        # --- Handle Specific Nested Objects ---
        elif attr_name in ['Session', 'Sender', 'Parent'] and isinstance(val, win32com.client.CDispatch):
             explore_object(val, full_path, max_depth, current_depth + 1)
        # Special case for Session.CurrentUser
        elif attr_name == 'Session' and isinstance(val, win32com.client.CDispatch):
             current_user = safe_get_attr(val, 'CurrentUser', full_path)
             if current_user:
                 explore_object(current_user, f"{full_path}.CurrentUser", max_depth, current_depth + 1)

def scan_inbox_properties(limit=50, output_file="inbox_properties_detailed.json", max_depth=2):
    """Main function to scan inbox and explore properties."""
    # Clear global stats for fresh run
    property_stats.clear()

    outlook, namespace = connect_to_outlook()
    if not namespace:
        logger.error("Could not connect to Outlook namespace.")
        return False

    inbox = get_default_folder(namespace, olFolderInbox)
    if not inbox:
        logger.error("Could not access Inbox folder.")
        return False

    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True) # Sort newest first
    except Exception as e:
        logger.error(f"Error accessing or sorting inbox items: {e}")
        return False

    total = items.Count
    limit = min(limit, total)
    logger.info(f"Scanning properties of {limit} messages out of {total} in Inbox (max_depth={max_depth})...")

    processed_count = 0
    for i in range(limit):
        item = None
        try:
            item = items.Item(i + 1)
            # Only handle mail items (Class = 43)
            item_class = safe_get_attr(item, "Class")
            if item_class != 43:
                logger.debug(f"Skipping item {i+1}, not a MailItem (Class={item_class})")
                continue

            # Start exploration from the root 'item'
            explore_object(item, "item", max_depth=max_depth, current_depth=0)
            processed_count += 1
            if processed_count % 20 == 0:
                 logger.info(f"Processed {processed_count}/{limit} items...")

        except pywintypes.com_error as ce:
            logger.warning(f"COM error reading item index {i+1}: {ce}")
        except Exception as e:
            logger.error(f"General error reading item index {i+1}: {e}", exc_info=True) # Log tracebacks for unexpected errors

    logger.info(f"Finished scanning {processed_count} items.")

    # --- Prepare output ---
    output_data = {}
    # Sort properties alphabetically for readability
    sorted_props = sorted(property_stats.keys())

    for prop_path in sorted_props:
        stats = property_stats[prop_path]
        output_data[prop_path] = {
            "count": stats["count"],
            # Convert sets to sorted lists for JSON
            "types": sorted(list(stats["types"])),
            "sample_values": sorted(list(stats["values"]))
        }

    # --- Save output ---
    try:
        out_path = Path(output_file)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(output_data, f, indent=2)
        logger.info(f"Saved property catalog to {out_path.resolve()}")
    except Exception as e:
        logger.error(f"Failed to write output file '{output_file}': {e}")
        return False

    logger.info(f"Cataloged {len(output_data)} unique property paths.")
    return True

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Scan Outlook inbox messages and catalog available properties recursively.")
    parser.add_argument("--limit", type=int, default=50, help="Maximum number of messages to scan.")
    parser.add_argument("--depth", type=int, default=2, help="Maximum recursion depth for exploring nested objects.")
    parser.add_argument("--output", type=str, default="inbox_properties_detailed.json", help="Output JSON file path.")
    args = parser.parse_args()

    scan_inbox_properties(limit=args.limit, output_file=args.output, max_depth=args.depth) 