# ===================================================================
#
#  Teams Message Automation Script (Python Conversion - Auto Launch & Focus)
#
#  Description: This script automates sending a message to a specific
#               chat in Microsoft Teams using pywinauto for UI Automation.
#               Now auto-launches and focuses Teams if not running.
#
#  Instructions:
#  1. Open Microsoft Teams (or let the script do it for you).
#  2. Run this script.
#
# ===================================================================

import time
import os
import pywinauto
from pywinauto.application import Application
from pywinauto.findwindows import find_elements, ElementNotFoundError

# --- Configuration ---
chat_name = "*Johnny DoeBoy*"  # Using partial regex match for robustness (e.g., handles status changes)
message_box_name = "*message*"  # Partial match
send_button_name = "*Send*"  # Partial match
message_to_send = "This is a message sent automatically via a Python script. Have a great day!"
search_depth = 20  # Depth for UI tree search; adjust as needed
launch_wait = 10  # Seconds to wait after launching Teams (adjust if needed)

print("Starting Teams Automation Script...")

# --- Get or Launch the Main Teams Window ---
teams_window = None
try:
    # Try to find and connect to existing Teams window
    elements = find_elements(title_re=".*Microsoft Teams.*", backend="uia")
    if elements:
        app = Application(backend="uia").connect(handle=elements[0].handle)
        teams_window = app.top_window()
        print(f"Found existing Teams window: {teams_window.window_text()}")
    else:
        raise ElementNotFoundError("No Teams window found.")
except Exception:
    # Launch Teams if not found
    print("Teams not running. Launching it...")
    teams_exe = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Teams\Update.exe")
    launch_cmd = f'"{teams_exe}" --processStart "Teams.exe"'
    app = Application(backend="uia").start(launch_cmd)
    time.sleep(launch_wait)  # Wait for Teams to start up
    teams_window = app.top_window()
    print("Launched Teams successfully.")

# Bring the window to the foreground
if teams_window:
    teams_window.set_focus()
    print("Focused Teams window.")
else:
    print("Error: Could not get Teams window.")
    exit()

# --- Step 1: Find and Select the Chat ---
try:
    print(f"Searching for chat matching: '{chat_name}'...")
    # Search descendants for TreeItem (chat); use title_re for regex match
    chat_elements = teams_window.descendants(title_re=chat_name, control_type="TreeItem", depth=search_depth)
    if not chat_elements:
        # Fallback: Search by control_type only and filter
        print("Chat not found by name. Trying fallback search...")
        chat_elements = teams_window.descendants(control_type="TreeItem", depth=search_depth)
        chat_elements = [el for el in chat_elements if "Johnny DoeBoy" in el.window_text()]
    if not chat_elements:
        raise Exception(f"Could not find chat matching '{chat_name}'. Check if visible.")
    
    group_chat = chat_elements[0]  # Take first match
    group_chat.select()  # Select the chat
    print("Successfully selected chat.")
    time.sleep(2)  # Wait for chat pane to load
except Exception as e:
    print(f"CRITICAL: Could not find or select the chat. Error: {e}")
    exit()

# --- Step 2: Find the Message Box and Type a Message ---
try:
    print(f"Searching for message box matching: '{message_box_name}'...")
    # Search for Edit control by title_re
    message_inputs = teams_window.descendants(title_re=message_box_name, control_type="Edit", depth=search_depth)
    if not message_inputs:
        # Fallback: Last Edit control (often the message box)
        print("Message box not found by name. Trying fallback...")
        message_inputs = teams_window.descendants(control_type="Edit", depth=search_depth)
        message_inputs = message_inputs[-1:] if message_inputs else []
    if not message_inputs:
        raise Exception(f"Could not find message input matching '{message_box_name}'.")
    
    message_input = message_inputs[0]
    message_input.set_edit_text(message_to_send)
    print("Successfully entered text into the message box.")
except Exception as e:
    print(f"CRITICAL: Could not find or set text in message input. Error: {e}")
    exit()

# --- Step 3: Find and Click the Send Button ---
try:
    print(f"Searching for send button matching: '{send_button_name}'...")
    # Search for Button by title_re
    send_buttons = teams_window.descendants(title_re=send_button_name, control_type="Button", depth=search_depth)
    if not send_buttons:
        # Fallback: Buttons containing "Send"
        print("Send button not found by name. Trying fallback...")
        send_buttons = teams_window.descendants(control_type="Button", depth=search_depth)
        send_buttons = [btn for btn in send_buttons if "Send" in btn.window_text()]
    if not send_buttons:
        raise Exception(f"Could not find Send button matching '{send_button_name}'.")
    
    send_button = send_buttons[0]
    send_button.click()
    print("SUCCESS: Message sent!")
except Exception as e:
    print(f"CRITICAL: Could not find or invoke Send button. Error: {e}")

print("Script finished.")