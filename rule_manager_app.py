"""
Streamlit application for creating Outlook rules based on sender statistics.
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import logging
import pywintypes

# Project modules
from outlook_utils import connect_to_outlook, get_default_folder, get_folder_by_path, create_sender_rule # Assuming create_sender_rule will be added

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('rule_manager_app')
st.set_page_config(layout="wide")

# --- Streamlit App UI --- 
st.title("Outlook Rule Manager from Sender Stats")

st.write("""
1.  Run `python inbox_sender_stats.py --limit <num> --output sender_stats.csv` to generate the analysis file.
2.  Upload the generated `sender_stats.csv` file below.
3.  Select senders and choose an action to create Outlook rules.
""")

uploaded_file = st.file_uploader("Upload Sender Stats CSV", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        df = df.fillna('') # Replace NaN with empty strings for display/processing

        st.info(f"Loaded {len(df)} senders from {uploaded_file.name}")

        # Add a selection column using data_editor
        df["Select"] = False
        # Define column configuration for the editor
        column_config = {
            "Select": st.column_config.CheckboxColumn("Select Sender?", default=False),
            "sender": st.column_config.TextColumn("Sender", width="large"),
            "total_emails": st.column_config.NumberColumn("Total Emails", format="%d"),
            "unread_percent": st.column_config.NumberColumn("% Unread", format="%.1f%%"),
            "subject_similarity_percent": st.column_config.NumberColumn("% Fuzzy Subject", format="%.1f%%")
        }

        st.write("### Sender Statistics:")
        edited_df = st.data_editor(
            df,
            column_order=("Select", "sender", "total_emails", "unread_percent", "subject_similarity_percent"),
            column_config=column_config,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic" # Allow dynamic rows, though loading from file sets initial rows
        )

        selected_senders_df = edited_df[edited_df["Select"]]
        selected_sender_list = selected_senders_df['sender'].tolist()

        if not selected_sender_list:
            st.warning("Select one or more senders using the checkboxes above to create rules.")
        else:
            st.write("### Create Rules for Selected Senders:")
            st.write(f"Selected Senders: {len(selected_sender_list)}")
            # Optionally display selected senders
            with st.expander("Show selected senders"): 
                 st.dataframe(selected_senders_df[[col for col in selected_senders_df.columns if col != 'Select']], hide_index=True, use_container_width=True)

            action = st.radio(
                "Choose Action:",
                ("Move to Folder", "Delete Permanently", "Mark as Read"),
                key='action_choice'
            )

            target_folder_path = ""
            if action == "Move to Folder":
                target_folder_path = st.text_input(
                    "Target Folder Path (relative to Inbox or full path):",
                    placeholder="e.g., Archive/Newsletters or Junk Email"
                )

            rule_name_prefix = st.text_input("Rule Name Prefix:", value="AutoRule:")

            st.warning("🚨 **Warning:** Creating rules will modify your Outlook settings.", icon="⚠️")

            if st.button("Create Rules", key='create_button'):
                if action == "Move to Folder" and not target_folder_path:
                    st.error("Please specify a Target Folder Path for the 'Move' action.")
                else:
                    # --- Rule Creation Logic ---
                    created_count = 0
                    error_count = 0
                    outlook = None
                    namespace = None
                    rules = None
                    needs_save = False

                    st.write("Attempting to create rules...")
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    try:
                        outlook, namespace = connect_to_outlook()
                        if not namespace:
                            st.error("Failed to connect to Outlook. Ensure Outlook is running.")
                        else:
                            rules = namespace.DefaultStore.GetRules()
                            total_selected = len(selected_sender_list)
                            
                            for i, sender in enumerate(selected_sender_list):
                                status_text.text(f"Processing sender {i+1}/{total_selected}: {sender}")
                                rule_name = f"{rule_name_prefix} {sender[:40]}" # Truncate long names
                                
                                try:
                                    success = create_sender_rule(
                                        namespace=namespace,
                                        rules=rules, # Pass rules object
                                        sender_identifier=sender,
                                        action_type=action,
                                        target_folder_path=target_folder_path if action == "Move to Folder" else None,
                                        rule_name=rule_name
                                    )
                                    if success:
                                        created_count += 1
                                        needs_save = True # Mark that we need to save changes
                                        logger.info(f"Rule '{rule_name}' created in memory for sender '{sender}'")
                                    else:
                                        error_count += 1
                                        logger.warning(f"Failed to create rule object for sender '{sender}'")
                                except Exception as rule_exc:
                                    error_count += 1
                                    logger.error(f"Error creating rule for sender '{sender}': {rule_exc}", exc_info=True)
                                    st.warning(f"Error creating rule for '{sender}': {rule_exc}")
                                
                                progress_bar.progress((i + 1) / total_selected)
                            
                            # Save rules *once* after loop if any rules were created
                            if needs_save:
                                status_text.text("Saving rules to Outlook...")
                                rules.Save()
                                logger.info(f"Successfully saved {created_count} rules to Outlook.")
                                st.success(f"Successfully created and saved {created_count} rules!")
                            elif created_count == 0 and error_count == 0:
                                 status_text.text("No rules needed creating (e.g., might exist already).") # Or check if success was False above
                            
                            if error_count > 0:
                                st.error(f"Failed to create {error_count} rules. Check logs for details.")
                            status_text.text("Rule creation process finished.")
                                
                    except pywintypes.com_error as ce:
                        st.error(f"Outlook COM Error: {ce}. Ensure Outlook is running and accessible.")
                        logger.error(f"Outlook COM Error during rule creation: {ce}", exc_info=True)
                    except Exception as e:
                        st.error(f"An unexpected error occurred: {e}")
                        logger.error(f"Unexpected error during rule creation: {e}", exc_info=True)
                    finally:
                        # Release COM objects (optional, Python usually handles it)
                        rules = None
                        namespace = None
                        outlook = None

    except pd.errors.EmptyDataError:
        st.error("The uploaded CSV file is empty or invalid.")
    except Exception as e:
        st.error(f"An error occurred processing the file: {e}")
        logger.error(f"Error processing uploaded file: {e}", exc_info=True)
else:
    st.info("Upload the CSV file generated by `inbox_sender_stats.py` to begin.") 