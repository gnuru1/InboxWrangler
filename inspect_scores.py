import pickle
import pprint
from pathlib import Path
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Define the path to the sender scores file
DATA_DIR = Path("./email_data")
SCORES_FILE = DATA_DIR / "sender_scores.pkl"

def inspect_scores():
    """Loads and prints the contents of the sender_scores.pkl file."""
    if not SCORES_FILE.exists():
        logger.error(f"Sender scores file not found: {SCORES_FILE}")
        print(f"\nError: File not found at {SCORES_FILE}")
        return

    try:
        with open(SCORES_FILE, 'rb') as f:
            sender_scores = pickle.load(f)
        
        logger.info(f"Successfully loaded {len(sender_scores)} sender records from {SCORES_FILE}")
        
        print("\n--- Sender Scores Data ---")
        if sender_scores:
            pprint.pprint(sender_scores, indent=2)
        else:
            print("The sender scores file is empty.")
            
    except (EOFError, pickle.UnpicklingError) as p_err:
        logger.error(f"Error unpickling {SCORES_FILE}: {p_err}")
        print(f"\nError: Could not read the pickle file. It might be corrupted or empty.")
    except Exception as e:
        logger.error(f"An unexpected error occurred while reading {SCORES_FILE}: {e}", exc_info=True)
        print(f"\nAn unexpected error occurred: {e}")

if __name__ == "__main__":
    inspect_scores() 