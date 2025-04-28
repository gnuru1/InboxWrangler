import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = {
    'sender_weight': 0.4,
    'topic_weight': 0.25,
    'temporal_weight': 0.15,
    'message_state_weight': 0.1,  # New weight for message state factors
    'recipient_weight': 0.1,      # New weight for recipient information
    'high_priority_threshold': 0.8,
    'medium_priority_threshold': 0.5,
    'response_time_weight': 0.4,
    'response_rate_weight': 0.4,
    'response_length_weight': 0.2,
    'max_analysis_emails': 5000,  # Max emails to analyze per folder
    'min_emails_for_pattern': 5,  # Min emails needed to establish a pattern
    'days_for_temporal_analysis': 90,  # Analyze last 90 days for temporal patterns
    
    # LLM Configuration
    'use_llm_for_content': True,  # Whether to use LLM for content analysis
    'use_llm_fallback': True,     # Use traditional ML as fallback if LLM fails
    'llm_cache_dir': './llm_cache',  # Default relative cache dir (often overridden)
    'llm_required': False,        # If True, fails when LLM unavailable; if False, uses fallbacks
    'enhanced_fallback': True,    # Use enhanced traditional NLP when LLM unavailable
    
    # Traditional NLP settings (used in fallback mode)
    'nlp_extract_topics_count': 5,       # Number of topics to extract in fallback mode
    'nlp_detect_action_items': True,     # Enable rule-based action item detection 
    'nlp_use_keyword_boost': True,       # Boost scores based on keyword matching
    
    # Message state factors
    'unread_bonus': 0.1,            # Bonus for unread emails
    'flagged_bonus': 0.15,          # Bonus for flagged emails
    'due_today_bonus': 0.25,        # Bonus for emails due today
    'due_soon_bonus': 0.15,         # Bonus for emails due soon (next 2 days)
    'high_importance_bonus': 0.2,   # Bonus for high importance emails
    'off_hours_bonus': 0.05,        # Bonus for emails received outside business hours
    
    # Recipient information factors
    'to_me_bonus': 0.15,            # Bonus for emails sent directly to me
    'direct_to_me_bonus': 0.1,      # Additional bonus for emails with few recipients
    'many_recipients_penalty': 0.1, # Penalty for mass emails
    'cc_me_penalty': 0.05,          # Penalty for emails where I'm in CC
    
    # LLM Service specific defaults (can be overridden by main config's llm_config section)
    'llm_config': {
        'api_type': 'local',
        'api_endpoint': 'http://localhost:1234/v1/chat/completions', # LM Studio default
        'model': 'local-model', # Placeholder - Set to your actual model name
        'max_tokens': 1500,
        'temperature': 0.1,
        'use_cache': True,
        'timeout': 120, # Default timeout for LLM queries
        
        # Copilot Chat Bridge configuration
        'use_copilot_proxy': False,  # Set to True to use Copilot Chat as LLM
        'copilot_proxy': {
            'work_dir': './copilot_work',     # Directory for working files
            'cache_dir': './copilot_cache',   # Directory for caching results
            'wait_time': 15,                  # Time to wait for Copilot to respond (seconds)
            'use_cache': True                 # Whether to cache results
        }
    },
}

def load_config(config_path='./config.json'):
    """
    Load configuration from a JSON file, merging with defaults.

    Args:
        config_path (str or Path): Path to the configuration file.

    Returns:
        dict: The loaded and merged configuration.
    """
    config_file = Path(config_path)
    config = DEFAULT_CONFIG.copy() # Start with defaults

    if config_file.exists():
        try:
            with open(config_file, 'r') as f:
                user_config = json.load(f)

            # Deep merge user config onto defaults (simple merge for top level, llm_config)
            config.update(user_config)
            if 'llm_config' in user_config:
                # Ensure llm_config exists in defaults before updating
                if 'llm_config' in config:
                     config['llm_config'].update(user_config['llm_config'])
                else:
                     config['llm_config'] = user_config['llm_config'] # If default didn't have it

            logger.info(f"Loaded configuration from {config_file}")

        except json.JSONDecodeError as e:
            logger.error(f"Error decoding JSON from config file '{config_file}': {e}. Using default config.")
        except Exception as e:
            logger.error(f"Error loading config file '{config_file}': {e}. Using default config.")
    else:
        logger.info(f"Config file '{config_file}' not found. Using default configuration.")

    # Ensure llm_cache_dir exists (might be set in user config or default)
    # This might be better handled during Organizer initialization based on data_dir
    # try:
    #     llm_cache_path = Path(config.get('llm_cache_dir', DEFAULT_CONFIG['llm_cache_dir']))
    #     llm_cache_path.mkdir(parents=True, exist_ok=True)
    # except Exception as e:
    #     logger.warning(f"Could not create llm_cache_dir '{config.get('llm_cache_dir')}': {e}")

    return config

# Example usage (optional, for testing)
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    # Create a dummy config file for testing
    dummy_config = {
        'sender_weight': 0.6,
        'llm_config': {
            'model': 'override-model',
            'temperature': 0.5
        }
    }
    with open('./dummy_config.json', 'w') as f:
        json.dump(dummy_config, f, indent=4)

    print("--- Loading Default Config ---")
    default_cfg = load_config('./non_existent_config.json')
    print(json.dumps(default_cfg, indent=2))

    print("\n--- Loading Dummy Config --- ")
    loaded_cfg = load_config('./dummy_config.json')
    print(json.dumps(loaded_cfg, indent=2))

    # Clean up dummy file
    Path('./dummy_config.json').unlink() 