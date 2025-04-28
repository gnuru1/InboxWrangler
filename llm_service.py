import os
import re
import json
import logging
import hashlib
import requests
from pathlib import Path

logger = logging.getLogger(__name__)

class LLMService:
    """
    Handles interactions with the LLM service for enhanced natural language understanding.
    This implementation supports both local LLM instances and API-based services.
    """
    def __init__(self, config=None):
        """
        Initialize LLM service with configuration

        Parameters:
        - config: Dictionary containing LLM configuration
          - api_type: 'local', 'openai', 'anthropic', etc.
          - api_key: API key for the service (if applicable)
          - api_endpoint: API endpoint URL
          - model: Model name to use
          - max_tokens: Maximum response tokens
        """
        self.config = config or {
            'api_type': 'local',
            'api_endpoint': 'http://localhost:8000/v1/chat/completions',
            'model': 'gpt-4o',
            'max_tokens': 1000,
            'temperature': 0.0
        }

        # Set up API key from config or environment
        self.api_key = self.config.get('api_key')
        if not self.api_key and 'api_type' in self.config and self.config['api_type'] != 'local':
            env_key = f"{self.config['api_type'].upper()}_API_KEY"
            self.api_key = os.environ.get(env_key)

        # Set up cache directory
        self.cache_dir = Path(self.config.get('cache_dir', './llm_cache'))
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        logger.info(f"LLM Service initialized with {self.config.get('api_type')} backend")

    def query(self, prompt, system_prompt=None, cache_key=None):
        """
        Send a query to the LLM service and return structured response

        Parameters:
        - prompt: User prompt text
        - system_prompt: Optional system prompt to guide LLM behavior
        - cache_key: Optional key for caching responses

        Returns:
        - Structured response from LLM (dict or text)
        """
        # Generate cache key if not provided
        if not cache_key and self.config.get('use_cache', True):
            cache_key = hashlib.md5(f"{system_prompt}|{prompt}".encode()).hexdigest()

        # Check cache first if cache_key provided
        if cache_key:
            cache_file = self.cache_dir / f"{cache_key}.json"
            if cache_file.exists():
                try:
                    with open(cache_file, 'r') as f:
                        return json.load(f)
                except Exception as e:
                    logger.warning(f"Could not load LLM cache file {cache_file}: {e}")
                    pass # Proceed to query API if cache load fails

        # Default system prompt for email analysis
        if system_prompt is None:
            system_prompt = """
            You are an email analysis assistant that extracts key information from emails.
            Provide outputs in structured JSON format based on the request.
            Focus on objective analysis of the content.
            """

        try:
            api_type = self.config.get('api_type', 'local')

            if api_type == 'local':
                response = self._query_local(prompt, system_prompt)
            elif api_type == 'openai':
                response = self._query_openai(prompt, system_prompt)
            elif api_type == 'anthropic':
                response = self._query_anthropic(prompt, system_prompt)
            else:
                logger.error(f"Unsupported LLM API type: {api_type}")
                return {"error": f"Unsupported LLM API type: {api_type}"}

            # Cache response if cache_key provided
            if cache_key:
                try:
                    with open(cache_file, 'w') as f:
                        json.dump(response, f)
                except Exception as e:
                    logger.debug(f"Error caching LLM response to {cache_file}: {e}")

            return response

        except requests.exceptions.RequestException as req_e:
            logger.error(f"Network error querying LLM ({api_type}) endpoint {self.config.get('api_endpoint', 'N/A')}: {req_e}")
            return {"error": f"Network error: {req_e}"}
        except Exception as e:
            logger.error(f"Error querying LLM ({api_type}): {e}", exc_info=True)
            return {"error": str(e)}

    def _query_local(self, prompt, system_prompt):
        """
        Query a local LLM API (e.g., LM Studio, Ollama, or similar local deployment)
        """
        endpoint = self.config.get('api_endpoint', 'http://localhost:8000/v1/chat/completions')
        payload = {
            "model": self.config.get('model', 'local-model'),
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "max_tokens": self.config.get('max_tokens', 1000),
            "temperature": self.config.get('temperature', 0.0)
        }
        logger.debug(f"Querying local LLM at {endpoint} with model {payload['model']}")
        response = requests.post(
            endpoint,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=self.config.get('timeout', 120) # Add timeout
        )

        if response.status_code == 200:
            result = response.json()
            content = result.get('choices', [{}])[0].get('message', {}).get('content', '')
            logger.debug(f"Local LLM response received. Length: {len(content)}")
            return self._parse_llm_content(content)
        else:
            logger.error(f"Local LLM API error ({endpoint}): {response.status_code}, {response.text}")
            return {"error": f"Local LLM API error: {response.status_code}"}

    def _query_openai(self, prompt, system_prompt):
        """
        Query OpenAI API
        """
        if not self.api_key:
            logger.error("OpenAI API key not found.")
            return {"error": "OpenAI API key not found"}

        endpoint = self.config.get('api_endpoint', "https://api.openai.com/v1/chat/completions")
        payload = {
            "model": self.config.get('model', 'gpt-4o'),
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "max_tokens": self.config.get('max_tokens', 1000),
            "temperature": self.config.get('temperature', 0.0)
        }
        logger.debug(f"Querying OpenAI API at {endpoint} with model {payload['model']}")
        response = requests.post(
            endpoint,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            },
            json=payload,
            timeout=self.config.get('timeout', 60)
        )

        if response.status_code == 200:
            result = response.json()
            content = result.get('choices', [{}])[0].get('message', {}).get('content', '')
            logger.debug(f"OpenAI response received. Length: {len(content)}")
            return self._parse_llm_content(content)
        else:
            logger.error(f"OpenAI API error: {response.status_code}, {response.text}")
            return {"error": f"OpenAI API error: {response.status_code}"}

    def _query_anthropic(self, prompt, system_prompt):
        """
        Query Anthropic API (Claude)
        """
        if not self.api_key:
            logger.error("Anthropic API key not found.")
            return {"error": "Anthropic API key not found"}

        endpoint = self.config.get('api_endpoint', "https://api.anthropic.com/v1/messages")
        payload = {
            "model": self.config.get('model', 'claude-3-opus-20240229'),
            "messages": [
                # Note: Anthropic API structure might differ slightly, adjust if needed
                # Typically system prompt is outside messages or handled differently
                # Assuming standard ChatML like structure for now
                 {"role": "user", "content": prompt}
            ],
            "system": system_prompt, # Pass system prompt separately
            "max_tokens": self.config.get('max_tokens', 1000),
            "temperature": self.config.get('temperature', 0.0)
        }
        logger.debug(f"Querying Anthropic API at {endpoint} with model {payload['model']}")
        response = requests.post(
            endpoint,
            headers={
                "Content-Type": "application/json",
                "x-api-key": self.api_key,
                "anthropic-version": self.config.get('anthropic_version', "2023-06-01")
            },
            json=payload,
            timeout=self.config.get('timeout', 60)
        )

        if response.status_code == 200:
            result = response.json()
            # Anthropic response structure might be different
            content = result.get('content', [{}])[0].get('text', '')
            logger.debug(f"Anthropic response received. Length: {len(content)}")
            return self._parse_llm_content(content)
        else:
            logger.error(f"Anthropic API error: {response.status_code}, {response.text}")
            return {"error": f"Anthropic API error: {response.status_code}"}

    def _parse_llm_content(self, content):
        """ Attempts to parse LLM response content as JSON, returns raw content on failure. """
        stripped_content = content.strip()
        # More robust JSON extraction: find first '{' and last '}'
        start_index = stripped_content.find('{')
        end_index = stripped_content.rfind('}')
        if start_index != -1 and end_index != -1 and end_index > start_index:
            json_str = stripped_content[start_index : end_index + 1]
            try:
                return json.loads(json_str)
            except json.JSONDecodeError as json_e:
                logger.warning(f"Failed to parse LLM response as JSON: {json_e}. Content: {content[:100]}...")
                # Fall through to return raw content if parsing fails
        return content # Return raw content if not JSON-like or parsing failed

    def analyze_email_content(self, subject, body, sender=None):
        """
        Analyze email content using LLM to extract key information

        Parameters:
        - subject: Email subject
        - body: Email body text
        - sender: Optional sender information

        Returns:
        - Dictionary containing extracted information or error dict
        """
        # Truncate body if too long
        max_body_length = self.config.get('llm_max_body_length', 1500) # Configurable limit
        truncated_body = body[:max_body_length]
        if len(body) > max_body_length:
            truncated_body += "... [truncated]"

        # Create hash-based cache key
        cache_key = hashlib.md5(f"{subject}|{truncated_body}|{sender}".encode()).hexdigest()

        system_prompt = """
        You are an AI assistant specializing in email analysis. Analyze the provided email and extract the following information:

        1. Main topics: List the 1-3 primary topics discussed as a JSON list of strings.
        2. Action items: Identify any tasks, requests, or required actions as a JSON list of strings.
        3. Urgency level: Classify as 'urgent', 'high', 'medium', or 'low' as a JSON string.
        4. Sentiment: Classify as 'positive', 'neutral', 'negative', or 'mixed' as a JSON string.
        5. Key entities: Identify important people, organizations, projects mentioned as a JSON list of strings.
        6. Category: Classify the email using one of the following categories: 'personal', 'professional', 'promotional', 'transactional', 'newsletter', 'spam', 'other' as a JSON string.

        Return ONLY a valid JSON object containing these keys: "topics", "action_items", "urgency", "sentiment", "entities", "category".
        Be concise and focus on objectively evident information from the email. If no action items are found, return an empty list.
        If unsure about a field, use a reasonable default (e.g., 'medium' urgency, 'neutral' sentiment, 'general' category, empty lists for topics/entities/action_items).
        Example JSON output:
        {"topics": ["Project Alpha Update", "Meeting Request"], "action_items": ["Schedule meeting for next week", "Send updated report by EOD"], "urgency": "high", "sentiment": "neutral", "entities": ["Project Alpha", "John Doe"], "category": "professional"}
        """

        prompt = f"Analyze the following email:\n\nFrom: {sender or 'Unknown'}\nSubject: {subject}\n\nBody:\n{truncated_body}"

        analysis = self.query(prompt, system_prompt, cache_key)

        # Validate and structure the response
        if isinstance(analysis, dict) and 'error' not in analysis:
            # Ensure all expected keys are present with correct types
            expected_keys_types = {
                "topics": list, "action_items": list, "urgency": str,
                "sentiment": str, "entities": list, "category": str
            }
            valid_analysis = {}
            all_keys_valid = True
            for key, expected_type in expected_keys_types.items():
                if key in analysis and isinstance(analysis[key], expected_type):
                    valid_analysis[key] = analysis[key]
                else:
                    logger.warning(f"LLM analysis missing or invalid type for key '{key}'. Using default.")
                    # Provide defaults based on type
                    valid_analysis[key] = [] if expected_type == list else "unknown"
                    if key == 'urgency': valid_analysis[key] = 'medium'
                    if key == 'sentiment': valid_analysis[key] = 'neutral'
                    if key == 'category': valid_analysis[key] = 'general'
                    all_keys_valid = False
            # Return the validated/defaulted dictionary
            return valid_analysis
        elif isinstance(analysis, dict) and 'error' in analysis:
            # LLM query itself failed
             logger.error(f"LLM query failed for email analysis: {analysis['error']}")
             return analysis # Propagate error dictionary
        else:
            # Fallback if LLM didn't return a dict or returned non-JSON string
            logger.warning(f"LLM analysis did not return a valid dictionary. Response: {str(analysis)[:100]}...")
            # Return a default structure indicating failure but matching expected format
            return {
                "topics": [], "action_items": [], "urgency": "medium",
                "sentiment": "neutral", "entities": [], "category": "general",
                "error": "LLM did not return valid structured data"
            }

    def generate_folder_name(self, email_cluster):
        """
        Generate a meaningful folder name for a cluster of emails

        Parameters:
        - email_cluster: List of email data dictionaries (should have 'subject', optionally 'topics', 'sender')

        Returns:
        - Folder name suggestion (string) or None if error
        """
        if not email_cluster or len(email_cluster) == 0:
            return "Miscellaneous" # Default for empty cluster

        # Prepare context for the LLM
        context = "Suggest a concise folder name (2-4 words) for emails with these subjects/topics:\n\n"
        for i, email_data in enumerate(email_cluster[:10]): # Limit context size
            subject = email_data.get('subject', 'No Subject')
            topics = email_data.get('topics')
            sender = email_data.get('sender')
            context += f"{i+1}. Subject: {subject}"
            if topics: context += f" (Topics: {', '.join(topics)})"
            if sender: context += f" (From: {sender})"
            context += "\n"

        system_prompt = """
        You are an AI assistant helping to organize emails. Based on the provided list of email subjects and topics,
        suggest a single, concise, and meaningful folder name (ideally 2-4 words) that accurately represents
        the common theme or purpose of this group of emails.
        Focus on the primary shared concept. Avoid generic names like "Updates" or "Misc".
        Return ONLY the suggested folder name as a plain string, without any extra text, quotes, or explanation.
        Example output: Project Alpha
        """

        # Create hash-based cache key based on the context string
        cache_key = hashlib.md5(context.encode()).hexdigest()

        folder_name_response = self.query(context, system_prompt, cache_key)

        # Handle response
        if isinstance(folder_name_response, dict) and 'error' in folder_name_response:
            logger.error(f"Error generating folder name: {folder_name_response['error']}")
            return "Miscellaneous" # Fallback on error
        elif isinstance(folder_name_response, str) and len(folder_name_response.strip()) > 0:
            # Clean up the suggested name
            folder_name = folder_name_response.strip('"\'\n ') # Corrected: Removed extra parenthesis inside strip argument
            # Basic sanitization (replace invalid chars for filenames)
            folder_name = re.sub(r'[<>:"/\\|?*]', '_', folder_name)
            # Limit length
            folder_name = folder_name[:50]
            # Avoid overly generic names sometimes returned
            if folder_name.lower() in ["miscellaneous", "general", "updates", "emails"]:
                 return "Miscellaneous"
            return folder_name
        else:
            logger.warning(f"LLM did not return a valid folder name string. Response: {folder_name_response}")
            return "Miscellaneous" # Fallback if response is unexpected

# Add the factory function here
def create_llm_service(config):
    """
    Factory function to create an LLMService instance.

    Args:
        config (dict): The configuration dictionary for the LLM service.

    Returns:
        LLMService: An instance of the LLMService.
    """
    return LLMService(config) 