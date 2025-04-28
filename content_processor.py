import re
import logging
from collections import Counter

import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.stem import WordNetLemmatizer

from sklearn.feature_extraction.text import TfidfVectorizer

logger = logging.getLogger(__name__)

# Initialize NLP tools (consider making this configurable or lazy-loading)
try:
    nltk.data.find('tokenizers/punkt')
    nltk.data.find('corpora/stopwords')
    nltk.data.find('corpora/wordnet')
except LookupError:
    logger.info("Downloading NLTK data (punkt, stopwords, wordnet)...")
    nltk.download('punkt', quiet=True)
    nltk.download('stopwords', quiet=True)
    nltk.download('wordnet', quiet=True)
    logger.info("NLTK data download complete.")

STOP_WORDS = set(stopwords.words('english'))
LEMMATIZER = WordNetLemmatizer()

# Compiled regex patterns (moved from main class)
PATTERNS = {
    'url': re.compile(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'),
    'email': re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'),
    'phone': re.compile(r'\b(?:\+\d{1,2}\s?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b'),
    'special_chars': re.compile(r'[^\w\s]'), # Keep alphanumeric and whitespace
    'extra_whitespace': re.compile(r'\s+')
}

def preprocess_text(text):
    """
    Preprocess text for NLP analysis: lowercase, remove URLs/emails/phones,
    remove special chars, tokenize, remove stopwords, lemmatize.

    Args:
        text (str): The input text.

    Returns:
        str: The preprocessed text.
    """
    if not isinstance(text, str):
        logger.warning(f"preprocess_text received non-string input: {type(text)}. Returning empty string.")
        return ""

    # Convert to lowercase
    text = text.lower()

    # Remove URLs, email addresses, and phone numbers
    text = PATTERNS['url'].sub(' ', text)
    text = PATTERNS['email'].sub(' ', text)
    text = PATTERNS['phone'].sub(' ', text)

    # Remove special characters (leaving alphanumeric and whitespace)
    text = PATTERNS['special_chars'].sub(' ', text)

    # Remove extra whitespace
    text = PATTERNS['extra_whitespace'].sub(' ', text).strip()

    # Tokenize
    try:
        tokens = word_tokenize(text)
    except Exception as e:
        logger.debug(f"NLTK word_tokenize failed: {e}")
        tokens = text.split() # Fallback to simple split

    # Remove stopwords and lemmatize
    processed_tokens = []
    for token in tokens:
        if token not in STOP_WORDS and len(token) > 1 and token.isalpha(): # Ensure token is alphabetic
            try:
                lemma = LEMMATIZER.lemmatize(token)
                processed_tokens.append(lemma)
            except Exception as e:
                logger.debug(f"NLTK lemmatize failed for token '{token}': {e}")
                processed_tokens.append(token) # Keep original token on error

    return ' '.join(processed_tokens)

def extract_topics_tfidf(text, num_topics=3):
    """
    Extract top topics using TF-IDF (traditional NLP approach as fallback).

    Args:
        text (str): The preprocessed text.
        num_topics (int): The maximum number of topics to return.

    Returns:
        list: A list of top topic strings.
    """
    if not text:
        return []

    try:
        vectorizer = TfidfVectorizer(max_features=1000, stop_words='english')
    except Exception as e:
        logger.error(f"Failed to initialize TfidfVectorizer: {e}")
        words = text.split()
        word_counts = Counter(w for w in words if len(w) > 1)
        return [w for w, _ in word_counts.most_common(num_topics)]

    try:
        sentences = sent_tokenize(text)
    except Exception as e:
        logger.debug(f"NLTK sent_tokenize failed: {e}")
        sentences = text.split('.')

    if len(sentences) < 2:
        words = text.split()
        word_counts = Counter(w for w in words if len(w) > 1)
        return [w for w, _ in word_counts.most_common(num_topics)]

    try:
        tfidf_matrix = vectorizer.fit_transform(sentences)
        feature_names = vectorizer.get_feature_names_out()

        if tfidf_matrix.shape[1] == 0 or len(feature_names) == 0:
            logger.debug("TF-IDF resulted in zero features.")
            words = text.split()
            word_counts = Counter(w for w in words if len(w) > 1)
            return [w for w, _ in word_counts.most_common(num_topics)]

        tfidf_sums = tfidf_matrix.sum(axis=0).A1
        actual_num_topics = min(num_topics, len(feature_names))
        top_indices = tfidf_sums.argsort()[-actual_num_topics:][::-1]
        top_terms = [feature_names[i] for i in top_indices]
        return top_terms

    except ValueError as ve:
        logger.debug(f"TF-IDF ValueError: {ve}. Falling back to word counts.")
        words = text.split()
        word_counts = Counter(w for w in words if len(w) > 1)
        return [w for w, _ in word_counts.most_common(num_topics)]
    except Exception as e:
        logger.error(f"Error during TF-IDF topic extraction: {e}")
        words = text.split()
        word_counts = Counter(w for w in words if len(w) > 1)
        return [w for w, _ in word_counts.most_common(num_topics)]

def analyze_urgency_rules(subject, body):
    """
    Analyze urgency level using rule-based approach.

    Args:
        subject (str): Email subject.
        body (str): Email body.

    Returns:
        str: Urgency level ('high', 'medium', or 'low').
    """
    subject_str = str(subject) if subject else ""
    body_str = str(body) if body else ""
    text = (subject_str + " " + body_str).lower()

    high_urgency_patterns = [
        r'\b(urgent|asap|immediately|emergency|now)\b',
        r'\b(due today|due tomorrow)\b',
        r'\b(critical|crucial|vital)\b'
    ]
    medium_urgency_patterns = [
        r'\b(important|priority|attention)\b',
        r'\b(please respond|please reply|needs response)\b',
        r'\b(deadline|due this week|due soon)\b'
    ]

    try:
        high_count = sum(1 for pattern in high_urgency_patterns if re.search(pattern, text))
        medium_count = sum(1 for pattern in medium_urgency_patterns if re.search(pattern, text))
    except Exception as e:
        logger.error(f"Regex error during urgency analysis: {e}")
        return 'medium'

    if high_count >= 2 or (high_count >= 1 and medium_count >= 1):
        return 'high'
    elif medium_count >= 1 or high_count >= 1:
        return 'medium'
    else:
        return 'low'

def categorize_email_rules(subject, body):
    """
    Categorize email using rule-based approach.

    Args:
        subject (str): Email subject.
        body (str): Email body.

    Returns:
        str: Category name (e.g., 'newsletter', 'professional', 'general').
    """
    subject_str = str(subject) if subject else ""
    body_str = str(body) if body else ""
    text = (subject_str + " " + body_str).lower()

    categories = {
        'newsletter': [r'\b(newsletter|update|digest)\b', r'unsubscribe', r'\b(weekly|monthly|quarterly)\s+update'],
        'promotional': [r'\b(offer|discount|sale|promo|marketing)\b', r'unsubscribe', r'\b(limited time|exclusive)\b'],
        'personal': [r'\b(hey|hi|hello|greetings)\b.*', r'\b(how are you|hope you|thinking of you)\b', r'family|friend|personal'], # Simplified personal match
        'professional': [r'\b(meeting|discussion|project|report|business|client|colleague)\b', r'\b(regards|sincerely|best|team)\b'],
        'transactional': [r'\b(order|invoice|receipt|payment|transaction|shipping|booking|confirmation)\b', r'\b(confirm|confirmed)\b'],
        'spam': [r'\b(viagra|cialis|pharmacy|loan|mortgage|refinance|degree|online degree)\b', r'click here', r'unsubscribe at'] # Basic spam keywords
    }

    category_scores = {}
    try:
        for category, patterns in categories.items():
            score = sum(3 if re.search(pattern, subject_str.lower()) else
                        1 if re.search(pattern, text) else
                        0
                        for pattern in patterns)
            # Boost score for unsubscribe links in promo/newsletter
            if category in ['promotional', 'newsletter'] and 'unsubscribe' in text:
                score += 2
            category_scores[category] = score
    except Exception as e:
        logger.error(f"Regex error during categorization: {e}")
        return 'general'

    # Select category with highest score > 0
    max_score = max(category_scores.values()) if category_scores else 0
    if max_score > 0:
        # Prioritize spam if its score is high enough
        if category_scores.get('spam', 0) >= 3:
            return 'spam'
        # Get all categories with the max score
        top_categories = [cat for cat, score in category_scores.items() if score == max_score]
        # Simple priority: transactional > professional > personal > others
        priority = ['transactional', 'professional', 'personal', 'newsletter', 'promotional']
        for p_cat in priority:
            if p_cat in top_categories:
                return p_cat
        return top_categories[0] # Return first match if no priority hit
    else:
        return 'general'

def extract_action_items_rules(text):
    """
    Extract potential action items using rule-based approach.
    
    Args:
        text (str): Email text (subject + body)
        
    Returns:
        list: A list of potential action items found
    """
    if not text:
        return []
        
    action_items = []
    
    # Common action item patterns
    action_patterns = [
        r'(?:please|kindly|can you|could you)[^.!?]*\?',  # Please/can you do X?
        r'(?:please|kindly|can you|could you)[^.!?]*(?:\.|$)',  # Please do X.
        r'(?:need to|needs to|must|should)[^.!?]*(?:\.|$)',  # Need to do X.
        r'(?:don\'t forget to|remember to)[^.!?]*(?:\.|$)',  # Remember to do X.
        r'(?:action (?:needed|required|item))[^.!?]*(?:\.|$)',  # Action needed: X.
        r'deadline[^.!?]*(?:\.|$)',  # Deadline for X.
        r'by (?:monday|tuesday|wednesday|thursday|friday|saturday|sunday|tomorrow|next week|today|eod|eow)',  # By Monday...
        r'due (?:date|by)[^.!?]*(?:\.|$)'  # Due by...
    ]
    
    try:
        # Split into sentences for better context
        sentences = sent_tokenize(text)
    except Exception as e:
        logger.debug(f"NLTK sent_tokenize failed: {e}")
        sentences = text.split('.')
    
    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue
            
        for pattern in action_patterns:
            try:
                matches = re.findall(pattern, sentence, re.IGNORECASE)
                for match in matches:
                    if len(match) > 10:  # Avoid tiny fragments
                        # Clean up the action item
                        action = match.strip()
                        # Capitalize first letter
                        if action and len(action) > 0:
                            action = action[0].upper() + action[1:]
                        action_items.append(action)
            except Exception as e:
                logger.debug(f"Regex error in action item extraction: {e}")
                
    # Deduplicate and limit
    seen = set()
    unique_actions = []
    for action in action_items:
        # Create a simplified comparison key
        key = re.sub(r'\s+', ' ', action.lower()).strip()
        if key not in seen and len(key) > 0:
            seen.add(key)
            unique_actions.append(action)
    
    return unique_actions[:5]  # Limit to top 5 most likely action items

def process_email_content(email_item, llm_service, config):
    """
    Extract and analyze email content using LLM if enabled, otherwise fallback.

    Args:
        email_item: The Outlook mail item.
        llm_service (LLMService): The initialized LLM service instance.
        config (dict): The application configuration.

    Returns:
        dict: A dictionary containing the content analysis results.
    """
    default_analysis = {
        'processed_text': "", 'word_count': 0, 'topics': [], 'action_items': [],
        'urgency': 'medium', 'sentiment': 'neutral', 'entities': [], 'category': 'general'
    }
    try:
        # Get basic metadata safely using a utility if available, or getattr
        subject = getattr(email_item, 'Subject', "")
        body = getattr(email_item, 'Body', "")
        sender = "Unknown Sender"
        try:
            sender = getattr(email_item, 'SenderEmailAddress', getattr(email_item, 'SenderName', "Unknown Sender"))
        except pywintypes.com_error:
            sender = getattr(email_item, 'SenderName', "Unknown Sender")

        # Preprocess text
        processed_text = preprocess_text(subject + " " + body)
        content_analysis = default_analysis.copy()
        content_analysis.update({
            'processed_text': processed_text,
            'word_count': len(body.split()),
            'subject_length': len(subject),
            'body_length': len(body)
        })

        # Determine if we should attempt LLM analysis
        use_llm = config.get('use_llm_for_content', True)
        llm_required = config.get('llm_required', False)
        enhanced_fallback = config.get('enhanced_fallback', True)
        nlp_topic_count = config.get('nlp_extract_topics_count', 3)
        detect_action_items = config.get('nlp_detect_action_items', True)
        
        # Try LLM if enabled and available
        llm_analysis = None
        llm_error = None
        if use_llm and llm_service:
            try:
                llm_analysis = llm_service.analyze_email_content(subject, body, sender)
            except Exception as e:
                llm_error = str(e)
                logger.error(f"Error during LLM content analysis call: {e}")
                llm_analysis = {"error": llm_error}  # Treat error like a failed LLM response
        
        # Handle LLM results or switch to fallback
        if llm_analysis and isinstance(llm_analysis, dict) and 'error' not in llm_analysis:
            logger.debug(f"Using LLM analysis results for subject: {subject[:30]}...")
            # Overwrite defaults with valid LLM results
            content_analysis.update(llm_analysis)
        elif not use_llm or not llm_service or (use_llm and config.get('use_llm_fallback', True)):
            # Use fallback in these cases:
            # 1. LLM not enabled (not use_llm)
            # 2. LLM service not available (not llm_service)
            # 3. LLM enabled but fallback also enabled (use_llm && use_llm_fallback)
            
            # Log appropriate message based on reason for fallback
            if not use_llm:
                logger.debug("LLM disabled by configuration, using rule-based analysis.")
            elif not llm_service:
                logger.info("LLM service unavailable, using rule-based analysis.")
            else:
                logger.warning(f"LLM analysis failed or invalid, using rule-based fallback. Error: {llm_error or 'Unknown error'}")
                
            # Build fallback content analysis
            fallback_analysis = {
                'topics': extract_topics_tfidf(processed_text, num_topics=nlp_topic_count),
                'urgency': analyze_urgency_rules(subject, body),
                'category': categorize_email_rules(subject, body),
                'sentiment': 'neutral',  # Default sentiment
                'entities': []  # Default empty entities list
            }
            
            # Add action items if enabled
            if detect_action_items:
                fallback_analysis['action_items'] = extract_action_items_rules(subject + " " + body)
            else:
                fallback_analysis['action_items'] = []
                
            # Apply the fallback analysis
            content_analysis.update(fallback_analysis)
        elif llm_required:
            # LLM is required but failed and fallback is disabled
            logger.error(f"LLM required by configuration but failed: {llm_error or 'Unknown error'}")
            content_analysis.update({
                "error": f"LLM analysis required but failed: {llm_error or 'Unknown error'}"
            })
        
        # Apply any keyword-based boosting if enabled
        if config.get('nlp_use_keyword_boost', True) and not use_llm:
            # Example of simple keyword boosting logic for urgency
            if re.search(r'\b(urgent|asap|immediately|emergency)\b', (subject + " " + body).lower()):
                if content_analysis['urgency'] == 'medium':
                    content_analysis['urgency'] = 'high'
        
        return content_analysis

    except Exception as e:
        logger.error(f"Unexpected error in process_email_content: {e}", exc_info=True)
        return {**default_analysis, "error": f"Failed to process content: {str(e)}"} 