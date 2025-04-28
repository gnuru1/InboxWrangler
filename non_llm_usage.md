# Using Hybrid Outlook Organizer Without LLM Access

The Hybrid Outlook Organizer is designed to work effectively even without LLM (Large Language Model) access. This document explains how to configure and use the system in traditional NLP mode.

## Configuration for Non-LLM Operation

To optimize the system for operation without LLM, you can create or modify your `config.json` file with these settings:

```json
{
  "use_llm_for_content": false,
  "enhanced_fallback": true,
  "nlp_extract_topics_count": 5,
  "nlp_detect_action_items": true,
  "nlp_use_keyword_boost": true
}
```

### Key Configuration Options

| Option | Description | Recommended Value |
|--------|-------------|-------------------|
| `use_llm_for_content` | Whether to attempt using LLM for content analysis | `false` |
| `enhanced_fallback` | Use the enhanced traditional NLP methods | `true` |
| `nlp_extract_topics_count` | Number of topics to extract from each email | `3-5` |
| `nlp_detect_action_items` | Use rule-based detection for action items | `true` |
| `nlp_use_keyword_boost` | Apply keyword-based boosting to scores | `true` |

## How It Works Without LLM

When operating without LLM capabilities, the system uses several traditional natural language processing techniques:

1. **TF-IDF Analysis** for topic extraction
2. **Rule-based Pattern Matching** for:
   - Action item detection
   - Email categorization
   - Urgency detection
3. **Metadata Analysis** using:
   - Sender importance (based on your response patterns)
   - Temporal factors (recency, time of day)
   - Message state (flagged, unread, etc.)
   - Recipient information (To/CC fields)

## Running the Application Without LLM

You can run the application with the `--config` parameter pointing to your LLM-disabled configuration:

```bash
python main.py analyze --config non_llm_config.json
python main.py organize --apply --config non_llm_config.json
```

If you've already set up your configuration file, the system will automatically switch to fallback mode when LLM services are unavailable.

## Performance Expectations

Without LLM, you can expect:

- **Content Understanding**: Good categorization of emails (professional, personal, newsletters, etc.)
- **Sender Importance**: Fully functional (based on your response patterns)
- **Temporal Analysis**: Fully functional (recency, urgency based on keywords)
- **Message State Analysis**: Fully functional (flags, due dates, etc.)
- **Recipient Analysis**: Fully functional (To/CC fields analysis)

The main difference is in the richness of content understanding, particularly:
- Less nuanced action item extraction
- Less sophisticated topic discovery
- More reliance on keyword patterns rather than semantic understanding

## Tips for Optimal Non-LLM Performance

1. **Run Analysis First**: Always run `analyze` before `organize` to build up sender scores
2. **Review Recommendations**: Run with `organize` (no `--apply`) first to review recommendations
3. **Consider Folder Structure**: Create a logical folder structure for categorization
4. **Adjust Weights**: You may want to increase `sender_weight` and `temporal_weight` slightly when operating without LLM

## Future LLM Integration

When you're ready to integrate LLM capabilities in the future:

1. Set up your LLM service (local or API-based)
2. Update your configuration with LLM settings
3. Set `use_llm_for_content` to `true`

The system will seamlessly integrate LLM capabilities while preserving all your learned patterns and user data.

## Troubleshooting

If you encounter issues while running in non-LLM mode:

1. Check the log file (`outlook_organizer.log`) for detailed error messages
2. Ensure NLTK data is downloaded (run `python -c "import nltk; nltk.download('punkt'); nltk.download('stopwords'); nltk.download('wordnet')"`)
3. If scoring seems off, try adjusting weights in your configuration 