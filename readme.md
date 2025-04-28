# Hybrid LLM-Enhanced Email Organizer

*Note - LLM integration has had zero testing.  All testing thus far using NLTK methods.

A sophisticated email management solution that leverages both traditional machine learning and optional LLM capabilities to intelligently organize Microsoft Outlook inboxes based on your unique behavior patterns.

## Overview

This tool analyzes your email interaction history (especially your sent mail and inbox reading patterns) to understand your personal priorities, then organizes your inbox based on these patterns rather than just addressing unread message volume. The hybrid approach combines the best of both worlds:

-   **Traditional ML/Rule-based Analysis** for efficiency, reliability, privacy, and strong baseline performance (sender importance, message state, keywords, temporal factors, recipient context).
-   **LLM Enhancement (Optional)** for deep semantic understanding, nuanced topic extraction, action item identification, and intelligent folder naming when enabled and available.

## Key Features

-   **Behavior-based prioritization**: Learns from your actual usage patterns (response times, read ratios, initiation frequency).
-   **Sent mail analysis**: Determines who you respond to quickly and thoroughly, and who you initiate contact with.
-   **Automated Contact Normalization**: Automatically identifies and maps different display names to a single canonical email address for consistent scoring.
-   **Content understanding**: Uses LLM (optional) or enhanced NLP rules to recognize action items, urgency signals, and categories.
-   **Dynamic folder creation**: Creates an intuitive folder structure based on priority, category, and content.
-   **Message State Awareness**: Considers unread status (with penalties for ignored emails), flags, due dates, and importance levels in prioritization.
-   **Recipient Analysis**: Factors in whether you are in To/CC, and the number of recipients.
-   **Privacy-focused**: Core analysis runs locally. Optional API-based LLM integration and local LLM server support (including Copilot Chat Bridge).
-   **Configurable Scoring**: Fine-tune scoring weights and thresholds via `config.json`.

## Why This Approach Works

Traditional email rules are limited by their rigid nature. Our hybrid system uncovers your *implicit* priorities by analyzing:

-   **Who you actually reply to** and how quickly (not just who you *say* is important).
-   **Who you initiate emails to** (indicating proactive importance).
-   **Which emails you ignore** over multiple checks (applying penalties).
-   **How you interact with different categories** of messages (newsletters vs. direct requests).
-   **What topics and content patterns** are meaningful to you (via NLP/LLM).
-   **Explicit signals** like flags, due dates, and high importance settings.
-   **Contextual clues** like recipient lists (direct message vs. mass email).

The result is a highly personalized organization system that matches your natural workflow.

## Installation

### Prerequisites

-   Windows 10 or later
-   Microsoft Outlook 2016 or later (must be installed, configured, and running during script execution)
-   Python 3.8+
-   Git (for cloning)

### Setup

1.  Clone the repository:
    ```bash
    git clone https://github.com/yourusername/hybrid-email-organizer.git
    cd hybrid-email-organizer
    ```

2.  Create and activate a virtual environment:
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    ```
    *(Note: Use `source venv/bin/activate` on Linux/macOS if applicable)*

3.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Download NLP Data (Required for Non-LLM mode / Fallback):**
    ```bash
    python -c "import nltk; nltk.download('punkt'); nltk.download('stopwords'); nltk.download('wordnet')"
    ```

5.  **Configure LLM Access (Optional but Recommended for full features):**
    LLM usage is controlled by settings in `config.json` (see Configuration section). If using an API service:
    *   **Option 1: API Service (OpenAI, Anthropic):**
        *   Set environment variables (recommended for security):
            ```bash
            set OPENAI_API_KEY=your_key_here
            # or
            set ANTHROPIC_API_KEY=your_key_here
            ```
        *   *Alternatively*, you can place the key directly in the `llm_config` section of `config.json`, but this is less secure.
    *   **Option 2: Local LLM Server:**
        *   Set up [LM Studio](https://lmstudio.ai/), [Ollama](https://ollama.ai/), or similar local LLM server.
        *   Configure the `api_type: "local"` and the correct `api_endpoint` in `config.json` under `llm_config`.
    *   **Option 3: Copilot Chat Bridge:**
        *   Ensure `copilot_chat_bridge.py` exists in the project directory.
        *   Set `use_copilot_proxy: true` in the `llm_config` section of `config.json`. Configure proxy settings if needed.

6.  **Customize Configuration (Optional):**
    *   Edit `config.json` to adjust weights, thresholds, and LLM settings (see Configuration section below). Create `config.json` if it doesn't exist (it will use defaults initially).

## Usage

**Important:** Ensure Microsoft Outlook is running before executing any commands.

All commands should be run from the `hybrid-email-organizer` directory where `main.py` is located.

### 1. Initial Analysis

Run the analysis to learn your email patterns, build sender scores, and initialize the contact normalization map. This reads your Sent Items and Inbox history. **Run this first.**

```bash
python main.py analyze --limit 5000
```

-   `analyze`: The command to perform the analysis.
-   `--limit <number>`: (Optional) Maximum number of emails to analyze *per folder* (Inbox, Sent Items). Default: `5000` (from `config.json`).
-   `--config <path>`: (Optional) Path to a custom configuration file (default: `./config.json`).
-   `--data-dir <path>`: (Optional) Directory to store analysis data (default: `./email_data`).
-   `--debug`: (Optional) Enable detailed debug logging to console and `outlook_organizer.log`.

This command creates/updates data files (like `sender_scores.pkl`, `email_tracking.pkl`, `contact_map.pkl`) in the specified `--data-dir`.

### 2. Generate Organization Report (Preview)

See what the organizer *would* do without making any changes. Useful for tuning configuration.

```bash
python main.py report --limit 100 --output ./reports
```

-   `report`: The command to generate recommendations.
-   `--limit <number>`: (Optional) Maximum number of emails to include in the report (default: 100).
-   `--output <path>`: (Optional) Directory to save the report files (CSV and HTML) (default: `./reports`).
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

This generates `.csv` and `.html` files in the output directory showing suggested actions for each email.

### 3. Organize Inbox (Dry Run)

Preview the organization changes *in the console log* without actually moving emails or creating folders/tasks. **This is the default behavior of the `organize` command.**

```bash
python main.py organize --limit 200
```

-   `organize`: The command to perform organization.
-   `--limit <number>`: (Optional) Process only the latest `<number>` emails in the Inbox root. If omitted, processes all items in the Inbox root.
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

Review the log output (e.g., `RECOMMENDATION: Move 'Subject...' to 'Folder...'`).

### 4. Apply Organization Changes

Organize your inbox by creating folders, moving emails, setting flags, and creating tasks based on the analysis and configuration. **This makes actual changes to your Outlook Inbox.**

```bash
python main.py organize --apply --limit 200
```

-   `organize`: The command to perform organization.
-   `--apply`: **Required** to actually make changes. Without this flag, it's a dry run (step 3).
-   `--limit <number>`: (Optional) Process only the latest `<number>` emails in the Inbox root.
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

The script will log the actions it takes (moving, flagging, etc.) and provide a summary at the end.

### 5. Recommend Contacts (Experimental)

Recommend contacts based on scoring thresholds (useful for identifying key contacts).

```bash
python main.py recommend --limit 10 --threshold 0.5
```

-   `recommend`: The command to list recommended contacts.
-   `--limit <number>`: (Optional) Maximum number of contacts to recommend (default: 10).
-   `--threshold <float>`: (Optional) Minimum score required for recommendation (default: 0.5).
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

### 6. Inspect Email Scores (Debugging/Tuning)

Analyze the detailed score calculation for the most recent emails in your Inbox to understand how components and weights contribute to the final score. Useful for tuning `config.json`.

```bash
python inspect_email_score.py --limit 5
```

-   `inspect_email_score.py`: The script to run.
-   `--limit <number>`: **Required.** Number of recent emails to inspect.
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

This script prints a detailed breakdown for each email, including:
    -   Subject, Sender, Received Time
    -   Final Calculated Score
    -   Table showing each scoring component (Sender, Topic, Temporal, Message State, Recipient), its raw score, the weight applied from `config.json`, and its contribution to the final score.

### 7. Analyze Inbox Sender Statistics (Cleanup Aid)

Get a quick overview of sender statistics in your current Inbox, focusing on volume, read/unread percentage, and subject line similarity (using fuzzy matching) to identify potential bulk cleanup targets.

```bash
python inbox_sender_stats.py --limit 1000 --top-senders 20 --fuzzy-threshold 85
```

-   `inbox_sender_stats.py`: The script to run.
-   `--limit <number>`: **Required.** Number of recent emails to analyze.
-   `--top-senders <number>`: (Optional) Number of top senders (by volume) to display (default: 20).
-   `--fuzzy-threshold <0-100>`: (Optional) Similarity threshold for grouping subjects (default: 85). Lower values group more aggressively.
-   `--config <path>`, `--data-dir <path>`, `--debug`: (Optional) Global options.

This script outputs a table showing the top senders, their total email count (within the limit), their unread percentage, and the percentage of their emails that fall into the largest "fuzzy subject" cluster.

### 8. Generate Standalone HTML Sender Report (Quick Overview)

Generate a quick HTML report of sender statistics directly from your Inbox *without* relying on prior analysis data or contact normalization. Useful for a fast overview based on raw sender information.

```bash
python html_sender_report.py --limit 1000 --output sender_report.html --top-senders 20 --fuzzy-threshold 85
```

-   `html_sender_report.py`: The standalone script to run.
-   `--limit <number>`: **Required.** Number of recent emails to analyze.
-   `--output <filename.html>`: **Required.** Path to save the generated HTML report.
-   `--top-senders <number>`: (Optional) Number of top senders (by volume) to display (default: 20).
-   `--fuzzy-threshold <0-100>`: (Optional) Similarity threshold for grouping subjects (default: 85).
-   `--debug`: (Optional) Enable detailed debug logging.

**Note:** This script uses the *raw* sender names/emails from Outlook and does *not* perform contact normalization like the main `analyze` command or the `inbox_sender_stats.py` script. The output is an HTML file.

## Configuration (`config.json`)

Customize the system's behavior by editing `config.json`. If the file doesn't exist, the defaults will be used. Here's an example reflecting many available options (based on the provided `config.json`):

```json
{
    "sender_weight": 0.35,
    "topic_weight": 0.25,
    "temporal_weight": 0.15,
    "message_state_weight": 0.15,
    "recipient_weight": 0.1,

    "high_priority_threshold": 0.8,
    "medium_priority_threshold": 0.5,

    "reply_pattern_score_factor": 1.0,
    "initiation_score_factor": 0.6,
    "read_kept_score_factor": 0.4,
    "min_emails_for_pattern": 1,

    "reply_time_weight": 0.4,
    "reply_rate_weight": 0.4,
    "reply_length_weight": 0.2,

    "max_analysis_emails": 5000,
    "days_for_temporal_analysis": 90,

    "use_llm_for_content": false,
    "use_llm_fallback": true,
    "llm_cache_dir": "./llm_cache",
    "llm_required": false,
    "enhanced_fallback": true,

    "nlp_extract_topics_count": 5,
    "nlp_detect_action_items": true,
    "nlp_use_keyword_boost": true,

    "unread_penalty": 0.15,         // Penalty applied if email is unread
    "ignore_penalty": 0.35,         // Additional penalty factor if seen unread multiple times
    "read_kept_bonus": 0.3,         // Bonus if read and kept in inbox
    "flagged_bonus": 0.15,
    "due_today_bonus": 0.25,        // For flagged items due today/overdue
    "due_soon_bonus": 0.15,         // For flagged items due in next 2 days
    "high_importance_bonus": 0.2,   // For emails marked High Importance
    "off_hours_bonus": 0.05,        // Slight boost for emails received outside 8am-6pm

    "to_me_bonus": 0.15,            // If current user is in 'To' field
    "direct_to_me_bonus": 0.1,      // Additional bonus if 'To' field has few recipients (<=3)
    "many_recipients_penalty": 0.1, // Penalty if many recipients (>10)
    "cc_me_penalty": 0.05,          // Penalty if current user is in 'CC' field

    "llm_config": {
        "api_type": "local",
        "api_endpoint": "http://localhost:1234/v1/chat/completions",
        "model": "local-model",
        "max_tokens": 1500,
        "temperature": 0.1,
        "use_cache": true,
        "timeout": 120,

        "use_copilot_proxy": false, // Enable to use Copilot Chat Bridge
        "copilot_proxy": {          // Settings for Copilot Chat Bridge
            "work_dir": "./copilot_work",
            "cache_dir": "./copilot_cache",
            "wait_time": 15,
            "use_cache": true
        }
    }
}
```

### Key Parameters:

| Parameter                   | Description                                                          | Default (from config.json) |
| :-------------------------- | :------------------------------------------------------------------- | :------------------------- |
| `sender_weight`             | Weight for sender score (behavior analysis)                          | 0.35                       |
| `topic_weight`              | Weight for content category/relevance                                | 0.25                       |
| `temporal_weight`           | Weight for recency/urgency signals                                   | 0.15                       |
| `message_state_weight`      | Weight for message state (unread, flag, due date, importance)        | 0.15                       |
| `recipient_weight`          | Weight for recipient info (To/CC, number of recipients)              | 0.1                        |
| `high_priority_threshold`   | Minimum score for "High Priority" folder/actions                     | 0.8                        |
| `medium_priority_threshold` | Minimum score for "Medium Priority" folder/actions                   | 0.5                        |
| `reply_pattern_score_factor`| Multiplier for reply pattern component of sender score               | 1.0                        |
| `initiation_score_factor`   | Multiplier for initiation component of sender score                  | 0.6                        |
| `read_kept_score_factor`    | Multiplier for read/kept component of sender score                   | 0.4                        |
| `min_emails_for_pattern`    | Min emails needed from/to sender for some pattern calculations       | 1                          |
| `reply_time_weight`         | Weight of response time within reply pattern score                   | 0.4                        |
| `reply_rate_weight`         | Weight of response frequency within reply pattern score              | 0.4                        |
| `reply_length_weight`       | Weight of response length within reply pattern score                 | 0.2                        |
| `max_analysis_emails`       | Max emails analyzed per folder during `analyze` command              | 5000                       |
| `days_for_temporal_analysis`| Lookback window for some temporal metrics                            | 90                         |
| `use_llm_for_content`       | Master switch to enable/disable LLM for content analysis             | false                      |
| `llm_required`              | If true, script fails if LLM is configured but unavailable           | false                      |
| `use_llm_fallback`          | If LLM enabled but fails, use traditional NLP?                       | true                       |
| `enhanced_fallback`         | Use enhanced NLP (rules, TF-IDF) when LLM off or failed              | true                       |
| `nlp_extract_topics_count`  | Number of topics via TF-IDF in fallback mode                         | 5                          |
| `nlp_detect_action_items`   | Use rule-based action item detection in fallback mode                | true                       |
| `nlp_use_keyword_boost`     | Boost scores based on keyword matches in fallback mode               | true                       |
| `unread_penalty`            | Penalty subtracted from state score if unread                        | 0.15                       |
| `ignore_penalty`            | Additional penalty factor if email seen unread multiple times        | 0.35                       |
| `read_kept_bonus`           | Bonus added to state score if read and kept in inbox                 | 0.3                        |
| `flagged_bonus`             | Bonus added to state score if flagged                                | 0.15                       |
| `due_today_bonus`           | Bonus for flagged items due today/overdue                            | 0.25                       |
| `due_soon_bonus`            | Bonus for flagged items due soon (1-2 days)                          | 0.15                       |
| `high_importance_bonus`     | Bonus for emails marked High Importance                              | 0.2                        |
| `off_hours_bonus`           | Bonus for emails received outside business hours (8am-6pm)           | 0.05                       |
| `to_me_bonus`               | Bonus if current user in 'To' field                                  | 0.15                       |
| `direct_to_me_bonus`        | Additional bonus if few recipients (<=3) in 'To' field               | 0.1                        |
| `many_recipients_penalty`   | Penalty if many recipients (>10)                                     | 0.1                        |
| `cc_me_penalty`             | Penalty if current user in 'CC' field                                | 0.05                       |
| `llm_config`                | Sub-dictionary for LLM provider details                              | (see below)                |

**Note:** The `llm_cache_dir` path is relative to the main project directory, not `data_dir`.

## LLM Configuration (`llm_config` section)

Configure your specific LLM provider within the `llm_config` section of `config.json`.

**Security Note:** Avoid committing API keys directly into `config.json`. Use environment variables where possible.

### Example: Local LLM Server (LM Studio, Ollama)

```json
"llm_config": {
  "api_type": "local",
  "api_endpoint": "http://localhost:1234/v1/chat/completions", // Adjust port if needed
  "model": "your-local-model-name", // Specify the model loaded in your server
  "max_tokens": 1500,
  "temperature": 0.1,
  "use_cache": true,
  "timeout": 120,
  "use_copilot_proxy": false // Ensure this is false if using direct local access
}
```

### Example: OpenAI

```json
"llm_config": {
  "api_type": "openai",
  // "api_key": "YOUR_OPENAI_KEY", // Less secure: Use environment variable OPENAI_API_KEY instead
  "model": "gpt-4o", // Or other desired model
  "max_tokens": 1500,
  "temperature": 0.0,
  "use_cache": true,
  "timeout": 60,
  "use_copilot_proxy": false // Ensure this is false
}
```

### Example: Anthropic Claude

```json
"llm_config": {
  "api_type": "anthropic",
  // "api_key": "YOUR_ANTHROPIC_KEY", // Less secure: Use environment variable ANTHROPIC_API_KEY instead
  "model": "claude-3-opus-20240229", // Or other desired model
  "max_tokens": 1500,
  "temperature": 0.1,
  "use_cache": true,
  "timeout": 60,
  "anthropic_version": "2023-06-01", // API version header
  "use_copilot_proxy": false // Ensure this is false
}
```

### Example: Copilot Chat Bridge

```json
"llm_config": {
  // Other settings like api_type, model etc. are ignored when use_copilot_proxy is true
  "use_copilot_proxy": true,
  "copilot_proxy": {
    "work_dir": "./copilot_work",    // Directory for temporary files
    "cache_dir": "./copilot_cache", // Directory for caching responses
    "wait_time": 15,                // Seconds to wait for Copilot response
    "use_cache": true
  },
  "use_cache": true // Main cache setting still relevant
}
```


## Running Without LLM

The system functions effectively using traditional NLP techniques even without an LLM.

1.  **Configure for Non-LLM:** Set `"use_llm_for_content": false` in your `config.json`. Ensure `"enhanced_fallback": true` is also set (which is the default). Alternatively, use the provided `non_llm_config.json`:
    ```bash
    python main.py analyze --config non_llm_config.json
    python main.py organize --apply --config non_llm_config.json
    ```
2.  **Automatic Fallback:** If `use_llm_for_content` is `true` but the LLM service fails (and `llm_required` is `false`, `use_llm_fallback` is `true`), the system will automatically revert to traditional NLP methods for content analysis for the affected emails.

## How It Works

### 1. Behavior Pattern Analysis (`analyzer.py`)

The system analyzes key data sources from your Outlook history:

-   **Sent Mail Patterns**: Response times, lengths, and frequencies to/from recipients. Calculates `sender_scores`.
-   **Inbox Reading Behaviors**: Tracks which emails you read/ignore per sender (`email_tracking.pkl`). Calculates read/kept stats (`inbox_behavior.pkl`).
-   **Contact Normalization**: Builds and uses a `contact_map.pkl` to map display names to email addresses, ensuring consistent scoring across identities.
-   **Folder Structure**: Analyzes your existing folder hierarchy (`folder_structure.pkl`).
-   Saves analysis results to the `data_dir`. Implements post-load validation to handle data format changes.

### 2. Hybrid Content Analysis (`content_processor.py`)

For each email, content is analyzed using:

-   **Traditional NLP** (Baseline / Fallback):
    -   Rule-based pattern matching for category (newsletter, professional, etc.), urgency, and action items.
    -   Keyword boosting.
-   **LLM Enhancement** (If `use_llm_for_content` is true and LLM available):
    -   Sends subject/body/metadata to LLM (`llm_service.py` or `copilot_chat_bridge.py`).
    -   Extracts topics, action items, sentiment, urgency, category, and entities based on LLM understanding.
    -   Can generate context-aware folder names.

### 3. Multi-Dimensional Scoring (`scorer.py`)

Each email receives a normalized score (0-1) based on a weighted combination of:

-   **Sender Score**: Importance based on historical interactions (replies, initiations, read/kept status). (Weight: `sender_weight`)
-   **Topic Score**: Relevance/importance based on content category and LLM/rule-based analysis. (Weight: `topic_weight`)
-   **Temporal Score**: Recency and urgency signals (days old, keywords, LLM urgency). (Weight: `temporal_weight`)
-   **Message State Score**: Considers unread status (incl. ignore penalty), flags, due dates, and importance markers. (Weight: `message_state_weight`)
-   **Recipient Score**: Factors in To/CC placement and number of recipients relative to the current user. (Weight: `recipient_weight`)

### 4. Organization System (`organizer.py`, `scorer.py`)

Based on the final score and content analysis, a recommended action is determined:

-   **Folder Assignment**:
    -   `Due Today`: For flagged items due today/overdue.
    -   `High Priority`: Score >= `high_priority_threshold`.
    -   `Medium Priority`: Score >= `medium_priority_threshold`.
    -   `Important`: If marked High Importance but below priority thresholds.
    -   `Action Required`: If specific action items detected and not already high/medium prio.
    -   `Category Folders`: (e.g., `Newsletter`, `Promotional`, `Personal`, `Professional`). LLM can suggest subfolders (e.g., `Professional/Project Alpha`).
    -   `Archive/Category`: For low-priority categories potentially suitable for archiving.
-   **Flagging**: Automatically flags Medium and High Priority emails, and those identified as 'Important'.
-   **Task Creation**: Creates Outlook tasks for High Priority emails and those with detected action items.
-   **Execution**: The `organize --apply` command creates necessary folders (via `outlook_utils.py`) and performs the move/flag/task actions.

## Privacy Considerations

-   **Local Operation**: The core analysis runs entirely on your local machine.
-   **Data Storage**: Analysis data (`sender_scores.pkl`, `contact_map.pkl`, etc.) is stored locally in the directory specified by `--data-dir` (default: `./email_data`).
-   **LLM Data**:
    -   If using an **API Service** (OpenAI, Anthropic), email content (subject, body, sender) is sent to the external service for analysis. Consult their privacy policies.
    -   If using a **Local LLM Server** or **Copilot Chat Bridge**, data remains on your local machine/network (subject to Microsoft's Copilot privacy policy if using the bridge).
    -   If **LLM is disabled** (`"use_llm_for_content": false`), no email content is sent externally by this tool.
-   **Caching**: LLM responses can be cached locally (`llm_cache_dir`, `copilot_cache`) to reduce API calls and costs, storing LLM-generated analysis tied to email content hashes.

Choose the configuration (local LLM, no LLM, API, Copilot) that best suits your privacy requirements.

## Advanced Usage & Troubleshooting

-   **Custom Configuration/Data Paths:** Use `--config` and `--data-dir` flags with commands.
-   **Debugging:** Use the `--debug` flag for verbose logging to console and `outlook_organizer.log`.
-   **Diagnostics Script:** Run `python diagnostics.py` for insights into loaded data, contact registry, and potential configuration issues. It saves reports to `diagnostics_output/`.
-   **Connection Issues:** Ensure Outlook is running, not modal, consider admin privileges, check Outlook security settings for programmatic access (use caution).
-   **Slow Performance:** Reduce `max_analysis_emails`, use `--limit` flags, run `analyze` less often, consider local LLM if network/API is slow.
-   **LLM Issues:** Verify API keys/endpoints/model names/Copilot setup, check connectivity, server logs, set `llm_required: false`, `use_llm_fallback: true`, or use `non_llm_config.json`.
-   **Scoring/Organization Issues:** Run `report`, check logs (`--debug`), adjust `config.json` weights/thresholds, ensure `analyze` was run recently, validate `config.json` syntax.

## License

[MIT License](LICENSE)

## Acknowledgments

This project utilizes several open-source libraries and frameworks, including:
-   pywin32: For interacting with the Outlook COM object model.
-   pandas, numpy: For data analysis.
-   NLTK: For natural language processing tasks in fallback mode.
-   requests: For communicating with LLM APIs.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an issue for discussion.