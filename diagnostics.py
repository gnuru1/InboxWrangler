#!/usr/bin/env python
"""
Email Analysis Diagnostics Utility

This script provides diagnostic information about the email analysis system,
including configuration, data files, and contact insights.
"""

import argparse
import json
import pickle
import logging
import glob
import csv
from datetime import datetime
from pathlib import Path
import collections
import re
import os
import sys
import operator
from typing import Dict, List, Any, Optional, Tuple, Set, Union

# Check for required packages
try:
    import pandas as pd
    import numpy as np
    from tabulate import tabulate
except ImportError:
    print("Error: Required packages not found. Please install them using:")
    print("pip install pandas numpy tabulate")
    sys.exit(1)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger('diagnostics')

class DiagnosticsError(Exception):
    """Custom exception for diagnostics errors"""
    pass

def load_config(config_path: str) -> Dict:
    """Load configuration file"""
    try:
        if not os.path.exists(config_path):
            logger.warning(f"Config file not found at {config_path}, using default config")
            return {}
        
        with open(config_path, 'r') as f:
            config = json.load(f)
            logger.info(f"Loaded configuration from {config_path}")
            return config
    except Exception as e:
        logger.error(f"Error loading config: {str(e)}")
        return {}

def load_pickle_file(file_path: str) -> Any:
    """Load data from a pickle file"""
    try:
        if not os.path.exists(file_path):
            logger.warning(f"Pickle file not found at {file_path}")
            return None
        
        with open(file_path, 'rb') as f:
            data = pickle.load(f)
            logger.info(f"Loaded data from {file_path}")
            return data
    except Exception as e:
        logger.error(f"Error loading pickle file {file_path}: {str(e)}")
        return None

def find_latest_report(reports_dir: str) -> str:
    """Find the most recent recommendation report"""
    try:
        reports = glob.glob(os.path.join(reports_dir, "recommendations_*.csv"))
        if not reports:
            logger.warning(f"No recommendation reports found in {reports_dir}")
            return None
        
        latest_report = max(reports, key=os.path.getctime)
        logger.info(f"Latest report found: {latest_report}")
        return latest_report
    except Exception as e:
        logger.error(f"Error finding latest report: {str(e)}")
        return None

def load_csv_to_dataframe(file_path: str) -> pd.DataFrame:
    """Load CSV file into a pandas DataFrame"""
    try:
        if not os.path.exists(file_path):
            logger.warning(f"CSV file not found at {file_path}")
            return pd.DataFrame()
        
        df = pd.read_csv(file_path)
        logger.info(f"Loaded CSV from {file_path} with {len(df)} rows")
        return df
    except Exception as e:
        logger.error(f"Error loading CSV file {file_path}: {str(e)}")
        return pd.DataFrame()

def convert_datetime_for_json(obj):
    """Convert datetime objects to ISO format strings for JSON serialization"""
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable")

def clean_data_for_json(data):
    """Recursively clean data structures for JSON serialization"""
    if isinstance(data, dict):
        return {k: clean_data_for_json(v) for k, v in data.items() if k != '_loaded_data'}
    elif isinstance(data, list):
        return [clean_data_for_json(item) for item in data]
    elif isinstance(data, (np.int64, np.int32, np.float64, np.float32)):
        return float(data) if np.issubdtype(type(data), np.floating) else int(data)
    elif isinstance(data, (datetime, pd.Timestamp)):
        return data.isoformat()
    elif isinstance(data, (np.ndarray, pd.Series)):
        return clean_data_for_json(data.tolist())
    elif pd.isna(data):
        return None
    return data

def aggregate_email_tracking(email_tracking_data: Dict) -> Dict:
    """Aggregate email tracking data by sender"""
    if not email_tracking_data:
        return {}
    
    sender_counts = collections.defaultdict(int)
    sender_timestamps = collections.defaultdict(list)
    
    for email_id, data in email_tracking_data.items():
        sender = data.get('sender', '').lower()
        if sender:
            sender_counts[sender] += 1
            if 'received_time' in data and data['received_time']:
                sender_timestamps[sender].append(data['received_time'])
    
    # Calculate first and last seen dates
    sender_stats = {}
    for sender, count in sender_counts.items():
        timestamps = sorted(sender_timestamps[sender]) if sender in sender_timestamps else []
        sender_stats[sender] = {
            'email_count': count,
            'first_seen': min(timestamps) if timestamps else None,
            'last_seen': max(timestamps) if timestamps else None,
            'days_active': (max(timestamps) - min(timestamps)).days if len(timestamps) > 1 else 0
        }
    
    return sender_stats

def build_contacts_registry(
    sender_scores: Dict,
    email_tracking: Dict,
    recommendations_df: pd.DataFrame
) -> Dict:
    """
    Build a comprehensive contacts registry by merging data from multiple sources
    """
    contacts = {}
    
    # Add contacts from sender scores
    if sender_scores:
        for email, score_data in sender_scores.items():
            email = email.lower()
            if email not in contacts:
                contacts[email] = {
                    'email': email,
                    'sender_score': score_data.get('score', 0),
                    'score_components': score_data.get('components', {}),
                    'last_scored': score_data.get('last_updated')
                }
    
    # Build a map of display names to email addresses
    contact_map = {}
    if email_tracking:
        # First try to extract direct correspondences from the tracking data
        for entry_id, data in email_tracking.items():
            sender_name = data.get('sender_name', '')
            sender_email = data.get('sender_email', '')
            
            if sender_name and sender_email and '@' in sender_email:
                contact_map[sender_name.lower()] = sender_email.lower()
    
    # Merge tracking data with normalization
    tracking_agg = aggregate_email_tracking(email_tracking)
    for email, tracking_data in tracking_agg.items():
        email = email.lower()
        
        # Normalize email: if it's a display name, try to find its email
        normalized_email = email
        if '@' not in email and email in contact_map:
            normalized_email = contact_map[email]
            logger.debug(f"Normalized contact: '{email}' -> '{normalized_email}'")
        
        if normalized_email in contacts:
            contacts[normalized_email].update(tracking_data)
        else:
            contacts[normalized_email] = {
                'email': normalized_email,
                'sender_score': 0,
                **tracking_data
            }
    
    # Merge recommendation data
    if not recommendations_df.empty and 'Email' in recommendations_df.columns:
        for _, row in recommendations_df.iterrows():
            email = row['Email'].lower() if isinstance(row.get('Email'), str) else None
            
            # Normalize email: if it's a display name, try to find its email
            normalized_email = email
            if email and '@' not in email and email in contact_map:
                normalized_email = contact_map[email]
                logger.debug(f"Normalized contact in recommendations: '{email}' -> '{normalized_email}'")
            
            if normalized_email and normalized_email in contacts:
                contacts[normalized_email]['recommended'] = True
                contacts[normalized_email]['recommendation_score'] = row.get('Score', 0)
            elif normalized_email:
                contacts[normalized_email] = {
                    'email': normalized_email,
                    'recommended': True,
                    'recommendation_score': row.get('Score', 0),
                    'sender_score': 0
                }
    
    # Format the data
    for email, data in contacts.items():
        # Ensure all contacts have consistent fields
        data.setdefault('email_count', 0)
        data.setdefault('sender_score', 0)
        data.setdefault('recommended', False)
        data.setdefault('recommendation_score', 0)
    
    logger.info(f"Built contacts registry with {len(contacts)} contacts")
    return contacts

def infer_insights(contacts: Dict, config: Dict) -> Dict:
    """Infer insights from the contacts registry"""
    insights = {
        'contacts': {
            'total': len(contacts),
            'with_scores': sum(1 for c in contacts.values() if c.get('sender_score', 0) > 0),
            'recommended': sum(1 for c in contacts.values() if c.get('recommended', False)),
            'active': sum(1 for c in contacts.values() if c.get('email_count', 0) >= 5),
            'inactive': sum(1 for c in contacts.values() if c.get('email_count', 0) < 5),
            'high_value': sum(1 for c in contacts.values() if c.get('sender_score', 0) >= 0.7),
            'medium_value': sum(1 for c in contacts.values() if 0.4 <= c.get('sender_score', 0) < 0.7),
            'low_value': sum(1 for c in contacts.values() if 0 < c.get('sender_score', 0) < 0.4),
            'no_score': sum(1 for c in contacts.values() if c.get('sender_score', 0) == 0),
        },
        'email_activity': {
            'total_emails': sum(c.get('email_count', 0) for c in contacts.values()),
            'avg_emails_per_contact': np.mean([c.get('email_count', 0) for c in contacts.values()]),
            'median_emails_per_contact': np.median([c.get('email_count', 0) for c in contacts.values()]),
        },
        'scoring': {
            'avg_score': np.mean([c.get('sender_score', 0) for c in contacts.values() if c.get('sender_score', 0) > 0]),
            'median_score': np.median([c.get('sender_score', 0) for c in contacts.values() if c.get('sender_score', 0) > 0]),
            'recommended_avg_score': np.mean([c.get('recommendation_score', 0) for c in contacts.values() if c.get('recommended', False)]),
        },
        'config_analysis': {
            'min_emails_for_pattern': config.get('min_emails_for_pattern', 5),
            'sender_weight': config.get('sender_weight', 0.4),
            'topic_weight': config.get('topic_weight', 0.25),
            'temporal_weight': config.get('temporal_weight', 0.15),
            'recipient_weight': config.get('recipient_weight', 0.1),
            'state_weight': config.get('state_weight', 0.1),
        }
    }
    
    # Check if weights sum to 1.0
    weights = [
        insights['config_analysis'].get('sender_weight', 0),
        insights['config_analysis'].get('topic_weight', 0),
        insights['config_analysis'].get('temporal_weight', 0),
        insights['config_analysis'].get('recipient_weight', 0),
        insights['config_analysis'].get('state_weight', 0)
    ]
    
    weight_sum = sum(weights)
    insights['config_analysis']['weights_sum'] = weight_sum
    insights['config_analysis']['weights_balanced'] = abs(weight_sum - 1.0) < 0.001
    
    # Calculate score distribution
    scores = [c.get('sender_score', 0) for c in contacts.values() if c.get('sender_score', 0) > 0]
    if scores:
        bins = [0, 0.2, 0.4, 0.6, 0.8, 1.0]
        hist, _ = np.histogram(scores, bins=bins)
        insights['score_distribution'] = {
            'ranges': ['0.0-0.2', '0.2-0.4', '0.4-0.6', '0.6-0.8', '0.8-1.0'],
            'counts': hist.tolist()
        }
    
    # Time-based analysis
    today = datetime.now()
    last_30_days = sum(1 for c in contacts.values() 
                       if c.get('last_seen') and (today - c.get('last_seen')).days <= 30)
    last_90_days = sum(1 for c in contacts.values() 
                       if c.get('last_seen') and (today - c.get('last_seen')).days <= 90)
    
    insights['time_analysis'] = {
        'active_last_30_days': last_30_days,
        'active_last_90_days': last_90_days,
    }
    
    logger.info("Generated insights from contacts data")
    return insights

def calculate_config_sensitivity(contacts: Dict, config: Dict) -> Dict:
    """
    Calculate how sensitive the scores are to configuration changes
    """
    if not contacts or not config:
        return {}
    
    # Only consider contacts with scores
    scored_contacts = {k: v for k, v in contacts.items() if v.get('sender_score', 0) > 0}
    if not scored_contacts:
        return {}
    
    # Get baseline scores
    baseline_scores = {email: data.get('sender_score', 0) for email, data in scored_contacts.items()}
    
    # Define sensitivity tests (parameter adjustments)
    tests = {
        'sender_weight_up': {'param': 'sender_weight', 'change': 0.1},
        'sender_weight_down': {'param': 'sender_weight', 'change': -0.1},
        'topic_weight_up': {'param': 'topic_weight', 'change': 0.1},
        'topic_weight_down': {'param': 'topic_weight', 'change': -0.1},
        'min_emails_up': {'param': 'min_emails_for_pattern', 'change': 1},
        'min_emails_down': {'param': 'min_emails_for_pattern', 'change': -1},
    }
    
    # Calculate potential impact of each test
    sensitivity = {}
    
    # Simulate impact based on score components if available
    for test_name, test_params in tests.items():
        param = test_params['param']
        change = test_params['change']
        
        # Check if the parameter exists in config
        if param not in config:
            continue
            
        # Estimate impact
        impacts = []
        for email, data in scored_contacts.items():
            components = data.get('score_components', {})
            if not components:
                continue
                
            if param.endswith('weight') and components:
                # For weights, estimate the impact on final score
                component_name = param.replace('_weight', '')
                component_score = components.get(component_name, 0)
                current_weight = config.get(param, 0)
                new_weight = max(0, min(1, current_weight + change))
                
                # Simple impact calculation
                impact = (new_weight - current_weight) * component_score
                impacts.append(abs(impact))
        
        if impacts:
            sensitivity[test_name] = {
                'parameter': param,
                'current_value': config.get(param, 0),
                'test_value': config.get(param, 0) + change,
                'avg_impact': np.mean(impacts),
                'max_impact': max(impacts) if impacts else 0
            }
    
    logger.info(f"Calculated configuration sensitivity for {len(sensitivity)} tests")
    return sensitivity

def generate_contact_summaries(contacts: Dict, top_n: int = 20) -> Dict:
    """
    Generate summary information for top contacts
    """
    if not contacts:
        return {}
    
    # Sort contacts by sender_score (descending)
    sorted_contacts = sorted(
        contacts.values(),
        key=lambda x: x.get('sender_score', 0),
        reverse=True
    )
    
    # Get top N high-scoring contacts
    top_contacts = sorted_contacts[:top_n]
    
    # Generate summaries for top contacts
    contact_summaries = []
    for contact in top_contacts:
        email = contact.get('email', '')
        if not email:
            continue
            
        summary = {
            'email': email,
            'sender_score': contact.get('sender_score', 0),
            'email_count': contact.get('email_count', 0),
            'recommended': contact.get('recommended', False),
            'recommendation_score': contact.get('recommendation_score', 0),
            'first_seen': contact.get('first_seen'),
            'last_seen': contact.get('last_seen'),
            'days_active': contact.get('days_active', 0),
            'score_components': contact.get('score_components', {})
        }
        contact_summaries.append(summary)
    
    logger.info(f"Generated summaries for {len(contact_summaries)} top contacts")
    return contact_summaries

def save_json_output(data: Dict, output_path: str) -> None:
    """Save the diagnostics data as JSON"""
    try:
        with open(output_path, 'w') as f:
            json.dump(clean_data_for_json(data), f, indent=2, default=convert_datetime_for_json)
        logger.info(f"Saved JSON output to {output_path}")
    except Exception as e:
        logger.error(f"Error saving JSON output: {str(e)}")

def save_html_output(data: Dict, output_path: str) -> None:
    """Save the diagnostics data as an HTML report"""
    try:
        insights = data.get('insights', {})
        contacts = data.get('contacts', {})
        top_contacts = data.get('top_contacts', [])
        
        # Create HTML content
        html_content = [
            "<!DOCTYPE html>",
            "<html lang='en'>",
            "<head>",
            "  <meta charset='UTF-8'>",
            "  <meta name='viewport' content='width=device-width, initial-scale=1.0'>",
            "  <title>Email Analysis Diagnostics Report</title>",
            "  <style>",
            "    body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }",
            "    h1, h2, h3 { color: #333; }",
            "    .container { max-width: 1200px; margin: 0 auto; }",
            "    .card { border: 1px solid #ddd; border-radius: 4px; padding: 15px; margin-bottom: 20px; }",
            "    .card h3 { margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 10px; }",
            "    table { border-collapse: collapse; width: 100%; }",
            "    th, td { text-align: left; padding: 8px; border-bottom: 1px solid #ddd; }",
            "    th { background-color: #f2f2f2; }",
            "    tr:hover { background-color: #f5f5f5; }",
            "    .metric { display: inline-block; width: 200px; margin: 10px; padding: 15px; ",
            "             border: 1px solid #ddd; border-radius: 4px; text-align: center; }",
            "    .metric-value { font-size: 24px; font-weight: bold; margin: 10px 0; }",
            "    .metric-label { font-size: 14px; color: #666; }",
            "    .chart { height: 200px; margin: 20px 0; }",
            "  </style>",
            "</head>",
            "<body>",
            "  <div class='container'>",
            f"    <h1>Email Analysis Diagnostics Report</h1>",
            f"    <p>Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>"
        ]
        
        # Contact metrics section
        if insights and 'contacts' in insights:
            html_content.extend([
                "    <div class='card'>",
                "      <h3>Contact Metrics</h3>",
                "      <div style='display: flex; flex-wrap: wrap;'>",
                f"        <div class='metric'><div class='metric-value'>{insights['contacts'].get('total', 0)}</div><div class='metric-label'>Total Contacts</div></div>",
                f"        <div class='metric'><div class='metric-value'>{insights['contacts'].get('with_scores', 0)}</div><div class='metric-label'>Contacts with Scores</div></div>",
                f"        <div class='metric'><div class='metric-value'>{insights['contacts'].get('recommended', 0)}</div><div class='metric-label'>Recommended Contacts</div></div>",
                f"        <div class='metric'><div class='metric-value'>{insights['contacts'].get('high_value', 0)}</div><div class='metric-label'>High-Value Contacts</div></div>",
                f"        <div class='metric'><div class='metric-value'>{insights['email_activity'].get('total_emails', 0)}</div><div class='metric-label'>Total Emails</div></div>",
                "      </div>",
                "    </div>"
            ])
        
        # Score distribution
        if insights and 'score_distribution' in insights:
            ranges = insights['score_distribution'].get('ranges', [])
            counts = insights['score_distribution'].get('counts', [])
            
            if ranges and counts and len(ranges) == len(counts):
                html_content.extend([
                    "    <div class='card'>",
                    "      <h3>Score Distribution</h3>",
                    "      <table>",
                    "        <tr>",
                    "          <th>Score Range</th>",
                    "          <th>Number of Contacts</th>",
                    "        </tr>"
                ])
                
                for i, range_label in enumerate(ranges):
                    html_content.append(f"        <tr><td>{range_label}</td><td>{counts[i]}</td></tr>")
                
                html_content.append("      </table>")
                html_content.append("    </div>")
        
        # Configuration analysis
        if insights and 'config_analysis' in insights:
            config = insights['config_analysis']
            html_content.extend([
                "    <div class='card'>",
                "      <h3>Configuration Analysis</h3>",
                "      <table>",
                "        <tr><th>Parameter</th><th>Value</th></tr>",
                f"        <tr><td>Sender Weight</td><td>{config.get('sender_weight', 0)}</td></tr>",
                f"        <tr><td>Topic Weight</td><td>{config.get('topic_weight', 0)}</td></tr>",
                f"        <tr><td>Temporal Weight</td><td>{config.get('temporal_weight', 0)}</td></tr>",
                f"        <tr><td>Recipient Weight</td><td>{config.get('recipient_weight', 0)}</td></tr>",
                f"        <tr><td>State Weight</td><td>{config.get('state_weight', 0)}</td></tr>",
                f"        <tr><td>Min Emails for Pattern</td><td>{config.get('min_emails_for_pattern', 5)}</td></tr>",
                f"        <tr><td>Weights Sum to 1.0</td><td>{config.get('weights_balanced', False)}</td></tr>",
                "      </table>",
                "    </div>"
            ])
        
        # Top contacts
        if top_contacts:
            html_content.extend([
                "    <div class='card'>",
                "      <h3>Top Contacts</h3>",
                "      <table>",
                "        <tr>",
                "          <th>Email</th>",
                "          <th>Score</th>",
                "          <th>Email Count</th>",
                "          <th>Recommended</th>",
                "          <th>Last Seen</th>",
                "        </tr>"
            ])
            
            for contact in top_contacts:
                last_seen = contact.get('last_seen')
                last_seen_str = last_seen.strftime('%Y-%m-%d') if isinstance(last_seen, datetime) else 'N/A'
                html_content.append(
                    f"        <tr>"
                    f"<td>{contact.get('email', '')}</td>"
                    f"<td>{contact.get('sender_score', 0):.3f}</td>"
                    f"<td>{contact.get('email_count', 0)}</td>"
                    f"<td>{'Yes' if contact.get('recommended', False) else 'No'}</td>"
                    f"<td>{last_seen_str}</td>"
                    f"</tr>"
                )
            
            html_content.append("      </table>")
            html_content.append("    </div>")
        
        # Close tags
        html_content.extend([
            "  </div>",
            "</body>",
            "</html>"
        ])
        
        # Save HTML file
        with open(output_path, 'w') as f:
            f.write('\n'.join(html_content))
        
        logger.info(f"Saved HTML report to {output_path}")
    except Exception as e:
        logger.error(f"Error saving HTML output: {str(e)}")

def save_csv_output(data: Dict, output_path: str) -> None:
    """Save the contacts data as CSV"""
    try:
        # Save contacts as CSV
        contacts = data.get('contacts', {})
        if contacts:
            contacts_list = []
            for email, contact_data in contacts.items():
                # Clean and flatten the data
                cleaned_data = {}
                for key, value in contact_data.items():
                    if key == 'score_components':
                        # Flatten score components
                        for comp_key, comp_value in value.items():
                            cleaned_data[f"component_{comp_key}"] = comp_value
                    elif isinstance(value, datetime):
                        cleaned_data[key] = value.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        cleaned_data[key] = value
                
                contacts_list.append(cleaned_data)
            
            # Convert to DataFrame and save as CSV
            if contacts_list:
                df = pd.DataFrame(contacts_list)
                df.to_csv(output_path, index=False)
                logger.info(f"Saved contacts CSV to {output_path}")
            else:
                logger.warning("No contact data to save to CSV")
        else:
            logger.warning("No contacts data to save to CSV")
    except Exception as e:
        logger.error(f"Error saving CSV output: {str(e)}")

def main():
    """Main function to run diagnostics"""
    parser = argparse.ArgumentParser(description='Email Analysis Diagnostics Utility')
    
    parser.add_argument('--config', type=str, default='config.json',
                        help='Path to configuration file')
    parser.add_argument('--data-dir', type=str, default='email_data',
                        help='Directory containing pickle data files')
    parser.add_argument('--reports-dir', type=str, default='.',
                        help='Directory containing recommendation reports')
    parser.add_argument('--output-dir', type=str, default='diagnostics_output',
                        help='Directory for diagnostic output files')
    parser.add_argument('--output-format', type=str, nargs='+', 
                        choices=['json', 'html', 'csv'], default=['json'],
                        help='Output format(s) for diagnostics data')
    parser.add_argument('--top-contacts', type=int, default=20,
                        help='Number of top contacts to include in detailed output')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Ensure output directory exists
    os.makedirs(args.output_dir, exist_ok=True)
    
    try:
        # 1. Load configuration
        config = load_config(args.config)
        
        # 2. Load data files
        data_dir = args.data_dir
        sender_scores_path = os.path.join(data_dir, 'sender_scores.pkl')
        email_tracking_path = os.path.join(data_dir, 'email_tracking.pkl')
        
        sender_scores = load_pickle_file(sender_scores_path) or {}
        email_tracking = load_pickle_file(email_tracking_path) or {}
        
        # 3. Load latest recommendation report
        latest_report = find_latest_report(args.reports_dir)
        recommendations_df = load_csv_to_dataframe(latest_report) if latest_report else pd.DataFrame()
        
        # 4. Build contacts registry
        contacts = build_contacts_registry(sender_scores, email_tracking, recommendations_df)
        
        # 5. Generate insights
        insights = infer_insights(contacts, config)
        
        # 6. Calculate config sensitivity
        sensitivity = calculate_config_sensitivity(contacts, config)
        
        # 7. Generate contact summaries
        top_contacts = generate_contact_summaries(contacts, args.top_contacts)
        
        # 8. Compile diagnostics data
        diagnostics_data = {
            'timestamp': datetime.now(),
            'contacts': contacts,
            'insights': insights,
            'config_sensitivity': sensitivity,
            'top_contacts': top_contacts,
            'data_files': {
                'sender_scores': sender_scores_path,
                'email_tracking': email_tracking_path,
                'latest_report': latest_report
            }
        }
        
        # 9. Save outputs in requested formats
        for output_format in args.output_format:
            output_filename = f"diagnostics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{output_format}"
            output_path = os.path.join(args.output_dir, output_filename)
            
            if output_format == 'json':
                save_json_output(diagnostics_data, output_path)
            elif output_format == 'html':
                save_html_output(diagnostics_data, output_path)
            elif output_format == 'csv':
                save_csv_output(diagnostics_data, output_path)
        
        logger.info(f"Diagnostics completed successfully. Output files saved to {args.output_dir}")
        
    except Exception as e:
        logger.error(f"Error running diagnostics: {str(e)}")
        import traceback
        logger.debug(traceback.format_exc())
        sys.exit(1)

if __name__ == '__main__':
    main() 