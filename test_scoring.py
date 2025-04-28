"""
Simple test script for the enhanced email scoring logic
"""

import os
import sys
import logging
from datetime import datetime, timedelta

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Import required modules
from config import load_config
from scorer import score_email, recommend_action

class MockEmailItem:
    """Mock email item for testing"""
    def __init__(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self, key, value)
        
    def __str__(self):
        return f"MockEmailItem(Subject='{getattr(self, 'Subject', 'No Subject')}')"

class MockSession:
    """Mock Outlook session"""
    def __init__(self, current_user_address="user@example.com"):
        self.CurrentUser = MockCurrentUser(current_user_address)

class MockCurrentUser:
    """Mock current user"""
    def __init__(self, address="user@example.com"):
        self.Address = address

def test_scoring_with_message_state():
    """Test scoring with message state factors"""
    config = load_config()
    
    # Mock data
    sender_scores = {
        "important@example.com": {"normalized_score": 0.8},
        "regular@example.com": {"normalized_score": 0.5},
        "low@example.com": {"normalized_score": 0.3}
    }
    
    email_patterns = {
        "sender_read_ratios": {
            "important@example.com": {"read_ratio": 0.9, "reopen_ratio": 0.6},
            "regular@example.com": {"read_ratio": 0.7, "reopen_ratio": 0.2}
        }
    }
    
    now = datetime.now()
    yesterday = now - timedelta(days=1)
    
    # Create test cases with different properties
    test_cases = [
        # Regular email
        MockEmailItem(
            Subject="Regular Test Email",
            Body="This is a test email with regular content.",
            SenderEmailAddress="regular@example.com",
            ReceivedTime=yesterday,
            Unread=False,
            FlagStatus=0,
            Importance=1,
            ConversationID="conv1",
            Session=MockSession()
        ),
        
        # High priority email - unread and flagged
        MockEmailItem(
            Subject="Important Unread Flagged Email",
            Body="This is a high priority email that needs action ASAP.",
            SenderEmailAddress="important@example.com",
            ReceivedTime=now,
            Unread=True,
            FlagStatus=2,  # olFlagMarked
            FlagDueBy=now + timedelta(days=2),
            Importance=2,  # olImportanceHigh
            ConversationID="conv2",
            To="user@example.com",
            CC="",
            Recipients=MockEmailItem(Count=3),
            Session=MockSession()
        ),
        
        # Due today email
        MockEmailItem(
            Subject="Due Today Email",
            Body="This email has a task due today.",
            SenderEmailAddress="regular@example.com",
            ReceivedTime=yesterday,
            Unread=False,
            FlagStatus=2,  # olFlagMarked
            FlagDueBy=now.replace(hour=0, minute=0, second=0, microsecond=0),
            Importance=1,
            ConversationID="conv3",
            Session=MockSession()
        ),
        
        # CC'd mass email
        MockEmailItem(
            Subject="CC Mass Email",
            Body="This is a mass email where you're CC'd.",
            SenderEmailAddress="low@example.com",
            ReceivedTime=yesterday,
            Unread=True,
            FlagStatus=0,
            Importance=1,
            To="team@example.com",
            CC="user@example.com",
            Recipients=MockEmailItem(Count=15),
            Session=MockSession()
        ),
        
        # Off-hours email
        MockEmailItem(
            Subject="Off-hours Email",
            Body="This email was received outside business hours.",
            SenderEmailAddress="regular@example.com",
            ReceivedTime=yesterday.replace(hour=22, minute=30),  # Set to 10:30 PM
            Unread=False,
            FlagStatus=0,
            Importance=1,
            Session=MockSession()
        )
    ]
    
    # Test each case
    for i, email in enumerate(test_cases):
        logger.info(f"\nTesting case {i+1}: {email.Subject}")
        
        # Score the email
        score_data = score_email(email, sender_scores, email_patterns, config)
        
        if score_data:
            # Print score components
            logger.info(f"Final score: {score_data['final_score']:.2f}")
            
            components = score_data['components']
            for comp_name, comp_value in components.items():
                logger.info(f"  - {comp_name}: {comp_value:.2f}")
            
            # Get recommendation
            recommendation = recommend_action(email, sender_scores, email_patterns, config)
            
            logger.info(f"Recommendation:")
            logger.info(f"  - Folder: {recommendation['folder']}")
            logger.info(f"  - Flag: {recommendation['flag']}")
            logger.info(f"  - Create task: {recommendation['create_task']}")
            logger.info(f"  - Auto archive: {recommendation['auto_archive']}")
        else:
            logger.error(f"Scoring failed for email: {email.Subject}")

def main():
    """Run tests"""
    logger.info("Starting email scoring tests...")
    test_scoring_with_message_state()
    logger.info("Testing complete.")

if __name__ == "__main__":
    main() 