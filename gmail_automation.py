"""
Gmail Email Automation Script
Sends emails to people listed in Dev.xlsx file
"""

import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import getpass
import json
from datetime import datetime
import logging

class GmailAutomation:
    def __init__(self, config_file="email_config.json"):
        """Initialize the Gmail automation with configuration"""
        self.config_file = config_file
        self.config = self.load_config()
        
        # Set up logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('email_automation.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def load_config(self):
        """Load email configuration from JSON file"""
        default_config = {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "excel_file": "Dev.xlsx",
            "email_column": "email",  # Update this based on your Excel column name
            "name_column": "name",    # Update this based on your Excel column name
            "active_column": "active",  # Column name for active flag (1 = send email, 0 = skip)
            "subject_template": "Application for Data Engineer role",
            "body_template_file": "email_template.txt",  # Path to email template file
            "attachment_file": "Resume_Mohit.pdf"  # Path to attachment file
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                # Merge with defaults
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
            except Exception as e:
                self.logger.warning(f"Error loading config file: {e}. Using defaults.")
                return default_config
        else:
            # Create default config file
            with open(self.config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            self.logger.info(f"Created default config file: {self.config_file}")
            return default_config
    
    def load_email_template(self):
        """Load email template from file"""
        try:
            template_file = self.config.get("body_template_file", "email_template.txt")
            if os.path.exists(template_file):
                with open(template_file, 'r', encoding='utf-8') as f:
                    template = f.read()
                self.logger.info(f"Email template loaded from {template_file}")
                return template
            else:
                self.logger.warning(f"Template file '{template_file}' not found. Using default template.")
                return "Hi,\n\nThis is an automated email.\n\nBest regards,\nMohit Saini\nmohitmssaini1@gmail.com"
        except Exception as e:
            self.logger.error(f"Error loading email template: {e}")
            return "Hi,\n\nThis is an automated email.\n\nBest regards,\nMohit Saini\nmohitmssaini1@gmail.com"
    
    def read_excel_data(self):
        """Read data from Excel file"""
        try:
            if not os.path.exists(self.config["excel_file"]):
                raise FileNotFoundError(f"Excel file '{self.config['excel_file']}' not found")
            
            # Try to read the Excel file
            df = pd.read_excel(self.config["excel_file"])
            
            # Display file structure for user verification
            self.logger.info(f"Excel file '{self.config['excel_file']}' loaded successfully!")
            self.logger.info(f"Columns found: {list(df.columns)}")
            self.logger.info(f"Number of rows: {len(df)}")
            
            # Check if required columns exist
            if self.config["email_column"] not in df.columns:
                self.logger.warning(f"Email column '{self.config['email_column']}' not found.")
                self.logger.info("Available columns: " + ", ".join(df.columns))
                # Try to find email column automatically
                email_cols = [col for col in df.columns if 'email' in col.lower() or 'mail' in col.lower()]
                if email_cols:
                    self.config["email_column"] = email_cols[0]
                    self.logger.info(f"Using '{email_cols[0]}' as email column")
                else:
                    raise ValueError("No email column found. Please update the config file.")
            
            if self.config["name_column"] not in df.columns:
                self.logger.warning(f"Name column '{self.config['name_column']}' not found.")
                # Try to find name column automatically
                name_cols = [col for col in df.columns if 'name' in col.lower()]
                if name_cols:
                    self.config["name_column"] = name_cols[0]
                    self.logger.info(f"Using '{name_cols[0]}' as name column")
                else:
                    # Use first column as name
                    self.config["name_column"] = df.columns[0]
                    self.logger.info(f"Using '{df.columns[0]}' as name column")
            
            # Remove rows with empty email addresses
            df = df.dropna(subset=[self.config["email_column"]])
            df = df[df[self.config["email_column"]].str.strip() != ""]
            
            # Filter for Active flag = 1 (if applicable)
            active_column = None
            
            # Apply Active flag filtering for both Dev.xlsx and Prod.xlsx
            # First try the configured active column
            if self.config["active_column"] in df.columns:
                active_column = self.config["active_column"]
                self.logger.info(f"Found configured active column: '{active_column}'")
            else:
                # Look for Active column (case-insensitive)
                self.logger.info(f"Looking for active column (configured: '{self.config['active_column']}')")
                for col in df.columns:
                    if 'active' in col.lower():
                        active_column = col
                        self.config["active_column"] = col  # Update config with found column
                        self.logger.info(f"Found active column: '{active_column}'")
                        break
            
            if active_column:
                original_count = len(df)
                self.logger.info(f"Active column '{active_column}' found with values: {df[active_column].value_counts().to_dict()}")
                
                # Filter for Active = 1 (handle both numeric and string values)
                # First try numeric comparison
                try:
                    df_filtered = df[df[active_column] == 1]
                    if len(df_filtered) == 0:
                        # If no numeric 1s, try string comparison
                        df_filtered = df[df[active_column].astype(str).str.strip().isin(['1', '1.0', 'True', 'true', 'YES', 'yes'])]
                    df = df_filtered
                except:
                    # Fallback to string comparison only
                    df = df[df[active_column].astype(str).str.strip().isin(['1', '1.0', 'True', 'true', 'YES', 'yes'])]
                
                filtered_count = len(df)
                self.logger.info(f"Filtered by Active flag: {original_count} -> {filtered_count} records")
                self.logger.info(f"Using column '{active_column}' for Active flag filtering")
                
                if filtered_count == 0:
                    self.logger.warning("No active records found! All records have Active = 0")
                    return pd.DataFrame()  # Return empty dataframe
            else:
                self.logger.warning(f"Active column '{self.config['active_column']}' not found.")
                self.logger.info("Available columns: " + ", ".join(df.columns))
                # Ask user if they want to continue
                response = input("No Active flag column found. Send to all recipients? (yes/no): ")
                if response.lower() != 'yes':
                    self.logger.info("Email sending cancelled - no Active flag column.")
                    return pd.DataFrame()  # Return empty dataframe
            
            self.logger.info(f"Found {len(df)} recipients after filtering")
            return df
            
        except Exception as e:
            self.logger.error(f"Error reading Excel file: {e}")
            raise
    
    def get_gmail_credentials(self):
        """Get Gmail credentials from user"""
        print("\n" + "="*50)
        print("GMAIL AUTHENTICATION")
        print("="*50)
        print("For Gmail authentication, you need an App Password.")
        print("To create an App Password:")
        print("1. Go to your Google Account settings")
        print("2. Security > 2-Step Verification (must be enabled)")
        print("3. App passwords > Generate a new app password")
        print("4. Select 'Mail' and your device")
        print("5. Use the generated 16-character password below")
        print("="*50)
        
        email = input("Enter your Gmail address: ").strip()
        password = getpass.getpass("Enter your Gmail App Password: ")
        
        return email, password
    
    def create_email(self, recipient_email, recipient_name, sender_email):
        """Create email message with optional attachment"""
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        
        # Set subject (no personalization needed for this template)
        subject = self.config["subject_template"]
        msg['Subject'] = subject
        
        # Load email body from template file
        body = self.load_email_template()
        
        # The template is static and doesn't need name personalization
        # as it's a professional introduction email
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Add attachment if specified in config
        if "attachment_file" in self.config and self.config["attachment_file"]:
            attachment_path = self.config["attachment_file"]
            if os.path.exists(attachment_path):
                try:
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    filename = os.path.basename(attachment_path)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {filename}',
                    )
                    msg.attach(part)
                    self.logger.info(f"Attached file: {filename}")
                except Exception as e:
                    self.logger.warning(f"Failed to attach file {attachment_path}: {e}")
            else:
                self.logger.warning(f"Attachment file not found: {attachment_path}")
        
        return msg
    
    def send_emails(self, test_mode=True, max_emails=None):
        """Send emails to all recipients"""
        try:
            # Read Excel data
            df = self.read_excel_data()
            
            if len(df) == 0:
                self.logger.error("No valid email addresses found in Excel file")
                return
            
            # Check attachment file
            if "attachment_file" in self.config and self.config["attachment_file"]:
                attachment_path = self.config["attachment_file"]
                if os.path.exists(attachment_path):
                    file_size = os.path.getsize(attachment_path) / (1024 * 1024)  # Size in MB
                    self.logger.info(f"Attachment file found: {attachment_path} ({file_size:.2f} MB)")
                else:
                    self.logger.warning(f"Attachment file not found: {attachment_path}")
                    response = input(f"Attachment file '{attachment_path}' not found. Continue without attachment? (yes/no): ")
                    if response.lower() != 'yes':
                        self.logger.info("Email sending cancelled due to missing attachment.")
                        return

            # Get Gmail credentials
            sender_email, password = self.get_gmail_credentials()
            
            # Limit emails in test mode
            if test_mode and max_emails:
                df = df.head(max_emails)
                self.logger.info(f"Test mode: Sending to first {len(df)} recipients")
            
            # Create secure connection
            context = ssl.create_default_context()
            
            # Connect to Gmail SMTP server
            self.logger.info("Connecting to Gmail SMTP server...")
            with smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"]) as server:
                server.starttls(context=context)
                server.login(sender_email, password)
                self.logger.info("Successfully connected to Gmail!")
                
                sent_count = 0
                failed_count = 0
                
                for index, row in df.iterrows():
                    try:
                        recipient_email = str(row[self.config["email_column"]]).strip()
                        recipient_name = str(row[self.config["name_column"]]).strip() if pd.notna(row[self.config["name_column"]]) else ""
                        
                        if not recipient_email or recipient_email.lower() == 'nan':
                            self.logger.warning(f"Skipping row {index + 1}: No email address")
                            continue
                        
                        # Clean up name field
                        if recipient_name.lower() == 'nan' or not recipient_name:
                            recipient_name = ""
                        
                        # Create and send email
                        msg = self.create_email(recipient_email, recipient_name, sender_email)
                        
                        if test_mode:
                            self.logger.info(f"TEST MODE - Would send to: {recipient_email} ({recipient_name})")
                            sent_count += 1
                        else:
                            server.send_message(msg)
                            self.logger.info(f"Email sent to: {recipient_email} ({recipient_name})")
                            sent_count += 1
                        
                    except Exception as e:
                        failed_count += 1
                        self.logger.error(f"Failed to send email to {recipient_email}: {e}")
                
                self.logger.info(f"\nEmail sending completed!")
                self.logger.info(f"Successfully sent: {sent_count}")
                self.logger.info(f"Failed: {failed_count}")
                
        except Exception as e:
            self.logger.error(f"Error in send_emails: {e}")
            raise

def main():
    """Main function to run the email automation"""
    print("Gmail Email Automation")
    print("=" * 30)
    
    try:
        # Ask user to select Excel file
        print("\nSelect Excel file:")
        print("1. Dev.xlsx (Development/Test data)")
        print("2. Prod.xlsx (Production data)")
        
        while True:
            file_choice = input("Enter your choice (1 or 2): ").strip()
            if file_choice in ['1', '2']:
                break
            print("Please enter 1 or 2")
        
        # Set the Excel file based on choice
        excel_file = "Dev.xlsx" if file_choice == '1' else "Prod.xlsx"
        print(f"Selected file: {excel_file}")
        
        # Initialize automation
        automation = GmailAutomation()
        
        # Override the excel file in config
        automation.config["excel_file"] = excel_file
        
        # Set appropriate column names based on file choice
        if excel_file == "Prod.xlsx":
            automation.config["email_column"] = "Email"
            automation.config["name_column"] = "Name"
            automation.config["active_column"] = "Active"  # Prod.xlsx also has Active column!
            print("Note: Prod.xlsx will use Active flag filtering (only Active = 1 will receive emails)")
        else:
            automation.config["email_column"] = "email"
            automation.config["name_column"] = "name"
            automation.config["active_column"] = "active"
            print("Note: Dev.xlsx will use Active flag filtering (only Active = 1 will receive emails)")
        
        # Ask user for mode
        print("\nSelect mode:")
        print("1. Test mode (preview emails without sending)")
        print("2. Send emails")
        
        while True:
            choice = input("Enter your choice (1 or 2): ").strip()
            if choice in ['1', '2']:
                break
            print("Please enter 1 or 2")
        
        test_mode = choice == '1'
        
        if test_mode:
            max_emails = input("How many emails to preview? (Enter number or press Enter for all): ").strip()
            max_emails = int(max_emails) if max_emails.isdigit() else None
            automation.send_emails(test_mode=True, max_emails=max_emails)
        else:
            confirm = input("Are you sure you want to send emails? (yes/no): ").strip().lower()
            if confirm == 'yes':
                automation.send_emails(test_mode=False)
            else:
                print("Email sending cancelled.")
    
    except Exception as e:
        print(f"Error: {e}")
        print("Please check the log file 'email_automation.log' for more details.")

if __name__ == "__main__":
    main()