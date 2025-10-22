import gspread
from google.auth import default
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st
import json
import os

class GoogleSheetsService:
    def __init__(self):
        self.client = None
        self.credentials = None
        
    def authenticate_with_service_account(self, credentials_json):
        """Authenticate using service account credentials JSON"""
        try:
            # Helper to validate required fields in credentials dict
            def _validate_credentials_dict(creds: dict) -> bool:
                required_keys = ['type', 'project_id', 'private_key', 'client_email', 'token_uri']
                missing = [k for k in required_keys if k not in creds or not creds.get(k)]
                if missing:
                    st.error(
                        "google_service_account secrets are incomplete. Missing: " + ", ".join(missing)
                    )
                    return False
                return True

            # Parse credentials from JSON string or file
            if isinstance(credentials_json, str):
                if os.path.isfile(credentials_json):
                    # It's a file path
                    credentials = Credentials.from_service_account_file(
                        credentials_json,
                        scopes=[
                            'https://www.googleapis.com/auth/spreadsheets',
                            'https://www.googleapis.com/auth/drive'
                        ]
                    )
                else:
                    # It's a JSON string
                    credentials_dict = json.loads(credentials_json)
                    if not _validate_credentials_dict(credentials_dict):
                        return False
                    credentials = Credentials.from_service_account_info(
                        credentials_dict,
                        scopes=[
                            'https://www.googleapis.com/auth/spreadsheets',
                            'https://www.googleapis.com/auth/drive'
                        ]
                    )
            else:
                # It's already a dict
                if not _validate_credentials_dict(credentials_json):
                    return False
                credentials = Credentials.from_service_account_info(
                    credentials_json,
                    scopes=[
                        'https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive'
                    ]
                )
            
            self.client = gspread.authorize(credentials)
            self.credentials = credentials
            return True
            
        except Exception as e:
            st.error(f"Authentication failed: {str(e)}")
            return False
    
    def upload_quali_nv_to_google_sheets(self, excel_path="excel/finale_jour.xlsx", spreadsheet_id="1cRNqohML-mZ2mMqXIfQlPPDQUmRGLmRzDfp3j0wpte4", worksheet_name="quali SOM VMM"):
        """
        Upload QUALI NV sheet data to specific Google Sheets ID and worksheet
        """
        try:
            # Get credentials from Streamlit secrets
            try:
                credentials_dict = dict(st.secrets["google_service_account"])
            except KeyError:
                print("Error: google_service_account not found in secrets. Please add your service account credentials to secrets.toml")
                return False
            
            # Authenticate with service account
            if not self.authenticate_with_service_account(credentials_dict):
                print("Error: Failed to authenticate with Google Sheets")
                return False
            
            print("✅ Authentication successful!")
            
            # Read the processed Excel file - specifically QUALI NV sheet
            if not os.path.exists(excel_path):
                print("Error: Processed Excel file not found. Please process the file first.")
                return False
            
            # Convert QUALI NV sheet to DataFrame
            df_quali = pd.read_excel(excel_path, sheet_name='QUALI NV')
            print(f"✅ QUALI NV sheet loaded with {len(df_quali)} rows")
            
            # Open the specific spreadsheet by ID
            spreadsheet = self.open_spreadsheet(spreadsheet_id)
            if not spreadsheet:
                print(f"Error: Failed to open spreadsheet with ID: {spreadsheet_id}")
                return False
            
            print(f"✅ Spreadsheet opened successfully!")
            
            # Upload QUALI NV data to the specific worksheet
            if self.upload_dataframe_to_sheet(spreadsheet, worksheet_name, df_quali):
                print(f"✅ QUALI NV data uploaded successfully to worksheet '{worksheet_name}'!")
                
                # Get and display spreadsheet URL
                spreadsheet_url = self.get_spreadsheet_url(spreadsheet)
                if spreadsheet_url:
                    print(f"🎉 QUALI NV upload completed!")
                    print(f"📊 Spreadsheet URL: {spreadsheet_url}")
                    return spreadsheet_url
                else:
                    return True
            else:
                print(f"Error: Failed to upload QUALI NV data to worksheet '{worksheet_name}'")
                return False
                
        except Exception as e:
            print(f"Error uploading QUALI NV to Google Sheets: {e}")
            return False
    
    def create_spreadsheet(self, title):
        """Create a new Google Spreadsheet"""
        try:
            if not self.client:
                raise Exception("Not authenticated. Please authenticate first.")
            
            spreadsheet = self.client.create(title)
            return spreadsheet
            
        except Exception as e:
            st.error(f"Failed to create spreadsheet: {str(e)}")
            return None
    
    def open_spreadsheet(self, spreadsheet_id_or_url):
        """Open an existing Google Spreadsheet by ID or URL"""
        try:
            if not self.client:
                raise Exception("Not authenticated. Please authenticate first.")
            
            # Extract spreadsheet ID from URL if needed
            if 'docs.google.com/spreadsheets' in spreadsheet_id_or_url:
                # Extract ID from URL
                spreadsheet_id = spreadsheet_id_or_url.split('/d/')[1].split('/')[0]
            else:
                spreadsheet_id = spreadsheet_id_or_url
            
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            return spreadsheet
            
        except Exception as e:
            print(f"Failed to open spreadsheet: {str(e)}")
            return None
    
    def upload_dataframe_to_sheet(self, spreadsheet, worksheet_name, dataframe):
        """Upload a pandas DataFrame to a specific worksheet"""
        try:
            if not spreadsheet:
                raise Exception("No spreadsheet provided")
            
            # Get all worksheets first
            all_worksheets = spreadsheet.worksheets()
            worksheet_names = [ws.title for ws in all_worksheets]
            print(f"Available worksheets: {worksheet_names}")
            
            # Try to find existing worksheet (case-insensitive)
            worksheet = None
            for ws in all_worksheets:
                if ws.title.lower() == worksheet_name.lower():
                    worksheet = ws
                    print(f"✅ Found existing worksheet: {ws.title}")
                    break
            
            if worksheet:
                # Clear existing data
                worksheet.clear()
                print("✅ Cleared existing data")
            else:
                print(f"Creating new worksheet: {worksheet_name}")
                # Create new worksheet
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=26)
                print("✅ New worksheet created")
            
            # Convert DataFrame to list of lists for gspread
            # Replace NaN values with empty strings to avoid JSON errors
            dataframe_clean = dataframe.fillna('')
            data = [dataframe_clean.columns.tolist()] + dataframe_clean.values.tolist()
            print(f"✅ Prepared data: {len(data)} rows, {len(data[0])} columns")
            
            # Upload data
            worksheet.update('A1', data)
            print("✅ Data uploaded successfully")
            
            return True
            
        except Exception as e:
            print(f"Failed to upload data to sheet: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def share_spreadsheet(self, spreadsheet, email, role='writer'):
        """Share spreadsheet with an email address"""
        try:
            if not spreadsheet:
                raise Exception("No spreadsheet provided")
            
            spreadsheet.share(email, perm_type='user', role=role)
            return True
            
        except Exception as e:
            st.error(f"Failed to share spreadsheet: {str(e)}")
            return False
    
    def get_spreadsheet_url(self, spreadsheet):
        """Get the URL of a spreadsheet"""
        try:
            return f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
        except Exception as e:
            print(f"Error getting spreadsheet URL: {str(e)}")
            return None
    
    def upload_excel_to_google_sheets(self, excel_path="excel/finale_jour.xlsx", spreadsheet_id="1cRNqohML-mZ2mMqXIfQlPPDQUmRGLmRzDfp3j0wpte4", worksheet_name="Suivi Test"):
        """
        Upload an Excel file to specific Google Sheets ID and worksheet
        """
        try:
            # Get credentials from Streamlit secrets
            try:
                credentials_dict = dict(st.secrets["google_service_account"])
            except KeyError:
                print("Error: google_service_account not found in secrets. Please add your service account credentials to secrets.toml")
                return False
            
            # Authenticate with service account
            if not self.authenticate_with_service_account(credentials_dict):
                print("Error: Failed to authenticate with Google Sheets")
                return False
            
            print("✅ Authentication successful!")
            
            # Read the processed Excel file
            if not os.path.exists(excel_path):
                print("Error: Processed Excel file not found. Please process the file first.")
                return False
            
            # Convert Excel to DataFrame
            df = pd.read_excel(excel_path)
            print(f"✅ Excel file loaded with {len(df)} rows")
            
            # Open the specific spreadsheet by ID
            spreadsheet = self.open_spreadsheet(spreadsheet_id)
            if not spreadsheet:
                print(f"Error: Failed to open spreadsheet with ID: {spreadsheet_id}")
                return False
            
            print(f"✅ Spreadsheet opened successfully!")
            
            # Upload data to the specific worksheet
            if self.upload_dataframe_to_sheet(spreadsheet, worksheet_name, df):
                print(f"✅ Data uploaded successfully to worksheet '{worksheet_name}'!")
                
                # Get and display spreadsheet URL
                spreadsheet_url = self.get_spreadsheet_url(spreadsheet)
                if spreadsheet_url:
                    print(f"🎉 Upload completed!")
                    print(f"📊 Spreadsheet URL: {spreadsheet_url}")
                    return spreadsheet_url
                else:
                    print("Warning: Could not get spreadsheet URL")
                    return True
            else:
                print("Error: Failed to upload data to Google Sheets")
                return False
                
        except Exception as e:
            print(f"Error uploading to Google Sheets: {str(e)}")
            import traceback
            traceback.print_exc()
            return False