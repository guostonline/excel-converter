#!/usr/bin/env python3
"""
Test script to verify Google Sheets permissions
"""

from google_sheets import GoogleSheetsService
import pandas as pd

def test_permissions():
    print("Testing Google Sheets permissions...")
    print("Service account email: n8n2-919@impactful-facet-218208.iam.gserviceaccount.com")
    print("Make sure this email has Editor access to your Google Sheet!")
    print()
    
    # Initialize service
    gs_service = GoogleSheetsService()
    
    # Authenticate
    if gs_service.authenticate_with_service_account("google.json"):
        print("✅ Authentication successful!")
    else:
        print("❌ Authentication failed!")
        return False
    
    # Try to open the spreadsheet
    spreadsheet_id = "1cRNqohML-mZ2mMqXIfQlPPDQUmRGLmRzDfp3j0wpte4"
    spreadsheet = gs_service.open_spreadsheet(spreadsheet_id)
    
    if spreadsheet:
        print("✅ Spreadsheet opened successfully!")
        print(f"Spreadsheet title: {spreadsheet.title}")
    else:
        print("❌ Failed to open spreadsheet!")
        return False
    
    # Create a simple test DataFrame
    test_data = pd.DataFrame({
        'Test Column 1': ['Test Value 1', 'Test Value 2'],
        'Test Column 2': ['Test Value 3', 'Test Value 4']
    })
    
    # Try to upload test data
    worksheet_name = "Suivi Test"
    success = gs_service.upload_dataframe_to_sheet(spreadsheet, worksheet_name, test_data)
    
    if success:
        print("✅ Test data uploaded successfully!")
        print("The Google Sheets integration is working correctly!")
        return True
    else:
        print("❌ Failed to upload test data!")
        print("Please make sure the service account has Editor permissions on the sheet.")
        return False

if __name__ == "__main__":
    test_permissions()