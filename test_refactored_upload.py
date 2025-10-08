#!/usr/bin/env python3
"""
Test script to verify the refactored Google Sheets upload functionality
"""

from google_sheets import GoogleSheetsService

def test_refactored_upload():
    """Test the new upload_excel_to_google_sheets method"""
    print("🧪 Testing refactored Google Sheets upload...")
    
    # Initialize Google Sheets service
    gs_service = GoogleSheetsService()
    
    # Test the upload
    result = gs_service.upload_excel_to_google_sheets()
    
    if result:
        print("✅ Refactored upload test PASSED!")
        print(f"📊 Result: {result}")
    else:
        print("❌ Refactored upload test FAILED!")
    
    return result

if __name__ == "__main__":
    test_refactored_upload()