#!/usr/bin/env python3
"""
Test script demonstrating direct Google Sheets integration
"""

from excel import Excel
import os

def test_excel_to_google_sheets():
    """
    Example of how to process Excel file and upload to Google Sheets programmatically
    """
    
    # Path to your Excel file
    excel_file_path = "path/to/your/excel_file.xlsx"  # Replace with actual path
    
    # Check if file exists (for demo purposes, we'll skip this)
    if not os.path.exists(excel_file_path):
        print("⚠️  Excel file not found. This is a demo script.")
        print("📝 To use this script:")
        print("   1. Replace 'excel_file_path' with your actual Excel file path")
        print("   2. Make sure google.json is in the project directory")
        print("   3. Run this script")
        return
    
    try:
        # Initialize Excel processor with rest days
        rest_days = 20  # Adjust as needed
        excel_processor = Excel(excel_file_path, rest_days=rest_days)
        
        print("📊 Processing Excel file...")
        
        # Extract day work information
        try:
            total_days, work_days = excel_processor.get_day_work()
            print(f"✅ Day work extracted: {total_days} Total Days, {work_days} Work Days")
        except Exception as e:
            print(f"⚠️  Could not extract day work: {e}")
        
        # Process the Excel file
        print("🔄 Processing Excel sheets...")
        success = excel_processor.fix_sheet(jour_rest=rest_days)
        
        if success:
            print("✅ Excel processing completed!")
            
            # Upload to Google Sheets automatically
            print("📤 Uploading to Google Sheets...")
            
            # Option 1: Simple upload with default settings
            result = excel_processor.upload_to_google_sheets()
            
            # Option 2: Custom upload with specific settings
            # result = excel_processor.upload_to_google_sheets(
            #     spreadsheet_name="My Custom Spreadsheet",
            #     worksheet_name="Data_Sheet",
            #     share_email="colleague@example.com"
            # )
            
            if result:
                if isinstance(result, str):
                    print(f"🎉 Success! Spreadsheet URL: {result}")
                else:
                    print("🎉 Upload completed successfully!")
            else:
                print("❌ Failed to upload to Google Sheets")
        else:
            print("❌ Excel processing failed")
            
    except Exception as e:
        print(f"❌ Error: {e}")

def demo_usage_examples():
    """
    Show different ways to use the Google Sheets integration
    """
    print("\n" + "="*60)
    print("📚 USAGE EXAMPLES")
    print("="*60)
    
    print("""
# Example 1: Basic usage
excel_processor = Excel("my_file.xlsx", rest_days=20)
excel_processor.fix_sheet(jour_rest=20)
excel_processor.upload_to_google_sheets()

# Example 2: Custom spreadsheet name
excel_processor.upload_to_google_sheets(
    spreadsheet_name="Monthly Report 2024"
)

# Example 3: Custom worksheet and sharing
excel_processor.upload_to_google_sheets(
    spreadsheet_name="Team Data",
    worksheet_name="January_Data", 
    share_email="manager@company.com"
)

# Example 4: Get the URL for further processing
url = excel_processor.upload_to_google_sheets(
    spreadsheet_name="Analysis Results"
)
if url:
    print(f"Share this link: {url}")
""")

if __name__ == "__main__":
    print("🚀 Excel to Google Sheets Integration Test")
    print("="*50)
    
    # Check if google.json exists
    if os.path.exists("google.json"):
        print("✅ google.json found")
        test_excel_to_google_sheets()
    else:
        print("❌ google.json not found")
        print("📝 Please add your Google Service Account credentials as 'google.json'")
    
    demo_usage_examples()