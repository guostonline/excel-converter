import streamlit as st
import os
import tempfile
import shutil
from excel import Excel
import json
import pandas as pd
from google_sheets import GoogleSheetsService

# Create necessary directories
os.makedirs("excel", exist_ok=True)

def create_days_json():
    """Create days.json file if it doesn't exist"""
    if not os.path.exists("days.json"):
        default_data = {
            "from_file": {"t": "24", "d": "4"}
        }
        with open("days.json", "w") as f:
            json.dump(default_data, f)

def main():
    st.set_page_config(
        page_title="Excel Converter",
        page_icon="📊",
        layout="centered"
    )
    
    st.title("📊 Excel File Converter")
    
    
    # Create days.json if it doesn't exist
    create_days_json()
    
    # Direct Excel converter functionality
    excel_converter_section()

def excel_converter_section():
    """Excel converter functionality"""
    
    # Initialize session state for processed data
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'output_path' not in st.session_state:
        st.session_state.output_path = None
    if 'excel_processor' not in st.session_state:
        st.session_state.excel_processor = None
    
    # File upload section
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx)",
        type=['xlsx'],
        help="Upload an Excel file to process and convert"
    )
    
    if uploaded_file is not None:
        # Display file info
        st.success(f"File uploaded: {uploaded_file.name}")
        
        
        
        
        # Initialize total_day_work with default value from days.json or fallback
        try:
            with open('days.json', 'r') as f:
                days_data = json.load(f)
                d_value = days_data.get('from_file', {}).get('d', 4)
                total_day_work = int(d_value) if isinstance(d_value, (str, int, float)) else 4
        except (FileNotFoundError, KeyError, ValueError, TypeError):
            total_day_work = 4  # fallback default
            
        # Calculate default rest days
        day_rest = 24 - total_day_work
        
        # User input for rest days (used in RAF calculation)
        jour_rest = int(st.text_input("Enter the number of rest days:", value=str(day_rest)))
        
        # Debug: Show the jour_rest value being used
       
        
        # Process button
        if st.button("Process Excel File", type="primary"):
            temp_path = None
            try:
                # Create a temporary file to save the uploaded file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    temp_path = tmp_file.name
                
                # Initialize Excel processor
                excel_processor = Excel(temp_path, rest_days=jour_rest)
                
                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Extract day work information
                status_text.text("Extracting day work information...")
                progress_bar.progress(25)
                
                try:
                    total_days, work_days = excel_processor.get_day_work()
                    st.success(f"✅ File processed successfully!. Day work extracted: {total_days} Total Days, {work_days} Work Days")
                except Exception as e:
                    st.warning(f"Could not extract day work information: {str(e)}")
                
                # Fix and process the sheet
                status_text.text("Processing Excel sheets...")
                progress_bar.progress(50)
                
                # Process with jour_rest parameter
                success = excel_processor.fix_sheet(jour_rest=jour_rest)
                progress_bar.progress(75)
                
                if success:
                    status_text.text("Processing completed successfully!")
                    progress_bar.progress(100)
                    
                    # Determine output filename
                    output_filename = "finale_jour.xlsx"
                    output_path = f"excel/{output_filename}"
                    
                    if os.path.exists(output_path):
                        
                        # Read the processed file into a dataframe for display
                        df = pd.read_excel(output_path)
                        
                        # Store processed data in session state
                        st.session_state.processed_data = df
                        st.session_state.output_path = output_path
                        st.session_state.excel_processor = excel_processor
                        
                        # Get QUALI NV data
                        df_quali = excel_processor.get_quali_nv_dataframe()
                        if df_quali is not None:
                            st.session_state.quali_nv_data = df_quali
                        
                       
                       
                        
                        # Display the processed dataframe
                        st.subheader("📊 Processed Data - AGADIR Sheet")
                        st.dataframe(df)
                        
                        # Display QUALI NV data if available
                        if df_quali is not None:
                            st.subheader("📈 QUALI NV Sheet Data")
                            st.dataframe(df_quali)
                            st.info(f"QUALI NV sheet contains {len(df_quali)} rows with sales performance metrics")
                        else:
                            st.warning("⚠️ QUALI NV sheet data could not be loaded")
                        
                        # Add button to send data to Google Sheets
                        st.subheader("Send to Google Sheets")
                        if st.button("📤 Send Data to Google Sheets", type="primary"):
                            with st.spinner("Uploading to Google Sheets..."):
                                try:
                                    # Initialize Google Sheets service and upload
                                    gs_service = GoogleSheetsService()
                                    upload_result = gs_service.upload_excel_to_google_sheets()
                                    if upload_result:
                                        st.success("✅ Data uploaded to Google Sheets successfully!")
                                        st.info("📊 Data has been sent to your Google Sheet: 'Suivi Test' worksheet")
                                    else:
                                        st.error("❌ Failed to upload to Google Sheets. Check your google.json credentials.")
                                except Exception as e:
                                    st.error(f"❌ Google Sheets upload failed: {str(e)}")
                        
                        # Provide download button
                        st.subheader("Download File")
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="📥 Download Processed Excel File",
                                data=f,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("Output file was not created. Please check the logs.")
                else:
                    st.error("Failed to process the Excel file.")
                    
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
            finally:
                # Clean up temporary file
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
    
    # Display processed data if it exists in session state
    if st.session_state.processed_data is not None:
        st.subheader("📊 Processed Data - AGADIR Sheet")
        st.dataframe(st.session_state.processed_data)
        
        # Display QUALI NV data if available in session state
        if hasattr(st.session_state, 'quali_nv_data') and st.session_state.quali_nv_data is not None:
            st.subheader("📈 QUALI NV Sheet Data")
            st.dataframe(st.session_state.quali_nv_data)
            st.info(f"QUALI NV sheet contains {len(st.session_state.quali_nv_data)} rows with sales performance metrics")
        
        # Add button to send data to Google Sheets
        st.subheader("Send to Google Sheets")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📤 Send AGADIR Data to Google Sheets", type="primary", key="upload_agadir"):
                with st.spinner("Uploading AGADIR data to Google Sheets..."):
                    try:
                        # Initialize Google Sheets service and upload AGADIR data
                        gs_service = GoogleSheetsService()
                        upload_result = gs_service.upload_excel_to_google_sheets()
                        if upload_result:
                            st.success("✅ AGADIR data uploaded to Google Sheets successfully!")
                            st.info("📊 AGADIR data has been sent to your Google Sheet: 'Suivi Test' worksheet")
                        else:
                            st.error("❌ Failed to upload AGADIR data to Google Sheets. Check your google.json credentials.")
                    except Exception as e:
                        st.error(f"❌ AGADIR Google Sheets upload failed: {str(e)}")
        
        with col2:
            if hasattr(st.session_state, 'quali_nv_data') and st.session_state.quali_nv_data is not None:
                if st.button("📈 Send QUALI NV Data to Google Sheets", type="secondary", key="upload_quali"):
                    with st.spinner("Uploading QUALI NV data to Google Sheets..."):
                        try:
                            # Initialize Google Sheets service and upload QUALI NV data
                            gs_service = GoogleSheetsService()
                            upload_result = gs_service.upload_quali_nv_to_google_sheets()
                            if upload_result:
                                st.success("✅ QUALI NV data uploaded to Google Sheets successfully!")
                                st.info("📊 QUALI NV data has been sent to your Google Sheet: 'QUALI NV' worksheet")
                            else:
                                st.error("❌ Failed to upload QUALI NV data to Google Sheets. Check your google.json credentials.")
                        except Exception as e:
                            st.error(f"❌ QUALI NV Google Sheets upload failed: {str(e)}")
            else:
                st.info("ℹ️ QUALI NV data not available for upload")
        
        # Provide download button
        if st.session_state.output_path and os.path.exists(st.session_state.output_path):
            st.subheader("Download File")
            with open(st.session_state.output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Processed Excel File",
                    data=f,
                    file_name="finale_jour.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="persistent_download"
                )

if __name__ == "__main__":
    main()
