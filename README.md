# 📊 Excel Converter Application

A powerful Streamlit-based web application for processing and converting Excel files with automated calculations and Google Sheets integration.

## 🚀 Features

### 📈 Excel Processing
- **File Upload**: Support for `.xlsx` Excel files
- **Automated Calculations**: RAF TSM (Risk Adjusted Factor - Time Series Model) calculations
- **Multi-Sheet Processing**: Handles both AGADIR and QUALI NV sheets
- **Data Validation**: Automatic error handling and data validation
- **Custom Rest Days**: Configurable rest days for accurate calculations

### 🔗 Google Sheets Integration
- **Dual Upload Options**: 
  - AGADIR data → "Suivi Test" worksheet
  - QUALI NV data → "quali SOM VMM" worksheet
- **Real-time Sync**: Direct upload to Google Sheets
- **Secure Authentication**: Google Service Account integration

### 💾 Data Management
- **Download Processed Files**: Export processed Excel files
- **Session Persistence**: Maintains data across user interactions
- **Temporary File Handling**: Secure temporary file management

## 🛠️ Installation

### Prerequisites
- Python 3.7+
- Google Cloud Service Account (for Google Sheets integration)

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/guostonline/excel-converter.git
   cd excel-converter
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure Google Sheets (Optional)**
   - Create a Google Cloud Service Account
   - Download the credentials JSON file
   - Rename it to `google.json` and place in the project root
   - Share your Google Sheet with the service account email

4. **Run the application**
   ```bash
   streamlit run app.py
   ```

## 📋 Dependencies

```
streamlit>=1.28.0
openpyxl>=3.1.0
pandas
gspread
google-auth
google-auth-oauthlib
google-auth-httplib2
```

## 🎯 Usage

### Basic Workflow

1. **Upload Excel File**
   - Click "Choose an Excel file (.xlsx)"
   - Select your Excel file containing AGADIR and QUALI NV sheets

2. **Configure Settings**
   - Enter the number of rest days (auto-calculated from file data)
   - Review the extracted day work information

3. **Process File**
   - Click "Process Excel File"
   - Monitor the progress bar for processing status

4. **Review Results**
   - View processed AGADIR sheet data
   - Review QUALI NV sheet metrics
   - Check calculated RAF TSM values

5. **Export Data**
   - **Download**: Get the processed Excel file
   - **Google Sheets**: Upload to your Google Sheets

### Advanced Features

#### RAF TSM Calculation
The application automatically calculates the Risk Adjusted Factor using the formula:
```
RAF TSM = (value - value²) / rest_days
```

#### Multi-Sheet Processing
- **AGADIR Sheet**: Main data processing with automated calculations
- **QUALI NV Sheet**: Sales performance metrics and analysis

## 📁 Project Structure

```
excel-converter/
├── app.py                    # Main Streamlit application
├── excel.py                  # Excel processing logic
├── google_sheets.py          # Google Sheets integration
├── requirements.txt          # Python dependencies
├── days.json                 # Configuration for day calculations
├── excel/                    # Output directory for processed files
│   └── finale_jour.xlsx      # Processed Excel output
├── test_*.py                 # Test files
└── README.md                 # This file
```

## 🔧 Configuration

### Days Configuration (`days.json`)
```json
{
  "from_file": {
    "t": "24",  // Total days
    "d": "4"    // Work days
  }
}
```

### Google Sheets Setup
1. Create a Google Cloud Project
2. Enable Google Sheets API
3. Create a Service Account
4. Add credentials to Streamlit secrets as `[google_service_account]` (Streamlit Cloud: App Settings → Secrets)
5. Share your target Google Sheet with the service account email

## 🚨 Security Features

- **Sensitive File Protection**: `.gitignore` prevents credential files from being committed
- **Temporary File Cleanup**: Automatic cleanup of uploaded files
- **Secure Authentication**: Google Service Account for API access

## 🐛 Troubleshooting

### Common Issues

1. **Google Sheets Upload Failed**
   - Verify Streamlit secrets contain a `[google_service_account]` section with full JSON fields (e.g., `type`, `project_id`, `private_key`, `client_email`, `token_uri`)
   - In Streamlit Cloud, set secrets via App Settings → Secrets (local `secrets.toml` is not deployed automatically)
   - Check if the Google Sheet is shared with the service account email
   - Ensure Google Sheets API is enabled

2. **Excel Processing Errors**
   - Verify Excel file contains AGADIR and QUALI NV sheets
   - Check file format is `.xlsx` (not `.xls`)
   - Ensure file is not corrupted

3. **Installation Issues**
   - Update pip: `pip install --upgrade pip`
   - Use virtual environment: `python -m venv venv`
   - Install dependencies individually if batch install fails

## 📊 Data Flow

```
Excel Upload → File Validation → Sheet Processing → RAF Calculations → 
Data Display → Google Sheets Upload → File Download
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

For support and questions:
- Create an issue on GitHub
- Check the troubleshooting section above
- Review the application logs in the Streamlit interface

## 🔄 Version History

- **v1.0.0**: Initial release with basic Excel processing
- **v1.1.0**: Added Google Sheets integration
- **v1.2.0**: Enhanced RAF TSM calculations and multi-sheet support
- **v1.3.0**: Security improvements and credential protection

---

**Made with ❤️ using Streamlit and Python**
