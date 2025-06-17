# wcexport - WebChart Export Utility Suite

A comprehensive suite of Python utilities for exporting, importing, and managing data from WebChart Electronic Health Record systems. This toolkit provides both GUI and command-line interfaces for various WebChart integration tasks.

## 🚀 Quick Start

### Windows Users
Download the pre-built executables from [Releases](https://github.com/mieweb/wcexport/releases):
- `wcexport.exe` - Full chart export application
- `docexport.exe` - Document-specific export application

### Python Users
```bash
# Clone the repository
git clone https://github.com/mieweb/wcexport.git
cd wcexport

# Install dependencies
pip install -r requirements.txt

# Run the main application
python wcexport.py

# Or run document export
python docexport.py
```

## 📋 Components Overview

| Component | Type | Purpose |
|-----------|------|---------|
| **wcexport.py** | GUI | Full chart export with PDF generation |
| **docexport.py** | GUI | Document-specific export with date filtering |
| **xmlimport.py** | CLI | Batch import XML documents to WebChart |
| **wcdoc.py** | CLI | Simple document download utility |
| **wcjson.py** | CLI | Patient data retrieval via JSON API |
| **wcxml.py** | CLI | Structured document retrieval via XML API |
| **common.py** | Module | Core business logic and GUI framework |

## 🏗️ Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   wcexport.py   │    │  docexport.py   │    │   xmlimport.py  │
│ Full Chart GUI  │    │ Document GUI    │    │ XML Import CLI  │
└─────────┬───────┘    └─────────┬───────┘    └─────────┬───────┘
          │                      │                      │
          └──────────────────────┼──────────────────────┘
                                 │
                         ┌───────▼───────┐
                         │   common.py   │
                         │ Core Business │
                         │    Logic      │
                         └───────┬───────┘
                                 │
                         ┌───────▼───────┐
                         │  WebChart     │
                         │    System     │
                         └───────────────┘
```

[📖 View Full Architecture Documentation](docs/architecture.md)

## 🔧 Installation & Setup

### Requirements
- Python 3.8+ (tested with 3.13.2)
- WebChart system access with appropriate permissions
- 'Appliance Synchronization' permission for chart exports

### System Dependencies
```bash
# Install Python dependencies
pip install -r requirements.txt
```

**Required Python Packages:**
- `Requests==2.32.3` - HTTP client library
- `tk==0.1.0` - Tkinter GUI framework

### WebChart Prerequisites
1. **Valid WebChart URL** - Must include the 'webchart.cgi' path
2. **User Account** with appropriate permissions:
   - 'Appliance Synchronization' permission for exports
   - Access to system reports and print definitions
3. **System Report** (for full chart export) - See [System Report Configuration](#system-report-configuration)
4. **Print Definition** (for full chart export) - Configured in WebChart admin

## 📖 Usage Guide

### Full Chart Export (wcexport.py)

Export complete patient charts as PDF files using WebChart's print functionality.

```bash
# GUI Application
python wcexport.py

# Pre-fill URL and username
python wcexport.py "https://your-server.com/webchart.cgi" "username"
```

**Required Inputs:**
- WebChart URL
- Username & Password
- System Report name (default: "WebChart Export")
- Print Definition name (default: "WebChart Export")
- Output Directory (default: "wcexport")

**Features:**
- Multi-threaded export processing
- Progress tracking with real-time updates
- Automatic external file download (CDA, CCR, etc.)
- Comprehensive logging
- Export cancellation support

### Document Export (docexport.py)

Export specific document types with date range filtering.

```bash
python docexport.py
```

**Required Inputs:**
- WebChart URL & Credentials
- Date range (begin/end dates)
- Document types (CDA, CCR, or custom CGI)

**Features:**
- Date-based document filtering
- Scheduled export capabilities
- Document type selection
- Custom CGI parameter support

### XML Document Import (xmlimport.py)

Batch import XML documents into WebChart systems.

```bash
python xmlimport.py <WebChartURL> <document_path> <mr_csv_file> <username>
```

**Parameters:**
- `WebChartURL`: Full WebChart system URL
- `document_path`: Directory containing XML files
- `mr_csv_file`: CSV mapping document names to MR numbers
- `username`: WebChart username

**CSV Format:**
```csv
document_name,mr_number
patient1_cda.xml,12345
patient2_ccr.xml,67890
```

### Command Line Utilities

**Simple Document Download:**
```bash
# Edit wcdoc.py configuration first
python wcdoc.py
```

**JSON API Patient Data:**
```bash
# Edit wcjson.py configuration first  
python wcjson.py
```

**Structured XML Document Retrieval:**
```bash
# Edit wcxml.py configuration first
python wcxml.py
```

## ⚙️ Configuration

### System Report Configuration

The system report must contain specific columns for proper chart identification:

**Required Columns:**
- `pat_id` - Patient/chart internal identifier (required)
- `filename` - Desired output filename (optional, defaults to pat_id)

**Optional Columns:**
Any additional columns with values starting with `?` are treated as relative URLs for downloading additional files (CDA, CCR, etc.).

**Example System Report SQL:**
```sql
SELECT 
    p.pat_id,
    CONCAT(p.last_name, '_', p.first_name, '_', p.pat_id) as filename,
    CASE WHEN EXISTS(SELECT 1 FROM documents d WHERE d.pat_id = p.pat_id AND d.storage_type = 19) 
         THEN CONCAT('?f=export&type=cda&pat_id=', p.pat_id) 
         ELSE NULL END as cda_export
FROM patients p 
WHERE p.status = 'Active'
```

### Print Definition Setup

1. Access WebChart Admin → System Administration → Manage Print Definitions
2. Create or configure a print definition named "WebChart Export" (or your preferred name)
3. Ensure the definition generates PDF output
4. Grant appropriate user permissions

### Output Directory Structure

```
~/wcexport/                    # Default output directory
├── patient1_chart.pdf         # Main chart PDFs
├── patient2_chart.pdf
├── patient1_chart_cda.xml     # External files (CDA)
├── patient2_chart_ccr.xml     # External files (CCR)
└── export.log                 # Export logs (when file logging enabled)
```

## 🔍 Troubleshooting

### Common Issues

**Authentication Failures:**
- Verify WebChart URL includes 'webchart.cgi'
- Check username/password credentials
- Ensure user has 'Appliance Synchronization' permission

**System Report Errors:**
- Verify system report exists and is accessible
- Check that report contains required 'pat_id' column
- Ensure report returns valid CSV data

**Export Failures:**
- Check print definition exists and is accessible
- Verify network connectivity to WebChart server
- Check output directory permissions

**SSL Certificate Issues:**
- The application disables SSL verification for self-signed certificates
- For additional security, configure proper SSL certificates on WebChart server

### Verbose Logging

Enable verbose logging in the GUI applications for detailed troubleshooting:
1. Check "Verbose Logging" option in the interface
2. Review the log area for detailed request/response information
3. Check log files in the output directory when file logging is enabled

### Performance Optimization

**For Large Exports:**
- Use smaller batch sizes by splitting system reports
- Monitor system resources during multi-threaded operations
- Consider running exports during off-peak hours

**Network Optimization:**
- Ensure stable network connection to WebChart server
- Consider running from same network as WebChart server
- Monitor for network timeouts in verbose logs

## 📚 Documentation

- [📖 Architecture Documentation](docs/architecture.md) - System design and data flow
- [🔧 Component Documentation](docs/components.md) - Detailed component reference
- [⚡ API Documentation](docs/api.md) - Developer API reference
- [🤝 Contributing Guidelines](CONTRIBUTING.md) - Development and contribution guide

## 🛠️ Development

### Building from Source

```bash
# Clone repository
git clone https://github.com/mieweb/wcexport.git
cd wcexport

# Install development dependencies
pip install -r requirements.txt
pip install pyinstaller

# Build executables
pyinstaller --onefile wcexport.py
pyinstaller --onefile docexport.py

# Executables will be in dist/ directory
```

### Project Structure

```
wcexport/
├── wcexport.py          # Main GUI application
├── docexport.py         # Document export GUI
├── common.py            # Core business logic
├── xmlimport.py         # XML import utility
├── wcdoc.py             # Simple document download
├── wcjson.py            # JSON API utility  
├── wcxml.py             # XML API utility
├── requirements.txt     # Python dependencies
├── Jenkinsfile         # CI/CD pipeline
├── docs/               # Documentation
│   ├── architecture.md
│   ├── components.md
│   └── api.md
└── .github/workflows/  # GitHub Actions
    └── windows-build.yml
```

## 📄 License

This project is open source. Please refer to the repository license for specific terms.

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details on:
- Setting up the development environment
- Code standards and style guidelines
- Testing procedures
- Submitting pull requests

## 💬 Support

For support and questions:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review [Documentation](docs/) for detailed information
3. Submit issues via GitHub Issues
4. Contact the development team through appropriate channels

---

**WebChart Export Utility Suite** - Streamlining healthcare data integration and export workflows.

