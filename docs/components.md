# Component Documentation

## GUI Applications

### wcexport.py - Main Chart Export Application

**Purpose**: Primary GUI application for comprehensive chart exports from WebChart systems.

**Key Features**:
- Full chart export with PDF generation
- System report-based chart selection
- Progress tracking and logging
- Command-line argument support for URL and username
- Threaded export operations

**Usage**:
```bash
# Run with GUI
python wcexport.py

# Pre-fill URL and username
python wcexport.py "https://example.com/webchart.cgi" "username"
```

**Required Inputs**:
- WebChart URL (must include 'webchart.cgi')
- Username and Password
- System Report name (default: "WebChart Export")
- Print Definition name (default: "WebChart Export")
- Output Directory (default: "wcexport")

**Class**: Uses `MainWin(win, fullExport=True)` from common.py

---

### docexport.py - Document Export Application

**Purpose**: Simplified GUI for exporting specific document types from WebChart.

**Key Features**:
- Date range-based document filtering
- Document type selection (CDA, CCR)
- Scheduled export capabilities
- Custom CGI parameter support

**Usage**:
```bash
python docexport.py
```

**Required Inputs**:
- WebChart URL
- Username and Password  
- Begin/End Date range
- Document Types (CDA, CCR, or custom CGI)

**Class**: Uses `MainWin(win, fullExport=False)` from common.py

---

## Command Line Utilities

### xmlimport.py - XML Document Import

**Purpose**: Command-line utility for importing XML documents into WebChart systems.

**Usage**:
```bash
python xmlimport.py <WebChartURL> <document_path> <mr_csv_file> <username>
```

**Parameters**:
- `WebChartURL`: Full URL to WebChart system
- `document_path`: Directory containing XML files to import
- `mr_csv_file`: CSV file mapping document names to MR numbers
- `username`: WebChart username

**Supported File Types**: `.xml` files

**Features**:
- Batch import processing
- CSV-based MR number mapping
- Import logging to `import.log`
- Error handling and success reporting

**CSV Format**:
```csv
document_name,mr_number
document1.xml,12345
document2.xml,67890
```

---

### wcdoc.py - Simple Document Download

**Purpose**: Basic command-line utility for downloading documents with hardcoded parameters.

**Configuration**: Edit the following constants in the file:
```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'your_username'
PASSWORD = 'your_password'
```

**Usage**:
```bash
python wcdoc.py
```

**Features**:
- Patient search by last name
- Document metadata retrieval
- Automatic document download
- Hardcoded search patterns for specific patients

---

### wcjson.py - JSON API Patient Data

**Purpose**: Retrieve patient data via WebChart's JSON API.

**Configuration**: Edit constants in the file:
```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'your_username'  
PASSWORD = 'your_password'
```

**Usage**:
```bash
python wcjson.py
```

**Features**:
- Patient search by last name patterns
- JSON API integration
- Automated session management
- Configurable search parameters

---

### wcxml.py - Structured Document XML API

**Purpose**: Retrieve structured clinical documents via WebChart's XML API.

**Configuration**: Edit constants in the file:
```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'your_username'
PASSWORD = 'your_password'
SDATE = 'start_date'  # Format: YYYY-MM-DD
EDATE = 'end_date'    # Format: YYYY-MM-DD
NAMES = ['PatientLastName1', 'PatientLastName2']
```

**Supported Data Types**:
- Patient Demographics (Name, Sex, DOB, Race, Ethnicity, Language)
- Clinical Data (Problems, Medications, Allergies, Lab Values, Vital Signs)
- Procedures and Immunizations
- Care Team and Medical Equipment
- Goals, Health Concerns, Assessments, Plans

**Usage**:
```bash
python wcxml.py
```

**Output**: Creates structured directories with XML files for each patient and data type.

---

## Core Module

### common.py - MainWin Class

**Purpose**: Core business logic and GUI framework for WebChart interactions.

#### Key Classes and Functions

**MainWin Class**:
```python
MainWin(win, fullExport=True)
```

**Core Methods**:

- `validateCredentials()`: Authenticates with WebChart and establishes session
- `getSystemReport()`: Executes system reports and parses CSV results
- `export()`: Main export orchestration with threading
- `getURLResponse(url, data, retries=3)`: HTTP request handling with retry logic
- `log(message, verbose=False)`: Thread-safe logging with timestamps

**Configuration Methods**:
- `validateInputs()`: Validates required fields before export
- `exportWrapper()`: Export workflow with scheduling support

**Utility Functions**:
- `sanitizeFilename(filename)`: Cleans filenames for filesystem compatibility
- `getSSLContext()`: Creates SSL context with disabled certificate verification  
- `getOutDir(cls, path)`: Resolves output directory paths

#### GUI Components

**Input Fields**:
- URL, Username, Password
- System Report and Print Definition (full export)
- Date ranges and document types (document export)
- Output directory and verbose logging

**Progress Tracking**:
- Progress bar with current item display
- Scrolled text log area
- Thread-safe queue-based updates

**Export Controls**:
- Export/Cancel button with state management
- Scheduling options for automated exports

#### Threading Architecture

**Main Features**:
- `ThreadPoolExecutor` for concurrent operations
- Thread-safe `queue.Queue` for GUI updates
- Graceful cancellation support
- Progress tracking across worker threads

**Session Management**:
- Automatic session establishment and maintenance
- Session refresh on authentication failures
- Cookie-based session persistence

#### Error Handling

**HTTP Error Handling**:
- Retry logic for failed requests
- Partial read recovery
- Session timeout recovery

**User Error Handling**:
- Input validation with user-friendly messages
- Export cancellation support
- Comprehensive error logging

#### WebChart Integration

**Authentication Flow**:
1. Credential validation via login request
2. Session cookie extraction and storage
3. Session ID attachment to subsequent requests
4. Automatic re-authentication on session expiry

**Export Operations**:
1. System report execution for chart list
2. Print job generation for each chart
3. PDF download with progress tracking
4. External URL processing (CDA, CCR, etc.)

**Document Operations**:
1. Date-based document search
2. Document metadata extraction
3. Bulk document download
4. Progress reporting