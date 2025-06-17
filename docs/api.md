# API Documentation

## MainWin Class API Reference

The `MainWin` class in `common.py` provides the core API for WebChart interactions.

### Constructor

```python
MainWin(win, fullExport=True)
```

**Parameters**:
- `win`: Tkinter window object for GUI attachment
- `fullExport`: Boolean flag for full chart export mode vs document export mode

**Returns**: MainWin instance

---

### Authentication Methods

#### validateCredentials()

```python
def validateCredentials(self) -> bool
```

**Purpose**: Authenticates with WebChart system and establishes session.

**Returns**: 
- `True` if authentication successful
- `False` if authentication failed

**Side Effects**:
- Sets `self.session_id` on successful authentication
- Displays error messages on failure

**Usage**:
```python
if app.validateCredentials():
    print("Authentication successful")
else:
    print("Authentication failed")
```

---

### Data Retrieval Methods

#### getSystemReport()

```python
def getSystemReport(self) -> bool
```

**Purpose**: Executes the configured system report and parses results.

**Requirements**:
- Valid authentication session
- System report name in `self.report`

**Returns**:
- `True` if report executed successfully
- `False` if report failed

**Side Effects**:
- Populates `self.charts` list with chart data
- Each chart contains: `pat_id`, `filename`, `urls`

**Expected CSV Format**:
```csv
pat_id,filename,additional_url1,additional_url2
123,Patient_Smith,?doc_id=456,?cda_export=789
124,Patient_Jones,?doc_id=457,
```

---

#### getURLResponse()

```python
def getURLResponse(self, url, data=None, retries=3) -> tuple
```

**Purpose**: Makes HTTP requests to WebChart with automatic retry and session management.

**Parameters**:
- `url`: Target URL string
- `data`: Dictionary of POST parameters (optional)
- `retries`: Number of retry attempts on failure (default: 3)

**Returns**: 
- Tuple of `(response_bytes, response_object)`

**Features**:
- Automatic session ID injection
- UTF-8 BOM handling
- Retry logic with re-authentication
- SSL context management

**Usage**:
```python
response_data, response_obj = app.getURLResponse(
    "https://example.com/webchart.cgi",
    {"f": "admin", "s": "system_report"}
)
```

---

### Export Methods

#### export()

```python
def export(self) -> None
```

**Purpose**: Main export orchestration method.

**Behavior**:
- **Full Export Mode**: Exports charts based on system report
- **Document Export Mode**: Exports documents based on date range and type

**Features**:
- Multi-threaded processing with ThreadPoolExecutor
- Progress tracking and user feedback
- Graceful cancellation support
- Error handling and logging

**Threading**:
- Uses `concurrent.futures.ThreadPoolExecutor` with max 2 workers
- Thread-safe communication via `queue.Queue`
- Progress updates from worker threads

---

#### exportWrapper()

```python
def exportWrapper(self) -> None
```

**Purpose**: Export workflow wrapper with validation and scheduling.

**Features**:
- Input validation before export
- Scheduled export support
- User confirmation for scheduled operations
- Export state management

---

### Utility Methods

#### log()

```python
def log(self, message, verbose=False) -> None
```

**Purpose**: Thread-safe logging with timestamp and verbosity control.

**Parameters**:
- `message`: Log message string
- `verbose`: Optional verbose flag (default: False)

**Features**:
- Automatic timestamp prefixing
- GUI text area updates
- Optional file logging
- Verbose message filtering

**Usage**:
```python
app.log("Export started")
app.log("Debug information", verbose=True)
```

---

### Input Validation Methods

#### validateInputs()

```python
def validateInputs(self) -> bool
```

**Purpose**: Validates required input fields before export operations.

**Full Export Mode Requirements**:
- URL, Username, Password
- System Report name
- Print Definition name

**Document Export Mode Requirements**:
- URL, Username, Password
- At least one document type selected OR custom CGI parameters

**Returns**: 
- `True` if all required inputs are valid
- `False` if validation fails (with user notification)

---

## Utility Functions API

### sanitizeFilename()

```python
def sanitizeFilename(filename) -> str
```

**Purpose**: Sanitizes filenames for filesystem compatibility.

**Parameters**:
- `filename`: Original filename string

**Returns**: Sanitized filename with invalid characters replaced by underscores

**Valid Characters**: Letters, digits, spaces, and `_-,.()[]`

**Usage**:
```python
safe_name = sanitizeFilename("Patient: John Doe (ID#123)")
# Result: "Patient_ John Doe _ID_123_"
```

---

### getSSLContext()

```python
def getSSLContext() -> ssl.SSLContext
```

**Purpose**: Creates SSL context with disabled certificate verification.

**Returns**: SSL context configured for internal/self-signed certificates

**Usage**: Used automatically by `getURLResponse()` method.

---

### getOutDir()

```python
def getOutDir(cls, path) -> str
```

**Purpose**: Resolves output directory path handling both absolute and relative paths.

**Parameters**:
- `cls`: Class reference (unused)
- `path`: Directory path string

**Returns**: 
- Absolute path if input is absolute
- Path relative to user home directory if input is relative

**Usage**:
```python
output_dir = getOutDir(None, "wcexport")
# Result: "/home/user/wcexport"

output_dir = getOutDir(None, "/tmp/exports")  
# Result: "/tmp/exports"
```

---

## WebChart Integration API

### System Report Format

**Required Columns**:
- `pat_id`: Patient ID for chart identification (required)
- `filename`: Desired output filename (optional, defaults to pat_id)

**Optional Columns**:
Any additional columns with values starting with `?` are treated as relative URLs for additional file downloads.

**Example**:
```csv
pat_id,filename,cda_export,ccr_export
12345,Smith_John,?f=export&type=cda&pat_id=12345,?f=export&type=ccr&pat_id=12345
12346,Jones_Mary,?f=export&type=cda&pat_id=12346,
```

### Print Definition Requirements

Print definitions must be configured in WebChart with:
- Name matching the input field (default: "WebChart Export")
- Appropriate permissions for the user account
- PDF output format

### Document Search Parameters

**Date Range Format**: 
- Separate month, day, year fields
- Time components: "00:00" for start, "23:59" for end

**Document Types**:
- `cda`: Clinical Document Architecture (storage_type=19)
- `ccr`: Continuity of Care Record (storage_type=21)

### Session Management

**Authentication Flow**:
1. POST to WebChart URL with `login_user` and `login_passwd`
2. Extract session cookie from `Set-Cookie` header
3. Include `session_id` in all subsequent requests
4. Handle session expiry with automatic re-authentication

**Session Parameters**:
- Cookie format: `PHPSESSID=<session_id>`
- Automatic injection into request data
- Session validation via `X-lg_status` header