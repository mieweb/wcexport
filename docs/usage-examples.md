# Usage Examples

This document provides practical examples for using each utility in the wcexport suite.

## 🖥️ GUI Applications

### wcexport.py - Full Chart Export

**Basic Usage:**
```bash
# Start with empty form
python wcexport.py

# Pre-fill URL and username
python wcexport.py "https://demo.webchart.com/webchart.cgi" "john.doe"
```

**Example Configuration:**
- **WebChart URL**: `https://demo.webchart.com/webchart.cgi`
- **Username**: `john.doe`
- **Password**: `[user enters securely]`
- **System Report**: `WebChart Export`
- **Print Definition**: `WebChart Export`
- **Output Directory**: `wcexport` (creates `~/wcexport/`)

**Example System Report (SQL):**
```sql
SELECT 
    p.pat_id,
    CONCAT(p.last_name, '_', p.first_name, '_', p.pat_id) as filename,
    CASE WHEN EXISTS(SELECT 1 FROM documents d WHERE d.pat_id = p.pat_id AND d.storage_type = 19) 
         THEN CONCAT('?f=export&type=cda&pat_id=', p.pat_id) 
         ELSE NULL END as cda_export,
    CASE WHEN EXISTS(SELECT 1 FROM documents d WHERE d.pat_id = p.pat_id AND d.storage_type = 21) 
         THEN CONCAT('?f=export&type=ccr&pat_id=', p.pat_id) 
         ELSE NULL END as ccr_export
FROM patients p 
WHERE p.status = 'Active' 
AND p.last_activity > DATE_SUB(NOW(), INTERVAL 1 YEAR)
ORDER BY p.last_name, p.first_name
LIMIT 100
```

**Output Structure:**
```
~/wcexport/
├── Smith_John_12345.pdf           # Main chart PDF
├── Smith_John_12345_cda_export    # CDA document (if available)
├── Jones_Mary_12346.pdf           # Main chart PDF  
├── Jones_Mary_12346_ccr_export    # CCR document (if available)
└── [timestamp].log                # Export log (if file logging enabled)
```

### docexport.py - Document Export

**Basic Usage:**
```bash
python docexport.py
```

**Example Configuration:**
- **WebChart URL**: `https://demo.webchart.com/webchart.cgi`
- **Username**: `jane.smith`
- **Password**: `[user enters securely]`
- **Begin Date**: `01/01/2024`
- **End Date**: `12/31/2024`
- **Document Types**: ☑️ CDA, ☑️ CCR
- **Schedule**: No Schedule (immediate export)

**Custom CGI Example:**
For advanced users, you can specify custom CGI parameters:
```
pat_id=12345&doc_type=custom&format=xml
```

**Output:**
Documents are saved with format: `{MR Number}_{Last}_{First}_{Doc ID}`

---

## 🔧 Command Line Utilities

### xmlimport.py - XML Document Import

**Basic Usage:**
```bash
python xmlimport.py "https://demo.webchart.com/webchart.cgi" "/path/to/xml/files" "mr_mapping.csv" "import.user"
```

**Directory Structure:**
```
/path/to/xml/files/
├── patient_001_cda.xml
├── patient_002_ccr.xml  
├── patient_003_cda.xml
└── mr_mapping.csv
```

**mr_mapping.csv Example:**
```csv
document_name,mr_number
patient_001_cda.xml,MR001234
patient_002_ccr.xml,MR001235
patient_003_cda.xml,MR001236
```

**Execution Example:**
```bash
$ python xmlimport.py "https://demo.webchart.com/webchart.cgi" "./import_docs" "mr_mapping.csv" "admin.user"
Please enter the webchart password for user [ admin.user ]: => [password entered securely]

Processing 3 XML files...
Success: patient_001_cda.xml => MR001234 => Document imported successfully
Success: patient_002_ccr.xml => MR001235 => Document imported successfully  
Failed import patient_003_cda.xml => MR001236 with error: Patient not found

Document import process complete:
Uploaded: 2
Skipped: 0
Errors: 1

See import.log for details
```

### wcdoc.py - Simple Document Download

**Configuration Required:**
Edit the constants at the top of `wcdoc.py`:

```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'your_username'
PASSWORD = 'your_password'
```

**Usage:**
```bash
python wcdoc.py
```

**Example Output:**
```
Initializing session
Getting Patients
  
Querying for patients: Last Name LIKE "Newman"
Getting Documents for Patient: 12345
Downloading Document: 67890
Success: Downloaded Newman_67890.xml

Querying for patients: Last Name LIKE "Larson"  
Getting Documents for Patient: 12346
No documents exist for that patient that meet the criteria.
```

### wcjson.py - JSON API Patient Data

**Configuration Required:**
Edit the constants in `wcjson.py`:

```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'api.user'
PASSWORD = 'api.password'
```

**Usage:**
```bash
python wcjson.py
```

**Example Output:**
```
Initializing session

Querying for patients: Last Name LIKE "Hart"
{
  "db": [
    {
      "pat_id": "12345",
      "last_name": "Hart", 
      "first_name": "John",
      "birth_date": "1985-03-15",
      "gender": "M"
    }
  ]
}

Querying for patients: Last Name LIKE "Pregnant"
{
  "db": [
    {
      "pat_id": "12346",
      "last_name": "Pregnant",
      "first_name": "Jane",
      "birth_date": "1990-07-22", 
      "gender": "F"
    }
  ]
}
```

### wcxml.py - Structured Document XML API

**Configuration Required:**
Edit the constants in `wcxml.py`:

```python
URL = 'https://your-webchart-system.com/webchart.cgi'
USERNAME = 'xml.user'
PASSWORD = 'xml.password'
SDATE = '2024-01-01'      # Start date
EDATE = '2024-12-31'      # End date  
NAMES = ['Smith', 'Jones', 'Brown']  # Patient last names to process
```

**Usage:**
```bash
python wcxml.py
```

**Example Output Structure:**
```
output/
├── Smith,John,A_12345/
│   ├── Patient Name.xml
│   ├── Sex.xml
│   ├── Date of Birth.xml
│   ├── Problems.xml
│   ├── Medications.xml
│   ├── Lab Values_Result.xml
│   └── Vital Signs.xml
├── Jones,Mary,B_12346/
│   ├── Patient Name.xml
│   ├── Sex.xml
│   ├── Problems.xml
│   └── Medications.xml
└── Brown,Robert,C_12347/
    ├── Patient Name.xml
    └── Sex.xml
```

---

## 🛠️ Advanced Usage Scenarios

### Scheduled Document Export

**Scenario**: Automatically export CDA documents every day at 2 AM.

**Configuration in docexport.py:**
1. Set date range for current day
2. Select CDA document type
3. Choose "Schedule At" → "Date"
4. Set time to 02:00
5. Click "Export"

### Bulk Chart Export with Custom Filenames

**System Report Example:**
```sql
SELECT 
    p.pat_id,
    CONCAT(
        YEAR(p.birth_date), '_',
        p.gender, '_',
        REPLACE(p.last_name, ' ', '_'), '_',
        REPLACE(p.first_name, ' ', '_')
    ) as filename
FROM patients p 
WHERE p.department = 'Cardiology'
AND p.last_visit >= '2024-01-01'
ORDER BY p.last_name
```

**Result**: Files named like `1985_M_Smith_John.pdf`

### Multi-Format Document Export

**System Report with Multiple Formats:**
```sql
SELECT 
    p.pat_id,
    CONCAT(p.last_name, '_', p.first_name) as filename,
    CONCAT('?f=chart&s=export&format=cda&pat_id=', p.pat_id) as cda_export,
    CONCAT('?f=chart&s=export&format=ccr&pat_id=', p.pat_id) as ccr_export,
    CONCAT('?f=chart&s=export&format=pdf&pat_id=', p.pat_id) as pdf_export
FROM patients p
WHERE p.status = 'Active'
```

### Error Recovery and Retry

**Manual Retry Process:**
1. Check export logs for failed items
2. Create new system report with only failed pat_ids
3. Re-run export with verbose logging enabled
4. Review detailed logs for specific failure reasons

### Performance Optimization

**Large Dataset Export:**
1. **Split by Date Ranges**: Create multiple system reports for different time periods
2. **Batch Processing**: Limit system reports to 50-100 patients at a time
3. **Off-Peak Hours**: Schedule large exports during low-usage periods
4. **Monitor Resources**: Watch CPU and memory usage during exports

---

## 🔍 Troubleshooting Examples

### Common Issues and Solutions

**Issue**: Authentication Failed
```bash
# Error in logs:
[2024-01-15 10:30:00] Login failed for https://demo.webchart.com/webchart.cgi

# Solution:
1. Verify URL includes 'webchart.cgi'
2. Test credentials in WebChart web interface  
3. Check user has 'Appliance Synchronization' permission
4. Enable verbose logging to see detailed error
```

**Issue**: System Report Not Found
```bash
# Error in logs:
[2024-01-15 10:30:00] System report [ Export Report ] does not contain the required "pat_id" column

# Solution:
1. Verify system report name matches exactly
2. Check system report contains 'pat_id' column
3. Test system report manually in WebChart
4. Ensure report returns CSV format
```

**Issue**: Print Job Failures
```bash
# Error in logs:
[2024-01-15 10:30:00] Failed to find job_url input or a pjob_id in the response for chart 12345

# Solution:
1. Verify print definition exists and is accessible
2. Check print definition generates PDF output
3. Test print definition manually in WebChart
4. Ensure user has printing permissions
```

**Issue**: Network Timeouts
```bash
# Error in logs:
[2024-01-15 10:30:00] Request timeout for chart 12345

# Solution:  
1. Increase retry count in code (modify retries parameter)
2. Check network connectivity to WebChart server
3. Run exports during off-peak hours
4. Contact WebChart administrator about server load
```

---

## 📊 Output Examples

### Successful Export Log
```
[2024-01-15 10:30:00] Application started.
[2024-01-15 10:30:15] Logging in
[2024-01-15 10:30:16] Running system report [ WebChart Export ]
[2024-01-15 10:30:17] Retrieved list of 25 charts
[2024-01-15 10:30:17] Starting export
[2024-01-15 10:30:18] Pat ID 12345, {'cda_export': '?f=export&type=cda&pat_id=12345'}
[2024-01-15 10:30:25] Pat ID 12346, {'ccr_export': '?f=export&type=ccr&pat_id=12346'}
[2024-01-15 10:30:32] Pat ID 12347, {}
[2024-01-15 10:30:45] Export completed successfully. 25 charts processed.
```

### Failed Export Log with Verbose
```
[2024-01-15 10:30:00] Application started.
[2024-01-15 10:30:15] Logging in
[2024-01-15 10:30:15] Attempt 1/3: Sending request to https://demo.webchart.com/webchart.cgi
[2024-01-15 10:30:16] Response code: 200
[2024-01-15 10:30:16] Request successful.
[2024-01-15 10:30:16] Running system report [ WebChart Export ]
[2024-01-15 10:30:17] Access denied to chart: 12345
[2024-01-15 10:30:18] Print for chart 12346 failed due to document errors, but what printed successfully was saved
[2024-01-15 10:30:25] Export completed with errors. 23 of 25 charts processed successfully.
```

This comprehensive set of examples should help users understand how to effectively use each utility in the wcexport suite.