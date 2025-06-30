# wcexport Architecture Documentation

## System Architecture Overview

The wcexport suite consists of multiple Python utilities designed to interact with WebChart systems for data export, import, and retrieval operations.

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   wcexport.py   │    │  docexport.py   │    │   xmlimport.py  │
│                 │    │                 │    │                 │
│ Full Chart      │    │ Document-only   │    │ XML Document    │
│ Export GUI      │    │ Export GUI      │    │ Import CLI      │
└─────────┬───────┘    └─────────┬───────┘    └─────────┬───────┘
          │                      │                      │
          └──────────────────────┼──────────────────────┘
                                 │
                         ┌───────▼───────┐
                         │   common.py   │
                         │               │
                         │ MainWin Class │
                         │ Core Business │
                         │ Logic         │
                         └───────┬───────┘
                                 │
         ┌───────────────────────┼───────────────────────┐
         │                       │                       │
┌────────▼────────┐    ┌─────────▼─────────┐    ┌───────▼───────┐
│    wcdoc.py     │    │    wcjson.py      │    │   wcxml.py    │
│                 │    │                   │    │               │
│ Document        │    │ JSON API          │    │ XML API       │
│ Download CLI    │    │ Patient Data CLI  │    │ Structured    │
│                 │    │                   │    │ Document CLI  │
└─────────────────┘    └───────────────────┘    └───────────────┘
```

## Component Overview

### GUI Applications
- **wcexport.py**: Primary GUI application for comprehensive chart exports
- **docexport.py**: Simplified GUI for document-specific exports

### Command Line Utilities  
- **xmlimport.py**: Import XML documents into WebChart systems
- **wcdoc.py**: Download specific documents with hardcoded parameters
- **wcjson.py**: Query patient data via JSON API
- **wcxml.py**: Retrieve structured documents via XML API

### Core Module
- **common.py**: Contains the MainWin class with shared business logic

## Data Flow Architecture

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│    User     │────▶│   GUI/CLI   │────▶│  common.py  │
│  Interface  │     │Application  │     │   MainWin   │
└─────────────┘     └─────────────┘     └──────┬──────┘
                                               │
                                               ▼
┌──────────────────────────────────────────────────────────────┐
│                    WebChart System                           │
├──────────────────┬───────────────────┬──────────────────────┤
│  Authentication  │   System Reports  │    Document APIs     │
│     & Sessions   │   & Print Jobs    │   & File Downloads   │
└──────────────────┴───────────────────┴──────────────────────┘
                                               │
                                               ▼
┌──────────────────────────────────────────────────────────────┐
│                    Local File System                        │
├──────────────────┬───────────────────┬──────────────────────┤
│   PDF Charts     │   XML Documents   │   CSV Reports        │
│   (wcexport/)    │   (import logs)   │   (system reports)   │
└──────────────────┴───────────────────┴──────────────────────┘
```

## Authentication Flow

```
┌─────────────┐
│  User Input │
│ Credentials │
└──────┬──────┘
       │
       ▼
┌─────────────┐    Login Request     ┌─────────────┐
│ validateCre-│───────────────────▶  │  WebChart   │
│ dentials()  │                     │   System    │
└──────┬──────┘    Session Cookie   └─────────────┘
       │          ◀─────────────────
       ▼
┌─────────────┐
│ Store       │
│ session_id  │
│ for future  │
│ requests    │
└─────────────┘
```

## Export Process Flow

### Full Chart Export (wcexport.py)
```
User Input → Validate Credentials → Run System Report → 
For Each Chart:
  ├─ Generate Print Job
  ├─ Download PDF
  ├─ Download External URLs (CDA, CCR, etc.)
  └─ Update Progress
```

### Document Export (docexport.py)  
```
User Input → Validate Credentials → Search Documents by Date Range →
For Each Document:
  ├─ Download Document File
  └─ Update Progress
```

## Key Design Patterns

### Singleton Pattern
- Single MainWin instance manages all WebChart interactions
- Session management through single session_id

### Template Method Pattern
- Common validation and authentication flow
- Specialized export methods for different use cases

### Observer Pattern
- Progress updates via queue-based communication
- Thread-safe logging and UI updates

## Threading Architecture

```
┌─────────────────┐
│   Main Thread   │
│   (GUI/Event)   │
└─────────┬───────┘
          │
          ▼
┌─────────────────┐     ┌─────────────────┐
│ Export Thread   │────▶│   Thread Pool   │
│ (ThreadPool     │     │   (Workers)     │
│  Executor)      │     │                 │
└─────────┬───────┘     └─────────────────┘
          │
          ▼
┌─────────────────┐
│     Queue       │
│ (Thread-safe    │
│ Communication)  │
└─────────────────┘
```

## Security Considerations

- SSL certificate verification disabled for internal systems
- Session-based authentication with timeout handling
- Password fields masked in GUI
- Credential validation before operations
- No persistent credential storage

## Error Handling Strategy

- Retry mechanism for failed HTTP requests
- Graceful degradation for partial failures
- Comprehensive logging with timestamps
- User-friendly error messages
- Session re-establishment on auth failures