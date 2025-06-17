# Contributing to wcexport

We welcome contributions to the wcexport project! This document provides guidelines for contributing to the codebase.

## 🚀 Getting Started

### Development Environment Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/mieweb/wcexport.git
   cd wcexport
   ```

2. **Set up Python environment:**
   ```bash
   # Recommended: Create virtual environment
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   
   # Install dependencies
   pip install -r requirements.txt
   ```

3. **Install development tools:**
   ```bash
   pip install pyinstaller  # For building executables
   ```

### Project Structure Understanding

Before contributing, familiarize yourself with the project structure:

```
wcexport/
├── wcexport.py          # Main GUI - full chart export
├── docexport.py         # GUI - document export  
├── common.py            # Core business logic (MainWin class)
├── xmlimport.py         # CLI - XML document import
├── wcdoc.py             # CLI - simple document download
├── wcjson.py            # CLI - JSON API utility
├── wcxml.py             # CLI - XML API utility
├── requirements.txt     # Python dependencies
├── docs/                # Documentation
└── .github/workflows/   # CI/CD configuration
```

## 🔧 Development Guidelines

### Code Style

**Python Style:**
- Follow PEP 8 style guidelines
- Use meaningful variable and function names
- Add docstrings for public methods and classes
- Keep functions focused and reasonably sized

**Example:**
```python
def sanitizeFilename(filename):
    """
    Sanitizes a filename by replacing invalid characters.
    
    Args:
        filename (str): The original filename
        
    Returns:
        str: Sanitized filename safe for filesystem use
    """
    return ''.join([x if x in validChars else '_' for x in filename])
```

### GUI Development

**Tkinter Guidelines:**
- Use consistent widget naming conventions
- Implement proper event handling
- Ensure thread-safe GUI updates via queue mechanism
- Follow existing layout patterns

**Threading:**
- Use `ThreadPoolExecutor` for concurrent operations
- Communicate with GUI thread via `queue.Queue`
- Handle cancellation gracefully
- Provide progress feedback

### Error Handling

**Best Practices:**
- Use specific exception types when possible
- Provide meaningful error messages to users
- Log detailed error information for debugging
- Implement retry logic for network operations

**Example:**
```python
try:
    response = self.getURLResponse(url, data)
except Exception as e:
    self.log(f"Failed to connect to WebChart: {e}")
    tkm.showerror(message=f'Connection failed: {e}')
    return False
```

### WebChart Integration

**API Interaction:**
- Always include session management
- Handle authentication failures gracefully
- Implement appropriate retry mechanisms
- Respect rate limits and system load

**Data Handling:**
- Validate CSV data structure before processing
- Handle UTF-8 BOM in WebChart responses
- Sanitize filenames for cross-platform compatibility
- Process large datasets efficiently

## 🧪 Testing

### Manual Testing

Before submitting changes, test the following scenarios:

**GUI Applications:**
1. **Authentication Testing:**
   - Valid credentials
   - Invalid credentials  
   - Network connectivity issues
   - Session timeout scenarios

2. **Export Testing:**
   - Small dataset (1-5 charts)
   - Medium dataset (10-50 charts)
   - Error handling (missing files, permission issues)
   - Progress tracking accuracy
   - Export cancellation

3. **Configuration Testing:**
   - Various output directory configurations
   - Different system report formats
   - Custom print definitions
   - Verbose logging functionality

**CLI Utilities:**
1. Test with sample data files
2. Verify error handling for missing files
3. Check output format and logging
4. Test with various WebChart configurations

### Automated Testing

Currently, the project lacks automated tests. Contributions to add testing infrastructure are welcome:

**Testing Framework Suggestions:**
- `pytest` for unit tests
- `unittest.mock` for WebChart API mocking
- GUI testing with `tkinter.test` or similar

**Priority Testing Areas:**
- `common.py` MainWin class methods
- File sanitization functions
- CSV parsing logic
- Authentication flow

## 📝 Documentation

### Code Documentation

**Docstring Standards:**
```python
def getURLResponse(self, url, data=None, retries=3):
    """
    Makes HTTP requests to WebChart with retry logic.
    
    Args:
        url (str): Target URL for the request
        data (dict, optional): POST parameters. Defaults to None.
        retries (int, optional): Number of retry attempts. Defaults to 3.
        
    Returns:
        tuple: (response_bytes, response_object)
        
    Raises:
        Warning: For HTTP or authentication errors
        Exception: For network connectivity issues
    """
```

**Update Documentation:**
- Update relevant `.md` files for functional changes
- Add examples for new features
- Update API documentation for interface changes

### Commit Messages

Use clear, descriptive commit messages:

```
feat: Add retry logic to WebChart authentication

- Implement 3-attempt retry for failed login requests
- Add exponential backoff between attempts  
- Improve error logging for authentication failures

Fixes #123
```

**Commit Message Format:**
- `feat:` - New features
- `fix:` - Bug fixes
- `docs:` - Documentation changes
- `refactor:` - Code refactoring without functional changes
- `test:` - Adding or modifying tests
- `chore:` - Maintenance tasks

## 🚢 Submitting Changes

### Pull Request Process

1. **Create Feature Branch:**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Changes:**
   - Follow coding guidelines
   - Add/update documentation
   - Test thoroughly

3. **Commit Changes:**
   ```bash
   git add .
   git commit -m "feat: description of changes"
   ```

4. **Push Branch:**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **Create Pull Request:**
   - Use descriptive title and description
   - Reference related issues
   - Include testing information
   - Add screenshots for GUI changes

### Pull Request Requirements

**Checklist:**
- [ ] Code follows project style guidelines
- [ ] Changes are documented appropriately
- [ ] Manual testing completed
- [ ] No new warnings or errors introduced
- [ ] Related documentation updated
- [ ] Commit messages are clear and descriptive

**Review Process:**
- Maintainers will review code for quality and compatibility
- Feedback will be provided for necessary changes
- Approved PRs will be merged into the main branch

## 🐛 Bug Reports

### Before Reporting

1. Check existing issues for duplicates
2. Test with latest version
3. Gather relevant system information

### Bug Report Template

```markdown
**Bug Description:**
Clear description of the bug

**Steps to Reproduce:**
1. Step one
2. Step two
3. Step three

**Expected Behavior:**
What should have happened

**Actual Behavior:**
What actually happened

**Environment:**
- OS: [Windows/Linux/macOS]
- Python Version: [e.g., 3.13.2]
- wcexport Version: [commit hash or release version]
- WebChart Version: [if known]

**Log Output:**
Paste relevant log output or error messages

**Additional Context:**
Any other relevant information
```

## 💡 Feature Requests

### Suggesting Features

1. Check existing issues and discussions
2. Provide detailed use case description
3. Consider implementation complexity
4. Discuss with maintainers before large changes

### Feature Request Template

```markdown
**Feature Description:**
Clear description of the proposed feature

**Use Case:**
Why is this feature needed? What problem does it solve?

**Proposed Solution:**
How might this feature be implemented?

**Alternatives Considered:**
Other approaches that were considered

**Additional Context:**
Screenshots, mockups, or other relevant information
```

## 📞 Communication

### Getting Help

- **Documentation:** Check [docs/](docs/) directory first
- **Issues:** Search existing GitHub issues
- **Discussions:** Use GitHub Discussions for questions
- **Code Review:** Comment on pull requests for code-specific questions

### Community Guidelines

- Be respectful and professional
- Provide constructive feedback
- Help others when possible
- Follow project's code of conduct

## 🔧 Build and Release

### Building Executables

```bash
# Build Windows executables
pyinstaller --onefile wcexport.py
pyinstaller --onefile docexport.py

# Executables will be in dist/ directory
```

### Release Process

1. **Version Management:**
   - Update version numbers in relevant files
   - Update CHANGELOG.md with release notes

2. **Testing:**
   - Comprehensive testing on target platforms
   - Verify executable builds work correctly

3. **Documentation:**
   - Update README.md if needed
   - Ensure all documentation is current

4. **Release Creation:**
   - Create GitHub release with appropriate tags
   - Include built executables as release assets
   - Provide release notes

### CI/CD Pipeline

The project uses GitHub Actions for automated builds:
- **Trigger:** Release creation
- **Process:** Build Windows executables and attach to release
- **Configuration:** `.github/workflows/windows-build.yml`

## 📄 License and Legal

By contributing to wcexport, you agree that your contributions will be subject to the same license as the project. Ensure that you have the right to contribute any code or content you submit.

---

Thank you for contributing to wcexport! Your efforts help improve healthcare data integration and workflow efficiency.