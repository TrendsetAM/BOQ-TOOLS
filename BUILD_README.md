# BOQ-Tools Executable Build Guide

This guide explains how to create a standalone executable (.exe) file for the BOQ-Tools application.

## Prerequisites

- **Python 3.8 or higher** installed on your system
- **pip** package manager
- **Windows 10/11** (for .exe creation)

## Quick Start

### Option 1: Using the Batch File (Recommended for Windows)

1. Double-click `build_exe.bat`
2. Wait for the build process to complete
3. Find your executable in the `dist` folder

### Option 2: Using Python Script

1. Open Command Prompt or PowerShell
2. Navigate to the project directory
3. Run: `python build_exe.py`

## Build Files Explanation

### Core Build Files

- **`build_exe.py`** - Main build script that handles the entire process
- **`build_exe.bat`** - Windows batch file for easy building
- **`boq_tools.spec`** - PyInstaller configuration file
- **`requirements.txt`** - Updated with all necessary dependencies

### Build Process Steps

1. **Environment Check** - Verifies Python version compatibility
2. **Dependency Installation** - Installs all required packages
3. **Icon Creation** - Creates a default application icon if none exists
4. **PyInstaller Execution** - Builds the standalone executable
5. **Verification** - Tests the created executable
6. **Installer Script Creation** - Creates an optional installer script

## Build Options

### Debug Mode
```bash
python build_exe.py --debug
```
Enables detailed debugging information during the build process.

### Skip Cleaning
```bash
python build_exe.py --no-clean
```
Skips cleaning previous build directories (faster for subsequent builds).

## Output Files

After a successful build, you'll find:

- **`dist/BOQ-Tools.exe`** - The standalone executable
- **`build/`** - Temporary build files (can be deleted)
- **`create_installer.iss`** - Inno Setup installer script
- **`build.log`** - Build process log file

## File Size and Performance

- **Expected size**: 80-200 MB (single-file executable with all dependencies)
- **Startup time**: 5-15 seconds (first run extracts to temp directory)
- **Memory usage**: Similar to running the Python script directly
- **GUI Mode**: Automatically launches when double-clicked (no console window)

## Creating an Installer (Optional)

1. Download and install [Inno Setup](https://jrsoftware.org/isinfo.php)
2. Open `create_installer.iss` in Inno Setup
3. Compile the script to create `BOQ-Tools-Setup.exe`

## Troubleshooting

### Common Issues

#### "Python is not installed or not in PATH"
- Install Python from [python.org](https://python.org)
- Make sure to check "Add Python to PATH" during installation

#### "Failed to install dependencies"
- Run: `pip install --upgrade pip`
- Try: `pip install -r requirements.txt` manually
- Check your internet connection

#### "PyInstaller failed"
- Check `build.log` for detailed error messages
- Try running with `--debug` flag
- Ensure all dependencies are properly installed

#### Large executable size
- This is normal for Python applications with many dependencies
- The executable includes the Python interpreter and all libraries

#### Slow startup
- First run is typically slower due to extraction
- Subsequent runs should be faster
- Consider using `--onedir` mode for faster startup (modify spec file)

### Advanced Configuration

#### Modifying the Spec File

Edit `boq_tools.spec` to customize:

- **Icon**: Change the icon path
- **Console mode**: Set `console=True` for debugging
- **Additional files**: Add more data files or resources
- **Hidden imports**: Add modules that PyInstaller misses

#### Example modifications:

```python
# Enable console output for debugging
exe = EXE(
    # ... other parameters ...
    console=True,  # Change to True
    # ... rest of parameters ...
)

# Add custom data files
datas=[
    ('config', 'config'),
    ('resources', 'resources'),
    ('your_custom_folder', 'your_custom_folder'),  # Add this line
],
```

## Distribution

### Single File Distribution
The created `BOQ-Tools.exe` is a standalone executable that can be distributed without Python installation.

### What's Included
- Python interpreter
- All required libraries (tkinter, pandas, openpyxl, etc.)
- Application code
- Configuration files
- Resources
- Automatic GUI launch (no command line needed)

### System Requirements for End Users
- Windows 10/11 (64-bit)
- No Python installation required
- No additional software needed
- Minimum 4GB RAM recommended
- 300MB free disk space (for extraction and data files)

### Standalone Operation
- **Single file**: Just copy `BOQ-Tools.exe` to any Windows PC
- **No installation**: Double-click to run immediately
- **Auto-configuration**: Creates user config directory on first run
- **Portable**: Can run from USB drive or network location

## Security Considerations

- The executable may be flagged by antivirus software (false positive)
- Consider code signing for distribution
- Test on clean systems before distribution

## Version Information

- **Application Version**: 1.0.0
- **Build System**: PyInstaller 5.13.0+
- **Python Version**: 3.8+

## Support

If you encounter issues:

1. Check the `build.log` file for detailed error messages
2. Ensure all dependencies are correctly installed
3. Try building with the `--debug` flag
4. Verify your Python installation is working correctly

## License

This build system is part of the BOQ-Tools project and follows the same license terms. 