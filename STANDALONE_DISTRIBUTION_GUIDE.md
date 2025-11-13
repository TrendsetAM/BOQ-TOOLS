# BOQ-Tools Standalone Distribution Guide

## Overview

This guide explains how to distribute the BOQ-Tools application as a standalone executable that automatically launches in GUI mode and works on any Windows PC without requiring Python installation.

## What You Get

After building, you'll have:
- **`BOQ-Tools.exe`** - A single, standalone executable file
- **Automatic GUI launch** - Double-click to run, no command line needed
- **Complete portability** - Works on any Windows 10/11 PC
- **No dependencies** - All required libraries are bundled

## Distribution Steps

### 1. Build the Executable

```bash
# Quick method (Windows)
double-click build_exe.bat

# Or using Python
python build_exe.py
```

### 2. Test the Executable

The build process automatically runs tests, but you should also:

1. **Double-click the executable** to verify GUI launches
2. **Test basic functionality** (open files, process data)
3. **Test on a clean system** without Python installed

### 3. Distribute the Executable

**Simple Distribution:**
- Copy `dist/BOQ-Tools.exe` to target PCs
- Users can run it directly by double-clicking

**Professional Distribution:**
- Use the generated `create_installer.iss` with Inno Setup
- Creates a proper Windows installer with shortcuts

## For End Users

### System Requirements
- Windows 10/11 (64-bit)
- 4GB RAM minimum (8GB recommended)
- 300MB free disk space
- No Python or other software needed

### How to Use

1. **First Run:**
   - Double-click `BOQ-Tools.exe`
   - Application creates config directory in your user folder
   - GUI launches automatically

2. **Subsequent Runs:**
   - Double-click to launch
   - Faster startup after first run
   - Settings are preserved between sessions

3. **Data Storage:**
   - Config files: `C:\Users\[YourName]\BOQ-Tools\config\`
   - Log files: `C:\Users\[YourName]\BOQ-Tools\logs\`
   - Processed files: Wherever you choose to save them

### Troubleshooting for End Users

#### "Windows protected your PC" message
- This is normal for unsigned executables
- Click "More info" then "Run anyway"
- Consider code signing for professional distribution

#### Slow startup
- First run extracts files to temp directory (slower)
- Subsequent runs should be faster
- Antivirus scanning may slow startup

#### Application won't start
- Check Windows Event Viewer for errors
- Ensure you have admin rights if needed
- Try running as administrator

#### Missing features or errors
- Check log files in `C:\Users\[YourName]\BOQ-Tools\logs\`
- Ensure you have sufficient disk space
- Contact support with log files

## Technical Details

### How It Works

1. **PyInstaller Packaging:**
   - Bundles Python interpreter and all dependencies
   - Creates single executable file
   - Extracts to temp directory on first run

2. **Automatic GUI Launch:**
   - Detects when running as executable (`sys.frozen`)
   - Automatically launches GUI mode
   - No console window appears

3. **Configuration Management:**
   - Creates user-specific config directory
   - Copies default settings on first run
   - Preserves user customizations

4. **Dependency Bundling:**
   - All Python packages included
   - GUI libraries (tkinter, ttkthemes)
   - Data processing libraries (pandas, openpyxl)
   - Image processing libraries (PIL)

### File Structure

```
BOQ-Tools.exe (standalone executable)
├── Python interpreter
├── Application code
├── Dependencies
│   ├── pandas
│   ├── openpyxl
│   ├── tkinter
│   ├── PIL
│   └── other libraries
├── Configuration files
└── Resources
```

### Security Considerations

- **Antivirus Detection:** May be flagged as suspicious (false positive)
- **Code Signing:** Consider signing for professional distribution
- **Firewall:** May prompt for network access (for future updates)
- **Permissions:** Runs with user privileges, no admin needed

## Advanced Distribution Options

### Code Signing (Recommended for Professional Use)

1. Obtain code signing certificate
2. Sign the executable:
   ```bash
   signtool sign /f certificate.pfx /p password BOQ-Tools.exe
   ```

### Creating MSI Installer

1. Use WiX Toolset or similar
2. Include executable and shortcuts
3. Handle registry entries if needed

### Network Deployment

1. Place executable on network share
2. Users can run directly from network
3. Consider performance implications

### Auto-Update System

1. Implement version checking
2. Download new executable when available
3. Replace old version with new one

## Maintenance

### Updating the Application

1. Make code changes
2. Rebuild executable: `python build_exe.py`
3. Test thoroughly
4. Distribute new version

### Version Management

- Update version in `main.py`
- Update installer script version
- Keep changelog for users

### User Support

- Provide clear instructions
- Include troubleshooting guide
- Collect log files for debugging
- Consider remote support tools

## Best Practices

### Before Distribution

- [ ] Test on clean Windows systems
- [ ] Verify all features work
- [ ] Check file size is reasonable
- [ ] Test with different user accounts
- [ ] Verify antivirus compatibility

### For Professional Use

- [ ] Code sign the executable
- [ ] Create proper installer
- [ ] Include documentation
- [ ] Set up support channels
- [ ] Plan update mechanism

### User Experience

- [ ] Provide clear instructions
- [ ] Include system requirements
- [ ] Offer troubleshooting guide
- [ ] Consider creating video tutorials
- [ ] Gather user feedback

## Conclusion

The standalone executable provides a professional, easy-to-distribute solution for BOQ-Tools. Users can run it immediately without any technical setup, making it ideal for business environments where Python installation isn't feasible.

The automatic GUI launch ensures users get the expected experience without needing to understand command-line options or technical details. 