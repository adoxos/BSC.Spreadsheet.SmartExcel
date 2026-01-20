# Changelog

All notable changes to the SmartExcel add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-01-20

### Changed
- Rolled the add-in back to target .NET Framework 4.8 (net48) for maximum compatibility with legacy Excel environments.
- Updated Excel-DNA configuration files and documentation to reflect the .NET Framework runtime requirements.

## [1.0.0] - 2026-01-20

### Added
- Initial release of SmartExcel add-in
- Custom "SmartExcel" ribbon tab in Excel
- **Split Rows** function:
  - Split delimiter-separated strings into multiple rows
  - Support for Comma, Semicolon, Pipe, Tab, Line break, and Custom delimiters
  - Automatic duplication of other column data to new rows
  - Column selection interface
- **Change Delimiter** function:
  - Replace delimiters in a column with different delimiters
  - Support for Comma, Semicolon, Pipe, Tab, Line break, and Custom delimiters
  - From/To delimiter selection
  - Column selection interface
- Windows Forms dialogs for user interaction
- Comprehensive README with installation and usage instructions
- USER_GUIDE.md with detailed examples and troubleshooting
- MIT License
- Build configuration for both 32-bit and 64-bit Excel
- **Installation scripts**:
  - `install-x64.bat` - Automated installation script for 64-bit Excel
  - `uninstall-x64.bat` - Uninstallation script
  - Scripts prompt user for organization name (no hardcoded values)
  - Scripts copy files to `%LOCALAPPDATA%\[OrganizationName]\Add-ins\`
- **END_USER_INSTALLATION_GUIDE.md** - Comprehensive guide for end users with step-by-step installation and usage instructions

### Changed
- Updated to .NET 8 (net8.0-windows) from .NET Framework 4.7.2
- Updated Excel-DNA to version 1.8.0

### Technical Details
- Built with Excel-DNA 1.8.0
- Target Framework: .NET 8 (net8.0-windows)
- Uses Excel-DNA dynamic COM access (no Microsoft Office Interop reference required)
- COM-visible ribbon controller
- Optimized for performance with screen updating and calculation control
