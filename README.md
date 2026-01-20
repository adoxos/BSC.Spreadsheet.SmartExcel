# BSC.Spreadsheet.SmartExcel
Add-ins for MS Excel

## SmartExcel Add-in

An Excel-DNA add-in that provides powerful data manipulation tools through a custom ribbon tab.

### Features

#### 1. Split Rows Function
Split delimiter-separated strings into new rows while preserving other cell data.
- Select the column containing delimited data
- Choose delimiter type: Comma, Semicolon, Pipe, Tab, Line break, or Custom
- Automatically creates new rows with split values
- Copies all other cell values to the newly created rows

#### 2. Change Delimiter Function
Replace delimiters in a column with a different delimiter.
- Select the column to process
- Choose source delimiter (From)
- Choose target delimiter (To)
- Supports: Comma, Semicolon, Pipe, Tab, Line break, or Custom delimiters

### Installation

#### For End Users

**See the [END_USER_INSTALLATION_GUIDE.md](END_USER_INSTALLATION_GUIDE.md) for simple, step-by-step instructions.**

#### For Developers/IT Administrators

**Option 1: Automated Installation (Recommended)**

1. Build and publish the project:
   ```bash
   dotnet publish SmartExcel.csproj -c Release -o ./publish
   ```

2. Run the installation script from the publish folder:
   - Double-click `install-x64.bat` (for 64-bit Excel)
   - Enter your organization name when prompted (or press Enter for default)
   - The script will automatically copy files to `%LOCALAPPDATA%\[OrganizationName]\Add-ins\`
   - Follow the on-screen instructions to enable the add-in in Excel

**Option 2: Manual Installation**

1. Build and publish the project (see above)
2. Manually copy files from the `publish` folder to your preferred location
3. In Excel, go to File → Options → Add-ins
4. Click "Go..." next to "Manage: Excel Add-ins"
5. Click "Browse..." and select `SmartExcel64.xll` (for 64-bit Excel)
6. Check the box next to SmartExcel and click OK

#### For Distribution

After running `dotnet publish`:
1. Zip the entire `publish` folder (includes installation scripts and END_USER_INSTALLATION_GUIDE.md)
2. Share the zip file with users
3. Users extract the zip and follow the END_USER_INSTALLATION_GUIDE.md

#### Uninstallation

Run `uninstall-x64.bat` from the publish folder to remove the add-in files from your system.

### Usage

Once installed, you'll see a new "SmartExcel" tab in the Excel ribbon with two buttons:

- **Split Rows**: Opens a dialog to split delimited strings into multiple rows
- **Change Delimiter**: Opens a dialog to replace delimiters in selected column

### Building from Source

Requirements:
- .NET 8 SDK or later
- Visual Studio 2022 or later (or .NET SDK with MSBuild)
- Windows OS (for packing functionality)

Build and publish command:
```bash
dotnet publish SmartExcel.csproj -c Release -o ./publish
```

This will create a `publish` folder containing:
- `SmartExcel64.xll` - 64-bit Excel add-in
- `SmartExcel.xll` - 32-bit Excel add-in  
- `install-x64.bat` - Installation script
- `uninstall-x64.bat` - Uninstallation script
- All required dependencies

**For distribution**: Zip the entire `publish` folder and share with users. Users can extract the zip and run `install-x64.bat` to install.

**Note on Packing:** The project has `RunExcelDnaPack` set to `false` by default for cross-platform builds. On Windows, you can set it to `true` in the .csproj file to enable automatic packing of the XLL with embedded resources. The unpacked XLL files work perfectly for installation.

### Technical Details

- Built with Excel-DNA 1.8.0
- Target Framework: .NET 8 (net8.0-windows)
- Uses Office Interop for Excel automation
- Windows Forms for user interface dialogs

