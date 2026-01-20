# BSC.Spreadsheet.SmartExcel
Add-ins for MS Excel

## SmartExcel Add-in

An Excel-DNA add-in that provides powerful data manipulation tools through a custom ribbon tab.

### Features

#### 1. Split Rows Function
Split delimiter-separated strings into new rows while preserving other cell data.
- Select the column containing delimited data
- Choose delimiter type: Comma, Semicolon, Pipe, Tab, or Custom
- Automatically creates new rows with split values
- Copies all other cell values to the newly created rows

#### 2. Change Delimiter Function
Replace delimiters in a column with a different delimiter.
- Select the column to process
- Choose source delimiter (From)
- Choose target delimiter (To)
- Supports: Comma, Semicolon, Pipe, Tab, or Custom delimiters

### Installation

1. Build the project using Visual Studio or MSBuild
2. Navigate to `bin/Debug/net472/` (or `bin/Release/net472/`)
3. Copy `SmartExcel64.xll` (for 64-bit Excel) or `SmartExcel.xll` (for 32-bit Excel)
4. In Excel, go to File → Options → Add-ins
5. Click "Go..." next to "Manage: Excel Add-ins"
6. Click "Browse..." and select the XLL file
7. Check the box next to SmartExcel and click OK

### Usage

Once installed, you'll see a new "SmartExcel" tab in the Excel ribbon with two buttons:

- **Split Rows**: Opens a dialog to split delimited strings into multiple rows
- **Change Delimiter**: Opens a dialog to replace delimiters in selected column

### Building from Source

Requirements:
- .NET Framework 4.7.2 or later
- Visual Studio 2019 or later (or .NET SDK with MSBuild)

Build command:
```bash
dotnet build SmartExcel.csproj -c Release
```

### Technical Details

- Built with Excel-DNA 1.6.0
- Target Framework: .NET Framework 4.7.2
- Uses Office Interop for Excel automation
- Windows Forms for user interface dialogs

