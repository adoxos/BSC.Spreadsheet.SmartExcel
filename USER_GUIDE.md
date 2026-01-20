# SmartExcel User Guide

## Overview

SmartExcel is an Excel add-in that provides powerful tools for manipulating delimited data directly within Excel. The add-in adds a custom "SmartExcel" tab to your Excel ribbon with two main functions.

## Installation Instructions

### Step 1: Build the Add-in (if not already built)

If you have the source code:
1. Open a command prompt in the project directory
2. Run: `dotnet build SmartExcel.csproj -c Release`
3. The XLL files will be created in `bin/Release/net472/`

### Step 2: Install in Excel

1. Open Microsoft Excel
2. Go to **File → Options → Add-ins**
3. At the bottom of the dialog, select **"Excel Add-ins"** from the "Manage:" dropdown
4. Click **"Go..."**
5. Click **"Browse..."**
6. Navigate to the build output folder:
   - For 64-bit Excel: Select `SmartExcel64.xll`
   - For 32-bit Excel: Select `SmartExcel.xll`
7. Click **OK**
8. Ensure the SmartExcel checkbox is checked
9. Click **OK** to close the dialog

### Step 3: Verify Installation

- You should now see a **"SmartExcel"** tab in your Excel ribbon
- The tab contains a **"Data Tools"** group with two buttons

## Function 1: Split Rows

### Purpose
Split delimiter-separated values in a column into multiple rows, duplicating data from other columns.

### Example Use Case
Convert this:
```
Name         | Hobbies
John Smith   | Reading,Swimming,Hiking
Jane Doe     | Cooking,Gaming
```

Into this:
```
Name         | Hobbies
John Smith   | Reading
John Smith   | Swimming
John Smith   | Hiking
Jane Doe     | Cooking
Jane Doe     | Gaming
```

### How to Use

1. Open your Excel workbook with data
2. Navigate to any sheet with data
3. Click the **"SmartExcel"** tab in the ribbon
4. Click the **"Split Rows"** button
5. In the dialog:
   - **Select Column**: Choose the column containing delimited data (e.g., "Column B")
   - **Select Delimiter**: Choose from:
     - Comma (,)
     - Semicolon (;)
     - Pipe (|)
     - Tab
     - Custom (enter your own delimiter)
6. Click **"Execute"**
7. The operation will:
   - Split the values in the selected column
   - Create new rows for each split value
   - Copy all other column data to the new rows

### Important Notes
- The function processes data from bottom to top to avoid row shifting issues
- All data in other columns will be preserved and duplicated for each new row
- Screen updating is disabled during processing for better performance

## Function 2: Change Delimiter

### Purpose
Replace one delimiter with another in a selected column.

### Example Use Case
Convert this:
```
Tags
tag1,tag2,tag3
apple,orange,banana
```

Into this:
```
Tags
tag1; tag2; tag3
apple; orange; banana
```

### How to Use

1. Open your Excel workbook with data
2. Navigate to any sheet with data
3. Click the **"SmartExcel"** tab in the ribbon
4. Click the **"Change Delimiter"** button
5. In the dialog:
   - **Select Column**: Choose the column to process
   - **From Delimiter**: Choose the current delimiter to replace
   - **To Delimiter**: Choose the new delimiter to use
   - Both support:
     - Comma (,)
     - Semicolon (;)
     - Pipe (|)
     - Tab
     - Custom (enter your own delimiter)
6. Click **"Execute"**
7. All instances of the "from" delimiter in the selected column will be replaced with the "to" delimiter

### Important Notes
- Only cells containing the "from" delimiter will be modified
- Empty cells are skipped
- The operation uses Excel's calculation and screen updating optimization for better performance

## Tips and Best Practices

1. **Backup Your Data**: Always save a copy of your workbook before performing operations on important data

2. **Test on Sample Data**: Try the functions on a small sample of data first to ensure they work as expected

3. **Column Selection**: Make sure to select the correct column - the functions work on the entire column's used range

4. **Custom Delimiters**: When using custom delimiters, ensure you enter the exact character(s) used in your data

5. **Performance**: For large datasets, the operations may take a few seconds. Excel will show "Not Responding" but will complete the operation

## Troubleshooting

### The SmartExcel tab doesn't appear
- Verify the add-in is enabled in **File → Options → Add-ins**
- Restart Excel
- Check that you're using the correct version (32-bit vs 64-bit)

### "No data found in the active sheet" message
- Ensure your active worksheet contains data
- Check that the data is in Excel's used range

### Operation seems to hang
- For large datasets, operations may take time
- Wait for Excel to complete the operation
- Check Excel's status bar for progress

### Unexpected results
- Verify you selected the correct column
- Check that you're using the right delimiter
- Ensure your data doesn't contain nested or escaped delimiters

## System Requirements

- Microsoft Excel 2010 or later (Windows)
- .NET Framework 4.7.2 or later
- Windows operating system

## Support and Feedback

For issues, questions, or feature requests, please contact the repository maintainer or create an issue in the GitHub repository.

## Version Information

- **Version**: 1.0.0
- **Built with**: Excel-DNA 1.6.0
- **Target Framework**: .NET Framework 4.7.2
