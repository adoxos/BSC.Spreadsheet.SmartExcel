# SmartExcel Add-in - End User Installation Guide

## Quick Start Installation (Recommended)

### Step 1: Extract the Files
1. Extract the zip file you received to a folder on your computer
2. You should see files including `install-x64.bat`, `SmartExcel64.xll`, and others

### Step 2: Run the Installation Script
1. **Double-click** the `install-x64.bat` file
2. When prompted, enter your organization name (e.g., "Statkraft AS", "Contoso", etc.)
   - If you don't have an organization name, just press **Enter** to use the default location
3. The script will:
   - Create the necessary folders
   - Copy all required files
   - Show you where the files were installed

### Step 3: Enable the Add-in in Excel
1. Open **Microsoft Excel**
2. Click **File** → **Options**
3. In the Excel Options window, click **Add-ins** (in the left sidebar)
4. At the bottom of the window, find the **Manage** dropdown
5. Select **Excel Add-ins** from the dropdown
6. Click the **Go...** button
7. In the Add-ins dialog:
   - Click **Browse...**
   - Navigate to the folder shown at the end of the installation script
     - Example: `C:\Users\YourName\AppData\Local\YourOrganization\Add-ins\`
   - Select **SmartExcel64.xll**
   - Click **OK**
8. Make sure the checkbox next to **SmartExcel** is checked
9. Click **OK** to close the dialog

### Step 4: Verify Installation
- Look for a new **SmartExcel** tab in your Excel ribbon
- The tab should contain two buttons: **Split Rows** and **Change Delimiter**

---

## Manual Installation (Alternative Method)

If you prefer not to use the installation script, follow these steps:

### Step 1: Choose Installation Location
1. Open File Explorer
2. Navigate to: `%LOCALAPPDATA%`
   - Tip: You can copy this path and paste it into the File Explorer address bar
3. Create a folder structure like: `YourOrganization\Add-ins\`
   - Example: `C:\Users\YourName\AppData\Local\Contoso\Add-ins\`

### Step 2: Copy Files
Copy these files from the extracted folder to your installation location:
- `SmartExcel64.xll` (main add-in file)
- `SmartExcel.dll`
- `SmartExcel64.dna`
- `SmartExcel.xll` and `SmartExcel.dna` (optional, only if 32-bit Excel support is needed)

### Step 3: Enable in Excel
Follow **Step 3** from the Quick Start Installation above.

---

## Using the SmartExcel Add-in

Once installed, you'll have access to two powerful tools:

### Split Rows
**Purpose:** Split text with delimiters (like commas) into separate rows

**How to use:**
1. Click the **SmartExcel** tab in Excel
2. Click **Split Rows**
3. Select the column containing delimited data
4. Choose your delimiter type (Comma, Semicolon, Pipe, Tab, Line break, or Custom)
5. Click **Execute**

**Example:**
- Before: One row with "Apple,Banana,Orange" in column A
- After: Three rows with "Apple", "Banana", "Orange" each in their own row

### Change Delimiter
**Purpose:** Replace one delimiter with another in your data

**How to use:**
1. Click the **SmartExcel** tab in Excel
2. Click **Change Delimiter**
3. Select the column to process
4. Choose the **From** delimiter (what you want to replace)
5. Choose the **To** delimiter (what you want to replace it with)
6. Click **Execute**

**Example:**
- Before: "Apple,Banana,Orange" (comma-separated)
- After: "Apple;Banana;Orange" (semicolon-separated)

---

## Troubleshooting

### The SmartExcel tab doesn't appear in Excel
**Solution:**
1. Check that the add-in is enabled:
   - Go to File → Options → Add-ins
   - Look for SmartExcel in the **Active Application Add-ins** list
   - If it's in **Inactive** or **Disabled** items, click Go... and enable it
2. Try restarting Excel
3. Make sure you're using 64-bit Excel (this add-in requires 64-bit Excel)

### "File not found" error when browsing for the add-in
**Solution:**
- Make sure you navigated to the correct folder shown during installation
- Check that the `SmartExcel64.xll` file exists in that location
- If using manual installation, verify all files were copied

### Security warning when enabling the add-in
**Solution:**
- This is normal for Excel add-ins
- Click **Enable Content** or **Trust this add-in** if prompted
- Contact your IT administrator if you cannot enable the add-in due to security policies

### The add-in crashes or doesn't work
**Solution:**
1. Make sure you have **.NET Framework 4.8** installed (Windows 10 1903 and later include it by default)
   - If needed, download the offline installer from: https://dotnet.microsoft.com/download/dotnet-framework/net48
2. Restart Excel after installing .NET Framework 4.8
3. Try uninstalling and reinstalling the add-in

---

## Uninstalling the Add-in

### Method 1: Using the Uninstall Script
1. Navigate to the folder where you extracted the files
2. Double-click `uninstall-x64.bat`
3. Enter the same organization name you used during installation
4. Confirm the uninstallation when prompted

### Method 2: Manual Uninstall
1. In Excel, go to File → Options → Add-ins
2. Select **Excel Add-ins** from the Manage dropdown and click **Go...**
3. Uncheck **SmartExcel** and click **OK**
4. Close Excel
5. Delete the folder where the add-in files were installed

---

## System Requirements

- **Operating System:** Windows 10 or later
- **Excel Version:** Microsoft Excel 2010 or later (64-bit version required)
- **.NET Runtime:** .NET Framework 4.8 (pre-installed on modern Windows 10/11)
- **Permissions:** Ability to install software in your local AppData folder

---

## Getting Help

If you encounter issues not covered in this guide:
1. Check the `USER_GUIDE.md` file for detailed technical information
2. Contact your system administrator
3. Refer to the repository documentation for advanced troubleshooting

---

## Version Information

- **SmartExcel Version:** 1.0.0
- **Last Updated:** January 2026
