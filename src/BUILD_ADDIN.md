# How to Build the PowerPoint Add-In (.ppam)

## Step 1: Prepare the VBA
1. Ensure your PowerPoint file has the latest `ReferenceManager.bas` imported (I have already updated it with the necessary "Callback" functions).

## Step 2: Download RibbonX Editor
1. Go to: [https://github.com/fernandreu/office-ribbonx-editor/releases](https://github.com/fernandreu/office-ribbonx-editor/releases)
2. Download the `.exe` (e.g., `OfficeRibbonXEditor-Installer-....exe` or the `.zip` version specific to your framework, usually .NET 6 or .NET Framework).
3. Install/Run it.

## Step 3: Inject the UI
1. Close PowerPoint completely.
2. Open **Office RibbonX Editor**.
3. Click `Open` and select your PowerPoint Macro-Enabled Presentation (`.pptm`).
4. Right-click on the file name in the left tree view.
5. Select `Insert Office 2010+ Custom UI Part`.
6. A new entry `customUI14.xml` appears. Double-click it.
7. Paste the content of `customUI.xml` (provided in this folder) into the editor.
8. Click `Save` (Disk icon) and close the Editor.

## Step 4: Create the Add-In
1. Open your PowerPoint file again. You should now see a **"Literature"** tab at the top!
2. Go to `File` > `Save As`.
3. Choose file type: **PowerPoint Add-In (*.ppam)**.
4. Save it (e.g., `ReferenceManager.ppam`).

## Step 5: Install (for you or others)
1. Open PowerPoint.
2. Go to `File` > `Options` > `Add-Ins`.
3. At the bottom, manage: **PowerPoint Add-ins** -> `Go...`.
4. Click `Add New...` and pick your `.ppam` file.
5. Done! The "Literature" tab is now permanently available in all your presentations.

## ⚠️ Troubleshooting: "Macro cannot be found" or Security Error

If you click the buttons and get an error about macros being disabled:

**Solution 1: Unblock the File (Critical for downloaded files)**
1. Right-click your `.ppam` file in Windows Explorer.
2. Choose `Properties`.
3. At the bottom of the "General" tab, look for a checkbox **"Unblock"** (Security: This file came from another computer...).
4. Check it and click OK.

**Solution 2: Trusted Locations (Best Practice)**
1. In PowerPoint, go to `File` > `Options` > `Trust Center`.
2. Click `Trust Center Settings...`.
3. Go to `Trusted Locations`.
4. Click `Add new location...`.
5. Browse to the folder where you saved your `.ppam` file.
6. Check **"Subfolders of this location are also trusted"**.
7. Click OK -> OK -> OK.
8. Restart PowerPoint.