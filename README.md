# Client Export Tool
**Developed by Chetan Mehta**

## Usage & Distribution Notice

This tool is currently under active development and is provided **only to individuals who received this build directly from the developer**.

Please **do not redistribute, repackage, modify, or resell** this software.

Use of this tool is limited to approved recipients of this build.

---

# Overview

The **Client Export Tool** is an Excel add-in that creates a **sanitized client-safe version of a workbook** by removing proprietary logic, internal comments, hidden data, and other sensitive elements.

It is designed to allow teams to safely share spreadsheets externally without exposing internal models, calculations, or development artifacts.

The tool operates on the currently open workbook and produces a **clean exported copy** in the same folder as the source file.

---

# What This Tool Does

When the export process runs, the tool performs a series of cleaning and transformation steps to ensure that only the intended data remains visible.

## Data Sanitization

- Converts **all formulas to values** while preserving cell formatting.
- Removes **all comments and notes** across the workbook.
- Deletes **hidden worksheets** (including very hidden sheets).
- Removes **hyperlinks**.
- Clears **document metadata** such as author, company, and internal descriptions.
- Deletes **external workbook links**.
- Removes **external data connections and queries**.
- Removes **external named ranges and references**.

## Grouped Row/Column Handling

Users can choose how grouped rows and columns are processed:

Options include:

- **Clear grouped data**
  - Deletes the contents of grouped cells while preserving formatting.

- **Flatten grouped data**
  - Converts formulas to values within grouped cells.

Additionally, users can choose whether to:

- **Keep groups collapsed**
- **Expand groups**

This ensures grouped sections can either remain visually identical or be expanded before export.

## Hidden Row/Column Handling

Hidden rows and columns can be processed independently from grouped data.

Options include:

- **Clear hidden data**
- **Flatten hidden data**
- **Keep hidden rows/columns hidden**
- **Expand (unhide) hidden rows/columns**

This allows internal data sections to be fully removed or safely exposed depending on export needs.

## Image Handling

Images (such as screenshots or imported graphics) can be processed with three options:

- **Keep all images**
- **Keep only images inside the printable area**
- **Delete all images**

## Chart Handling

Charts can be handled separately from images with the same options:

- **Keep all charts**
- **Keep charts only inside the printable area**
- **Delete all charts**

## Print Area Cleanup

For worksheets with defined print areas, the tool:

- Deletes rows and columns **outside the print area**
- Removes objects outside the printable region (depending on selected options)

This helps ensure exported workbooks contain only the content intended for presentation.

---

# Export Behavior

The tool does **not modify the original workbook**.

Instead, it:

1. Creates a copy of the current workbook
2. Applies all sanitization rules
3. Saves the cleaned workbook alongside the original file

The output file will be named:

```
OriginalWorkbookName_CLIENT_EXPORT.xlsx
```

---

# Installation Instructions

Go to the [latest release](https://github.com/cheetah-tw/excel-client-export/releases/latest) and, under **Assets**, download the zip for your computer:

- **Windows:** `ClientExportTool-<version>-Windows.zip` (the add-in + `Install-Windows.cmd`)
- **Mac:** `ClientExportTool-<version>-macOS.zip` (the add-in + `install-mac.command`)

Ignore the "Source code" downloads; those are for development.

Windows and macOS mark files downloaded from the internet as untrusted, and Excel refuses to load an add-in that is still marked. The installers copy the add-in to Excel's add-in folder and remove that mark for you.

## Windows

1. Right-click the downloaded Windows zip and choose **Extract All… → Extract**. Don't run the installer from inside the zip.
2. In the extracted folder, double-click **`Install-Windows.cmd`**.
   - If Windows shows **"Windows protected your PC"**, click **More info → Run anyway**.
   - If it shows **"Open File - Security Warning"**, click **Run**.
   - These warnings appear because the installer was downloaded. They are expected.
3. If Excel is open, the installer asks you to close it. Save your work, close Excel, and press **Enter**.
4. When the window says **Done!**, press any key to close it and open Excel.

The tool appears in the Excel ribbon under:

**Review → Client Export → Create Client Export**

The installer puts the add-in in `%APPDATA%\Microsoft\AddIns` (one of Excel's trusted locations), turns it on in Excel, and removes leftover entries from earlier installs, such as one pointing at your Downloads folder.

### Manual install (Windows)

Only needed if you can't run the installer.

1. Copy `ClientExportTool.xlam` to `C:\Users\YourUsername\AppData\Roaming\Microsoft\AddIns`.
2. Right-click the copied file → **Properties** → tick **Unblock** at the bottom of the General tab → **OK**. (If there is no Unblock box, the file is already unblocked.)
3. Open Excel → **File → Options → Add-ins** → **Manage: Excel Add-ins → Go**.
4. Click **Browse**, select `ClientExportTool.xlam`, and click **OK**.
5. Restart Excel.

---

## Mac

1. Open the downloaded Mac zip. Safari usually unzips it automatically; otherwise double-click it in Finder.
2. Open **Terminal** (press **Cmd + Space**, type `Terminal`, press **Return**).
3. Type `bash` followed by a **space**, then drag **`install-mac.command`** from Finder into the Terminal window, and press **Return**.
   - Running it this way avoids macOS blocking the downloaded script.
   - If macOS asks whether **Terminal** may access data from other apps, click **Allow**. (The installer needs this to put the add-in in Excel's folder.)
4. If Excel is open, the installer asks you to quit it (**Cmd + Q**) and press **Return**.
5. First install only: in Excel, choose **Tools → Excel Add-ins…**, tick **ClientExportTool**, and click **OK**.

The tool appears in the Excel ribbon under:

**Review → Client Export → Create Client Export**

### Manual install (Mac)

Only needed if you can't run the installer.

1. In Finder, choose **Go → Go to Folder…** and enter `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins` (create the `Add-Ins` folder if it doesn't exist).
2. Copy `ClientExportTool.xlam` into that folder.
3. In Terminal, run:
   `xattr -d com.apple.quarantine ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/ClientExportTool.xlam`
4. In Excel, choose **Tools → Excel Add-ins…**, tick **ClientExportTool**, and click **OK**.

---

# Usage

1. Open the workbook you want to export.
2. Ensure the workbook has been **saved**.
3. Click:

**Review → Client Export → Create Client Export**

4. Choose the export options in the settings window.
5. Click **Create Export**.

The cleaned workbook will be created in the same folder as the original file.

---

# Notes

- After installation, do **not rename or move the add-in file** in Excel's add-in folder.
- The workbook being exported must be saved first so the tool can create the export copy in the same directory.
- If Excel blocks macros, click **Enable Content** when prompted.
- **Mac:** Excel for Mac only writes to folders you've allowed. The first time you export from a folder, Excel asks for access to it; click **Grant Access**. For a workbook in OneDrive, Excel may also ask once for access to the `CloudStorage` folder so it can find your synced files.

---

# Updating from a Previous Version

Download the new release zip for your computer and run the installer again, exactly as in the installation steps above. It replaces the old version in place.

- **Windows:** run `Install-Windows.cmd`. Nothing else is needed; the add-in stays turned on.
- **Mac:** run `install-mac.command`. If **ClientExportTool** is already ticked in **Tools → Excel Add-ins…**, nothing else is needed.

> **Important:** Excel must be fully closed while updating. Excel caches add-ins in memory and keeps running the old version until it is restarted. The installers wait for you to close Excel before continuing.

If you installed an earlier version by hand from another folder (for example Downloads), the Windows installer cleans that up. On Mac, untick the old entry in **Tools → Excel Add-ins…** if two appear.

---

# Troubleshooting

**"This file type is not supported in protected view" or "Microsoft Excel cannot access the file '…\Downloads\ClientExportTool.xlam'"** (Windows)
The add-in was loaded directly from Downloads while Windows still marked it as downloaded. Run `Install-Windows.cmd`. It installs an unblocked copy and removes the old Downloads entry so the error stops appearing when Excel starts.

**Red bar: "Microsoft has blocked macros from running because the source of this file is untrusted"**
Same cause and same fix: run the installer for your system (or follow the manual install steps, including the Unblock / `xattr` step).

**The Client Export button isn't on the Review tab**
Check that **ClientExportTool** is ticked in the Excel Add-ins dialog (**File → Options → Add-ins → Manage: Excel Add-ins → Go** on Windows, **Tools → Excel Add-ins…** on Mac), then restart Excel.

**The Windows installer says it can't find `ClientExportTool.xlam`**
Extract the zip first, and keep `Install-Windows.cmd` in the same folder as the `.xlam`.

**"Excel doesn't have permission to save in this folder"** (Mac)
Excel asked for access to the folder and it wasn't granted. Run the export again and click **Grant Access**.

**"Could not resolve a local file path for this workbook"**
The workbook is open from OneDrive or SharePoint but the tool couldn't find its synced copy on this computer. Make sure the folder is synced (OneDrive app running, folder not set to online-only), then open the workbook from File Explorer (Windows) or Finder (Mac) and export again.

---

# Support

If you encounter issues or have questions about this tool, please contact the developer.
