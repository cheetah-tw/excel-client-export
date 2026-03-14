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

## Windows

1. Save `ClientExportTool.xlam` somewhere permanent on your computer.

Recommended location:

`C:\Users\YourUsername\AppData\Roaming\Microsoft\AddIns`

2. Open **Excel**.

3. Go to:

**File → Options → Add-ins**

4. At the bottom of the window:

**Manage: Excel Add-ins → Go**

5. Click **Browse**.

6. Select `ClientExportTool.xlam`.

7. Click **OK**.

8. Restart Excel if prompted.

After installation, the tool will appear in the Excel ribbon under:

**Review → Client Export → Create Client Export**

---

## Mac

1. Save `ClientExportTool.xlam` somewhere permanent on your computer.

2. Open **Excel**.

3. Go to the menu bar and select:

**Tools → Excel Add-ins**

4. Click **Browse**.

5. Select `ClientExportTool.xlam`.

6. Click **OK**.

After installation, the tool will appear in the Excel ribbon under:

**Review → Client Export → Create Client Export**

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

- The `.xlam` file must remain on your computer in the same location after installation.
- Do **not rename or move the add-in file** after installation.
- The workbook being exported must be saved first so the tool can create the export copy in the same directory.
- If Excel blocks macros, click **Enable Content** when prompted.

---

# Support

If you encounter issues or have questions about this tool, please contact the developer.
