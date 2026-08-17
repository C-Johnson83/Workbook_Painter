# Workbook Painter

Workbook Painter is an Excel VBA tool for creating, managing, and applying custom Office color themes to Excel workbooks.

Themes are stored in a CSV file and can be selected, organized, favorited, or applied randomly from the Workbook Painter interface.

## Requirements

* Microsoft Excel for Windows
* VBA macros enabled
* Internet access for AI-powered features
* A Google AI Studio API key

## Installation

1. Download the Workbook Painter ZIP file from GitHub.
2. Extract the entire ZIP to a folder. **Do not open the workbook from inside the ZIP.**
3. Keep the supplied files together in the extracted folder:

   * `Workbook_Painter.xlsm`
   * `PaintbrushThemes.csv`
   * `Painter.ico`
4. Open `Workbook_Painter.xlsm` in Excel.
5. Enable macros/content if prompted.

On first launch, Workbook Painter automatically creates:

```text
C:\Painter\
├── Workbook_Painter.xlsm
├── PaintbrushThemes.csv
└── Painter.ico
```

It also creates a **Workbook Painter** shortcut on the Windows Desktop using the supplied icon.

The Desktop shortcut is recreated whenever Workbook Painter opens, so it will be restored automatically if it is deleted.

After installation, use the Desktop shortcut to launch Workbook Painter.

## Google AI API Key

AI-powered features require a Google AI Studio API key.

Create or manage your API key here:

https://aistudio.google.com/app/api-keys

Once you have a key:

1. Open `Workbook_Painter.xlsm`.
2. Press **Alt + F11** to open the Visual Basic Editor.
3. Locate the **PaintbrushHelpers** module.
4. At the top of the module, locate:

```vb
Public Const API_KEY As String = ""
```

5. Paste your API key between the quotation marks:

```vb
Public Const API_KEY As String = "YOUR_API_KEY_HERE"
```

6. Save the workbook.

> **Security:** Do not commit or distribute your personal API key with the workbook. If Workbook Painter is stored in a public GitHub repository, the distributed workbook should contain a blank API key.

## File Locations

Workbook Painter uses the following constants in the **PaintbrushHelpers** module:

```vb
Option Explicit

Public CURRENT_THEME As String
Public CURRENT_STEP As String

Public Const API_KEY As String = ""

Public Const THEME_FILE As String = "C:\Painter\PaintbrushThemes.csv"
Public Const PAINTER_FOLDER As String = "C:\Painter"
Public Const WORKBOOK_PATH As String = PAINTER_FOLDER & "\Workbook_Painter.xlsm"
Public Const ICON_PATH As String = PAINTER_FOLDER & "\Painter.ico"
```

These constants control where Workbook Painter stores its files.

If you want to install Workbook Painter somewhere other than `C:\Painter`, change these paths in **PaintbrushHelpers** before running the installation from the downloaded copy.

Keep the paths consistent so that `THEME_FILE`, `WORKBOOK_PATH`, and `ICON_PATH` all point to files inside the intended Painter folder.

## Theme File

Themes are stored in:

```text
C:\Painter\PaintbrushThemes.csv
```

During the initial installation, Workbook Painter copies the supplied `PaintbrushThemes.csv` into the Painter directory.

If the installed CSV is later deleted or missing, Workbook Painter automatically creates a replacement CSV containing the required headers.

The CSV contains the theme colors along with supporting information such as category, favorites, tags, usage count, and last-used date.

## Important Notes

* Always extract the ZIP before opening the workbook for the first time.
* Macros must be enabled for installation and Workbook Painter features to function.
* Keep `Workbook_Painter.xlsm`, `PaintbrushThemes.csv`, and `Painter.ico` together for the initial installation.
* After installation, launch Workbook Painter using the Desktop shortcut.
* Back up `PaintbrushThemes.csv` if you have created custom themes. If the file is deleted, Workbook Painter can recreate the file structure, but it cannot recover deleted custom themes.
* Never publish or share your personal Google AI Studio API key.

## Quick Setup

**Download → Extract → Open `Workbook_Painter.xlsm` → Enable Macros → Add API Key → Save → Use Desktop Shortcut**
