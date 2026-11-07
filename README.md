<img width="2172" height="724" alt="Image" src="https://github.com/user-attachments/assets/2df8698b-26e5-42df-9778-fda65ab99337" />

## Overview

**VBA UserForm Transpiler** is a GUI-based tool that provides a unified interface for converting Excel VBA UserForms into GUI code for various target platforms and programming languages.

It allows you to convert UserForms created in Microsoft Excel VBA into GUI code for various target platforms and programming languages.

| Target | Generated Code | GUI Framework |
|---|---|---|
| Python | `.py` | Tkinter |
| PowerShell | `.ps1` / `.bat` | WinForms |
| HTML | `.html` | HTML / CSS |

<img width="1677" height="587" alt="Image" src="https://github.com/user-attachments/assets/491f411e-6e84-4dbf-a921-eeb193c4542e" />
<img width="1248" height="448" alt="Image" src="https://github.com/user-attachments/assets/9e38f2f8-5232-4adf-86d3-083dd095abfb" />

## Features

- **Unified Interface** — Manage multiple conversion engines through a single, unified GUI.
- **Multiple Output Targets** — Convert VBA UserForms into GUI code for various target platforms and programming languages (e.g., Python/Tkinter, PowerShell/WinForms, and HTML/CSS).
- **Workbook Selection** — Select the Excel workbook containing the UserForms you want to convert.
- **UserForm Selection** — Choose a specific UserForm from the selected workbook.
- **Reload** — Refresh the available workbooks and UserForms.
- **Conversion Options** — Configure additional options depending on the selected target.
- **Image Handling** — Configure how images embedded in UserForms are exported or referenced.


## Usage
<img width="423" height="499" alt="Image" src="https://github.com/user-attachments/assets/22eee50d-b973-45b7-a921-efc6fdee03f1" />

The application provides a simple workflow for converting UserForms.

Before starting, open the Excel workbook containing the UserForms you want to convert in Microsoft Excel.

1. Select the source workbook from the list of currently open workbooks.
2. Select the UserForm you want to convert.
3. Choose the target platform.
4. Configure the available conversion options.
5. Click **Convert**.

The interface automatically presents the options relevant to the selected conversion target.

## Requirements

### Excel Security Settings

To use this tool, you must enable the following setting in Excel:
1.  Open a workbook.
2.  Go to **File → Options → Trust Center → Trust Center Settings**.
3.  Select **Macro Settings**.
4.  Check **Trust access to the VBA project object model**.
5.  Click **OK**.

### Operating Environments

- Microsoft Windows
- Microsoft Excel with VBA support
- An Excel workbook containing the UserForms to be converted

The original conversion engines have been tested with various Windows and Microsoft Office versions, including Windows 10/11 and both 32-bit and 64-bit Office installations.

Microsoft Excel 2016 or later is recommended.

<details>
<summary>Running the Application in Older Versions of Microsoft Office</summary>

To use this application with Office 2000, Office XP, or Office 2003, you must install the **Microsoft Office Compatibility Pack** on your PC.

Download the file from [here](http://web.archive.org/web/20120201204132/http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=3), install it, and then restart your PC.

If you want to download the file directly, copy and paste the following URL into your browser's address bar:
```
http://web.archive.org/web/20111016032501/http://download.microsoft.com/download/9/2/2/9222d67f-7630-4f49-bd26-476b51517fc1/FileFormatConverters.exe
```
After installing the Microsoft Office Compatibility Pack, opening an `.xlsm` file will convert it to `.xls` format. However, the converted `.xls` file will be placed in a temporary folder, which prevents the application from working correctly.

Before using the application, please use **Save As** to save the `.xls` file to a different folder.

</details>

### Generated Code Runtime Requirements

Depending on the selected target:

| Target | Runtime |
|---|---|
| Python / Tkinter | Python 3.4 or later |
| PowerShell / WinForms | Windows PowerShell 5.1 or later |
| HTML / CSS | A compatible web browser |

Additional requirements may apply depending on the generated code and execution environment.

## Conversion Options

Available options depend on the selected output target.

| Option | Python (Tkinter) | PowerShell (WinForms) | HTML/CSS |
|---|:---:|:---:|:---:|
| Use Class/Prefix | ✓ | ✓ | ✓ |
| Save As `.bat` | — | ✓ | — |
| No Mainloop | ✓ | ✓ | — |
| Unique Style Name | ✓ | — | — |
| Language Attribute | — | — | ✓ |
| Image Mode | ✓ | ✓ | ✓ |

### Option Descriptions

- **Use Class/Prefix** — Generate the output using a class-based structure or apply a specified prefix to generated identifiers, depending on the target platform. This helps organize the generated code and avoid naming conflicts.

- **Save As `.bat`** — The generated PowerShell script will be saved as a `.bat` file that can be executed by double-clicking.

- **No Mainloop** — Omit the main event loop or display method from the generated code. This is useful when integrating the generated GUI into an existing application or managing the GUI lifecycle externally.

- **Unique Style Name** — Generate unique names for Tkinter styles to prevent naming conflicts with other styles in the application.

- **Language Attribute** — Specify the language attribute (`lang`) of the generated HTML document. This can be used to indicate the language of the document to browsers and other tools.

- **Image Mode** — Specify how images embedded in VBA UserForms are handled during conversion. Available modes depend on the selected target platform:
  - `file` — Export images as separate files.
  - `disabled` — Disable image processing.
  - `reference-only` — Generate references to existing image files without exporting them.
  - `base64` — Embed images directly into the generated code as Base64-encoded data.
  - `base64-dict` — Store Base64-encoded image data in an associative array (dictionary) supported by the target language.
  - `base64-json` — Store Base64-encoded image data in an external JSON file.
  - `base64-json-reference` — Reference an existing JSON image file without exporting it.

## Output Directory

All generated files are automatically saved in a dedicated output folder created in the same directory as `VBA-UserForm-Transpiler.xlsm`.

The output folder name depends on the selected conversion target:

| Target | Output Folder |
|---|:---:|
| Python (Tkinter) | `VBAForm2Tkinter_output` |
| PowerShell (WinForms) | `VBAForm2PS_output` |
| HTML/CSS | `VBAForm2HTML_output` |

### Directory Structure

The following example shows the output directory structure for Python (Tkinter):

```
ApplicationFolder/
├─ VBA-UserForm-Transpiler.xlsm
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ exported images...
```

All generated files, including exported images and auxiliary files, are stored in the corresponding output folder.

## How It Works

VBA UserForm Transpiler integrates multiple conversion engines to support different target platforms and programming languages:

```text
                  VBA UserForm Transpiler
                            |
                            |
                    Select Workbook
                            |
                            v
                     Select UserForm
                            |
                            v
                    Select Output Target
                            |
              +-------------+-------------+
              |             |             |
              v             v             v
          Python        PowerShell      HTML/CSS
          Tkinter        WinForms        Browser
              |             |             |
              v             v             v
           output.py     output.ps1     output.html
                         output.bat
```

Each conversion target uses its corresponding conversion engine rather than attempting to translate VBA directly into a single intermediate format.

This approach preserves the functionality and conversion behavior of the original projects while providing a consistent user experience.

## Conversion Limitations

This tool focuses on converting the visual structure and properties of VBA UserForms.

The following points should be considered:

- Generated GUI layouts are approximations of the original VBA UserForms.
- Not every VBA control or property can be represented identically in the target framework.
- Unsupported controls may cause conversion to fail.
- Some properties behave differently between VBA and the target GUI framework.
- VBA event procedures and application-specific business logic are not automatically translated into equivalent event handlers or executable code in the target environment.
- Additional programming may be necessary to make the generated GUI fully functional.

The generated code should therefore be treated as a starting point for porting a VBA-based GUI, rather than a complete replacement for the original VBA application.

For detailed information about supported controls and conversion limitations, please refer to the documentation of each original project.

## Original Projects

This project integrates the following open-source projects:

| Project | Description |
|---|---|
| [VBAForm2Tkinter](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter) | VBA UserForm to Python Tkinter |
| [VBAForm2PowerShell](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell) | VBA UserForm to PowerShell WinForms |
| [VBAForm2HTML](https://github.com/GUI-Conversion-Tools/VBAForm2HTML) | VBA UserForm to HTML/CSS |

Each project provides its own detailed documentation, including supported controls, conversion behavior, parameters, and limitations.

If you need more detailed information about a particular conversion target, please consult its original repository.


---
