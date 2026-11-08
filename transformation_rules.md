# Transformation Rules
- [Python (Tkinter)](#Python-Tkinter)
- [PowerShell (WinForms)](#PowerShell-WinForms)
- [HTML](#HTML)

# Python (Tkinter)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Z-order of controls
- Control colors (foreground)
- Control colors (background)
- Text display (`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`UserForm`, `Frame`, `TextBox`, `Label`, `ListBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox` [.MultiLine=False], `ComboBox`, `CheckBox`, `ToggleButton`, `OptionButton`, `ListView`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`, `ListView`, `TreeView`
- Selection state of `OptionButton`, `CheckBox` and `ToggleButton`
- Transparent background setting specified in `.BackStyle`(`Label`, `TextBox`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `Image`, `ComboBox`)
- Images Embedded in Controls (`Image`)
- `.Visible` property
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar` property (`TextBox` [.MultiLine=False])
- `.ScrollBars` property (`TextBox` [.MultiLine=True])
- `.WordWrap` property (`TextBox` [.MultiLine=True])
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`, `ListView`)
- `.PictureAlignment` property (`Image`)
- `.Scroll` property (`TreeView`)
- `.Expanded` property (`TreeView.Nodes`)

> Note:  
> - When `.BackStyle = fmBackStyleTransparent`, true transparency is not supported in Tkinter. Instead, the background color is substituted as follows:
>   -   If the parent control has a `.BackColor`, that color is used.
>   -   If the parent is a `Page` (which does not expose `.BackColor`), a system default color (`&H8000000F&`) is used as a fallback, which matches the visual background color of the `Page`.
> 
> - Since Tkinter does not support displaying resized images natively, when exporting images with the `.PictureSizeMode` property set to `fmPictureSizeModeStretch` or `fmPictureSizeModeZoom`, the embedded images are exported after being resized to match the control's dimensions. This preserves the appearance of the original VBA UserForm as closely as possible.

## Supported Controls
| VBA Form Class | Tkinter Class|
| ------ | ------ |
| `Label` | `tk.Label` |
| `CommandButton` | `tk.Button` |
| `Frame` (without Caption) | `tk.Frame` |
| `Frame` (with any Caption) | `tk.LabelFrame` |
| `TextBox` (`MultiLine=False`) | `tk.Entry` |
| `TextBox` (`MultiLine=True`) | `tk.Text` |
| `SpinButton` | `tk.Spinbox` |
| `ListBox` | `tk.Listbox` |
| `CheckBox` | `tk.Checkbutton` |
| `ToggleButton` | `tk.Checkbutton`(`indicatoron=0`) |
| `OptionButton` | `tk.Radiobutton` |
| `Image` | `tk.Canvas` |
| `ScrollBar` | `ttk.Scale` |
| `ComboBox` | `ttk.Combobox` |
| `MultiPage` | `ttk.Notebook` |
| `ListView`(`.View=lvwReport`) | `ttk.Treeview`(`show="headings"`) |
| `TreeView` | `ttk.Treeview`(`show="tree"`) |

> Note:
>- `SpinButton` behaves differently in VBA and Tkinter, so appearance may vary depending on placement.
>- `ScrollBar` in VBA has up/down adjustment buttons, but Tkinter’s `Scale` does not.


# PowerShell (WinForms)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Z-order of controls
- Control colors (foreground) (Excluding `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList])
- Control colors (background) (Excluding `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList], `ScrollBar`)
- Text display (`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`Frame [without Caption]`, `TextBox`, `Label`, `ListBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox`, `CheckBox`, `ToggleButton`, `OptionButton`,  `ListView`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`, `ListView`, `TreeView`
- Selection state of `OptionButton`, `CheckBox` and `ToggleButton`
- Transparent background setting specified in `.BackStyle`(Excluding `ComboBox` [.Style = fmStyleDropDownList])
- Images Embedded in Controls (`Image`)
- `.Visible` property
- `.Orientation`/`.Min`/`.Max` property (`ScrollBar`)
- `.Alignment` property (`CheckBox`. `OptionButton`)
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar` property (`TextBox`)
- `.ScrollBars` property (`TextBox`)
- `.WordWrap` property (`TextBox`)
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`, `ListView`)
- `.PictureAlignment`/`.PictureSizeMode` property (`Image`)
- `.Scroll` property (`TreeView`)
- `.Expanded` property (`TreeView.Nodes`)

>Note:
>
>-   When `.BackStyle` is `fmBackStyleOpaque`, the control’s own `.BackColor` is used directly.
>-   When `.BackStyle` is `fmBackStyleTransparent`:
>        -   For controls that support transparency in WinForms (e.g., `Label`, `CommandButton`, `CheckBox`, `OptionButton`, `Image`, etc), `.BackColor` is set to `"Transparent"`.
>        -   For controls that **do not support transparency in WinForms** (`TextBox`, `ComboBox`, `ToggleButton`), the background is substituted:
>            -   If the parent control has a `.BackColor`, that color is used.
>            -   If the parent is a `Page` (which does not expose `.BackColor`), a system default color (`&H8000000F&`) is used as a fallback, which matches the visual background color of the `Page`.
>
>-   `.PictureSizeMode`/`.PictureAlignment` is mapped to the corresponding WinForms `.SizeMode` behavior:
>        -   `fmPictureSizeModeClip` → `"Normal"` or `"CenterImage"` (depends on `.PictureAlignment`)
>        -   `fmPictureSizeModeStretch` → `"StretchImage"`
>        -   `fmPictureSizeModeZoom` → `"Zoom"`
>
>        -   When `.PictureSizeMode = fmPictureSizeModeClip`:
>            -   `.PictureAlignment = fmPictureAlignmentCenter` → `"CenterImage"`
>            -   `.PictureAlignment = fmPictureAlignmentTopLeft` → `"Normal"`
>            -   Other alignment values are not supported in WinForms `PictureBox`, and are converted to `"Normal"` (top-left).
>        -   For `"StretchImage"` and `"Zoom"` modes, `.PictureAlignment` is ignored.
>-   MultiPage controls with `.TabOrientation` set to `fmTabOrientationLeft`  or  `fmTabOrientationRight` render tab text vertically in WinForms, unlike VBA, which keeps it horizontal.

## Supported Controls
| VBA Form Class | WinForms Class|
| ------ | ------ |
| `Label` | `Label` |
| `CommandButton` | `Button` |
| `Frame` (without Caption) | `Panel` |
| `Frame` (with any Caption) | `GroupBox` |
| `TextBox` | `TextBox` |
| `SpinButton` | `NumericUpDown` |
| `ListBox` | `ListBox` |
| `CheckBox` | `CheckBox` |
| `ToggleButton` | `CheckBox`<br>(`.Appearance = "Button"`) |
| `OptionButton` | `RadioButton` |
| `Image` | `PictureBox` |
| `ScrollBar` | `HScrollBar` / `VScrollBar` |
| `ComboBox` | `ComboBox` |
| `MultiPage` | `TabControl` |
| `ListView`(`.View=lvwReport`) | `ListView` |
| `TreeView` | `TreeView` |

> Note:
>- `SpinButton` behaves differently in VBA and WinForms, so appearance may vary depending on placement.<br>

# HTML

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Z-order of controls
- Control colors (foreground)
- Control colors (background) (Excluding `MultiPage`, `ScrollBar`)
- Text display (`UserForm`, `Frame`, `Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`UserForm`, `Frame`, `TextBox`, `Label`, `ListBox`, `ComboBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox`, `ComboBox`, `CheckBox`, `ToggleButton`, `OptionButton`, `ListBox`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`, `ListView`, `TreeView`
- Selection state of `OptionButton`, `CheckBox`, `ToggleButton`
- Transparent background setting specified in `.BackStyle`(`Label`, `TextBox`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `Image`, `ComboBox`)
- Images Embedded in Controls (`Image`)
- `.Visible` property
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`)
- `.PasswordChar` property (`TextBox` [.MultiLine=False])
- `.ScrollBars` property (`TextBox` [.MultiLine=True])
- `.WordWrap` property (`TextBox` [.MultiLine=True])
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`)
- `.PictureAlignment` property (`Image`)
- `.Scroll` property (`TreeView`)
- `.Expanded` property (`TreeView.Nodes`)

> Notes:
>
>-   For `ListBox`, `.MultiSelect` values `fmMultiSelectMulti` and `fmMultiSelectExtended` are both converted to HTML multiple-selection mode (`multiple`). After conversion, selecting multiple items follows standard browser behavior and typically requires holding `Ctrl` (or the platform equivalent modifier key).
>-   For `TextBox`, `.PasswordChar` is reflected only by converting the control to a password input (`type="password"`); the actual masking character itself is not preserved.
>-   For `MultiPage`, `.TabOrientation` supports `fmTabOrientationTop` and `fmTabOrientationBottom` only. Any other value is treated as `fmTabOrientationTop`.


## Supported Controls
| VBA Form Class | CSS Element|
| ------ | ------ |
| `Label` | `<div>` |
| `CommandButton` | `<button>` |
| `Frame` | `<fieldset>` + `<legend>` |
| `TextBox` (`.MultiLine=False`) | `<input>` |
| `TextBox` (`.MultiLine=True`) | `<textarea>` |
| `SpinButton` | `<input type="number">` |
| `ListBox` | `<select size=[n]>` |
| `CheckBox` | `<label>` + `<input type="checkbox">` |
| `ToggleButton` | `<input type="checkbox">` (`display:none`) + `<label>` |
| `OptionButton` | `<label>` + `<input type="radio">` |
| `Image` | `<div>` |
| `ScrollBar` | `<input type="range">` |
| `ComboBox` (`.Style = fmStyleDropDownCombo`) | `<input list="[datalist name]">` |
| `ComboBox` (`.Style = fmStyleDropDownList`) | `<select>` |
| `MultiPage` | `<div>` (Tabs) + `<div>` (Pages) + `<button>` (Scroll Buttons)|
| `ListView` | `<div>` + `<table>` |
| `TreeView` | `<div>` + `<details>` + `<summary>`|

> Note:
>
>`SpinButton` behaves differently in VBA and CSS, so appearance may vary depending on placement.<br>