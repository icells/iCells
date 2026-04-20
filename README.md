# iCells - Excel Enhancement Add-in

**iCells** is a powerful VSTO (Visual Studio Tools for Office) add-in for Microsoft Excel, providing 200+ productivity features including data processing, text extraction, random data generation, VBA security tools, AI integration, and more.

- **Supported Excel**: 2013 and later
- **Framework**: .NET Framework 4.8
- **Languages**: Simplified Chinese / English

> [**中文文档**](README_CN.md)

---

## Table of Contents

- [Feature Demos](#feature-demos)
  - [File Directory Browser](#file-directory-browser)
  - [Reading Mode (Spotlight)](#reading-mode-spotlight)
  - [Copy to Visible Cells](#copy-to-visible-cells)
  - [Insert Sequence](#insert-sequence)
  - [Data Extraction & Undo](#data-extraction--undo)
  - [Chinese Pinyin Conversion](#chinese-pinyin-conversion)
  - [Random Data Generation](#random-data-generation)
  - [Random Cell Selection](#random-cell-selection)
  - [Input Restriction](#input-restriction)
  - [Show Hidden Rows & Columns](#show-hidden-rows--columns)
  - [Unlock VBA Project Password](#unlock-vba-project-password)
  - [VBE Theme Management](#vbe-theme-management)
  - [VBA Module Visibility](#vba-module-visibility)
  - [Theme Switching](#theme-switching)
  - [i18n - Chinese & English](#i18n---chinese--english)
- [Full Feature List](#full-feature-list)
- [Community](#community)

---

## Feature Demos

### File Directory Browser

Browse files and folders directly in Excel. Navigate directories, extract file paths and names, batch rename files, and batch create folders from cell contents.

![File Directory](demos/file-directory.gif)

---

### Reading Mode (Spotlight)

Highlight the current row/column with a customizable spotlight overlay. Supports adjustable transparency, multiple color schemes, and a fixed spotlight toggle for focused data review.

![Reading Mode](demos/reading-mode.gif)

---

### Copy to Visible Cells

Copy data and paste it only into visible (non-hidden) cells, skipping filtered or hidden rows. Essential for working with filtered datasets.

![Copy to Visible Cells](demos/copy-to-visible-cells.gif)

---

### Insert Sequence

Insert various types of sequences into cells with one click: numeric (1,2,3...), circled numbers, Roman numerals, Chinese numerals, Heavenly Stems, Earthly Branches, English letters, months, weekdays, and more. Supports undo.

![Insert Sequence](demos/insert-sequence.gif)

---

### Data Extraction & Undo

Extract specific data from cells: numbers, letters, Chinese characters, phone numbers, emails, ID numbers, bank card numbers, postal codes, and more. All extraction operations support undo.

![Data Extraction & Undo](demos/extract-and-undo.gif)

---

### Chinese Pinyin Conversion

Convert Chinese characters to Pinyin (romanization), Pinyin with tone marks, Pinyin initials, or stroke counts.

![Chinese Pinyin](demos/chinese-pinyin.gif)

---

### Random Data Generation

Generate random integers, decimals, letters, Chinese names, Chinese characters, and even random math problems. Fully configurable range and precision settings.

![Random Data](demos/random-data.gif)

---

### Random Cell Selection

Randomly select cells from a given range - useful for random sampling, lottery-style picks, and audit spot checks.

![Random Cell Selection](demos/random-cell-selection.gif)

---

### Input Restriction

Apply data validation rules to cells: restrict to numbers only, text only, email format, phone format, IP address format, yes/no, male/female, no duplicates, and more. One-click removal of all restrictions.

![Input Restriction](demos/input-restriction.gif)

---

### Show Hidden Rows & Columns

Quickly show/hide all hidden rows and columns, toggle "very hidden" sheets, and manage worksheet visibility with one click.

![Show Hidden Rows & Columns](demos/show-hidden-rows-cols.gif)

---

### Unlock VBA Project Password

Remove VBA project passwords from locked workbooks, allowing access to protected VBA code. Also supports making VBA projects invisible/not viewable and repairing visibility.

![Unlock VBA Password](demos/unlock-vba-password.gif)

---

### VBE Theme Management

Customize the Visual Basic Editor (VBE) color theme - apply dark mode, custom syntax highlighting, and personalized editor appearance.

![VBE Theme Management](demos/vbe-theme-management.gif)

---

### VBA Module Visibility

Show or hide specific VBA modules in the VBE project explorer for better code organization and project management.

![VBA Module Visibility](demos/vba-module-visibility.gif)

---

### Theme Switching

Switch between multiple UI themes: Light, Dark, Green, Deep Blue, or create your own custom theme with the built-in color picker.

![Theme Switching](demos/theme-switching.gif)

---

### i18n - Chinese & English

Full bilingual support with dynamic language switching between Simplified Chinese and English for all UI elements.

![i18n Support](demos/i18n-support.gif)

---

## Full Feature List

### Data Processing
| Feature | Description |
|---------|-------------|
| Text Extraction | Extract numbers, letters, Chinese characters, phone numbers, emails, ID numbers, bank cards, postal codes |
| Data Format | Set text/date format, convert currency units, reverse data order |
| Text Cleanup | Remove spaces, line breaks, non-printing characters, formulas |
| Merge & Split | Merge cells by content, split merged cells and fill, combine worksheets/workbooks |
| Copy to Visible Cells | Paste only into visible (non-hidden) cells |

### Number & Code Validation
| Feature | Description |
|---------|-------------|
| ID Number | Validate, extract birthday/age/gender/zodiac/constellation |
| Bank Card | Validate card numbers and credit reporting codes |
| Phone Number | Validate, format with spaces/hyphens, mask middle digits |

### Chinese Text
| Feature | Description |
|---------|-------------|
| Pinyin | Full pinyin, pinyin with tones, pinyin initials |
| Stroke Count | Get stroke count for Chinese characters |

### Random Data
| Feature | Description |
|---------|-------------|
| Random Generation | Integers, decimals, letters, Chinese names, Chinese characters, math problems |
| Random Selection | Randomly pick cells from a range |

### Sequences & Insertion
| Feature | Description |
|---------|-------------|
| Sequences | Numeric, circled, Roman, Chinese numerals, Heavenly Stems, Earthly Branches, letters, months, weekdays |
| Special Insert | Dropdown lists, alternating rows, symbols, checkboxes, radio buttons |

### Input Restriction
| Feature | Description |
|---------|-------------|
| Validation Rules | Numbers only, text only, email, phone, IP address, no duplicates, yes/no, male/female |

### Financial
| Feature | Description |
|---------|-------------|
| Currency | Uppercase RMB, Chinese numeral style, abbreviations |
| Rounding | Standard, banker's, round up/down, custom decimal places |

### Image & PDF
| Feature | Description |
|---------|-------------|
| Export | Save selection as image/PDF |
| Image Tools | Compress, fit to cell, resize to standard photo dimensions |
| PDF Tools | Convert images to PDF, merge PDFs, extract images from PDF |
| QR Code | Generate QR codes (with/without logo), barcodes, QR code reader |
| OCR | Recognize and extract text from images |

### Visual Aids
| Feature | Description |
|---------|-------------|
| Spotlight | Reading mode with adjustable color and transparency |
| Arrow Indicators | Highlight with arrow overlays |
| Navigation | Navigation pane, show/hide sheets, show mouse position |
| Snapshots | Save and restore worksheet views |

### Security
| Feature | Description |
|---------|-------------|
| Protection | Remove worksheet/workbook protection (single or batch) |
| VBA Tools | Unlock VBA password, hide/show modules, VBE themes |
| Privacy | Clear document properties and private information |

### Database & SQL
| Feature | Description |
|---------|-------------|
| ACE/SQL | Execute SQL queries, batch SQL, multi-database operations |
| SQLite | SQLite query, batch execution, multi-database support |

### AI Integration
| Feature | Description |
|---------|-------------|
| AI Chat | Built-in AI chat panel with multi-provider support |
| Providers | OpenAI, Claude, Gemini, DeepSeek, Qwen, Zhipu, SiliconFlow, Ollama, ONNX, Custom |
| Code Execution | Run VB.NET, VBA, and C# code directly from Excel |

### File & Directory
| Feature | Description |
|---------|-------------|
| File Browser | Navigate directories, extract file paths/names |
| Batch Operations | Batch rename files/folders, batch create folders |
| Auto Backup | Configurable automatic workbook backup |

### UI & Localization
| Feature | Description |
|---------|-------------|
| Themes | Light, Dark, Green, Deep Blue, Custom |
| Languages | Simplified Chinese, English |

---

## Community

Join the iCells QQ group for discussion, feedback, and support:

**QQ Group: 17576056**

<p align="center">
  <img src="demos/qq-group.jpg" alt="iCells QQ Group" width="300">
</p>

---

## License

Copyright &copy; 2020-2025 iCells Developer. All rights reserved.
