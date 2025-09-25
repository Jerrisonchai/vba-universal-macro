# VBA Universal Macro 🚀

A reusable collection of **universal Excel VBA macros** that can be integrated into any VBA project to enhance productivity, logging, performance, and user interaction.

This repository is designed to serve as a **utility toolkit** for business analysts, developers, and power users who want to quickly add features like logging, performance optimization, shape interactions, and workbook automation into their Excel projects.

---

## ✨ Features

- **Execution Logging**
  - `capturetime` and `captureendtime` track macro start/end with timestamps, user info, and workbook details.
- **Dynamic Button Interaction**
  - Shapes automatically toggle colors when clicked (`MyShape_Click`, `MyFont_Click`).
- **Performance Optimization**
  - Enable or disable Excel settings (`ScreenUpdating`, `Events`, etc.) for faster macro runs.
- **Sheet Protection**
  - Protect/unprotect all sheets in a workbook with a single macro.
- **Helper Functions**
  - Convert column numbers to letters (`Col_Letter`).
  - Check the status of Excel events/updates (`CheckEvents`).
- **Workbook Automation**
  - Automatically logs session start/end and resets dashboard button styles when the workbook is opened.

---

## 📂 Macro List

### Logging
- `capturetime` → Logs macro start.
- `captureendtime` → Logs macro end.

### Shape Interaction
- `MyShape_Click` → Toggle shape fill color.
- `MyFont_Click` → Toggle shape font color.

### Performance
- `OptimizedMode` → Enable/disable optimization.
- `OptOn` / `OptOff` → Quick shortcuts.
- `CheckEvents` → Show optimization status.

### Protection
- `ProtectAllSheets` → Protects all sheets with one command.

### Utility
- `Col_Letter` → Converts column numbers to Excel column letters.

---

## ⚡ Workbook-Level Integration

Add the following code to your **`ThisWorkbook` module** to enhance workbook behavior:

```vba
Option Explicit
Private Sub Workbook_Open()
    Call capturetime
    
    Dim sourceWorkbook As Workbook
    Dim sourceSheet1 As Worksheet
    Dim shape1 As Shape
    
    Set sourceWorkbook = ThisWorkbook
    Set sourceSheet1 = sourceWorkbook.Sheets("Dashboard")
    
    OptimizedMode False
    
    For Each shape1 In sourceSheet1.Shapes
        On Error Resume Next
        shape1.Fill.ForeColor.RGB = RGB(0, 0, 255)
        shape1.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    Next
    
    sourceSheet1.Activate
    
    Set sourceWorkbook = Nothing
    Set sourceSheet1 = Nothing
    Set shape1 = Nothing
    
    Call captureendtime
    
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub
```
## 🔧 What it does:
- Logs workbook open/start and close/end.
- Resets dashboard buttons to default colors.
- Activates the Dashboard sheet on open.
- Ensures a clean, optimized start for your workbook.

## 🛠 Setup Instructions
1. Open your workbook in Excel VBA Editor (ALT + F11).
2. Create a new module → Paste the Universal Macro code.
3. In ThisWorkbook module → Paste the Workbook_Open code.
4. Add a worksheet named LOG (used for execution logs).
5. (Optional) Add a sheet named Dashboard if you want buttons and shapes to reset automatically.
6. Update the password in ProtectAllSheets and logging macros (admin by default).

## 📊 Example LOG Output
| Macro Name  | Timestamp           | User | Status            | Workbook           |
| ----------- | ------------------- | ---- | ----------------- | ------------------ |
| MyMacroName | 2025-09-19 10:15 AM | Jerr | Run Macro Started | FinanceReport.xlsm |
| MyMacroName | 2025-09-19 10:16 AM | Jerr | Run Macro Ended   | FinanceReport.xlsm |

## 🔒 Security Notes
- Default sheet protection password is set to admin.
- Change this in code before deploying to production.

## 💡 Use Cases
- Corporate Excel dashboards.
- Audit logging for compliance.
- Performance-heavy reporting macros.
- Interactive dashboards with shapes/buttons.
- General VBA utility functions for any project.

## 📌 Next Steps
- Add more universal macros (e.g., error logging, API connectors).
- Convert this into a .bas file for easier import/export.
- Extend to Excel Add-In (.xlam) format for enterprise distribution.

## 🤝 Contribution
- Feel free to fork, improve, and submit PRs!
- Let’s build a universal VBA toolkit for analysts and developers worldwide 🌍.
