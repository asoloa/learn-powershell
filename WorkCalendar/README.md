# Excel Timesheet → Calendar Report Generator  
_A modern, automated PowerShell solution for transforming timesheet data into a calendar-style Excel report_

This project converts raw employee timesheet data into a fully formatted **calendar report**.  
It leverages PowerShell and Excel COM automation for fast, in-memory data processing and clean, professional output.

---

## ⚡ Features

### **Automated Excel Processing**
- Reads an input Excel timesheet and creates a normalized **Data Sheet**  
- Generates a **Calendar Sheet** with employee/project rows by day  
- Automatically calculates the month and days from the latest date in the dataset  

### **Smart Data Handling**
- Uses a 2D array for fast in-memory operations  
- Automatically fills empty cells using previous-column values  
- Supports leave types (can be easily extended):
  - **VL** → Vacation Leave  
  - **SL** → Sick Leave  
  - **UNK** → Unknown/other leave  

### **Professional Formatting**
- Employee headers bolded with indentation  
- Projects indented below employees  
- Borders applied across the full calendar  
- Weekends shaded automatically  
- Conditional formatting applied for leave types  
- Both Calendar and Data sheets are protected

### **PowerShell Best Practices**
- Modular functions with clear responsibilities: `Get-EmployeeData`, `Populate-Calendar`, `Set-LeaveFormatting`, `Shade-Weekends`  
- Efficient in-memory string and array handling  
- Clean separation of **data processing**, **calendar construction**, and **formatting**  
- COM objects properly released to avoid memory leaks  

---

## ▶️ Usage & Testing

1. Download the zipped `.xls` test data <a href="https://bit.ly/asoloa-wc-timesheet-data" target="_blank">here</a>.
2. Launch PowerShell (ExecutionPolicy must allow scripts)  
3. Run:
```powershell
.\WorkCalendar.ps1
# OR
powershell -ExecutionPolicy Bypass .\WorkCalendar.ps1
```
4. Select the Excel timesheet file when prompted
5. Output is saved to: `<script-directory>/extracted/yyyyMMdd-HHmmss.xlsx`

---

_This tool was developed and tested in PowerShell 5.1._