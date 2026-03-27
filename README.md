# Event1644Reader

A C++ Win32 port of the original [Event1644Reader.ps1 v1.04](https://support.microsoft.com/en-us/kb/3060643) PowerShell script by Ming Chen. Scans Windows `.evtx` event log files for **Event ID 1644** (Active Directory LDAP search events), extracts key fields to CSV, and optionally imports the results into Excel with pre-built pivot tables for analysis.

## What is Event 1644?

Event 1644 is logged by Active Directory Domain Services when expensive, inefficient, or high-volume LDAP searches are detected. Enabling and analyzing these events helps administrators identify problematic LDAP queries impacting domain controller performance. See [KB3060643](https://support.microsoft.com/en-us/kb/3060643) for a detailed walkthrough.

## Features

- Scans all `.evtx` files in a specified directory for Event 1644
- Extracts 19 fields per event including extended fields from KB2800945+:
  - LDAPServer, TimeGenerated, ClientIP, ClientPort, StartingNode, Filter, SearchScope, AttributeSelection, ServerControls, VisitedEntries, ReturnedEntries
  - UsedIndexes, PagesReferenced, PagesReadFromDisk, PagesPreReadFromDisk, CleanPagesModified, DirtyPagesModified, SearchTimeMS, AttributesPreventingOptimization
- Exports results to `1644-<logname>.csv` files (UTF-8 encoded)
- Imports CSVs into Excel via COM automation and generates 6 pivot table sheets:
  - **2.TopIP-StartingNode** — grouped by StartingNode, Filter, ClientIP
  - **3.TopIP** — grouped by ClientIP, Filter
  - **4.TopIP-Filters** — grouped by Filter, ClientIP
  - **5.TopTime-IP** — total/average search time by ClientIP
  - **6.TopTime-Filters** — total/average search time by Filter
  - **7.TimeRanks** — search time distribution ranking
  - **8.SandBox** — empty worksheet for custom analysis
- Raw data sheet (1.RawData) with auto-filter and frozen header row

## Requirements

- **OS:** Windows (uses Win32 APIs: `wevtapi`, `shlwapi`, `oleauto`)
- **Compiler:** MSVC (`cl.exe`) or MinGW g++ with Windows SDK headers
- **Excel** (optional): Excel 2013 or later for pivot table generation. 64-bit Excel is recommended for larger datasets.

## Building

### With MSVC (Visual Studio Developer Command Prompt)

```cmd
cl /EHsc /W4 /DUNICODE /D_UNICODE Event1644Reader.cpp ole32.lib oleaut32.lib wevtapi.lib shlwapi.lib
```

### With g++ (MinGW/MSYS2 UCRT64)

```cmd
g++ -DUNICODE -D_UNICODE -o Event1644Reader.exe Event1644Reader.cpp -lole32 -loleaut32 -lwevtapi -lshlwapi
```

### With VS Code

Use the included build task (`Ctrl+Shift+B`) which invokes g++ to compile the active file.

## Usage

1. Run `Event1644Reader.exe`
2. Enter the path to the directory containing `.evtx` files (or press Enter to use the program's directory)
3. The tool scans each `.evtx` file and generates `1644-*.csv` files for any logs containing Event 1644
4. If Excel is installed, CSVs are automatically imported into a new workbook with pivot tables
5. Enter a filename to save the Excel report
6. Choose whether to delete the intermediate CSV files

```
Event1644Reader: See https://support.microsoft.com/en-us/kb/3060643 for sample walk through and pivotTable tips.
Enter local, mapped or UNC path to Evtx(s). Remove trailing blank.
For Example (c:\CaseData)
Or press [Enter] if evtx is in the program folder.
> c:\CaseData
Reading DirectoryService.evtx
    Event 1644 found (243 events), generated 1644-DirectoryService.csv
Import csv to excel.
Customizing XLS.
Enter a FileName to save extracted event 1644 xlsx:
> MyReport
Saving file to c:\CaseData\MyReport.xlsx
Delete generated 1644-*.csv? ([Enter]/[Y] to delete, [N] to keep csv)
>
Script completed.
```

## Notes

- Pre-Windows Server 2008 `.evt` files must be converted to `.evtx` format first using a later OS version
- Pre-2008 logs may not contain all 16 extended data fields; some pivot table columns may be empty
- The original PowerShell script (`Event1644Reader.ps1`) is also included for reference

## References

- [KB3060643 — How to turn on LDAP Event Logging](https://support.microsoft.com/en-us/kb/3060643)
- [KB2800945 — Extended Event 1644 fields](https://support.microsoft.com/en-us/kb/2800945)
