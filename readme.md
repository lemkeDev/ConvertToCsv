# Excel to CSV Converter

A PowerShell script that batch converts selected Excel (`.xls` and `.xlsx`) files to CSV.

## Features

* Select multiple Excel files with a file picker.
* Converts the **first worksheet** of each workbook.
* Processes files concurrently (up to **4 Excel instances** by default).
* Creates an **Uploads** folder in the source directory and saves all CSVs there.
* Cleans output filenames by removing spaces, hyphens (`-`), and parentheses (`(` `)`).

## Requirements

* Windows
* Microsoft Excel
* PowerShell 5.1+

## Usage

Run:

```powershell
.\ExcelToCsv.ps1
```

Select your Excel files, and the script will create the following structure:

```text
Source Folder/
├── File1.xlsx
├── File2.xls
└── Uploads/
    ├── File1.csv
    └── File2.csv
```

The number of concurrent conversions can be changed by modifying:

```powershell
$maxConcurrentJobs = 4
```
