Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select Excel files"
$dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
$dialog.Multiselect = $true

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    exit
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($file in $dialog.FileNames) {
    $workbook = $null

    try {
        $folder = Split-Path $file

        # Create Uploads folder if it doesn't exist
        $uploadFolder = Join-Path $folder "Uploads"
        if (-not (Test-Path $uploadFolder)) {
            New-Item -ItemType Directory -Path $uploadFolder | Out-Null
        }

        # Remove spaces and hyphens from filename
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $cleanName = $name -replace '[ ()-]', ''

        $csvPath = Join-Path $uploadFolder "$cleanName.csv"

        # Open workbook
        $workbook = $excel.Workbooks.Open($file)

        # Save first worksheet as CSV
        $workbook.Worksheets.Item(1).SaveAs($csvPath, 6)  # 6 = xlCSV

        $workbook.Close($false)
        $workbook = $null

        Write-Host "Converted:"
        Write-Host "  Source: $file"
        Write-Host "  Output: $csvPath"
        Write-Host ""
    }
    catch {
        Write-Warning "Failed to convert: $file"
        Write-Warning $_.Exception.Message

        if ($workbook -ne $null) {
            $workbook.Close($false)
        }
    }
}

$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "All selected files have been processed."
