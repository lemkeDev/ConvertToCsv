Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select Excel files"
$dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
$dialog.Multiselect = $true

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    exit
}

$files = $dialog.FileNames
$maxConcurrentJobs = 4

$scriptBlock = {
    param($file)

    $excel = $null
    $workbook = $null

    try {
        $folder = Split-Path $file

        $uploadFolder = Join-Path $folder "Uploads"
        if (-not (Test-Path $uploadFolder)) {
            New-Item -Path $uploadFolder -ItemType Directory -Force | Out-Null
        }

        $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $cleanName = $name -replace '[ ()-]', ''
        $csvPath = Join-Path $uploadFolder "$cleanName.csv"

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($file)
        $workbook.Worksheets.Item(1).SaveAs($csvPath, 6)

        $workbook.Close($false)
        $workbook = $null

        $excel.Quit()
        $excel = $null

        Write-Output "SUCCESS: $cleanName.csv"
    }
    catch {
        Write-Output "FAILED: $file"
        Write-Output $_.Exception.Message
    }
    finally {
        if ($workbook -ne $null) {
            try { $workbook.Close($false) } catch {}
        }

        if ($excel -ne $null) {
            try { $excel.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$jobs = @()

foreach ($file in $files) {
    while (($jobs | Where-Object State -eq Running).Count -ge $maxConcurrentJobs) {
        Start-Sleep -Milliseconds 500

        $finished = $jobs | Where-Object State -ne Running
        foreach ($job in $finished) {
            Receive-Job $job
            Remove-Job $job
            $jobs = $jobs | Where-Object Id -ne $job.Id
        }
    }

    $jobs += Start-Job -ScriptBlock $scriptBlock -ArgumentList $file
}

while ($jobs.Count -gt 0) {
    Start-Sleep -Milliseconds 500

    $finished = $jobs | Where-Object State -ne Running
    foreach ($job in $finished) {
        Receive-Job $job
        Remove-Job $job
        $jobs = $jobs | Where-Object Id -ne $job.Id
    }
}

Write-Host ""
Write-Host "All selected files have been processed."