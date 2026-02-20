param([string[]]$files)

$config = Get-Content "$PSScriptRoot\config.json" | ConvertFrom-Json
$results = New-Object System.Collections.Generic.List[PSObject]

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($filePath in $files) {
    $fullPath = Resolve-Path $filePath
    $folder = Split-Path $fullPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    $reportName = [System.IO.Path]::GetFileName($fullPath)

    try {
        $workbook = $excel.Workbooks.Open($fullPath)
        $sheet = $workbook.Sheets | Where-Object { $_.Name -eq $config.pageName }

        if ($null -eq $sheet) {
            throw "Sheet '$($config.pageName)' not found."
        }

        # Handle naming (1), (2), etc.
        $targetExt = $config.extension.TrimStart('.')
        $newFileName = "$baseName.$targetExt"
        $counter = 1
        while (Test-Path "$folder\$newFileName") {
            $newFileName = "$baseName ($counter).$targetExt"
            $counter++
        }
        $outputPath = "$folder\$newFileName"

        # Export Logic
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        $data = $usedRange.Value2

        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $writer = New-Object System.IO.StreamWriter($outputPath, $false, $utf8NoBom)

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = for ($c = 1; $c -le $colCount; $c++) {
                $val = $data[$r, $c]
                if ($null -eq $val) { "" } else { $val.ToString() }
            }
            $writer.WriteLine(($line -join $config.delimiter))
        }
        $writer.Close()
        
        $workbook.Close($false)
        $results.Add([PSCustomObject]@{ File = $reportName; Status = "OK" })
    }
    catch {
        $results.Add([PSCustomObject]@{ File = $reportName; Status = "Error: $($_.Exception.Message)" })
        if ($workbook) { $workbook.Close($false) }
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Final Dialog
$msg = $results | Out-String
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($msg, "Conversion Summary")