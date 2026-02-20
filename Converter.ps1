param([string[]]$files)

# 1. Setup Logging
Write-Host "--- Excel Converter Starting ---" -ForegroundColor Cyan
$configPath = Join-Path $PSScriptRoot "config.json"

if (-not (Test-Path $configPath)) {
    Write-Host "ERROR: config.json not found at $configPath" -ForegroundColor Red
    pause; exit
}

$config = Get-Content $configPath | ConvertFrom-Json
$results = New-Object System.Collections.Generic.List[PSObject]

# 2. Start Excel
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
} catch {
    Write-Host "ERROR: Could not start Excel COM object." -ForegroundColor Red
    pause; exit
}

foreach ($filePath in $files) {
    Write-Host "Target: $filePath"
    
    try {
        # Resolve path and ensure file is local (force download if OneDrive)
        $item = Get-Item -LiteralPath $filePath
        $baseName = $item.BaseName
        $folder = $item.DirectoryName
        
        # Create Local Temp copy
        $tempPath = Join-Path $env:TEMP "$($baseName)_$([guid]::NewGuid().Guid).xlsx"
        Copy-Item -LiteralPath $item.FullName -Destination $tempPath -Force
        
        Write-Host "Opening file..." -NoNewline
        $workbook = $excel.Workbooks.Open($tempPath)
        Write-Host " [Done]" -ForegroundColor Green
        
        $targetSheetName = $config.pageName.Trim()
        $sheet = $workbook.Sheets | Where-Object { $_.Name.Trim() -ieq $targetSheetName }

        if ($null -eq $sheet) {
            throw "Sheet '$targetSheetName' not found."
        }

        # Handle Output Naming
        $targetExt = $config.extension.TrimStart('.')
        $newFileName = "$baseName.$targetExt"
        $counter = 1
        while (Test-Path (Join-Path $folder $newFileName)) {
            $newFileName = "$baseName ($counter).$targetExt"
            $counter++
        }
        $outputPath = Join-Path $folder $newFileName

        # Convert
        Write-Host "Saving to: $newFileName..."
        $usedRange = $sheet.UsedRange
        $data = $usedRange.Value2
        
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $writer = New-Object System.IO.StreamWriter($outputPath, $false, $utf8NoBom)
        
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = for ($c = 1; $c -le $colCount; $c++) {
                $val = if ($rowCount -eq 1 -and $colCount -eq 1) { $data } else { $data[$r, $c] }
                if ($null -eq $val) { "" } else { $val.ToString() }
            }
            $writer.WriteLine(($line -join $config.delimiter))
        }
        $writer.Close()
        
        $results.Add([PSCustomObject]@{ File = $item.Name; Status = "OK" })
        $workbook.Close($false)

    } catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $results.Add([PSCustomObject]@{ File = $filePath; Status = "Error: $($_.Exception.Message)" })
    } finally {
        if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
    }
}

# 3. Cleanup
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "--- Process Finished ---" -ForegroundColor Cyan

# 4. Final Summary
$msg = $results | Out-String
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($msg, "Conversion Summary")

# Keep window open if there were errors
if ($results | Where-Object { $_.Status -ne "OK" }) {
    Write-Host "Press any key to close this window..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}