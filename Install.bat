@echo off
SET "scriptPath=%~dp0Converter.ps1"
SET "escapedPath=%scriptPath:\=\\%"

:: Update .xlsx
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\ExcelToTab" /ve /t REG_SZ /d "Excel to Tab delimited" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\ExcelToTab\command" /ve /t REG_SZ /d "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%escapedPath%\" \"%%1\"" /f

:: Update .xlsm
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsm\shell\ExcelToTab" /ve /t REG_SZ /d "Excel to Tab delimited" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsm\shell\ExcelToTab\command" /ve /t REG_SZ /d "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%escapedPath%\" \"%%1\"" /f

echo Registry Updated Successfully!
pause