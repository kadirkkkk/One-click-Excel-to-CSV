@echo off
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\ExcelToTab" /f
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsm\shell\ExcelToTab" /f
echo Uninstalled successfully.
pause