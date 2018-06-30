#usage     > PS ..\PS1\scripts_ps> .\ssps.ps1 -file "..\xls\cstmxl.xlsx" -sheet "CustomSheet" -row 2 -col 9
#output    > 0000000123
param(
        [string]$file,
        [string]$sheetname,
        [int]$row,
        [int]$col
        )
    $objExcel=New-Object -ComObject Excel.Application;$objExcel.Visible=$false;$objExcel.DisplayAlerts=$false
    Write-Host($objExcel.Workbooks.Open($file).Sheets.Item($sheetname).Cells.Item($row,$col).text)
    $objExcel.quit();
