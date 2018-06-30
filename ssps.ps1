GET-XLS-Cell-value -file ".\cstmxl.xlsx" -sheet "CustomSheet" -row 2 -col 9

Function GET-XLS-Cell-value()
    {
    param(
        [string]$file,
        [string]$sheetname,
        [int]$row,
        [int]$col
        )
    $objExcel=New-Object -ComObject Excel.Application;$objExcel.Visible=$false;$objExcel.DisplayAlerts=$false
    Write-Host($objExcel.Workbooks.Open($file).Sheets.Item($sheetname).Cells.Item($row,$col).text)
    $objExcel.quit();  
    }         