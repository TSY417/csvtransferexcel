$csvFolder = "C:\MergeData\2023\transfer"
$excelFolder = "C:\MergeData\2023\excel"

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

foreach ($csvFile in Get-ChildItem -Path $csvFolder -Filter *.csv) {
    $csvFilePath = $csvFile.FullName
    $excelFilePath = Join-Path $excelFolder ($csvFile.BaseName + ".xlsx")

    $workbook = $excel.Workbooks.Open($csvFilePath)
    $workbook.SaveAs($excelFilePath, 51)  # 51代表Excel的XLSX格式
    $workbook.Close()
}

$excel.Quit()
