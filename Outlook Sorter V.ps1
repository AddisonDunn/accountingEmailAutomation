#Function to look through excel file and turn contents of first column into list 
$Excel = New-Object -ComObject Excel.Application 
$Excel.Visible = $true
$Excel.DisplayAlerts = $false
$ExcelWorkBook = $Excel.Workbooks.Open("C:\Users\michael.vabner\Documents\Book2.xlsx") 
$ExcelWorkSheet = $Excel.Sheets.item("Sheet1") 
$ExcelWorkSheet.activate() 
$arrBlackListEmails = @()
$i = 1
$string = $ExcelWorkSheet.Cells.Item($i, 1).Value()
Do 
{
    $arrBlackListEmails += $ExcelWorkSheet.Cells.Item($i, 1).Value()
    $i = $i + 1
}
Until ($ExcelWorkSheet.Cells.Item($i, 1).Value() -eq $null)



#How to create a folder based on name after "@" and before the ".com"

$testemail = "bill@microsoft.com"
$testemail = $testemail -replace ".*@" -replace ".com.*"
New-Item C:\$testemail -type directory 
