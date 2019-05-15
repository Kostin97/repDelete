$filePath ="C:\Users\ïîçèðîí\Desktop\config.xlsx"
$Excel = New-Object -COM "Excel.Application"
$Excel.Visible = $true
$WorkBooks = $Excel.Workbooks.Open($filePath )
$WorkSheet = $WorkBooks.Sheets.Item(1)
 
$column = 1
$columnDay = 2
 

$b1, $b2, $b3, $b4, $b5, $b6, $b7, $b8, $b9, $b10  = 1..11 | % {$WorkSheet.Cells.Item($_,$column).value2}
$a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10  = 1..11 | % {$WorkSheet.Cells.Item($_,$columnDay).Value2}
$excel.ActiveWorkbook.close()
$excel.Quit()

$collectionPath = $b1, $b2, $b3, $b4, $b5, $b6, $b7, $b8, $b9, $b10
$collectionDay = $a1, $a2, $a3, $a4, $a5, $a6, $a7, $a8, $a9, $a10


foreach ($item in $collectionPath)
{
$index = $collectionPath.IndexOf($item)
if($item -ne $null){
 $item
  Get-ChildItem -Path $item -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $(Get-Date).AddDays(-$collectionDay.GetValue($index)) } | Remove-Item –Force
}
}

