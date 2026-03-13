$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open('d:\문서\MyClaudeCode\실행예산서\★실행결의서_남부내륙철도(김천~거제) 건설사업 제2공구 노반신설 기타공사.xlsb')
$sheet = $workbook.Sheets.Item('토목실행')
Write-Host '--- Row 4 contents ---'
for ($i = 1; $i -le 30; $i++) {
    $val = $sheet.Cells.Item(4, $i).Text
    $formula = $sheet.Cells.Item(4, $i).Formula
    Write-Host "Col $i : text='$val', formula='$formula'"
}
Write-Host '--- Row 100 contents ---'
for ($i = 1; $i -le 30; $i++) {
    $val = $sheet.Cells.Item(100, $i).Text
    $formula = $sheet.Cells.Item(100, $i).Formula
    Write-Host "Col $i : text='$val', formula='$formula'"
}
$workbook.Close($false)
$excel.Quit()
