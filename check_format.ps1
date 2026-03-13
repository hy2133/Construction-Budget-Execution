$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open('d:\문서\MyClaudeCode\실행예산서\★실행결의서_남부내륙철도(김천~거제) 건설사업 제2공구 노반신설 기타공사.xlsb')

function Get-FormatConditions($sheetName) {
    Write-Host "--- Checking Conditional Formatting on $sheetName ---"
    $sheet = $workbook.Sheets.Item($sheetName)
    $fcs = $sheet.Cells.FormatConditions
    Write-Host "Count: $($fcs.Count)"
    foreach ($fc in $fcs) {
        try {
            Write-Host "Type: $($fc.Type)"
            Write-Host "AppliesTo: $($fc.AppliesTo.Address)"
            if ($fc.Type -eq 1) {
                # xlCellValue
                Write-Host "Operator: $($fc.Operator)"
                Write-Host "Formula1: $($fc.Formula1)"
            }
        }
        catch {
            Write-Host "Error reading properties: $_"
        }
        Write-Host "---"
    }
}

Get-FormatConditions "공내역서"
Get-FormatConditions "토목실행"

$workbook.Close($false)
$excel.Quit()
