Attribute VB_Name = "M02_담당자최종실행복사"
Sub 담당자실행값복사()
Attribute 담당자실행값복사.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 담당자실행값복사 매크로
'

'
    ' 1. 현재 활성화된 시트 데이터 값만 복사
    Range("E12:G200").Copy
    Range("I12").PasteSpecial Paste:=xlPasteValues
    
    ' 2. "대비표" 시트 데이터 값만 복사
    Sheets("대비표").Range("C4:C25").Copy
    Sheets("대비표").Range("F4").PasteSpecial Paste:=xlPasteValues
    
    ' 3. 클립보드 초기화
    Application.CutCopyMode = False
    
    ' 4. "부대경상비" 시트로 이동
    Sheets("부대경상비").Select

End Sub
