Attribute VB_Name = "M05_특허공종금액"
Sub 특허공종내역생성()

    Dim shMacro As Worksheet, shSummary As Worksheet, shHado As Worksheet
    Dim hado As String, PatentStart As Long, PatentEnd As Long, cnt As Long
     
    Set shMacro = Worksheets("부가기능(매크로)")
    Set shSummary = Worksheets("(특허)요약")
    
    With shMacro
        PatentStart = .Range("c4").Value
        PatentEnd = .Range("c5").Value
    End With
    
    Application.Calculation = xlCalculationManual

    shSummary.Range("c4:G33").ClearContents

    ' 1단계: 시트 생성 및 기본값 입력
    For cnt = PatentStart To PatentEnd
        hado = shMacro.Cells(cnt + 7, 3).Value
        Worksheets("(특허)-원본").Copy Before:=Worksheets("(특허)-원본")
        
        Set shHado = ActiveSheet
        With shHado
            .Name = hado
            .Cells(3, 9).Value = hado
        End With
    Next cnt

    ' 2단계: 요약 시트에 수식 연결
    For cnt = PatentStart To PatentEnd
        hado = shMacro.Cells(cnt + 7, 3).Value
        With shSummary
            .Cells(cnt + 3, 3).Value = hado
            .Cells(cnt + 3, 4).Formula = "='" & hado & "'!d22"
            .Cells(cnt + 3, 5).Formula = "='" & hado & "'!d23"
            .Cells(cnt + 3, 6).Formula = "='" & hado & "'!d24"
        End With
    Next cnt

    ' 3단계: 요약 시트 정렬
    With shSummary.Sort
        .SortFields.Clear
        .SortFields.Add Key:=shSummary.Range("D4:D33"), Order:=xlDescending
        .SetRange shSummary.Range("C4:G33")
        .Header = xlGuess
        .Apply
    End With

    Application.Calculation = xlCalculationAutomatic
    shMacro.Select

End Sub


Sub 특허공종내역생성2()
'
' 매크로4 매크로 최적화
'
    Dim i As Long
    Dim shMacro As Worksheet, shSummary As Worksheet
    
    Set shMacro = Worksheets("부가기능(매크로)")
    Set shSummary = Worksheets("(특허)요약")
    
    i = shMacro.Cells(5, 3).Value

    ' 1. (특허)요약 시트 정제
    shSummary.Range("C5:AE43").ClearContents
    
    If i > 0 Then
        ' 2. 업체명 복사 (부가기능(매크로) 8행~i+7행 -> (특허)요약 4행)
        shMacro.Range(shMacro.Cells(8, 3), shMacro.Cells(i + 7, 3)).Copy
        shSummary.Range("C4").PasteSpecial Paste:=xlPasteValues
        
        ' 3. 수식 입력 및 일괄 처리
        With shSummary
            ' D4, E4, F4 수식 입력
            .Range("D4:F4").FormulaR1C1 = "=+RC[5]"
            
            ' i가 1이 아닐 경우 4행부터 i+3행까지 아래로 채우기 (D열~AE열)
            If i <> 1 Then
                .Range(.Cells(4, 4), .Cells(i + 3, 31)).FillDown
            End If
        End With
    End If
    
    Application.CutCopyMode = False
    shMacro.Select
     
End Sub

