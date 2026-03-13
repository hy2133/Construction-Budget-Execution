Attribute VB_Name = "M03_실행내역수식복사"
Sub Copy()
    Dim rtn As Double
    rtn = MsgBox("집계 생성과 수식복사를 하시겠습니까?" & vbLf & vbLf & "집계생성과 수식복사 -> 예 " & vbLf & "수식만 복사 --> 아니요", vbYesNoCancel, "수식복사의 종류 선택")

    Select Case rtn
        Case vbYes: Copy2
        Case vbNo: Copy1
    End Select
End Sub

Sub Copy2()
    Dim L1, L2, L3, l4, StartRowN, EndRowN As Long
    Dim S1, S2, s3, S4, s5 As Integer
    Dim T1, T2, T3, T4, T5, T6 As Integer

    ' 공사비계 위치 찾기 및 행 삽입
    For L3 = 5 To 65536
        If Sheets("토목실행").Cells(L3, 10).Value = "직 접 공 사 비 계" Then Exit For
    Next L3

    If L3 < 35 Then
        Rows("14:" & (14 + 34 - L3)).Insert Shift:=xlDown
    End If
    Range("j15:s34").ClearContents

    ' 줄수 계산
    For T1 = 5 To 30000
        If Sheets("토목실행").Cells(T1, 10).Value = "END" Then
            EndRowN = T1 - 1
            Exit For
        End If
    Next T1

    For T2 = 5 To 30000
        If Sheets("토목실행").Cells(T2, 9).Value = 2 Then Exit For
    Next T2

    ' 임시 시트 정리 및 데이터 복사
    Sheets("임시").Range("b6:g1001").ClearContents
    Cells(T2, 8).FormulaR1C1 = "=ROW()"
    Range(Cells(T2, 8), Cells(T1, 8)).FillDown
    
    With Sheets("토목실행")
        .Range(.Cells(T2, 1), .Cells(T1, 73)).AutoFilter Field:=9, Criteria1:="<>"
        .Range(.Cells(T2, 9), .Cells(T1 - 1, 9)).Copy
        Sheets("임시").Range("b2").PasteSpecial Paste:=xlPasteValues
        .Range(.Cells(T2, 8), .Cells(T1 - 1, 8)).Copy
        Sheets("임시").Range("c2").PasteSpecial Paste:=xlPasteValues
    End With

    ' 임시 시트 채우기
    For T6 = 2 To 1000
        If Sheets("임시").Cells(T6, 2).Value = "" Then
            Sheets("임시").Range(Sheets("임시").Cells(3, 4), Sheets("임시").Cells(T6 - 1, 7)).FillDown
            Exit For
        End If
    Next T6

    Sheets("토목실행").ShowAllData
    Range(Cells(T2, 8), Cells(T1 - 1, 8)).ClearContents

    ' 수식 복사 파트
    Dim col1 As Variant
    Range(Cells(5, 15), Cells(EndRowN, 15)).Copy
    For Each col1 In Array("Q5", "S5", "W5", "AA5", "AC5", "AE5", "AG5", "AI5", "AK5", "AM5", "AO5", "AQ5", "AV5")
        Range(col1).PasteSpecial Paste:=xlPasteFormulas
    Next col1
    
    Range("u5").Copy Range("u5:u" & Sheets("임시").Range("m1") + 5)
    Range("x5").Copy Range("x5:x" & Sheets("임시").Range("m1") + 5)

    Dim iCol As Integer
    Range(Cells(12, 15), Cells(EndRowN, 15)).Copy
    For iCol = 50 To 72 Step 2
        Cells(12, iCol).PasteSpecial Paste:=xlPasteFormulas
    Next iCol

    ' 시트 정리
    Sheets("대비표").Range("a5:c33").ClearContents
    Sheets("대비표").Range("a4:c" & 3 + Sheets("임시").Range("m1")).FillDown

    Application.CutCopyMode = False
End Sub

Sub Copy1()
    Dim EndRowN As Long, T1 As Long
    For T1 = 5 To 30000
        If Sheets("토목실행").Cells(T1, 10).Value = "END" Then
            EndRowN = T1 - 1
            Exit For
        End If
    Next T1
    
    Dim col1 As Variant
    Range(Cells(5, 15), Cells(EndRowN, 15)).Copy
    For Each col1 In Array("Q5", "S5", "W5", "AA5", "AC5", "AE5", "AG5", "AI5", "AK5", "AM5", "AO5", "AQ5", "AV5")
        Range(col1).PasteSpecial Paste:=xlPasteFormulas
    Next col1
    
    Application.CutCopyMode = False
End Sub
