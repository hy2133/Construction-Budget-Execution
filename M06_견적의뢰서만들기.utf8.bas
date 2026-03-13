' ==========================================
' 견적공종가져오기 최적화
' ==========================================
Sub 견적공종가져오기()
    Dim shTomo As Worksheet, shMacro As Worksheet
    Dim i As Long
    
    Set shTomo = Sheets("토목실행")
    Set shMacro = Sheets("부가기능(매크로)")
    
    On Error GoTo ErrorHandler
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    ' 1. 데이터 복사 및 중복 제거
    shTomo.Range("B22:B10000").Copy
    With shMacro.Range("H9")
        .PasteSpecial Paste:=xlPasteValues
        .Offset(-1, 0).Resize(9993, 1).RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    Application.CutCopyMode = False

    ' 2. 불필요한 숫자 데이터 삭제 (H8:H50 범위)
    For i = 8 To 50
        If IsNumeric(shMacro.Cells(i, 8).Value) Then
            shMacro.Cells(i, 8).ClearContents
        End If
    Next i

    ' 3. 빈 셀 삭제 및 위로 밀기
    On Error Resume Next
    shMacro.Range("H8:H100").SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    On Error GoTo ErrorHandler

    ' 4. (특허) 포함 항목 상단 정렬
    Dim lastRow As Long
    lastRow = shMacro.Cells(shMacro.Rows.Count, "H").End(xlUp).Row
    If lastRow >= 8 Then
        With shMacro
            ' 임시 보조 열(I열)에 가중치 부여: (특허) 포함 시 1, 미포함 시 2
            Dim r As Range
            For Each r In .Range("H8:H" & lastRow)
                If InStr(1, r.Value, "(특허)", vbTextCompare) > 0 Then
                    r.Offset(0, 1).Value = 1
                Else
                    r.Offset(0, 1).Value = 2
                End If
            Next r
            
            ' H-I 열을 I열 기준으로 오름차순 정렬
            With .Sort
                .SortFields.Clear
                .SortFields.Add Key:=shMacro.Range("I8:I" & lastRow), Order:=xlAscending
                .SetRange shMacro.Range("H8:I" & lastRow)
                .Header = xlNo
                .Apply
            End With
            
            ' 임시 열 삭제
            .Range("I8:I" & lastRow).ClearContents
        End With
    End If

    ' 선택 구역 해제 (단일 셀 선택)
    shMacro.Cells(8, 8).Select

CleanUp:
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "견적공종가져오기 중 오류 발생: " & Err.Description, vbCritical
    Resume CleanUp
End Sub



Sub 견적의뢰서생성()

' 2025.01.14 수정

    Dim L1, L2, cnt, StartRowN, EndRowN As Long
   
    Dim StartRowN1, EndRowN1 As Long
   
    Dim tR, tmp As Range
   
    Dim Gongjong, Mypath As String
   
    Dim Mywork, Myfile, TGfile As String
   
    Dim MyPjtName As String
   
    Dim a1, a2, a3, a4 As Interior
    
    Dim T1, T2, T3, T4, T5, T6 As Integer
    
    
    MyPjtName = Sheets("부대경상비").Range("c2")
   
    Mypath = ActiveWorkbook.Path
    Myfile = ActiveWorkbook.Name
   
   

   Sheets("임시").Range("b6:g1001").ClearContents

   Sheets("토목실행").Select
  
  
'
' 전체 줄수 알아내기
     
     For T1 = 5 To 30000
       
        If Sheets("토목실행").Cells(T1, 10) = "END" Then
           
           EndRowN = T1 - 1
    
        Exit For
        
        End If
       
      Next T1

'내역의 첫번째 줄 알아내기
      For T2 = 5 To 30000
       
        If Sheets("토목실행").Cells(T2, 9) = 2 Then
           
           EndRowN2 = T2
    
        Exit For
        
        End If
       
      Next T2
'
    
    
    Cells(T2, 8).FormulaR1C1 = "=ROW()"
    Range(Cells(T2, 8), Cells(T1, 8)).Select
    Selection.FillDown
    ActiveSheet.Range(Cells(T2, 1), Cells(T1, 73)).AutoFilter Field:=9, Criteria1:="<>"
    Cells(T2, 9).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Cells(T2, 9), Cells(T1 - 1, 9)).Select
    Selection.Copy
    Range(Cells(T2, 9), Cells(T1, 9)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("임시").Select
    Range("b2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("토목실행").Select
    Range(Cells(T2, 8), Cells(T1 - 1, 8)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("임시").Select
    Range("c2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
      For T6 = 2 To 1000
       
        If Sheets("임시").Cells(T6, 2) = "" Then
           
           Range(Cells(3, 4), Cells(T6 - 1, 7)).Select
           Selection.FillDown
               
        Exit For
        
        End If
       
      Next T6
    

    Sheets("토목실행").Select
    ActiveSheet.ShowAllData

    Range(Cells(T2, 8), Cells(T1 - 1, 8)).ClearContents
    
    
    Sheets("부가기능(매크로)").Select
    



'직접공사비 내역부분 확인==========================================================================
   
    For L1 = 5 To 65536
       
        If Sheets("토목실행").Cells(L1, 10) = "직 접 공 사 비 계" Then
       
           StartRowN = L1
    
        Exit For
        
        End If
       
    Next L1
    
    For L2 = 5 To 65536
       
        If Sheets("토목실행").Cells(L2, 25) = "END" Then
           
           EndRowN = L2 - 1
    
        Exit For
        
        End If
       
    Next L2


    StartRowN1 = Cells(4, 7)

    EndRowN1 = Cells(5, 7)

'직접공사비 내역부분 확인 끝=======================================================================


'토철공사 견적의뢰서 만들기(기본적으로 필수생성)===================================================


        Sheets("견적의뢰서").Select
        
        Cells(5, 3).Value = "토철 공사 실행견적 요청"
        
        Worksheets("공내역서").Copy After:=Worksheets("공내역서")  '만들 쉬트를 복사하기
 
        ActiveSheet.Name = "토목실행(토철)"
        
        Mywork = ActiveSheet.Name
        
        Cells(1, 10).Value = "공사명 : " & MyPjtName & " (HDC현대산업개발 의뢰)"
        
        Range(Cells(5, 1), Cells(4 + 34 - L1, 1)).Select
        
        Selection.EntireRow.Delete
        
        Sheets("토목실행").Select
        
        Range(Cells(4, 9), Cells(L2, 27)).Select
      
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("i4").Select
  
        ActiveSheet.Paste

    
        Sheets("토목실행").Select
        
        Range(Cells(4, 25), Cells(L2, 25)).Select
        
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("z4").Select
  
        ActiveSheet.Paste
        
        Sheets("토목실행").Select
        
        Range(Cells(4, 1), Cells(L2, 1)).Select
        
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("x4").Select
  
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        Selection.Style = "Comma [0]"
        
        Range(Cells(L2 + 1, 1), Cells(15000, 1)).Select
        
        Selection.EntireRow.Delete

        Range("w30").Select
        
        Range(Cells(4, 19), Cells(L2 - 1, 19)).Select
        
        Selection.Copy
        
        Range("w4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range(Cells(4, 24), Cells(L2 - 1, 24)).AutoFilter Field:=15, Criteria1:="<>1", Operator:=xlAnd
     
    Range(Cells(4, 24), Cells(L2 - 1, 24)).Select
    
    Application.CutCopyMode = False
     
    Selection.ClearContents
    
    Range(Cells(4, 24), Cells(L2 - 1, 24)).AutoFilter Field:=15
        
    Range(Cells(4, 24), Cells(L2 - 1, 24)).Replace 1, "토철"
        
        
        
    Range(Cells(2, 22), Cells(L2, 23)).Select

    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
        
    Range(Cells(4, 22), Cells(L2, 22)).ClearContents
     
    Range(Cells(4, 27), Cells(L2, 27)).ClearContents
        
    Cells(22, 10).Select
        
  
    Fast
    
    Sheets("공내역서").Select
  
    Range(Cells(4, 25), Cells(L2, 25)).Select
  
    Selection.Copy
  
    Sheets(Mywork).Select
        
    Range("y4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
  
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.00%"
  
    Columns("N:S").Select
    Selection.EntireColumn.Hidden = True
    
    Range(Cells(EndRowN + 1, 10), Cells(EndRowN + 1, 27)).Select
    Selection.FillRight
    
    Range("v42").Select
    
    Sheets(Array("견적의뢰서", Mywork, "관급자재")).Select
    Sheets(Mywork).Activate
    Sheets(Array("견적의뢰서", Mywork, "관급자재")).Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\(견적의뢰_토철)_" & MyPjtName & "_HDC현대산업개발.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    
    
    Application.DisplayAlerts = False
    
    TGfile = ActiveWorkbook.Name
    
    Windows(TGfile).Activate
    
    'DATA 연결해제
        
        ActiveWorkbook.BreakLink Name:= _
            Mypath & "\" & Myfile _
            , Type:=xlExcelLinks
    
    Sheets(3).Select
    
    Range("v44").Select
    
    Sheets(1).Select
    
    ActiveWorkbook.Save
    

    
    ActiveWindow.Close
        
    Worksheets(Mywork).Select
        
    Worksheets(Mywork).Delete
    
    Slow
    'Application.DisplayAlerts = True
        
    Sheets("부가기능(매크로)").Select
    
    
'토철공사 견적의뢰서 만들기(기본적으로 필수생성) 끝===================================================
    
    
'선택한 항목 공사 견적의뢰서 만들기===================================================================
    
    
    For cnt = StartRowN1 To EndRowN1

        Gongjong = Cells(cnt + 8, 7)
        
        Sheets("임시").Range("c1").Value = Gongjong

        Sheets("견적의뢰서").Select
        
        Cells(5, 3).Value = Gongjong & " 공사 실행견적 요청"
        
        Worksheets("공내역서").Copy After:=Worksheets("공내역서")  '만들 쉬트를 복사하기
 
        ActiveSheet.Name = "토목실행(" & Gongjong & ")"
        
        Mywork = ActiveSheet.Name
        
        Cells(1, 10).Value = "공사명 : " & MyPjtName & " (HDC현대산업개발 의뢰)"
        
        Range(Cells(5, 1), Cells(4 + 34 - L1, 1)).Select
        
        Selection.EntireRow.Delete
        
        Sheets("토목실행").Select
        
        Range(Cells(4, 9), Cells(L2, 27)).Select
      
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("i4").Select
  
        ActiveSheet.Paste
    

        Sheets("토목실행").Select
        
        Range(Cells(4, 25), Cells(L2, 25)).Select
        
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("z4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        
        Sheets("토목실행").Select
        
        Range(Cells(4, 2), Cells(L2, 2)).Select
        
        Selection.Copy
        
        Sheets(Mywork).Select
        
        Range("x4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        
        
        Range(Cells(L2 + 1, 1), Cells(15000, 1)).Select
        
        Selection.EntireRow.Delete

        Range("w30").Select
        
        Range(Cells(4, 19), Cells(L2 - 1, 19)).Select
        
        Selection.Copy
        
        Range("w4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range(Cells(4, 24), Cells(L2 - 1, 24)).AutoFilter Field:=15, Criteria1:="<>" & Gongjong _
        , Operator:=xlAnd
     
    Range(Cells(4, 24), Cells(L2 - 1, 24)).Select
    
    Application.CutCopyMode = False
     
    Selection.ClearContents
    
    Range(Cells(4, 24), Cells(L2 - 1, 24)).AutoFilter Field:=15
        
    Range(Cells(2, 22), Cells(L2, 23)).Select

    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
        
        
    Range(Cells(4, 22), Cells(L2, 22)).ClearContents
     
    Range(Cells(4, 27), Cells(L2, 27)).ClearContents
        
    Cells(22, 10).Select
        
    Fast
    
    Sheets("공내역서").Select
  
    Range(Cells(4, 25), Cells(L2, 25)).Select
  
    Selection.Copy
  
    Sheets(Mywork).Select
        
    Range("y4").Select
  
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
        
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.00%"
        
    Columns("N:S").Select
    Selection.EntireColumn.Hidden = True

    Range("V32").Select


      For a3 = 2 To 1000
       
        If Sheets("임시").Cells(a3, 2) = "END" Then
            
          ' a4 = a3

        Exit For
        
        End If
       
      Next a3
   
   
   For a1 = 1 To a3 - 1
   
    If Sheets("임시").Cells(a1, 7) = "숨김" Then
    
          Rows(Sheets("임시").Cells(a1, 3) & ":" & Sheets("임시").Cells(a1 + 1, 3) - 1).Select
          Selection.EntireRow.Hidden = True
          
    End If
    
  Next a1

    Range(Cells(EndRowN + 1, 10), Cells(EndRowN + 1, 27)).Select
    Selection.FillRight
    
    Range("v42").Select
    
    Sheets(Array("견적의뢰서", Mywork, "관급자재")).Select
    Sheets(Mywork).Activate
    Sheets(Array("견적의뢰서", Mywork, "관급자재")).Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\(견적의뢰_" & Gongjong & ")_" & MyPjtName & "_HDC현대산업개발.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    
    
    Application.DisplayAlerts = False
    
    TGfile = ActiveWorkbook.Name
    
    Windows(TGfile).Activate
    
    'DATA 연결해제
        
        ActiveWorkbook.BreakLink Name:= _
            Mypath & "\" & Myfile _
            , Type:=xlExcelLinks
    
    Sheets(3).Select
    
    Range("v44").Select
    
    Sheets(1).Select
    
    
    
    ActiveWorkbook.Save
    
    ActiveWindow.Close
        
    Worksheets(Mywork).Select
        
    Worksheets(Mywork).Delete
    
    
    Slow
    'Application.DisplayAlerts = True
        
    Sheets("부가기능(매크로)").Select

    Next cnt


'선택한 항목 공사 견적의뢰서 만들기 끝===================================================================


End Sub











