' ==========================================
' 견적공종가져오기 최적화
' ==========================================
Sub 견적공종가져오기()
    Dim shTomo As Worksheet, shMacro As Worksheet
    Dim shStart As Worksheet
    Dim i As Long
    
    Set shStart = ActiveSheet
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

    ' 선택 구역 해제 생략 (Select 제거)

CleanUp:
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    If Not shStart Is Nothing Then shStart.Activate
    Exit Sub

ErrorHandler:
    MsgBox "견적공종가져오기 중 오류 발생: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Sub 견적의뢰서생성()

' 2025.01.14 수정 (최적화 및 결과 보고 기능 추가)

    Dim L1 As Long, L2 As Long, cnt As Long, StartRowN As Long, EndRowN As Long
    Dim StartRowN2 As Long, EndRowN2 As Long
    Dim StartRowN1 As Long, EndRowN1 As Long
    Dim shStart As Worksheet
   
    Set shStart = ActiveSheet
    Dim Gongjong As String, Mypath As String
    Dim Mywork As String, Myfile As String, TGfile As String
    Dim MyPjtName As String
    
    Dim T1 As Long, T2 As Long, T6 As Long
    
    Dim totalCnt As Long, successCnt As Long, failCnt As Long
    Dim statusMsg As String
    
    ' 기존 성능 옵션 상태 저장
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    
    With Application
        oldScreenUpdating = .ScreenUpdating
        oldCalculation = .Calculation
        oldEvents = .EnableEvents
        oldDisplayAlerts = .DisplayAlerts
        
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    On Error GoTo ErrorHandler
    
    Dim shTomo As Worksheet, shTemp As Worksheet, shMacro As Worksheet, shInq As Worksheet, shGong As Worksheet
    Set shTomo = Sheets("토목실행")
    Set shTemp = Sheets("임시")
    Set shMacro = Sheets("부가기능(매크로)")
    Set shInq = Sheets("견적의뢰서")
    Set shGong = Sheets("공내역서")
    
    MyPjtName = Sheets("부대경상비").Range("c2").Value
    Mypath = ActiveWorkbook.Path
    Myfile = ActiveWorkbook.Name

    shTemp.Range("b6:g1001").ClearContents

    ' 전체 줄수 알아내기
    For T1 = 5 To 30000
        If shTomo.Cells(T1, 10).Value = "END" Then
            EndRowN = T1 - 1
            Exit For
        End If
    Next T1

    ' 내역의 첫번째 줄 알아내기
    For T2 = 5 To 30000
        If shTomo.Cells(T2, 9).Value = 2 Then
            EndRowN2 = T2
            Exit For
        End If
    Next T2
    
    ' 임시 데이터 작업
    With shTomo
        .Cells(T2, 8).Value = T2
        .Range(.Cells(T2, 8), .Cells(T1, 8)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, Step:=1, Trend:=False
        .Range(.Cells(T2, 1), .Cells(T1, 73)).AutoFilter Field:=9, Criteria1:="<>"
        
        .Range(.Cells(T2, 9), .Cells(T1, 9)).Copy
        shTemp.Range("b2").PasteSpecial Paste:=xlPasteValues
        
        .Range(.Cells(T2, 8), .Cells(T1 - 1, 8)).Copy
        shTemp.Range("c2").PasteSpecial Paste:=xlPasteValues
        
        If .AutoFilterMode Then .ShowAllData
        .Range(.Cells(T2, 8), .Cells(T1 - 1, 8)).ClearContents
    End With
    
    For T6 = 2 To 1000
        If shTemp.Cells(T6, 2).Value = "" Then
            shTemp.Range(shTemp.Cells(3, 4), shTemp.Cells(T6 - 1, 7)).FillDown
            Exit For
        End If
    Next T6

    ' 직접공사비 내역부분 확인
    For L1 = 5 To 65536
        If shTomo.Cells(L1, 10).Value = "직 접 공 사 비 계" Then
            StartRowN = L1
            Exit For
        End If
    Next L1
    
    For L2 = 5 To 65536
        If shTomo.Cells(L2, 25).Value = "END" Then
            EndRowN = L2 - 1
            Exit For
        End If
    Next L2

    StartRowN1 = shMacro.Cells(4, 7).Value
    EndRowN1 = shMacro.Cells(5, 7).Value

    ' 1. 토철공사 견적의뢰서 만들기
    totalCnt = totalCnt + 1
    On Error Resume Next
    shInq.Cells(5, 3).Value = "토철 공사 실행견적 요청"
    shGong.Copy After:=shGong
    
    Set wsNew = ActiveSheet
    wsNew.Name = "토목실행(토철)"
    Mywork = wsNew.Name
    
    wsNew.Cells(1, 10).Value = "공사명 : " & MyPjtName & " (HDC현대산업개발 의뢰)"
    wsNew.Range(wsNew.Cells(5, 1), wsNew.Cells(4 + 34 - L1, 1)).EntireRow.Delete
    
    shTomo.Range(shTomo.Cells(4, 9), shTomo.Cells(L2, 27)).Copy wsNew.Range("i4")
    shTomo.Range(shTomo.Cells(4, 25), shTomo.Cells(L2, 25)).Copy wsNew.Range("z4")
    
    shTomo.Range(shTomo.Cells(4, 1), shTomo.Cells(L2, 1)).Copy
    wsNew.Range("x4").PasteSpecial Paste:=xlPasteFormulas
    wsNew.Range("x4:x" & wsNew.Cells(wsNew.Rows.Count, "X").End(xlUp).Row).Style = "Comma [0]"
    
    wsNew.Range(wsNew.Cells(L2 + 1, 1), wsNew.Cells(15000, 1)).EntireRow.Delete
    
    shTomo.Range(shTomo.Cells(4, 19), shTomo.Cells(L2 - 1, 19)).Copy
    wsNew.Range("w4").PasteSpecial Paste:=xlPasteFormulas
    
    With wsNew.Range(wsNew.Cells(4, 24), wsNew.Cells(L2 - 1, 24))
        .AutoFilter Field:=15, Criteria1:="<>1"
        .ClearContents
        .AutoFilter Field:=15
        .Replace 1, "토철"
    End With
    
    wsNew.Range(wsNew.Cells(2, 22), wsNew.Cells(L2, 23)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=0
        
    wsNew.Range(wsNew.Cells(4, 22), wsNew.Cells(L2, 22)).ClearContents
    wsNew.Range(wsNew.Cells(4, 27), wsNew.Cells(L2, 27)).ClearContents
    
    shGong.Range(shGong.Cells(4, 25), shGong.Cells(L2, 25)).Copy
    wsNew.Range("y4").PasteSpecial Paste:=xlPasteFormulas
    With wsNew.Range("y4:y" & L2)
        .Style = "Percent"
        .NumberFormatLocal = "0.00%"
    End With
  
    wsNew.Columns("N:S").Hidden = True
    wsNew.Range(wsNew.Cells(EndRowN + 1, 10), wsNew.Cells(EndRowN + 1, 27)).FillRight
    
    Sheets(Array("견적의뢰서", Mywork, "관급자재")).Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=Mypath & "\(견적의뢰_토철)_" & MyPjtName & "_HDC현대산업개발.xlsx", FileFormat:=xlOpenXMLWorkbook
    
    wbNew.BreakLink Name:=Mypath & "\" & Myfile, Type:=xlExcelLinks
    wbNew.Close SaveChanges:=True
    
    wsNew.Delete
    If Err.Number = 0 Then successCnt = successCnt + 1 Else failCnt = failCnt + 1
    On Error GoTo ErrorHandler

    ' 2. 선택한 항목 공사 견적의뢰서 만들기
    For cnt = StartRowN1 To EndRowN1
        totalCnt = totalCnt + 1
        On Error Resume Next
        
        Gongjong = shMacro.Cells(cnt + 8, 7).Value
        shTemp.Range("c1").Value = Gongjong
        shInq.Cells(5, 3).Value = Gongjong & " 공사 실행견적 요청"
        
        shGong.Copy After:=shGong
        Set wsNew = ActiveSheet
        wsNew.Name = "토목실행(" & Gongjong & ")"
        Mywork = wsNew.Name
        
        wsNew.Cells(1, 10).Value = "공사명 : " & MyPjtName & " (HDC현대산업개발 의뢰)"
        wsNew.Range(wsNew.Cells(5, 1), wsNew.Cells(4 + 34 - L1, 1)).EntireRow.Delete
        
        shTomo.Range(shTomo.Cells(4, 9), shTomo.Cells(L2, 27)).Copy wsNew.Range("i4")
        shTomo.Range(shTomo.Cells(4, 25), shTomo.Cells(L2, 25)).Copy wsNew.Range("z4")
        
        shTomo.Range(shTomo.Cells(4, 2), shTomo.Cells(L2, 2)).Copy
        wsNew.Range("x4").PasteSpecial Paste:=xlPasteFormulas
        
        wsNew.Range(wsNew.Cells(L2 + 1, 1), wsNew.Cells(15000, 1)).EntireRow.Delete
        
        shTomo.Range(shTomo.Cells(4, 19), shTomo.Cells(L2 - 1, 19)).Copy
        wsNew.Range("w4").PasteSpecial Paste:=xlPasteFormulas
        
        With wsNew.Range(wsNew.Cells(4, 24), wsNew.Cells(L2 - 1, 24))
            .AutoFilter Field:=15, Criteria1:="<>" & Gongjong
            .ClearContents
            .AutoFilter Field:=15
        End With
        
        wsNew.Range(wsNew.Cells(2, 22), wsNew.Cells(L2, 23)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=0
        
        wsNew.Range(wsNew.Cells(4, 22), wsNew.Cells(L2, 22)).ClearContents
        wsNew.Range(wsNew.Cells(4, 27), wsNew.Cells(L2, 27)).ClearContents
        
        shGong.Range(shGong.Cells(4, 25), shGong.Cells(L2, 25)).Copy
        wsNew.Range("y4").PasteSpecial Paste:=xlPasteFormulas
        With wsNew.Range("y4:y" & L2)
            .Style = "Percent"
            .NumberFormatLocal = "0.00%"
        End With
        
        wsNew.Columns("N:S").Hidden = True

        ' 임시 시트 기반 숨김 로직 (기본 로직 유지)
        Dim a3 As Long, a1 As Long
        For a3 = 2 To 1000
            If shTemp.Cells(a3, 2).Value = "END" Then Exit For
        Next a3
       
        For a1 = 1 To a3 - 1
            If shTemp.Cells(a1, 7).Value = "숨김" Then
                wsNew.Rows(shTemp.Cells(a1, 3).Value & ":" & shTemp.Cells(a1 + 1, 3).Value - 1).Hidden = True
            End If
        Next a1

        ' 계층 구조 유지 숨기기 알고리즘
        Dim rH As Long, fRow As Long, minSpc As Integer, curSpc As Integer, vJ As String
        fRow = 4
        For rH = 4 To L2
            If Replace(wsNew.Cells(rH, 10).Text, " ", "") = "총공사비계" Then
                fRow = rH + 1
                Exit For
            End If
        Next rH
        
        minSpc = 999
        For rH = L2 To fRow Step -1
            vJ = wsNew.Cells(rH, 10).Text
            curSpc = 0
            Do While Mid(vJ, curSpc + 1, 1) = " "
                curSpc = curSpc + 1
            Loop
            If Trim(vJ) = "" Then curSpc = 999
            
            If Trim(wsNew.Cells(rH, 24).Text) <> "" Then
                wsNew.Rows(rH).Hidden = False
                minSpc = curSpc
            Else
                If curSpc < minSpc And minSpc <> 999 Then
                    wsNew.Rows(rH).Hidden = False
                    minSpc = curSpc
                Else
                    wsNew.Rows(rH).Hidden = True
                End If
            End If
        Next rH

        wsNew.Range(wsNew.Cells(EndRowN + 1, 10), wsNew.Cells(EndRowN + 1, 27)).FillRight
        
        Sheets(Array("견적의뢰서", Mywork, "관급자재")).Copy
        Set wbNew = ActiveWorkbook
        wbNew.SaveAs Filename:=Mypath & "\(견적의뢰_" & Gongjong & ")_" & MyPjtName & "_HDC현대산업개발.xlsx", FileFormat:=xlOpenXMLWorkbook
        
        wbNew.BreakLink Name:=Mypath & "\" & Myfile, Type:=xlExcelLinks
        wbNew.Close SaveChanges:=True
        
        wsNew.Delete
        
        If Err.Number = 0 Then successCnt = successCnt + 1 Else failCnt = failCnt + 1
        On Error GoTo ErrorHandler
    Next cnt

CleanUp:
    With Application
        .ScreenUpdating = oldScreenUpdating
        .Calculation = oldCalculation
        .EnableEvents = oldEvents
        .DisplayAlerts = oldDisplayAlerts
    End With
    
    statusMsg = "견적의뢰서 생성 완료" & vbCrLf & vbCrLf & _
                "전체 대상: " & totalCnt & "건" & vbCrLf & _
                "성공: " & successCnt & "건" & vbCrLf & _
                "실패: " & failCnt & "건"
    
    If failCnt > 0 Then
        MsgBox statusMsg, vbExclamation, "작업 결과 보고"
    Else
        MsgBox statusMsg, vbInformation, "작업 결과 보고"
    End If
    
    ' 원래 시트로 안전하게 복귀
    If Not shStart Is Nothing Then shStart.Activate
    Exit Sub

ErrorHandler:
    MsgBox "견적의뢰서 생성 중 예상치 못한 오류 발생: " & Err.Description, vbCritical
    Resume CleanUp
End Sub