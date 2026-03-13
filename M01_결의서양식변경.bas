Attribute VB_Name = "M01_결의서양식변경"
Sub Fast()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
End Sub

Sub Slow()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub 양식()
'
' 양식 매크로
'
    Fast
    
    Range("A45:A46").Copy Destination:=Range("A22")
    Range("B45:M56").Copy Destination:=Range("B22")
    
    Range("A1").Select
    Slow
    
    ' 최근 입찰 10개 정보 업데이트
    Recentbid
End Sub

Sub 설계()
'
' 설계 매크로
'
    Fast
   
    Range("A58:A59").Copy Destination:=Range("A22")
    Range("B22:M33").UnMerge
    Range("B58:M69").Copy Destination:=Range("B22")
    
    Range("A1").Select
    Slow
End Sub

Sub Recentbid()
    Dim i As Long, X As VbMsgBoxResult, TGLine As Long
    Dim a As String, b As String, t As String, TGfile As String, Myfile As String
   
    Myfile = ActiveWorkbook.Name
    Sheets("결의서").Select

    a = Cells(3, 26).Value    '최근
    b = Cells(5, 26).Value    '구분
    
    X = MsgBox("최근 = " & a & vbCrLf & vbCrLf _
               & "구분 = " & b & vbCrLf & vbCrLf _
               & "최근 입찰 정보를 연동하시겠습니까?", vbYesNo + vbQuestion)
   
    If X = vbYes Then
        ' 03.공동도급 현황 파일에서 최근입찰 10개 정보 연동
        If b = "경상북샘플" Or b = "경상북샘플(공)" Or b = "경상 샘플" Then
            b = "샘플1"
        Else
            b = "샘플2"
        End If
        
        ' 네트워크 경로에서 파일 열기
        On Error Resume Next
        Workbooks.Open Filename:="\\218.153.46.220\infra\공동도급\03.공동도급현황.xlsx"
        If Err.Number <> 0 Then
            MsgBox "공동도급현황 파일을 열 수 없습니다.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        
        TGfile = ActiveWorkbook.Name
        
        With Workbooks(TGfile).Sheets("공동도급")
            .Cells(1, 16).Value = b
            .Cells(2, 16).Value = a
            .Range("B3:M12").Copy
        End With
        
        With Workbooks(Myfile).Sheets("결의서")
            .Range("B24").PasteSpecial Paste:=xlPasteValues
            .Range("B24").Select
        End With
        
        Application.CutCopyMode = False
        
        Application.DisplayAlerts = False
        Workbooks(TGfile).Close SaveChanges:=False
        Application.DisplayAlerts = True
        
        Workbooks(Myfile).Save
    End If
End Sub
