Attribute VB_Name = "M10_실행직비검토"
Sub Verification()
'
' 부대경상비의 직비부분 집계금액과 토목실행 내역금액의 비교
'
    Dim i As Long, L As Long
    Dim Asum As Double, Bsum As Double
    Dim shBudget As Worksheet, shTomo As Worksheet
    
    Set shBudget = Sheets("부대경상비")
    Set shTomo = Sheets("토목실행")

   '부대경상비 집계금액 합계
    For L = 14 To 30
       If shBudget.Cells(L, 3).Value = "** 업 체 잡 비" Then Exit For
    Next L
    Asum = Application.Sum(shBudget.Range(shBudget.Cells(14, 7), shBudget.Cells(L - 1, 7)))
    
   '토목실행 내역금액 합계
    Bsum = 0
    For i = 4 To shTomo.Cells(shTomo.Rows.Count, 22).End(xlUp).Row
       If shTomo.Cells(i, 22).Value = "END" Then Exit For
       If shTomo.Cells(i, 22).Value <> 0 Then
           Bsum = Bsum + shTomo.Cells(i, 23).Value
       End If
    Next i
    
   '합계비교 메세지 출력
    If Bsum - Asum = 0 Then
        MsgBox "부대경상비 집계금액 = " & Format(Asum, "#,##0") & vbCrLf & vbCrLf _
               & "실행내역 합계금액 = " & Format(Bsum, "#,##0") & vbCrLf & vbCrLf _
               & "합계금액 차이 = " & Format(Bsum - Asum, "#,##0") & vbCrLf & vbCrLf _
               & "OK!! 완벽하시네요!!"
    Else
        MsgBox "부대경상비 집계금액 = " & Format(Asum, "#,##0") & vbCrLf & vbCrLf _
               & "실행내역 합계금액 = " & Format(Bsum, "#,##0") & vbCrLf & vbCrLf _
               & "합계금액 차이 = " & Format(Bsum - Asum, "#,##0") & vbCrLf & vbCrLf _
               & "금액이 맞지않습니다."
    End If
    
End Sub
