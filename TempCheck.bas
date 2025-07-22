Attribute VB_Name = "TempCheck"
Sub ScanModule1()
    Dim cm As Object      '← ここを Object に
    Set cm = ThisWorkbook.VBProject.VBComponents("Module1").CodeModule
    
    Dim i As Long, txt As String
    For i = 1 To cm.CountOfLines
        txt = cm.Lines(i, 1)
        
        '--- SheetExists / End Function 行だけを出力
        If InStr(txt, "SheetExists") > 0 Or InStr(txt, "End Function") > 0 Then
            Debug.Print i, txt
        End If
        
        '--- コメントでも宣言行でもない行を「※疑い」として出力
        If Len(txt) > 0 _
           And Left$(Trim$(txt), 1) <> "'" _
           And Not txt Like "End*" _
           And Not txt Like "Option*" Then
            Debug.Print "※疑い", i, txt
        End If
    Next i
End Sub

Sub RemoveFullWidthSpaces()
    Dim cm As Object, i As Long, txt As String
    Set cm = ThisWorkbook.VBProject.VBComponents("Module1").CodeModule
    For i = 1 To cm.CountOfLines
        txt = cm.Lines(i, 1)
        If InStr(txt, ChrW(&H3000)) > 0 Then          'ChrW(&H3000)=全角スペース
            txt = Replace(txt, ChrW(&H3000), " ")
            cm.ReplaceLine i, txt
        End If
    Next i
    MsgBox "全角スペースを半角に置換しました。", vbInformation
End Sub

