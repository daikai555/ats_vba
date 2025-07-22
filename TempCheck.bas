Attribute VB_Name = "TempCheck"
Sub ScanModule1()
    Dim cm As Object      '�� ������ Object ��
    Set cm = ThisWorkbook.VBProject.VBComponents("Module1").CodeModule
    
    Dim i As Long, txt As String
    For i = 1 To cm.CountOfLines
        txt = cm.Lines(i, 1)
        
        '--- SheetExists / End Function �s�������o��
        If InStr(txt, "SheetExists") > 0 Or InStr(txt, "End Function") > 0 Then
            Debug.Print i, txt
        End If
        
        '--- �R�����g�ł��錾�s�ł��Ȃ��s���u���^���v�Ƃ��ďo��
        If Len(txt) > 0 _
           And Left$(Trim$(txt), 1) <> "'" _
           And Not txt Like "End*" _
           And Not txt Like "Option*" Then
            Debug.Print "���^��", i, txt
        End If
    Next i
End Sub

Sub RemoveFullWidthSpaces()
    Dim cm As Object, i As Long, txt As String
    Set cm = ThisWorkbook.VBProject.VBComponents("Module1").CodeModule
    For i = 1 To cm.CountOfLines
        txt = cm.Lines(i, 1)
        If InStr(txt, ChrW(&H3000)) > 0 Then          'ChrW(&H3000)=�S�p�X�y�[�X
            txt = Replace(txt, ChrW(&H3000), " ")
            cm.ReplaceLine i, txt
        End If
    Next i
    MsgBox "�S�p�X�y�[�X�𔼊p�ɒu�����܂����B", vbInformation
End Sub

