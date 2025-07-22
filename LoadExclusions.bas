Attribute VB_Name = "LoadExclusions"
'--- LoadExclusions : �ݒ�V�[�g�̏��O�������X�g�������Ŏ擾
Private Function LoadExclusions() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = Worksheets("�ݒ�")
    Dim hdr As Range, r As Long, code As String, qty As Long
    
    Set hdr = ws.Columns(1).Find("���O����", LookAt:=xlWhole)
    If hdr Is Nothing Then
        Set LoadExclusions = dict: Exit Function
    End If
    
    r = hdr.Row + 1
    Do While Len(ws.Cells(r, "A").Value) > 0
        code = Trim$(ws.Cells(r, "A").Value)
        qty = Val(ws.Cells(r, "B").Value)
        If qty > 0 Then dict(code) = dict(code) + qty
        r = r + 1
    Loop
    
    Set LoadExclusions = dict
End Function
