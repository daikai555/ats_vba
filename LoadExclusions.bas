Attribute VB_Name = "LoadExclusions"
'--- LoadExclusions : 設定シートの除外銘柄リストを辞書で取得
Private Function LoadExclusions() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = Worksheets("設定")
    Dim hdr As Range, r As Long, code As String, qty As Long
    
    Set hdr = ws.Columns(1).Find("除外銘柄", LookAt:=xlWhole)
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
