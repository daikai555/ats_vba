Attribute VB_Name = "Module1"
Option Explicit

'=====  ���ʃ��[�e�B���e�B  ======================================

'--- SheetExists : �w��V�[�g�̗L�����m�F
Private Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(name) Is Nothing
    On Error GoTo 0
End Function



'=====  �O���[�o������t���O  =====================================

Public AutoRun As Boolean   'True = �������� / False = ��~

'=====  �Ď��V�[�g���בւ�  =======================================

Sub AutoSortByFlag()
    Const FLAG_COL As Long = 6     'F��
    Const FALLBACK_COL As Long = 27 'AA��i�㏸��%�j
    
    Dim ws As Worksheet: Set ws = Worksheets("�Ď�")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then Exit Sub
    
    Dim scoreCol As Long, hdr As Range, keyWords, k As Long
    keyWords = Array("�㏸��", "�����X�R�A")
    For k = LBound(keyWords) To UBound(keyWords)
        Set hdr = ws.Rows(1).Find(keyWords(k), LookAt:=xlPart)
        If Not hdr Is Nothing Then scoreCol = hdr.Column: Exit For
    Next k
    If scoreCol = 0 Then scoreCol = FALLBACK_COL
    
    Dim pri: pri = Array("�ۗL��", "������", "��", "���m", "����", "��", "")
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add ws.Range(ws.Cells(2, FLAG_COL), ws.Cells(lastRow, FLAG_COL)), _
                        xlSortOnValues, xlAscending, , xlSortNormal
        .SortFields(.SortFields.Count).CustomOrder = Join(pri, ",")
        .SortFields.Add ws.Range(ws.Cells(2, scoreCol), ws.Cells(lastRow, scoreCol)), _
                        xlSortOnValues, xlDescending
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        .Header = xlYes
        .Apply
    End With
End Sub

'=====  �Ǘ��n���[�e�B���e�B  =====================================

Private Sub CenterBold(rg As Range)
    With rg
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
End Sub

Sub ImportCsvToActiveSheet()
    Dim fp As String: fp = Application.GetOpenFilename("CSV (*.csv),*.csv")
    If fp = "False" Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim tmp As Workbook, arr, i&, mx&
    Set tmp = Workbooks.Open(fp, ReadOnly:=True, Local:=True)
        mx = tmp.Sheets(1).Cells(tmp.Sheets(1).Rows.Count, 1).End(xlUp).Row
        arr = tmp.Sheets(1).Range("A1").Resize(mx, 1).Value
    tmp.Close False
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Range("A2:A501").ClearContents
    For i = 1 To UBound(arr, 1)
        If i > 500 Then Exit For
        ws.Cells(i + 1, 1).Value = arr(i, 1)
    Next i
    CenterBold ws.Range("A2:A501")
    
    On Error Resume Next
        Worksheets("Ticks").Rows("2:" & Rows.Count).ClearContents
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    MsgBox "CSV ��荞�݊���", vbInformation
End Sub

Sub StartMonitoring(): AutoRun = True:  MsgBox "�Ď��J�n", vbInformation: End Sub
Sub StopMonitoring():  AutoRun = False: MsgBox "�Ď���~", vbInformation: End Sub

Sub ClearPreviousLog()
    On Error Resume Next
        Worksheets("���O").Cells.Clear
        Worksheets("Positions").Cells.Clear
    On Error GoTo 0
    MsgBox "�O�������N���A", vbInformation
End Sub

'=====  �t���O�X�V (Orders / Positions �Q��) ======================

Sub UpdateFlags()
    If Not SheetExists("Orders") Or Not SheetExists("Positions") Then Exit Sub
    
    Dim wsW As Worksheet: Set wsW = Worksheets("�Ď�")
    Dim wsO As Worksheet: Set wsO = Worksheets("Orders")
    Dim wsP As Worksheet: Set wsP = Worksheets("Positions")
    Dim ex As Object:     Set ex = LoadExclusions()
    
    Dim posD As Object: Set posD = CreateObject("Scripting.Dictionary")
    Dim ob  As Object: Set ob = CreateObject("Scripting.Dictionary")
    Dim os  As Object: Set os = CreateObject("Scripting.Dictionary")
    
    Dim i&, code$, qty#
    
    With wsP
        For i = 2 To .Cells(.Rows.Count, 4).End(xlUp).Row
            code = Trim$(.Cells(i, 4).Value)
            qty = Val(.Cells(i, 5).Value)
            If ex.Exists(code) Then qty = qty - ex(code)
            If qty > 0 Then posD(code) = qty
        Next i
    End With
    
    With wsO
        For i = 2 To .Cells(.Rows.Count, 4).End(xlUp).Row
            If Val(.Cells(i, 12).Value) = 0 Then
                code = Trim$(.Cells(i, 4).Value)
                If .Cells(i, 7).Value = 1 Then ob(code) = True
                If .Cells(i, 7).Value = 2 Then os(code) = True
            End If
        Next i
    End With
    
    Dim last&, wsR As Range
    last = wsW.Cells(wsW.Rows.Count, 1).End(xlUp).Row
    Application.EnableEvents = False
    For i = 2 To last
        code = Trim$(wsW.Cells(i, 1).Value)
        Set wsR = wsW.Cells(i, 6)
        If posD.Exists(code) Then
            wsR.Value = "�ۗL��": wsR.Offset(0, 1).Value = posD(code)
        ElseIf ob.Exists(code) Or os.Exists(code) Then
            wsR.Value = "������": wsR.Offset(0, 1).Value = 0
        Else
            wsR.Value = "": wsR.Offset(0, 1).Value = 0
        End If
    Next i
    Application.EnableEvents = True
End Sub

