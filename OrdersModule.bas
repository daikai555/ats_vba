Attribute VB_Name = "OrdersModule"
'###OrdersModule_full �������� �|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
Option Explicit
'===============================================================================
' �������Ɋւ��邷�ׂĂ̏������W��
' �E�e�X�g���[�h����
' �E���g�p���� ID �擾
' �E�������ʌv�Z
' �E�M�p�V�K�i���j����  ���󔄂�ǉ����͂������g��
' �E�����^��胍�O
'===============================================================================

'�� �e�X�g���[�h����i�ݒ�V�[�g B21 �Z������ 1=�e�X�g/0=�{�ԁj
Function IsTestMode() As Boolean
    IsTestMode = (GetSettingValue("�e�X�g���[�h") = 1)
End Function

'---------------------------------------------------------------
' ���g�p�̔��� ID ���擾�iRssOrderIDList �𗘗p�j
'---------------------------------------------------------------
Function GetNextOrderID() As Long
    Dim used As Variant, i As Long
    used = Application.Run("RssOrderIDList")   '�z�� or ������
    For i = 1 To 999999
        If IsError(Application.Match(i, used, 0)) Then
            GetNextOrderID = i
            Exit Function
        End If
    Next i
End Function

'---------------------------------------------------------------
' ���]�͏���ƌ��ݒl���甭�����ʂ��Z�o
'---------------------------------------------------------------
Function CalcOrderQty(curPrice As Double) As Long
    Dim cap As Double, unitLot As Long, qty As Long
    cap = GetSettingValue("���]�͏���i���~�j") * 10000
    unitLot = GetSettingValue("�ŏ������P��")
    If unitLot <= 0 Then unitLot = 100
    
    qty = Int(cap / curPrice / unitLot) * unitLot
    CalcOrderQty = Application.Max(unitLot, qty)
End Function

'---------------------------------------------------------------
' �Ď��V�[�g F �� =�u���v�s��M�p�V�K�iRssMarginOpenOrder_v�j�Ŕ���
'---------------------------------------------------------------
Sub PlaceBuyOrders()
    Dim ws As Worksheet: Set ws = Worksheets("�Ď�")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim priceType As Long: priceType = GetSettingValue("�������i�敪")    '0=���s 1=�w�l
    Dim creditType As Long: creditType = GetSettingValue("�M�p�敪")      '1=���x 2=���
    Dim acctType As Long: acctType = 0    '�M�p���������
    
    Dim maxPos As Long: maxPos = GetSettingValue("�ő哯���ۗL��")
    Dim curPos As Long: curPos = Application.CountIf(ws.Range("G2:G" & lastRow), ">*0") '�ۗL����>0
    
    Dim i As Long, code As String, curP As Double
    Dim qty As Long, price As Double, orderID As Long
    
    For i = 2 To lastRow
        '�\�\ ����t���O
        If ws.Cells(i, "F").Value <> "��" Then GoTo ContinueLoop
        If curPos >= maxPos Then GoTo ContinueLoop
        
        code = Format(ws.Cells(i, "A").Value, "0000")
        curP = Val(ws.Cells(i, "D").Value)
        If curP <= 0 Then GoTo ContinueLoop
        
        qty = CalcOrderQty(curP)
        If qty <= 0 Then GoTo ContinueLoop
        
        '�\�\ ���s or �w�l
        If priceType = 0 Then
            price = 0
        Else
            price = SafeRss(code, "�ŗǔ��C�z�l1")
            If price = 0 Then price = curP
        End If
        
        orderID = GetNextOrderID()
        
        '�\�\ ����
        If IsTestMode() Then
            AppendLog "TEST-BUY", code, qty, price
        Else
            On Error Resume Next
            Application.Run "RssMarginOpenOrder_v", _
                            orderID, code, qty, price, priceType, _
                            creditType, acctType, 0, 0   'SOR=0, ����
            On Error GoTo 0
            AppendLog "BUY", code, qty, price
        End If
        
        ws.Cells(i, "F").Value = "������"
        curPos = curPos + 1
        
ContinueLoop:
    Next i
End Sub

'---------------------------------------------------------------
' �����E��胍�O�isheet name: ���O�j
'---------------------------------------------------------------
Sub AppendLog(act As String, code As String, qty As Long, price As Double)
    Dim wsL As Worksheet
    On Error Resume Next
    Set wsL = Worksheets("���O")
    If wsL Is Nothing Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "���O"
        Set wsL = Worksheets("���O")
    End If
    On Error GoTo 0
    
    If wsL.Cells(1, "A").Value <> "����" Then
        wsL.[A1:E1].Value = Array("����", "�敪", "�R�[�h", "����", "���i")
    End If
    
    Dim n As Long: n = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
    wsL.Cells(n, 1).Resize(1, 5).Value = Array(Now, act, code, qty, price)
End Sub
'###OrdersModule_full �����܂� �|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|


