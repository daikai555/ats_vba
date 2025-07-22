Attribute VB_Name = "OrdersModule"
'###OrdersModule_full ここから －－－－－－－－－－－－－－－－－－－－
Option Explicit
'===============================================================================
' 実売買に関するすべての処理を集約
' ・テストモード判定
' ・未使用発注 ID 取得
' ・発注数量計算
' ・信用新規（買）発注  ※空売り追加時はここを拡張
' ・発注／約定ログ
'===============================================================================

'▼ テストモード判定（設定シート B21 セル等に 1=テスト/0=本番）
Function IsTestMode() As Boolean
    IsTestMode = (GetSettingValue("テストモード") = 1)
End Function

'---------------------------------------------------------------
' 未使用の発注 ID を取得（RssOrderIDList を利用）
'---------------------------------------------------------------
Function GetNextOrderID() As Long
    Dim used As Variant, i As Long
    used = Application.Run("RssOrderIDList")   '配列 or 文字列
    For i = 1 To 999999
        If IsError(Application.Match(i, used, 0)) Then
            GetNextOrderID = i
            Exit Function
        End If
    Next i
End Function

'---------------------------------------------------------------
' 建余力上限と現在値から発注数量を算出
'---------------------------------------------------------------
Function CalcOrderQty(curPrice As Double) As Long
    Dim cap As Double, unitLot As Long, qty As Long
    cap = GetSettingValue("建余力上限（万円）") * 10000
    unitLot = GetSettingValue("最小売買単位")
    If unitLot <= 0 Then unitLot = 100
    
    qty = Int(cap / curPrice / unitLot) * unitLot
    CalcOrderQty = Application.Max(unitLot, qty)
End Function

'---------------------------------------------------------------
' 監視シート F 列 =「買」行を信用新規（RssMarginOpenOrder_v）で発注
'---------------------------------------------------------------
Sub PlaceBuyOrders()
    Dim ws As Worksheet: Set ws = Worksheets("監視")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim priceType As Long: priceType = GetSettingValue("注文価格区分")    '0=成行 1=指値
    Dim creditType As Long: creditType = GetSettingValue("信用区分")      '1=制度 2=一般
    Dim acctType As Long: acctType = 0    '信用＝特定口座
    
    Dim maxPos As Long: maxPos = GetSettingValue("最大同時保有数")
    Dim curPos As Long: curPos = Application.CountIf(ws.Range("G2:G" & lastRow), ">*0") '保有数量>0
    
    Dim i As Long, code As String, curP As Double
    Dim qty As Long, price As Double, orderID As Long
    
    For i = 2 To lastRow
        '―― 判定フラグ
        If ws.Cells(i, "F").Value <> "買" Then GoTo ContinueLoop
        If curPos >= maxPos Then GoTo ContinueLoop
        
        code = Format(ws.Cells(i, "A").Value, "0000")
        curP = Val(ws.Cells(i, "D").Value)
        If curP <= 0 Then GoTo ContinueLoop
        
        qty = CalcOrderQty(curP)
        If qty <= 0 Then GoTo ContinueLoop
        
        '―― 成行 or 指値
        If priceType = 0 Then
            price = 0
        Else
            price = SafeRss(code, "最良売気配値1")
            If price = 0 Then price = curP
        End If
        
        orderID = GetNextOrderID()
        
        '―― 発注
        If IsTestMode() Then
            AppendLog "TEST-BUY", code, qty, price
        Else
            On Error Resume Next
            Application.Run "RssMarginOpenOrder_v", _
                            orderID, code, qty, price, priceType, _
                            creditType, acctType, 0, 0   'SOR=0, 当日
            On Error GoTo 0
            AppendLog "BUY", code, qty, price
        End If
        
        ws.Cells(i, "F").Value = "発注中"
        curPos = curPos + 1
        
ContinueLoop:
    Next i
End Sub

'---------------------------------------------------------------
' 発注・約定ログ（sheet name: ログ）
'---------------------------------------------------------------
Sub AppendLog(act As String, code As String, qty As Long, price As Double)
    Dim wsL As Worksheet
    On Error Resume Next
    Set wsL = Worksheets("ログ")
    If wsL Is Nothing Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "ログ"
        Set wsL = Worksheets("ログ")
    End If
    On Error GoTo 0
    
    If wsL.Cells(1, "A").Value <> "時刻" Then
        wsL.[A1:E1].Value = Array("時刻", "区分", "コード", "数量", "価格")
    End If
    
    Dim n As Long: n = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
    wsL.Cells(n, 1).Resize(1, 5).Value = Array(Now, act, code, qty, price)
End Sub
'###OrdersModule_full ここまで －－－－－－－－－－－－－－－－－－－－


