Attribute VB_Name = "SettingsModule"
'###SettingsModule_add ここから −−−−−−−−−−−−−−−−−−−−
Option Explicit
'===============================================================================
' 設定値取得 & 楽天RSS 補助関数
'===============================================================================

'---------------------------------------------------------------
' 設定シート (名前 = "設定") の A列=項目名 / B列=値 を取得
'   引数 key  : 取得したい項目名（完全一致）
'   返値      : 見つかれば値、無ければ 0
'---------------------------------------------------------------
Public Function GetSettingValue(key As String) As Variant
    Dim ws As Worksheet: Set ws = Worksheets("設定")
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetSettingValue = 0
    Else
        GetSettingValue = rng.Offset(0, 1).Value
    End If
End Function

'---------------------------------------------------------------
' 楽天RSS のセル関数ラッパー
'   例: SafeRss("7203","現在値") → 数値(0許容)
'---------------------------------------------------------------
Public Function SafeRss(code As String, item As String) As Double
    On Error Resume Next
    SafeRss = Val(Application.Run("RssCell", code, item))
    On Error GoTo 0
End Function
'###SettingsModule_add ここまで −−−−−−−−−−−−−−−−−−−−

