Attribute VB_Name = "SettingsModule"
'###SettingsModule_add ‚±‚±‚©‚ç ||||||||||||||||||||
Option Explicit
'===============================================================================
' İ’è’læ“¾ & Šy“VRSS •â•ŠÖ”
'===============================================================================

'---------------------------------------------------------------
' İ’èƒV[ƒg (–¼‘O = "İ’è") ‚Ì A—ñ=€–Ú–¼ / B—ñ=’l ‚ğæ“¾
'   ˆø” key  : æ“¾‚µ‚½‚¢€–Ú–¼iŠ®‘Sˆê’vj
'   •Ô’l      : Œ©‚Â‚©‚ê‚Î’lA–³‚¯‚ê‚Î 0
'---------------------------------------------------------------
Public Function GetSettingValue(key As String) As Variant
    Dim ws As Worksheet: Set ws = Worksheets("İ’è")
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetSettingValue = 0
    Else
        GetSettingValue = rng.Offset(0, 1).Value
    End If
End Function

'---------------------------------------------------------------
' Šy“VRSS ‚ÌƒZƒ‹ŠÖ”ƒ‰ƒbƒp[
'   —á: SafeRss("7203","Œ»İ’l") ¨ ”’l(0‹–—e)
'---------------------------------------------------------------
Public Function SafeRss(code As String, item As String) As Double
    On Error Resume Next
    SafeRss = Val(Application.Run("RssCell", code, item))
    On Error GoTo 0
End Function
'###SettingsModule_add ‚±‚±‚Ü‚Å ||||||||||||||||||||

