Attribute VB_Name = "SettingsModule"
'###SettingsModule_add �������� �|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
Option Explicit
'===============================================================================
' �ݒ�l�擾 & �y�VRSS �⏕�֐�
'===============================================================================

'---------------------------------------------------------------
' �ݒ�V�[�g (���O = "�ݒ�") �� A��=���ږ� / B��=�l ���擾
'   ���� key  : �擾���������ږ��i���S��v�j
'   �Ԓl      : ������Βl�A������� 0
'---------------------------------------------------------------
Public Function GetSettingValue(key As String) As Variant
    Dim ws As Worksheet: Set ws = Worksheets("�ݒ�")
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetSettingValue = 0
    Else
        GetSettingValue = rng.Offset(0, 1).Value
    End If
End Function

'---------------------------------------------------------------
' �y�VRSS �̃Z���֐����b�p�[
'   ��: SafeRss("7203","���ݒl") �� ���l(0���e)
'---------------------------------------------------------------
Public Function SafeRss(code As String, item As String) As Double
    On Error Resume Next
    SafeRss = Val(Application.Run("RssCell", code, item))
    On Error GoTo 0
End Function
'###SettingsModule_add �����܂� �|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|

