Attribute VB_Name = "mdlPUBLIC"
Option Explicit

Public Const G_UA_PWD           As String = "FA74C8A530DE7E088B1ACA673DD6297D"
Public Const G_UA_KEY           As String = "0016FDE250354FA9A4BA45433DBCC35D"
Public Const G_INTERFACE_KEY    As String = "EBA1D9B8CCCB4FD0804672DEDB222CFB"
Public Const G_APP_KEY          As String = "FD304782E75C41FDB14CB7A92A8A0B97"

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
'       ʵ�����ݴ洢����
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function
