Attribute VB_Name = "mdlPUBLIC"
Option Explicit

Public Const G_UA_PWD           As String = "FA74C8A530DE7E088B1ACA673DD6297D"
Public Const G_UA_KEY           As String = "0016FDE250354FA9A4BA45433DBCC35D"
Public Const G_INTERFACE_KEY    As String = "EBA1D9B8CCCB4FD0804672DEDB222CFB"
Public Const G_APP_KEY          As String = "FD304782E75C41FDB14CB7A92A8A0B97"

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
'       实际数据存储长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
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
