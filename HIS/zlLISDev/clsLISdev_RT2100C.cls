VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_RT2100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:雷杜酶标仪（RT-2100C）
'厂商:Rayto(雷杜)
'作者:曾超
'用户:望江医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
    Dim strLine As String, varLine As Variant, var项目 As Variant, strData As String
    Dim str_标本号 As String, str_日期 As String, str_项目 As String
    Dim intRow As Integer, strTmp As String, strR As String
    Dim intLoop As Integer
    Dim aItem() As String
    Dim intCount As Integer
    Dim str医院 As String
    
    On Error GoTo errH
    aItem = Split(strSource, ",")
    
    '没有传完一个时退出
    If UBound(aItem) < 15 Then
        strReserved = strSource
        Exit Sub
    End If
    
    str医院 = ReadIni("RT2100C", "医院名称", App.Path & "\RT2100C.ini")
    
    If str医院 = "" Then
        str医院 = "默认"
    End If
    Call WriteIni("RT2100C", "医院名称", str医院, App.Path & "\RT2100C.ini")
    
    For intLoop = 0 To UBound(aItem)
        If InStr(aItem(intLoop), "R") > 0 Then
            If UBound(aItem) - intLoop > 12 Then
                If str医院 = "默认" Or str医院 = "" Then
                    '写入一个标本
                    strData = strData & "||" & aItem(intLoop + 5) & "|" & aItem(intLoop + 6) & "| |血清|0|" & aItem(intLoop + 7) & _
                    "|" & aItem(intLoop + 11) & "^" & aItem(intLoop + 12) & "^^"
                ElseIf str医院 = "望江" Then
                    '写入一个标本
                    strData = strData & "||" & aItem(intLoop + 5) & "|" & aItem(intLoop + 6) & "| |血清|0|" & aItem(intLoop + 7) & _
                    "|" & aItem(intLoop + 12)
                End If
            Else
                strReserved = strReserved & "," & aItem(intLoop)
            End If
        Else
            If UBound(aItem) - intLoop < 11 And strReserved <> "" Then
                strReserved = strReserved & "," & aItem(intLoop)
            End If
        End If
    Next
    
    strResult = Mid(strData, 3)
    strReserved = Mid(strReserved, 2)
    Exit Sub
errH:
    WriteLog "Rt2100c", Err.Description, ""
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function


