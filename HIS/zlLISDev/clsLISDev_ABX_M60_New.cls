VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_M60_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ABX Micros 60
'厂商:ABX
'作者:陈东
'用户:武昌铁路医院

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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varLine As Variant, intRow As Integer, strLine As String
    Dim str标本号 As String, str日期 As String, strItems As String, lngItem As Long
    
    lngS = InStr(strSource, Chr(2))
    lngE = InStr(strSource, Chr(3))
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
        varLine = Split(strCode, Chr(&HD))
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            If InStr("!,2,3,4,5,6,7,8,@,A,B,C,#,%,',"""",$,&", Mid(strLine, 1, 1)) > 0 Then
                '项目
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid(strLine, 3))
            ElseIf Mid(strLine, 1, 1) = "q" Then
                '日期
                str日期 = Format(CDate(Mid(strLine, 9, 2) & "-" & Mid(strLine, 6, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
            ElseIf Mid(strLine, 1, 1) = "u" Then
                '标本号
                str标本号 = Val(Mid(strLine, 3))
            End If
        Next
        
        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = str日期 & "|" & str标本号 & "| |全血|0" & strItems
        End If
    End If
    
    
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function
