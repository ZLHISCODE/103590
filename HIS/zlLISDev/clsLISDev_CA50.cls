VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:CA50
'厂商:希森美康
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
        
    Dim strCode As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim str标本号 As String, str项目 As String, str结果 As String, str日期 As String, str质控 As String
    Dim strItems As String
    
    strReserved = strSource
    
    If strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        lngS = InStr(strReserved, Chr(2))
        lngE = InStr(strReserved, Chr(3))
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
    
        str质控 = Mid(strCode, 10, 1)
        If str质控 = "C" Then
            str质控 = "1"
        Else
            str质控 = "0"
        End If
        str日期 = "20" & Mid(strCode, 15, 2) & "-" & Mid(strCode, 13, 2) & "-" & Mid(strCode, 11, 2)
        If Not IsDate(str日期) Then
            str日期 = Format(Now, "yyyy-MM-dd")
        Else
            str日期 = Format(CDate(str日期), "yyyy-MM-dd")
        End If
        
        str标本号 = Val(Mid(strCode, 32, 10))
        strCode = Mid(strCode, 54, InStr(strCode, Chr(3)) - 54)
        strItems = ""
        Do While Len(strCode) > 0
            str项目 = Mid(strCode, 1, 9)
            str结果 = Mid(str项目, 4)
            str项目 = Left(str项目, 3)
            If str项目 = "041" Then
                str结果 = Format(Val(str结果) / 10, "0.0")
            ElseIf str项目 = "043" Then
                str结果 = Format(Val(str结果) / 100, "0.00")
            ElseIf str项目 = "044" Then
                str结果 = Format(Val(str结果) / 100, "0.00")
            ElseIf str项目 = "045" Then
                str结果 = Format(Val(str结果) / 10, "0.0")
            ElseIf str项目 = "051" Then
                str结果 = Format(Val(str结果) / 10, "0.0")
            End If
            strItems = strItems & "|" & str项目 & "|" & Trim(str结果)
            strCode = Mid(strCode, 10)
            
        Loop
        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = str日期 & "|" & Val(str标本号) & "| |血液|" & str质控 & strItems
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
