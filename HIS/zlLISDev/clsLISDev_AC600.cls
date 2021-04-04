VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AC600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AC600 糖化血红蛋白仪
'厂商:
'作者:敬秦
'用户:川大望江校医院

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
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim strCode As String
    Dim i As Integer, lngE As Long, lngS As Long
    Dim strDate As String, strSample As String, StrQ As String, Items As String, strE As String
    

    If strSource <> "" Then
        strReserved = strReserved & strSource
    End If

    On Error GoTo errHand
    
    Call writelislog("收到的数据：" & strReserved)
    lngE = InStr(strReserved, Chr(&H2))
    lngS = InStr(strReserved, Chr(&H3))
    If lngE > 0 And lngS > 0 Then
        strCode = strReserved
        strReserved = Mid(strReserved, lngS + 1)
        strDate = Format(Mid(strCode, 2, 19), "yyyy-mm-dd hh:mm:ss")
        StrQ = Mid(strCode, 21, 15)

        strSample = Val(Right(Split(StrQ, "-")(1), 4))
        StrQ = Trim(Left(Split(StrQ, "-")(1), 3))
        If StrQ = "ST" Then
            strE = "1"
        ElseIf StrQ = "CAL" Then
            StrQ = "1"
            strE = "0"
        Else
            StrQ = "0"
            strE = "0"
        End If
        Items = Items & "|HbA1c|" & Val(Mid(Mid(strCode, 36, 26), 9, 6))
        Items = Items & "|IFCC|" & Val(Mid(Mid(strCode, 62, 23), 16, 4))
        Items = Items & "|NGSP|" & Val(Mid(Mid(strCode, 85), 26, 5))
        If Items <> "" Then
            strResult = strDate & "|" & Val(strSample) & IIf(strE = "1", "^1^", "") & "| |血清|" & IIf(StrQ = "1", "1", "0") & Items
        End If
    End If
    
    Exit Sub
errHand:
    WriteLog "clsLISDev_AC600", Err.Description, "待解析结果：" & strCode
End Sub
Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_AC600_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function
