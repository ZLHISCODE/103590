VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Everesys_A1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'类型:化学发光测定仪
'仪器:everesys a1800 I
'厂商:苏州长光华医生物医学工程 HYBIOME
'作者:敬秦
'用户:武侯区第三人民医院
'时间:2014-12-04

Private mstrCode As String    '待解析串
Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.
Private mstrWaitSend As String '待发送标本
Private mstrQCNO As String      '存QC编号，要转为固定的标本号

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
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim strCode As String, strLine As String, strCheel As String, strValue As String, strItem As String, strArray() As String
    Dim intLoop As Integer
    Dim lngS As Long, lngE As Long
    Dim strSample As String, str急诊 As String, strQC As String, strBH  As String
    Dim strDate As String
    Dim objFileSystem As New FileSystemObject, objStream As textStream
    
    strCode = strCode & strSource
    If Not objFileSystem.FileExists(App.Path & "\" & "A1800.ini") Then '获取数据库信息
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "A1800.ini")
            objStream.WriteLine "[A1800]"
            objStream.WriteLine "STAT=3000"
            objStream.Close
            Set objStream = Nothing
    End If
    strBH = ReadIni("A1800", "STAT", App.Path & "\" & "A1800.ini")
    lngS = InStr(strCode, Chr(&H2))
    lngE = InStr(strCode, Chr(&H3))
    If lngS > 0 And lngE > 0 Then
        strArray = Split(strCode, Chr(&HD))
        For intLoop = 0 To UBound(strArray)
            strLine = strArray(intLoop)
            If InStr(strLine, "O|") > 0 Then '查找标本号和时间
                strSample = Split(Split(strLine, "|")(2), "^")(0)
                str急诊 = Split(Split(strLine, "|")(2), "^")(1)
                strDate = Split(strLine, "|")(6)
                strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
'            ElseIf InStr(strLine, "R|") > 0 Then '获取检验数据
            ElseIf strLine Like "R|*" Then
                strCheel = Split(Split(strLine, "|")(2), "^")(1)
                strValue = Replace(Split(strLine, "|")(3), "^", "")
                strItem = strItem & "|" & strCheel & "|" & strValue
            End If
        Next
        If strSample <> "" And strItem <> "" Then
            strResult = strDate & "|" & IIf(str急诊 = "0", Val(strBH) + Val(strSample) & "^1^", strSample) & "| |血清|0" & strItem
            Call writelislog(strResult)
        End If
    End If
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_A1800_"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function
