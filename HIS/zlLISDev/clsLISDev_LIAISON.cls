VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LIAISON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器名称: LIAISON 化学发光
'仪器型号: LBYN7500B
'厂商:普利生
'医院:四川省第五人民医院
'作者:敬秦
'日期:2014-10-08

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
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
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
    Dim lng_S As Long, lng_E As Long, intLoop As Integer
    Dim strCode As String, strSample As String, strDate As String, strLine() As String, strValue As String
    Dim strBuffer As String
    Dim i As Integer

    
    strReserved = strReserved & strSource
    Call writelislog("接收：" & strReserved)

    If InStr(strReserved, Chr(&H5)) > 0 Then                                         '收到结束符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&HA) & Chr(&HD)) > 0 Then                              '收到行结束符 03 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&H3)) > 0 Then                                         '收到开始符 03 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&H4)) > 0 Then                                         '收到开始符 04 ,应答 06
        strCmd = Chr(&H6)
    End If
    
    lng_S = InStr(strReserved, Chr(&H5)) '开始符
    lng_E = InStr(strReserved, Chr(&H4)) '结束符
    i = 1
    If lng_S > 0 And lng_E > 0 Then
        Do Until strReserved = ""
            lng_S = InStr(strReserved, Chr(&H5)) '开始符
            lng_E = InStr(strReserved, Chr(&H4)) '结束符
            strBuffer = Mid$(strReserved, lng_S + 1, lng_E + lng_S - 1)
            Call writelislog("处理：" & strBuffer)
            strReserved = Mid$(strReserved, lng_E + 1) '下次处理
            Call writelislog("后续处理：" & strReserved)
            strLine = Split(strBuffer, Chr(&HD))
            For intLoop = LBound(strLine) To UBound(strLine)
                If InStr(strLine(intLoop), "4R|") > 0 And i = 1 Then
                    strDate = Split(strLine(intLoop), "|")(12)
                    strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
                    i = i + 1
                End If
                If InStr(strLine(intLoop), "O|") > 0 Then
                    strSample = Split(strLine(intLoop), "|")(2)
                End If
                If InStr(strLine(intLoop), "Liaison") > 0 And InStr(strLine(intLoop), "1H") = 0 Then
                    strValue = strValue & "|" & Replace(Split(strLine(intLoop), "|")(2), "^", "") & "|" & Split(strLine(intLoop), "|")(3)
                End If
            Next
            If Val(strSample) <> 0 And strValue <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSample & "| |血清|0" & strValue
            End If
        Loop
        If Left(strResult, 2) = "||" Then
            strResult = Mid(strResult, 3)
        End If
    End If
    Exit Sub

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_LIAISON"
    Close #lngErrFileNo
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_LIAISON_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFilename As String
    strFilename = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function
