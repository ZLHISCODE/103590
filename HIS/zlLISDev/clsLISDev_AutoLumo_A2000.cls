VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AutoLumo_A2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AutoLumo A2000 酶标仪 按测试发送数据
'厂商:安图生物工程有限公司
'作者:敬秦
'用户:双流中医院
'日期:2014-07-24
Private mstrCode As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
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

    Dim strCode As String, STRRESULTS() As String, strValue() As String, strLine As String
    Dim intLoop As Integer, strOD As Double, strSCO As Double, strCutOff As Double
    Dim intID As Long, strSample As String, strNumber As String, strCheal As String
    Dim intSerial As Integer, lngS As Long, lngE As Long

    strCode = strCode & strSource
    strReserved = strReserved & strSource
    lngS = InStr(strReserved, "{")
    lngE = InStr(strReserved, "}")
    Call writelislog("收到：" & strReserved)
    '    If lngS = 0 Or lngE = 0 Then
    '        Exit Sub
    '    End If
    If lngS > 0 And lngE > 0 Then
        If Len(strReserved) = 0 Then Exit Sub
        strValue = Split(strReserved, "}")
        For intLoop = 0 To UBound(strValue)
            strLine = strValue(intLoop)
            If strLine <> "" Then
                strLine = Replace(strLine, "{", "")
                intSerial = Val(Split(strLine, ",")(0)) '序号
                If intSerial = 5 Then '普通数据
                    strSample = Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) '标本号
                    strCheal = Mid(Split(strLine, ",")(3), InStr(Split(strLine, ",")(3), "]") + 1) '通道号
                    strNumber = Split(strLine, ",")(4) '病历号
                    strOD = Val(Split(strLine, ",")(5))
                    strSCO = Val(Split(strLine, ",")(6))
                    strCutOff = Format(strOD / IIf(strSCO = 0, 1, strSCO), "0.000")
                    '                    strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & strSample & "|" & " |血液|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    If Len(strSample) = 12 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|^^" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) & "|" & " |血液|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    Else
                        strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) & "^^|" & " |血液|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    End If
                    If intSerial = 5 Then '测试发送样本
                        strCmd = strCmd & "{6,0," & strNumber & ",[S]" & strSample & ",[S],[S]" & strCheal & "}"
                    End If
                End If
                If intSerial = 9 Then '开启双向
                    strResult = Format(Now, "yyyy-mm-dd") & "|^^" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1)
                    strCmd = "1|"
                    Call writelislog("strResult:" & strResult)
                End If
            End If
        Next
        strReserved = ""
    End If
    If Left(strResult, 2) = "||" Then
        strResult = Mid(strResult, 3)
    End If
    Exit Sub

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_AutoLumo_A2000"
    Close #lngErrFileNo
End Sub
Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    ''   strSample:解析需要向仪器发送的标本信息
    
    '解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String, strSampleNO As String, strItem As String, sampleno As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    
    blnSuccess = True: SendSample = ""
    Call writelislog("strSample:" & strSample)
    Call writelislog("strResponse:" & strResponse)
    blnSuccess = True
    aItem = Split(strSample, "|")
    If strResponse Like "{*[S]*}" Then
        strSampleID = Mid(Split(strResponse, ",")(2), InStr(Split(strResponse, ",")(2), "]") + 1, 12)
        strSampleNO = Mid(Split(strResponse, ",")(0), InStr(Split(strResponse, ",")(0), "{") + 1)
        sampleno = aItem(1)
        If Len(strSampleID) = 12 And Val(strSampleNO) = 9 Then '组织指令
            For i = 10 To UBound(aItem)
'                SendSample = "{10,0,[S]" & strSampleID & ",1,0,[S],3,"
                strItem = strItem & "[S]" & aItem(i) & ",1,"
            Next
            If Len(strItem) > 0 Then
                strItem = Mid(strItem, 1, Len(strItem) - 1)
                SendSample = "{10,0,[S]" & strSampleID & ",1,0,[S]," & Val(UBound(aItem) - 9) & "," & strItem & "}"
'                SendSample = "{10,0,[S]" & sampleno & ",1,0,[S]," & "3" & "," & strItem & "}"
                Call writelislog("SendSample:" & SendSample)
            End If
            iSendStep = 0
        End If
   End If
End Function
Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_A2000"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine ("--------------------------------------------------------------------------------------------------------------------------------------------------------------")
    objStream.WriteLine (strLog)
    'objStream.WriteLine ("--------------------------------------------------------------------------------------------------------------------------------------------------------------")
    objStream.Close
    Set objStream = Nothing
End Function
