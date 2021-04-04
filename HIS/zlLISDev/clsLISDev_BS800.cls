VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BS800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BS-800,BS-400
'厂商:迈瑞
'用户:湖北省荆州市江陵县人民医院
'开发人员：张网云
'开发日期：2014-09-03
Private strDeviceString '发送信息的设备串
Private mstrCode As String
Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    On Error GoTo hErr '打开错误陷阱
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
    Dim lng_S As Long, lng_E As Long, int_Loop As Long, i As Integer
    Dim str_Code As String, str_Line As String, var_Code As Variant, strTmp As String, varTmp As Variant, strTmp2 As String
    Dim str_时间 As String, str_样本号 As String, str_检验人 As String, str_标本 As String, str_是否质控 As String, str_结果 As String
    Dim str_条码 As String
    Dim str_消息类型 As String, str_应答 As String
    Dim iResultType As Integer  '1为检验结果，2为仪器发送请求
    '--解析
    mstrCode = mstrCode & strSource
    lng_S = InStr(strSource, Chr(&HB))
    
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
    End If
    
    '数据完整
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 2)
        var_Code = Split(str_Code, Chr(&HD))
        
        If str_Code <> "" Then
            str_消息类型 = Split(var_Code(0), "|")(8)
        End If
        Select Case Mid(str_消息类型, 1, 3)
            Case "ORU"         '检验结果
              If Split(var_Code(0), "|")(15) = 2 Then '判断质控
                  iResultType = 1
                
                
                str_Line = var_Code(int_Loop)
                    'If Mid(str_Line, 2, 4) = "MSH|" Then
                    '保存本次消息相关信息
                    varTmp = Split(str_Line, "|")
                    str_消息类型 = varTmp(8)
                    str_应答 = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                    For i = 6 To UBound(varTmp)
                        strTmp2 = varTmp(i)
                        If InStr(1, strTmp2, "ORU^") > 0 Then
                            strTmp2 = Replace(strTmp2, "ORU", "ACK")
                        End If
                        '生成应答消息
                        str_应答 = str_应答 & "|" & strTmp2
                    Next
                    str_应答 = str_应答 & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                    str_应答 = str_应答 & Chr(&H1C) & Chr(&HD)
                        
                
                
                
                  str_时间 = Split(var_Code(0), "|")(6)
                  str_时间 = Mid(str_时间, 1, 4) & "-" & Mid(str_时间, 5, 2) & "-" & Mid(str_时间, 7, 2) & " " & Mid(str_时间, 9, 2) & ":" & Mid(str_时间, 11, 2) & ":00"
                  str_结果 = str_时间 & "|" & Mid(Split(var_Code(0), "|")(6), 1, 8) & "| |血清|1" & "|" & Split(var_Code(1), "|")(2) & "|" & Split(var_Code(1), "|")(20)
                  strResult = strResult & "||" & str_结果
                  Exit Do
              Else
                iResultType = 1
                For int_Loop = LBound(var_Code) To UBound(var_Code)
                    str_Line = var_Code(int_Loop)
                    If Mid(str_Line, 2, 4) = "MSH|" Then
                    '保存本次消息相关信息
                    varTmp = Split(str_Line, "|")
                    str_消息类型 = varTmp(8)
                    str_应答 = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                    For i = 6 To UBound(varTmp)
                        strTmp2 = varTmp(i)
                        If InStr(1, strTmp2, "ORU^") > 0 Then
                            strTmp2 = Replace(strTmp2, "ORU", "ACK")
                        End If
                        '生成应答消息
                        str_应答 = str_应答 & "|" & strTmp2
                    Next
                    str_应答 = str_应答 & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                    str_应答 = str_应答 & Chr(&H1C) & Chr(&HD)
                        
                    ElseIf Mid(str_Line, 1, 4) = "OBR|" Then
                        '取样本号
                        str_样本号 = Split(str_Line, "|")(3)
                        If Split(str_Line, "|")(5) = "Y" Then
                            str_样本号 = str_样本号 & "^1"
                        End If
                        If Split(str_Line, "|")(2) <> "" Then
                            If InStr(str_样本号, "^1") > 0 Then
                                str_样本号 = str_样本号 & "^" & Split(str_Line, "|")(2)
                            Else
                                str_样本号 = str_样本号 & "^^" & Split(str_Line, "|")(2)
                            End If
                            
                        End If
                        
                        '取检验时间
                        str_时间 = Split(str_Line, "|")(7)
                        If str_时间 = "" Then
                            str_时间 = Split(str_Line, "|")(14)
                        End If
                        If Len(str_时间) = 14 Then
                            str_时间 = Mid(str_时间, 1, 4) & "-" & Mid(str_时间, 5, 2) & "-" & Mid(str_时间, 7, 2) & " " & Mid(str_时间, 9, 2) & ":" & Mid(str_时间, 11, 2)
                        ElseIf Len(str_时间) = 8 Then
                            str_时间 = Mid(str_时间, 1, 4) & "-" & Mid(str_时间, 5, 2) & "-" & Mid(str_时间, 7, 2)
                        End If
                        
                    ElseIf Mid(str_Line, 1, 4) = "OBX|" Then
                        str_结果 = str_结果 & "|" & Split(str_Line, "|")(3) & "|" & Split(str_Line, "|")(5)
                    End If
                Next
                strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "| |血清|0" & str_结果
                lng_S = InStr(mstrCode, Chr(&HB))
                If lng_S > 0 Then
                    lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
                Else
                    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
                End If
             End If
            Case "QRY"         '仪器请求
                iResultType = 2
                
                For int_Loop = LBound(var_Code) To UBound(var_Code)
                    str_Line = var_Code(int_Loop)
                    If Mid(str_Line, 2, 4) = "MSH|" Then
                        '保存本次消息相关信息
                        varTmp = Split(str_Line, "|")
                        str_消息类型 = varTmp(8)
                        str_应答 = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                        For i = 6 To UBound(varTmp)
                            strTmp2 = varTmp(i)
                            If InStr(1, strTmp2, "QRY^") > 0 Then
                                strTmp2 = Replace(strTmp2, "QRY", "QCK")
                            End If
                            '生成应答消息
                            str_应答 = str_应答 & "|" & strTmp2
                        Next
                        str_应答 = str_应答 & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                        str_应答 = str_应答 & "ERR|0|" & Chr(&HD)
                        str_应答 = str_应答 & "QAK|SR|OK|" & Chr(&HD)
                        str_应答 = str_应答 & Chr(&H1C) & Chr(&HD)
                        
                    ElseIf Mid(str_Line, 1, 4) = "QRD|" Then
                        str_条码 = Split(str_Line, "|")(8)
                    ElseIf InStr(str_Line, "QRF") > 0 Then
                        str_样本号 = Split(str_Line, "|")(4)
                    End If
                Next
                strDeviceString = str_应答   '发送信息的设备串
                
                ''因为实际测试双向时，必须要传入条码号，空的条码号会产生相应的问题，为了保证后面能够正常处理，空起的条码号传入标本号
                ''因此需要在进行结果传输的时候，对条码号和标本号相同的标本进行处理
                If Mid(Trim(str_条码), 1, 8) = Format(Now, "yyyymmdd") Then
                    str_条码 = ""
                End If
                
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & str_样本号 & "^^" & Trim(str_条码)
                lng_S = InStr(mstrCode, Chr(&HB))
                If lng_S > 0 Then
                    lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
                Else
                    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
                End If
            Case Else
                lng_S = 0
        End Select
    Loop
    Select Case iResultType
            Case 1
                strCmd = "0|" & str_应答
            Case 2
                strCmd = "1|" & str_应答
            Case Else
                strCmd = "0|"
                strResult = ""
        End Select
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    '------------------------------------------------
    Exit Sub
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_BS800"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
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
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    Dim strBanCode As String '自定义条码
    Dim int_Loop As Integer
    blnSuccess = True
    
    strBanCode = Format(Now, "yyyymmdd")
    Select Case iSendStep
      '  Case 0 '开始发送
     '       iSendStep = 1
        Case Else '发送数据
            If strResponse Like "*MSH*" Or Len(Trim(strResponse)) = 0 Then '允许发送
                If strResponse Like "*MSH*" Or Len(Trim(strResponse)) = 0 Then
                    '请求结束
                    '生成设备串
                    aItem = Split(strSample, "|")
                    '标本或SampleID
                    strBuffer = aItem(1)
                    strSampleID = aItem(5)
                    '读取反馈的架号和杯号
                    strRackNO = ""
                    If strDeviceString <> "" Then
                        SendSample = Chr(&HB) & "MSH|^~\&|||Mindray|BS-800|" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "||DSR^Q03|" & Split(strDeviceString, "|")(9) & "|P|2.3.1||||||ASCII|||" & Chr(&HD)
                    
                        SendSample = SendSample & "MSA|AA|" & Split(strDeviceString, "|")(9) & "|Message accepted|||0|" & Chr(&HD)
                        
                        SendSample = SendSample & "ERR|0|" & Chr(&HD)
                        
                        SendSample = SendSample & "QAK|SR|OK|" & Chr(&HD)
                        
                        SendSample = SendSample & "QRD|" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "|R|D|1|||RD|" & strSampleID & "|OTH|||T|" & Chr(&HD)
                        
                        Dim strTimeTmp As String
                        strTimeTmp = Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "")
                        SendSample = SendSample & "QRF|BS-800|" & strTimeTmp & "|" & strTimeTmp & "|||RCT|COR|ALL||" & Chr(&HD)
                        
                        For int_Loop = 1 To 28
                            Dim strTmp As String
                            Select Case int_Loop
                                Case 3              '病人姓名
                                    strTmp = "DSP|3||" & aItem(2) & "|||" & Chr(&HD)
                                Case 21             '样本条码
                                    strTmp = "DSP|21||" & IIf(Trim(aItem(5)) = "", strBanCode & aItem(1), aItem(5)) & "|||" & Chr(&HD)
                                Case 22             '样本编号
                                    strTmp = "DSP|22||" & aItem(1) & "|||" & Chr(&HD)
                                Case 23             '送检时间
                                    strTmp = "DSP|23||" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "|||" & Chr(&HD)
                                Case 24             '是否急诊
                                    If aItem(4) = "1" Then
                                        strTmp = "DSP|24||Y|||" & Chr(&HD)
                                    Else
                                        strTmp = "DSP|24||N|||" & Chr(&HD)
                                    End If
                                Case 26
                                    strTmp = "DSP|26||" & aItem(3) & "|||" & Chr(&HD)
                                Case Else
                                    strTmp = "DSP|" & int_Loop & "|||||" & Chr(&HD)
                            End Select
                            SendSample = SendSample & strTmp
                        Next
                        
                        For i = 10 To UBound(aItem)
                            SendSample = SendSample & "DSP|" & int_Loop & "||" & aItem(i) & "^^^|||" & Chr(&HD)
                            int_Loop = int_Loop + 1
                        Next
                        SendSample = SendSample & "DSC||" & Chr(&HD) & Chr(&H1C) & Chr(&HD)
                    End If
                    iSendStep = 0
                Else
                    SendSample = Chr(6)
                    iSendStep = iSendStep + 1
                End If
            Else
                SendSample = Chr(6)
                iSendStep = 0
                blnSuccess = False '发送失败
            End If
    End Select
End Function


