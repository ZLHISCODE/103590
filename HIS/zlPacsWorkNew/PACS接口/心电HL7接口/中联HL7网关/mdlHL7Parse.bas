Attribute VB_Name = "mdlHL7Parse"
Option Explicit

Public Function GetMsgACK(strData As String, strACK As String) As Boolean
'-----------------------------------------------------------------------------
'功能:从HL7消息中，提取信息，组织成响应消息
'参数:
'       strData 【IN】  ---消息文本
'       strACK  【OUT】---- ACK消息
'返回：True -- 成功接收；False -- 出错拒绝
'-----------------------------------------------------------------------------
    Dim strField() As String
    Dim i As Integer
    Dim strSendingApp As String
    Dim strSendingFac As String
    Dim strMsgControlID As String
    Dim strErrMsg As String
    Dim strMSH As String
    Dim strMSA As String
    Dim lngMsgFullType As Long
    
    On Error GoTo err
    
    'ACK消息，从原消息中提取MSH-3，MSH-4，MSH-10，其他的是固定值
    'MSH-3  ，发送程序名称
    'MSH-4  ，发送设备名称
    'MSH-10 ，消息控制ID
    
    '检查消息的完整性,'0 -- 是完整消息；1 -- 是消息头；2 -- 是消息尾；3 -- 是消息中间段；4 -- 错误
    lngMsgFullType = funMsgFullType(strData)
    
    '记录处理日志
    Call WriteProcessLog("GetMsgACK", "组织响应消息", "需要响应的消息是：" & strData & vbCrLf & "这个消息的完整性是：" & lngMsgFullType, 2)
    
    If lngMsgFullType = 0 Or lngMsgFullType = 1 Then
        '完整消息或者消息头，进行解析
        '提取原消息中的MSH-3，MSH-4，MSH-10
        '因为只是解析消息中的MSH段，因此直接使用“|”作为分隔符
        strField = Split(strData, "|")
        If Trim(strField(0)) = Chr(11) & "MSH" Then
            If UBound(strField) > 10 Then
                strSendingApp = strField(2)
                strSendingFac = strField(3)
                strMsgControlID = strField(9)
                
                '组织ACK返回消息
                strMSH = Chr(11) & "MSH|^~\&|ZLHIS|HIS001|" & strSendingApp & "|" & strSendingFac & "|" & _
                            getDateTimeString(Now) & "||ACK|" & getDateTimeString(Now) & "|P|2.4|" & Chr(13)
                If lngMsgFullType = 0 Then
                    strMSA = "MSA|AA|" & strMsgControlID & "|" & Chr(28) & Chr(13)
                    GetMsgACK = True
                Else
                    strMSA = "MSA|AR|" & strMsgControlID & "|" & Chr(28) & Chr(13)
                    GetMsgACK = False
                End If
                strACK = strMSH & strMSA
                '组织好ACK响应，直接退出
                Exit Function
            Else
                'MSH段消息不够
                strErrMsg = "MSH段中找不到消息控制ID"
            End If
        Else
            'MSH段错误
            strErrMsg = "找不到MSH段"
        End If
    Else
        strErrMsg = "消息不完整"
    End If
    
    '其他错误，单独组织一个MSH消息头
    strMSH = Chr(11) & "MSH|^~\&|ZLHIS|HIS001|SendingApp|SendingFac|" & _
                            getDateTimeString(Now) & "||ACK|" & getDateTimeString(Now) & "|P|2.4|" & Chr(13)
    strMSA = "MSA|AR||" & strErrMsg & Chr(28) & Chr(13)
    strACK = strMSH & strMSA
    GetMsgACK = False
    
    Exit Function
err:
    '不处理，返回空消息
    '记录错误日志
    Call WriteLog(1001, err.Number, "产生ACK响应出错，待处理的消息是： " & strData & "。错误描述是：" & err.Description)
End Function

Public Function getDateTimeString(strDateTime As String) As String
'-----------------------------------------------------------------------------
'功能:返回格式化好的时间字符串，到毫秒级别
'参数:
'       strDateTime  ---日期时间文本
'返回：格式化的文本
'-----------------------------------------------------------------------------
    
    getDateTimeString = Format(strDateTime, "YYYYMMDDHHMMSS")
    
End Function



Public Function getXPN(strXPN As String) As String
'-----------------------------------------------------------------------------
'功能:从PN或者XPN字段中，读取姓名
'参数:
'       strXPN  ---接收到的PN或者XPN类型的字符串
'返回：姓名
'-----------------------------------------------------------------------------
    Dim arrName() As String
    Dim strName As String
    Dim i As Integer
    
    '格式“张^三”
    On Error GoTo err
    arrName = Split(strXPN, "^")
    
    If UBound(arrName) >= 0 Then
        strName = arrName(0)
        For i = 1 To UBound(arrName) - 1
            strName = strName & arrName(i)
        Next i
    End If
    
    Exit Function
err:
    '记录错误日志
    Call WriteLog(1002, err.Number, "getXPN,解析姓名出错，待解析的姓名是： " & strXPN & "。错误描述是：" & err.Description)
End Function

Public Function getCMName(strCMName As String) As String
'-----------------------------------------------------------------------------
'功能:从CM类型的字段中，提取姓名，适用于OBR-32，OBR-33
'参数:
'       strCMName  ---接收到的CM类型的姓名字符串
'返回：姓名
'-----------------------------------------------------------------------------
    Dim arrName() As String
    Dim strName As String
    Dim i As Integer
    Dim iEnd As Integer
    
    '格式“^张^三^^^^^^^^1^”,"李四^^^^^^^^^^^"
    
    On Error GoTo err
    arrName = Split(strCMName, "^")
    iEnd = UBound(arrName) - 1
    If iEnd > 4 Then iEnd = 4
    
    If UBound(arrName) >= 1 Then
        strName = arrName(0)
        For i = 1 To iEnd
            strName = strName & arrName(i)
        Next i
        getCMName = strName
    End If
    
    Exit Function
err:
    '记录错误日志
    Call WriteLog(1003, err.Number, "getCMName,解析姓名出错，待解析的姓名是： " & strCMName & "。错误描述是：" & err.Description)

End Function

Public Function funParseInMsg(strMsg As String) As Long
'-----------------------------------------------------------------------------
'功能:解析和处理接收到的HL7消息
'参数:  strMsg -- 消息文本
'返回： 0 -- 成功；1 -- 失败,消息类型不支持
'-----------------------------------------------------------------------------
    Dim strSegments() As String
    Dim strFields() As String
    Dim i As Integer
    Dim strMsgType As String
    Dim strPatientID As String
    Dim strOrderID As String
    Dim strDoctor As String
    Dim strResultURL As String
    Dim strResultDiag As String
    Dim strSQL As String
    
    On Error GoTo err
    
    '暂时只处理ORU-R01消息
    
    '将消息按照回车分段
    strSegments = Split(strMsg, Chr(13))
    
    '根据段标志，循环分析每一段的消息，提取其中的信息
    For i = 0 To UBound(strSegments) - 1
        strFields = Split(strSegments(i), "|")
        
        If UBound(strFields) > -1 Then
            If Trim(strFields(0)) = Chr(11) & "MSH" Then
                'MSH段
                '提取MSH-9，消息类型，判断是否 “ORU-R01”消息
                If UBound(strFields) >= 8 Then
                    strMsgType = strFields(8)
                    If strMsgType <> "ORU^R01" Then
                        Call WriteProcessLog("funParseInMsg", "消息类型不支持", "消息类型是：" & strMsgType & "，服务程序无法解析这条消息。", 2)
                        funParseInMsg = 1   '消息类型不支持
                        Exit Function
                    End If
                End If
            ElseIf strFields(0) = vbLf & "PID" Or strFields(0) = "PID" Then
                'PID段
                '提取PID-2,Patient ID 患者ID 作为病人ID的判断条件
                If UBound(strFields) >= 2 Then
                    strPatientID = strFields(2)
                End If
            ElseIf strFields(0) = vbLf & "OBR" Or strFields(0) = "OBR" Then
                'OBR段
                '提取OBR-2，开单者医嘱号码,医嘱ID，作为查询并记录结果的索引
                '提取OBR-32，结果主要负责人+，作为执行费用操作员名称
                If UBound(strFields) >= 32 Then
                    strOrderID = strFields(2)
                    strDoctor = getCMName(strFields(32))
                End If
            ElseIf strFields(0) = vbLf & "OBX" Or strFields(0) = "OBX" Then
                'OBX段
                'OBX-11"观察结果状态"，缺省值“F”表示正常，其他值有可能表示结果需要被更新或者替换。不判断这个值，所有结果直接替换。
                
                '提取心电返回的URL结果连接
                '提取OBX-2，Value Type 值类型,类型值=“RP”的是返回的URL连接
                '提取OBX-3，Observation Identifier 观察标识符,标识符=“MUSEWebURL”的是URL连接
                '提取OBX-5，Observation Value 观察值，观察值的内容就是URL连接
                'URL连接需要注意，服务器名称可能是机器名，需要转成IP地址，链接串中的\T\需要转义成&
                
                '提取心电反馈的报告描述
                'OBX|88|FT|ECGMEASANDDIAG||Test Reason : ~Blood Pressure : ***/*** mmHG~Vent. Rate : 079 BPM     Atrial Rate : 079 BPM~   P-R Int : 150 ms          QRS Dur : 086 ms~    QT Int : 394 ms       P-R-T Axes : 065 013 034 degrees~   QTc Int : 451 ms~~窦性心律 ~~Referred By:             //心电图显示的检测的参数
                'Overread By: 勇娟 郭||||||D|        //下医嘱医生
                '提取OBX-2，Value Type 值类型,类型值=“FT”的是返回的报告描述
                '提取OBX-3，Observation Identifier 观察标识符,标识符=“ECGMEASANDDIAG”的是报告描述
                '提取OBX-5，Observation Value 观察值，观察值的内容就是报告描述，整个值作为一个文本解析，使用回车替换“~”符号，保存到报告单的“所见”中。
                
                If UBound(strFields) >= 5 Then
                    '检查结果URL
                    If strFields(2) = "RP" And strFields(3) = "MUSEWebURL" Then
                        strResultURL = strFields(5)
                        strResultURL = Replace(strResultURL, "\T\", "&")
                        strResultURL = Replace(strResultURL, "'", "‘")
                    End If
                    
                    'ECG诊断
                    If strFields(2) = "FT" And strFields(3) = "ECGMEASANDDIAG" Then
                        strResultDiag = strFields(5)
                        '解析诊断内容，使用回车替换“~”符号
                        strResultDiag = Replace(strResultDiag, "~", vbCrLf)
                        '放置内容串出错，使用双字节的“‘”代替单字节的"'"
                        strResultDiag = Replace(strResultDiag, "'", "’")
                    End If
                End If
            End If
        End If
    Next i
    
    '检查提取的信息是否正确，正确则保存到数据库中
    If strResultURL <> "" Then
        '心电图的检查结果，保存到“病人医嘱发送.执行说明”中。
        strSQL = "zlhis.b_Hl7interface.Recevieresult(" & strOrderID & ", '" & strDoctor & "','" & strResultURL & "')"
        
        '记录处理日志
        Call WriteProcessLog("funParseInMsg", "准备保存心电检查结果", "调用存储过程 =" & strSQL, 3)
        
        gzlDatabase.ExecuteProcedure strSQL, "接收到心电检查结果"
                
        '记录消息记录
        Call WriteMessageLog("接收心电检查结果", "医嘱ID = " & strOrderID & "，检查医生=" & strDoctor & "，结果链接=" & strResultURL)
    End If
    If strResultDiag <> "" Then
        '心电诊断描述，保存到报告单的“所见”提纲中
        strSQL = "zlhis.b_Hl7interface.SendReport(" & strOrderID & ",'" & strResultDiag & "',NULL,'" & strDoctor & "')"
        
        '记录处理日志
        Call WriteProcessLog("funParseInMsg", "准备保存心电报告描述", "调用存储过程 =" & strSQL, 3)
        
        gzlDatabase.ExecuteProcedure strSQL, "接收心电报告描述"
        
        '记录消息记录
        Call WriteMessageLog("接收心电报告描述", "医嘱ID = " & strOrderID & "，检查医生=" & strDoctor & "，报告描述=" & strResultDiag)
    End If
    
    
    Exit Function
err:
    '记录错误日志
    Call WriteLog(1004, err.Number, "funParseInMsg解析消息出错，待解析的前半段消息是： " & Left(strMsg, 250) & "。错误描述是：" & err.Description)
    Call WriteLog(1004, err.Number, "funParseInMsg解析消息出错，医嘱ID = " & strOrderID & "，检查医生=" & strDoctor & "，结果链接=" & strResultURL & "，报告描述=" & strResultDiag)
End Function

Public Function funParseACK(strMsg As String, strACK As String) As Long
'-----------------------------------------------------------------------------
'功能:解析和处理接收到的ACK消息，通过消息控制ID判断ACK是否正确接收了
'参数:  strMsg -- 发送的消息文本
'       strACK -- 接收到的ACK文本
'返回： 0 -- 成功；1 -- 失败,发送的消息不正确;2 -- 接收到的不是ACK消息;3 -- 收到ACK消息，但是没有被对方接收
'-----------------------------------------------------------------------------
    Dim i As Integer
    Dim strSegments() As String
    Dim strFields() As String
    Dim strMsgControlID As String
    Dim blnSendMsgOK As Boolean
    Dim blnIsACK As Boolean
    Dim blnACKOK As Boolean
    
    On Error GoTo err
    
    '将消息按照回车分段
    strSegments = Split(strMsg, Chr(13))
    
    '提取消息控制ID
    If UBound(strSegments) <> -1 Then
        strFields = Split(strSegments(i), "|")
        If UBound(strFields) > 10 Then
            If strFields(0) = Chr(11) & "MSH" Then
                strMsgControlID = strFields(9)
                blnSendMsgOK = True
            End If
        End If
    End If
    
    If blnSendMsgOK = True Then
        '解析ACK消息
        strSegments = Split(strACK, Chr(13))
        
        For i = 0 To UBound(strSegments) - 1
            strFields = Split(strSegments(i), "|")
            If UBound(strFields) > -1 Then
                If Trim(strFields(0)) = Chr(11) & "MSH" Then
                    'MSH段，MSG-9,消息类型
                    If UBound(strFields) > 8 Then
                        If strFields(8) = "ACK" Then
                            blnIsACK = True
                        Else
                            Exit For
                        End If
                    End If
                ElseIf strFields(0) = vbLf & "MSA" Or strFields(0) = "MSA" Then
                    'MSA段，MSA-1 确认代码；MSA-2 消息控制ID
                    If UBound(strFields) >= 2 Then
                        If strFields(1) = "AA" And (strFields(2) = strMsgControlID Or strFields(2) = strMsgControlID & Chr(28)) Then
                            'AA表示正常的ACK
                            blnACKOK = True
                        ElseIf strFields(1) = "AE" And (strFields(2) = strMsgControlID Or strFields(2) = strMsgControlID & Chr(28)) And (UCase(strFields(3)) = UCase("Duplicate Order Record")) Then
                            'AE表示出错的ACK，但是如果错误描述是Duplicate Order Record，说明这个医嘱已经发送成功了，算是发送成功，不用重新发送了。
                            blnACKOK = True
                        End If
                    End If
                End If
            End If
        Next i
        
        If blnIsACK = True Then
            If blnACKOK = True Then
                '正常接收
                funParseACK = 0
            Else
                '接收出现错误
                Call WriteLog(1005, err.Number, "funParseACK，接收到ACK消息，但是没有成功，ACK消息是： " & strACK)
                funParseACK = 3
                Exit Function
            End If
        Else
            '接收到的消息不是ACK消息
            Call WriteLog(1006, err.Number, "funParseACK，接收到的消息不是ACK消息，该消息是： " & strACK)
            funParseACK = 2
            Exit Function
        End If
        
    Else
        '发送的消息本身就不正确，记录错误日志
        Call WriteLog(1007, err.Number, "funParseACK，发送的消息格式不正确，发送的消息是： " & strMsg)
        '返回错误信息
        funParseACK = 1
        Exit Function
    End If
    
    Exit Function
err:
    '记录错误日志
    Call WriteLog(1008, err.Number, "funParseACK,解析消息出错，待解析的消息是： " & strMsg & "。错误描述是：" & err.Description)
End Function
