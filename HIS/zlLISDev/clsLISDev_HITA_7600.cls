VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串

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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnMod2 As Boolean '模式2在标本号前加01
    Dim blnQc As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode 、2:New Mode
    Dim iResultType As Integer '结果类型：1－检验结果、2－仪器请求
    Dim i As Long
    Dim strACK As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 0
    
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&HD) & Chr(&HA)) > 0 Or InStr(strSource, Chr(&H4)) > 0 Then
        strACK = Chr(&H6)
    End If
    
    iStrPoint = 1: blnReadItem = False
    '取消请求符
    strSource = Replace(strSource, Chr(5), "")
    '取消结束符
    strSource = Replace(strSource, Chr(4), "")
    '取消开始符
    strSource = Replace(strSource, Chr(2), "")
    '取消回车换行符
    strSource = Replace(strSource, vbCrLf, "")
    '取消换行符
    strSource = Replace(strSource, Chr(10), "")
    '取消制表符
    strSource = Replace(strSource, Chr(9), "")
    '取消通信连接符
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(23))
        If i = 0 Then Exit Do
        If i + 3 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 4)
        Else
            Exit Do
        End If
    Loop
    
    Do While iStrPoint <= Len(strSource)
        '查找信息的开始位置：以H开始
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Mid(strBuffer, 1, 2) Like "*H*" Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint
            
            '记录分隔符
            If Mid(strBuffer, 1, 2) Like "*H" Then
                strFldDelim = Mid(strBuffer, 3, 1) '字段分隔符
                strRepeatDelim = Mid(strBuffer, 4, 1) '重复内容分隔符
                strCompDelim = Mid(strBuffer, 5, 1) '组间内容分隔符
            Else
                strFldDelim = Mid(strBuffer, 2, 1) '字段分隔符
                strRepeatDelim = Mid(strBuffer, 3, 1) '重复内容分隔符
                strCompDelim = Mid(strBuffer, 4, 1) '组间内容分隔符
            End If
            
            '数据是否完整
            If InStr(iStrPoint, strSource, vbCr & "L" & strFldDelim) > 0 Then
                aItemData = Split(strBuffer, strFldDelim)
                Select Case Split(aItemData(10), strCompDelim)(0)
                    Case "TSREQ" '仪器请求申请信息
                        '如果当前是返回结果，则不再处理请求，保留内容直接返回
                        If iResultType = 1 Then Exit Do
                    
                        '信息完整，不再保留
                        iCurrGrpEntry = 0
                        iResultType = 2
                        
                        '查找标本信息
                        iStrPoint = iPos
                        Do While iStrPoint <= Len(strSource)
                            iPos = iStrPoint
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 1, 1) = "Q" Then
                                aItemData = Split(strBuffer, strFldDelim)
                                If InStr(strBuffer, strFldDelim & "ALL" & strFldDelim) > 0 Then
                                    iMode = 2 'New Mode
                                    '检验时间
                                    '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                            
                                    '样本序号，必须为数字
                                    strData = strData & "|" & Val(Split(aItemData(2), strCompDelim)(3))
                                            
                                    '检验人
                                    strData = strData & "| " '本仪器不支持
                                    '标本
                                    strData = strData & "|血液" '忽略
                                    '质控品
                                    strData = strData & "|0" '忽略
                                Else
                                    iMode = 1 'Current Mode
                                    '检验时间
                                    '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                            
                                    '样本序号，必须为数字
                                    strData = strData & "|" & Val(Split(aItemData(2), strCompDelim)(2))
                                            
                                    '检验人
                                    strData = strData & "| " '本仪器不支持
                                    '标本
                                    strData = strData & "|血液" '忽略
                                    '质控品
                                    strData = strData & "|0" '忽略
                                End If
                                '一个标本结束
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                            Else
                                If Mid(strBuffer, 1, 1) <> "P" Then
                                    iStrPoint = iPos
                                    Exit Do
                                End If
                            End If
                        Loop
                    Case "RSUPL" '仪器结果上传
                        '如果当前是返回结果，则不处理请求，保留内容直接返回
                        If iResultType = 2 Then Exit Do
                    
                        '信息完整，不再保留
                        iCurrGrpEntry = 0
                        iResultType = 1
                        
                        '查找标本信息
                        iStrPoint = iPos: blnReadItem = False
                        Do While iStrPoint <= Len(strSource)
                            If Not blnReadItem Then '当前是否正在读取检验项目
                                Do While iStrPoint <= Len(strSource)
                                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                    If Mid(strBuffer, 1, 1) = "O" Then Exit Do
                                Loop
                                If iStrPoint > Len(strSource) Then
                                    '当前信息无效，不再处理。
                                    Exit Do
                                End If
                                
                                aItemData = Split(strBuffer, strFldDelim)
                                If InStr(aItemData(3), strCompDelim) > 0 Then
                                    iMode = 2 'New Mode
                                    '检验时间
                                    '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                    strData = Mid(aItemData(22), 1, 4) & "-" & Mid(aItemData(22), 5, 2) & "-" & _
                                        Mid(aItemData(22), 7, 2) & " " & Mid(aItemData(22), 9, 2) & ":" & _
                                        Mid(aItemData(22), 11, 2) & ":00" ' & Mid(aItemData(22), 13, 2)
                                            
                                    '样本序号，必须为数字
                                    strData = strData & "|" & Val(Split(aItemData(3), strCompDelim)(0))
                                    If UCase(aItemData(5)) = "S" Then '急诊
                                        strData = strData & "^1"
                                    End If
                                            
                                    '检验人
                                    strData = strData & "| " '本仪器不支持
                                    '标本
                                    Select Case Split(aItemData(3), strCompDelim)(4)
                                        Case "S1"
                                            strData = strData & "|血液"
                                        Case "S2"
                                            strData = strData & "|尿液"
                                        Case "S3"
                                            strData = strData & "|脊髓"
                                        Case Else
                                            strData = strData & "|体液"
                                    End Select
                                    '质控品
                                    strData = strData & IIf(aItemData(11) = "Q", "|1", "|0")
                                Else
                                    iMode = 1 'Current Mode
                                    '检验时间
                                    '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                    strData = Mid(aItemData(22), 1, 4) & "-" & Mid(aItemData(22), 5, 2) & "-" & _
                                        Mid(aItemData(22), 7, 2) & " " & Mid(aItemData(22), 9, 2) & ":" & _
                                        Mid(aItemData(22), 11, 2) & ":00" ' & Mid(aItemData(22), 13, 2)
                                            
                                    '样本序号，必须为数字
                                    strData = strData & "|" & Val(Split(aItemData(2), strCompDelim)(0))
                                    If UCase(aItemData(5)) = "S" Then '急诊
                                        strData = strData & "^1"
                                    End If
                                            
                                    '检验人
                                    strData = strData & "| " '本仪器不支持
                                    '标本
                                    Select Case Val(Split(aItemData(2), strCompDelim)(2))
                                        Case 1
                                            strData = strData & "|血液"
                                        Case 2
                                            strData = strData & "|尿液"
                                        Case 3
                                            strData = strData & "|脊髓"
                                        Case Else
                                            strData = strData & "|体液"
                                    End Select
                                    '质控品
                                    strData = strData & IIf(aItemData(11) = "Q", "|1", "|0")
                                End If
                                
                                '准备接收
                                blnReadItem = True
                            Else
                                Do While iStrPoint <= Len(strSource)
                                    iPos = iStrPoint '保留当前位置
                                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                    If Mid(strBuffer, 1, 1) = "R" Then
                                    '检验结果
                                        aItemData = Split(strBuffer, strFldDelim)
                                        strData = strData & "|" & Val(Mid(Split(aItemData(2), "/")(0), 4))
                                        If InStr(aItemData(3), "^") > 0 Then
                                            strData = strData & "|" & Format(Val(Split(aItemData(3), "^")(1)), "0.00")
                                        Else
                                            strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                        End If
                                    Else
                                        If Mid(strBuffer, 1, 1) <> "C" Then
                                            blnReadItem = False
                                            Exit Do
                                        End If
                                    End If
                                Loop
                                '添加计算项 2008-07-11
                                Dim varCacle As Variant, str项目 As String
                                Dim dblGLO As Double '球蛋白=　总蛋白TP(678)-白蛋白ALB(413)
                                Dim dblTP As Double, dblALB As Double
                                
                                Dim dblIBIL As Double '间接胆红素
                                Dim dblTBIL As Double, dblDBIL As Double
                                
                                
                                Dim dblAST As Double, dblALT As Double
                                Dim dblAPOA As Double, dblAPOB As Double
                                Dim dblA1C3 As Double, dblHB As Double
                                
                                varCacle = Split(strData, "|")
                                If UBound(varCacle) > 5 Then
                                    For i = 5 To UBound(varCacle)
                                        If i Mod 2 = 1 Then
                                            If Trim(varCacle(i)) = "5" Then dblTP = Val(varCacle(i + 1))
                                            If Trim(varCacle(i)) = "6" Then dblALB = Val(varCacle(i + 1))
                                            
                                            If Trim(varCacle(i)) = "8" Then dblTBIL = Val(varCacle(i + 1))
                                            If Trim(varCacle(i)) = "9" Then dblDBIL = Val(varCacle(i + 1))
                                            
                                            If Trim(varCacle(i)) = "1" Then dblALT = Val(varCacle(i + 1))
                                            If Trim(varCacle(i)) = "18" Then dblAST = Val(varCacle(i + 1))
                                            
                                            If Trim(varCacle(i)) = "313" Then dblAPOA = Val(varCacle(i + 1))
                                            If Trim(varCacle(i)) = "314" Then dblAPOB = Val(varCacle(i + 1))
                                            
                                            If Trim(varCacle(i)) = "74" Then dblA1C3 = Val(varCacle(i + 1))
                                            If Trim(varCacle(i)) = "125" Then dblHB = Val(varCacle(i + 1))
                                        End If
                                    Next
                                    
                                    If dblTP <> 0 And dblALB <> 0 Then
                                        dblGLO = dblTP - dblALB
                                        strData = strData & "|GLO|" & Format(dblGLO, "0.000")
                                        strData = strData & "|A/G|" & Format(dblALB / dblGLO, "0.000")
                                    End If
                                    
                                    If dblTBIL <> 0 And dblDBIL <> 0 Then
                                        strData = strData & "|IBIL|" & Format(dblTBIL - dblDBIL, "0.000")
                                    End If
                                    
                                    If dblAST <> 0 And dblALT <> 0 Then
                                        strData = strData & "|X001|" & Format(dblAST / dblALT, "0.000")
                                    End If
                                    
                                    If dblAPOA <> 0 And dblAPOB <> 0 Then
                                        strData = strData & "|APOA/APOB|" & Format(dblAPOA / dblAPOB, "0.000")
                                    End If
                                    
                                    If dblA1C3 <> 0 And dblHB <> 0 Then
                                        strData = strData & "|HBA1C|" & Format(0.82 * ((dblA1C3 * 100) / dblHB) + 2.76, "0.000")
                                    End If
                                End If
                                
                                '一个标本结束
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                                
                                iStrPoint = iPos
                                If InStr("PO", Mid(strBuffer, 1, 1)) = 0 Then
                                '不是下一个标本
                                    Exit Do
                                End If
                            End If
                        Loop
                    Case Else
                        '信息完整，不再保留
                        iCurrGrpEntry = 0
                        iStrPoint = iPos
                End Select
            Else
                Exit Do
            End If
        Else
            iStrPoint = iPos
        End If
    Loop
            
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
        '向设备发送‘数据收到’命令
        Select Case iResultType
            Case 1 '返回结果
                strCmd = "0|" & strACK
            Case 2 '处理仪器请求
                strCmd = "1|" & strACK
            Case Else
                strCmd = strACK
        End Select
    Else
        '向设备发送‘数据收到’命令
        Select Case iResultType
            Case 1 '返回结果
                strCmd = "0|" & strACK
            Case 2 '处理仪器请求
                strCmd = "1|" & strACK
            Case Else
                strCmd = strACK
        End Select
    End If
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    Dim blnEnd As Boolean
    blnSuccess = True
    Select Case iSendStep
        Case 0 '开始发送
            SendSample = Chr(5)
            iSendStep = 1
        Case Else '发送数据
            If strResponse Like "*" & Chr(6) & "*" Then '允许发送
                If iSendStep = 1 Then '第一次发送
                    '生成设备串
                    strDeviceString = GetDeviceString(strSample, blnUndo, iType)
                End If
                strLastString = Mid(strDeviceString, (iSendStep - 1) * 240 + 1, 240)
                strLastString = iSendStep & strLastString
                If iSendStep * 240 >= Len(strDeviceString) Then '最后一段或只有一段
                    strLastString = strLastString & Chr(&H3)
                    blnEnd = True
                Else
                    strLastString = strLastString & Chr(&H17)
                    blnEnd = False
                End If
                strLastString = Chr(&H2) & strLastString & GetCheckCode(strLastString) & Chr(&HD) & Chr(&HA)
                
                SendSample = strLastString
                If blnEnd Then
                    SendSample = SendSample & Chr(&H4)  '发送结束
                    iSendStep = 0
                Else
                    iSendStep = iSendStep + 1
                End If
            Else
'                If iSendStep = 1 Then '在开始即不允许发送
                    SendSample = Chr(&H6) & Chr(&H4) '发送结束
                    iSendStep = 0
                    blnSuccess = False '发送失败
'                Else '重新发送
'                    SendSample = strLastString '重新发送上次内容
'                End If
            End If
    End Select
End Function

Private Function GetDeviceString(ByVal strSample As String, Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    Dim i As Long, aItem() As String
    Dim strItem As String, iSampleType As Integer
    Dim blnNewMode As Boolean
    
    blnNewMode = False
    
    If blnNewMode Then
        'New Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1|||||||||||||"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & Space(22) & "|" & _
            aItem(1) & "^^^^"
        Select Case True
            Case aItem(3) Like "*血*"
                GetDeviceString = GetDeviceString & "S1"
                iSampleType = 1
            Case aItem(3) Like "*尿*"
                GetDeviceString = GetDeviceString & "S2"
                iSampleType = 2
            Case aItem(3) Like "*脊髓*"
                GetDeviceString = GetDeviceString & "S3"
                iSampleType = 3
            Case Else
                GetDeviceString = GetDeviceString & "S5"
                iSampleType = 5
        End Select
        GetDeviceString = GetDeviceString & "^SC"
    
        GetDeviceString = GetDeviceString & "|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||" & iSampleType & "||||||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "C|1|I|^^^^|G" & vbCr & "L|1|N" & vbCr
    Else
        'Current Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & _
            aItem(1) & "^" & Space(22) & "^"
        GetDeviceString = GetDeviceString & "1" '标本都默认为血液，否则仪器不认？
        iSampleType = 1
'        Select Case True
'            Case aItem(3) Like "*血*"
'                GetDeviceString = GetDeviceString & "1"
'                iSampleType = 1
'            Case aItem(3) Like "*尿*"
'                GetDeviceString = GetDeviceString & "2"
'                iSampleType = 2
'            Case aItem(3) Like "*脊髓*"
'                GetDeviceString = GetDeviceString & "3"
'                iSampleType = 3
'            Case Else
'                GetDeviceString = GetDeviceString & "5"
'                iSampleType = 5
'        End Select
        GetDeviceString = GetDeviceString & "^^"
    
        GetDeviceString = GetDeviceString & "|R1|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||||||^^^^||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "L|1|N" & vbCr
    End If
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '从指定位置开始读取一行，同时指针定位到下一个换行之后
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

