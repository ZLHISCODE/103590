VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU640"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘厦 双向
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim intQC As Integer       '质控位数
    Dim intDataType As Integer '数据类型 0＝结果、1＝请求
    Dim intEnd As Integer      '最后位置
    Dim strResultTmp As String    '临时保存结果
    Dim intRulueLenth As Integer    '结果长度
    
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    Dim intE As Integer, intP As Integer, intQ As Integer
    Dim intStart As Integer         '检验项目通道码的起始位置(以 E 所在位置为标准)
    Dim str质控编码规则 As String
    
    intE = Val(ReadIni("AU640", "E", App.Path & "\ccccccc"))
    If intE = 0 Then intE = 8000:   Call WriteIni("AU640", "E", 8000, App.Path & "\AU640.ini")
    
    intQ = Val(ReadIni("AU640", "Q", App.Path & "\AU640.ini"))
    If intQ = 0 Then intE = 6000:   Call WriteIni("AU640", "Q", 6000, App.Path & "\AU640.ini")
    
    intP = Val(ReadIni("AU640", "P", App.Path & "\AU640.ini"))
    If intP = 0 Then intE = 3000:   Call WriteIni("AU640", "P", 3000, App.Path & "\AU640.ini")
    
    intRulueLenth = Val(ReadIni("AU640", "结果长度", App.Path & "\AU640.ini"))
    If intRulueLenth = 0 Then intRulueLenth = 8:   Call WriteIni("AU640", "结果长度", 8, App.Path & "\AU640.ini")
    
    intStart = Val(ReadIni("AU640", "项目起始位置", App.Path & "\AU640.ini"))
    If intStart = 0 Then intStart = 0: Call WriteIni("AU640", "项目起始位置", 0, App.Path & "\AU640.ini")
    
    str质控编码规则 = ReadIni("AU640", "质控编码规则", App.Path & "\AU640.ini")
    If str质控编码规则 = "" Then str质控编码规则 = "默认": Call WriteIni("AU640", "质控编码规则", str质控编码规则, App.Path & "\AU640.ini")
    
    
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '查找一组数据的开始位置
            strSource = Replace(strSource, "DB", "")
            strSource = Replace(strSource, "DE", "")
            strSource = Replace(strSource, "D01B", "")
            strSource = Replace(strSource, "D01E", "")
            strSource = Replace(strSource, "RB", "")
            strSource = Replace(strSource, "RE", "")
            
            intDataType = -1
            iPos = InStr(iStrPoint, strSource, "R")
            If iPos > 0 Then
                '仪器请求
                intDataType = 1
            Else
                '仪器结果
                iPos = InStr(iStrPoint, strSource, "D")
                If iPos > 0 Then intDataType = 0
            End If
            Select Case intDataType
                Case 0 '结果
                    '记录一组结果的入口
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '数据是否完整
                    iPos = InStr(iStrPoint, strSource, "")
                    If iPos > 0 Then
                        '去掉两次传送中连接信息
                        Do While True
                            iPos = InStr(strSource, "")
                            If iPos = 0 Then Exit Do
                            
                            strSource = Mid(strSource, 1, iPos - 1) & Mid(strSource, iPos + 2 + 41)
                        Loop
                
                        '检验时间
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 9
                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
                        If Not IsNumeric(Mid(strBuffer, 1, 1)) Then
                            Select Case Mid(strBuffer, 1, 1)
                                Case "E"
                                    strData = strData & "|" & intE + Val(Mid(strBuffer, 2)) & "^1"
                                Case "P"
                                    strData = strData & "|" & intP + Val(Mid(strBuffer, 2)) & "^1"
                                Case "Q"
                                    blnQc = True
                                    Select Case str质控编码规则
                                        Case "默认", ""
                                            strData = strData & "|" & intQ + Val(Mid(strBuffer, 2)) + _
                                                  Val(Mid(strSource, iStrPoint, 24)) & "^1"
                                        Case "塘厦"
                                            strData = strData & "|" & intQ + Val(Mid(strSource, iStrPoint, 24)) * 10 + _
                                                    Val(Mid(strBuffer, 2)) & "^1"
                                    End Select
                            End Select
                        Else
                            strData = strData & "|" & Val(strBuffer) & "^0"
                        End If
                        'SampleID或条码,20位
                        strData = strData & "^" & ReadSubString(strSource, iStrPoint, 20)
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血清" '本仪器不支持
                        '质控品
                        If blnQc = True Then
                            strData = strData & IIf(blnQc, "|1", "|0")
                            intEnd = InStr(iStrPoint, strSource, "E")
                            Select Case Val(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                                Case 1
                                    strData = strData & "^" & 402
                                Case 2
                                    strData = strData & "^" & 403
                                Case 3
                                    strData = strData & "^" & 404
                                Case 4
                                    strData = strData & "^" & 405
                                Case 5
                                    strData = strData & "^" & 406
                                Case 6
                                    strData = strData & "^" & 407
                                Case Else
                                    strData = strData & "^" & Val(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                            End Select
                            
                        Else
                            intEnd = InStr(iStrPoint, strSource, "E")
                            strData = strData & "|0"
                            iStrPoint = iStrPoint + (intEnd - iStrPoint)
                        End If
                        'strData = strData & "|0" 'IIf(Mid(strSource, iStrPoint, 1) = "E", "|1", "|0")
                        
                        '准备接收
                        iStrPoint = iStrPoint + 1
                        blnReadItem = True
                    Else
                        Exit Do
                    End If
                Case 1 '请求
                    '记录一组结果的入口
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '数据是否完整
                    iPos = InStr(iStrPoint, strSource, "")
                    If iPos > 0 Then
                        '去掉两次传送中连接信息
                        Do While True
                            iPos = InStr(strSource, "")
                            If iPos = 0 Then Exit Do
                            
                            strSource = Mid(strSource, 1, iPos - 1) & Mid(strSource, iPos + 2 + 41)
                        Loop
                
                        '检验时间
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 9
                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
                        If Not IsNumeric(Mid(strBuffer, 1, 1)) Then
                            strData = strData & "|" & Val(Mid(strBuffer, 2)) & "^1"
                        Else
                            strData = strData & "|" & Val(strBuffer) & "^0"
                        End If
                        'SampleID或条码,20位
                        strData = strData & "^" & ReadSubString(strSource, iStrPoint, 20)
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血清" '本仪器不支持
                        '质控品
                        strData = strData & "|0" 'IIf(Mid(strSource, iStrPoint, 1) = "E", "|1", "|0")
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        '该组结果已解析
                        iCurrGrpEntry = 0
                        If InStr(iStrPoint, strSource, "R") = 0 Then Exit Do
                    Else
                        Exit Do
                    End If
                Case Else
                    Exit Do
            End Select
        Else
            iStrPoint = iStrPoint + intStart
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            
            If strBuffer = "" Then '结果结束
'                '计算白球蛋白(3-4)，通道码：50
'                iPos = InStr(28, strData, "|03|")
'                sng1 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
'                iPos = InStr(28, strData, "|04|")
'                sng2 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
'                strData = strData & "|50|" & Mid(CStr(sng1 - sng2), 1, 8)
'                '糖化血红蛋白()
                
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                '检验项目
                iStrPoint = iStrPoint - 1
                strData = strData & "|" & ReadSubString(strSource, iStrPoint, 2)
                strResultTmp = ReadSubString(strSource, iStrPoint, intRulueLenth)
                
                iStrPoint = iStrPoint - intStart
                If InStr(strResultTmp, "!") > 0 Then
                    'strData = strData & "|" & Val(strResultTmp) & "^!^^"
                    strData = strData & "|" & "????"
                Else
                    strData = strData & "|" & Val(strResultTmp)
                End If
                
            End If
        End If
    Loop
            
    If blnReadItem Then
        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
        '该组结果已解析
        iCurrGrpEntry = 0
    End If
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '向设备发送‘数据收到’命令
        Select Case intDataType
            Case 0 '返回结果
                strCmd = "0|" & Chr(&H6&)
            Case 1 '处理仪器请求
                strCmd = "1|"
            Case Else
                strCmd = ""
        End Select
    End If

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
    
    blnSuccess = True
    Select Case iSendStep
'        Case 0 '开始发送
'            SendSample = Chr(6)
'            iSendStep = 1
        Case Else '发送数据
            If strResponse Like "*R*" Or Len(Trim(strResponse)) = 0 Then '允许发送
                If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                    '请求结束
                    '生成设备串
                    aItem = Split(strSample, "|")
                    '标本或SampleID
                    strBuffer = aItem(1)
                    strSampleID = aItem(5)
                    '读取反馈的架号和杯号
                    strRackNO = ""
                    If Len(Trim(strResponse)) > 0 Then
                        strResponse = Replace(strResponse, "RB", "")
                        iPos = InStr(strResponse, "R")
                        strRackNO = Mid(strResponse, iPos + 3, 11)
                    End If
                    If Len(Trim(strSampleID)) > 0 Then
                        'SampleID
                        SendSample = "S " & Right(Space(11) & strRackNO, 11) & Right(Space(20) & strSampleID, 20) & Space(4) & "E"
                    Else
                        SendSample = "S        " & Right(String(4, "0") & strBuffer, 4) & Space(20) & Space(4) & "E"
                    End If
                    For i = 10 To UBound(aItem)
                        SendSample = SendSample & aItem(i)
                    Next
                    SendSample = SendSample & "SE"
                
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

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
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


