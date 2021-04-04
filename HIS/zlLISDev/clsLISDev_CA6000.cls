VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA6000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘厦 双向
Private mstrSendString As String

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
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim strSampleType As String '样本类型
    Dim strItemValue As String
    Dim iPos  As Long, intDataType As Integer '数据类型 0＝结果、1＝请求
    Dim strBarcode As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '查找一组数据的开始位置
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
                    If InStr(iStrPoint, strSource, "") > 0 Then
                        '检验时间
                        iStrPoint = iStrPoint + 8
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '样本类型
                        
                        strData = ReadSubString(strSource, iStrPoint, 2)   '年
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '月
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) '小时
                        strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '分秒
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 6
                        strBarcode = Trim(Replace(UCase(ReadSubString(strSource, iStrPoint, 13)), "QC", ""))
                        strData = strData & "|" & Val(Right(strBarcode, 9))
                        If Len(strBarcode) > 10 Then
                            strData = strData & "^^" & strBarcode
                        End If
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血液" '本仪器不支持
                        '质控品
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:常规、E:STAT分析、S:标准曲线、C:质控
                        
                        '准备接收
                        iStrPoint = iStrPoint + 12
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
                        '检验时间
                        iStrPoint = iStrPoint + 8
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '样本类型
                        
                        strData = ReadSubString(strSource, iStrPoint, 2)   '年
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '月
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) '小时
                        strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '分秒
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 6
                        strBarcode = Trim(ReadSubString(strSource, iStrPoint, 13))
                        strData = strData & "|" & Val(Right(strBarcode, 9))
                        If Len(strBarcode) > 10 Then
                            strData = strData & "^^" & strBarcode
                        End If
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血液" '本仪器不支持
                        '质控品
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:常规、E:STAT分析、S:标准曲线、C:质控
                        
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
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '结果结束
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                '检验项目
                iStrPoint = iStrPoint - 1
                strBuffer = ReadSubString(strSource, iStrPoint, 3)
                strData = strData & "|" & strBuffer
                '确定小数位置
                Select Case Mid(strBuffer, 3, 1)
                    Case "3", "4", "5" '比率、标准比
                        strItemValue = ReadSubString(strSource, iStrPoint, 6)
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & Val(strItemValue) / 100
                        Else
                            strData = strData & "|*"
                        End If
                    Case Else
                        Select Case strBuffer
                            Case "062" 'Fpg
                                strItemValue = ReadSubString(strSource, iStrPoint, 6)
                                If IsNumeric(strItemValue) Then
                                    strData = strData & "|" & Val(strItemValue) / 100
                                Else
                                    strData = strData & "|*"
                                End If
                            Case Else
                                strItemValue = ReadSubString(strSource, iStrPoint, 6)
                                If IsNumeric(strItemValue) Then
                                    strData = strData & "|" & Val(strItemValue) / 10
                                Else
                                    strData = strData & "|*"
                                End If
                        End Select
                End Select
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
                strCmd = "1|" & Chr(&H6&)
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
    Dim iPos As Long, strItem As String, strSampleNO As String
    
    blnSuccess = True
    Select Case iSendStep
        Case 0 '开始发送
            If strResponse Like "*R*" Or Len(Trim(strResponse)) = 0 Then '允许发送
                If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                    If Len(Trim(strSample)) > 0 Then
                        '请求结束
                        '生成设备串
                        aItem = Split(strSample, "|")
                        '标本或SampleID
                        strBuffer = aItem(1)
                        strSampleID = aItem(5)
                        '读取反馈的架号和杯号
                        strBuffer = "210101"
                        If Len(Trim(strResponse)) > 0 Then
                            iPos = InStr(strResponse, "R")
                            strBuffer = Mid(strResponse, iPos + 3, 6)
                        End If
                        SendSample = "S2" & Right(Space(6) & strBuffer, 6) & "U"
                        If Len(Trim(strResponse)) > 0 Then
                            iPos = InStr(strResponse, "R")
                            strBuffer = Mid(strResponse, iPos + 10, 29)
                            SendSample = SendSample & strBuffer & "B" & Space(11)
                        Else
                            SendSample = SendSample & Format(aItem(0), "yymmdd") & Format(Now, "hhmm") & "000101" & _
                                Right(Space(13) & strSampleID, 13) & "B" & Space(11)
                        End If
                        
                        For i = 10 To UBound(aItem)
                            strItem = Right(String(2, "0") & Left(aItem(i), 2), 2) & "0" & Space(6)
                            If InStr(SendSample, strItem) = 0 Then
                                SendSample = SendSample & strItem
                            End If
                        Next
                        SendSample = SendSample & ""
                    
                        iSendStep = 0
                        mstrSendString = SendSample
                    ElseIf Len(Trim(strResponse)) > 0 Then
                        '请求结束
                        '生成设备串
                        '读取反馈的架号和杯号
                        iPos = InStr(strResponse, "R")
                        strBuffer = Mid(strResponse, iPos + 3, 6)
                        
                        SendSample = "S2" & Right(Space(6) & strBuffer, 6) & "U"
                            
                        iPos = InStr(strResponse, "R")
                        strBuffer = Mid(strResponse, iPos + 10, 16)
                        strSampleNO = Mid(strResponse, iPos + 29, 10)
                        If Mid(strResponse, iPos + 27, 2) <> "88" Then
                            '非质控标本,标本号加10000
                            strSampleNO = Right(Space(10) & (Val(strSampleNO) + 10000), 10)
                        Else
                            strSampleNO = Right(Space(10) & "QC" & Val(strSampleNO), 10)
                        End If
                        SendSample = SendSample & strBuffer & "   " & strSampleNO & "C" & Space(11) & "040      050      060      510      "
                        
                        SendSample = SendSample & ""
                    
                        iSendStep = 0
                        mstrSendString = SendSample
                    End If
                Else
                    SendSample = ""
                    iSendStep = 0
                    blnSuccess = False '发送失败
                End If
            Else
                SendSample = ""
                iSendStep = 0
            End If
        Case Else
            If InStr(strResponse, Chr(&H6)) = 0 Then
                '未成功，重发
                SendSample = Chr(&H15)
                iSendStep = 0
            Else
                '发送成功
                SendSample = ""
                iSendStep = 0
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
