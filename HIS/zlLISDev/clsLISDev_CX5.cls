VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CX5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'双向
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
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    Dim strTP As Long, strALP As Long, strTB As Long, strDB As Long, strIB As Long, strAG As Long, strGLO As Long
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strChennel As String, lngResult As Double
    Dim intDataType As Integer '数据类型 0＝结果、1＝请求
    Dim strbarcode As String
    Dim intRequest As Integer   '0请求 1结果
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1
    strTB = 0
    strDB = 0
    strTP = 0
    strALP = 0
    Do While iStrPoint <= Len(strSource)
        strBuffer = ReadSubString(strSource, iStrPoint, 1)
        '判断一组结果是否开始
        If strBuffer = "[" Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint - 1
            
            iStrPoint = iStrPoint + 3
            strBuffer = ReadSubString(strSource, iStrPoint, 3)
            Select Case strBuffer
            
            Case "702"   '701，请求，702，结果
                intRequest = 1
                intDataType = 0
                iStrPoint = iStrPoint + 1
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "03" Then
                    If Len(strSource) - iStrPoint >= 80 Then
                        '检验时间
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '日
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '月
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '年
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 30
                        
                        strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                        strData = strData & "|" & IIf(Len(strSampleID) = 0, strTmpID, strSampleID)
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血液" '本仪器不支持
                        '质控品
                        strData = strData & "|0" '本仪器不支持，设为非质控品
                        
                        '检验项目
                        iStrPoint = iStrPoint + 1
                        strChennel = ReadSubString(strSource, iStrPoint, 3)
                        strData = strData & "|" & strChennel
                        iStrPoint = iStrPoint + 19
                        lngResult = Val(ReadSubString(strSource, iStrPoint, 9))
                        strData = strData & "|" & lngResult
                        Select Case strChennel
                            Case "11A"
                                strTB = lngResult
                            Case "12A"
                                strDB = lngResult
                            Case "07A"
                                strTP = lngResult
                            Case "08A"
                                strALP = lngResult
                        End Select
                        If strTB <> 0 And strDB <> 0 Then
                            strIB = strTB - strDB
                            strTB = 0
                            strDB = 0
                            strData = strData & "|IBIL|" & strIB
                        End If
                        If strTP <> 0 And strALP <> 0 Then
                            strGLO = strTP - strALP
                            strAG = strALP / strGLO
                            strTP = 0
                            strALP = 0
                            strData = strData & "|A/G|" & strAG & "|GLO|" & strGLO
                        End If
                        '该组结果已解析
                        iCurrGrpEntry = 0
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                    End If
                Else
                    '该组结果无效
                    If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                End If
            Case "701" '非结果
                intRequest = 0
                intDataType = 1
                iStrPoint = iStrPoint + 1
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "06" Then '请求反馈的条码，一组最多七个 [-0,701,06,samp1------,samp2------,samp3------,samp4------,samp5------,samp6------,samp7------]
                    Do While iStrPoint < Len(strSource)
                        
                        iStrPoint = iStrPoint + 1
                        strbarcode = Trim(ReadSubString(strSource, iStrPoint, 11)) '条码
                        strSampleID = Val(Mid(strbarcode, Len(strbarcode) - 1))
                        If Val(strbarcode) <> 0 Then
                            strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            strData = strData & "|" & strSampleID & "^0^" & strbarcode
                                    
                            '检验人
                            strData = strData & "| " '本仪器不支持
                            '标本
                            strData = strData & "|血液" '本仪器不支持
                            '质控品
                            strData = strData & "|0"
                            strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        '该组结果已解析
                        End If
                        iCurrGrpEntry = 0
                    Loop
                End If
            End Select
            '指针定位到下一个换行之后
            MoveToNextLine strSource, iStrPoint
        End If
    Loop
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    If intRequest = 0 Then
        '向设备发送‘数据收到’命令
        strCmd = "1|" & Chr(&H6) & Chr(&H3)
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
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim blnEnd As Boolean
    blnSuccess = True
    Select Case iSendStep
        Case 0 '开始发送
            SendSample = Chr(4) & Chr(1)
            iSendStep = 1
        Case Else '发送数据
            If strResponse Like "*" & Chr(6) & "*" Then '允许发送
                If iSendStep = 1 Then '第一次发送
                    '生成设备串
                    strDeviceString = GetDeviceString(strSample, blnUndo, iType)
                End If
'                strLastString = Mid(strDeviceString, (iSendStep - 1) * 240 + 1, 240)
'                strLastString = iSendStep & strLastString
'                If iSendStep * 240 >= Len(strDeviceString) Then '最后一段或只有一段
'                    strLastString = strLastString & Chr(&H3)
'                    blnEnd = True
'                Else
'                    strLastString = strLastString & Chr(&H17)
'                    blnEnd = False
'                End If
                strLastString = strDeviceString & GetCheckCode(strDeviceString) & Chr(&HD) & Chr(&HA)
                
                SendSample = strLastString
'                If blnEnd Then
'                    SendSample = SendSample & Chr(&H4)  '发送结束
'                    iSendStep = 0
'                Else
'                    iSendStep = iSendStep + 1
'                End If
            Else
'                If iSendStep = 1 Then '在开始即不允许发送
                    SendSample = Chr(&H4) '发送结束
                    iSendStep = 0
                    blnSuccess = False '发送失败
'                Else '重新发送
'                    SendSample = strLastString '重新发送上次内容
'                End If
            End If
    End Select
End Function
Private Function GetDeviceString(ByVal strSample As String, Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    Dim I As Long, aItem() As String
    Dim strItem As String, iSampleType As Integer
    Dim blnNewMode As Boolean
    
    '    [00,701,01,01,03,1,RO,SE,samp3------,CONTROL_NAME________,
'-------------------------,SAMPLE_COMMENT_TWO_______,
'------------------,PATIENT_F_NAME_,M,------------,
'------------------,------,----,--------------------,
'000,4,------,M,-------------------------,-------,
'----,----,------,001,01B-,0]2F<CR><LF>

    blnNewMode = False
    
    If blnNewMode Then
        'New Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "[00,701,01,00,00,1,RO,SE," & aItem(1) & Space(11 - Len(aItem(1))) & "," & _
        aItem(1) & Space(20 - Len(aItem(1))) & "," & Space(25) & "," & Space(25) & "," & Space(18) & "," & Space(15) & "," & Space(1) & "," & Space(12) & _
         "," & Space(18) & "," & Format(aItem(0), "ddmmyy") & "," & Format(aItem(0), "hhmm") & "," & Space(20) & "," & Space(3) & "," & Space(1) & "," & Space(6) & "," & Space(1) & "," & Space(25) & _
         "," & Space(7) & "," & Space(4) & "," & Space(4) & "," & Space(6) & "," & Space(3)
            
        '取通道数据
        strItem = ""
        For I = 10 To UBound(aItem)
            strItem = strItem & "," & aItem(I) & ",0"
        Next
        GetDeviceString = GetDeviceString & strItem
        GetDeviceString = GetDeviceString & "]"
    End If
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = "]" Then
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

Private Function GetCheckCode(ByVal strSource As String) As String
'获取校验码
    Dim lngAsc As Long, I As Integer
    
    For I = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, I, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

