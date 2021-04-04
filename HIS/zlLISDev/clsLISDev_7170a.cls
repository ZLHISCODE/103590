VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_7170a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    'AnalysestrSource = ":s 1010 1                3 060506103610 6 315.9  8  69.8  9  39.9 10   2.8 11     4 12    94 14   1.2 15   4.0 16   7.0 24    14 "
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    Dim strTP As Double, stralb As Double  '保留计算项目的原始值 25-26
    Dim str高密 As Double, str低密 As Double  '34/39
    Dim str通道, str结果 As String
    Dim str球蛋白, str血球比, str高低密 As Double, str常规 As String
    Dim bln常规 As Boolean, blnQc As Boolean
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    Dim int序号 As Integer
    
    blnQc = False
    iStrPoint = 1
    Do While iStrPoint <= Len(strSource)
        strBuffer = ReadSubString(strSource, iStrPoint, 1)
'        If InStr(strSource, "") <= 0 Then
'            strReserved = strSource
'            Exit Do
'        End If
        '判断一组结果是否开始
        If strBuffer = "" Then
            '记录一组结果的入口
            
            strTP = 0: stralb = 0: str高密 = 0: str低密 = 0: bln常规 = False: str常规 = ""
            iCurrGrpEntry = iStrPoint - 1 '63530918
            '判断常规与急诊
            iStrPoint = iStrPoint + 1
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            Select Case UCase(strBuffer)
                Case "N"
                    str常规 = ""
                Case "Q"
                    str常规 = "^1"
                Case "F"
                    str常规 = ""
                    blnQc = True
            End Select
            iStrPoint = iStrPoint + 26
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
                If Len(strSource) - iStrPoint >= 20 Then
                    If Not blnQc Then
                    '检验时间
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '月
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '年
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) & ":" & ReadSubString(strSource, iStrPoint, 2)
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint - 36
                    Else
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        iStrPoint = iStrPoint - 23
                    End If
                    '样本号
'                    iStrPoint = iStrPoint + 1
                    strSampleID = Trim(ReadSubString(strSource, iStrPoint, 4)) & str常规
                    'strData = strData & "|" & IIf(Len(strSampleID) = 0, strTmpID, strSampleID)
                    strData = strData & "|" & strSampleID
                    '检验人
                    strData = strData & "| " '本仪器不支持
                    '标本
                    strData = strData & "|血液" '本仪器不支持
                    '质控品
                        If blnQc Then
                           strData = strData & "|1"
                           iStrPoint = iStrPoint + 33
                        Else
                           strData = strData & "|0"
                           iStrPoint = iStrPoint + 35
                        End If
                    
                    '检验项目
                    int序号 = 0
                    Do While ReadSubString(strSource, iStrPoint, 1) <> ""
                        If int序号 > 50 Then Exit Do
                        iStrPoint = iStrPoint - 1
                        str通道 = Trim(ReadSubString(strSource, iStrPoint, 3))
                        strData = strData & "|" & str通道
                        str结果 = Val(Trim(ReadSubString(strSource, iStrPoint, 7)))
                        strData = strData & "|" & str结果
                        Select Case str通道
                               Case 25
                                   strTP = str结果
                               Case 26
                                   stralb = str结果
                               Case 34
                                   str高密 = str结果
                               Case 39
                                   str低密 = str结果
                        End Select
                    int序号 = int序号 + 1
                    Loop
                        '该组结果已解析
                    If Not blnQc Then
                           str球蛋白 = strTP - stralb
                        If strTP <> 0 And stralb <> 0 Then
                           str血球比 = Format(stralb / str球蛋白, "0.00")
                           strData = strData & "|GLO|" & str球蛋白 & "|A/G|" & str血球比
                        End If
                        If str低密 <> 0 And str高密 <> 0 Then
                           str高低密 = Format(str高密 / str低密, "0.00")
                           strData = strData & "|H/L|" & str高低密
                        End If
                    End If
                    iCurrGrpEntry = 0
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                    
                End If
            '指针定位到下一个换行之后
         '   MoveToNextLine strSource, iStrPoint
        End If
    Loop
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '向设备发送‘数据收到’命令
         strCmd = Chr(&H6&)
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
    SendSample = ""
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








