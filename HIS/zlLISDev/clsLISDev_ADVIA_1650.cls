VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_1650"
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
    '1R 010102420070803N11            01-01                                  M   20070803 1.013  1M    21.2     4M     4.5     7M   43.20    10M    23.2    13M    37.7 *  16M    60.5 *R 19M     174    22M   76.42    28M    4.17    31M    7.47    34M     300    37M   221.4    40M      37    43M   87.40    49M    1.62    52M    3.72    67M    1.67    70M    1.00    73M    0.97    76M  131.39    85M   74.94 D 324M     121l  325M     3.0l  326M      90l   
    Dim iStrPoint As Integer '字符指针
    Dim iStrEndLeng As Integer '字串长度
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim strSampleType As String '样本类型
    Dim varTmp() As String      '标本集
    Dim intloop As Integer      '循环数
    Dim blnVar As Boolean       '是否是值
    Dim intHightLow As Integer  '=1高 =0低
    Dim blnQc  As Boolean       '是否质控
    Dim intStartPos As Integer
    Dim strHead  As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    '回应请求
    If InStr(strSource, Chr(&H5&)) > 0 Then
        strCmd = Chr(&H6&)
        strReserved = ""
        Exit Sub
    End If
    
    '没有请求时退出（可能有两种情况 &H17 和 &H3)
    If InStr(strSource, Chr(&H17&)) = 0 And InStr(strSource, Chr(&H3&)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strHead = Mid(strSource, 1, 30)
    
    iStrPoint = 1
    
    iStrPoint = IIf(InStr(strHead, "N") = 0, InStr(strHead, "I"), InStr(strHead, "N"))
    If iStrPoint = 0 Then
        iStrPoint = InStr(strHead, "C")
    End If
    iStrPoint = iStrPoint - 8
    
'    iStrPoint = InStr(strSource, Chr(&H20&)) + 2
'    iStrPoint = iStrPoint + 6
    strBuffer = ReadSubString(strSource, iStrPoint, 8)
    strBuffer = Mid(strBuffer, 1, 4) & "-" & Mid(strBuffer, 5, 2) & "-" & Mid(strBuffer, 7)
    '时间
    strData = Format(strBuffer, "YYYY-MM-DD") & " " & Format(Now, "HH:MM:SS")
    
    iStrPoint = InStr(strHead, "N") + 1
    If iStrPoint = 1 Then iStrPoint = InStr(strHead, "I") + 1
    '处理质控
    If iStrPoint = 1 Then
        If InStr(strSource, "PA") > 0 Then iStrPoint = InStr(strSource, "PA") + 2: intHightLow = 0
        If InStr(strSource, "PB") > 0 Then iStrPoint = InStr(strSource, "PB") + 2: intHightLow = 1
        If InStr(strSource, "PC") > 0 Then iStrPoint = InStr(strSource, "PC") + 2: intHightLow = 2
        If InStr(strSource, "PD") > 0 Then iStrPoint = InStr(strSource, "PD") + 2: intHightLow = 3
        If InStr(strSource, "PE") > 0 Then iStrPoint = InStr(strSource, "PE") + 2: intHightLow = 4
        If InStr(strSource, "PF") > 0 Then iStrPoint = InStr(strSource, "PF") + 2: intHightLow = 5
        blnQc = True
    End If
    
    iStrEndLeng = InStr(iStrPoint, strSource, Chr(&H20&))
    iStrEndLeng = iStrEndLeng - iStrPoint
    strBuffer = ReadSubString(strSource, iStrPoint, iStrEndLeng)
    '标本号
    If blnQc = False Then
        strData = strData & "|" & Val(Mid(strBuffer, 2))
    Else
        If intHightLow = 0 Then
            strData = strData & "|" & 9100 + Val(strBuffer)
        ElseIf intHightLow = 1 Then
            strData = strData & "|" & 9200 + Val(strBuffer)
        ElseIf intHightLow = 2 Then
            strData = strData & "|" & 9300 + Val(strBuffer)
        ElseIf intHightLow = 3 Then
            strData = strData & "|" & 9400 + Val(strBuffer)
        ElseIf intHightLow = 4 Then
            strData = strData & "|" & 9500 + Val(strBuffer)
        Else
            strData = strData & "|" & 9600 + Val(strBuffer)
        End If
    End If
    '检验人
    strData = strData & "| "
    '血清
    strData = strData & "|血清"
    '质控品
    strData = strData & IIf(blnQc, "|1", "|0")
    
    varTmp = Split(strSource, "M")
    
    If strSource Like "*1R*" Then intStartPos = 1
    If strSource Like "*2R*" Then intStartPos = 0
    
    For intloop = intStartPos To UBound(varTmp)
        If intloop = intStartPos Then
            '开始单独处理
            strData = strData & "|" & Trim(Right(varTmp(intloop), 3)) & "M"
        ElseIf intloop = UBound(varTmp) Then
            '结束单独处理
            strData = strData & "|" & Val(Trim(Mid(varTmp(intloop), 1, Len(varTmp(intloop)) - 4)))
        Else
            '中间部份处理
            strData = strData & "|" & Val(Trim(Mid(varTmp(intloop), 1, Len(varTmp(intloop)) - 4))) & "|" & Trim(Right(varTmp(intloop), 3)) & "M"
            
        End If
        
    Next
    Debug.Print strData
'    varTmp = Split(Mid(strSource, iStrPoint))
'    For intloop = 0 To UBound(varTmp)
'        '碰到结束符时退出
'        If InStr(varTmp(intloop), Chr(&H3&)) > 0 Or InStr(varTmp(intloop), Chr(&H17&)) > 0 Then
'            Exit For
'        End If
'
'        If blnVar = True And varTmp(intloop) <> "" And InStr(varTmp(intloop), "M") = 0 Then
'            strData = strData & "|" & Val(varTmp(intloop))
'            blnVar = False
'        End If
'
'        If blnVar = False Then
'            If InStr(varTmp(intloop), "M") Then
'                If varTmp(intloop) <> "M" Then
'                    strData = strData & "|" & varTmp(intloop)
'                    blnVar = True
'                Else
'                    blnVar = False
'                End If
'            Else
'                blnVar = False
'            End If
'        Else
'            If InStr(varTmp(intloop), "M") Then
'                If varTmp(intloop) <> "M" Then
'                    strData = strData & "| "
'                    blnVar = False
'                End If
'            End If
'        End If
'    Next
'    If Mid(strData, Len(strData)) = "M" Then
'        strData = strData & "| "
'    End If
    strResult = strData
    strReserved = ""
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
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




