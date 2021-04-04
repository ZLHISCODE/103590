VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_U168"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: U168
'厂商:
'作者: 曾超
'用户: 南京江北人民医院
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
    Dim strTmpReturn As String, blnEnd As Boolean
    Dim lngLength As Long, str日期 As String, str标本号 As String, str检验人标本质控 As String, strLine As String
    Dim strItem As String, i As Integer, blnLineEnd As Boolean
    Dim strTmp As String, intStart As Integer, intEnd As Integer
    Dim intPosition As Integer                  '当前点
    Dim intloop As Integer
    Dim astrLine() As String
    Dim strItemVal As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    '没有回车换行符时退出
    If InStr(strSource, vbCr) <= 0 Then strReserved = strSource: Exit Sub
    '查找开始位置
    astrLine = Split(strSource, vbCrLf)
    intStart = 0
    For intloop = 0 To UBound(astrLine)
        If IsNumeric(Trim(Mid(astrLine(intloop), 1, 2))) = True Then
            intStart = InStr(strSource, Mid(astrLine(intloop), 1, 4))
            Exit For
        End If
    Next
    '没有找到开始位置时退出
    If intStart = 0 Then strReserved = strSource: Exit Sub
    
    '查找结束位置
    intEnd = InStr(strSource, "ASC")
    If intEnd <= 0 Then strReserved = strSource: Exit Sub
    '未找到结束位置退出
    'If InStr(intEnd, strSource, "neg") <= 0 Then strReserved = strSource: Exit Sub
    
    strLine = Mid(strSource, intStart, intEnd + 24 - intStart)
    
    '只解析一个,其他返回下次解析
    strReserved = Replace(Mid(strSource, intEnd + 30), vbCrLf & vbCrLf, "")
    
    astrLine = Split(strLine, vbCrLf)
    
    strData = Format(Mid(astrLine(0), 4, 2) & "-" & TrunMonth(Mid(astrLine(0), 7, 3)) & "-" & Mid(astrLine(0), 11, 2), "yyyy-mm-dd") & " " & Trim(Mid(astrLine(0), 13))
    strData = strData & "|" & Val(Mid(astrLine(2), 10, 4)) & "| |尿|0"
    
    For intloop = 3 To UBound(astrLine)
        If intloop = 7 Or intloop = 9 Then
            strItemVal = Format(Val(Trim(Mid(astrLine(intloop), 6))), "0.00")
            strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItemVal
        Else
            strItem = Replace(Trim(Mid(astrLine(intloop), 6, 9)), "-", "－")
            strItemVal = Format(Val(Trim(Mid(astrLine(intloop), 15))), "0.00")
            If Val(strItemVal) = 0 Then
                strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItem & "  " & Trim(Mid(astrLine(intloop), 15))
            Else
                strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItem & "  " & strItemVal
            End If
        End If
    Next
    strResult = strData
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


Private Function TrunMonth(ByVal strMonth As String) As String
    '转换中文的月份为数字
    Select Case UCase(strMonth)
        Case "一月": TrunMonth = "01"
        Case "二月": TrunMonth = "02"
        Case "三月": TrunMonth = "03"
        Case "四月": TrunMonth = "04"
        Case "五月": TrunMonth = "05"
        Case "六月": TrunMonth = "06"
        Case "七月": TrunMonth = "07"
        Case "八月": TrunMonth = "08"
        Case "九月": TrunMonth = "09"
        Case "十月": TrunMonth = "10"
        Case "十一": TrunMonth = "11"
        Case "十二月": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function


