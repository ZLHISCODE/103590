VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ROCHE_UR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘厦
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
    Dim blnEndSample As Boolean '标本结束
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer, strFlag As String
    Dim lngNo As Long
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '查找一组数据的开始位置
            iStrPoint = InStr(iStrPoint, strSource, "")
            '判断一组结果是否开始
            If iStrPoint > 0 Then
                '记录一组结果的入口
                iStrPoint = iStrPoint + 1
                iCurrGrpEntry = iStrPoint - 1
                '数据是否完整
                iPos = InStr(iStrPoint, strSource, "")
                If iPos > 0 Then
                    '查找当前标本起始位置
                    Do While iStrPoint < iPos
                        strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                        If UCase(Mid(strBuffer, 1, 7)) = "SEQ.NO." Then Exit Do
                    Loop
                    If iStrPoint < iPos Then
                        '当前标本信息有效
                        lngNo = Val(Mid(strBuffer, 8))
                        '检验时间
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                        aDatas = Split(Split(strBuffer, " ")(0), ".")
                        strData = Format(aDatas(2) & "-" & aDatas(1) & "-" & aDatas(0) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        strData = strData & "|" & lngNo
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|尿液" '本仪器不支持
                        '质控品
                        strData = strData & "|0"
                        
                        '准备接收
                        blnReadItem = True
                    End If
                End If
            Else
                Exit Do
            End If
        Else
            iStrPointBK = iStrPoint
            strBuffer = ReadLine(strSource, iStrPointBK)
            
            '标本是否结束
            iPos = InStr(strBuffer, "")
            blnEndSample = True
            If iPos = 0 Then iPos = Len(strBuffer) + 1: blnEndSample = False
            
            strBuffer = Mid(strBuffer, 1, iPos - 1)
            If Len(strBuffer) > 0 Then '处理结果
                strData = strData & "|" & Trim(Mid(strBuffer, 2, 4))
                
                strItemResult = Trim(Mid(strBuffer, 7, Len(strBuffer) - 12))
                strFlag = Trim(Mid(strBuffer, 19))
                '取消计量单位
                For i = Len(strItemResult) To 1 Step -1
                    If Mid(strItemResult, i, 1) = " " Or IsNumeric(Mid(strItemResult, i, 1)) Then Exit For
                Next
                If i = 0 Then
                    strItemResult = IIf(Left(strBuffer, 1) = "*", IIf(strFlag = "", "+", strFlag), "-") '& strItemResult
                Else
                    strItemResult = Trim(Mid(strItemResult, 1, i)) & IIf(Left(strBuffer, 1) = "*", IIf(strFlag = "", "(+)", "(" & strFlag & ")"), "")
                End If
'                If Not IsNumeric(strItemResult) Then
'                    strItemResult = IIf(UCase(strItemResult) Like "NEG*" Or UCase(strItemResult) Like "NORM*", "-", strItemResult)
'                End If
                strData = strData & "|" & strItemResult
            End If
            
            If blnEndSample Then '标本已结束
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                iStrPoint = iStrPointBK
            End If
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

Public Function SendSample(ByVal strSample As String) As String
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

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos > 0 Then
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 2
    Else
        iPos = InStr(iPoint, strSource, Chr(13))
        If iPos = 0 Then iPos = InStr(iPoint, strSource, Chr(10))
        If iPos = 0 Then iPos = Len(strSource) + 1
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

