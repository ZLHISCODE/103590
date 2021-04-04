VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_array"
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
    
    Dim iStrPoint As Integer, iPos As Integer '字符指针
    Dim strBuffer As String, strData As String, strChannel As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    Dim str检验时间 As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1
    If InStr(1, strSource, "[88]") > 0 Then
        Do While iStrPoint <= Len(strSource)
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            '判断一组结果是否开始
            If strBuffer = "[" Then
                '记录一组结果的入口
                iCurrGrpEntry = iStrPoint - 1
                
                'iStrPoint = iStrPoint + 7
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "20" Then
                    If Len(strSource) - iStrPoint >= 80 Then
                        '检验时间
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '日
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '月
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '年
                        iStrPoint = iStrPoint + 1
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) & ":" & ReadSubString(strSource, iStrPoint, 2)
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        str检验时间 = strData
                        iStrPoint = iStrPoint + 3

                        '样本序号，必须为数字
                         Do While iStrPoint <= Len(strSource)
                             strData = ""
                             Do While iStrPoint <= Len(strSource)
                                iPos = iStrPoint '保留当前位置
                                strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                If Mid(strBuffer, 3, 2) = "21" Then
                                   iStrPoint = iPos + 8
                                   strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                                   strData = str检验时间 & "|" & strSampleID
                                    '检验人
                                   strData = strData & "| " '本仪器不支持
                                    '标本
                                   strData = strData & "|血液" '本仪器不支持
                                    '质控品
                                   strData = strData & "|0" '本仪器不支持，设为非质控品
                                   Exit Do
                                End If
                             Loop
                            
                            '检验项目
                            Do While iStrPoint <= Len(strSource)
                                iPos = iStrPoint '保留当前位置
                                strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                If Mid(strBuffer, 3, 2) = "22" Then
                                   
                                   strChannel = Trim(ReadSubString(strBuffer, 10, 9))
                                   strData = strData & "|" & strChannel
                                   strData = strData & "|" & Trim(ReadSubString(strBuffer, 22, 6))
                                ElseIf Mid(strBuffer, 3, 2) = "99" Then
                                   Exit Do
                                End If
                             Loop
        
                            '该组结果已解析
                            iCurrGrpEntry = 0
                            
                            strResult = IIf(Len(strResult) > 0, IIf(Len(strData) > 0, strResult & "||", strResult), "") & strData
                        Loop
                    End If
                Else
                    '该组结果无效
                    If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                End If
                '指针定位到下一个换行之后
                MoveToNextLine strSource, iStrPoint
            End If
        Loop
        '返回未解析的原始串
        If iCurrGrpEntry > 0 Then
            strReserved = Mid(strSource, iCurrGrpEntry)
        End If
        '向设备发送‘数据收到’命令
        strCmd = Chr(&H6&)
    Else
        strReserved = strSource
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
Private Function ReadLine(ByVal strSource, iPoint As Integer) As String
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


