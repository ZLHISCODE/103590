VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UF100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: UF-100
'厂商: sysmex
'作者: (未知)
'用户: 医大二院
'修改: 2007-08-16 陈东
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
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "DS*" & Chr(3) & Chr(2) & "DP*" & Chr(3) & Chr(2) & "DC*" & Chr(3) & "*" Then
        strTmpReturn = strSource
        
        Do While Not blnEnd
            
            lngLength = InStr(strTmpReturn, Chr(3)) - InStr(strTmpReturn, Chr(2)) + 1
            strLine = Mid(strTmpReturn, InStr(strTmpReturn, Chr(2)), lngLength)
            strTmpReturn = Mid(strTmpReturn, InStr(strTmpReturn, Chr(3)) + 1)
            
            If strLine Like "*DS*" Then
                strItem = "": str日期 = "": str标本号 = "": str检验人标本质控 = "": blnLineEnd = False
                str日期 = Format(CDate(Mid(strLine, 11, 2) & "-" & Mid(strLine, 13, 2) & "-" & Mid(strLine, 15, 2)), "yyyy-MM-dd")
                If Mid(strLine, 40, 1) = " " Then
                    '大医二院格式
                    str标本号 = Val(Mid(strLine, 27, 13))
                ElseIf Mid(strLine, 44, 1) = " " Then
                    '2008-3-11 增加成都德阳格式
                    str标本号 = Val(Mid(strLine, 31, 13))
                End If
                str检验人标本质控 = " |尿液|" & IIf(Mid(strLine, 10, 1) = "C", 1, 0)
                
                Select Case Mid(strLine, 47, 1)
                    Case "0"
                        strItem = "|RBCInfo|RBC Negative"
                    Case "1"
                        strItem = "|RBCInfo|Normocytic"
                    Case "2"
                        strItem = "|RBCInfo|Microcytic"
                    Case "3"
                        strItem = "|RBCInfo|Non-classified"
                End Select
            ElseIf strLine Like "*DP*" Then
                strLine = Replace(Mid(strLine, 12), Chr(3), "")
                For i = 0 To 4
                    strItem = strItem & "|" & Mid(strLine, 1 + 12 * i, 4) & "|" & Val(Mid(strLine, 5 + 12 * i, 8))
                Next
            ElseIf strLine Like "*DC*" Then
                strLine = Replace(Mid(strLine, 12), Chr(3), "")
                If strLine <> "" Then
                    For i = 0 To Len(strLine) / 4
                        If Mid(strLine, 1 + 4 * i, 4) <> "" Then
                            strItem = strItem & "|" & Mid(strLine, 1 + 4 * i, 4) & "|+"
                        End If
                    Next
                End If
                blnLineEnd = True
            End If
            
            If Val(str标本号) > 0 And strItem <> "" And blnLineEnd Then
                strResult = strResult & "||" & str日期 & "|" & Val(str标本号) & "|" & str检验人标本质控 & strItem
            End If
            
            If Not strTmpReturn Like "*" & Chr(2) & "*" & Chr(3) & "*" Or blnLineEnd Then
                blnEnd = True
            End If
        Loop
'        strTmpReturn = Mid(strSource, InStr(strSource, Chr(2)) + 192)
'        strSource = Mid(strSource, InStr(strSource, Chr(2)), 192)
        
            'Beging 原来的程序
'            iStrPoint = 1: blnReadItem = False
'            Do While iStrPoint <= Len(strSource)
'                If Not blnReadItem Then '查找一组数据的开始位置
'                    strBuffer = ReadSubString(strSource, iStrPoint, 2)
'                    iStrPoint = iStrPoint - 1
'                    '判断一组结果是否开始
'                    If strBuffer = "DS" Then
'                        '记录一组结果的入口
'                        iCurrGrpEntry = iStrPoint - 1
'                        '数据是否完整
'                        If Len(strSource) - iStrPoint >= 117 Then
'                            '检验时间
'                            iStrPoint = iStrPoint + 7
'                            strSampleType = ReadSubString(strSource, iStrPoint, 1)   '样本类型
'
'                            strData = ReadSubString(strSource, iStrPoint, 2)   '年
'                            strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '月
'                            strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
'                            strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) '小时
'                            strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '分秒
'                            '时间格式必须为：yyyy-mm-dd hh:mm:ss
'                            strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
'
'                            '样本序号，必须为数字
'                            iStrPoint = iStrPoint + 6
'                            strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 13))
'
'                            '检验人
'                            strData = strData & "| " '本仪器不支持
'                            '标本
'                            strData = strData & "| " '本仪器不支持
'                            '质控品
'                            strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:常规、E:STAT分析、S:标准曲线、C:质控
'
'                            '准备接收
'                            iStrPoint = iStrPoint + 80
'                            blnReadItem = True
'                        End If
'                    End If
'                Else
'                    strBuffer = ReadSubString(strSource, iStrPoint, 1)
'                    If strBuffer = "" Then '结果结束
'                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'
'                        blnReadItem = False
'                        '该组结果已解析
'                        iCurrGrpEntry = 0
'                    Else
'                        '检验项目:0201红细胞、0202白细胞
'                        iStrPoint = iStrPoint - 1
'                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
'                        strData = strData & "|" & strBuffer
'                        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 8))
'                    End If
'                End If
'            Loop
'
'            If blnReadItem Then
'                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'                '该组结果已解析
'                iCurrGrpEntry = 0
'            End If
'            '返回未解析的原始串
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '向设备发送‘数据收到’命令
'                strCmd = Chr(&H6&)
'            End If
            'End 原来的程序
            '
            If strTmpReturn <> "" Then strReserved = strTmpReturn
    Else
        strReserved = strSource
        strCmd = Chr(&H6&)
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
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
