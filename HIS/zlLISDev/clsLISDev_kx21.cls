VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_kx21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘夏
'---仪器传输格式： K1000 (不带图形)
'D1U0703235000000000024000000C011900403001360040000993003370034000157000660005400880000080000600105001330019210117003900
'2007-03-23 00:00:00|24| |血液|0|1|11.9|2|4.03|3|136|4|0.4|5|99.3|6|33.7|7|34|8|157|9|0.066|10|0.054|11|0.88|12|0.8|13|0.6|14|10.5|15|0.133|16|19.21|17|11.7|18|0.39

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
    Dim i As Integer
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
    '06/ 8/1117:50           97ZRSZ   Q  8.3 3.66  1240.358 97.8 33.9  346   890.2090.0550.736  1.7  0.5  6.10.154 55.6---.----.--.--- 51.2161.7               +                                 MPMPMP
    '[kx21]D1U0610125000000000099000000S010400522001240038700741002380032000298003690005400577000380000600060004150010900089001700

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long
    Dim strItemValue As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    intSpot = InStr(strSource, "")
    
    iStrPoint = 1
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "D")
        '判断一组结果是否开始
        If iStrPoint > 0 Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint
            
            If InStr(iStrPoint, strSource, "") > 0 Then
                '检验时间
                iStrPoint = iStrPoint + 4
                strData = ReadSubString(strSource, iStrPoint, 2)   '年
                strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '月
                strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                
                strData = strData & " " & Format(Time, "hh:mm:00") '时分
                '时间格式必须为：yyyy-mm-dd hh:mm:ss
                strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                
                '样本序号，必须为数字 8位
                iStrPoint = iStrPoint + 5 '定位到标本位
                strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 8))
                '检验人
                strData = strData & "| " '本仪器不支持
                '标本
                strData = strData & "|血液" '本仪器不支持
                '质控品
                strData = strData & "|0" '本仪器不支持，设为非质控品
                
                '检验项目
                iStrPoint = iStrPoint + 7
                For i = 1 To 21
                    strData = strData & "|" & i
                    Select Case i
                        Case 3, 7, 8
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & Val(strItemValue) / 10
                            Else
                                strData = strData & "|*"
                            End If
                        Case 1, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & Val(strItemValue) / 100
                            Else
                                strData = strData & "|*"
                            End If
                        Case 2
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & Val(strItemValue) / 1000
                            Else
                                strData = strData & "|*"
                            End If
                        Case 4, 18
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & Val(strItemValue) / 10000
                            Else
                                strData = strData & "|*"
                            End If
                        Case Else
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & Val(strItemValue) / 100
                            Else
                                strData = strData & "|*"
                            End If
                    End Select
                Next
                '该组结果已解析
                iCurrGrpEntry = 0
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
            Else
                Exit Do
            End If
          
            '指针定位到下一个换行之后
        Else
            Exit Do
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


