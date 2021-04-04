VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA1500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'大医二院
'2007-07-09 陈东 修改
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
    Dim strBuffer As String, strData As String, strTmp As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim strSampleType As String '样本类型
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*D**" Then
        iStrPoint = 1: blnReadItem = False
        Do While iStrPoint <= Len(strSource)
            If Not blnReadItem Then '查找一组数据的开始位置
                strBuffer = ReadSubString(strSource, iStrPoint, 1)
                '判断一组结果是否开始
                If strBuffer = "D" Then
                    '记录一组结果的入口
                    iCurrGrpEntry = iStrPoint - 1
                    '数据是否完整
                    If Len(strSource) - iStrPoint >= 38 Then
                        '检验时间
                        iStrPoint = iStrPoint + 7
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '样本类型
                        
    '                    strData = ReadSubString(strSource, iStrPoint, 2)   '年
    '                    strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '月
    '                    strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                        strData = ReadSubString(strSource, iStrPoint, 2)   '月
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '日
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '年
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) '小时
                        strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '分秒
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        iStrPoint = iStrPoint + 6
                        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 13))
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血液" '本仪器不支持
                        '质控品
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:常规、E:STAT分析、S:标准曲线、C:质控
                        
                        '准备接收
                        iStrPoint = iStrPoint + 12
                        blnReadItem = True
                    End If
                End If
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
                            strTmp = ReadSubString(strSource, iStrPoint, 6)
                            If Val(strTmp) = 0 Then
                                strData = strData & "|" & Trim(strTmp)
                            Else
                                strData = strData & "|" & Val(strTmp) / 100
                            End If
                        Case Else
                            strTmp = ReadSubString(strSource, iStrPoint, 6)
                            If Val(strTmp) = 0 Then
                                strData = strData & "|" & Trim(strTmp)
                            Else
                                Select Case strBuffer
                                     Case "062" 'Fpg
                                        strData = strData & "|" & Val(strTmp) / 1000
                                    Case Else
                                        strData = strData & "|" & Val(strTmp) / 10
                                End Select
                            End If
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
            strCmd = Chr(&H6&)
        End If
    Else
        '数据不全,直接反回
        strReserved = strSource
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
