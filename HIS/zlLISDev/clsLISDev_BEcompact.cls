VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BEcompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'重庆市二院

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
    Const strSplit = "|"
    Dim i As Integer, iPos As Integer
    Dim aItemData() As String
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String, blnEnd As Boolean
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
        '----- 开始解析结果
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的消息
            strLine = GetStr_Section(strReserved, Chr(&H2), Chr(&H3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)

    
            iStrPoint = 1
            Do While iStrPoint <= Len(strLine)
                strBuffer = Trim(ReadLine(strLine, iStrPoint))
                Do While Len(strBuffer) > 10
                    aItemData = Split(strBuffer, strSplit)
                    If UBound(aItemData) < 7 Then '无效数据
                       iCurrGrpEntry = iPos
                       Exit Do
                    End If
                    i = Trim(aItemData(2)) '每一行都取标本信息，i作为计算通道的基数
                      '检验时间
                       '时间格式必须为：yyyy-mm-dd hh:mm:ss
                    
                       strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                       iStrPoint = iStrPoint + 1
                       '样本序号，必须为数字
                     
                       strSampleID = Val(Trim(aItemData(1)))
                       strData = strData & "|" & strSampleID
                       '检验人
                       strData = strData & "| " '本仪器不支持
                       '标本
                       strData = strData & "|血液" '本仪器不支持
                       '质控品
                       strData = strData & "|0" '本仪器不支持，设为非质控品
                          
                        strData = strData & "|" & 3 * i + 1
                        strData = strData & "|" & Val(Trim(aItemData(5)))
                        strData = strData & "|" & 3 * i + 2
                        strData = strData & "|" & Val(Trim(aItemData(6)))
                        strData = strData & "|" & 3 * i + 3
                        strData = strData & "|" & Val(Trim(aItemData(7)))
                        '该组结果已解析
                        iCurrGrpEntry = 0
                
                    If Len(strData) > 0 Then strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                     strData = ""
                     iPos = iStrPoint
                     strBuffer = Trim(ReadLine(strLine, iStrPoint))
                 Loop
            Loop
            '返回未解析的原始串
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '向设备发送‘数据收到’命令
'                strCmd = Chr(&H6&)
'            End If
            
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
        strCmd = Chr(&H6&)
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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




