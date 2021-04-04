VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ATB"
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

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strBarcode As String                    '条码
    Dim strVerifyDate As String                 '检验时间
    Dim strSampleNumber As String               '样本序号
    Dim strVerifyName As String                 '检验人
    Dim intStart As Integer                     '记录临时分解开始位置
    Dim intEnd As Integer                       '记录临时分解结整位置
    Dim strSpecimen As String                   '送检标本
    Dim intLoop As Integer                      '临时循环变理
    
    Dim blnReadItem As Boolean                  '当前是否正在读取检验项目
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    intSpot = InStr(strSource, "")
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        
        If Not blnReadItem Then '查找一组数据的开始位置
        
            iStrPoint = InStr(iStrPoint, strSource, "NO:")
            '判断一组结果是否开始
            If iStrPoint > 0 Then
                '记录一组结果的入口
                iCurrGrpEntry = iStrPoint
                
                iPos = InStr(iStrPoint, strSource, "")        '不完整就退出
                If iPos = 0 Then Exit Do
                If InStr(iPos, strSource, "") > 0 Then
                
                    iStrPoint = InStr(iStrPoint, strSource, "标本编号:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '标本序号
                    strSampleNumber = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    iStrPoint = InStr(iStrPoint, strSource, "鉴定结果:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '送检标本（这里指细菌）
                    strSpecimen = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    iStrPoint = InStr(iStrPoint, strSource, "报告日期:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '检验时间
                    strVerifyDate = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    strData = strData & strVerifyDate & " " & Format(Time, "hh:mm:ss")
                    '时间格式必须为：yyyy-mm-dd hh:mm:ss
                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    
                    strData = strData & "|" & strSampleNumber & "^0"
                    
                    strData = strData & "| |" & strSpecimen & "|0"
                    
                    '找到项目开始的位置
                    
                    iStrPoint = InStr(iStrPoint, strSource, vbCrLf & vbCrLf)
                    
                    For intLoop = 1 To 6
                        strTmp = ReadSubString(strSource, iStrPoint, 1)
                        If Asc(strTmp) >= 65 And Asc(strTmp) <= 122 Then
                            Exit For
                        End If
                    Next
                    
                    iStrPoint = iStrPoint - 1
                    blnReadItem = True
                End If
            Else
                Exit Do
            End If
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '结果结束
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                strData = ""
                blnReadItem = False
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                                        
                '通道码
                iStrPoint = iStrPoint - 1
                intEnd = InStr(iStrPoint, strSource, " ")
                strData = strData & "|" & ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                
                
                '耐药
                iStrPoint = InStr(iStrPoint, strSource, "   ") + 3
                intEnd = InStr(iStrPoint, strSource, " ")
                strTmp = Trim(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                Select Case strTmp
                    Case "耐药"
                        strTmp = "R"
                    Case "中介"
                        strTmp = "I"
                    Case "敏感"
                        strTmp = "S"
                End Select
                strData = strData & "|" & strTmp
                
                '方法
                intEnd = InStr(iStrPoint, strSource, ":") + 1
                strTmp = Trim(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                Select Case strTmp
                    Case "MIC:"
                        strTmp = 1
                    Case "DISK:"
                        strTmp = 2
                    Case "K-B:"
                        strTmp = 3
                End Select
                strData = strData & "^" & strTmp
                
                '结果
                intEnd = InStr(iStrPoint, strSource, " ")
                
                strTmp = ReadSubString(strSource, iStrPoint, IIf(intEnd - iStrPoint = 0, 1, intEnd - iStrPoint))
                strData = strData & "^" & Trim(strTmp)
                iStrPoint = InStr(iStrPoint, strSource, vbCrLf) + 2
                
            End If
        End If
    Loop
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '向设备发送‘数据收到’命令
    strCmd = Chr(&H6&)
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


