VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_PCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘厦
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
    Dim aDatas() As Variant, strItemResult As String, i As Integer
    Dim lngNo As Long
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "Sample information:")
        '判断一组结果是否开始
        If iStrPoint > 0 Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint
            '数据是否完整
            iPos = InStr(iStrPoint, strSource, "Cycles" & Chr(vbKeyTab) & "F1")
            If iPos > 0 Then
                '查找当前标本起始位置
                Do While iStrPoint < iPos
                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                    If UCase(Mid(strBuffer, 1, 6)) = "NUMBER" Then Exit Do
                Loop
                If iStrPoint < iPos Then
                    Do While True
                        strBuffer = Trim(ReadLine(strSource, iStrPoint))
                        If Len(strBuffer) = 0 Then Exit Do
                        
                        aDatas = getArray(strBuffer)
                        If UBound(aDatas) >= 4 Then
                            If aDatas(1) = "Sample" Then
                                '检验时间
                                '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                '样本序号，必须为数字
                                Select Case Mid(aDatas(2), 1, 2)
                                    Case "BQ", "BS", "NQ", "NS", "TQ", "TS", "CT"
                                        strData = strData & "|" & Val(Mid(aDatas(2), 3))
                                    Case Else
                                        strData = strData & "|" & Val(Mid(aDatas(2), 2))
                                End Select
                                'strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", Val(Mid(aDatas(2), 3)), Val(Mid(aDatas(2), 2)))
                                '检验人
                                strData = strData & "| " '本仪器不支持
                                '标本
                                strData = strData & "|血清" '本仪器不支持
                                '质控品
                                strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", 1, 0)
                                
                                '通道码
                                Select Case Mid(aDatas(2), 1, 2)
                                    Case "BQ", "BS", "NQ", "NS", "TQ", "TS", "CT"
                                        strData = strData & "|" & Mid(aDatas(2), 1, 2)
                                    Case Else
                                        strData = strData & "|" & Mid(aDatas(2), 1, 1)
                                End Select
                            
'                                strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", Mid(aDatas(2), 1, 2), Mid(aDatas(2), 1, 1))
                                
                                '结果
                                strData = strData & "|" & aDatas(3)
                
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                            End If
                        End If
                    Loop
                End If
                
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
            
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '向设备发送‘数据收到’命令
    strCmd = Chr(&H6&)
End Sub

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject, objStream As TextStream
    Dim lngResultCount As Long
    Dim strSource As String
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iPos As Long
    Dim aDatas() As Variant
    Dim strIDCode As String  '标识码
    
    Set objStream = objFile.OpenTextFile(strFile)
    strSource = objStream.ReadAll
    objStream.Close
        
    '初始化返回值
    iStrPoint = 1
    lngResultCount = -1
        
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "Sample information:")
        '判断一组结果是否开始
        If iStrPoint > 0 Then
            '数据是否完整
            iPos = InStr(iStrPoint, strSource, "Standard Line Data")
            If iPos > 0 Then
                '查找当前标本起始位置
                Do While iStrPoint < iPos
                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                    If UCase(Mid(strBuffer, 1, 6)) = "NUMBER" Then Exit Do
                Loop
                If iStrPoint < iPos Then
                    Do While True
                        strBuffer = Trim(ReadLine(strSource, iStrPoint))
                        If Len(strBuffer) = 0 Then Exit Do
                        
                        aDatas = getArray(strBuffer)
                        If UBound(aDatas) >= 2 Then
                            If aDatas(1) = "Sample" Then
                                '检验时间
                                '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                '样本序号，必须为数字，质控从10000开始
                                Select Case Mid(aDatas(2), 1, 2)
                                    Case "BQ"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 100
                                    Case "BS"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 200
                                    Case "NQ"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 300
                                    Case "NS"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 400
                                    Case "TQ"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 500
                                    Case "TS"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 600
                                    Case "CT"
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)) + 700
                                    Case Else
                                        strData = strData & "|" & Val(Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1))
                                End Select
'                                strData = strData & "|" & Mid(aDatas(2), CheckNumber(CStr(aDatas(2))) + 1)
                                '检验人
                                strData = strData & "| " '本仪器不支持
                                '标本
                                strData = strData & "|血液" '本仪器不支持
                                '质控品
                                strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", 1, 0)
                                
                                '通道码
                                strData = strData & "|" & Mid(aDatas(2), 1, CheckNumber(CStr(aDatas(2))))
                                
                                '结果
                                If UBound(aDatas) > 2 Then
                                    If UBound(aDatas) >= 4 And Mid(aDatas(2), 1, 2) = "BS" Then
                                        strData = strData & "|" & aDatas(4)
                                    Else
                                        strData = strData & "|" & aDatas(3)
                                    End If
                                Else
                                    strIDCode = Mid(aDatas(2), 1, CheckNumber(CStr(aDatas(2))))
                                    
                                    Select Case strIDCode
                                        Case "MP", "HPV", "TB"
                                            strData = strData & "|阴性"
                                        Case Else
                                            strData = strData & "|0"
                                    End Select
                                End If
                                
                                lngResultCount = lngResultCount + 1
                                ReDim Preserve aResults(lngResultCount)
                                aResults(lngResultCount) = strData
                            End If
                        End If
                    Loop
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    ResultFromFile = aResults
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

Private Function getArray(ByVal strData As String) As Variant()
    Dim i As Long
    Dim aTmp() As String, aData() As Variant
    
    strData = Replace(strData, Chr(vbKeyTab), " ")
    aTmp = Split(strData, " "): aData = Array()
    
    For i = 0 To UBound(aTmp)
        If Len(Trim(aTmp(i))) > 0 Then
            ReDim Preserve aData(UBound(aData) + 1)
            aData(UBound(aData)) = aTmp(i)
        End If
    Next
    getArray = aData
End Function
Private Function CheckNumber(strNumber As String) As Integer
    '功能:          取第一个为数字的字符位数
    '参数:          取字符的字串
    
    Dim intLoop As Integer
    
    For intLoop = 1 To Len(strNumber)
        If IsNumeric(Mid(strNumber, intLoop, 1)) = True Then
            If intLoop > 1 Then
                CheckNumber = intLoop - 1
            Else
                CheckNumber = intLoop
            End If
            Exit Function
        End If
    Next
    CheckNumber = Len(strNumber)
End Function
