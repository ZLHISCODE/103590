VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Coba121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'塘厦 双向
Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串

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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnMod2 As Boolean '模式2在标本号前加01
    Dim blnQc As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode 、2:New Mode
    Dim iResultType As Integer '结果类型：1－检验结果、2－仪器请求
    Dim i As Long
    Dim str39 As String, str32 As String, strItemCode As String
    
    
    Dim objFile As New Scripting.FileSystemObject, objStream As TextStream, strFile As String
    
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 0
    
    iStrPoint = 1: blnReadItem = False
    '取消请求符
    strSource = Replace(strSource, Chr(5), "")
    '取消结束符
    strSource = Replace(strSource, Chr(4), "")
    '取消开始符
    strSource = Replace(strSource, Chr(2), "")
    '取消回车换行符
    strSource = Replace(strSource, vbCrLf, "")
    '取消换行符
    strSource = Replace(strSource, Chr(10), "")
    '取消制表符
    strSource = Replace(strSource, Chr(9), "")
    '取消通信连接符
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(23))
        If i = 0 Then Exit Do
        If i + 3 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 4)
        Else
            Exit Do
        End If
    Loop
    
    '记录分隔符
    strFldDelim = "|"
    strRepeatDelim = "\" '重复内容分隔符
    strCompDelim = "^" '组间内容分隔符
    Do While iStrPoint <= Len(strSource)
        '查找信息的开始位置：以H开始
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Right(Split(strBuffer, strFldDelim)(0), 2) Like "*H" Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint
            
            '数据是否完整
            If InStr(iStrPoint, strSource, "L" & strFldDelim & "1" & strFldDelim & "N") > 0 Then  ' "L|1|N"是每个标本数据的结尾
            
                                       
                'aItemData = Split(strBuffer, strFldDelim)
                '信息完整，不再保留
                iCurrGrpEntry = 0  '当前这组结果的入口字符位置
                iResultType = 1    '结果类型：1－检验结果、2－仪器请求
                
                '查找标本信息
                iStrPoint = iPos: blnReadItem = False
                Do While iStrPoint <= Len(strSource)
                    If Not blnReadItem Then '当前是否正在读取检验项目
                    

                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Right(Split(strBuffer, strFldDelim)(0), 1) = "P" Then Exit Do
                        Loop
                        If iStrPoint > Len(strSource) Then
                            '当前信息无效，不再处理。
                            Exit Do
                        End If
                        
                        aItemData = Split(strBuffer, strFldDelim)
                        
                        '检验时间
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")

                        '样本序号，必须为数字
                        If aItemData(3) <> "" Then  '如果是读取条码病人
                            strData = strData & "|" & Val(Right(aItemData(3), 9))
                            If Len(aItemData(3)) > 10 Then
                                strData = strData & "^^" & aItemData(3)
                            End If
                            
                        Else
                         
                            Do While iStrPoint <= Len(strSource)
                                strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                If Right(Split(strBuffer, strFldDelim)(0), 1) = "O" Then Exit Do
                            Loop
                            aItemData = Split(strBuffer, strFldDelim)
                            aItemData = Split(aItemData(3), strCompDelim)
                            
                            strData = strData & "|" & Val(aItemData(1))
                        End If


                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血清"
                        '质控品
                        strData = strData & "|0"     '暂不处理质控
                        
                        '准备接收
                        
                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Right(Split(strBuffer, strFldDelim)(0), 1) = "C" Then Exit Do
                        Loop
                        
                        If iStrPoint > Len(strSource) Then
                            '当前信息无效，不再处理。
                            Exit Do
                        End If
                        
                        blnReadItem = True
                        
                        str39 = "": str32 = ""
                      
                       
                    Else

                        Do While iStrPoint <= Len(strSource)
                            
                            iPos = iStrPoint '保留当前位置
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            
                            aItemData = Split(strBuffer, strFldDelim)
                            'If Right(Split(strBuffer, strFldDelim)(0), 1) = "R" Then
                            If Right(aItemData(0), 1) = "R" Then
                                '修改内容：判断项目是否有结果值，无结果值项目不解析
                                If aItemData(3) <> "" Then
                                    '项目
                                    strItemCode = Split(aItemData(2), strCompDelim)(3) '第四个元素
                                    strData = strData & "|" & strItemCode
                                    '结果
                                     strData = strData & "|" & aItemData(3)
                                End If
                            Else
                                If Right(Split(strBuffer, strFldDelim)(0), 1) <> "C" Then
                                    blnReadItem = False
                                    Exit Do
                                End If
                            End If
                        Loop
                        
                        
                        '一个标本结束
                        '计算项目
                        If Trim(str39) <> "" And Trim(str32) <> "" And Val(str32) <> 0 Then
                            strData = strData & "|990|" & (Val(str39) / Val(str32))
                        End If
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        
                        iStrPoint = iPos
                        If InStr("PO", Right(Split(strBuffer, strFldDelim)(0), 1)) = 0 Then
                        '不是下一个标本
                            Exit Do
                        End If
                    End If
                Loop
            Else
                Exit Do
            End If
        Else
            iStrPoint = iPos
        End If
    Loop
            
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
        '向设备发送‘数据收到’命令
        Select Case iResultType
            Case 1 '返回结果
                strCmd = "0|" & Chr(&H6&)
            Case 2 '处理仪器请求
                strCmd = "1|" & Chr(&H6&)
            Case Else
                strCmd = Chr(&H6&)
        End Select
    Else
        '向设备发送‘数据收到’命令
        Select Case iResultType
            Case 1 '返回结果
                strCmd = "0|" & Chr(&H6&)
            Case 2 '处理仪器请求
                strCmd = "1|" & Chr(&H6&)
            Case Else
                strCmd = Chr(&H6&)
        End Select
    End If
                    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim blnEnd As Boolean
    blnSuccess = True
    Select Case iSendStep
        Case 0 '开始发送
            SendSample = Chr(5)
            iSendStep = 1
        Case Else '发送数据
            If strResponse Like "*" & Chr(6) & "*" Then '允许发送
                If iSendStep = 1 Then '第一次发送
                    '生成设备串
                    strDeviceString = GetDeviceString(strSample, blnUndo, iType)
                End If
                strLastString = Mid(strDeviceString, (iSendStep - 1) * 240 + 1, 240)
                strLastString = iSendStep & strLastString
                If iSendStep * 240 >= Len(strDeviceString) Then '最后一段或只有一段
                    strLastString = strLastString & Chr(&H3)
                    blnEnd = True
                Else
                    strLastString = strLastString & Chr(&H17)
                    blnEnd = False
                End If
                strLastString = Chr(&H2) & strLastString & GetCheckCode(strLastString) & Chr(&HD) & Chr(&HA)
                
                SendSample = strLastString
                If blnEnd Then
                    SendSample = SendSample & Chr(&H4)  '发送结束
                    iSendStep = 0
                Else
                    iSendStep = iSendStep + 1
                End If
            Else
'                If iSendStep = 1 Then '在开始即不允许发送
                    SendSample = Chr(&H6) & Chr(&H4) '发送结束
                    iSendStep = 0
                    blnSuccess = False '发送失败
'                Else '重新发送
'                    SendSample = strLastString '重新发送上次内容
'                End If
            End If
    End Select
End Function

Private Function GetDeviceString(ByVal strSample As String, Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
'    Dim i As Long, aItem() As String
'    Dim strItem As String, iSampleType As Integer
'    Dim blnNewMode As Boolean
'
'    blnNewMode = False
'
'    If blnNewMode Then
'        'New Mode
'        aItem = Split(strSample, "|")
'
'        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
'        GetDeviceString = GetDeviceString & vbCr & "P|1|||||||||||||"
'        GetDeviceString = GetDeviceString & vbCr & "O|1|" & Space(22) & "|" & _
'            aItem(1) & "^^^^"
'        Select Case True
'            Case aItem(3) Like "*血*"
'                GetDeviceString = GetDeviceString & "S1"
'                iSampleType = 1
'            Case aItem(3) Like "*尿*"
'                GetDeviceString = GetDeviceString & "S2"
'                iSampleType = 2
'            Case aItem(3) Like "*脊髓*"
'                GetDeviceString = GetDeviceString & "S3"
'                iSampleType = 3
'            Case Else
'                GetDeviceString = GetDeviceString & "S5"
'                iSampleType = 5
'        End Select
'        GetDeviceString = GetDeviceString & "^SC"
'
'        GetDeviceString = GetDeviceString & "|"
'        strItem = ""
'        For i = 10 To UBound(aItem)
'            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
'        Next
'        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
'        GetDeviceString = GetDeviceString & strItem
'
'        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
'        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
'        GetDeviceString = GetDeviceString & "||||N" & "||||" & iSampleType & "||||||||||O" & vbCr
'
'        GetDeviceString = GetDeviceString & "C|1|I|^^^^|G" & vbCr & "L|1|N" & vbCr
'    Else
'        'Current Mode
'        aItem = Split(strSample, "|")
'
'        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
'        GetDeviceString = GetDeviceString & vbCr & "P|1"
'        GetDeviceString = GetDeviceString & vbCr & "O|1|" & _
'            aItem(1) & "^" & Space(22) & "^"
'        GetDeviceString = GetDeviceString & "1" '标本都默认为血液，否则仪器不认？
'        iSampleType = 1
''        Select Case True
''            Case aItem(3) Like "*血*"
''                GetDeviceString = GetDeviceString & "1"
''                iSampleType = 1
''            Case aItem(3) Like "*尿*"
''                GetDeviceString = GetDeviceString & "2"
''                iSampleType = 2
''            Case aItem(3) Like "*脊髓*"
''                GetDeviceString = GetDeviceString & "3"
''                iSampleType = 3
''            Case Else
''                GetDeviceString = GetDeviceString & "5"
''                iSampleType = 5
''        End Select
'        GetDeviceString = GetDeviceString & "^^"
'
'        GetDeviceString = GetDeviceString & "|R1|"
'        strItem = ""
'        For i = 10 To UBound(aItem)
'            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
'        Next
'        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
'        GetDeviceString = GetDeviceString & strItem
'
'        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
'        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
'        GetDeviceString = GetDeviceString & "||||N" & "||||||||^^^^||||||O" & vbCr
'
'        GetDeviceString = GetDeviceString & "L|1|N" & vbCr
'    End If
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
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

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
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

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '从指定位置开始读取一行，同时指针定位到下一个换行之后
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos >= Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function



