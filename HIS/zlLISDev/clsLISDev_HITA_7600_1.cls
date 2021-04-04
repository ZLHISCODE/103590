VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7600_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串
Private mstrData As String

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
    Dim strACK As String
    Dim aLineData() As String
    Dim strTmp As String
    
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 0
    
    
    If InStr(Right(strSource, 10), Chr(&H5)) > 0 Or InStr(Right(strSource, 10), Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    If InStr(strSource, "1H|") = 0 Or InStr(strSource, Chr(&H4)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(23))
        If i = 0 Then Exit Do
        If i + 3 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, InStr(i, strSource, Chr(&H2)) + 2)
        Else
            Exit Do
        End If
    Loop
    
    '得到时间和标本号
    aLineData = Split(strSource, "O|")
    strTmp = Replace$(Replace(Replace(aLineData(1), vbCr, ""), vbLf, ""), Chr(&H2), "")
    aItemData = Split(strTmp, "|")
    strData = Mid(aItemData(21), 1, 4) & "-" & Mid(aItemData(21), 5, 2) & "-" & Mid(aItemData(21), 7, 2) & " " & _
                          Mid(aItemData(21), 9, 2) & ":" & Mid(aItemData(21), 11, 2) & ":" & Mid(aItemData(21), 13, 2)
    If aItemData(2) = "Q" Then
        If InStr(aLineData(1), "P2") > 0 Then
            strData = strData & "|" & Split(aItemData(1), "^")(0) & "1| |血液|0"
        Else
            strData = strData & "|" & Split(aItemData(1), "^")(0) & "| |血液|0"
        End If
    Else
        strData = strData & "|" & Split(aItemData(1), "^")(0) & IIf(aItemData(4) = "S", "^1", "") & "| |血液|0"
    End If
    
    '得到结果
    aLineData = Split(Replace(strSource, "|R|", ""), "R|")
    For i = 1 To UBound(aLineData)
        strTmp = Replace$(Replace(Replace(aLineData(i), vbCr, ""), vbLf, ""), Chr(&H2), "")
        aItemData = Split(strTmp, "|")
        strData = strData & "|" & Replace(Split(aItemData(1), "^")(3), "/", "") & "|" & aItemData(2)
    Next
    
'    aLineData = Split(strSource, Chr(&HD))
'    For i = 0 To UBound(aLineData)
'
'        strTmp = Replace(Replace(aLineData(i), vbCr, ""), vbLf, "")
'        Debug.Print strTmp
'        aItemData = Split(strTmp, "|")
'        If UBound(aItemData) > 5 Then
'            '标本号和时间
'            If aItemData(0) = "O" Then
'
'                strTmp = Replace$(Replace(Replace(aLineData(i), vbCr, ""), vbLf, ""), Chr(&H2), "")
'
'                aItemData = Split(strTmp, "|")
'
'                strData = Mid(aItemData(22), 1, 4) & "-" & Mid(aItemData(22), 5, 2) & "-" & Mid(aItemData(22), 7, 2) & " " & _
'                          Mid(aItemData(22), 9, 2) & ":" & Mid(aItemData(22), 11, 2) & ":" & Mid(aItemData(22), 13, 2)
'
'                If aItemData(3) = "Q" Then
'                    If InStr(aLineData(i + 1), "P2") > 0 Then
'                        strData = strData & "|" & Split(aItemData(2), "^")(0) & "1| |血液|0"
'                    Else
'                        strData = strData & "|" & Split(aItemData(2), "^")(0) & "| |血液|0"
'                    End If
'                Else
'                    strData = strData & "|" & Split(aItemData(2), "^")(0) & IIf(aItemData(5) = "S", "^1", "") & "| |血液|0"
'                End If
'            End If
'            '结果
'            If aItemData(0) = "R" Then
'                strTmp = Replace$(Replace(Replace(aLineData(i), vbCr, ""), vbLf, ""), Chr(&H2), "")
'                If InStr(strTmp, Chr(&H17)) > 0 Then
'                    strTmp = Mid$(strTmp, 1, InStr(strTmp, Chr(&H17)) - 1) & Mid$(strTmp, InStr(strTmp, Chr(&H17)) + 4)
'                End If
'                aItemData = Split(strTmp, "|")
'                strData = strData & "|" & Replace(Split(aItemData(2), "^")(3), "/", "") & "|" & aItemData(3)
'            End If
'        End If
'
'    Next
    If strData <> "" Then
        strResult = strData
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
    Dim i As Long, aItem() As String
    Dim strItem As String, iSampleType As Integer
    Dim blnNewMode As Boolean
    
    blnNewMode = False
    
    If blnNewMode Then
        'New Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1|||||||||||||"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & Space(22) & "|" & _
            aItem(1) & "^^^^"
        Select Case True
            Case aItem(3) Like "*血*"
                GetDeviceString = GetDeviceString & "S1"
                iSampleType = 1
            Case aItem(3) Like "*尿*"
                GetDeviceString = GetDeviceString & "S2"
                iSampleType = 2
            Case aItem(3) Like "*脊髓*"
                GetDeviceString = GetDeviceString & "S3"
                iSampleType = 3
            Case Else
                GetDeviceString = GetDeviceString & "S5"
                iSampleType = 5
        End Select
        GetDeviceString = GetDeviceString & "^SC"
    
        GetDeviceString = GetDeviceString & "|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||" & iSampleType & "||||||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "C|1|I|^^^^|G" & vbCr & "L|1|N" & vbCr
    Else
        'Current Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & _
            aItem(1) & "^" & Space(22) & "^"
        GetDeviceString = GetDeviceString & "1" '标本都默认为血液，否则仪器不认？
        iSampleType = 1
'        Select Case True
'            Case aItem(3) Like "*血*"
'                GetDeviceString = GetDeviceString & "1"
'                iSampleType = 1
'            Case aItem(3) Like "*尿*"
'                GetDeviceString = GetDeviceString & "2"
'                iSampleType = 2
'            Case aItem(3) Like "*脊髓*"
'                GetDeviceString = GetDeviceString & "3"
'                iSampleType = 3
'            Case Else
'                GetDeviceString = GetDeviceString & "5"
'                iSampleType = 5
'        End Select
        GetDeviceString = GetDeviceString & "^^"
    
        GetDeviceString = GetDeviceString & "|R1|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||||||^^^^||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "L|1|N" & vbCr
    End If
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
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function



