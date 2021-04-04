VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ASTM_Diff2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    Dim aItemCan() As String, str标本信息 As String
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean, blnFirst As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode 、2:New Mode
    Dim iResultType As Integer '结果类型：1－检验结果、2－仪器请求
    Dim i As Long, strChannel As String, strSample As String, StrUnit As String '结果是否数字
    Dim str_WBC As String, str_RBC As String, str_PLT_R As String, str_PLT_F As String, Str_Img As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 1 '初始是返回结果
    
    iStrPoint = 1: blnReadItem = False
    '取消请求符
    strSource = Replace(strSource, Chr(5), "")
    '取消结束符
   ' strSource = Replace(strSource, Chr(4), "")
    '取消开始符
    strSource = Replace(strSource, Chr(2), "")
    '取消回车换行符
    'strSource = Replace(strSource, vbCrLf, "")
    '取消换行符
    strSource = Replace(strSource, Chr(10), "")
    '取消制表符
    strSource = Replace(strSource, Chr(9), "")
    '取消通信连接符
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(3))
        If i = 0 Then Exit Do
        If i + 2 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 3)
        Else
            Exit Do
        End If
    Loop
    
    Do While iStrPoint <= Len(strSource)
        '查找信息的开始位置：以H开始
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Mid(strBuffer, 1, 2) Like "*H*" Then
            '记录一组结果的入口
            iCurrGrpEntry = iStrPoint
            
            '记录分隔符
            If Mid(strBuffer, 1, 2) Like "*H" Then
                strFldDelim = Mid(strBuffer, 3, 1) '字段分隔符
                strRepeatDelim = Mid(strBuffer, 4, 1) '重复内容分隔符
                strCompDelim = Mid(strBuffer, 5, 1) '组间内容分隔符
            Else
                strFldDelim = Mid(strBuffer, 2, 1) '字段分隔符
                strRepeatDelim = Mid(strBuffer, 3, 1) '重复内容分隔符
                strCompDelim = Mid(strBuffer, 4, 1) '组间内容分隔符
            End If
            
            '数据是否完整
            If InStr(iStrPoint, strSource, "") > 0 Then
                '信息完整，不再保留
                iCurrGrpEntry = 0
                
                '查找标本信息
                iStrPoint = iPos: blnReadItem = False: blnFirst = True
                
                Do While iStrPoint <= Len(strSource)
                    If Not blnReadItem Then '当前是否正在读取检验项目
                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "O" Then Exit Do
                        Loop
                        If iStrPoint > Len(strSource) Then
                            '当前信息无效，不再处理。
                            Exit Do
                        End If
                        
                        aItemData = Split(strBuffer, strFldDelim)
                        
                        '样本序号，必须为数字
                        strSample = Val(aItemData(2))
                        'aItemData = Split(strSample, strCompDelim)
                        'strSample = Val(Trim(aItemData(2)))
                        
                        blnReadItem = True '标本号结束，转结果
                        '准备接收
                    Else
                        Do While iStrPoint <= Len(strSource)
                            iPos = iStrPoint '保留当前位置
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "R" Then
                            '检验结果,第一行要取时间
                               If blnFirst Then
                                    aItemData = Split(strBuffer, strFldDelim)
                                   '检验时间,时间格式必须为：yyyy-mm-dd hh:mm:ss
                                    strData = aItemData(12)
                                    strData = Format(Mid(strData, 1, 4) & "-" & Mid(strData, 5, 2) & "-" & Mid(strData, 7, 2) & " " & Mid(strData, 9, 2) & ":" & Mid(strData, 11, 2) & ":" & Mid(strData, 13, 2), "yyyy-mm-dd hh:mm:ss")
                                    strData = strData & "|" & strSample
                                                            '检验人
                                    strData = strData & "| " '本仪器不支持
                                    '标本
                                    strData = strData & "|血液"
                                    '质控品
                                    strData = strData & "|0"
                                    str标本信息 = strData
                                    aItemCan = Split(aItemData(2), strCompDelim)
                                    strChannel = UCase(Trim(aItemCan(3)))
                                    strData = strData & "|" & strChannel
                                    strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                    blnFirst = False
                                Else
                                    aItemData = Split(strBuffer, strFldDelim)
                                     aItemCan = Split(aItemData(2), strCompDelim)
                                    strChannel = UCase(Trim(aItemCan(3)))
                                   ' If strChannel > 24 Then Exit Do
                                   '图形数据
                                    Select Case strChannel
                                        Case UCase("WBC Histo")
'                                            strData = strData & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_WBC = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case UCase("RBC Histo")
'                                            strData = strData & "||" & str标本信息 & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_RBC = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case UCase("Plt Raw Histo")
                                            str_PLT_R = aItemData(3)
                                        Case UCase("Plt Fit Histo")
'                                            strData = strData & "||" & str标本信息 & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_PLT_F = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case Else
                                            strData = strData & "|" & strChannel
                                            StrUnit = UCase(aItemData(4))
                                            strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                    End Select
                                   
                                End If
                            Else
                                If Mid(strBuffer, 3, 1) <> "|" And InStr(strBuffer, "|") > 0 Then
                                    If str_WBC <> "" And str_RBC = "" Then str_WBC = str_WBC & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                    If str_RBC <> "" And str_PLT_F = "" Then str_RBC = str_RBC & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                    If str_PLT_F <> "" Then str_PLT_F = str_PLT_F & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                End If
                            End If
                            iCurrGrpEntry = 0
                        Loop
                        '一个标本结束
                        If str_WBC <> "" Then
                            Str_Img = Str_Img & "^WBC;0;200,200,19.90625-50,39.8125-100,59.71875-150,79.625-200,99.53125-250,119.4375-300,139.34375-350,159.25-400,179.15625-" & de_code(str_WBC)
                        End If
                        If str_RBC <> "" Then
                            Str_Img = Str_Img & "^RBC;0;200,185,18.75-50,56.25-100,112.5-200,168.75-250" & de_code(str_RBC)
                        End If
                        If str_PLT_R <> "" And str_PLT_F <> "" Then
                            Str_Img = Str_Img & "^PLT;4;180,180,5-2,12.5-5,25-10,37.5-15,50-,62.5-25,75-30;0,0,0,0,0," & Replace(Mid(de_code(str_PLT_R), 2), ";", ",") & "|" & Replace(Mid(de_code(str_PLT_F), 2), ";", ",")
                        End If
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData & _
                                    IIf(Str_Img <> "", vbCrLf & Mid(Str_Img, 2), "")
                        
                    End If
                Loop
            Else
                iStrPoint = iPos
                iCurrGrpEntry = 1
                Exit Do
            End If
        Else
          iStrPoint = iPos
        End If
    Loop
            
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '向设备发送‘数据收到’命令
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
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
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = CStr(Hex(lngAsc Mod 256))
End Function

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

Private Function to_Hex(str记录 As String) As String
Dim i  As Integer
    to_Hex = ""
    For i = 1 To Len(str记录)
            to_Hex = to_Hex & Hex(Asc(Mid(str记录, i, 1)))
    Next
End Function







