VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cobas6000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:临床免疫分析系 Cobas 6000
'厂商:Roche 罗氏
'作者:陈东
'用户:广西民族医院

Private mstrCode     As String    '待解析串

Private mstrRack     As String    '扫条码时传的盘号,架号等需要返回的信息.

Private mstrWaitSend As String '待发送标本

Private mstrQCNO     As String      '存QC编号，要转为固定的标本号

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

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
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
    Dim lng_S            As Long, lng_E As Long, i_Count As Integer

    Dim str_In           As String, str_Line As String, str_Bacode As String, var_line As Variant

    Dim str_Date         As String, str_SampleNo As String, str_Item As String, str_Values As String

    Dim str_WaitSendLine As String, str_JiZhen As String, str_QC As String

    Dim str_tmp          As String

    '-----------------------------------------------------------------------------
    Dim varQC            As Variant, strIniFile As String, strQC As String, intLoop As Integer

    Dim objFileSystem    As New Scripting.FileSystemObject

    Dim objStream        As Scripting.TextStream, strLine As String
    
    str_In = strSource
    '-----------------------------------------------------------------------------------
    '--   收到仪器的应答，发送未发完的数据
    strCmd = ""

    '-----------------------------------------------------------------------------------
    
    If InStr(str_In, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_In, Chr(&H17)) > 0 Then                  '收到ETB(17) 应答
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_In, Chr(&H3)) > 0 Then                 '收到ETX(03)
        strCmd = Chr(&H6)
    End If

    If Len(str_In) > 0 Then mstrCode = mstrCode & str_In
    
    '去掉
    Do While mstrCode Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
        lng_S = InStr(mstrCode, Chr(&H17))
        lng_E = InStr(lng_S, mstrCode, Chr(&H2)) + 1
        mstrCode = Mid(mstrCode, 1, lng_S - 1) & Mid(mstrCode, lng_E + 1)
    Loop
    
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))

    If lng_S > 0 And lng_E > 0 Then
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If InStr(str_Line, "|TSREQ^REAL|") > 0 Then         '双向通讯- 开始
            mstrRack = ""
            var_line = Split(str_Line, Chr(&HD))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)

                If str_Line Like "Q|*" Then
                    str_Bacode = Trim(Split(str_Line, "|")(2))
                    mstrRack = Split(str_Bacode, "^")(3) & "^" & Split(str_Bacode, "^")(4) & "^" & Split(str_Bacode, "^")(5) & "^^" & Split(str_Bacode, "^")(7) & "^" & Split(str_Bacode, "^")(8)
                    str_Bacode = Split(str_Bacode, "^")(2)
                    
                    If Len(Trim(str_Bacode)) >= 12 And Val(str_Bacode) > 0 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_Bacode)
                    End If
                End If

            Next
            
            If strResult <> "" Then
                strResult = Mid(strResult, 3)
                strCmd = "1|"    '双向通讯请求
            End If

        ElseIf InStr(str_Line, "|RSUPL^") > 0 Then     '检验结果

            '-----------------------------------------------
            strIniFile = App.Path & "\Cobas6000.ini"
            strQC = ""

'            If objFileSystem.FileExists(strIniFile) Then
'                Set objStream = objFileSystem.OpenTextFile(strIniFile)
'
'                Do While Not objStream.AtEndOfStream
'                    strLine = objStream.ReadLine
'
'                    If InStr(strLine, "=") > 0 Then
'                        strQC = strQC & "|" & strLine
'                    End If
'
'                Loop

'            Else
'                Call objFileSystem.CreateTextFile(strIniFile)
'                Set objStream = objFileSystem.OpenTextFile(strIniFile, ForAppending)
'                objStream.WriteLine "[水平1]"
'                objStream.WriteLine "5200 = PPU"
'                objStream.WriteLine "[水平2]"
'                objStream.WriteLine "6100 = PNU"
'
'                strQC = "5200 = PPU|6100 = PNU"
'            End If

            If strQC <> "" Then varQC = Split(Mid(strQC, 2), "|")
            '-----------------------------------------------------------------------------
            var_line = Split(str_Line, Chr(&HD))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)

                If str_Line Like "O|*" Then
                    str_Date = Trim(Split(str_Line, "|")(7))

                    If Len(str_Date) <> 14 Then str_Date = Split(str_Line, "|")(22)
                    str_Date = Left(str_Date, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2)
                    
                    str_SampleNo = Split(str_Line, "|")(2)

                    If Trim(str_SampleNo) = "" Then str_SampleNo = " " & Trim(Split(Split(str_Line, "|")(3), "^")(0))
                    If Left(str_SampleNo, 1) <> " " Then str_SampleNo = " " & str_SampleNo
                    str_JiZhen = Trim(Split(str_Line, "|")(5))
                    str_QC = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")
                    
'                    If str_QC = "1" Then
'                        If strQC <> "" Then
'
'                            For intLoop = LBound(varQC) To UBound(varQC)
'
'                                If InStr("," & UCase(Trim(Split(varQC(intLoop), "=")(1))) & ",", "," & UCase(Trim(str_SampleNo)) & ",") > 0 Then
'                                    str_SampleNo = " " & Val(Split(varQC(intLoop), "=")(0))
'
'                                    Exit For
'
'                                End If
'
'                            Next
'
'                        Else
'
'                            If UCase(Trim(str_SampleNo)) = "PPU" Then
'                                str_SampleNo = "5200" '& Split(Trim(Split(str_Line, "|")(3)), "^")(0)
'
'                            ElseIf UCase(Trim(str_SampleNo)) = "PNU" Then
'
'                                str_SampleNo = "6100" '& Split(Trim(Split(str_Line, "|")(3)), "^")(0)
'                            End If
'                        End If
'
'                        If Val(str_SampleNo) <= 0 Then
'                            str_SampleNo = " " & Split(Trim(Split(str_Line, "|")(3)), "^")(0)
'                        End If
'                    End If
                    
                    If Trim(str_SampleNo) <> "" Then
                        If Left(str_SampleNo, 1) = " " And Len(Trim(str_SampleNo)) < 12 Then
                            '标本号
                            str_SampleNo = Trim(str_SampleNo) & IIf(str_JiZhen = "S", "^1", "")
                            
                        Else
                            '条码
                            str_SampleNo = "^" & IIf(str_JiZhen = "S", "^1", "^") & Trim(str_SampleNo)
                        End If

                    Else
                        str_SampleNo = ""
                    End If
                    
                    str_Item = ""

                ElseIf str_Line Like "R|*" Then

                    str_Values = Trim(Split(str_Line, "|")(3))

                    If str_Values <> "" Then
                        If InStr(str_Values, "^") > 0 Then str_Values = Split(str_Values, "^")(1) '--- 090407 有免疫的结果，为 -1^0.715 的格式。
                        str_tmp = Replace(Split(str_Line, "|")(2), "^^^", "")

                        If InStr(str_tmp, "/") > 0 Then str_tmp = Split(str_tmp, "/")(0)
                        If InStr(Split(str_Line, "|")(6), "HH") > 0 Then str_Values = ">" & str_Values
                        If InStr(Split(str_Line, "|")(6), "LL") > 0 Then str_Values = "<" & str_Values
                        str_Item = str_Item & "|" & str_tmp & "|" & str_Values
                    End If
                End If

            Next

            If IsDate(CDate(str_Date)) And str_SampleNo <> "" And str_Item <> "" Then
                strResult = str_Date & "|" & str_SampleNo & "| |血清|" & str_QC & str_Item
                WriteLog "收到仪器传回的结果 ", strResult, ""
                '加上 计算项目
                var_line = Split(Mid(str_Item, 2), "|")
                
                For intLoop = LBound(var_line) To UBound(var_line)

                    If Trim(var_line(intLoop)) = "320" Then 'Cysc
                        If Val(var_line(intLoop + 1)) <> 0 Then
                            strResult = strResult & "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4)
                            WriteLog "计算结果 ", "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4), ""

                            Exit For

                        End If
                    End If

                Next

            End If

        End If
    End If
    
End Sub

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
    
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个无素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。

    Dim var_Sample   As Variant, i_Count As Integer, i_Loop As Integer, str_Record As String, str_SampleNo As String

    Dim i_JiZhen     As Integer, str_Bacode As String, str_CheckItem As String, str_SampleType As String

    Dim str_SendItem As String, str_EndString As String, str_tmp As String, lngRow As Long

    Dim lng_S        As Long, str_In As String
    
    If strSampleInfo <> "" Then WriteLog "双向通讯(Cobas6000 Begin)：", strSampleInfo, mstrWaitSend
    var_Sample = Split(strSampleInfo, "||")
    blnSuccess = True

    If iSendStep = 0 Then
        SendSample = ""

        For i_Count = LBound(var_Sample) To UBound(var_Sample)
            str_Record = var_Sample(i_Count)
            
            If Val(Split(str_Record, "|")(1)) = 0 Then
                str_SampleNo = Trim(Split(str_Record, "|")(5))
            Else
                str_SampleNo = Val(Split(str_Record, "|")(1))
            End If

            str_SampleType = Split(str_Record, "|")(3)
            i_JiZhen = Val(Split(str_Record, "|")(4))
            str_Bacode = Trim(Split(str_Record, "|")(5))
            str_CheckItem = ""
            
            For i_Loop = 10 To UBound(Split(str_Record, "|"))

                If Val(Split(str_Record, "|")(i_Loop)) > 0 Then
                    If InStr(str_CheckItem, "^^^" & Replace(Replace(Split(str_Record, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\") <= 0 Then
                        '两个血糖，只传一个
                        str_CheckItem = str_CheckItem & "^^^" & Replace(Replace(Split(str_Record, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\"
                    End If
                End If

            Next
            
            If str_CheckItem <> "" And str_Bacode <> "" Then
                str_CheckItem = Mid(str_CheckItem, 1, Len(str_CheckItem) - 1)
                
                str_SendItem = "H|\^&|||H7600^1||||||TSDWN^REPLY" & Chr(&HD) & "P|1" & Chr(&HD)
                str_SendItem = str_SendItem & "O|1|" & str_Bacode & Space(22 - Len(str_Bacode))
                str_SendItem = str_SendItem & "|" & mstrRack & "|" & str_CheckItem
                str_SendItem = str_SendItem & "|" & IIf(i_JiZhen = 1, "S", "R")               'S-急诊,R-常规
                str_SendItem = str_SendItem & "||" & Format(Now, "yyyyMMddHHmmss")
                str_SendItem = str_SendItem & "||||N||||1||||||||||O" & Chr(&HD)
                str_SendItem = str_SendItem & "C|1|L|^^^^|G" & Chr(&HD)
                str_SendItem = str_SendItem & "L|1|N" & Chr(&HD) & Chr(&H3)
                          
                str_EndString = ""
                lngRow = 0

                Do While str_SendItem <> ""
                    
                    If Len(str_SendItem) <= 230 Then
                        lngRow = lngRow + 1
                        str_EndString = str_EndString & Chr(&H2) & lngRow & str_SendItem & GetCheckCode(lngRow & str_SendItem) & vbCrLf & ","
                        str_SendItem = ""
                    Else
                        lngRow = lngRow + 1
                        str_tmp = Mid(str_SendItem, 1, 230) & Chr(&H17)
                        str_EndString = str_EndString & Chr(&H2) & lngRow & str_tmp & GetCheckCode(lngRow & str_tmp) & vbCrLf & ","
                        str_SendItem = Mid(str_SendItem, 231)
                    End If

                Loop
    
                If SendSample = "" Then
                    SendSample = Chr(&H5)
                    iSendStep = iSendStep + 1
                    mstrWaitSend = mstrWaitSend & str_EndString & Chr(&H4) '加上 消息结束标志(EOT)
                Else
                    mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_EndString & Chr(&H4)
                End If
            End If

        Next

        If strSampleInfo <> "" Then WriteLog "双向通讯(Cobas6000 End)：", SendSample, mstrWaitSend
    Else
        str_In = strResponse
        SendSample = ""

        Do While InStr(str_In, Chr(&H6)) > 0                   '收到ACK(06)
            lng_S = InStr(str_In, Chr(&H6))
            str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)

            If mstrWaitSend <> "" Then
                
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                
                If InStr(mstrWaitSend, ",") > 0 Then
                    SendSample = SendSample & Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
                    SendSample = SendSample & mstrWaitSend
                    mstrWaitSend = ""
                End If
                
            End If
    
        Loop
        
        If mstrWaitSend = "" Then iSendStep = 0
        
    End If
    
End Function

Private Function GetCheckCode(ByVal str_In As String) As String

    '获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_In)
        lngAsc = lngAsc + Asc(Mid(str_In, i, 1))
    Next

    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

