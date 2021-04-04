VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CoBas8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:临床免疫分析系 Cobas 8000
'厂商:Roche 罗氏
'作者:陈德荣
'用户:阆中市人民医院
'修改:蒋淑君于2013年9月26日修改双向代码,SendSample函数全部代码

Private mstrCode As String    '待解析串

Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.

Private mstrWaitSend As String    '待发送标本

Private mstrQCNO As String      '存QC编号，要转为固定的标本号

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
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_条码 As String, var_line As Variant
    Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
    Dim str_WaitSendLine As String, str_急诊 As String, str_质控 As String
    Dim str_Tmp As String
    '-----------------------------------------------------------------------------
    Dim varQC As Variant, strIniFile As String, strQC As String, intLoop As Integer
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    str_in = str_in & strSource
    '    Call writelislog("strsource", strSource)
    '-----------------------------------------------------------------------------------
    '--   收到仪器的应答，发送未发完的数据
    strCmd = ""
    Do While InStr(str_in, Chr(&H6)) > 0                   '收到ACK(06)
        lng_S = InStr(str_in, Chr(&H6))
        str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)

        If mstrWaitSend <> "" Then
            If Left(mstrWaitSend, 1) = "," Then
                mstrWaitSend = Mid(mstrWaitSend, 2)
            End If
            If InStr(mstrWaitSend, ",") > 0 Then
                str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
            Else
                str_WaitSendLine = mstrWaitSend
                mstrWaitSend = ""
            End If

            If str_WaitSendLine <> "" Then
                If strCmd = "" Then strCmd = "1|"
                strCmd = strCmd & str_WaitSendLine
            End If
        End If

    Loop
    '-----------------------------------------------------------------------------------

    Debug.Print strCmd
    If InStr(str_in, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
        strCmd = Chr(&H6)
    End If

    If InStr(str_in, Chr(&H17)) > 0 Then                  '收到ETB(17) 应答
        strCmd = Chr(&H6)
    End If

    If InStr(str_in, Chr(&H3)) > 0 Then                 '收到ETX(03)
        strCmd = Chr(&H6)
    End If
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in

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

        If InStr(str_Line, "|TSREQ|") > 0 Then         '双向通讯- 开始
            mstrRack = ""
            var_line = Split(str_Line, Chr(&HD))

            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If InStr(str_Line, "Q|1") > 0 Then
                    '                If str_Line Like "Q|*" Then
                    str_条码 = Trim(Split(str_Line, "|")(2))
                    mstrRack = Split(str_条码, "^")(3) & "^" & Split(str_条码, "^")(4) & "^" & Split(str_条码, "^")(5) & "^^" & Split(str_条码, "^")(7) & "^" & Split(str_条码, "^")(8) & "^" & Split(str_条码, "^")(9)
                    str_条码 = Split(str_条码, "^")(2)


                    If Len(Trim(str_条码)) >= 12 And Val(str_条码) > 0 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_条码)
                    End If
                End If
            Next
            '            Call writelislog("str_条码", str_条码)
            If strResult <> "" Then
                strResult = Mid(strResult, 3)
                '                Call writelislog("strResult", strResult)
                strCmd = "1|"    '双向通讯请求
                '                Call writelislog("strCmd", strCmd)
            End If

        ElseIf InStr(str_Line, "|RSUPL") > 0 Then     '检验结果
            '-----------------------------------------------
            strIniFile = App.Path & "\Cobas8000.ini"
            strQC = ""
            If objFileSystem.FileExists(strIniFile) Then
                Set objStream = objFileSystem.OpenTextFile(strIniFile)
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine
                    If InStr(strLine, "=") > 0 Then
                        strQC = strQC & "|" & strLine
                    End If
                Loop
            Else
                Call objFileSystem.CreateTextFile(strIniFile)
                Set objStream = objFileSystem.OpenTextFile(strIniFile, ForAppending)
                objStream.WriteLine "[水平1]"
                objStream.WriteLine "5200 = PPU"
                objStream.WriteLine "[水平2]"
                objStream.WriteLine "6100 = PNU"

                strQC = "5200 = PPU|6100 = PNU"
            End If
            If strQC <> "" Then varQC = Split(Mid(strQC, 2), "|")
            '-----------------------------------------------------------------------------
            var_line = Split(str_Line, Chr(&HD))

            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If InStr(str_Line, "O|") > 0 Then
                    '                If str_Line Like "O|*" Then
                    str_急诊 = Trim(Split(str_Line, "|")(5))
                    str_质控 = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")
                    str_日期 = Trim(Split(str_Line, "|")(7))   '20120327154714
                    If Len(str_日期) <> 14 Then str_日期 = Split(str_Line, "|")(22)
                    str_日期 = Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & _
                             Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)

                    str_标本号 = Split(str_Line, "|")(2)
                    If Trim(str_标本号) = "" Then str_标本号 = " " & Trim(Split(Split(str_Line, "|")(3), "^")(0))

                    If str_质控 = "1" Then
                        If strQC <> "" Then
                            For intLoop = LBound(varQC) To UBound(varQC)
                                If InStr("," & UCase(Trim(Split(varQC(intLoop), "=")(1))) & ",", "," & UCase(Trim(Split(str_标本号, "^")(0))) & ",") > 0 Then
                                    str_标本号 = " " & Val(Split(varQC(intLoop), "=")(0))
                                    Exit For
                                End If
                            Next
                        Else
                            If UCase(Trim(str_标本号)) = "PPU" Then
                                str_标本号 = "5200"    '& Split(Trim(Split(str_Line, "|")(3)), "^")(0)
                            ElseIf UCase(Trim(str_标本号)) = "PNU" Then
                                str_标本号 = "6100"    '& Split(Trim(Split(str_Line, "|")(3)), "^")(0)
                            End If
                        End If
                    End If
                    If Val(Trim(str_标本号)) > 0 Then
                        If Len(Trim(str_标本号)) < 12 Then
                            '                        If Left(str_标本号, 1) = " " And Len(Trim(str_标本号)) < 12 Then
                            '标本号
                            str_标本号 = IIf(str_急诊 = "S", 3000 + Val(Trim(str_标本号)), Trim(str_标本号))
                        Else
                            '条码
                            str_标本号 = "^" & IIf(str_急诊 = "S", "^1", "^") & Trim(str_标本号)
                        End If
                    Else
                        str_标本号 = ""
                    End If

                    str_项目 = ""
                ElseIf InStr(str_Line, "R|") > 0 Then
                    If str_质控 = "1" And Not IsDate(CDate(str_日期)) Then
                        str_日期 = Trim(Split(str_Line, "|")(12))
                        str_日期 = Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & _
                                 Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                    End If
                    str_结果 = Trim(Split(str_Line, "|")(3))
                    If str_结果 <> "" Then
                        If InStr(str_结果, "^") > 0 Then str_结果 = Split(str_结果, "^")(1)    '--- 090407 有免疫的结果，为 -1^0.715 的格式。
                        str_Tmp = Replace(Split(str_Line, "|")(2), "^^^", "")
                        If InStr(str_Tmp, "/") > 0 Then str_Tmp = Split(str_Tmp, "/")(0)
                        str_项目 = str_项目 & "|" & str_Tmp & "|" & str_结果
                    End If
                End If
            Next
            If IsDate(CDate(str_日期)) And str_标本号 <> "" And str_项目 <> "" Then
                strResult = str_日期 & "|" & str_标本号 & "| |血清|" & str_质控 & str_项目
                WriteLog "收到仪器传回的结果 ", strResult, ""
                '加上 计算项目
                var_line = Split(Mid(str_项目, 2), "|")

                For intLoop = LBound(var_line) To UBound(var_line)
                    If Trim(var_line(intLoop)) = "320" Then    'Cysc
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


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

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

    Dim var_标本 As Variant, i_Count As Integer, i_Loop As Integer, str_记录 As String, str_标本号 As String
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String
    Dim str_SendItem As String, str_EndString As String, str_Tmp As String, lngRow As Long
    '变量用于处理冲突
    Static saveSend As String

    If mstrWaitSend = "" Then
        mstrWaitSend = saveSend
    End If

    '    Call writelislog("SendSample0", "")

    If strSampleInfo <> "" Then WriteLog "双向通讯(Cobas8000 Begin)：", strSampleInfo, mstrWaitSend
    var_标本 = Split(strSampleInfo, "||")
    '    Call writelislog("SendSample1", strSampleInfo)
    SendSample = ""
    blnSuccess = True    '设为不取消发送
    '    Call writelislog("iSendStep", CStr(iSendStep))

    If iSendStep = 0 Then
        For i_Count = LBound(var_标本) To UBound(var_标本)
            str_记录 = var_标本(i_Count)

            If Val(Split(str_记录, "|")(1)) = 0 Then
                str_标本号 = Trim(Split(str_记录, "|")(5))
            Else
                str_标本号 = Val(Split(str_记录, "|")(1))
            End If

            '            Call writelislog("str_标本号", str_标本号)

            str_标本类型 = Split(str_记录, "|")(3)
            i_急诊 = Val(Split(str_记录, "|")(4))
            str_条码 = Trim(Split(str_记录, "|")(5))
            str_检验项目 = ""

            For i_Loop = 10 To UBound(Split(str_记录, "|"))
                If Val(Split(str_记录, "|")(i_Loop)) > 0 Then
                    If InStr(str_检验项目, "^^^" & Replace(Replace(Split(str_记录, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\") <= 0 Then
                        '两个血糖，只传一个
                        str_检验项目 = str_检验项目 & "^^^" & Replace(Replace(Split(str_记录, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\"
                    End If
                End If
            Next

            '            Call writelislog("str_检验项目", str_检验项目)

            If str_检验项目 <> "" And str_条码 <> "" Then
                str_检验项目 = Mid(str_检验项目, 1, Len(str_检验项目) - 1)

                str_SendItem = "1H|\^&|||host|||||cobas 8000^1.02|TSUPL|P|1|" & Chr(&HD)
                str_SendItem = str_SendItem & "P|1|" & Chr(&HD)
                str_SendItem = str_SendItem & "O|1|" & str_条码 & "|" & mstrRack & "|" & str_检验项目
                '                str_SendItem = str_SendItem & "|" & IIf(i_急诊 = 1, "S", "R")               'S-急诊,R-常规
                str_SendItem = str_SendItem & "|||||||O||||1||||||||||F|" & Chr(&HD)
                str_SendItem = str_SendItem & "C|1|L|^^^|G|" & Chr(&HD)
                str_SendItem = str_SendItem & "L|1|N" & Chr(&HD) & Chr(&H3)
                '                Call writelislog("str_SendItem", str_SendItem)

                str_SendItem = str_EndString & Chr(&H2) & str_SendItem & GetCheckCode(str_SendItem) & vbCrLf & ","
                '                Call writelislog("str_SendItem4", str_EndString)

                If SendSample = "" Then
                    SendSample = SendSample & Chr(&H5)
                    iSendStep = 1    '开始通信
                    mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)    '加上 消息结束标志(EOT)

                    saveSend = mstrWaitSend
                    '                     Call writelislog("mstrWaitSend1", mstrWaitSend)
                Else
                    mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_SendItem & Chr(&H4)

                    saveSend = mstrWaitSend
                    '                     Call writelislog("mstrWaitSend2", mstrWaitSend)
                End If

                '                Call writelislog("SendSample3", SendSample)
            End If
        Next
    ElseIf (InStr(strResponse, Chr(&H5)) > 0) Then
        SendSample = Chr(&H6)
        saveSend = mstrWaitSend
        mstrWaitSend = ""
    Else
        If InStr(strResponse, Chr(&H6)) > 0 Then    '收到06,将下一条要发送到仪器的指令取出,放到SendSample
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)

                End If
                If InStr(mstrWaitSend, ",") > 0 Then
                    SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))

                    '                   Call writelislog("SendSample4", SendSample)
                    '                   Call writelislog("mstrWaitSend3", mstrWaitSend)
                Else
                    SendSample = mstrWaitSend
                    mstrWaitSend = ""
                    saveSend = ""
                End If
            End If
        End If
        '         Call writelislog("SendSample5", SendSample)
        '没有待发送指令,退出双向通讯.
        If mstrWaitSend = "" Then iSendStep = 0
        Exit Function
    End If
    If strSampleInfo <> "" Then WriteLog "双向通讯(Cobas8000 End)：", SendSample, mstrWaitSend
End Function

Private Function GetCheckCode(ByVal str_in As String) As String
'获取校验码
    Dim lngAsc As Long, i As Integer

    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function






