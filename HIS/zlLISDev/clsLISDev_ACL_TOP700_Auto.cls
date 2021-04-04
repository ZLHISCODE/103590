VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL_TOP700_Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:贝克曼ACL700TOP血凝仪
'用户:湖北省丹江口市第一医院
'作者:Zwy
'双向
Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤

Private mlngStep_Out As Long  '传双向消息步骤
Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.

Private mstrWaitSend As String '待发送标本

Private mstrIdCode As String '消息ID

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)

    Dim lng_S As Long    '开始
    Dim lng_E As Long    '结束
    Dim lng_SS As Long, lng_ss1 As Long
    '一个完整的记录
    Dim strNo As String, strDate As String, strTestType As String, strItem As String, strValues As String
    Dim str_Line As String, strTmp As String
    Dim lngNo As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, strValueEnd As String
    Dim strCode As String
    Dim lngComType As Long '0-单向模式 1-双向模式
    Dim var_Code As Variant
    Dim m As Long
    
    
    strTestType = " |血液|0"
    On Error GoTo errH
    '-----应答
    strSource = Replace(strSource, Chr(&H4) & Chr(&H5), Chr(&H5))
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
        '    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        '        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If strSource <> "" Then strCmd = Chr(&H6)

    mstrCode = mstrCode & strSource
    '---取完整的数据
    lng_S = InStr(mstrCode, Chr(&H5))   '以&H5开头，返回完整段落的开始位置
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '以&H4结束，返回完整段落的结束位置
    Else
        lng_E = 0
    End If


    Do While lng_S > 0 And lng_E > 0    '如果开始、结束都不为空
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)    '返回从开始位置到结束位置之间的字符
        mstrCode = Mid(mstrCode, lng_E + 1)
        Do While InStr(str_Code, Chr(&H17)) > 0
            lng_SS = InStr(str_Code, Chr(&H17))
            str_Code = Mid(str_Code, 1, lng_SS - 1) & Mid(str_Code, lng_SS + 7)    '取7位校验码
        Loop

        var_data = Split(str_Code, Chr(&HD))    '取每段数据
        mstrIdCode = ""
        For i = LBound(var_data) To UBound(var_data)
            str_Line = var_data(i)

            If InStr(str_Line, "H|") > 0 Then
                strDate = Trim(Split(str_Line, "|")(13))    '获取日期时间
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid( _
                         strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, _
                                                                             2) & ":" & Mid(strDate, 13, 2)    '拆分日期出来“yyyy-mm-dd”
                strDate = Format(strDate, "yyyy-MM-dd hh:mm:ss")
                If Not IsDate(strDate) Then    '如果日期为空，则显示为当前日期
                    strDate = Format(Now, "yyyy-MM-dd hh:mm:ss")
                End If
                
                '获取消息ID，用于后面通信
                'mstrIdCode = Split(str_Line, "|")(2)
                
            ElseIf InStr(str_Line, "O|") = 1 Then
                If Len(Trim(Split(str_Line, "|")(2))) < 8 Then
                    strNo = Trim(Split(str_Line, "|")(2))
                    strCode = ""
                Else
                    strNo = ""
                    strCode = Trim(Split(str_Line, "|")(2))
                End If
            '增加双向部分的处理,安小龙改动,等于4改为等于1
            ElseIf InStr(str_Line, "Q|") = 1 Then
                '判断收到双向指令后，做如下处理
                
'                H | @ ^ \ | <0_0> <1025080549_50> | | ACL-TOP-03 | | | | | LIS-HOST-04 | | P | 1394-97 | 19990913174650 <CR>
'                Q | 1 | ^ 4243 ^ 876271 @ ^ 0434 @ ^ 0435 @ ^ 6742 ^ 878432 | | | | | | | | | | O@N <CR>
'                L | 1 | N <CR>
                '可能存在多个标本查询的情况
                strNo = Split(Trim(Split(str_Line, "|")(2)), "^")(0)
                
                var_Code = Split(Trim(Split(str_Line, "|")(2)), "@")
                For m = LBound(var_Code) To UBound(var_Code)
                    strCode = Trim(Split(var_Code(m), "^")(1))
                    
                    If strCode <> "" Then
                        If Len(strCode) > 11 Then
                            strValueEnd = strValueEnd & "||" & Format(strDate, "yyyy-MM-dd") & "|^^" & strCode
                        Else
                            strValueEnd = strValueEnd & "||" & Format(strDate, "yyyy-MM-dd") & "|" & strCode & "^^"
                        End If
                    End If
                    
                Next
                
                lngComType = 1
                
            ElseIf InStr(str_Line, "R|") = 1 Then    '取分组的检验项目


                strItem = Split(Split(str_Line, "|")(2), "^")(3)
                If strItem = 131 Or strItem = 2011 Then
                    strItem = strItem & Split(str_Line, "|")(4)
                End If
                strValues = Trim(Split(str_Line, "|")(3))

                '最后结果拼接起来
                If (Val(strNo) > 0 Or Val(strCode) > 0) And strItem <> "" And (strValues <> "" And strValues <> "FAILED") Then
                    strNo = Replace(strNo, "^^", "") & "^^" & strCode
                    strValueEnd = strValueEnd & "||" & strDate & "|" & strNo & "|" & strTestType & "|" & strItem & "|" & strValues
                End If
                
            End If

        Next

        lng_S = InStr(mstrCode, Chr(&H5))   '以&H5开头，返回完整段落的开始位置
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '以&H4结束，返回完整段落的结束位置
        Else
            lng_E = 0
        End If
        
    Loop

    Select Case lngComType
        Case 0
            '安小龙改动,原strCmd = ""
            strCmd = strCmd
        Case 1
            strCmd = "1|" & strCmd
    End Select
    
    strResult = Mid(strValueEnd, 3)
    
    Exit Sub
errH:
End Sub

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
    Dim var_Tamp1 As Variant    '标本
    Dim i_Count As Integer, i_Loop As Integer
    Dim strSample As String, strNo As String
    Dim IntJizhen As Integer    '急诊
    Dim strCode As String, strItem As String
    Dim strTestType As String
    Dim str_SendItem As String, str_Line As String
    Dim FN As Integer
    
    var_Tamp1 = Split(strSampleInfo, "||")
    SendSample = ""
    
    blnSuccess = True '设为不取消发送
    If iSendStep = 0 Then
        For i_Count = LBound(var_Tamp1) To UBound(var_Tamp1)
            strSample = var_Tamp1(i_Count)
            If Val(Split(strSample, "|")(1)) = 0 Then
                strNo = Trim(Split(strSample, "|")(5))
            Else
                strNo = Val(Split(strSample, "|")(1))
            End If
            strTestType = Split(strSample, "|")(3)
            IntJizhen = Val(Split(strSample, "|")(4))
            strCode = Trim(Split(strSample, "|")(5))
            strItem = ""
            For i_Loop = 10 To UBound(Split(strSample, "|"))
                If Val(Split(strSample, "|")(i_Loop)) > 0 And Trim(Split(strSample, "|")(i_Loop)) <> "131INR" And Trim(Split(strSample, "|")(i_Loop)) <> "131s" And Trim(Split(strSample, "|")(i_Loop)) <> "131%" Then
                        strItem = strItem & "^^^" & Val(Split(strSample, "|")(i_Loop)) & "@"
                End If
            Next
            
            '双向模式下，仪器输标本号的情况
            If strCode = "" Then
                strCode = strNo
            End If
            
            If strItem <> "" And strCode <> "" Then
                strItem = Mid(strItem, 1, Len(strItem) - 1)
                If i_Count = LBound(var_Tamp1) Then
                    '第一段
                    FN = 0
                    FN = IIf(FN = 7, 0, FN + 1)
                    'strSample
                    mstrIdCode = "<0_0><" & Mid(Format(Split(Now, "|")(0), "yyyymmddhhmmss"), 3, 10) & "_" & Mid(Format(Split(Now, "|")(0), "yyyymmddhhmmss"), 13, 2) & ">"
                    str_Line = FN & "H|@^\|" & mstrIdCode & "||LIS|||||ACL TOP||P|1394-97|" & Format(Split(Now, "|")(0), "yyyymmddhhmmss") & Chr(&HD) & Chr(&H17)
                    str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                End If
                    '第二段
                    FN = IIf(FN = 7, 0, FN + 1)
                    str_Line = FN & "P|" & i_Count + 1 & "||" & strCode & "||^" & Split(Split(strSample, "|")(7), "^")(3) & "||19800228|" & IIf(Split(Split(strSample, "|")(7), "^")(0) = "男", "M", "F") & "|||||542" & Chr(&HD) & Chr(&H17)
                    str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                    '第三段
                    FN = IIf(FN = 7, 0, FN + 1)
                    str_Line = FN & "O|1|" & strCode & "|" & "" & "|" & strItem & "|" & IIf(IntJizhen = 1, "S", "R") & "|" & Format(Split(Now, "|")(0), "yyyymmddhhmmss") & "|||||A||||P||||||||||Q" & Chr(&HD) & Chr(&H17)
                    str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
                If i_Count = UBound(var_Tamp1) Then
                    '第四段
                    FN = IIf(FN = 7, 0, FN + 1)
                    str_Line = FN & "L|1|F" & Chr(&HD) & Chr(&H3)
                    str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA)
                    
                    If SendSample = "" Then
                        '加上 消息标志(05，04)
                        SendSample = SendSample & Chr(&H5)
                        iSendStep = 1 '开始通信
                        mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)
                    Else
                        mstrWaitSend = mstrWaitSend & Chr(&H5) & str_SendItem & Chr(&H4)
                    End If
                End If

            End If
        Next
    Else
        If InStr(strResponse, Chr(&H6)) > 0 Then
            '收到06，将下一条要发送到仪器的指令取出，放到SendSample
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                If InStr(mstrWaitSend, ",") > 0 Then
                    SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
                    SendSample = mstrWaitSend
                    mstrWaitSend = ""
                End If
            End If
        Else
         iSendStep = 0
        End If
        '没有待发送指令，退出双向通讯。不能忘了
        If mstrWaitSend = "" Then iSendStep = 0
        Exit Function
    End If
End Function


'获取校验码
Private Function GetCheckCode(ByVal str_in As String) As String
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function


''''当仪器手动或者自动下载LIS这边的订单请求时
'H | @ ^ \ | <0_0> <1025080549_50> | | ACL-TOP-07 | | | | | LIS-HOST-06 | | P | 1394-97 | 20000913174650 <CR>
'Q | 1 | ALL | | | | | | | | | | O@N <CR>
'L | 1 | N <CR>

'此时主机如果拒绝传输可以应答
'H | @ ^ \ | <0_0> <1025080549_50> | | ACL-TOP-11 | | | | | LIS-HOST-93 | | P | 1394-97 | 20010414182154 <CR>
'Q | 1 | | | | | | | | | | | A <CR>
'C | 1 | I | timeout ^ last request has been cancelled | P <CR>
'L | 1 | N <CR>


'如果主动传输时可以应答（查询）
'H | @ ^ \ |<0_0> <1025080549_50> | | LIS-HOST-03 | | | | | ACL-TOP-21 | | P | 1394-97 | 20000614080500 <CR>
'P | 1 | | 80501 | | Anderson ^ Jim | | 19800228 | M | | | | | 542 <CR>
'O | 1 | 6483 | | ^ ^ ^ 211 @ ^ ^ ^ 063 | R | 20000614043211 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'O | 2 | 8651 | | ^ ^ ^ 310 | S | 20000614043514 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'P | 2 | | | | Foster ^ Rene | | 19740717 | F | | | | | 374 <CR>
'O | 1 | 0012 | | ^ ^ ^ 512 | R | 20000614044212 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'O | 2 | 0012 | | ^ ^. ^ 254 | R | 20000614044325 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'P | 3 | | | | Clinton ^ Jeremy | | 19580223 | M <CR>
'O | 1 | 6537 | 310648 | ^ ^ ^ 076 | S | 20000614044212 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'O | 2 | 6537 | 310648 | ^ ^ ^ 098 | R | 20000614045021 | | | | | C | | | | P | | | | | | | | | | Q <CR>
'O | 3 | 6912 | | ^ ^ ^ 301 @ ^ ^ ^ 357 | R | 20000614052158 | | | | | A | | | | P | | | | | | | | | | Q <CR>
'P | 4 | | | | Turner ^ Jim | | 19890821 | M | | | | | 271 <CR>
'O | 1 | 0509 | 479481 | ^ ^ ^ 002 | | 20000614055734 | | | | | C | | | | P | | | | | | | | | | Q <CR>
'L | 1 | F <CR>

'其中需要注意的是：P代表病人信息（P|序号|病人标识号||姓^名||出生日期|性别|||||医生ID）
'O代表订单信息(O|序号|标本号|仪器标本ID|^^^通道码1@^^^通道码2|处理模式(S or R)|检验时间|||||A||||P||||||||||Q)


'仪器返回标本在仪器上的仪器标本ID
'H | @ ^ \ | <0_0> <1025080549_50> | | ACL-TOP-21 | | | | | LIS-HOST-03 | | P | 1394-97 | 20000614080624 <CR>
'P | 1 | | | | Anderson ^ Jim | | 19800228 | M | | | | | 542 <CR>
'O | 1 | 6483 | 312890 | ^ ^ ^ 063 | R | 20000614043211 | | | | | P | | | | P | | | | | | | | | | O@I <CR>
'P | 2 | | | | Foster ^ Rene | | 19740717 | F | | | | | 374 <CR>
'O | 1 | 0012 | 312892 | ^ ^ ^ 512 | R | 20000614044212 | | | | | P | | | | P | | | | | | | | | | O@I <CR>
'L | 1 | N <CR>


'H | @ ^ \ | <0_0> <1025080549_50> | | LIS-HOST-04| | | | | ACL-TOP-12 | | P | 1394-97 | 20030330033003 <CR>
'L | 1 | I <CR>


