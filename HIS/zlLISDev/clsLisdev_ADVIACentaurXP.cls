VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisdev_ADVIACentaurXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器名称: 化学发光免疫分析仪
'仪器型号: ADVIA centaur XP
'厂商:德国西门子
'作者:
'用户:丹东中心医院(XP)

Private mstrCode As String  '待解析数据
Private mstrWaitSend As String  '待发送数据

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
    mstrWaitSend = ""
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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim var_Code As Variant, int_Loop As Integer
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim strSampleNo As String, strTestItem As String, strTestValue As String, strDate As String
    Dim str_SendLine As String, str_Item As String
    Dim str_Tmp As String
    On Error GoTo hErr
    strResult = "": strCmd = ""
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
'    If Right(strSource, 1) = Chr(&H4) Then
'        '收到一个完整的数据，这时如果还有未发送数据，进入双向模式
'        If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
'            SaveLog "XP-缓存中有未发送数据，切换到双向模式 ", strSource, mstrWaitSend
'            mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
'            strCmd = "1|"
'            strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
'            Exit Sub
'        ElseIf mstrWaitSend <> "" Then
'            '没有发送完的数据，丢弃
'            SaveLog "XP-缓存中有未发送的不完整数据，丢弃", strSource, mstrWaitSend
'            mstrWaitSend = ""
'        End If
'    End If
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '收到结束符 05 ,应答 06
        strCmd = Chr(&H6)
    ElseIf InStr(mstrCode, Chr(&HA)) > 0 Then                               '收到行结束符 03 ,应答 06
         strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
 
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = 0
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
 
    Do While lng_S > 0 And lng_E > 0    '收到一个完整的记录,开始解析
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        Do While str_Code Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
            lng_S = InStr(str_Code, Chr(&H17))
            lng_E = InStr(lng_S, str_Code, Chr(&H2)) + 1
            str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_E + 1)
        Loop
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)

            If str_Line Like "O|*" Then
               str_Tmp = Split(str_Line, "|")(2)
               str_Tmp = Split(str_Tmp, "^")(0)
               str_Tmp = Val(str_Tmp)
               If strSampleNo <> str_Tmp Then
                    If str_Item <> "" And Val(strSampleNo) > 0 And strDate <> "" Then
                        If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                        strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |血清|0" & str_Item
                    End If
                    strSampleNo = str_Tmp
                    str_Item = ""
               End If
            ElseIf str_Line Like "R|*|*^DOSE|*" Then
               strDate = Split(str_Line, "|")(12)
               If strDate <> "" Then
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    If Not IsDate(strDate) Then strDate = ""
               End If
               strTestItem = Split(str_Line, "|")(2)
               strTestItem = Trim(Split(strTestItem, "^")(3))
               strTestValue = Trim(Split(str_Line, "|")(3))
               If strTestValue <> "" Then
                    str_Item = str_Item & "|" & Replace(strTestItem, "^", "") & "|" & strTestValue
               End If
            ElseIf str_Line Like "Q|*|ALL|*" Then
                '双向
                strSampleNo = Trim(Split(str_Line, "|")(3))
                If strSampleNo <> "" Then
                    If Left$(strSampleNo, 1) = "^" Then strSampleNo = Mid$(strSampleNo, 2)
                    If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then
                        strCmd = "1|"
                        strResult = "||" & Format(Now, "yyyy-MM-dd") & "|" & "^^" & strSampleNo
                    End If
                End If
                
            End If
        Next
        
        If strCmd <> "1|" Then
            If str_Item <> "" And Val(strSampleNo) > 0 And strDate <> "" Then
                If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |血清|0" & str_Item
            End If
        Else
            Exit Do
        End If
        '--------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = 0
        If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    SaveLog "Centaur-收到" & strSource & ",第" & CStr(Erl()) & "行," & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    Dim varInData As Variant, iLoop As Integer, strData As String, i As Integer
    Dim strNo As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String

    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '刚进入双向模式，还没有生成发送数据

106     SaveLog "Centaur-进入双向，" & strSample & ",缓存" & mstrWaitSend
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
            '--条码
114         strNo = Trim(Split(strData, "|")(5))
            '---急诊
116         If Val(Split(strData, "|")(4)) = 1 Then
118             strEmergency = "S"
            Else
120             strEmergency = "R"
            End If
            '----指标
         
122         strItems = ""
124         For i = 10 To UBound(Split(strData, "|"))
126             strChannel = Trim$("" & Split(strData, "|")(i))
128             strItems = strItems & "^^^" & strChannel & "\"
            Next
        
130         strSend = ""
132         If strItems <> "" And strNo <> "" Then
134             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
            
136             strSend = "H|\^&" & Chr(&HD) & "P|1" & Chr(&HD)
138             strSend = strSend & "O|1|" & strNo & "||" & strItems & "|" & strEmergency & "||||||||||||||||||||O\Q" & Chr(&HD)
140             strSend = strSend & "L|1" & Chr(&HD) & Chr(&H3)
            
142             iRow = 0: strSegment = ""
144             Do While strSend <> ""
146                 If Len(strSend) < 230 Then
148                     iRow = iRow + 1
150                     strSegment = Chr(&H5) & strSegment & Chr(&H2) & iRow & strSend & GetCheckCode(iRow & strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
152                     strSend = ""
                    Else
154                     iRow = iRow + 1
156                     strTmp = Mid(strSend, 1, 230) & Chr(&H17)
158                     strSegment = strSegment & Chr(&H2) & iRow & strTmp & GetCheckCode(iRow & strTmp) & Chr(&HD) & Chr(&HA)
160                     strSend = Mid(strSend, 231)
                    End If
                Loop
                
162             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If

164         SendSample = GetSendData
166         If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式

        Next
168     SaveLog "Centaur-进入双向，数据已生成,发送:" & SendSample & ",状态" & iSendStep & ",缓存:" & mstrWaitSend
    Else
        '收到仪器应答，从缓存中取得指令发送

170     If InStr(strResponse, Chr(&H6)) > 0 Then
172         strTmp = "收到06"
174         SendSample = GetSendData
176         If mstrWaitSend = "" Then iSendStep = 0
178     ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
            '收到仪器请求,本次通讯失败，丢弃未发送的数据，切换状态
180         strTmp = "收到05,优先处理，放弃本段数据，退出双向"
182         mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
183         mstrCode = mstrCode & Chr(&H5)
184         SendSample = Chr(&H6)
186         iSendStep = 0
188     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
190         strTmp = "收到04,退出双向,清空缓存"
192         SendSample = ""
194         mstrWaitSend = ""
196         iSendStep = 0
        Else
198         If mstrWaitSend = "" Then iSendStep = 0
200         strTmp = "收到其他:" & strResponse
        End If
        
        
202     If iSendStep = 0 Then
204         strTmp = strTmp & ",退出双向"
        Else
206         strTmp = strTmp & ",保持双向"
        End If
        
208     If SendSample <> "" Then
210         strTmp = strTmp & ",发送:" & SendSample
        Else
212         strTmp = strTmp & ",无数据发送"
        End If
214     If mstrWaitSend <> "" Then
216         strTmp = strTmp & ",缓存" & mstrWaitSend
        Else
218         strTmp = strTmp & ",无数据缓存"
        End If
220     SaveLog "Centaur-双向，" & strTmp
    End If
    Exit Function
hErr:
222     SaveLog "Centaur-第" & CStr(Erl()) & "行收到" & strResponse & ",出现错误:" & Err.Description & "，当前缓存" & mstrWaitSend
End Function

Private Function GetSendData() As String
    '从缓存中获取一段数据，返回。如果缓存中没有数据，返回空串

    
    If mstrWaitSend = "" Then Exit Function
    
    If Left$(mstrWaitSend, 1) = Chr(&H5) Then
        GetSendData = Chr(&H5)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    End If

End Function
Private Sub SaveLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    strOutput = Replace(strOutput, Chr(&HD), "<CR>")
    strOutput = Replace(strOutput, Chr(&HA), "<LF>")
    objStream.WriteLine (strDate & ":   " & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub
Private Function GetCheckCode(ByVal str_in As String) As String
    '计算校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

