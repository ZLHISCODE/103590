VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CI16000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: Axsym Ci16000
'厂商: 雅培
'作者: 陈东
'用户: 连山区人民医院
'创建: 2014-11-22
'修改记录:
'2014-11-22 连山区人民医院 CI16000 ，新版LIS双向时，不能用传入的日期直接传到仪器，因时间不是23:59:59秒。肯定超过仪器当前时间。仪器报警。


Private mstrData As String
Private mstrWaitSend As String

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

        Dim lngPos As Long
        Dim lngSeqi As Long
        Dim lngCount As Long
        Dim lngCount1 As Long
        Dim lngSeqj As Long
        Dim lngSeqjcnt As Long
        Dim lngfn As Long

    
        Dim varReply As Variant
        Dim strReplyMsg As String
        Dim strInstrReply As String
        Dim strChecksum As String
  
    
        Dim varRecord As Variant
        Dim varField As Variant
        Dim lngS As Long, lngE As Long  '开始，结束位置
        Dim strData As String
        Dim strLine As String
        Dim strBarcode As String    '条码
        Dim strSampleNo As String   '标本号
        Dim strQC As String         '是否质控
        Dim strDate As String       '日期
        Dim strValue As String, strChannel As String '结果，通道码
        Dim strItem As String
    
        On Error GoTo hErr
100     strData = strSource
102     strResult = ""
    
104     If strData <> "" Then
            '---   应答部分 ----
106         Call SaveLog("收到字符: " & strData)
108         mstrData = mstrData & strData
110         If InStr(strData, Chr(&H5)) > 0 Then
112             strCmd = Chr(&H6)
114             Call SaveLog("接收到仪器发送过来的通讯请求符ENQ。")
116         ElseIf InStr(strData, Chr(&H3)) > 0 Then
118             strCmd = Chr(&H6)
120             Call SaveLog("收到ETX，应答ACK后，等待通讯结束符EOT。")
122         ElseIf InStr(strData, Chr(&H17)) > 0 Then
124             strCmd = Chr(&H6)
126             Call SaveLog("收到ETB，应答ACK后，继续等待下一帧。")
            End If
    
    
128         lngS = InStr(mstrData, Chr(&H5))
130         lngE = 0
132         If lngS > 0 Then lngE = InStr(lngS, mstrData, Chr(&H4))
        
134         Do While lngS > 0 And lngE > 0
        
136             Call SaveLog("接收到一个完整数据")
138             strData = Mid(mstrData, lngS, lngE - lngS + 1)
140             mstrData = Mid(mstrData, lngE + 1)
            
142             Do While strData Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
144                 lngS = InStr(strData, Chr(&H17))
146                 lngE = InStr(lngS, strData, Chr(&H2)) + 1
148                 strData = Mid(strData, 1, lngS - 1) & Mid(strData, lngE + 1)
                Loop
            
                '------------------------------------------------------------------

150             Call SaveLog("开始解码")
152             varRecord = Split(strData, Chr(&HD))
154             strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
            
156             For lngSeqi = LBound(varRecord) To UBound(varRecord)
158                 strLine = varRecord(lngSeqi)
    
160                 If strLine Like "*" & Chr(&H2) & "?Q|*" Then
162                     varField = Split(strLine, "|")
164                     strBarcode = Trim("" & varField(2))
166                     Call SaveLog("发现条码“" & strBarcode & "”的请求指令，开始获取检验项目信息。")
168                     If Left(strBarcode, 1) = "^" Then strBarcode = Mid(strBarcode, 2)
                    
170                     strCmd = "1|"
172                     strResult = Format(Now, "yyyy-MM-dd") & "|^^" & strBarcode
                    
                        Exit Sub
174                 ElseIf strLine Like "*" & Chr(&H2) & "?O|*" Then

176                     If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
178                         strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
180                         Call SaveLog("发现“" & strDate & "”标本号为“" & strSampleNo & "”的检验结果“" & strItem & "”")
                           
                        End If
182                     strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
                        
184                     varField = Split(strLine, "|")
186                     strSampleNo = Trim("" & varField(2))
                        
188                     If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
190                     strQC = Trim("" & varField(11))
192                     If strQC = "Q" Then
194                         strQC = "1"
                        Else
196                         strQC = "0"
                        End If
                        
198                 ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then
200                     varField = Split(strLine, "|")
202                     strChannel = Trim("" & varField(2))
204                     strValue = Trim("" & varField(3))
206                     If strDate = "" Then
208                         strDate = Trim("" & varField(12))
210                         strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                        End If
212                     If Right(strChannel, 2) = "^F" Then
214                         strChannel = Split(strChannel, "^")(3)
216                         If strChannel <> "" And strValue <> "" Then
218                             strItem = strItem & "|" & strChannel & "|" & strValue
                            End If
                        End If
                        
                    End If
            
                Next
            
220             If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
222                 strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
224                 Call SaveLog("发现“" & strDate & "”标本号为“" & strSampleNo & "”的检验结果“" & strItem & "”")
226                 strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
                End If
                '------------------------------------------------------------------
                '检查是否还有完整的数据
228             lngS = InStr(mstrData, Chr(&H5))
230             lngE = 0
232             If lngS > 0 Then lngE = InStr(mstrData, Chr(&H4))
            Loop
        End If
234     If strResult <> "" Then strResult = Mid$(strResult, 3)
    
        Exit Sub
hErr:
236     Call SaveLog(CStr(Erl()) & "行," & Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    ''   strSample:解析需要向仪器发送的标本信息
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
    Dim strBarcode As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String, strDate As String, strSampleType As String
    Dim strSampeNo As String, intOdx As Integer
    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '刚进入双向模式，还没有生成发送数据
106     varInData = Split(strSample, "||")
108     For iLoop = LBound(varInData) To UBound(varInData)
110         strData = varInData(iLoop)
112         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMdd000000")
114         strSampeNo = Trim(Split(strData, "|")(1))
            '--标本
116         strSampleType = "" 'Trim(Split(strData, "|")(3))
            '--条码
118         strBarcode = Trim(Split(strData, "|")(5))
            '---急诊
120         If Val(Split(strData, "|")(4)) = 1 Then
122             strEmergency = "S"
            Else
124             strEmergency = ""
            End If
            '----指标
126         strItems = ""

128         strSend = ""
130         If UBound(Split(strData, "|")) >= 10 And strBarcode <> "" Then
132             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
134             iRow = 0
136             strSegment = Chr(&H5)
138             iRow = iRow + 1
140             strSend = iRow & "H|\^&||||||||||P|1" & Chr(&HD) & Chr(&H3)
142             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
            
                'P|1||A43749||Smith^Tom^J||19950315^120^H|性别M/F/U|||||医生||||||||||||位置<CR>
                'P|1||标本号|<CR>
144             iRow = iRow + 1
146             strSend = iRow & "P|1||" & strSampeNo & "||" & strSampeNo & "||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
148             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
            
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
                'O指定 项目过多会造成通讯故障，在此处理
                intOdx = 0
150             For i = 10 To UBound(Split(strData, "|"))
152                 strChannel = Trim$("" & Split(strData, "|")(i))
154                 strItems = strItems & "\^^^" & strChannel
                    
156                 If (i + 1) Mod 10 = 0 Or i = UBound(Split(strData, "|")) Then
158                     iRow = iRow + 1
                        intOdx = intOdx + 1
                        strItems = Mid(strItems, 2)
160                     strSend = iRow & "O|" & intOdx & "|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & strDate & _
                                "||||A||||||||||||||Q" & Chr(&HD) & Chr(&H3)
162                     strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                        strItems = ""
                    End If
                Next
                          
                'L
164             iRow = iRow + 1
166             strSend = iRow & "L|1" & Chr(&HD) & Chr(&H3)
168             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
170             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment


            End If
172         SendSample = GetSendData
174         If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式

        Next
176     If SendSample <> "" Then
178         SaveLog "双向，生成待发送数据,发送:" & SendSample & "缓存:" & mstrWaitSend
        Else
180         SaveLog "双向，生成待发送数据,缓存:" & mstrWaitSend
        End If
        
    Else
        '收到仪器应答，从缓存中取得指令发送

182     If InStr(strResponse, Chr(&H6)) > 0 Then
184         SendSample = GetSendData
186         If mstrWaitSend = "" Then iSendStep = 0
188         SaveLog "仪器应答06，通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
190     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
192         SendSample = GetSendData
194         If mstrWaitSend = "" Then iSendStep = 0
196         SaveLog "仪器应答04，表示希望暂停，根据协议仍当做06,通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
198     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
200         SendSample = GetSendData
202         If mstrWaitSend = "" Then iSendStep = 0
204         SaveLog "仪器应答15，通讯失败。发送:" & SendSample & "缓存:" & mstrWaitSend
206     ElseIf strResponse = Chr(&H5) Then
208         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
210             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
212             mstrWaitSend = ""
            End If
214         Call SaveLog("仪器应答05，因仪器有优先权，本次发送取消")
216         iSendStep = 0
218         SendSample = Chr(&H6)
        Else
220         Call SaveLog("仪器应答" & strResponse)
        End If
222     If mstrWaitSend = "" Then iSendStep = 0    '已发送完，退出双向模式
    End If
    Exit Function
hErr:
224     SaveLog "第" & CStr(Erl()) & "行错误" & Err.Description

End Function

'------- 以下是内部函数
Private Function GetCheckCode(ByVal str_in As String) As String
    '计算校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

Private Function GetSendData() As String
    '从缓存中获取一段数据，返回。如果缓存中没有数据，返回空串
    If mstrWaitSend = "" Then Exit Function
    If Left$(mstrWaitSend, 1) = Chr(&H5) Then
        GetSendData = Chr(&H5)
        mstrWaitSend = Mid(mstrWaitSend, 2)
        Call SaveLog("发送通讯结束符05，本次通讯开始。")
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
        Call SaveLog("发送通讯结束符04，本次通讯结束。")
    End If
End Function

Private Sub SaveLog(ByVal StrInput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------

    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject

    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub





