VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ASTM_i2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: Axsym i2000
'厂商: 雅培
'作者: 陈东
'用户: 本钢总医院
'创建: 2012-10-25
'修改记录:
'
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
    strData = strSource
    strResult = ""
    
    If strData <> "" Then
        '---   应答部分 ----
        Call SaveLog("收到字符: " & strData)
        mstrData = mstrData & strData
        If InStr(strData, Chr(&H5)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("接收到仪器发送过来的通讯请求符ENQ。")
        ElseIf InStr(strData, Chr(&H3)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("收到ETX，应答ACK后，等待通讯结束符EOT。")
        ElseIf InStr(strData, Chr(&H17)) > 0 Then
            strCmd = Chr(&H6)
            Call SaveLog("收到ETB，应答ACK后，继续等待下一帧。")
        End If
    
    
        lngS = InStr(mstrData, Chr(&H5))
        lngE = 0
        If lngS > 0 Then lngE = InStr(lngS, mstrData, Chr(&H4))
        
        Do While lngS > 0 And lngE > 0
        
            Call SaveLog("接收到一个完整数据")
            strData = Mid(mstrData, lngS, lngE - lngS + 1)
            mstrData = Mid(mstrData, lngE + 1)
            
            Do While strData Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
                lngS = InStr(strData, Chr(&H17))
                lngE = InStr(lngS, strData, Chr(&H2)) + 1
                strData = Mid(strData, 1, lngS - 1) & Mid(strData, lngE + 1)
            Loop
            
            '------------------------------------------------------------------

            Call SaveLog("开始解码")
            varRecord = Split(strData, Chr(&HD))
            strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
            
            For lngSeqi = LBound(varRecord) To UBound(varRecord)
                strLine = varRecord(lngSeqi)
    
                If strLine Like "*" & Chr(&H2) & "?Q|*" Then
                    varField = Split(strLine, "|")
                    strBarcode = Trim("" & varField(2))
                    Call SaveLog("发现条码“" & strBarcode & "”的请求指令，开始获取检验项目信息。")
                    If Left(strBarcode, 1) = "^" Then strBarcode = Mid(strBarcode, 2)
                    
                    strCmd = "1|"
                    strResult = Format(Now, "yyyy-MM-dd") & "|^^" & strBarcode
                    
                    Exit Sub
                ElseIf strLine Like "*" & Chr(&H2) & "?O|*" Then
                    varField = Split(strLine, "|")
                    strSampleNo = Trim("" & varField(2))
                    If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
                    strQC = Trim("" & varField(11))
                    If strQC = "Q" Then
                        strQC = "1"
                    Else
                        strQC = "0"
                    End If
                ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then
                    varField = Split(strLine, "|")
                    strChannel = Trim("" & varField(2))
                    strValue = Trim("" & varField(3))
                    If strDate = "" Then
                        strDate = Trim("" & varField(12))
                        strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    End If
                    If Right(strChannel, 2) = "^F" Then
                        strChannel = Split(strChannel, "^")(3)
                        If strChannel <> "" And strValue <> "" Then
                            strItem = strItem & "|" & strChannel & "|" & strValue
                        End If
                    End If
                End If
            
            Next
            
            If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
                Call SaveLog("发现“" & strDate & "”标本号为“" & strSampleNo & "”的检验结果“" & strItem & "”")
            End If
            '------------------------------------------------------------------
            '检查是否还有完整的数据
            lngS = InStr(mstrData, Chr(&H5))
            lngE = 0
            If lngS > 0 Then lngE = InStr(mstrData, Chr(&H4))
        Loop
    End If
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    
    Exit Sub
hErr:
    Call SaveLog(CStr(Erl()) & "行," & Err.Description)
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
    Dim strSampeNo As String
    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '刚进入双向模式，还没有生成发送数据
106     varInData = Split(strSample, "||")
108     For iLoop = LBound(varInData) To UBound(varInData)
110         strData = varInData(iLoop)
112         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMddHHmmss")
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
128         For i = 10 To UBound(Split(strData, "|"))
130             strChannel = Trim$("" & Split(strData, "|")(i))
132             strItems = strItems & "^^^" & strChannel & "\"
            Next
134         strSend = ""
136         If strItems <> "" And strBarcode <> "" Then
138             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
140             iRow = 0
142             strSegment = Chr(&H5)
144             iRow = iRow + 1
146             strSend = iRow & "H|\^&||||||||||P|1" & Chr(&HD) & Chr(&H3)
148             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                'P|1||A43749||Smith^Tom^J||19950315^120^H|性别M/F/U|||||医生||||||||||||位置<CR>
                'P|1||标本号|<CR>
150             iRow = iRow + 1
152             strSend = iRow & "P|1||" & strSampeNo & "||" & strSampeNo & "||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
154             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
156             iRow = iRow + 1
158             strSend = iRow & "O|1|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & strDate & _
                          "||||A||||||||||||||Q" & Chr(&HD) & Chr(&H3)
160             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
162             iRow = iRow + 1
164             strSend = iRow & "L|1" & Chr(&HD) & Chr(&H3)
166             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
168             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If
170         SendSample = GetSendData
172         If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式

        Next
174     If SendSample <> "" Then
176         SaveLog "双向，生成待发送数据,发送:" & SendSample & "缓存:" & mstrWaitSend
        Else
178         SaveLog "双向，生成待发送数据,缓存:" & mstrWaitSend
        End If
        
    Else
        '收到仪器应答，从缓存中取得指令发送

180     If InStr(strResponse, Chr(&H6)) > 0 Then
182         SendSample = GetSendData
184         If mstrWaitSend = "" Then iSendStep = 0
186         SaveLog "仪器应答06，通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
188     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
190         SendSample = GetSendData
192         If mstrWaitSend = "" Then iSendStep = 0
194         SaveLog "仪器应答04，表示希望暂停，根据协议仍当做06,通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
196     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
198         SendSample = GetSendData
200         If mstrWaitSend = "" Then iSendStep = 0
202         SaveLog "仪器应答15，通讯失败。发送:" & SendSample & "缓存:" & mstrWaitSend
204     ElseIf strResponse = Chr(&H5) Then
206         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
208             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
210             mstrWaitSend = ""
            End If
212         Call SaveLog("仪器应答05，因仪器有优先权，本次发送取消")
214         iSendStep = 0
216         SendSample = Chr(&H6)
        Else
218         Call SaveLog("仪器应答" & strResponse)
        End If
220     If mstrWaitSend = "" Then iSendStep = 0    '已发送完，退出双向模式
    End If
    Exit Function
hErr:
222     SaveLog "第" & CStr(Erl()) & "行错误" & Err.Description

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
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject

    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(Date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub

