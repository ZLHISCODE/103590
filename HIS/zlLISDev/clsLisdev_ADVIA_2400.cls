VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisdev_ADVIA_2400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ADVIA 2400 全自动化学发光仪
'厂商:拜尔 西门子
'作者:
'用户:丹东中心医院

Private mstrCode As String      '待解码数据
Private mstrWaitSend As String  '待发送数据
'Private mstrRackNo As String    '记录条码对应的盘号，杯号等信息
Private mstrEndBarcode As String '记录每次指令最后一个条码
Private mintRow As Integer       '顺序号

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
 
'    mstrRackNo = ""
    mstrWaitSend = ""
    mstrCode = ""
    mstrEndBarcode = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function
 
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
        '功能：解析数据
        '参数：
        'strSource：要解析的原始串
        'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
        '   每组检验结果以||分隔,元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
        '   第2个元素：检验人
        '   第3个元素：标本
        '   第4个元素：是否质控品
        '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
        '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
        'strReserved：最后不能完成解析的原始串，需要返回作后续处理
        'strCmd：如果需要，可返回向设备发送的命令
        Dim strCode As String, strQC As String, strSampleNo As String, strChannel As String, strValue As String
        Dim varCode As Variant, iRow As Integer, strLine As String, intStart As String, strItems As String
        Dim lngS As Long, lngE As Long, i As Integer, strDate As String
        Dim iCount As Integer, strSendBlock As String
        
        On Error GoTo hErr
100     strCmd = "": strResult = ""
102     mstrCode = mstrCode & strSource
104     If InStr(strSource, Chr(&H5)) > 0 Then
106         strCmd = Chr(&H6)
108     ElseIf InStr(strSource, Chr(&HA)) > 0 Then
110         strCmd = Chr(&H6)
        End If
    
112     lngE = 0
114     lngS = InStr(mstrCode, Chr(&H5))
116     If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
        
118     If Right(strSource, 1) = Chr(&H4) Then
            '收到一个完整的数据，这时如果还有未发送数据，进入双向模式
120         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
122             Call SaveLog(" 缓存中有未发送数据，切换到双向模式,缓存" & mstrWaitSend & ",序号" & mintRow)
                
124             If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
126                 mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
                Else
128                 mstrWaitSend = Chr(&H5) & "|" & mstrWaitSend
130                 mintRow = 0
                End If
132             strCmd = "1|"
134             strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
                Exit Sub
            End If
        End If
        
136     Do While lngS > 0 And lngE > 0
138         strCode = Mid(mstrCode, lngS, lngE - lngS + 1)
140         mstrCode = Mid$(mstrCode, lngE + 1)
            '-----------------------------------------------------
142         strQC = "0"
144         strSampleNo = ""
146         strDate = ""
148         strItems = ""
150         intStart = -1
152         iCount = 0
154         strValue = ""
156         varCode = Split(strCode, Chr(&H2))

158         For iRow = LBound(varCode) To UBound(varCode)
160             strLine = varCode(iRow)
162             If Mid$(strLine, 2, 1) = "R" Then
164                 If Not (Mid$(strLine, 44, 1) = "M" Or Mid$(strLine, 44, 1) = "D" Or Mid$(strLine, 44, 1) = "U") Then
166                     If strSampleNo <> "" And strDate <> "" And strItems <> "" Then
168                         If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
170                         strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|" & strQC & strItems
172                         strQC = "0"
174                         strSampleNo = ""
176                         strDate = ""
178                         strItems = ""
180                         intStart = -1
                        
                        End If
182                     If Mid$(strLine, 19, 1) = "N" Then
                            '--普通标本
184                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
186                         strSampleNo = Trim$(Mid$(strLine, 21, 13))
188                         intStart = 91
190                     ElseIf Mid$(strLine, 25, 1) = "N" Then
                            '--普通标本
192                         strSampleNo = Trim$(Mid$(strLine, 27, 13))
194                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
196                         intStart = 97
198                     ElseIf Mid$(strLine, 19, 1) = "I" Then
                            '---急诊
200                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
202                         strSampleNo = Trim$(Mid$(strLine, 21, 13)) & "^1^"
204                         intStart = 91
206                     ElseIf Mid$(strLine, 25, 1) = "I" Then
                            '---急诊
208                         strSampleNo = Trim$(Mid$(strLine, 27, 13)) & "^1^"
210                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
212                         intStart = 97
214                     ElseIf Mid$(strLine, 19, 1) = "C" Then
                            '---质控
216                         strQC = "1"
218                         strSampleNo = Trim$(Mid$(strLine, 21, 13))
220                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2)
222                         intStart = 91
224                     ElseIf Mid$(strLine, 25, 1) = "C" Then
                            '---质控
226                         strQC = "1"
228                         strSampleNo = Trim$(Mid$(strLine, 27, 13)) & "^1^"
230                         strDate = Mid$(strLine, 11, 4) & "-" & Mid$(strLine, 15, 2) & "-" & Mid$(strLine, 17, 2) & " " & Mid$(strLine, 19, 2) & ":" & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 23, 2)
232                         intStart = 97
                        End If
                        '3+1+8+3=15
234                     For i = intStart To Len(strLine) Step 15
236                         If i + 15 > Len(strLine) Then Exit For
238                         strChannel = Trim$(Mid$(strLine, i, 3))
240                         strValue = Trim$(Mid$(strLine, i + 4, 8))
                        
242                         If strChannel <> "" And IsNumeric(strValue) Then strItems = strItems & "|" & strChannel & "|" & strValue
                        Next
                    Else
244                     If intStart = 91 Then
246                         intStart = 41
248                     ElseIf intStart = 97 Then
250                         intStart = 47
                        End If
                    
252                     For i = intStart To Len(strLine) Step 15
254                         If i + 15 > Len(strLine) Then Exit For
256                         strChannel = Trim$(Mid$(strLine, i, 3))
258                         strValue = Trim$(Mid$(strLine, i + 4, 8))
                        
260                         If strChannel <> "" And IsNumeric(strValue) Then strItems = strItems & "|" & strChannel & "|" & strValue
                        Next
                    
                    End If
262             ElseIf Mid$(strLine, 2, 1) = "Q" And Mid$(strLine, 10, 1) = "0" Then
                    '双向,
                    '------
264                 SaveLog "2400-双向查询请求:" & strLine
266                 strValue = ""
268                 If Mid$(strLine, 1, 1) = "1" Then iCount = 0
270                 For i = 11 To Len(strLine) Step 13
272                     If i + 13 > Len(strLine) Then Exit For
274                     strValue = Trim(Mid$(strLine, i, 13))
276                     If Len(strValue) >= 12 Then
278                         strCmd = "1|"
280                         iCount = iCount + 1
282                         If iCount = 8 Then iCount = 0

    '274                         mstrRackNo = mstrRackNo & "," & Mid$(strLine, 4, 6) & "|" & iCount
284                         strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & strValue
286                         If Mid$(strLine, 1, 1) = "1" And i = 11 Then
288                             SaveLog "------------2400双向开始，添加05,开始条码," & strValue
290                             mstrWaitSend = mstrWaitSend & "|" & Chr(&H5)
292                             mintRow = 0
                            End If
                        Else
294                         SaveLog "2400因" & strValue & " 小于12位，不处理！"
                            
                        End If
                    Next
                    
    '286                 If mstrRackNo <> "" Then mstrRackNo = Mid$(mstrRackNo, 2)
296             ElseIf Mid$(strLine, 2, 1) = "Q" And Mid$(strLine, 10, 1) <> "0" Then
                    
298                 SaveLog "2400-双向查询请求(非条码),应答:04" & strLine
300                 'strCmd = Chr(&H4)   '发出终止通讯请求
                    '--
                End If
            Next
            
302         If strValue <> "" And strCmd = "1|" Then
304             If Len(strValue) >= 12 Then
306                 If InStr(mstrEndBarcode, "|" & strValue) <= 0 Then mstrEndBarcode = mstrEndBarcode & "|" & strValue
308                 SaveLog "2400双向缓存的未位条码为" & mstrEndBarcode '& ",缓存的架号为" & mstrRackNo
                End If
            End If
            
310         If strSampleNo <> "" And strDate <> "" And strItems <> "" Then
312             If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
314             strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|" & strQC & strItems
            End If
            '-----------------------------------------------------
316         lngE = 0
318         lngS = InStr(mstrCode, Chr(&H5))
320         If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
        Loop
    
322     If strResult <> "" Then strResult = Mid$(strResult, 3)
        
        Exit Sub
hErr:
    SaveLog "2400-第" & CStr(Erl()) & "行,出现错误提示," & Err.Description & ",收到的数据为" & strSource

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

    
    Dim strTmp As String
    On Error GoTo hErr
    
100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '刚进入双向模式，还没有生成发送数据
106     If strSample = "" Then
108         SendSample = GetSendData
110         If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式
112         SaveLog "2400-从单向直接切换，" & IIf(iSendStep = 0, "退出双向", "保持双向") & "状态,本次发送" & SendSample & ",未发送数据:" & mstrWaitSend
        Else
114         If GetSendCmd(strSample) Then
116             SendSample = GetSendData
118             If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式
            End If
120         SaveLog "2400-生成数据结束，" & IIf(iSendStep = 0, "退出双向", "保持双向") & "状态,本次发送" & SendSample & ",未发送数据:" & mstrWaitSend
        End If

    Else
        '收到仪器应答，从缓存中取得指令发送
122     If InStr(strResponse, Chr(&H6)) > 0 Then

124         SendSample = GetSendData
126         If mstrWaitSend = "" Then iSendStep = 0
128         strTmp = "06 "
130     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then

132         SendSample = GetSendData
134         If mstrWaitSend = "" Then iSendStep = 0
136         strTmp = "15 Err "
138     ElseIf InStr(strResponse, Chr(&H11)) > 0 Then

140         SendSample = GetSendData
142         If mstrWaitSend = "" Then iSendStep = 0
144         strTmp = "11 Next "
146     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
            '超时-放弃未发送的标本，
148         mstrWaitSend = ""
150         mstrEndBarcode = ""
152         iSendStep = 0
            
154         strTmp = "04 中止"
156     ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
            '收到05，暂停后面的双向传输，进入单向模式
            '清空未发送的数据
    '152         mstrWaitSend = ""
    '156         mstrEndBarcode = ""
158         If mstrWaitSend <> "" Then
160             mstrWaitSend = Chr(&H5) & "|" & mstrWaitSend
162             mintRow = 0
            End If
164         iSendStep = 0
166         strTmp = "05 请求"
168         mstrCode = mstrCode & Chr(&H5)
170         SendSample = Chr(&H6)
        Else
172         strTmp = strResponse
        End If
174     strTmp = "2400-继续双向，收到" & strTmp & "," & IIf(iSendStep = 0, "-----------退出双向", "保持双向") & "状态"
176     If SendSample <> "" Then
178         strTmp = strTmp & ",本次发送:" & SendSample
        Else
180         strTmp = strTmp & ",没有数据发送"
        End If
182     If mstrWaitSend <> "" Then
184         strTmp = strTmp & ",待发送数据:" & mstrWaitSend
        Else
186         strTmp = strTmp & ",数据已发完"
        End If
188     SaveLog strTmp
    End If
    Exit Function
hErr:
190     SaveLog "2400-继续双向第" & CStr(Erl()) & "行,出现错误提示" & Err.Description & ",收到:" & strResponse
End Function

Private Function GetSendCmd(ByVal strSample As String) As Boolean
    '生成待发送数据，如果是最后一个，则返回True
    Dim varInData As Variant, iLoop As Integer, strData As String
    Dim strNo As String, strBlood As String, strEmergency As String, strRackNo As String
    Dim strChannel As String, strSend As String, iCount As Integer, iFlag As Integer
    Dim iRow As Integer, i As Integer
    
106     SaveLog "2400-准备生成待发送数据," & strSample
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
            '--条码
114         strNo = Trim(Split(strData, "|")(5))
            '--标本类型
116         strBlood = Split(strData, "|")(3)
118         If InStr(strBlood, "尿") > 0 Then
120             strBlood = "2"
            Else
122             strBlood = "1"
            End If
            '---急诊
124         If Val(Split(strData, "|")(4)) = 1 Then
126             strEmergency = "I"
            Else
128             strEmergency = "N"
            End If
            '---盘号
            strRackNo = "0101" & Format(UBound(Split(strData, "|")) - 9, "000")
            '----指标
154         iCount = 0: strSend = "": iFlag = 0

156         For i = 10 To UBound(Split(strData, "|"))
158             If iCount = 0 Then
                        '首次数据-头
160                 strSend = "O " & strRackNo & strEmergency & "0" & strNo & String$(13 - Len(strNo), " ") & Space$(7) & Space$(16) & Space$(16)
162                 strSend = strSend & "M000" & Format(CDate(Split(strData, "|")(0)), "yyyyMMdd") & " 1.0" & strBlood & "1"
                End If
                    '--- 通道码为3位
164             strChannel = GetChannel(Trim$(Split(strData, "|")(i)))
166             If strChannel <> "" Then strSend = strSend & strChannel & "M"

168             iCount = iCount + 1
            Next
            
170         If strSend <> "" Then mstrWaitSend = mstrWaitSend & "|" & strSend

180         If mstrEndBarcode = "" Then SaveLog "2400-最后条码缓存为空,当前处理的条码是" & strNo
            
            '---本次　是指令中的最后一个条码，则清除
192         If InStr(mstrEndBarcode & "|", "|" & strNo & "|") > 0 Then
                If strSend <> "" Then mstrWaitSend = mstrWaitSend & "|" & Chr(&H4)
194             SaveLog "2400-最后一个条码，添加04,从:" & mstrEndBarcode & "中请除" & strNo
196             mstrEndBarcode = Replace(mstrEndBarcode, "|" & strNo, "")
                GetSendCmd = True
            End If
        Next
        
End Function

Private Function GetChannel(ByVal strIn As String) As String
    '生成通道码
    
    If Len(Trim(strIn)) <= "3" Then
        GetChannel = String(3 - Len(Trim(strIn)), "0") & Trim(strIn)
        If GetChannel = Space(3) Then
            SaveLog "2400-双向，通道码为空," & strIn
            GetChannel = ""
        End If
    Else
        SaveLog "2400-双向，通道码超长," & strIn
    End If

End Function
Private Function GetSendData() As String
    '从缓存中获取一段数据，返回。如果缓存中没有数据，返回空串
    
    Dim strSend As String
    
    If Left$(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid$(mstrWaitSend, 2)
    If mstrWaitSend = "" Then Exit Function
    If InStr(mstrWaitSend, "|") <= 0 Then
        strSend = mstrWaitSend
        mstrWaitSend = ""
    Else
        strSend = Mid$(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
        mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
    End If
    
    If Len(strSend) > 1 Then
        mintRow = mintRow + 1
        If mintRow = 8 Then mintRow = 0
        strSend = mintRow & strSend & " " & Chr(&H3)   '最后一段数据固定为 03
        strSend = Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
    
    End If
    GetSendData = strSend
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
    strFileName = App.Path & "\LisDev_" & Format(Date, "yyyyMMdd") & ".LOG"
    
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
