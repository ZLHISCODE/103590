VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_C601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:临床生化分析系 C601
'厂商:Roche 罗氏
'作者:陈东
'用户:
'2014-10-11 妇儿医院修改双向

'注意： 稀释与非稀释项目不能开在一个标本中，比如雌二醇和雌二醇稀释5倍不能开在一个标本中。
'       HCG结果中有HH，LL两种类型。HH为》，LL为《

Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤

Private mlngStep_Out As Long  '传双向消息步骤
Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.

Private mstrWaitSend As String '待发送标本
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
        Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String, str结果标志 As String
        Dim str_WaitSendLine As String, str_急诊 As String, str_质控 As String
        '-----------------------------------------------------------------------------
        Dim varQc As Variant, strIniFile As String, strQC As String, intloop As Integer
        Dim objFileSystem As New Scripting.FileSystemObject
        Dim objStream As Scripting.textStream, strLine As String
    
        '----------------------------
        '计算项目
        Dim strTP As String
        Dim strALB As String
        Dim strGLOB As String
        Dim strAST As String
        Dim strALT As String
        Dim strAPOA1 As String
        Dim strAPOB As String
        Dim strALBGLOB As String
        Dim strASTALT As String
        Dim strAPOA1APOB As String
        Dim strTBIL As String
        Dim strDBIL As String
        Dim strTDBIL As String
        Dim strHDL As String
        Dim strLDL As String
        Dim strHL As String
        Dim strTmp As String
    
100     str_in = strSource
        
102     If InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '仪器通讯结束
        
104         mlngStep_In = 0
        End If
    
106     If InStr(str_in, Chr(&H3)) > 0 Then                 '收到ETX(03)
108         lng_S = InStr(str_in, Chr(&H3))
110         mlngStep_Out = mlngStep_Out + 1
112         strCmd = Chr(&H6)
        End If
    
'114     If InStr(str_in, Chr(&H6)) > 0 Then                 '收到ACK(06)
'116         lng_S = InStr(str_in, Chr(&H6))
'118         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
'120         mlngStep_Out = mlngStep_Out + 1
'
'122         If mstrWaitSend <> "" Then
'124             If Left(mstrWaitSend, 1) = "," Then
'126                 mstrWaitSend = Mid(mstrWaitSend, 2)
'                End If
'128             If InStr(mstrWaitSend, ",") > 0 Then
'130                 str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
'132                 mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
'                Else
'134                 str_WaitSendLine = mstrWaitSend
'136                 mstrWaitSend = ""
'                End If
'
'138             If InStr(str_WaitSendLine, "|TSDWN^REPLY") > 0 Then
'140                 strCmd = Chr(&H5) & str_WaitSendLine
'                End If
'            End If
'        End If
    
142     If InStr(str_in, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
144         strCmd = Chr(&H6)
        End If
    
146     If InStr(str_in, Chr(&H17)) > 0 Then                  '收到ETB(17) 应答
148         strCmd = Chr(&H6)
        End If
    
    
150     If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
        '去掉
152     Do While mstrCode Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
154         lng_S = InStr(mstrCode, Chr(&H17))
156         lng_E = InStr(lng_S, mstrCode, Chr(&H2)) + 1
158         mstrCode = Mid(mstrCode, 1, lng_S - 1) & Mid(mstrCode, lng_E + 1)
        
        Loop
    
160     lng_S = InStr(mstrCode, Chr(&H5))
162     lng_E = InStr(mstrCode, Chr(&H4))
164     If lng_S > 0 And lng_E > 0 Then
166         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
168         mstrCode = Mid(mstrCode, lng_E + 1)
        
170         If InStr(str_Line, "|TSREQ^REAL|") > 0 Then         '双向通讯- 开始
172             mstrRack = ""
174             mlngStep_Out = 1
176             var_line = Split(str_Line, Chr(&HD))
            
178             For i_Count = LBound(var_line) To UBound(var_line)
180                 str_Line = var_line(i_Count)
182                 If str_Line Like "Q|*" Then
184                     str_条码 = Trim(Split(str_Line, "|")(2))
                    
186                     mstrRack = Split(str_条码, "^")(3) & "^" & Split(str_条码, "^")(4) & "^" & Split(str_条码, "^")(5) & "^^" & Split(str_条码, "^")(7) & "^" & Split(str_条码, "^")(8)
188                     str_条码 = Split(str_条码, "^")(2)
                    
    '1H|\^&|||cobas6000^1|||||host|TSREQ^REAL|P|1
    'Q|1|^^          260000055220^0^00031^4^^S1^SC||ALL||||||||O
    'L|1|N
    '35
    '1H|\^&|||cobas6000^1|||||host|TSREQ^REAL|P|1
    'Q|1|^^                   601^0^00034^1^^S1^SC||ALL||||||||O
    'L|1|N
    '96
    '


190                     If Val(str_条码) <> 0 Then
192                         If Len(CStr(Val(str_条码))) >= 12 Then
                                '条码
194                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_条码)
                            Else
                                '在条码模式下，临时用手工输入的标本号
196                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Trim(str_条码)
                            End If
                        Else
                            '标本号双工模式
198                         str_条码 = Trim(Split(str_Line, "|")(2))
200                         str_条码 = Trim(Split(str_条码, "^")(3))
202                         strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Trim(str_条码)
                        End If
                    

                    End If
                Next
            
204             If strResult <> "" Then
206                 strResult = Mid(strResult, 3)
208                 strCmd = "1|"       '向仪器发通讯开始请求 05
                End If

210         ElseIf InStr(str_Line, "|RSUPL^") > 0 Then     '检验结果
            '-----------------------------------------------
    '            strIniFile = App.Path & "\Roche_P800.ini"
    '            strQC = ""
    '            If objFileSystem.FileExists(strIniFile) Then
    '                Set objStream = objFileSystem.OpenTextFile(strIniFile)
    '                Do While Not objStream.AtEndOfStream
    '                    strLine = objStream.ReadLine
    '                    If InStr(strLine, "=") > 0 Then
    '                        strQC = strQC & "|" & strLine
    '                    End If
    '                Loop
    '            Else
    '               Call objFileSystem.CreateTextFile(strIniFile)
    '               Set objStream = objFileSystem.OpenTextFile(strIniFile, ForAppending)
    '               objStream.WriteLine "[水平1]"
    '               objStream.WriteLine "5200 = PPU"
    '               objStream.WriteLine "[水平2]"
    '               objStream.WriteLine "6100 = PNU"
    '
    '               strQC = "5200 = PPU|6100 = PNU"
    '            End If
    '            If strQC <> "" Then varQC = Split(Mid(strQC, 2), "|")
                '-----------------------------------------------------------------------------
212             var_line = Split(str_Line, Chr(&HD))
            
214             For i_Count = LBound(var_line) To UBound(var_line)
216                 str_Line = var_line(i_Count)
218                 If str_Line Like "O|*" Then
220                     str_日期 = Trim(Split(str_Line, "|")(7))
222                     If Len(str_日期) <> 14 Then str_日期 = Split(str_Line, "|")(22)
224                     str_日期 = Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
226                     str_质控 = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")
                    
228                     str_标本号 = Trim("" & Split(str_Line, "|")(2))
230                     If str_标本号 = "" Then
232                         str_标本号 = Trim("" & Split(str_Line, "|")(3))
                            '----修改
234                         str_标本号 = Split(str_标本号, "^")(0)
                        End If
236                     str_急诊 = Trim(Split(str_Line, "|")(5))
                    
238                     If Trim(str_标本号) <> "" Then
240                         If Len(Trim(str_标本号)) < 12 Then
                                '标本号
242                             str_标本号 = Trim(str_标本号) & IIf(str_急诊 = "S", "^1^", "")
                            Else
                                '条码
244                             str_标本号 = "^" & IIf(str_急诊 = "S", "^1", "^") & Trim(str_标本号)
                                '----修改
                                'str_标本号 = Trim(str_标本号)
                                'str_标本号 = Trim(str_标本号) & IIf(str_急诊 = "S", "^1^", "")
                            End If

                        End If
                    
246                     str_项目 = ""
248                 ElseIf str_Line Like "R|*" Then
250                     str_结果 = Trim(Split(str_Line, "|")(3))
252                     str结果标志 = Trim(Split(str_Line, "|")(6))
254                     If str结果标志 = "HH" Then
256                         str结果标志 = "＞"
258                     ElseIf str结果标志 = "LL" Then
260                         str结果标志 = "＜"
                        Else
262                         str结果标志 = ""
                        End If
                    
264                     If str_结果 <> "" Then
266                         If InStr(str_结果, "^") > 0 Then str_结果 = Split(str_结果, "^")(1) '--- 090407 有免疫的结果，为 -1^0.715 的格式。
                        
268                         strTmp = Replace(Split(str_Line, "|")(2), "^^^", "")
270                         If InStr(strTmp, "/") > 0 Then
272                             If Trim(Split(strTmp, "/")(1)) <> "" Then
                                    '处理 E2，带稀释倍数的项目
                                    '这种项目现在需要在检检中单独列出指标。通道码特殊处理为 12^5 这种样式
274                                 strTmp = Split(strTmp, "/")(0) & "^" & Trim(Split(strTmp, "/")(1))
                                Else
276                                 strTmp = Split(strTmp, "/")(0)
                                End If
                            
                            
                            End If
278                         str_项目 = str_项目 & "|" & strTmp & "|" & str结果标志 & str_结果
                        End If
                    End If
                Next
280             If IsDate(CDate(str_日期)) And str_标本号 <> "" And str_项目 <> "" Then
282                 strResult = str_日期 & "|" & str_标本号 & "| |血清|" & str_质控 & str_项目
'284                 WriteLog "收到仪器传回的结果 " & strResult, "C601"
                    '加上 计算项目
286                 var_line = Split(Mid(str_项目, 2), "|")
                
288                 For intloop = LBound(var_line) To UBound(var_line)
290                     If Trim(var_line(intloop)) = "320" Then 'Cysc
    '                        If Val(var_line(intLoop + 1)) <> 0 Then
    '                            strResult = strResult & "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4)
    '                            WriteLog "计算结果 ", "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4), ""
    '                            Exit For
    '                        End If
                        End If
                    
292                     If Trim(var_line(intloop)) = "413" Then 'ALB
294                         strALB = Val(var_line(intloop + 1))
                        End If
                    
296                     If Trim(var_line(intloop)) = "678" Then 'TP
298                         strTP = Val(var_line(intloop + 1))
                        End If
                    
300                     If Trim(var_line(intloop)) = "687" Then 'AST
302                         strAST = Val(var_line(intloop + 1))
                        End If
                    
304                     If Trim(var_line(intloop)) = "685" Then 'ALT
306                         strALT = Val(var_line(intloop + 1))
                        End If
                    
308                     If Trim(var_line(intloop)) = "168" Then 'APOA1
310                         strAPOA1 = Val(var_line(intloop + 1))
                        End If
                    
312                     If Trim(var_line(intloop)) = "151" Then 'APOB
314                         strAPOB = Val(var_line(intloop + 1))
                        End If
                    
316                     If Trim(var_line(intloop)) = "257" Then 'APOB
318                         strTBIL = Val(var_line(intloop + 1))
                        End If
                    
320                     If Trim(var_line(intloop)) = "293" Then 'APOB
322                         strDBIL = Val(var_line(intloop + 1))
                        End If
                    
                    Next
                
324                 If strALB <> "" And strTP <> "" Then
326                     strGLOB = strTP - strALB
328                     strGLOB = Format(strGLOB, "0.0")
330                     strResult = strResult & "|5501|" & strGLOB
                    End If
                
332                 If strALB <> "" And strGLOB <> "" Then 'ALB/GLOB
334                     strALBGLOB = strALB / strGLOB
336                     strALBGLOB = Format(strALBGLOB, "0.0")
338                     strResult = strResult & "|5502|" & strALBGLOB
                    End If
                
340                 If strALT <> "" And strAST <> "" Then 'AST/ALT
342                     strASTALT = strAST / strALT
344                     strASTALT = Format(strASTALT, "0.00")
346                     strResult = strResult & "|5503|" & strASTALT
                    End If
                
348                 If strAPOA1 <> "" And strAPOB <> "" Then 'APOA1/APOB
350                     strAPOA1APOB = strAPOA1 / strAPOB
352                     strAPOA1APOB = Format(strAPOA1APOB, "0.0")
354                     strResult = strResult & "|5504|" & strAPOA1APOB
                    End If
                
356                 If strTBIL <> "" And strDBIL <> "" Then 'strTBIL-strDBIL
358                     strTDBIL = strTBIL - strDBIL
360                     strTDBIL = Format(strTDBIL, "0.0")
362                     strResult = strResult & "|5505|" & strTDBIL
                    End If
                
364                 If strHDL <> "" And strLDL <> "" Then 'strHDL/strLDL
366                     strHL = strHDL / strLDL
368                     strHL = Format(strHL, "0.00")
370                     strResult = strResult & "|5506|" & strHL
                    End If
                
                End If

            End If
        End If
        Exit Sub
hErr:
372     WriteLog "E601解码", Erl() & "行," & Err.Description, "C601"
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
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
        Dim str_SendItem As String, str_EndString As String, str_tmp As String, lngRow As Long
        Dim strSendNull As String '未找到标本时传输的指令。
        Dim strRackNo As String '手工发送时指定架号，位号
        Dim strJ As String, strW As String
        Dim strTmp As String, strN As String
        
        Dim dateS As Date
        
        On Error GoTo hErr
        
100     WriteLog "双向(C601 开始)：", strSampleInfo, "C601"
'102     WriteLog "双向(C601 入参)：iSendStep=" & iSendStep & ",strResponse=" & strResponse & ",blnSuccess=" & blnSuccess, "C601"
104     blnSuccess = True
106         If iSendStep = 0 Then
                'If strSampleInfo <> "" Then WriteLog "双向(C601 开始)：" & strSampleInfo & " (" & mstrWaitSend & ")", "C601"
108             var_标本 = Split(strSampleInfo, "||")
            
110             SendSample = ""
112             For i_Count = LBound(var_标本) To UBound(var_标本)
114                 str_记录 = var_标本(i_Count)
                
116                 If Val(Split(str_记录, "|")(1)) = 0 Then
118                     str_标本号 = Trim(Split(str_记录, "|")(5))
                    Else
120                     str_标本号 = Val(Split(str_记录, "|")(1))
                    End If
122                 str_标本类型 = Split(str_记录, "|")(3)
124                 strRackNo = Trim(Split(str_记录, "|")(6))
126                 i_急诊 = Val(Split(str_记录, "|")(4))
                
128                 str_条码 = Trim(Split(str_记录, "|")(5))
130                 If str_条码 = "" Then str_条码 = str_标本号
                
132                 str_检验项目 = ""
                
134                 For i_Loop = 10 To UBound(Split(str_记录, "|"))
136                     If Val(Split(str_记录, "|")(i_Loop)) > 0 Then
138                         str_tmp = Split(str_记录, "|")(i_Loop)
140                         str_tmp = Replace(str_tmp, "not", "")
                        
142                         If InStr(str_检验项目, "^^^" & str_tmp & "^\") <= 0 Then
                                '两个血糖，只传一个
144                             str_检验项目 = str_检验项目 & "^^^" & str_tmp & "^\"
                            End If
                        End If
                    Next
                
146                 If str_检验项目 <> "" And str_条码 <> "" Then
148                     str_检验项目 = Mid(str_检验项目, 1, Len(str_检验项目) - 1)
150                     If InStr(mstrRack, "^") > 0 Then
152                         str_SendItem = "H|\^&|||H7600^1||||||TSDWN^REPLY" & Chr(&HD) & "P|1" & Chr(&HD)
154                         strN = "N"
                        Else
156                         str_SendItem = "H|\^&|||H7600^1||||||TSDWN^BATCH" & Chr(&HD) & "P|1" & Chr(&HD)
158                         strN = "A"
                        End If
                    
160                     str_SendItem = str_SendItem & "O|1|" & str_条码 & Space(24 - Len(str_条码))
162                     If InStr(mstrRack, "^") > 0 Then
164                         str_SendItem = str_SendItem & "|" & str_标本号 & Mid(mstrRack, InStr(mstrRack, "^")) & "|" & str_检验项目
                        Else
                            '手工发送的无Rack信息
166                         strJ = "": strW = ""
168                         If InStr(strRackNo, ",") > 0 Then
170                             strJ = Trim(Split(strRackNo, ",")(0))
172                             strW = Trim(Split(strRackNo, ",")(1))
                            
174                             If strJ <> "" Then strJ = "^" & Format(strJ, "00000")
176                             If strW <> "" Then strW = "^" & strW & "^^S1^SC"
                            End If
178                         str_SendItem = str_SendItem & "|" & str_标本号 & strJ & strW & "|" & str_检验项目
                        End If
180                     str_SendItem = str_SendItem & "|" & IIf(i_急诊 = 1, "S", "R")               'S-急诊,R-常规
182                     str_SendItem = str_SendItem & "||" & Format(Now, "yyyyMMddHHmmss")
184                     str_SendItem = str_SendItem & "||||" & strN & "||||1||||||||||O" & Chr(&HD)           '2015-03-07 将A改为N，白山路用A不能双向。
186                     str_SendItem = str_SendItem & "C|1|I|" & str_标本号 & "^^^^|G" & Chr(&HD)
188                     str_SendItem = str_SendItem & "L|1|N" & Chr(&HD) & Chr(&H3)
                              
190                     str_EndString = ""
192                     lngRow = 0
194                     Do While str_SendItem <> ""
                        
196                         If Len(str_SendItem) <= 230 Then
198                             lngRow = lngRow + 1
200                             str_EndString = str_EndString & Chr(&H2) & lngRow & str_SendItem & GetCheckCode(lngRow & str_SendItem) & vbCrLf
202                             str_SendItem = ""
                            Else
204                             lngRow = lngRow + 1
206                             str_tmp = Mid(str_SendItem, 1, 230) & Chr(&H17)
208                             str_EndString = str_EndString & Chr(&H2) & lngRow & str_tmp & GetCheckCode(lngRow & str_tmp) & vbCrLf
210                             str_SendItem = Mid(str_SendItem, 231)
                            End If
                        Loop
        
212                     mstrWaitSend = mstrWaitSend & "," & str_EndString & "," & Chr(&H4)
214                     SendSample = Chr(&H5)
216                     iSendStep = 1
                    Else
                        '未找到条码对应的项目
218                     If str_条码 <> "" Then
                            'strSendNull = GetIniKeyValue(App.Path & "\C601.ini", "应答", "无项目", Chr(&H4))
220                         SendSample = Chr(&H5)  'strSendNull
                            mstrWaitSend = mstrWaitSend & "," & Chr(&H4)
                            iSendStep = 1
                        End If
                    End If
                Next
222             If SendSample <> "" Then WriteLog "双向(E601 Send)：", SendSample, "C601"
                
224         ElseIf InStr(strResponse, Chr(&H6)) > 0 Then
226             If mstrWaitSend <> "" Then
                
228                 If Left(mstrWaitSend, 1) = "," Then
230                     mstrWaitSend = Mid(mstrWaitSend, 2)
                    End If
232                 If InStr(mstrWaitSend, ",") > 0 Then
234                     strTmp = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
236                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                    Else
238                     strTmp = mstrWaitSend
240                     mstrWaitSend = ""
                    End If
                    
242                 SendSample = strTmp
244                 WriteLog "双向(E601 发送)：", SendSample, "C601"
                End If
246             If mstrWaitSend = "" Then
248                 iSendStep = 0
250                 WriteLog "双向(E601 完成)", "C601", ""
                End If
252         ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
254             WriteLog "双向(E601 收到04 退出)", "C601", ""
256             mstrWaitSend = ""
258             iSendStep = 0
260         ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
262             WriteLog "双向(E601 收到05 退出)", "C601", ""
264             mstrWaitSend = ""
266             iSendStep = 0
268             SendSample = Chr(&H4)
            End If
270         'If SendSample <> "" Then
                '延时应答
272             'dateS = Now
                'Do
274             'Loop While DateDiff("s", dateS, Now) < 0.15
            'End If
'276     WriteLog "双向(C601 出参)：iSendStep=" & iSendStep & ",blnSuccess=" & blnSuccess, "C601"
        Exit Function
hErr:
278     WriteLog "E601-双向:", Erl() & "行," & Err.Description, "C601"
End Function


Private Function GetCheckCode(ByVal str_in As String) As String

    '获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function



