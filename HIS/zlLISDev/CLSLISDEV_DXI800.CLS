VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DXI800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:免疫分析DXI800(双向)
'厂商:BeackMan 贝克曼
'作者:陈东
'用户:妇儿医院
'日期:2014-10

Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤

Private mlngStep_Out As Long  '传双向消息步骤
Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.

Private mstrWaitSend As String '待发送标本

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
        Dim str_Last标本号 As String
    
100     strResult = ""
102     str_in = strSource
        
104     If InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '仪器通讯结束
106         mlngStep_In = 0
        End If
    
108     If InStr(str_in, Chr(&H3)) > 0 Then                 '收到ETX(03)
110         lng_S = InStr(str_in, Chr(&H3))
112         mlngStep_Out = mlngStep_Out + 1
114         strCmd = Chr(&H6)
        End If
    
116     If InStr(str_in, Chr(&H6)) > 0 Then                 '收到ACK(06)
118         lng_S = InStr(str_in, Chr(&H6))
        
120         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
122         strReserved = str_in
124         str_in = ""
        
126         mlngStep_Out = mlngStep_Out + 1
128         If mstrWaitSend <> "" Then
            
130             If Left(mstrWaitSend, 1) = "," Then
132                 mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
134             If InStr(mstrWaitSend, ",") > 0 Then
136                 str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
138                 mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
140                 str_WaitSendLine = mstrWaitSend
142                 mstrWaitSend = ""
                End If
            
144             If InStr(",1H,2P,3O,4L,", Left(str_WaitSendLine, 3)) > 0 Then
146                 strCmd = "1|" & str_WaitSendLine
                End If
148             WriteLog "收到06,发送到仪器的指令:" & str_WaitSendLine, "", ""
                Exit Sub
            End If
        End If
    
150     If InStr(str_in, Chr(&H5)) > 0 And strCmd = "" Then               '收到ENQ(05) 应答 ETX(03)
152         strCmd = Chr(&H6)
        End If
    
154     If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
156     lng_S = InStr(mstrCode, Chr(&H5))
158     lng_E = InStr(mstrCode, Chr(&H4))
160     If lng_S > 0 And lng_E > 0 Then
        
162         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
164         mstrCode = Mid(mstrCode, lng_E + 1)
166         WriteLog "开始-收到完整数据" & str_Line, "", ""
        
168         If str_Line Like "*" & Chr(&H2) & "?Q|*" Then          '双向通讯- 开始
170             mstrRack = ""
172             mlngStep_Out = 1
174             var_line = Split(str_Line, Chr(&HA))
            
176             For i_Count = LBound(var_line) To UBound(var_line)
178                 str_Line = var_line(i_Count)
180                 If str_Line Like Chr(&H2) & "?Q|*||ALL||||||||*" Then   '扫条码
182                     str_条码 = Split(str_Line, "|")(2)
                    
184                     mstrRack = "" 'Split(str_条码, "^")(2) & "^" & Split(str_条码, "^")(3) & "^" & Split(str_条码, "^")(4)
186                     str_条码 = Split(str_条码, "^")(1)
                    
188                     If Val(str_条码) <> 0 Then
190                         strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_条码)
                        End If
                    End If
                Next
            
192             If strResult <> "" Then
194                 strResult = Mid(strResult, 3)
196                 strCmd = "1|"       '向仪器发通讯开始请求 05
198                 WriteLog "结束-解码结果" & strResult, "", ""
                End If
200         ElseIf str_Line Like "*" & Chr(&H2) & "?R|*" Then     '检验结果
202             var_line = Split(str_Line, Chr(&HA))
            
204             For i_Count = LBound(var_line) To UBound(var_line)
206                 str_Line = var_line(i_Count)
208                 If str_Line Like Chr(&H2) & "?O|*" Then
210                     str_标本号 = Split(str_Line, "|")(2)
212                     str_急诊 = Trim(Split(str_Line, "|")(5))
214                     str_质控 = "0"
216                     If (Trim(Split(str_Line, "|")(11)) = "Q") Or (Trim(Split(str_Line, "|")(11)) = "X\Q") Then str_质控 = "1"

                        '质控标本
218                     If Trim(str_标本号) Like "PC*1" Then
220                         str_质控 = "1"
222                         str_标本号 = "3100"
224                     ElseIf Trim(str_标本号) Like "PC*2" Then
226                         str_质控 = "1"
228                         str_标本号 = "3200"
                        End If
                    
230                     If Val(Trim(str_标本号)) > 0 Then
232                         If Len(Trim(str_标本号)) < 12 Then
                                '标本号
234                             str_标本号 = Trim(str_标本号) & IIf(str_急诊 = "S", "^1", "")
                            Else
                                '条码
236                             str_标本号 = "^" & IIf(str_急诊 = "S", "^1", "^") & Trim(str_标本号)
                            End If
                        Else
238                         str_标本号 = ""
                        End If
                    
240                     str_项目 = ""
242                 ElseIf str_Line Like Chr(&H2) & "?R|*" Then
244                     str_日期 = Trim(Split(str_Line, "|")(12))
246                     str_日期 = Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
248                     str_结果 = Trim(Split(str_Line, "|")(3))
250                     If str_结果 <> "" Then
252                         If InStr(str_结果, "^") > 0 Then
254                             str_项目 = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & Split(str_结果, "^")(1)
                            Else
256                             str_项目 = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & str_结果
                            End If
258                         If IsDate(CDate(str_日期)) And str_标本号 <> "" And str_项目 <> "" Then
260                             If str_Last标本号 = str_标本号 Then
262                                 strResult = strResult & str_项目
                                Else
264                                 strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|" & str_质控 & str_项目
266                                 str_Last标本号 = str_标本号
                                End If
                            End If
                        End If
                    End If
                Next
268             If strResult <> "" Then strResult = Mid(strResult, 3)
270             WriteLog "结束-解码结果" & strResult, "", ""
            End If
        End If
        Exit Sub
hErr:
272     WriteLog "Analyse " & Erl() & "行，" & Err.Description, "", ""
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
        Dim str_SendItem As String, str_Line As String
    
        On Error GoTo hErr
100     WriteLog "开始" & strSampleInfo, "", ""
102     If iSendStep = 0 Then
104         var_标本 = Split(strSampleInfo, "||")
106         SendSample = ""
108         For i_Count = LBound(var_标本) To UBound(var_标本)
110             str_记录 = var_标本(i_Count)
            
112             If Val(Split(str_记录, "|")(1)) = 0 Then
114                 str_标本号 = Trim(Split(str_记录, "|")(5))
                Else
116                 str_标本号 = Val(Split(str_记录, "|")(1))
                End If
118             str_标本类型 = Split(str_记录, "|")(3)
120             i_急诊 = Val(Split(str_记录, "|")(4))
122             str_条码 = Trim(Split(str_记录, "|")(5))
124             str_检验项目 = ""
            
126             For i_Loop = 10 To UBound(Split(str_记录, "|"))
128                 If Trim(Split(str_记录, "|")(i_Loop)) <> "" Then
130                     str_检验项目 = str_检验项目 & "^^^" & Trim(Split(str_记录, "|")(i_Loop)) & "\"
                    End If
                Next
            
132             If str_检验项目 <> "" And str_条码 <> "" Then
134                 str_检验项目 = Mid(str_检验项目, 1, Len(str_检验项目) - 1)
                
136                 str_Line = "1H|\^&|" & Chr(&HD) & Chr(&H3)
138                 str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
140                 str_Line = "2P|1|" & str_标本号 & "|" & Chr(&HD) & Chr(&H3)
142                 str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
144                 str_Line = "3O|1|" & str_条码 & "|" & mstrRack & "|" & str_检验项目 & "|" & IIf(i_急诊 = 1, "S", "R")
146                 str_Line = str_Line & "||||||A||||Serum|" & Chr(&HD) & Chr(&H3)
148                 str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
150                 str_Line = "4L|1|F" & Chr(&HD) & Chr(&H3)
152                 str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
                
154                 If SendSample = "" Then
156                     SendSample = SendSample & Chr(&H5)                              '加上 消息结束标志(EOT)
158                     iSendStep = 1
160                     mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)
                    Else
162                     mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_SendItem & Chr(&H4)
                    End If
                End If
            Next
        Else
164         If InStr(strResponse, Chr(&H6)) > 0 Then
166             If mstrWaitSend <> "" Then
168                 If Left(mstrWaitSend, 1) = "," Then
170                     mstrWaitSend = Mid(mstrWaitSend, 2)
                    End If
172                 If InStr(mstrWaitSend, ",") > 0 Then
174                     SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
176                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                    Else
178                     SendSample = mstrWaitSend
180                     mstrWaitSend = ""
                    End If
182                 WriteLog "收到06,发送到仪器的指令," & SendSample, "", ""
                End If
            End If
184         If mstrWaitSend = "" Then iSendStep = 0
        End If
186     WriteLog "结束-发送到仪器的信息" & SendSample, "", ""
        Exit Function
hErr:
188     WriteLog "SendSample " & Erl() & "行," & Err.Description, "", ""
    
End Function

'获取校验码
Private Function GetCheckCode(ByVal str_in As String) As String
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

