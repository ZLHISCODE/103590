VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LX20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:临床生化分析系 DXC800/LX20
'厂商:美国贝克曼库尔特 Beckman coulter
'作者:陈东
'用户:烟台山医院

Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤
Private mstrACK As String     '是否应答ACK,是则应答ACK(06),否-应答 ETX(03)

Private mlngStep_Out As Long  '传双向消息步骤
Private mstr仪器号 As String

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
        Dim str_in As String, str_Line As String, str_条码 As String
        Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
        Dim str_WaitSendLine As String
    
100     str_in = strSource
102     If InStr(str_in, Chr(&H4) & Chr(&H1)) > 0 Then       'EOT SOH '仪器通讯请求 应答
104         lng_S = InStr(str_in, Chr(&H4) & Chr(&H1))
106         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 2)
108         mstrACK = Chr(&H3)  '04 01固定固定从06开始应答
110         strCmd = ACKChar
        
112         WriteLog "收到开始(EOT SOH)请求", "", str_in

114     ElseIf InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '仪器通讯结束
116         lng_S = InStr(str_in, Chr(&H4))
118         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
        
120         mlngStep_In = 0
122         Call ACKChar
124         If mstrWaitSend <> "" Then
126             strCmd = "1|"
128             strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
130             WriteLog "收到结束(EOT)请求.cmd=1", str_in, mstrWaitSend
            Else
132             WriteLog "收到结束(EOT)请求.", str_in, ""
            End If
        End If
    
134     If InStr(str_in, Chr(&H3)) > 0 Then                 '收到ETX(03)
136         lng_S = InStr(str_in, Chr(&H3))
138         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
140         mlngStep_Out = mlngStep_Out + 1
        End If
    
142     If InStr(str_in, Chr(&H6)) > 0 Then                 '收到ACK(06)
144         lng_S = InStr(str_in, Chr(&H6))
146         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
148         mlngStep_Out = mlngStep_Out + 1
        End If
    
150     If InStr(str_in, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
152         lng_S = InStr(str_in, Chr(&H5))
154         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
156         mstrACK = Chr(&H6)
158         strCmd = ACKChar
        End If
    
160     If InStr(str_in, Chr(&HA)) > 0 Then                 '收到换行(0A) 应答
162         lng_S = InStr(str_in, Chr(&HA))
164         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
166         strCmd = ACKChar
        End If
168     If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
170     lng_S = InStr(mstrCode, "[")
172     lng_E = InStr(mstrCode, "]")
174     If lng_S > 0 And lng_E > 0 And InStr(mstrCode, Chr(&HD)) > 0 Then
176         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
178         lng_E = InStr(mstrCode, Chr(&HD))
180         mstrCode = Mid(mstrCode, lng_E + 1)
        
182         If InStr(str_Line, ",801,06,") > 0 Then         '双向通讯- 开始
184             mstr仪器号 = ""
186             mlngStep_Out = 1
188             For i_Count = 0 To 3
190                 str_条码 = Split(str_Line, ",")(3 + i_Count)
192                 str_条码 = Trim(Replace(str_条码, "]", ""))
194                 If str_条码 <> "" Then
196                     strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & str_条码
                    End If
                Next
198             If strResult <> "" Then
200                 mstr仪器号 = Split(str_Line, ",")(0)
202                 strResult = Mid(strResult, 3)
204                 Call ACKChar
206                 strCmd = "1|" & ACKChar      '向仪器发通讯开始请求 EOT SOH
                End If
208             WriteLog "收到 801,06 双向开始 :", strCmd, mstrWaitSend
210         ElseIf InStr(str_Line, ",802,01,") > 0 Then     '仪器开始返回结果
212             WriteLog "收到 802,01 结果开始 :", strCmd, mstrWaitSend
214         ElseIf InStr(str_Line, ",802,05,") > 0 Then     '仪器返回结果完毕
216             WriteLog "收到 802,05 结果传完:", strCmd, mstrWaitSend
218         ElseIf InStr(str_Line, ",801,02,") > 0 Then     '仪器返回结果完毕
                'Call ACKChar
220             WriteLog "收到 801,02 条码确认", strCmd, mstrWaitSend 'strCmd = "1|" & ACKChar
222         ElseIf InStr(str_Line, ",802,03,") > 0 Then     '检验结果
224             str_日期 = Split(str_Line, ",")(3)
226             str_日期 = Right(str_日期, 4) & "-" & Mid(str_日期, 3, 2) & "-" & Left(str_日期, 2)
228             If Len(Split(str_Line, ",")(4)) >= 6 And IsNumeric(Split(str_Line, ",")(4)) Then
230                 str_日期 = str_日期 & " " & Mid(Split(str_Line, ",")(4), 1, 2) & ":" & Mid(Split(str_Line, ",")(4), 3, 2) & ":" & Mid(Split(str_Line, ",")(4), 5, 2)
                End If
232             str_标本号 = Trim(Split(str_Line, ",")(9))
234             str_项目 = Trim(Split(str_Line, ",")(10))
236             str_结果 = Trim(Split(str_Line, ",")(15))
238             If Len(Trim(str_标本号)) >= 12 Then str_标本号 = "^^" & str_标本号
240             strResult = str_日期 & "|" & str_标本号 & "| |血清|0|" & str_项目 & "|" & str_结果
            End If
        End If
242     If strCmd <> "" Then WriteLog "返回指令:", strCmd, mstrWaitSend
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
        Dim str_SendItem As String, str_WaitSendLine As String, str_Line As String, str_tmp As String
        On Error GoTo hErr
100     blnSuccess = True
102     If iSendStep = 0 Then
104         var_标本 = Split(strSampleInfo, "||")
106         WriteLog "SendSample(If begin)", strSampleInfo, strResponse
108         SendSample = ""
110         For i_Count = LBound(var_标本) To UBound(var_标本)
112             str_记录 = var_标本(i_Count)
            
114             If Val(Split(str_记录, "|")(1)) = 0 Then
116                 str_标本号 = Trim(Split(str_记录, "|")(5))
                Else
118                 str_标本号 = Val(Split(str_记录, "|")(1))
                End If
120             str_标本类型 = Split(str_记录, "|")(3)
                'SE 血清,SF 脑脊液,UR 随机尿,TU 定时尿,PL 血浆,OT 其他,BL 血液, 血

122             If InStr(str_标本类型, "血浆") > 0 Then
124                 str_标本类型 = "PL"
                ElseIf InStr(str_标本类型, "静脉血") > 0 Then
                    str_标本类型 = "SE"
126             ElseIf InStr(str_标本类型, "血清") > 0 Then
128                 str_标本类型 = "SE"
130             ElseIf InStr(str_标本类型, "血") > 0 Then
132                 str_标本类型 = "BL"
134             ElseIf InStr(str_标本类型, "脑脊液") > 0 Then
136                 str_标本类型 = "SF"
138             ElseIf InStr(str_标本类型, "定时尿") > 0 Then
140                 str_标本类型 = "TU"
142             ElseIf InStr(str_标本类型, "尿") > 0 Then
144                 str_标本类型 = "UR"
                Else
146                 str_标本类型 = "SE"
                End If
148             i_急诊 = Val(Split(str_记录, "|")(4))
150             str_条码 = Trim(Split(str_记录, "|")(5))
152             str_检验项目 = ""
154             str_记录 = Replace(str_记录, "|NGSP HBA1", "") '去掉三院的计算项目
156             For i_Loop = 10 To UBound(Split(str_记录, "|"))
158                 str_tmp = Trim(Split(str_记录, "|")(i_Loop))
160                 If Len(str_tmp) <= 3 Then
162                     str_tmp = str_tmp & Space(4 - Len(str_tmp))
                    End If
                    
164                 str_检验项目 = str_检验项目 & "," & str_tmp & ",0"
                    
                Next
            
            
166             If str_检验项目 <> "" And str_条码 <> "" Then
168                 str_检验项目 = Format(i_Loop - 10, "000") & str_检验项目
                
170                 str_SendItem = mstr仪器号 & ",801,01,0000,00,0"
172                 str_SendItem = str_SendItem & "," & IIf(i_急诊 = 1, "ST", "RO")             'ST-急诊,RO-常规
174                 str_SendItem = str_SendItem & "," & str_标本类型                                         '血清 SE,SF 脊液,UR 尿(随机),,
176                 str_SendItem = str_SendItem & "," & str_条码 & Space(15 - Len(str_条码))    '条码号,15位,不足15位补空格
178                 str_SendItem = str_SendItem & "," & Space(20)
180                 str_SendItem = str_SendItem & "," & Space(12)
182                 str_SendItem = str_SendItem & "," & Space(25)
184                 str_SendItem = str_SendItem & "," & Space(18)
186                 str_SendItem = str_SendItem & "," & Space(15)
188                 str_SendItem = str_SendItem & "," & Space(1)
190                 str_SendItem = str_SendItem & "," & str_标本号 & Space(15 - Len(str_标本号)) '标本号
192                 str_SendItem = str_SendItem & "," & Space(18)
194                 str_SendItem = str_SendItem & "," & Space(8)
196                 str_SendItem = str_SendItem & "," & Space(4)
198                 str_SendItem = str_SendItem & ",ZLSOFT" & Space(14)                         '位置
200                 str_SendItem = str_SendItem & "," & Space(3)                                '年龄
202                 str_SendItem = str_SendItem & ",5"                                          '年龄单位,默认5
204                 str_SendItem = str_SendItem & "," & Space(8)                                '生日
206                 str_SendItem = str_SendItem & ",M"                                          '性别
208                 str_SendItem = str_SendItem & "," & Space(45)
210                 str_SendItem = str_SendItem & ",111.1  "
212                 str_SendItem = str_SendItem & ",22.2"
214                 str_SendItem = str_SendItem & ",33.3"
216                 str_SendItem = str_SendItem & ",03"
218                 str_SendItem = str_SendItem & ",4.4444"
220                 str_SendItem = str_SendItem & "," & str_检验项目 & "]"
222                 str_SendItem = str_SendItem & GetCheckCode(str_SendItem) & vbCrLf
                    
224                 mstrWaitSend = mstrWaitSend & "|" & Chr(&H4) & Chr(&H1) & "|" & str_SendItem & "|" & Chr(&H4)
                    
                End If
            Next
            
226         If mstrWaitSend <> "" Then
228             str_Line = strResponse
230             If InStr(str_Line, Chr(&H4) & Chr(&H1)) > 0 Then
232                 SendSample = ""
234                 iSendStep = 0
                    Exit Function
                Else
236                 If Left(mstrWaitSend, 1) = "|" Then
238                     mstrWaitSend = Mid(mstrWaitSend, 2)
                    End If
240                 If InStr(mstrWaitSend, "|") > 0 Then
242                     SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
244                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
                    Else
246                     SendSample = mstrWaitSend
248                     mstrWaitSend = ""
                    End If
                End If
                
250             If SendSample <> "" Then iSendStep = 1
                
            End If
252         WriteLog "SendSample(If End)：", SendSample, "状态：" & iSendStep & vbNewLine & "待发送:" & mstrWaitSend
        Else
254         WriteLog "SendSample(Else Begin)", strResponse, "状态：" & iSendStep & vbNewLine & "待发送:" & mstrWaitSend
256         str_Line = strResponse
258         If InStr(str_Line, Chr(&H4) & Chr(&H1)) > 0 Then
                '补足前面的04，01 以便下次发送
260             If Left(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid(mstrWaitSend, 2)
262             If InStr(mstrWaitSend, "|") > 0 Then
264                 str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                Else
266                 str_WaitSendLine = mstrWaitSend
                End If
268             If Left(str_WaitSendLine, 1) = "[" Then mstrWaitSend = Chr(&H4) & Chr(&H1) & "|" & mstrWaitSend
                
270             SendSample = ""
272             iSendStep = 0

274             WriteLog "SendSample(Else End)", strResponse, "发到一半收到数据：" & iSendStep & vbNewLine & "待发送:" & mstrWaitSend
                Exit Function
            End If
276         str_WaitSendLine = ""
278         If InStr(str_Line, Chr(&H3)) > 0 Or _
               InStr(str_Line, Chr(&H4)) > 0 Or _
               InStr(str_Line, Chr(&H6)) > 0 Then                            '仪器返回确认信息
               
280             If mstrWaitSend <> "" Then
282                 If Left(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid(mstrWaitSend, 2)
                    
284                 If InStr(mstrWaitSend, "|") > 0 Then
286                     str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
288                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
                    Else
290                     str_WaitSendLine = mstrWaitSend
292                     mstrWaitSend = ""
                    End If
294                 SendSample = str_WaitSendLine
                End If
            End If
296         If mstrWaitSend = "" Then iSendStep = 0
298         If InStr(SendSample, Chr(&H4)) > 0 Then iSendStep = 0             '发完一个标本，即退出双向状态。

300         WriteLog "SendSample(Else End)", strResponse, "状态：" & iSendStep & vbNewLine & "待发送" & mstrWaitSend
        End If
        Exit Function
hErr:
302     WriteLog "SendSample", CStr(Erl()) & " " & Err.Description, "状态:" & iSendStep & ",收到:" & strResponse & ",标本：" & strSampleInfo & ",待发:" & mstrWaitSend & ",本次发:" & SendSample
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function


Private Function GetCheckCode(ByVal str_in As String) As String
    '计算校验码
    Dim ls_ch As String
    Dim ll_i As Long
    Dim ll_len As Long
    Dim ll_sum As Long

    If ((IsNull(str_in)) Or (str_in = "")) Then
        GetCheckCode = ""
    End If

    ll_len = Len(str_in)

    For ll_i = 1 To ll_len
        ls_ch = Mid(str_in, ll_i, 1)
        ll_sum = ll_sum + Asc(ls_ch)
    Next

    ll_sum = 256 - (ll_sum Mod 256)
    GetCheckCode = Hex(ll_sum)
    If Len(GetCheckCode) = 1 Then GetCheckCode = "0" & GetCheckCode
    
End Function








