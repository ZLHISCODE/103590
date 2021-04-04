VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:临床生化分析系 MD2 (DXC800/LX20)
'厂商:美国贝克曼库尔特 Beckman coulter
'作者:陈东
'用户:西安公司方刚

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
    Dim str_In As String, str_Line As String, str_条码 As String
    Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
    
    str_In = strSource
    
    If InStr(str_In, Chr(&H4) & Chr(&H1)) > 0 Then       'EOT SOH '仪器通讯请求 应答
        lng_S = InStr(str_In, Chr(&H4) & Chr(&H1))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 2)
        strCmd = ACKChar
        
        If mlngStep_In <> 0 Then WriteLog "步骤" & mlngStep_In, "收到开始(EOT SOH)请求.", ""
        mlngStep_In = 1
    End If
        
    If InStr(str_In, Chr(&H4)) > 0 Then                 'EOT '仪器通讯结束
        lng_S = InStr(str_In, Chr(&H4))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        If mlngStep_In = 0 Then WriteLog "步骤" & mlngStep_In, "收到结束(EOT)请求.", ""
        mlngStep_In = 0
        Call ACKChar
        strCmd = Chr(&H4)
    End If
    
    If InStr(str_In, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
        lng_S = InStr(str_In, Chr(&H5))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        strCmd = Chr(&H3)
    End If
    
    If InStr(str_In, Chr(&HA)) > 0 Then                 '收到换行(0A) 应答
        lng_S = InStr(str_In, Chr(&HA))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        strCmd = ACKChar
    End If
    
    If Len(str_In) > 0 Then mstrCode = mstrCode & str_In
    
    lng_S = InStr(mstrCode, "[")
    lng_E = InStr(mstrCode, "]")
    If lng_S > 0 And lng_E > 0 And InStr(mstrCode, Chr(&HD)) > 0 Then
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        lng_E = InStr(mstrCode, Chr(&HD))
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If InStr(str_Line, ",801,06,") > 0 Then         '双向通讯- 开始
            mstr仪器号 = ""
            mlngStep_Out = 1
            For i_Count = 0 To 3
                str_条码 = Split(str_Line, ",")(3 + i_Count)
                str_条码 = Trim(Replace(str_条码, "]", ""))
                If str_条码 <> "" Then
                    strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & str_条码
                End If
            Next
            If strResult <> "" Then
                mstr仪器号 = Split(str_Line, ",")(0)
                strResult = Mid(strResult, 3)
                Call ACKChar
                strCmd = "1|" & ACKChar & Chr(&H4) & Chr(&H1)     '向仪器发通讯开始请求 EOT SOH
            End If
        ElseIf InStr(str_Line, ",801,02,") > 0 Then     '仪器返回确认信息

        ElseIf InStr(str_Line, ",802,01,") > 0 Then     '仪器开始返回结果

        ElseIf InStr(str_Line, ",802,05,") > 0 Then     '仪器返回结果完毕

        ElseIf InStr(str_Line, ",802,03,") > 0 Then     '检验结果
            str_日期 = Split(str_Line, ",")(3)
            str_日期 = Right(str_日期, 4) & "-" & Mid(str_日期, 3, 2) & "-" & Left(str_日期, 2)
            str_标本号 = Trim(Split(str_Line, ",")(9))
            str_项目 = Trim(Split(str_Line, ",")(10))
            str_结果 = Trim(Split(str_Line, ",")(15))
            strResult = str_日期 & "|^^" & str_标本号 & "| |血清|0|" & str_项目 & "|" & str_结果

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
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String, str仪器号 As String
    Dim str_SendItem As String, str_WaitSendLine As String, str_In As String
    
    blnSuccess = True
    If iSendStep = 0 Then
        var_标本 = Split(strSampleInfo, "||")
        SendSample = ""
        For i_Count = LBound(var_标本) To UBound(var_标本)
            str_记录 = var_标本(i_Count)
            
            If Val(Split(str_记录, "|")(1)) = 0 Then
                str_标本号 = Trim(Split(str_记录, "|")(5))
            Else
                str_标本号 = Val(Split(str_记录, "|")(1))
            End If
            str_标本类型 = Split(str_记录, "|")(3)
            i_急诊 = Val(Split(str_记录, "|")(4))
            str_条码 = Trim(Split(str_记录, "|")(5))
            str_检验项目 = ""
            For i_Loop = 10 To UBound(Split(str_记录, "|"))
                str_检验项目 = str_检验项目 & "," & Split(str_记录, "|")(i_Loop) & " ,0"
            Next
            
            
            If str_检验项目 <> "" And str_条码 <> "" Then
                str_检验项目 = Format(i_Loop - 10, "000") & str_检验项目
                If mstr仪器号 = "" Then
                    If Dir(App.Path & "\DM2.ini") <> "" Then
                        str仪器号 = ReadIni("DM2", "仪器号", App.Path & "\DM2.ini")
                    Else
                        str仪器号 = "0"
                        WriteIni "DM2", "仪器号", "0", App.Path & "\DM2.ini"
                    End If
                    If Not (str仪器号 Like "##" Or str仪器号 Like "#") Then
                        str仪器号 = "0"
                        WriteIni "DM2", "仪器号", "0", App.Path & "\DM2.ini"
                    End If
                    
                    mstr仪器号 = "[" & IIf(Len(str仪器号) = 1, " " & str仪器号, str仪器号)
                End If
                str_SendItem = mstr仪器号 & ",801,01,0000,00,0"
                str_SendItem = str_SendItem & "," & IIf(i_急诊 = 1, "ST", "RO")             'ST-急诊,RO-常规
                str_SendItem = str_SendItem & ",SE"                                         '血清 ,SF 脊液,UR 尿(随机),
                str_SendItem = str_SendItem & "," & str_条码 & Space(15 - Len(str_条码))    '条码号,15位,不足15位补空格
                str_SendItem = str_SendItem & "," & Space(20)
                str_SendItem = str_SendItem & "," & Space(12)
                str_SendItem = str_SendItem & "," & Space(25)
                str_SendItem = str_SendItem & "," & Space(18)
                str_SendItem = str_SendItem & "," & Space(15)
                str_SendItem = str_SendItem & "," & Space(1)
                str_SendItem = str_SendItem & "," & str_标本号 & Space(15 - Len(str_标本号)) '标本号
                str_SendItem = str_SendItem & "," & Space(18)
                str_SendItem = str_SendItem & "," & Space(8)
                str_SendItem = str_SendItem & "," & Space(4)
                str_SendItem = str_SendItem & ",ZLSOFT" & Space(14)                         '位置
                str_SendItem = str_SendItem & "," & Space(3)                                '年龄
                str_SendItem = str_SendItem & ",5"                                          '年龄单位,默认5
                str_SendItem = str_SendItem & "," & Space(8)                                '生日
                str_SendItem = str_SendItem & ",M"                                          '性别
                str_SendItem = str_SendItem & "," & Space(45)
                str_SendItem = str_SendItem & ",111.1  "
                str_SendItem = str_SendItem & ",22.2"
                str_SendItem = str_SendItem & ",33.3"
                str_SendItem = str_SendItem & ",03"
                str_SendItem = str_SendItem & ",4.4444"
                str_SendItem = str_SendItem & "," & str_检验项目 & "]"
                str_SendItem = str_SendItem & GetCheckCode(str_SendItem) & vbCrLf
                
                If SendSample = "" Then
                    SendSample = Chr(&H4) & Chr(&H1)
                    mstrWaitSend = mstrWaitSend & "|" & str_SendItem & "|" & Chr(&H4)                          '加上 消息结束标志(EOT)
                Else
                    mstrWaitSend = mstrWaitSend & "|" & str_SendItem & "|" & Chr(&H4)
                End If
                iSendStep = 1
            End If
        Next
    Else
        str_In = strResponse
        If InStr(str_In, Chr(&H6)) > 0 Or InStr(str_In, Chr(&H3)) > 0 Then
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "|" Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                If InStr(mstrWaitSend, "|") > 0 Then
                    str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
                Else
                    str_WaitSendLine = mstrWaitSend
                    mstrWaitSend = ""
                End If
                SendSample = str_WaitSendLine
            End If
        End If
        If mstrWaitSend = "" Then iSendStep = 0
    End If
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function


Private Function GetCheckCode(ByVal str_In As String) As String
    '计算校验码
    Dim ls_ch As String
    Dim ll_i As Long
    Dim ll_len As Long
    Dim ll_sum As Long

    If ((IsNull(str_In)) Or (str_In = "")) Then
        GetCheckCode = ""
    End If

    ll_len = Len(str_In)

    For ll_i = 1 To ll_len
        ls_ch = Mid(str_In, ll_i, 1)
        ll_sum = ll_sum + Asc(ls_ch)
    Next

    ll_sum = 256 - (ll_sum Mod 256)
    GetCheckCode = Hex(ll_sum)
    If Len(GetCheckCode) = 1 Then GetCheckCode = "0" & GetCheckCode

    
End Function


