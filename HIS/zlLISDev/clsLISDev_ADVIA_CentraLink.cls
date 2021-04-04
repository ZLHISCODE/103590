VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_CentraLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:西门子流水线
'厂商:
'作者:陈东
'用户:河科大一附院新区医院

Private mstrCode As String    '待解析串
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
        Dim str_In As String, str_Line As String, str_BarCode As String, var_line As Variant
        Dim str_Date As String, str_NO As String, str_Item As String, str_Value As String, str_ItemName As String
        Dim str_E As String, str_QC As String
        '-----------------------------------------------------------------------------
        Dim strQC As String
    
        On Error GoTo hErr
100     str_In = strSource
    
102     If InStr(str_In, Chr(&H5)) > 0 Then                 '收到ENQ(05) 应答 ETX(03)
104         strCmd = Chr(&H6)
106     ElseIf InStr(str_In, Chr(&H17)) > 0 Then            '收到ETB(17) 应答
108         strCmd = Chr(&H6)
110     ElseIf InStr(str_In, Chr(&HA)) > 0 Then
112         strCmd = Chr(&H6)
        End If
    
    
114     If Len(str_In) > 0 Then mstrCode = mstrCode & str_In
        '去掉 ETB
116     Do While mstrCode Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
118         lng_S = InStr(mstrCode, Chr(&H17))
120         lng_E = InStr(lng_S, mstrCode, Chr(&H2)) + 1
122         mstrCode = Mid(mstrCode, 1, lng_S - 1) & Mid(mstrCode, lng_E + 1)
        Loop
    
124     lng_S = InStr(mstrCode, Chr(&H5))
126     If lng_S > 0 Then
128         lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
130         lng_E = InStr(mstrCode, Chr(&H4))
        End If
132     Do While lng_S > 0 And lng_E > 0
    
134         str_NO = "":        str_Date = "":        str_Item = ""
136         str_E = "":        str_QC = "": str_ItemName = ""
138         str_Value = ""
        
        
140         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
142         mstrCode = Mid(mstrCode, lng_E + 1)
        
144         var_line = Split(str_Line, Chr(&HD))
        
146         For i_Count = LBound(var_line) To UBound(var_line)
148             str_Line = var_line(i_Count)
150             If Left(str_Line, 1) = Chr(&HA) Then str_Line = Mid(str_Line, 2)
152             If str_Line Like Chr(&H2) & "?O|*" Then
154                 str_Date = Trim(Split(str_Line, "|")(6))
                 
156                 str_Date = Left(str_Date, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2) & " " & Mid(str_Date, 9, 2) & ":" & Mid(str_Date, 11, 2) & ":" & Mid(str_Date, 13, 2)
                
158                 str_NO = Split(str_Line, "|")(2)
                
160                 str_E = Trim(Split(str_Line, "|")(5))
162                 str_QC = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")

164                 If Trim(str_NO) <> "" And Len(Trim(str_NO)) < 12 Then
                            '标本号
166                     str_NO = Trim(str_NO) & IIf(str_E = "S", "^1^", "")
168                 ElseIf Trim(str_NO) <> "" And Val(Trim(str_NO)) > 0 Then
                            '条码
170                     str_NO = "^" & IIf(str_E = "S", "1^", "^") & Trim(str_NO)
                    Else
172                     str_NO = ""
                    End If
                
174                 str_Item = ""
176             ElseIf str_Line Like Chr(&H2) & "?R|*" Then
178                 str_Value = Trim(Split(str_Line, "|")(3))
180                 If str_Value <> "" Then
182                     If InStr(str_Value, "^") > 0 Then str_Value = Split(str_Value, "^")(1) '--- 有结果为 -1^0.715 的格式。
184                     str_ItemName = Split(str_Line, "|")(2)
186                     If Split(str_ItemName, "^")(7) = "" Or Split(str_ItemName, "^")(7) = "DOSE" Then
188                         str_ItemName = Split(str_ItemName, "^")(3)
190                         str_Item = str_Item & "|" & str_ItemName & "|" & str_Value
                        End If
                    
                    End If
                End If
            Next
192         If IsDate(CDate(str_Date)) And str_NO <> "" And str_Item <> "" Then
194             strResult = strResult & "||" & str_Date & "|" & str_NO & "| |B|" & str_QC & str_Item
            End If
 
196         lng_S = InStr(mstrCode, Chr(&H5))
198         If lng_S > 0 Then
200             lng_E = InStr(lng_S, mstrCode, Chr(&H4))
            Else
202             lng_E = InStr(mstrCode, Chr(&H4))
            End If
        Loop
204     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
206     WriteLog "CentraLink", "第" & CStr(Erl()) & "行", Err.Description
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

    Dim var_Sample As Variant, i_Count As Integer, i_Loop As Integer, str_Record As String, str_NO As String
    Dim i_E As Integer, str_BarCode As String, str_TestItem As String, str_SampleType As String
    Dim str_SendItem As String, str_EndString As String, str_tmp As String, lngRow As Long
    Dim str_WaitSendLine As String, arrTestItem() As String, strLine As String
    On Error GoTo hErr
    blnSuccess = True
    ReDim arrTestItem(0) As String
    If iSendStep = 0 Then
        If strSampleInfo <> "" Then WriteLog "双向通讯(Begin)：", strSampleInfo, mstrWaitSend
        var_Sample = Split(strSampleInfo, "||")
        
        SendSample = ""
        For i_Count = LBound(var_Sample) To UBound(var_Sample)
            str_Record = var_Sample(i_Count)
            
            If Val(Split(str_Record, "|")(1)) = 0 Then
                str_NO = Trim(Split(str_Record, "|")(5))
            Else
                str_NO = Val(Split(str_Record, "|")(1))
            End If
            str_SampleType = Split(str_Record, "|")(3)
            i_E = Val(Split(str_Record, "|")(4))
            str_BarCode = Trim(Split(str_Record, "|")(5))
            
            str_TestItem = ""
            
            For i_Loop = 10 To UBound(Split(str_Record, "|"))
                If Trim("" & Split(str_Record, "|")(i_Loop)) <> "" Then
                    If InStr(str_TestItem, "^^^" & Replace(Replace(Split(str_Record, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\") <= 0 Then
                        '两个血糖，只传一个
                        str_TestItem = str_TestItem & "^^^" & Replace(Replace(Split(str_Record, "|")(i_Loop), "/1/not", ""), "1not", "") & "^\"
                        If Len(str_TestItem) > 160 Then
                            '分成160个字符一份
                             If arrTestItem(UBound(arrTestItem)) <> "" Then ReDim Preserve arrTestItem(UBound(arrTestItem) + 1)
                             arrTestItem(UBound(arrTestItem)) = str_TestItem
                             str_TestItem = ""
                        End If
                    End If
                End If
                    
            Next
            If str_TestItem <> "" Then
                If arrTestItem(UBound(arrTestItem)) <> "" Then ReDim Preserve arrTestItem(UBound(arrTestItem) + 1)
                arrTestItem(UBound(arrTestItem)) = str_TestItem
                str_TestItem = ""
            End If
                
            For lngRow = LBound(arrTestItem) To UBound(arrTestItem)
                str_TestItem = arrTestItem(lngRow)
                If str_TestItem <> "" And str_BarCode <> "" Then
                        
                    str_TestItem = Mid(str_TestItem, 1, Len(str_TestItem) - 1)
                    
                    strLine = "1H|\^&||||||||||P|1" & Chr(&HD) & Chr(&H3)
                    str_SendItem = "," & Chr(&H2) & strLine & GetCheckCode(strLine) & vbCrLf
                    
                    strLine = "2P|1|" & str_NO & "||||||" & Chr(&HD) & Chr(&H3)
                    str_SendItem = str_SendItem & "," & Chr(&H2) & strLine & GetCheckCode(strLine) & vbCrLf
                    
                    If str_SampleType = "血清" Then
                        str_SampleType = "Serum"
                    Else
                        str_SampleType = ""
                    End If
                    
                    strLine = "3O|1|" & str_BarCode & "||" & str_TestItem & "|" & IIf(i_E = 1, "S", "R") & "||" & Format(Now, "yyyyMMddHHmmss") & "||||A||||" & str_SampleType & "|" & Chr(&HD) & Chr(&H3)
                    str_SendItem = str_SendItem & "," & Chr(&H2) & strLine & GetCheckCode(strLine) & vbCrLf
                    
                    strLine = "4L|1" & Chr(&HD) & Chr(&H3)
                    str_SendItem = str_SendItem & "," & Chr(&H2) & strLine & "3D" & vbCrLf
        
                    If SendSample = "" Then
                        SendSample = SendSample & Chr(&H5)                          '加上 消息结束标志(EOT)
                        mstrWaitSend = mstrWaitSend & str_SendItem & "," & Chr(&H4)
                    Else
                        mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & str_SendItem & "," & Chr(&H4)
                    End If
                    iSendStep = 1
                End If
            Next
        Next
        If strSampleInfo <> "" Then WriteLog "双向通讯(End)：", SendSample, mstrWaitSend
    Else
        str_tmp = strResponse
        If InStr(str_tmp, Chr(&H6)) > 0 Then                 '收到ACK(06)
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
                
                SendSample = str_WaitSendLine
                
            End If
        End If
        If mstrWaitSend = "" Then iSendStep = 0
    End If
    Exit Function
hErr:
    WriteLog "CentraLink_SendSample", "第" & CStr(Erl()) & "行", Err.Description
End Function


Private Function GetCheckCode(ByVal str_In As String) As String

'获取校验码
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_In)
        lngAsc = lngAsc + Asc(Mid(str_In, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function





