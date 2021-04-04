VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_Glamour_4000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Glamour 4000 (魅力 4000生化分析仪)
'厂商:美国MD公司 阿根廷生产
'作者:陈东
'用户:曲阳医院

Private mstrCode As String

Private mstrWaitSend As String  '待发送的命令

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
   
    Dim lng_S As Long, lng_E As Long, str_In As String
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_SampleNo As String, str_Item As String, str_Value As String, str_Date As String
    Dim str_Bacode As String, str_SendLine As String
    
    strResult = ""
    str_In = strSource
    '-------------------------------------------------------------------------------------------------------------
    If InStr(str_In, Chr(&H5)) > 0 Then                                         '收到开始符 05 ,应答 06
        lng_S = InStr(str_In, Chr(&H5))
        str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
        strCmd = Chr(&H6)
    End If
    If InStr(str_In, Chr(&H3)) > 0 Then                                         '收到行结束符 03 ,应答 06
        lng_S = InStr(str_In, Chr(&H3))
        strCmd = Chr(&H6)
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Len(str_In) > 0 Then mstrCode = mstrCode & str_In
    
    If mstrCode Like "*1H|*L|1*" & Chr(&H4) & "*" Then                         '收到一个完整的记录,开始解析
        WriteLog "解析4000", mstrCode, ""
        Do While InStr(mstrCode, "L|1") > 0
            lng_S = InStr(mstrCode, "1H|")
            lng_E = InStr(mstrCode, "L|1")
            lng_E = InStr(lng_E, mstrCode, Chr(&H4))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            Do While str_Code <> ""
                lng_S = InStr(str_Code, Chr(&H2))
                lng_E = InStr(str_Code, Chr(&H3))
                
                If lng_S > 0 And lng_E > 0 Then
                    str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                    str_Code = Mid(str_Code, lng_E + 1)
'                    If InStr(str_Line, "CHO") > 0 Then Stop
                    If str_Line Like Chr(&H2) & "?O|*" Then                           '标本号,项目 行
                        str_SampleNo = Val(Split(str_Line, "|")(2))
                        If Len(str_SampleNo) = 13 Then                            '条码长度固定为13
                            str_SampleNo = "^^" & str_SampleNo
                        End If
                        If str_SampleNo = 0 Then str_SampleNo = ""
                        str_Item = Split(str_Line, "|")(4)
                    ElseIf InStr(str_Line, "R|") > 0 Then                       '项目 结果 行
                        If str_SampleNo <> "" And str_Item <> "" Then
                            If str_Item = Split(str_Line, "|")(2) Then
                                str_Date = Split(str_Line, "|")(12)
                                str_Value = Split(str_Line, "|")(3)
                                If str_Date <> "" Then
                                    str_Date = Mid(str_Date, 1, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2) & " " & Mid(str_Date, 9, 2) & ":" & Mid(str_Date, 11, 2) & ":" & Mid(str_Date, 13, 2)
'                                    If CDate(str_Date) <= CDate("2008-10-22") Then
                                    strResult = strResult & "||" & str_Date & "|" & str_SampleNo & "| |血清|0|" & Replace(str_Item, "^^^", "") & "|" & str_Value
'                                    Else
'                                        WriteLog "试用期已到", "请支持工作", ""
'                                    End If
                                End If
                            End If
                        End If
                        str_SampleNo = "": str_Item = ""
                        
                        
                    ElseIf InStr(str_Line, "Q|") > 0 Then                       '双向通讯- 条码信息
                        strCmd = "1|"
                        str_Bacode = Split(str_Line, "|")(2)
                        If str_Bacode <> "" Then
                            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^" & str_Bacode   '条码中包含有^号,所以这里少一个^
                        End If
                    End If
                Else
                    str_Code = ""
                End If
            Loop
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "解析4000", strResult, ""
    End If
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    
    Dim str_Record As String, str_SampleNo As String, str_SampleType As String, i_JiZhen As Integer, str_Bacode As String
    Dim str_SendItem As String, i_Loop As Integer, str_CheckItem As String
    Dim lngRow As Long, lngO As Long
    Dim str_In As String, lng_S As Long, lng_E As Long
    
    If strSampleInfo = "" Then Exit Function
    
    blnSuccess = True
    
    If iSendStep = 0 Then
        str_Record = strSampleInfo
        
        If UBound(Split(str_Record, "|")) < 10 Then Exit Function
        
        str_SampleNo = Val(Split(str_Record, "|")(1))
        str_SampleType = Split(str_Record, "|")(3)
        i_JiZhen = Val(Split(str_Record, "|")(4))
        str_Bacode = Trim(Split(str_Record, "|")(5))
        
        If Val(str_SampleNo) = 0 Then
            str_SampleNo = Right(str_Bacode, 10)
        End If
        
            If str_Bacode <> "" Then
            
            lngRow = 1
            
            str_SendItem = Chr(&H2) & lngRow & "H|\^&|" & Chr(&HD) & Chr(&H3)
            
            lngRow = lngRow + 1
            str_SendItem = str_SendItem & Chr(&H2) & lngRow & "P|" & "1|" & str_SampleNo & "|||Nobody" & "||19950101|M||||Dr|||||None|||||||Unknown^Blank" & Chr(&HD) & Chr(&H3)
    
            lngO = 1
            For i_Loop = 10 To UBound(Split(str_Record, "|"))
                str_CheckItem = Split(str_Record, "|")(i_Loop)
                If str_CheckItem <> "" Then
                    If lngRow = 7 Then lngRow = -1
                    lngRow = lngRow + 1
                    str_SendItem = str_SendItem & Chr(&H2) & lngRow & "O|" & lngO & "|" & str_Bacode & "||^^^" & str_CheckItem & "|||" & Format(Now, "yyyyMMdd") & "||||A||||Feo Color||||||||||Q" & Chr(&HD) & Chr(&H3)
                    lngO = lngO + 1
                End If
            Next
            
            If lngRow = 7 Then lngRow = -1
            lngRow = lngRow + 1
            
            str_SendItem = str_SendItem & Chr(&H2) & lngRow & "L|1|c" & Chr(&HD) & Chr(&H3) & Chr(&H2) & Chr(&H4) & Chr(&H3)
            mstrWaitSend = mstrWaitSend & str_SendItem
        End If
        
        If mstrWaitSend <> "" Then
            SendSample = Chr(&H5)
            iSendStep = 1
        End If
    Else
        str_In = strResponse
        SendSample = ""
        If InStr(str_In, Chr(&H6)) > 0 Or InStr(str_In, Chr(&H15)) > 0 Then           '双向通讯- 收到06,应答信息
            lng_S = InStr(str_In, Chr(&H6))
            If lng_S <= 0 Then
                lng_S = InStr(str_In, Chr(&H15))
            End If
            str_In = Mid(str_In, 1, lng_S - 1) & Mid(str_In, lng_S + 1)
            
            If mstrWaitSend <> "" Then
                lng_S = InStr(mstrWaitSend, Chr(&H2))
                lng_E = InStr(mstrWaitSend, Chr(&H3))
                
                str_SendItem = Mid(mstrWaitSend, lng_S, lng_E - lng_S + 1)
                mstrWaitSend = Mid(mstrWaitSend, lng_E + 1)
                If InStr(str_SendItem, Chr(&H4)) > 0 Then
                    SendSample = Chr(&H4)
                Else
                    SendSample = str_SendItem & GetCheckCode(str_SendItem) & vbCrLf
                End If
                
            End If
        End If
        If mstrWaitSend = "" Then iSendStep = 0
    End If
    
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
    '获取校验码
    Dim lngAsc As Long, i As Integer

    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc - 2 Mod 256)), 2)
End Function





