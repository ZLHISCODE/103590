VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SEAC_DELTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:意大利SEAC DELTA 特种蛋白分析仪
'厂商:意大利SEAC
'作者:陈东
'用户:德阳人民医院
'要设为流模式接收
Private mstrCode As String
Private mstrTotal As String '用于应答

Public Function GetAnswerCmd() As String
    '自动应答指令
    GetAnswerCmd = ",06"
End Function

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
    Dim str_Code As String, lng_S As Long, lng_E As Long
    Dim str_Tmp As String
    Dim var_Code As Variant, str_Line As String, i As Long, k As Long
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_last As String
    
    On Error GoTo errH
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    If strSource <> "" Then mstrTotal = mstrTotal & strSource
    '----------------------------------------
    lng_S = InStr(mstrTotal, ",01")
    Do While lng_S > 0
        mstrTotal = Mid(mstrTotal, lng_S + 1)
        lng_S = InStr(mstrTotal, ",01")
    Loop
    If Len(mstrTotal) / 3 >= 129 Then strCmd = ",06"
    '----------------------------------------
    lng_S = InStr(mstrCode, ",01")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ",04")
    Else
        lng_E = InStr(mstrCode, ",04")
    End If
    If lng_E > 0 Then
        If InStr(lng_E + 3, mstrCode, ",04") > 0 Then
            lng_E = InStr(lng_E + 3, mstrCode, ",04")
        Else
            If Mid(mstrCode, lng_E + 3) <> "" Then
                lng_E = 0
            End If
        End If
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 3)
        mstrCode = Mid(mstrCode, lng_E + 3)
        
        '-------------------------------------------------------
        str_Tmp = ""
        Do While InStr(str_Code, ",01") > 0
            lng_S = InStr(str_Code, ",01")
            If str_last <> Mid(str_Code, lng_S + 3 * 3, 128 * 3) Then
                str_tmp = str_tmp & Mid(str_Code, lng_S + 3 * 3, 128 * 3)
            End If
            str_last = Mid(str_Code, lng_S + 3 * 3, 128 * 3)
            
            str_Code = Mid(str_Code, lng_S + 3 * 3 + 129 * 3)
        Loop
        If str_Tmp <> "" Then str_Code = Mid_bin(str_Tmp, 1)
        WriteLog "seac_delta", "", str_Code
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like "[[]*[]]" Then
                str_标本号 = Split(var_Code(i + 1), "=")(1)
                str_日期 = Split(var_Code(i + 13), "=")(1) & "-" & Split(var_Code(i + 12), "=")(1) & "-" & Split(var_Code(i + 11), "=")(1)
                str_项目 = Split(var_Code(i + 14), "=")(1)
                str_结果 = Split(var_Code(i + 15), "=")(1)
                
                If Val(str_标本号) > 0 And IsDate(str_日期) And str_项目 <> "" And str_结果 <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0"
                    For k = LBound(Split(str_项目, ",")) To UBound(Split(str_项目, ","))
                        If Trim(Split(str_结果, ",")(k)) <> "" And Trim(Split(str_项目, ",")(k)) <> "" Then
                            strResult = strResult & "|" & Trim(Split(str_项目, ",")(k)) & "|" & Trim(Split(str_结果, ",")(k))
                        End If
                    Next
                End If
            End If
        Next
        '-------------------------------------------------------
        lng_S = InStr(mstrCode, ",01")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ",04")
        Else
            lng_E = InStr(mstrCode, ",04")
        End If
        If lng_E > 0 Then
            If InStr(lng_E + 3, mstrCode, ",04") > 0 Then
                lng_E = InStr(lng_E + 3, mstrCode, ",04")
            Else
                If Mid(mstrCode, lng_E + 3) <> "" Then
                    lng_E = 0
                End If
            End If
        End If
    
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "解析时出错", Err.Description, str_Code
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function

