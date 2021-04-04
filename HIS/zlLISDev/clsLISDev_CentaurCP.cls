VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CentaurCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:ADVIA centaur CP 化学发光免疫分析仪
'厂商:拜尔
'作者:陈东
'用户:贵阳市肿瘤医院

Private mstrCode As String

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
    
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim var_Code As Variant, str_Line As String, int_Loop As Integer
    Dim str_标本号 As String, str_日期 As String, str_项目 As String, str_结果 As String
    Dim str_Item As String
    On Error GoTo errH
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If str_Line Like "?O|*" Then
            
               If str_标本号 <> Split(str_Line, "|")(2) Then
                    If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Item
                    End If
                    str_标本号 = Split(str_Line, "|")(2)
                    str_Item = ""
               End If
               str_项目 = Split(str_Line, "|")(4)
            ElseIf str_Line Like "?R|*" Then
                If InStr(Split(str_Line, "|")(2), str_项目) = 1 And InStr(Split(str_Line, "|")(2), "^DOSE") > 0 Then
                    str_日期 = Split(str_Line, "|")(12)
                    If str_日期 <> "" And Len(str_日期) > 14 And IsNumeric(Mid(str_日期, 1, 14)) Then
                        str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                        If Not IsDate(str_日期) Then str_日期 = ""
                    Else
                        str_日期 = ""
                    End If
                    str_结果 = Trim(Split(str_Line, "|")(3))
                    If str_结果 <> "" Then
                        str_Item = str_Item & "|" & Replace(str_项目, "^", "") & "|" & str_结果
                    End If
                End If
            End If
        Next
        
        If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Item
        End If
        '--------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    Call WriteLog("CentaurCP", Err.Description & vbCrLf & mstrCode, str_Code)
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
    SendSample = ""
End Function

