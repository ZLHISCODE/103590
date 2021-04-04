VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_TBA_1200FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:TBA-1200FR 生化分析仪
'厂商:东芝
'作者:陈东
'修改人:王振涛,双向接口
'修改时间：2013-11-11
'用户:永川人民医院,新晃人民医院
Private mstrCode As String

Private mstrRes As String

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
    Dim str_Code As String
    Dim lng_S As Long, lng_E As Long, str_项目 As String, str_标本号 As String, str_日期 As String
    Dim str_Line As String, i As Integer, str_Tmp As String
    
    On Error GoTo errH
    If InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H2) & Chr(&H6) & Chr(&H3)
'    ElseIf InStr(strSource, Chr(&H2)) > 0 Then
'        strCmd = Chr(&H2) & Chr(&H6) & Chr(&H3)
    End If
    If strSource <> "" Then mstrCode = mstrCode & strSource
'    Debug.Print strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
   
    Do While lng_S > 0 And lng_E > 0
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------
        If str_Code Like "*" & Chr(&H2) & "R *" Then                                            '解析结果
            str_项目 = "": str_标本号 = "": str_日期 = ""
            
            str_标本号 = Trim(Mid(str_Code, 8, 8))
            str_日期 = Mid(str_Code, 35, 12)
            
            If IsNumeric(str_日期) Then
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                If Not IsDate(str_日期) Then str_日期 = ""
            Else
                str_日期 = ""
            End If
            
            str_Line = str_Code
            
            If InStr(str_Line, Chr(&H17)) > 0 Then str_Line = Mid(str_Line, 1, InStr(str_Line, Chr(&H17)))
            str_Line = Mid(str_Line, 50)
            
            For i = 0 To Len(str_Line) / 15
                If Len(str_Line) < 15 Then Exit For
                str_Tmp = Trim(Mid(str_Line, 5, 6))
                If str_Tmp <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 4)) & "|" & str_Tmp
                End If
                str_Line = Mid(str_Line, 16)
            Next
            '处理质控
            If str_项目 <> "" And Val(str_标本号) <= 0 Then
                str_标本号 = 10001
            End If
            
            If Val(str_标本号) > 0 And str_项目 <> "" And str_日期 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_项目
            End If
        ElseIf str_Code Like "*" & Chr(&H2) & "Q *" Then                                        '双向通讯
            str_项目 = "": str_标本号 = "": str_日期 = ""
            str_标本号 = Trim(Mid(str_Code, 8, 20))
            str_日期 = Format(Now, "YYYY-MM-DD hh:mm:ss")
            If Len(str_标本号) = 8 Then
                If Val(str_标本号) <> 0 Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号)
                End If
            ElseIf Len(str_标本号) = 14 Then
                If Val(str_标本号) <> 0 Then
                    strResult = strResult & "||" & str_日期 & "|^^" & Val(str_标本号)
                End If
            End If

            If strResult <> "" Then
                strCmd = "1|"                                                         '双向指令
                mstrRes = str_Code
            End If
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop

    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    Call WriteLog("TBA_1200FR.SendSample 错误: ", Err.Number, Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：病人姓名
    '   第3个元素：标本类型
    '   第4个元素：急诊标志
    '   第5个元素：样本条码
    '   第6个元素：盘号，杯号
    '   第7个元素：病人ID^性别^出生日期^年龄^姓名全拼^稀释倍数
    '   第8～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNo As String, iPos As Long
    
    On Error GoTo errHandle
    
    blnSuccess = True
    If strResponse <> "" Then
        strResponse = mstrRes
    End If
    If strResponse Like "*Q*" Or Len(Trim(strResponse)) = 0 Then    '允许发送
        If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
            '请求结束
            '生成设备串
            aItem = Split(strSample, "|")
            '标本或SampleID
            strBuffer = aItem(1)
            strSampleID = aItem(5)
            
            '读取反馈的架号和杯号
            strRackNo = ""
            
            
            If Len(Trim(strResponse)) > 0 Then
                strRackNo = Mid(strResponse, 28, 7)
            End If
            
            If Len(Trim(strSampleID)) > 0 Then
                'SampleID
                SendSample = "O " & Right("00000000000000" & strSampleID, 14) & Space(6) & strRackNo & "  1"
            Else
                SendSample = "O " & Right("00000000" & strBuffer, 8) & Space(12) & strRackNo & "  1"
            End If
            
            For i = 10 To UBound(aItem)
                SendSample = SendSample & Right(Space(4) & aItem(i), 4) & "1"
            Next
            
            SendSample = SendSample & ""
            Call WriteLog("TBA_1200FR", "TBA_1200FR.SendSample 发送的仪器数据: ", SendSample)
            iSendStep = 0
        Else
            SendSample = Chr(&H2) & Chr(&H6) & Chr(&H3)
            iSendStep = iSendStep + 1
        End If
    Else
        SendSample = Chr(&H2) & Chr(&H6) & Chr(&H3)
        iSendStep = 0
        blnSuccess = False    '发送失败
    End If
        
    Exit Function

errHandle:
    Call WriteLog("TBA_1200FR.SendSample 错误: ", Err.Number, Err.Description)
End Function





