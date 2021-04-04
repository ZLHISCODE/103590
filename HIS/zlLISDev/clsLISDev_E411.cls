VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_E411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:电化学发光 E411
'厂商:Roche 罗氏
'作者:陈东
'用户:榆林医院
'修改人:李建飞
'修改日期:2009-08-28

Private mstrCode As String    '待解析串

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
    str_in = strSource
        
    If InStr(str_in, Chr(&HA)) > 0 Then                 '收到ETX(03)
        strCmd = Chr(&H6)
    End If
    
    If InStr(str_in, Chr(&H5)) > 0 And strCmd = "" Then               '收到ENQ(05) 应答 ETX(03)
        strCmd = Chr(&H6)
    End If

    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
    If lng_S > 0 And lng_E > 0 Then
        
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "开始", "收到完整数据", str_Line
        
        If str_Line Like "*" & Chr(&H2) & "?R|*" Then     '检验结果
            var_line = Split(str_Line, Chr(&HA))
            
            For i_Count = LBound(var_line) To UBound(var_line)
                str_Line = var_line(i_Count)
                If str_Line Like Chr(&H2) & "?O|*" Then
                    str_日期 = Split(str_Line, "|")(6)
                    str_日期 = Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2)
                    
                    'str_标本号 = Mid(Split(str_Line, "|")(3), 1, InStr(Split(str_Line, "|")(3), "^") - 1)
                    str_标本号 = Replace(Split(str_Line, "|")(2), "@", "")
                    str_急诊 = Trim(Split(str_Line, "|")(5))
                    
                    str_质控 = "0"
                    If Trim(Split(str_Line, "|")(11)) = "Q" Then str_质控 = "1"
                    If Trim(Split(str_Line, "|")(11)) = "X\Q" Then str_质控 = "1"
                    
                    If Val(Trim(str_标本号)) > 0 Then
                        If Len(Trim(str_标本号)) < 12 Then
                            '标本号
                            str_标本号 = Trim(str_标本号) & IIf(str_急诊 = "S", "^1^", "")
                        Else
                            '条码
                            str_标本号 = "^" & IIf(str_急诊 = "S", "1^", "^") & Trim(str_标本号)
                        End If
                    Else
                        str_标本号 = ""
                    End If
                    
                    str_项目 = ""
                ElseIf str_Line Like Chr(&H2) & "?R|*" Then
                    str_结果 = Trim(Split(str_Line, "|")(3))
                    If str_结果 <> "" Then
                        If InStr(str_结果, "^") > 0 Then
                            str_项目 = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & Split(str_结果, "^")(1)
                        Else
                            str_项目 = "|" & Split(Split(str_Line, "|")(2), "^")(3) & "|" & str_结果
                        End If
                        If IsDate(CDate(str_日期)) And str_标本号 <> "" And str_项目 <> "" Then
                            If str_Last标本号 = str_标本号 Then
                                strResult = strResult & str_项目
                            Else
                                strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|" & str_质控 & str_项目
                                str_Last标本号 = str_标本号
                            End If
                        End If
                    End If
                End If
            Next
            If strResult <> "" Then strResult = Mid(strResult, 3)
            WriteLog "结束", "解码结果", strResult
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
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String
    Dim str_SendItem As String, str_Line As String
    
    SendSample = ""

    
End Function





