VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cd3700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血常规 cell-DYN3700
'厂商:美国雅培
'作者:陈东
'用户:陈辉
'修改时间:2009-09-02
'修改人:李建飞

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
    Dim str_Last标本号 As String, str_Tmp As String
    '保存图形数据
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str结果处理方式 As String       '保存结果处理方式
    
    str_in = strSource
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = 0
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    End If
    
    '读取结果处理方式
    str结果处理方式 = Trim(ReadIni("CD3700", "结果处理方式", App.Path & "\CD3700.ini"))
    If str结果处理方式 = "" Then str结果处理方式 = "默认": Call WriteIni("CD3700", "结果处理方式", str结果处理方式, App.Path & "\CD3700.ini")
    
    Do While lng_S > 0 And lng_E > 0
        
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "开始", "收到完整数据", str_Line
        str_Line = Replace(Replace(str_Line, """", ""), Chr(&H3), "")
       
        var_line = Split(str_Line, ",")
        str_Tmp = Trim(var_line(0))
        str_标本号 = Trim(var_line(4))
        str_日期 = var_line(7)
        str_日期 = Format(CDate(Replace(str_日期, "/", "-")), "yyyy-MM-dd") & " " & var_line(8)
        If Not IsDate(str_日期) Then str_日期 = ""
        '判断接受数据是检验结果还是图形
        Select Case str_Tmp
            Case Chr(&H2) & ""  '检验结果
                For i_Count = 9 To UBound(var_line)
                    str_结果 = Trim(var_line(i_Count))
                    If str_结果 <> "" Then
                        Select Case str结果处理方式
                            Case "默认"
                                str_项目 = str_项目 & "|" & i_Count - 8 & "|" & str_结果
                            Case "无特殊字符"
                                str_项目 = str_项目 & "|" & i_Count - 8 & "|" & IIf(IsNumeric(str_结果), str_结果, " ")
                        End Select
                       ' str_项目 = str_项目 & "|" & i_Count - 8 & "|" & str_结果
                    End If
                Next
            Case Chr(&H2) & "WIC"   'WBC图形
                str_项目 = "|图形通送|0"
                For i_Count = 10 To UBound(var_line) - 1
                     str_WBC = str_WBC & ";" & var_line(i_Count)
                Next
            Case Chr(&H2) & "RBC"   'RBC图形
                str_项目 = "|图形通送|0"
                For i_Count = 10 To UBound(var_line) - 1
                    str_RBC = str_RBC & ";" & var_line(i_Count)
                Next
            Case Chr(&H2) & "PLT"   'PLT图形
                str_项目 = "|图形通送|0"
                For i_Count = 10 To UBound(var_line) - 1
                    str_PLT = str_PLT & ";" & var_line(i_Count)
                Next
        End Select

        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = 0
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        End If
        
        If Val(str_标本号) > 0 And str_日期 <> "" And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血液|0" & str_项目
            '处理图形数据
            If str_WBC <> "" Then strResult = strResult & vbCrLf & "WBC;0;256,67.5,7.5-,15-100,22.5-," & _
                                                            "30-200,37.5-,45-300,52.5-,60-400,67.5-" & str_WBC
            If str_RBC <> "" Then strResult = strResult & vbCrLf & "RBC;0;256,70,7-,14-50,21-,28-100," & _
                                                            "35-,42-150,49-,56-200,63-,70-250" & _
                                                            str_RBC
            If str_PLT <> "" Then strResult = strResult & vbCrLf & "PLT;0;256,64,8-5,16-10," & _
                                                            "24-15,32-20," & _
                                                            "40-25,48-30,56-35," & _
                                                            "64-40" & str_PLT
        End If
        str_标本号 = "": str_日期 = "": str_项目 = "": str_WBC = "": str_RBC = "": str_PLT = ""
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    WriteLog "结束", "解码结果", strResult

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




