VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_mag2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'上海曲阳医院
'2011-07-14
'
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

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
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
 Dim lng_S As Long, lng_E As Long, str_Code As String, str_Line As String, var_Code As Variant
    Dim int_Loop As Integer, str_结果 As String, str_标本号 As String, str_项目 As String, str_日期 As String
    
    '-- 处理 应答
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H4)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HD)) > 0 Then strCmd = Chr(&H6)

    If Len(strSource) <> 0 Then mstrCode = mstrCode & strSource
    
    '--解析
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = 0
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        var_Code = Split(str_Code, Chr(&HD))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 2) = "O|" Then
                str_标本号 = Split(str_Line, "|")(2)
                str_日期 = "": str_项目 = "": str_结果 = ""
            ElseIf Left(str_Line, 2) = "R|" Then
                If str_日期 = "" Then
                    str_日期 = Trim(Split(str_Line, "|")(12))
                    If IsNumeric(str_日期) And Len(str_日期) > 13 Then
                        str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                    End If
                End If
                str_项目 = Trim(Split(str_Line, "|")(2))
                str_结果 = str_结果 & "|" & Replace(str_项目, "^", "") & "|" & Trim(Split(str_Line, "|")(3))

            End If
        Next
        
        If str_标本号 <> "" And str_结果 <> "" And str_日期 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_结果
            str_日期 = "": str_结果 = ""
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    
End Sub

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
    
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

    iSendStep = 0
    
End Function

