VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AVL9130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AVL 9130
'厂商:罗氏
'作者:陈东
'用户:晋煤集团医院

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
    Dim str_Code As String, lng_S As Long, lng_E As Long
    Dim var_line As Variant, str_Line As String, int_Loop As Integer
    Dim str_标本号 As String, str_日期 As String, str_项目 As String
    
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        var_line = Split(str_Code, Chr(&HD) & Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_line) To UBound(var_line)
            str_Line = var_line(int_Loop)
            If str_Line Like "*AVL 9130*" Then
                str_标本号 = "": str_日期 = "": str_项目 = ""
                str_日期 = var_line(int_Loop + 2)
                
                str_日期 = "20" & Mid(str_日期, 6, 2) & "-" & TrunMonth(Mid(str_日期, 3, 3)) & "-" & Left(str_日期, 2)
                If Not IsDate(str_日期) Then str_日期 = Format(Now, "yyyy-MM-dd")
                
                str_标本号 = var_line(int_Loop + 8)
                If str_标本号 Like "Sample No.*" Then
                    str_标本号 = Val(Replace(str_标本号, "Sample No.", ""))
                    If Val(str_标本号) <> 0 Then
                        
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0"
                        str_项目 = var_line(int_Loop + 11)
                        strResult = strResult & "|" & Trim(Split(str_项目, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_项目, "=")(1), "mmol/L", ""), "", "")))
                        str_项目 = var_line(int_Loop + 12)
                        strResult = strResult & "|" & Trim(Split(str_项目, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_项目, "=")(1), "mmol/L", ""), "", "")))
                        str_项目 = var_line(int_Loop + 13)
                        strResult = strResult & "|" & Trim(Split(str_项目, "=")(0)) & "|" & Val(Trim(Replace(Replace(Split(str_项目, "=")(1), "mmol/L", ""), "", "")))
                    End If
                End If
            End If
        Next
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    
    SendSample = ""
    
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    '转换英文的月份为数字
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function

