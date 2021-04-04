VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EasyBloodGas_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MEDICA EasyBloodGas
'厂商:MEDICA
'作者:陈东
'用户:曲阳医院

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String
    Dim str年 As String, str月 As String, str日 As String, str时分 As String, str指标 As String
    Dim blnEnd As Boolean '是否已解析完
    
    Dim lng_S As Long '开始
    Dim lng_E As Long '结束
    Dim str_Line As String
    Dim lng_标本号 As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_结果 As String
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    str指标 = ",pH,PCO2,PO2,pH(T),PCO2(T),PO2(T),HCO2,BEb,BEecf,SBC,%SO2c,O2ct,A-aDO2,RI"
    
    mstrCode = mstrCode & strSource
    '---取完整的数据
    lng_S = InStr(mstrCode, Chr(&HD) & Chr(&H3D))     '以 开头，返回完整段落的开始位置
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ";")  '以 结束，返回完整段落的结束位置
        If lng_E > 0 Then
            lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD))
        Else
            lng_E = 0
        End If
    Else
        lng_E = 0
    End If
    
    Do While lng_S > 0 And lng_E > 0 '如果开始、结束都不为空
        '----- 开始解析结果
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 5) '返回从开始位置到结束位置之间的字符
        mstrCode = Mid(mstrCode, lng_E + 5)
        strItem = ""
        str日期 = ""
        str标本号 = ""
        
        var_data = Split(str_Code, Chr(&HD)) '取每段数据
        For iRow = LBound(var_data) To UBound(var_data)
            str_Line = var_data(iRow)
            
            If InStr(str_Line, "ANALYSIS RESULTS") > 0 Then
                str标本号 = Val(Split((str_Line), Chr(&H20))(2))
                strItem = ""
                str检验人标本质控 = " |血液|0"
            ElseIf InStr(str_Line, ";") > 0 Then
                str日期 = Replace(Replace(str_Line, Chr(&H3F), ""), ";", "")
                str月 = Mid(str日期, 1, InStr(str日期, "-") - 1): str时分 = Split(str日期, " ")(1)
                str日 = Split(str日期, " ")(0)
                str日 = Split(str日, "-")(1)
                str年 = Split(str日期, " ")(0)
                str年 = Split(str年, "-")(2)
                str日期 = Format(CDate(str年 & "-" & TrunMonth(str月) & "-" & str日) & " " & str时分, "yyyy-MM-dd hh:mm")
            Else
                If InStr(str_Line, "pH(T)") > 0 Then
                    strItem = strItem & "|pH(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PCO[(T)") > 0 Then
                    strItem = strItem & "|PCO2(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PO[(T)") > 0 Then
                    strItem = strItem & "|PO2(T)|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PCO[") > 0 Then
                    strItem = strItem & "|PCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PO[") > 0 Then
                    strItem = strItem & "|PO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "pH") > 0 Then
                    strItem = strItem & "|pH|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "TCO[") > 0 Then
                    strItem = strItem & "|TCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "HCO]") > 0 Then
                    strItem = strItem & "|HCO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "BEb") > 0 Then
                    If Split(trimString(str_Line), " ")(1) = "-" Then
                        strItem = strItem & "|BEb|" & Split(trimString(str_Line), " ")(1) & Split(trimString(str_Line), " ")(2)
                    Else
                        strItem = strItem & "|BEb|" & Split(trimString(str_Line), " ")(1)
                    End If
                    
                ElseIf InStr(str_Line, "BEecf") > 0 Then
                    If Split(trimString(str_Line), " ")(1) = "-" Then
                        strItem = strItem & "|BEecf|" & Split(trimString(str_Line), " ")(1) & Split(trimString(str_Line), " ")(2)
                    Else
                        strItem = strItem & "|BEecf|" & Split(trimString(str_Line), " ")(1)
                    End If
                    
                ElseIf InStr(str_Line, "SBC") > 0 Then
                    strItem = strItem & "|SBC|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "%SO[c") > 0 Then
                    strItem = strItem & "|%SO2c|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "O2ct") > 0 Then
                    strItem = strItem & "|O2ct|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "A-aDO[") > 0 Then
                    strItem = strItem & "|A-aDO2|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "RI") > 0 Then
                    strItem = strItem & "|RI|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "Hb") > 0 Then
                    strItem = strItem & "|Hb|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "FIO[") > 0 Then
                    strItem = strItem & "|FIO|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "B.P.") > 0 Then
                    strItem = strItem & "|BP|" & Split(trimString(str_Line), " ")(1)
                ElseIf InStr(str_Line, "PAT.TEMP") > 0 Then
                    strItem = strItem & "|PATTEMP|" & Split(trimString(str_Line), " ")(1)
                End If
            End If
            
        Next
        
        If strItem <> "" And str日期 <> "" And Val(str标本号) > 0 Then
            strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
        End If
            
        '---取完整的数据
        lng_S = InStr(mstrCode, Chr(&HD) & Chr(&H3D))     '以 开头，返回完整段落的开始位置
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ";")  '以 结束，返回完整段落的结束位置
            If lng_E > 0 Then
                lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD))
            Else
                lng_E = 0
            End If
        Else
            lng_E = 0
        End If
    Loop
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    '向设备发送‘数据收到’命令
    'strCmd = Chr(&H6&)
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

Private Function getNumber_Right(ByVal strSource As String) As String
    '从右开始取数字，遇到空格结束
    Dim strTmp As String, lngLength As Long
    Dim strIn As String
    strIn = strSource
    
    If Len(strIn) > 0 Then
        For lngLength = 1 To Len(strSource)
            strTmp = Right(strIn, 1)
            If IsNumeric(strTmp) Then
                getNumber_Right = strTmp & getNumber_Right
                strIn = Mid(strIn, 1, Len(strIn) - 1)
            ElseIf strTmp = "." Then
                If InStr(getNumber_Right, ".") <= 0 Then
                    getNumber_Right = strTmp & getNumber_Right
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                Else
                    Exit For
                End If
            ElseIf strTmp = " " Then
                If getNumber_Right <> "" Then
                    Exit For
                Else
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                End If
            End If
        Next
    End If
    If getNumber_Right <> "" Then
        If InStr(strSource, "- ") > 0 Then getNumber_Right = "-" & getNumber_Right
    End If
End Function

Private Function trimString(ByVal Str As String) As String
        Dim strTemp       As String
        
        strTemp = Trim(Str)
        Do While InStr(1, strTemp, Space(2)) > 0
                strTemp = Replace(strTemp, Space(2), Space(1))
        Loop
        trimString = strTemp
End Function




