VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EasyBloodGas_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MEDICA EasyBloodGas 与以前的数据格式不一致
'厂商:MEDICA
'作者:李建飞
'用户:贵阳市肺科医院
'日期:2009-09-29

Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串
Private mStr位置 As String  '双向通信用。

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String
    Dim str年 As String, str月 As String, str日 As String, str指标 As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    str指标 = ",pH,PCO2,PO2,pH(T),PCO2(T),PO2(T),HCO2,BEb,BEecf,SBC,%SO2c,O2ct,A-aDO2,RI"
    If strSource Like "*-##-####;*" Then
        '----- 开始解析结果
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的消息
            strLine = GetStr_Section(strReserved, "=", ";")
            strReserved = Mid(strReserved, InStr(strReserved, ";") + 1)

            '--解析结果
            varLine = Split(strLine, Chr(&HD&))
            For iRow = LBound(varLine) To UBound(varLine)
                If varLine(iRow) Like "*ANALYSIS RESULTS*" Then
                    str标本号 = Val(Trim(Split(varLine(iRow), " ")(2)))
                    strItem = ""
                    str检验人标本质控 = " |血液|0"
                ElseIf varLine(iRow) Like "*-##-####;*" Then
                    str日期 = Replace(Replace(varLine(iRow), Chr(&H3F), ""), ";", "")
                    str月 = TrunMonth(Trim(Mid(str日期, 1, InStr(str日期, "-") - 1))): str日期 = Mid(str日期, InStr(str日期, "-") + 1)
                    str日 = Mid(str日期, 1, InStr(str日期, "-") - 1)
                    str年 = Mid(str日期, InStr(str日期, "-") + 1, 4)
                    If str月 Like "##" And str年 Like "####" And str日 Like "##" Then
                        str日期 = Format(CDate(str年 & "-" & str月 & "-" & str日), "yyyy-MM-dd")
                    Else
                        str日期 = Format(Now, "yyyy-MM-dd")
                    End If
                Else
                    strTmp = varLine(iRow)
                    If InStr(strTmp, "*****") <= 0 Then
                        If InStr(strTmp, "pH(T)") > 0 Then
                            strItem = strItem & "|pH(T)|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "PCO[(T)") > 0 Then
                            strItem = strItem & "|PCO2(T)|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "PO[(T)") > 0 Then
                            strItem = strItem & "|PO2(T)|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "PCO[") > 0 Then
                            strItem = strItem & "|PCO2|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "PO[") > 0 Then
                            strItem = strItem & "|PO2|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "pH") > 0 Then
                            strItem = strItem & "|pH|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "TCO[") > 0 Then
                            strItem = strItem & "|TCO2|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "HCO]") > 0 Then
                            strItem = strItem & "|HCO2|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "BEb") > 0 Then
                            strItem = strItem & "|BEb|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "BEecf") > 0 Then
                            strItem = strItem & "|BEecf|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "SBC") > 0 Then
                            strItem = strItem & "|SBC|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "%SO[c") > 0 Then
                            strItem = strItem & "|%SO2c|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "O2ct") > 0 Then
                            strItem = strItem & "|O2ct|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "A-aDO[") > 0 Then
                            strItem = strItem & "|A-aDO2|" & getNumber_Right(Mid(strTmp, 1, 13))
                        ElseIf InStr(strTmp, "RI") > 0 Then
                            strItem = strItem & "|RI|" & getNumber_Right(Mid(strTmp, 1, 13))
                        End If
                    End If
                End If
            Next
            If strItem <> "" And str日期 <> "" And Val(str标本号) > 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If Not strReserved Like "*-##-####;*" Then
                blnEnd = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
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
    Select Case Mid(Trim(UCase(strMonth)), 2, 3)
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



Public Function GetStr_Section(ByVal strSource As String, ByVal strStart As String, ByVal strEnd As String) As String
    '功能：取两个字符之间的内容返回,开始字符和结束字符可以相同
    'strSource: 源字符串
    'strStart : 开始字符
    'strEnd   ：结束字符
    '
    Dim lngLength As Long, strTmp As String, strTmpStart As String, i As Integer
    
    If strStart <> strEnd Then
        lngLength = InStr(strSource, strEnd) - InStr(strSource, strStart) + 1
    Else
        For i = -22350 To -22310
            strTmpStart = Chr(i)
            If InStr(strSource, strTmpStart) <= 0 And strStart <> strTmpStart Then
                Exit For
            End If
        Next
        strTmp = Mid(strSource, 1, InStr(strSource, strStart) - 1) & strTmpStart & Mid(strSource, InStr(strSource, strStart) + 1)
        lngLength = InStr(strTmp, strEnd) - InStr(strTmp, strTmpStart) + 1
    End If
    
    If lngLength < 0 Then
        GetStr_Section = Mid(strSource, InStr(strSource, strStart) + lngLength, Abs(lngLength))
    Else
        GetStr_Section = Mid(strSource, InStr(strSource, strStart), lngLength)
    End If
End Function




