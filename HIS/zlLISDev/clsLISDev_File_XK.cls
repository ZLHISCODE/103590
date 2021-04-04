VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:细菌鉴定/药敏分析仪 XK型 微生物
'厂商:山东鑫科生物科技有限公司
'作者:王振涛
'用户:山西河津铝厂职工医院
'时间:2012-12-26
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()

    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim lngResultCount As Long
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str标本号 As String, str药敏项目 As String
    Dim str日期 As String, str细菌 As String, str敏感度 As String
    If Dir(strFile) = "" Then Exit Function
    Set objStream = objFile.OpenTextFile(strFile, ForReading)

    lngResultCount = -1
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "@*" Then

            ''有两个细菌以上的情况。 先返回上一个细菌结果
'            If Val(str标本号) > 0 And str细菌 <> "" And str日期 <> "" And str药敏项目 <> "" Then
'                lngResultCount = lngResultCount + 1
'                ReDim Preserve aResults(lngResultCount)
'                aResults(lngResultCount) = str日期 & "|" & Val((str标本号)) & "| |" & str细菌 & "|0" & str药敏项目
'                str日期 = "": str标本号 = "": str细菌 = "": str药敏项目 = ""
'            End If

            '一个标本开始,细菌结果
            If Split(strLine, "|")(1) <> "" Then
                str标本号 = Trim(Split(strLine, "|")(1))
            End If
            str日期 = Split(strLine, "|")(2)
'            If str日期 = "2012-11-14" Then
'                MsgBox "-------"
'            End If
            If str日期 >= Format(dtStart, "yyyy-MM-dd") And str日期 <= Format(dtEnd, "yyyy-MM-dd") Then

                If lngSampleNo <> -1 Then    '要指定标本号
                    If str标本号 <> lngSampleNo Then
                        '标本号不同，不返回
                        str标本号 = ""
                        str日期 = ""
                        str细菌 = ""
                    End If
                End If
            Else
                '日期不符，不返回
                str标本号 = ""
                str日期 = ""
                str细菌 = ""
            End If
        ElseIf InStr(strLine, "#") > 0 Then
            str细菌 = Split(strLine, "|")(2)
        ElseIf strLine Like "%*" Then
            '药敏结果
            If str标本号 <> "" Then
                '修改细菌检验结果,其中r*,blac,esbl 解析成耐药by陈玉强.
                str敏感度 = Split(strLine, "|")(6)
                Select Case str敏感度
                    Case "敏感": str敏感度 = "S"
                    Case "耐药": str敏感度 = "R"
                    Case "中介": str敏感度 = "I"
                End Select
                str药敏项目 = str药敏项目 & "|" & Split(strLine, "|")(2) & "|" & str敏感度 & "^1" & "^" & Split(strLine, "|")(4)
             End If
        ElseIf strLine Like "$*" Or strLine = "" Then
            '一个标本结束
            If Val(str标本号) >= 0 And str细菌 <> "" And str日期 <> "" And str药敏项目 <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str日期 & "|" & Val(str标本号) & "| |" & str细菌 & "|0" & str药敏项目
                str日期 = "": str标本号 = "": str细菌 = "": str药敏项目 = ""
            End If
            
            If Val(str标本号) >= 0 And str细菌 <> "" And str日期 <> "" And str药敏项目 = "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str日期 & "|" & Val(str标本号) & "| |细菌生长情况|0|细菌生长|" & str细菌
                str日期 = "": str标本号 = "": str细菌 = "": str药敏项目 = ""
            End If
            
        End If
    Loop
    ResultFromFile = aResults

End Function









