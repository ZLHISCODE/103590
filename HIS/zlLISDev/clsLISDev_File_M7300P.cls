VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_M7300P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:光电m7300p
'厂商:
'作者:董雨
'用户:长安区医院
'时间:2014-10

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '   标准函数：从文件获取检验结果
    '   strFile：包含路径的文件名
    '   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
    '   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    '   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    '   返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '   如：第5i个元素为检验项目，第5i+1个元素为检验结果

    Dim str操作员 As String, str日期 As String, strTime As String, str标本号 As String, str检验结果 As String
    Dim i As Long

    Dim objFSO As FileSystemObject
    Dim objText As textStream
    Dim strLine As String

    Set objFSO = New FileSystemObject
    Dim aResults() As String
    Dim lngLoop As Long
    Dim strPath As String
    Dim strErr As String             '错误提示

    ReDim Preserve aResults(0)
    On Error GoTo errHandle
    strPath = Mid(strFile, 1, InStrRev(strFile, "\"))
    If dtEnd = CDate("3000-12-31") Then
        dtEnd = Now
    End If
    lngLoop = DateDiff("d", dtStart, dtEnd)
    If lngLoop < 0 Then
        strErr = "开始日期不能大于结束日期！"
        ResultFromFile = aResults
        Exit Function
    End If
    For i = 0 To lngLoop    '循环获取日期内文件
        strFile = Dir(strPath & Format(DateAdd("d", i, dtStart), "YYMMDD") & ".txt") '时间段内的文件
        Do While strFile <> ""
            Set objText = objFSO.OpenTextFile(strPath & strFile, _
                ForReading, False, -2)
            Do Until objText.AtEndOfStream
                strLine = Trim(objText.ReadLine)
                strLine = Replace(strLine, vbTab, " ")
                strLine = TrimString(strLine)
                If InStr(strLine, "接口号") > 0 And UBound(Split(strLine, " ")) >= 6 Then
                    If str日期 = "" Then str日期 = "3014-01-01"
                    If Val(str标本号) <> Val(Split(strLine, " ")(1)) Or _
                       (InStr(Split(strLine, " ")(4), "入库") > 0 And Format(str日期, "yyyy-mm-dd") <> Format(Split(strLine, " ")(5), "yyyy-mm-dd")) Or _
                       (InStr(Split(strLine, " ")(4), "入库") = 0 And Format(str日期, "yyyy-mm-dd") <> Format(Split(strLine, " ")(6), "yyyy-mm-dd")) Then
                        str标本号 = Val(Split(strLine, " ")(1))
                        If InStr(Split(strLine, " ")(4), "入库") > 0 Then
                            If UBound(Split(strLine, " ")) >= 6 Then
                                strTime = " " & Split(strLine, " ")(6)
                            Else
                                strTime = ""
                            End If
                            str日期 = Format(Split(strLine, " ")(5) & strTime, "yyyy-mm-dd hh:nn:ss")
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            str检验结果 = " "

                        Else
                            If UBound(Split(strLine, " ")) >= 7 Then
                                strTime = " " & Split(strLine, " ")(7)
                            Else
                                strTime = ""
                            End If
                            str日期 = Format(Split(strLine, " ")(6) & strTime, "yyyy-mm-dd hh:nn:ss")
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            str检验结果 = Split(strLine, " ")(4)
                            str检验结果 = IIf(str检验结果 = "", " ", str检验结果)

                        End If
                        aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血液|0|" _
                                 & UCase(Trim(Replace(Replace(Split(strLine, " ")(3), "接口号：", ""), ":", ""))) & "|" & str检验结果
                    Else
                        If InStr(Split(strLine, " ")(4), "入库") > 0 Then
                            str检验结果 = " "

                        Else
                            str检验结果 = Split(strLine, " ")(4)
                            str检验结果 = IIf(str检验结果 = "", " ", str检验结果)

                        End If
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" _
                                 & UCase(Trim(Replace(Replace(Split(strLine, " ")(3), "接口号：", ""), ":", ""))) & "|" & str检验结果

                    End If
                Else
                    str日期 = "3014-01-01"
                    str标本号 = ""
                End If

            Loop
            Set objText = Nothing
            strFile = Dir
        Loop
    Next i
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_M7300P" & vbCrLf & strErr, strFile, Err.Number & " " & Err.Description

End Function
  

Private Function TrimString(strIn As String) As String
'替换字符串中多个空格为单个空格
    Dim strTemp As String
    strTemp = strIn
    Do While InStr(1, strTemp, Space(2)) > 0
        strTemp = Replace(strTemp, Space(2), Space(1))
    Loop
    TrimString = strTemp
End Function
