VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:FIA200
'厂商:时间分辨仪
'作者:王振涛
'用户:泽州县医院
'日期:2012-05-3
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
    '标准函数：从文件获取检验结果
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
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果

    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str操作员 As String, str日期 As String
    Dim strFilePath As String, intCount As Long, lngLoop As Long

    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant
    Dim strItem As String
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)

    For intCount = 0 To lngLoop
        str日期 = Format(dtStart, "yyyyMMdd")
        strFile = Dir(strFilePath & "\*" & str日期 & "*.TXT")
        Do While strFile <> ""
            str日期 = Format(Mid(strFile, 1, 4) & "-" & Mid(strFile, 5, 2) & "-" & Mid(strFile, 7, 2) & " " & _
                      Mid(strFile, 9, 2) & ":" & Mid(strFile, 11, 2) & ":" & Mid(strFile, 13, 2), "yyyy-mm-dd hh:mm:ss")
            strItem = Mid(strFile, 16, 5)
            Open strFilePath & "\" & strFile For Input As #1    ' 打开文件。
            Do While Not EOF(1)    ' 循环至文件尾。
                Line Input #1, strLine    ' 读入一行数据并将其赋予某变量。

                If strLine <> "" Then
                    str标本号 = Mid(strLine, 1, InStr(strLine, " ") - 1)
                    If Val(str标本号) > 0 Then
                        str检验结果 = strItem & "|" & Mid(Replace(strLine, " ", ""), Len(str标本号) + 1)
                        If str日期 <> "" And strItem <> "" And str检验结果 <> "" Then
                                If aResults(UBound(aResults)) <> "" Then
                                    ReDim Preserve aResults(UBound(aResults) + 1)
                                End If
                                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血液|0|" & str检验结果
                        End If
                    End If
                End If
                lngRow = lngRow + 1
                DoEvents
            Loop
            Close #1    ' 关闭文件。
            strFile = Dir
        Loop

    Next

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_FIA200.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function









