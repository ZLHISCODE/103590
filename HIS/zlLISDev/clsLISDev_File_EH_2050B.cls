VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH_2050B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:尿液分析 全自动尿沉渣分析仪
'厂商:迈瑞
'作者:王振涛
'用户:朔州平鲁区人民医院
'日期:2013-06-14
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
    Dim StrTime As String
    Dim objStream As TextStream
    Dim strLine As String, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant, strItems As String, strImages As String
    Dim intStep As Integer, strItem As String, strImage As String
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)

    For intCount = 0 To lngLoop
        str日期 = Format(DateAdd("d", intCount, dtStart), "yyyyMMdd")
        strFile = Dir(strFilePath & "\*" & str日期 & "*.TXT")
        Do While strFile <> ""

            Open strFilePath & "\" & strFile For Input As #1    ' 打开文件。
            Do While Not EOF(1)    ' 循环至文件尾。
                Line Input #1, strLine    ' 读入一行数据并将其赋予某变量。
                strLine = Replace(strLine, vbTab, "")
                Select Case intStep
                Case 1
                    If InStr(strLine, "Sediment Data Start") > 0 Then
                        intStep = 2
                    Else
                        strItem = Trim(Split(strLine, ",")(2))
                        str检验结果 = Trim(Split(strLine, ",")(5))
                        If str检验结果 <> "" Then
                            str检验结果 = Trim(Mid(str检验结果, 1, Len(str检验结果) - 1))
                        End If
                    End If
                    If strItem <> "" And str检验结果 <> "" Then
                        strItems = strItems & "|" & strItem & "|" & str检验结果
                        strItem = ""
                        str检验结果 = ""
                    End If
                Case 2
                    If InStr(strLine, "Image filename Start") > 0 Then
                        intStep = 3
                    Else
                        strItem = Trim(Split(strLine, ",")(2))
                        str检验结果 = Trim(Split(strLine, ",")(3))
                    End If
                    If strItem <> "" And str检验结果 <> "" Then
                        strItems = strItems & "|" & strItem & "|" & str检验结果
                        strItem = ""
                        str检验结果 = ""
                    End If
                Case 3
                    If InStr(strLine, "--End--") > 0 Then
                        intStep = 4
                    Else
                        If strImage = "" Then
                            intCount = 1
                            strImage = Trim(Split(strLine, ",")(2))
                            strImages = intCount & ";110;0;" & Trim(Mid(strImage, 1, Len(strImage) - 1))
                        Else
                            intCount = intCount + 1
                            strImage = Trim(Split(strLine, ",")(2))
                            strImage = intCount & ";110;0;" & Trim(Mid(strImage, 1, Len(strImage) - 1))
                            strImages = strImages & "^" & strImage
                        End If
                    End If
                End Select

                If InStr(strLine, "ItemName=UI_ReportDate") > 0 Then
                    StrTime = Trim(Split(strLine, ",")(3))
                    StrTime = Mid(StrTime, 1, Len(StrTime) - 1)
                    StrTime = Format(StrTime, "yyyy-mm-dd HH:MM:SS")
                ElseIf InStr(strLine, "ItemName=SampleID") > 0 Then
                    str标本号 = Trim(Split(strLine, ",")(3))
                    str标本号 = Trim(Mid(str标本号, 1, Len(str标本号) - 1))
                ElseIf InStr(strLine, "Dry Chemical Data Start") > 0 Then
                    intStep = 1
                End If
                
                If intStep = 4 Then
                    If aResults(UBound(aResults)) <> "" Then
                        ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
                    aResults(UBound(aResults)) = StrTime & "|" & str标本号 & "|血液| |0" & strItems & vbCrLf & strImages
                End If
                
                DoEvents
            Loop
            Close #1    ' 关闭文件。
            strFile = Dir
        Loop

    Next

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "EH-2050B Plus", strFile, Err.Number & " " & Err.Description
End Function









