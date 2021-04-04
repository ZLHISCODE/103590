VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Lumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:LUmo化学发光仪
'厂商:62-AutoBio郑州安图生物工程有限公司
'作者:陈东
'用户:德阳医院
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
    Dim str日期 As String
    Dim strFilePath As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant
    Dim str标本号 As String
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varItem = Split(strLine, Chr(&H9))
        If UBound(varItem) >= 3 Then
            If IsNumeric(varItem(0)) And IsNumeric(varItem(2)) Then
                If Len(varItem(0)) = 11 Then
                    str日期 = "20" & Mid(varItem(0), 1, 2) & "-" & Mid(varItem(0), 3, 2) & "-" & Mid(varItem(0), 5, 2)
                    str标本号 = Mid(varItem(0), 7)
                    If Trim(varItem(1)) <> "" And Trim(varItem(2)) <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & varItem(1) & "|" & varItem(2)
                    End If
                End If
            End If
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_lumo.ResultFromFile", strFile & " " & strLine, Err.Number & " " & Err.Description
End Function

