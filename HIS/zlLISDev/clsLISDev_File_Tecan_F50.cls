VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Tecan_F50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:全自动酶联免疫工作站（Tecan Freedom EVOlyzer）
'厂商:瑞士帝肯TECAN
'作者:李骏驰
'用户:大理州中医院
'型号:F50


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
    Dim strFilePath As String, strFilename As String
    Dim i As Integer
    Dim objStream As textStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim iRow As Integer '行号
    Dim varLine As Variant
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    str日期 = Format(dtStart, "yyyy-MM-dd")

    strFilename = Dir(strFilePath & "\*.asc")
    Do While strFilename <> ""
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varLine = Split(strLine, Chr(&H9))
            If InStr(1, Trim(varLine(0)), "mth") > 0 Then
                strChenel = Mid(Trim(varLine(0)), 1, InStr(1, Trim(varLine(0)), "mth") - 2)
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varLine = Split(strLine, Chr(&H9))
            
            If Trim(varLine(0)) <> "" And InStr(1, Trim(varLine(0)), "mth") = 0 Then
                '标本号，数据行
                str标本号 = Trim(varLine(3))
                '检验结果模式  str定性结果 & "^" & str定量结果 & "^" & strCO & "^0"
                 
                str检验结果 = Trim(Trim(varLine(2))) & "^" & Trim(Trim(varLine(0))) & "^" & Trim(Trim(varLine(4))) & "^" & Trim(Trim(varLine(1)))
                
                If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd")
                
                If Val(str标本号) > 0 And str检验结果 <> "" And strChenel <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血液|0|" & strChenel & "|" & str检验结果
                End If
            End If
            
        Loop
        objStream.Close
        Set objStream = Nothing
        
        If Not objFile.FolderExists(strFilePath & "\" & Replace(str日期, "-", "")) Then objFile.CreateFolder (strFilePath & "\" & Replace(str日期, "-", ""))
        objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\" & Replace(str日期, "-", "") & "\" & strFilename, True
        objFile.DeleteFile strFilePath & "\" & strFilename
        strFilename = Dir(strFilePath & "\*.asc")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_EVOlyzer.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function
