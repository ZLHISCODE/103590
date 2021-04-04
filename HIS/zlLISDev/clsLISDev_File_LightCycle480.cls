VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LightCycle480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:荧光定量PCR分析仪 LightCycle480
'厂商:罗氏 Roche
'作者:陈德荣
'用户:德阳市人民医院
'日期:2012-09-12

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
    Dim strChenel As String, str日期 As String, cTDH As String
    Dim strFilePath As String
    
    Dim objStream As TextStream
    Dim strLine As String, varLine As Variant
    Dim str标本号 As String, str检验结果 As String, lngRow As Integer
    
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
        If (Not strLine Like "*Experiment:*") And (Not strLine Like "*Include*Color*") Then
          varLine = Split(strLine, vbTab)
          If UBound(varLine) >= 5 Then
             strChenel = Trim(varLine(3))      'HC1120706188
             strChenel = Trim(Mid(strChenel, 3))
             If IsNumeric(strChenel) Then
                str标本号 = strChenel  '1120706188
                str日期 = "20" & Mid(strChenel, 2, 2) & "-" & Mid(strChenel, 4, 2) & "-" & Mid(strChenel, 6, 2) & " " & Format(Now, "HH:MM:SS")
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                str检验结果 = Trim(varLine(5))
                If IsDate(CDate(str日期)) And str标本号 <> "" And str检验结果 <> "" Then
                    cTDH = "HBV"
                    If InStr(varLine(3), "HB") Then
                       cTDH = "HBV"
                    ElseIf InStr(varLine(3), "HC") Then
                       cTDH = "HCV"
                    End If
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & cTDH & "|" & str检验结果
                End If
             End If
          End If
        End If
        DoEvents
    Loop
    
    For lngRow = 0 To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycle480.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function





