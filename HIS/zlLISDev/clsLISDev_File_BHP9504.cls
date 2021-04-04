VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BHP9504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BHP9504化学发光仪
'厂商:北京滨松光子技术股份有限公司
'作者:陈德荣
'用户:汉源县人民医院

Private FileNames As String
Dim FileDate As Date
Dim fso As New FileSystemObject
Dim fld As Folder

Private Function Trunc(ByVal aDate As Date) As Date
  Dim Tmp As String
  Tmp = Format(aDate, "YYYY-MM-DD")
  Trunc = CDate(Tmp)
End Function

Private Function FindFile(ByVal sFol As String, ByVal sFile As String, ByVal aDate As Date, ByVal bDate As Date) As Currency
   Dim tFld As Folder, tFil As File, Filename As String
    
   On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   Filename = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(Filename) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
      Filename))
      Set tFil = fso.GetFile(fld.Path & "\" & Filename)
      FileDate = tFil.DateLastModified
      If Trunc(FileDate) >= Trunc(aDate) And Trunc(FileDate) <= Trunc(bDate) Then
        If FileNames <> "" Then FileNames = FileNames + "|"
        FileNames = FileNames & fso.BuildPath("", Filename)
      End If
      Filename = Dir() ' Get next file
      DoEvents
   Wend
   If fld.SubFolders.count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.Path, sFile, aDate, bDate)
      Next
   End If
   Exit Function
Catch: Filename = ""
       Resume Next
End Function


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strFilePath As String, strLog As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim str_Return As String, str_Tmp As String
    Dim lngRow As Integer, varLine As Variant
    Dim ExportFile As String
    Dim aFiles() As String, intloop As Integer, UseMode As Integer
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream2 As Scripting.TextStream
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strIniFile = App.Path & "\BHP9504.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream2 = objFileSystem.CreateTextFile(strIniFile)
       objStream2.WriteLine "[Options]"
       objStream2.WriteLine "ReadMode=0 ;0=读取本日文件 1=读取选定日期范围内的文件"
       objStream2.Close
       Set objStream2 = Nothing
    End If
    UseMode = Val(ReadIni("Options", "ReadMode", strIniFile))
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If UseMode = 0 Then
       FileNames = Format(Now, "YYYY-MM-DD") & ".txt"
    ElseIf UseMode = 1 Then
       Call FindFile(strFilePath, "*.txt", dtStart, dtEnd)
       If FileNames = "" Then
          FileNames = Format(dtEnd, "YYYY-MM-DD") & ".txt"
       End If
    End If
    If Right(strFilePath, 1) <> "\" Then strFilePath = strFilePath & "\"
    aFiles = Split(FileNames, "|")
    For intloop = LBound(aFiles) To UBound(aFiles)
      Set objStream = objFile.OpenTextFile(strFilePath & aFiles(intloop), ForReading)
      
      Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varLine = Split(strLine, Chr(&HD) & Chr(&HA))
        
        For lngRow = LBound(varLine) To UBound(varLine)
           str_Return = varLine(lngRow)
           str_Tmp = Trim(Mid(str_Return, 1, 12))
           
           If IsNumeric(str_Tmp) Then
              str日期 = Trim(Mid(str_Return, 14, 20))
              str标本号 = Trim(Mid(str_Tmp, 9))
              strChenel = Trim(Mid(str_Return, 38, 12))
              str检验结果 = Trim(Mid(str_Return, 52))
              If Trim(strChenel) <> "" And Trim(str检验结果) <> "" Then
                 If aResults(UBound(aResults)) <> "" Then
                  ReDim Preserve aResults(UBound(aResults) + 1)
                 End If
              End If
              aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & strChenel & "|" & str检验结果
              strLog = strLog & str日期 & "|" & str标本号 & "| |血清|0|" & strChenel & "|" & str检验结果
           End If
        Next
        DoEvents
      Loop
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_BHP9504.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function

