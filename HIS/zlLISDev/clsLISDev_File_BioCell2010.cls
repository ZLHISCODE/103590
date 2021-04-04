VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BioCell2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:博赛2010酶标仪
'厂商:郑州博赛生物工程有限责任公司
'作者:陈德荣
'用户:德阳第五医院

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
    Dim strFilePath As String, strLog As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim ItemName As Integer, ItemResult As Integer, ItemOD As Integer, ItemSCO As Integer '从INI读取通道号和项目结果的位置
    Dim lngRow As Long, varline As Variant
    Dim ExportFile As String, OD As String, SCO As String
    Dim aFiles() As String, intloop As Integer, UseMode As Integer, ResultMode As Integer, CheckDate As Integer
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strIniFile = App.Path & "\BioCell2010.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "ItemName=12       ;项目名称所在列数"
       objStream.WriteLine "ItemResult=15     ;项目值所在列数"
       objStream.WriteLine "ItemOD=14         ;OD值所在列数"
       objStream.WriteLine "ItemSCO=16        ;SCO值所在列数"
       objStream.WriteLine "CheckDate=0       ;检验日期所在列数"
       objStream.WriteLine "ReadMode=1        ;读取文件方式 0=只读取本日文件 1=读取指定日期范围文件"
       objStream.WriteLine "ResultMode=0      ;结果处理方式 0=不处理 1=特殊处理"
       objStream.Close
       Set objStream = Nothing
    End If
    
    ItemName = Val(ReadIni("Options", "ItemName", strIniFile))
    ItemResult = Val(ReadIni("Options", "ItemResult", strIniFile))
    CheckDate = Val(ReadIni("Options", "CheckDate", strIniFile))
    UseMode = Val(ReadIni("Options", "ReadMode", strIniFile))
    ResultMode = Val(ReadIni("Options", "ResultMode", strIniFile))
    If ItemName = 0 Then ItemName = 12
    If ItemResult = 0 Then ItemResult = 15
    If ItemOD = 0 Then ItemOD = 14
    If ItemSCO = 0 Then ItemSCO = 16
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If UseMode = 0 Then
       FileNames = Format(Now, "YYMMDD") & "dat.txt"
    ElseIf UseMode = 1 Then
       Call FindFile(strFilePath, "*.txt", dtStart, dtEnd)
       If FileNames = "" Then
          FileNames = Format(dtEnd, "YYMMDD") & "dat.txt"
       End If
    End If
    If Right(strFilePath, 1) <> "\" Then strFilePath = strFilePath & "\"
    aFiles = Split(FileNames, "|")
    For intloop = LBound(aFiles) To UBound(aFiles)
      Set objStream = objFile.OpenTextFile(strFilePath & aFiles(intloop), ForReading)
      
      Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varline = Split(strLine, Chr(&H9))
        If UBound(varline) >= ItemResult Then
           If IsNumeric(varline(0)) Then
              If (Len(varline(0)) = 10) Or (Len(varline(0)) = 11) Or (Len(varline(0)) = 9) Then
                 If CheckDate = 0 Then
                     str日期 = "20" & Mid(varline(0), 1, 2) & "-" & Mid(varline(0), 3, 2) & "-" & Mid(varline(0), 5, 2) & " " & Format(Now, "HH:MM:SS")
                 Else
                     str日期 = Trim(varline(CheckDate))
                 End If
                 str标本号 = Mid(varline(0), 7)
                 If Trim(varline(ItemName)) <> "" And Trim(varline(ItemResult)) <> "" Then
                    If aResults(UBound(aResults)) <> "" Then
                      ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
                    strChenel = Trim(varline(ItemName))
                    strChenel = Trim(Replace(strChenel, "梅毒抗体", "TP"))
                    str检验结果 = Trim(varline(ItemResult))
                    OD = Format(Trim(varline(ItemOD)), "#0.000#")
                    SCO = Format(Trim(varline(ItemSCO)), "#0.000#")
'                    If (ResultMode = 1) Then
'                       If (strChenel = "HBsAg") And (CCur(str检验结果) < 0.1) Then
'                          str检验结果 = "<0.10"
'                       ElseIf (strChenel = "Anti-HBs") And (CCur(str检验结果) < 5) Then
'                          str检验结果 = "<5.00"
'                       ElseIf (strChenel = "HbeAg") And (CCur(str检验结果) < 0.25) Then
'                          str检验结果 = "<0.25"
'                       ElseIf (strChenel = "Anti-HBe") And (CCur(str检验结果) < 1) Then
'                          str检验结果 = "<1.00"
'                       ElseIf (strChenel = "Anti-HBc") And (CCur(str检验结果) < 2) Then
'                          str检验结果 = "<2.00"
'                       End If
'                    End If
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & strChenel & "|" & str检验结果 & "^" & OD & "^0.00^" & SCO
                End If
              End If
           End If
        End If
        DoEvents
      Loop
    Next
    For lngRow = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_BioCell2010.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function


