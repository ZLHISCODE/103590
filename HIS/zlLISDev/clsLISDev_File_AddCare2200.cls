VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AddCare2200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Addcare2200
'厂商:烟台艾德康
'作者:陈德荣
'用户:德阳人民医院
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
    Dim strChenel As String, StrQc As String, str日期 As String
    Dim strFilePath As String, strLog As String
    Dim objStream As TextStream
    Dim strLine As String, strTmp As String
    Dim str标本号 As String, str检验结果 As String, QcSampleNo As String
    Dim lngRow As Integer, varLine As Variant, iPos As Integer
    Dim strOD As String             '原始结果(OD)
    Dim strCO As String             'OD/CUTOFF
    Dim strCutOff As String         '判断值(Cut Off)
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject

    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strIniFile = App.Path & "\AddCare2000.ini"

    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "QcSampleNo = 1000 ;质控标本号为 1000"
       objStream.Close
       Set objStream = Nothing
    End If
    QcSampleNo = ReadIni("Options", "QcSampleNo", strIniFile)
    QcSampleNo = Replace(QcSampleNo, ";质控标本号为 1000", "")
    QcSampleNo = Trim(QcSampleNo)

    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    lngRow = 0: StrQc = "": strTmp = ""
    Do Until objStream.AtEndOfStream
       strLine = objStream.ReadLine
       iPos = InStr(strLine, "检验项目:")
       If iPos > 0 Then
         strChenel = Trim(Mid(strLine, iPos + 5, 6))
         strChenel = Replace(strChenel, "试", "")
         strChenel = Replace(strChenel, "验", "")
       End If
       iPos = InStr(strLine, "检测时间:")
       If iPos > 0 Then
         str日期 = Trim(Mid(strLine, iPos + 5, 20))
       End If
       If lngRow >= 5 Then
         If Trim(strChenel) <> "" And Trim(str日期) <> "" Then
             If aResults(UBound(aResults)) <> "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
             End If
             str标本号 = Trim(Mid(strLine, 1, 8))
             If (str标本号 = "QC") Or Val(str标本号) > 0 Then
                If str标本号 = "QC" Then
                  StrQc = "1"
                  str标本号 = QcSampleNo + Val(Replace(str标本号, "QC", ""))
                Else
                  StrQc = "0"
                End If
                strTmp = Trim(Mid(strLine, 28, 8))
                strOD = Trim(Mid(strLine, 37, 8))
                strCutOff = Trim(Mid(strLine, 67, 8))
                strCO = Trim(Mid(strLine, 57, 8))
                If StrQc = "1" Then
                  str检验结果 = strCO & "^" & strOD & "^" & strCutOff & "^" & strCO
                Else
                  str检验结果 = strTmp & "^" & strOD & "^" & strCutOff & "^" & strCO
                End If
                
                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|" & StrQc & "|" & IIf(StrQc = "1", "q" & strChenel, strChenel) & "|" & str检验结果
             End If
        End If
       End If
       lngRow = lngRow + 1
       DoEvents
    Loop
 
    For lngRow = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_AddCare2000.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function





