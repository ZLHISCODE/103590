VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ARIS2X_Swin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:
'厂商:
'作者:敬秦
'用户:都江堰林业局医院
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperW" (ByVal lpsz As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private Function ReadUniFile(ByVal sFile As String) As String '处理unicode乱码的问题
    Dim a As Long
    a = FileLen(sFile)
    ReDim buff(a - 1) As Byte
    ReDim buff1(a - 3) As Byte
    Open sFile For Binary As #1
    Get #1, , buff
    Close #1
    CopyMemory buff1(0), buff(2), a - 2
    Dim s As String
    s = StrConv(buff1, vbNarrow)
    ReadUniFile = s
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
    Dim objFile As New FileSystemObject, strStream As textStream
    Dim strFileName As String, strFilePath As String
    Dim intDay As Integer, intLoop As Integer, i As Integer, j As Integer
    Dim strResult() As String, strDate As String, strItem As String, strGerm As String, strAnti As String, strValue As String
    Dim strLine As String, arrLine() As String, strSample As String, strFnc As String
    
    On Error GoTo ErrorHand
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    ReDim Preserve strResult(0)
    strResult(0) = ""
    j = 2
    intDay = Abs(DateDiff("d", dtStart, dtEnd))
    For intLoop = 0 To intDay
        strFileName = Dir(strFilePath & "\*_" & Format(dtStart + intLoop, "yyyymmdd") & "_*.txt")
        Do While strFileName <> ""
            strLine = ReadUniFile(strFilePath & "\" & strFileName)
            If strLine <> "" Then
                arrLine = Split(strLine, Chr(&H9))
                If UBound(arrLine) > 0 Then
                    For i = 0 To UBound(arrLine)
                        If i = 1 Then
                            strSample = arrLine(i)
                        End If
                        If i = 9 Then
                            strGerm = arrLine(i)
                        End If
                        If i = 39 Then
                            strDate = Format(arrLine(i), "yyyy-mm-dd hh:mm:ss")
                            If Not IsDate(strDate) Then
                                strDate = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            End If
                        End If
                        If i >= 40 Then
                            If j = 2 Then '循环插入数据
                                strAnti = arrLine(i)
                                If UCase(arrLine(i + 2)) = "SUSC" Then
                                    strItem = strItem & "|" & strAnti & "|" & "S^1^" & Replace(arrLine(i + 1), " ", "")
                                ElseIf UCase(arrLine(i + 2)) = "RESIST" Or UCase(arrLine(i + 2)) = "NOINTP" Then
                                    strItem = strItem & "|" & strAnti & "|" & "R^1^" & Replace(arrLine(i + 1), " ", "")
                                ElseIf UCase(arrLine(i + 2)) = "INTER" Then
                                    strItem = strItem & "|" & strAnti & "|" & "I^1^" & Replace(arrLine(i + 1), " ", "")
                                End If
                                j = -1
                            End If
                            j = j + 1
                        End If
'                        Call WriteLisLog1(arrLine(i))
                    Next
                End If
            End If
            If strItem <> "" And strGerm <> "" Then
                If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                strResult(UBound(strResult)) = strDate & "|" & strSample & "| |" & strGerm & "|0" & strItem
                strItem = ""
            End If
            strFileName = Dir   '查找下一个文件
        Loop
    Next
    ResultFromFile = strResult
    Exit Function
ErrorHand:
    Call WriteLog("调试日志", "", Err.Description)
    Err.Clear
End Function


