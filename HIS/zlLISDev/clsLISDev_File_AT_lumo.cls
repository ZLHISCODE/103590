VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AT_lumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:酶标仪
'厂商:安图
'作者:
'用户:周口用膳妇科
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

    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.textStream, strLine As String
    Dim str日期 As String, str检验项目 As String, str定性结果 As String, str定量结果 As String
    Dim blnOk As Boolean '是否按标本号提取数据
    Dim strFilePath As String
    Dim strCO As String
    Dim ii As Long
    Dim varcol As Variant
    Dim str_bz As String
    Dim strline1 As String

    strFilePath = UCase(objFile.GetParentFolderName(strFile))

    If Dir(strFile) = "" Then Exit Function
    strFile = UCase(strFile)

    '
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    ii = 0
    str_bz = "End"
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If Trim(strLine) = "" Then
            str_bz = "Begin"
            ii = 1
        ElseIf Trim(strLine) = "" Then
            ii = 0
            str_bz = "End"
        End If
        If str_bz = "Begin" And ii <= 5 Then
            strline1 = strline1 & "|" & strLine
            ii = ii + 1
        End If

        If str_bz = "End" And Len(Trim(strline1)) > 10 Then
            varcol = Split(strline1, "|")
            If Not Nvl(aResults(UBound(aResults))) = "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
            End If
            str日期 = Split(varcol(2), ",")(0)
            If IsDate(str日期) Then
                str日期 = Format(str日期, "yyyy-MM-dd HH:mm:ss")
            Else
                str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
            lngSampleNO = Trim(varcol(2))
            str检验项目 = Trim(varcol(5))
            str定性结果 = Trim(varcol(6))
            str定量结果 = Trim(varcol(9))
            strCO = Format(Trim(varcol(10)), "0.000")
            strline1 = ""
            '            aResults(UBound(aResults)) = str日期 & "|" & lngSampleNO & "| |血液|0|" & str检验项目 & "|" & str定性结果 & "^" & str定量结果 & "^" & strCO & "^0"
            aResults(UBound(aResults)) = str日期 & "|" & lngSampleNO & "| |血液|0|" & str检验项目 & "|" & str定性结果 & "^" & str定量结果 & "^" & strCO & "^0"

        End If
    Loop

    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Function
