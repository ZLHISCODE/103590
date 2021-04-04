VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_GFM3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'快速免疫分析仪
'仪器型号：GF-M3000
'山东高密
'用户：新绛县人民医院
'作者：郝立功
'日期：2014-12-02
'

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
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

    Dim str日期 As String, str检验项目 As String, str标本号 As String, strimg As String
    Dim strFilePath As String, strFilename As String
    Dim StrValue As String
    Dim str_Item As String, str_Value As String
    Dim lngS As Long, lngE As Long
    Dim j As Integer

    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = UCase(Dir(strFilePath & "\*"))
    ReDim aResults(0)

    Do While strFilename <> ""
        str日期 = "": str标本号 = "": str检验项目 = "": strimg = ""
        If UCase(strFilename) Like "RESULT.TXT" Then
            If objFile.FileExists(strFilePath & "\" & strFilename) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
                '                    j = 1
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine

                    If str日期 = "" Then

                        str日期 = Trim(Mid(Split(strLine, "测试日期")(1), 1, 10)) & " " & Trim(Mid(Split(strLine, "测试时间")(1), 1, 8))

                        str日期 = Format(str日期, "yyyy-mm-dd hh:mm:ss")

                    Else
                        str_Item = Split(strLine, ",")(0)

                        If Split(strLine, ",")(2) <> "PC" And Split(strLine, ",")(2) <> "NC" And Split(strLine, ",")(2) <> "QC" Then
                            str标本号 = Split(strLine, ",")(2)

                            str_Value = Split(strLine, ",")(7) & "^" & Split(strLine, ",")(4) & "^" & Split(strLine, ",")(5) & "^" & Split(strLine, ",")(6)

                            If str_Value = "" Then
                                str_Value = "0"
                            End If

                            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||" & str日期 & "|" & str标本号 & "|检验人|血液|0" & "|" & str_Item & "|" & str_Value
                        End If
                    End If

                Loop

                '                读了就把文件放到BAK下
                objStream.Close
                Set objStream = Nothing
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
                objFile.DeleteFile strFilePath & "\" & strFilename, True
            End If
        End If
        strFilename = Dir

    Loop
    aResults(UBound(aResults)) = Mid(aResults(UBound(aResults)), 3)
    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_GFM3000"
    Close #lngErrFileNo
End Function
