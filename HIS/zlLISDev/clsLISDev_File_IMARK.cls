VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_IMARK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'仪器:酶标仪 BIO_IMARK
'厂商:伯乐
'作者:王冬建
'用户:寿县县医院
Private mstrCode As String, str_Old As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
     

End Sub

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
    
    Dim objStream As textStream
    Dim strLine As String, varItem As Variant
    Dim str标本号 As String
    Dim lngn, Lngm As Long
    Dim FindDate As String
    On Error GoTo errHandle
    
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    lngn = CDate(Format(Now, "YYYY-MM-DD")) - dtStart
    For Lngm = 0 To lngn
        FindDate = dtStart + Lngm
        strFile = Mid(strFile, 1, InStrRev(strFile, "\")) + Format(FindDate, "yymmdd") + ".txt"
        If Dir(strFile) <> "" Then
                strFilePath = objFile.GetParentFolderName(strFile)
            If Len(strFilePath) = 0 Then strFilePath = App.Path
            Set objStream = objFile.OpenTextFile(strFile, ForReading)
            Do Until objStream.AtEndOfStream
                strLine = objStream.ReadLine
                varItem = Split(strLine, Chr(&H9))
                If strLine <> "" Then
                    str日期 = Format(FindDate, "yyyy-mm-dd")
                    str标本号 = varItem(1)
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    If str标本号 Like "QC*" Then '是否是质控标本
                        str标本号 = 997 + Mid(str标本号, InStr(str标本号, ":") + 1)
                        aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|1|" & varItem(2) & "|" & "" & "^" & Mid(varItem(3), 1, 5) & "^^"
                    ElseIf Val(str标本号) > 0 Then
                        aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & varItem(2) & "|" & varItem(4) & "^" & varItem(5) & "^" & varItem(6) & "^" & varItem(5) / varItem(6)
                    End If
                End If
            Loop
            ResultFromFile = aResults
            
            objStream.Close
        End If
    Next
    Exit Function
errHandle:
    WriteLog "clsLISDev_IMARK.ResultFromFile", strFile & " " & strLine, Err.Number & " " & Err.Description
End Function

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function
