VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_zybiioQ7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:zybiio-Q7免疫定量分析仪
'作者:李骏驰
'用户:大理白求恩医院

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
Dim XMLDoc As DOMDocument, XML指标 As DOMDocument, objFile As New Scripting.FileSystemObject
Dim strFilePath As String, strFileName As String
Dim xList  As IXMLDOMNodeList, yList As IXMLDOMNodeList
Dim aResults() As String, firstnode As IXMLDOMNode
Dim str标本号 As String, str检验结果 As String
Dim str日期 As String
Dim i As Long, j As Long
On Error GoTo errHandle

    Set XMLDoc = New DOMDocument
    Set XML指标 = New DOMDocument
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileName = Dir(strFilePath & "\*.xml")
    Do While strFileName <> ""
        If XMLDoc.Load(strFilePath & "\" & strFileName) Then
            str日期 = Replace(Replace(XMLDoc.selectNodes("Export/When").Item(0).Text, "/", " "), ".", ":")
            Set xList = XMLDoc.getElementsByTagName("Cassette")
            For i = 0 To xList.Length - 1
                For j = 0 To xList.Item(i).childNodes.Length - 1
                    If xList.Item(i).childNodes.Item(j).nodeName = "SampleId" Then
                        str标本号 = xList.Item(i).childNodes.Item(j).Text
                    End If
                    If xList.Item(i).childNodes.Item(j).nodeName = "Analyte" Then
                        str检验结果 = str检验结果 & "|" & xList.Item(i).childNodes.Item(j).childNodes(1).Text & "|" & xList.Item(i).childNodes.Item(j).childNodes(2).Text
                    End If
                Next j
                If Val(str标本号) > 0 And str检验结果 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血液|0" & str检验结果
                    str检验结果 = ""
                End If
            Next i
        End If
        If Not objFile.FolderExists(strFilePath & "\" & Format(str日期, "yyyy-mm-dd")) Then objFile.CreateFolder (strFilePath & "\" & Format(str日期, "yyyy-mm-dd"))
        objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\" & Format(str日期, "yyyy-mm-dd") & "\" & strFileName, True
        objFile.DeleteFile strFilePath & "\" & strFileName
        strFileName = Dir(strFilePath & "\*.xml")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_FILE_XN1000.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

