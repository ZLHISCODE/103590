VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_KOFA_AM300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:KOFA AM300 生化分析仪
'厂商:
'作者:欧海胜
'用户:东莞莞华医院

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

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strDate As String, strNo As String
    Dim strItem As String, strResult As String, strValue As String
    On Error GoTo errH

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path


    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    If Format(dtStart, "yyyy-mm-dd") = Format(Now, "yyyy-mm-dd") Then
        strSQL = "SELECT  a.SampleID as 标本号,b.ItemID as 通道码,a.ID,a.TestDate,b.ItemResult as 结果,b.TestDate as 测试时间 " & _
                 "FROM SampleBaseInfor a, SampleTestData b " & _
                 "where a.ID=b.SampleID and Format(a.TestDate, 'yyyyMMdd')>=" & Format(dtStart, "yyyyMMdd") & _
               " and Format(a.TestDate, 'yyyyMMdd') <=" & Format(dtEnd, "yyyyMMdd") & _
                 IIf(lngSampleNO = -1, "", " And val(a.SampleID)=" & lngSampleNO) & _
               "  ORDER BY CStr(a.TestDate & '  ' & a.SampleID) "
    Else
        strSQL = "SELECT  a.SampleID as 标本号,b.ItemID as 通道码,a.ID,a.TestDate,b.ItemResult as 结果,b.TestDate as 测试时间 " & _
                 "FROM SampleBaseInforHistory a, SampleTestDataHistory b " & _
                 "where a.ID=b.SampleID and Format(a.TestDate, 'yyyyMMdd')>=" & Format(dtStart, "yyyyMMdd") & _
               " and Format(a.TestDate, 'yyyyMMdd') <=" & Format(dtEnd, "yyyyMMdd") & _
                 IIf(lngSampleNO = -1, "", " And val(a.SampleID)=" & lngSampleNO) & _
               "  ORDER BY CStr(a.TestDate & '  ' & a.SampleID) "
    End If


    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    strDate = "": strNo = ""
    strResult = ""
    Do Until rsTmp.EOF
        If strDate <> Trim("" & rsTmp.Fields("测试时间")) Or strNo <> Trim("" & rsTmp.Fields("标本号")) Then
            If strResult <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = strResult

            End If
            strDate = Trim("" & rsTmp.Fields("测试时间"))
            strNo = Trim("" & rsTmp.Fields("标本号"))
            strResult = strDate & "|" & strNo & "| |血清|0"


        End If
        strItem = Trim("" & rsTmp.Fields("通道码"))
        strValue = Trim("" & rsTmp.Fields("结果"))

        If strItem <> "" And strValue <> "" Then
            strResult = strResult & "|" & strItem & "|" & strValue
        End If

        rsTmp.MoveNext
    Loop
    If strResult <> "" Then
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
        aResults(UBound(aResults)) = strResult
    End If
    ResultFromFile = aResults
    Exit Function
errH:

End Function




