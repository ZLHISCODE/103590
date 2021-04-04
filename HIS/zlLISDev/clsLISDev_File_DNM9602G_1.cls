VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DNM9602G_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:酶标仪
'厂商:北京普朗
'作者:王振涛
'用户:高平市妇幼保健院
'日期:2013-06-13

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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBh As String, strAnti As String, str结果 As String, strNO As String
    Dim lngResultCount As Long, str_项目 As String

    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=508"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    strSQL = "select * from tjsjb as A "

    strSQL = strSQL & " Where  A.RQ >=Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
    strSQL = strSQL & " And  A.RQ < Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "') " & _
             IIf(lngSampleNO = -1, "", " And val(A.blh)=" & Format(dtStart, "yyyymmdd") & Format(lngSampleNO, "0000"))
    WriteLog "DNM9602G", strSQL, ""

    strBh = "***"
    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "DNM9602G", "无记录！", ""
        Exit Function
    End If
    Do Until rsTmp.EOF
        strNO = Mid(rsTmp("blh"), Len(rsTmp("blh")) - 3, 4)
        If strBh <> strNO Then  '新标本
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)

            aResults(UBound(aResults)) = Format(Trim("" & rsTmp("rq")) & " " & Trim("" & rsTmp("sj")), "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         Val(strNO)    '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         "血清"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            strBh = strNO
        End If

        If strBh = strNO Then
            str_项目 = rsTmp("xmmc") & "|" & rsTmp("jcjg")
        End If
        If str_项目 <> "" Then
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str_项目
            str_项目 = ""
        End If

        WriteLog "DNM9602G", aResults(UBound(aResults)), ""
        rsTmp.MoveNext
    Loop

    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "DNM9602G", CStr(Erl()) & ":" & Err.Description, ""
End Function








