VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SM_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPaInfo As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        If dtStart < Date Then
            strSQL = "select checkDate as 时间,customername as 标本号,'血清' as 标本,'0' as 是否质控,decision as 结果,kindsymbol as 通道码 " & _
                     " from platerecord " & _
                     " where checkDate >= #" & Format(dtStart, "yyyy/mm/dd") & "#" & _
                     " Order By customername ,checktime"
        Else
            strSQL = "select checkDate as 时间,customername as 标本号,'血清' as 标本,'0' as 是否质控,decision as 结果,kindsymbol as 通道码 " & _
                     " from platerecord " & _
                     " where checkDate = #" & Format(dtStart, "yyyy/mm/dd") & "#" & _
                     " Order By customername,checktime "
            
        End If
        
        Set rsTmp = cnAccess.Execute(strSQL)
        lngSampleID = 0: lngResultCount = -1
        Do Until rsTmp.EOF
            If lngSampleID <> Val(rsTmp("标本号")) Then
                '新建头信息
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = Format(rsTmp("时间"), "yyyy-mm-dd hh:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("标本号") '标本号
                aResults(lngResultCount) = aResults(lngResultCount) & "| "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("标本") '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '质控
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("通道码") & "|" & Replace(Replace(rsTmp("结果"), "阴性", "阴性（-）"), "阳性", "阳性（+）")
            Else
                '指标信息
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("通道码") & "|" & Replace(Replace(rsTmp("结果"), "阴性", "阴性（-）"), "阳性", "阳性（+）")
            End If
            lngSampleID = Val(rsTmp("标本号"))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


