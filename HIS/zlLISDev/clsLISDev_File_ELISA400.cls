VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ELISA400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:elisa 400 酶免疫
'厂商:美国诺泰
'作者:罗云龙
'时间:2015-01-15
'用户:库车县第二人民医院
Public MdbPass As String



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
    Dim lngResultCount As Long
    On Error GoTo ErrH
    MdbPass = "20100909JVD"
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function

    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & MdbPass & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "select B.TEST_DATE As 检验日期,C.SAMPLENBR As 标本号,'血液' As 标本,A.METHOD_ID As 通道码,C.INTERPRETATION As 检验结果 from" & _
            " TB_METHODDEFINE A,TB_RESULT_PLATE B,TB_SAMPLEDATA C where A.METHOD_ID=B.METHOD_ID And B.PLATE_ID=C.PLATE_ID" & _
            " And CDate(Format(B.TEST_DATE,'YYYY-MM-DD'))>=CDate('" & Format(dtStart, "yyyy-MM-dd") & "') And CDate(Format(B.TEST_DATE,'YYYY-MM-DD'))<=CDate('" & Format(dtEnd, "yyyy-MM-dd") & "')" & _
            IIf(lngSampleNO = -1, "", " And val(C.SAMPLENBR)=" & lngSampleNO) & _
            " order by B.TEST_DATE,C.SAMPLENBR "
        Set rsTmp = cnAccess.Execute(strSQL)
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("检验日期"), "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(rsTmp("标本号")) '标本号
            aResults(lngResultCount) = aResults(lngResultCount) & "| "  '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("标本") '标本
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '质控
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("通道码") & "|" & IIf(IsNull(rsTmp("检验结果")), "*.**", rsTmp("检验结果"))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
ErrH:
    If Error = 1 Then
        Resume
    End If
End Function




