VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BIOFOSUN_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'用户：汉源县人民医院
'仪器型号：微生物BIOFOSUN-II
'作者：敬秦
'时间：2014/4/18
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
        Dim objSystemFile As New FileSystemObject, cnAccess As New ADODB.Connection, rsTmp As New ADODB.Recordset
        Dim textStream As textStream
        Dim strResults() As String
        Dim strSQL As String, lngRecourd As Integer, i As Integer
        Dim strFilePath As String, filename As String, strBth As String
        
        On Error GoTo h
        strFilePath = objSystemFile.GetParentFolderName(strFile)
        If Len(strFilePath) = 0 Then strFilePath = App.Path
        With cnAccess
            .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
            .Open
        End With
        If cnAccess.State = adStateClosed Then Exit Function
        strSQL = "select b.TestRecordNum,b.plateID,b.GroupNum,b.Concentration as 结果,b.SIR,b.ref_Value,b.AntibioticNum,a.enum,a.value as 鉴定药物, " & _
                        "c.testRecordNum , c.specimenNum, c.IDtype, c.DateTimeStamp " & _
                        "from (select * from tblEnumAntibiotic_copy where id>=109 and id<>161 and id <> 163) a,tblResultGroups_copy b, tblTestRecord c where " & _
                        "a.enum=b.AntibioticNum and c.testRecordNum=b.testRecordNum  " & vbNewLine
        strSQL = strSQL & " and format(c.DateTimeStamp,'yyyy-mm-dd') >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & vbNewLine
        strSQL = strSQL & " and format(c.DateTimeStamp,'yyyy-mm-dd') <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
        strSQL = strSQL & IIf(lngSampleNO = -1, "", " and c.specimenNum= " & lngSampleNO) & vbNewLine
        strSQL = strSQL & " order by b.TestRecordNum , b.plateID"
        Set rsTmp = cnAccess.Execute(strSQL)
        strBth = "***"
        lngRecourd = -1
        ReDim Preserve strResults(0)
        strResults(0) = ""
        Do Until rsTmp.EOF
            If strBth <> rsTmp("specimenNum").Value Then '产生新的标本号
                lngRecourd = lngRecourd + 1
                ReDim Preserve strResults(lngRecourd)
                strResults(lngRecourd) = Format(rsTmp("DateTimeStamp").Value, "yyyy-mm-dd hh:mm:ss") & "|" & rsTmp("specimenNum").Value & "| |" '标本号
                strResults(lngRecourd) = strResults(lngRecourd) & rsTmp("IDtype").Value & "|0|" '细菌
                strResults(lngRecourd) = strResults(lngRecourd) & rsTmp("鉴定药物").Value & "|" & rsTmp("sir").Value & "^1^" & rsTmp("结果").Value
                strBth = rsTmp("specimenNum").Value
            Else
                strResults(lngRecourd) = strResults(lngRecourd) & "|" & rsTmp("鉴定药物").Value & "|" & rsTmp("sir").Value & "^1^" & rsTmp("结果").Value
            End If
            rsTmp.MoveNext
        Loop
        For i = LBound(strResults) To UBound(strResults)
            Debug.Print strResults(i)
        Next
        rsTmp.Close
        cnAccess.Close
        ResultFromFile = strResults
        Exit Function
h:
    'MsgBox Err.Description, vbCritical, "Information"
    WriteLog "File_BIOFOSUN", CStr(Erl()) & ":" & Err.Description, ""
    Err.Clear
End Function


