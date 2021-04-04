VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ADC_CLIA_400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:安图全自动化学发光免疫分析仪ADC CLIA 400
'厂商:安图实验仪器（郑州）有限公司
'MDB文件提取
'用户：防城中医院
'集成时间：2014年2月14
Public PassWord As String

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

Public Function GetAnswerCmd() As String
    '定时发送的指令
    GetAnswerCmd = ""
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPaInfo As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim strData As String
    
    On Error GoTo errHandle
    '数据库密码
    PassWord = "20100909JVD"
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & PassWord & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        '获取通道码及结果
        strSQL = "SELECT A.SAMPLE_ID as 标本号, B.METHOD_ID as 通道号, A.CALC_DATA as 结果, C.TEST_DATE as 日期 " & _
                 " FROM TB_SAMPLEDATA A,TB_METHODDEFINE B,TB_RESULT_PLATE C where A.PLATE_ID=C.PLATE_ID And C.METHOD_ID=B.METHOD_ID AND Format(C.TEST_DATE,'yyyy-MM-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' And Format(C.TEST_DATE,'yyyy-MM-dd') <= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                  IIf(lngSampleNO = -1, "", " And A.SAMPLE_ID =" & lngSampleNO) & _
                  " order by A.SAMPLE_ID,C.TEST_DATE"
                  

        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        lngResultCount = -1
        
        
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("日期"), "yyyy-mm-dd") '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(rsTmp("标本号")) '标本号
            aResults(lngResultCount) = aResults(lngResultCount) & "| "  '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|血清"  '标本
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '质控
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("通道号") & "|" & Format(Val(IIf(IsNull(rsTmp("结果")), "0", rsTmp("结果"))), "###0.000") '结果
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
errHandle:
    
    WriteLog "clsLISDev_File_ADC_CLIA_400.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

