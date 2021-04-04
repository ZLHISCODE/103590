VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_QQ96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:QQ-96化学发光免疫分析仪
'型号:QQ-96
'厂商:石家庄康普生科技有限公司
'作者:李建飞
'用户:贵阳铜仁地区医院
'日期:2010-01-19
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
    'strFile：包含路径的文件名
    'lngSampleNo：标本号。为-1表示选取所有时间范围内的标本
    'dtStart：开始时间，如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    'dtEnd：结束时间。只有选取多个标本(lngSample=-1)时该参数有效。如果不指定该参数，则时间范围为>=dtStart
    '返回的检验结果数组（各解析程序必须按照以下标准组织结果）
    '元素之间以 | 分隔
    '第 0 个元素：检验时间
    '第 1 个元素：样本序号
    '第 2 个元素：检验人
    '第 3 个元素：标本
    '第 4 个元素：是否质控品
    '从第 5 个元素开始为检验结果，每两个元素表示一个检验项目
    '如：第 5i 个元素为检验项目，第 5i+1 个元素表示检验结果
    
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cnErs As New ADODB.Connection, strFileErs As String, rsErs As ADODB.Recordset, strImgErs As String
    Dim str_检验项目 As String
        
    strFilePath = objFile.GetParentFolderName(strFile)  '取路径
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
            .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
            .Open
            If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    ReDim aResults(0)
    strImgErs = ""
    
    If strFileErs <> "" Then
        strSQL = "select A.RQ As 检验时间 ,B.YBH As 样本序号,'血液' as 标本, C.ENG AS 检验项目, A.userCon as 检验结果 " & _
                 "From BL as A, PATIENT as B ,REAG as C " & _
                 "Where A.SampId = B.ID And C.ID = A.ReagId and Format(Cdate(A.RQ),'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-MM-dd") & "' " & _
                 "And Format(Cdate(A.RQ),'yyyy-mm-dd')<= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " and VAL(B.YBH)=" & lngSampleNO)
                 
        Set rsErs = cnErs.Execute(strSQL)
        
        While Not rsErs.EOF
            '根据文件结构读取检验信息
            If Val("" & rsErs!样本序号) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = Format(CDate(rsErs!检验时间), "yyyy-mm-dd") & "|" & Val("" & rsErs!样本序号) & "| |" & _
                                        rsErs!标本 & "|0|" & rsErs!检验项目 & "|" & rsErs!检验结果
                'WriteLog "解析 KPS-I", aResults(UBound(aResults)), ""
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults
End Function


