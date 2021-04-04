VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_QL800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:QL800微量元素分析仪
'厂商:济南齐力医疗器械有限公司
'作者:陈德荣
'用户:彭州妇幼保健院

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
    Dim strBh As String, str结果 As String
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT 编号,检验日期,锌浓度,铁浓度,钙浓度,镁浓度,磷浓度,硒浓度,铜浓度,铅浓度,锰浓度,镉浓度 From 检验报告 "
        strSQL = strSQL & " Where Format(检验日期,'yyyy-MM-dd 00:00:00')>='" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "'"
        strSQL = strSQL & " And Format(检验日期,'yyyy-MM-dd 23:59:59')<='" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(编号)=" & lngSampleNO)
        strSQL = strSQL & vbNewLine & "order by 编号"
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If strBh <> rsTmp("编号") Then '新标本
                
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("检验日期"), "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("编号")) '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|Zn_" & "|" & Val(IIf(IsNull(rsTmp!锌浓度), " ", rsTmp!锌浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Fe_" & "|" & Val(IIf(IsNull(rsTmp!铁浓度), " ", rsTmp!铁浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Ca_" & "|" & Val(IIf(IsNull(rsTmp!钙浓度), " ", rsTmp!钙浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Mg_" & "|" & Val(IIf(IsNull(rsTmp!镁浓度), " ", rsTmp!镁浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|P_" & "|" & Val(IIf(IsNull(rsTmp!磷浓度), " ", rsTmp!磷浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Se_" & "|" & Val(IIf(IsNull(rsTmp!硒浓度), " ", rsTmp!硒浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Cu_" & "|" & Val(IIf(IsNull(rsTmp!铜浓度), " ", rsTmp!铜浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Pb_" & "|" & Val(IIf(IsNull(rsTmp!铅浓度), " ", rsTmp!铅浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Mn_" & "|" & Val(IIf(IsNull(rsTmp!锰浓度), " ", rsTmp!锰浓度))
                aResults(lngResultCount) = aResults(lngResultCount) & "|Cd_" & "|" & Val(IIf(IsNull(rsTmp!镉浓度), " ", rsTmp!镉浓度))
                strBh = rsTmp("编号")
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function




