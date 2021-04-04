VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_PCR9800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:实时跟踪荧光基因扩增仪  Genelight9800
'厂商:Amplly 厦门安普利生物工程有限公司
'作者:陈东
'用户:晋煤集团总医院
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
    Dim strBh As String, strAnti As String, str结果 As String
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=amplly"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select Mid(A.patient_id,9) as 编号,A.report_Date as 日期,item_name as 项目,A.value_illno as 结果 "
        strSQL = strSQL & " From analysis A Where Format(A.report_Date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
        strSQL = strSQL & " And Format(A.report_Date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(mid(A.patient_id,9))=" & lngSampleNO)
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
                
                aResults(lngResultCount) = Format(rsTmp("日期"), "yyyy-mm-dd hh:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("编号")) '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "|血清"
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp!项目
                
                str结果 = CStr(rsTmp!结果)
                If str结果 = 0 Then
                    str结果 = "<1.00E+03"
                Else
                    str结果 = Format(Val(str结果) / 10 ^ (Len(str结果) - 1), "0.00") & "E+" & Format(Len(str结果) - 1, "00")
                End If
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & str结果
                strBh = rsTmp("编号")

            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
   
End Function
