VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DR6660"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:免疫  DR6660-2
'厂商:
'作者:王冬建
'用户:金寨县人民医院

Public Function GetStartCmd() As String
    '第一次启动时调用
End Function

Public Function GetEndCmd() As String
    '接口关闭的时候调用
End Function

Public Function GetAnswerCmd() As String
    '定时发送的指令
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
    Dim lngResultCount As Long
    On Error GoTo ErrH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select series_itemname as 项目,series_value as 结果,series_sampleID as 标本号,series_date as 日期 From series Where series_date >= '" & Format(dtStart, "yyyy-mm-dd") & "'"
        strSQL = strSQL & " And series_date <= '" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(Series_sampleID) = " & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1

        Do Until rsTmp.EOF
            
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsTmp("日期") & " " & "00:00:00", "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val(rsTmp("标本号")) '样本序号
            aResults(lngResultCount) = aResults(lngResultCount) & "| |血液|0|" & rsTmp("项目") & "|" & rsTmp("结果")
                    
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults
    Exit Function
ErrH:
    If Error = 1 Then
        Resume
    End If
End Function






