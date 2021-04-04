VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ATB_1525"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ATB_1525微生物分析仪
'厂商:
'作者:王冬建
'用户:寿县县医院

Public Function GetStartCmd() As String
    '第一次启动时调用
End Function

Public Function GetEndCmd() As String
    '接口关闭的时候调用
End Function

Public Function GetAnswerCmd() As String
    '定时发送的指令
End Function

Public Function ResultFromFile(ByVal strfile As String, ByVal lngSampleNO As Long, _
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
    Dim strBh As String, strAnti As String
    Dim lngResultCount As Long
    strFilePath = objFile.GetParentFolderName(strfile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    On Error GoTo hErr
    
    With cnAccess
        .ConnectionString = "DBQ=" & strfile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=scsl"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select a.reqNo as 关联id,a.samNo as 标本号,b.ExeDate as 报告时间,b.GermId as 细菌 From ResultBill a,ResultExe b Where Format(b.ExeDate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
        strSQL = strSQL & "And a.reqNo = b.reqNo" & " And Format(b.ExeDate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(a.samNo)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If strBh <> rsTmp("关联id") Then '新标本
                
                If Nvl(rsTmp("细菌")) = "" Then GoTo NextDo:
                
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("报告时间"), "yyyy-mm-dd hh:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("标本号")) '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "| "
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    IIf(IsNull(rsTmp("细菌")), " ", rsTmp("细菌")) '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                    
                strBh = rsTmp("关联id")
                strAnti = ""
                
                strSQL = "Select antiId,micExplain as 药敏结果,iif(testNo='K-B','3',iif(testNo='DISK','2',iif(testNo='MIC','1'))) as 方法,resShow From ResultMic Where reqNo=" & strBh
                Set rsAnti = cnAccess.Execute(strSQL)
                Do Until rsAnti.EOF
                    strAnti = strAnti & "|" & rsAnti!antiId & "KSS|" & rsAnti!药敏结果 & "^" & rsAnti!方法 & "^" & rsAnti!resShow
                    rsAnti.MoveNext
                Loop
                aResults(lngResultCount) = aResults(lngResultCount) & strAnti
            End If
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    '解析完成，关闭文件
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults
    Exit Function
hErr:
    WriteLog "clsLisDev_ATB_1525", Err.Description, ""
End Function


