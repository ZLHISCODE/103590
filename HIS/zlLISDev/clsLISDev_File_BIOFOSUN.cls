VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BIOFOSUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:细菌鉴定及药敏分析仪 BIOFOSUN
'厂商:
'作者:陈东
'用户:泽州县医院

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsSlav As ADODB.Recordset
    Dim i As Integer, strDate As String
    On Error GoTo errH
    
         
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "SELECT A.specimenNum,A.isolatenum,A.testDate,A.LIS_ID_ID" & _
             " FROM LIS_MAST as A Where A.testDate >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "#" & _
             " And A.TestDate<=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#"
    
    If lngSampleNo <> -1 Then
        strSQL = strSQL & " And  a.sepecimenNum = " & lngSampleNo
    End If
    WriteLog "File_BIOFOSUN", strSQL, ""
    
    ReDim aResults(0)
    
    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "File_BIOFOSUN", "无记录！", ""
        Exit Function
    End If
        
    Do Until rsTmp.EOF
         
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
        aResults(UBound(aResults)) = Format(rsTmp!testDate, "yyyy-MM-dd")  '检验时间
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("specimenNum")) '样本序号
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |" & rsTmp!LIS_ID_ID & "|0"
        strSQL = "SELECT A.Med_CODE,a.SIR,a.Concentration " & _
                 " FROM LIS_SLAVE as A where a.specimenNum = '" & rsTmp!specimenNum & "' and A.isolateNum='" & rsTmp!isolatenum & "'"
        strSQL = strSQL & " Order by A.med_group "
        
        Set rsSlav = cnAccess.Execute(strSQL)
        Do Until rsSlav.EOF
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsSlav!med_code & "|" & rsSlav!SIR & "^1^" & rsSlav!Concentration
            rsSlav.MoveNext
        Loop
        rsTmp.MoveNext
    Loop
        
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "File_BIOFOSUN", CStr(Erl()) & ":" & Err.Description, ""
End Function






