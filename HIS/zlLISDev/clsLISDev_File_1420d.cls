VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_1420d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:VICTOR2 D (1420d) 多标记分析仪(用于新生儿筛查)
'厂商:61-珀金埃尔默(PerkinElmer)
'作者:陈东
'用户:德阳人民医院

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

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strFilePath = objFile.GetParentFolderName(strFile)
    ReDim aResults(0)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=amplly"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select  Pat_Id as 标本号,Entry_Date as 日期,GEST ,""AGE@EDD"" as AGE,AFP,AFPMOM,HCG,HCGMOM,T18_RISK as T18,T21_RISK as T21,AGE_RISK,NTD "
        strSQL = strSQL & " From downs2t A Where A.entry_date >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
        strSQL = strSQL & " And A.entry_date< Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "')" & _
        IIf(lngSampleNO = -1, "", " And val(A.Pat_Id)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = Format(rsTmp("日期"), "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("标本号"))  '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血清"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|GEST|" & IIf(Trim("" & rsTmp!GEST) = "", " ", Trim("" & rsTmp!GEST))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AGE|" & IIf(Trim("" & rsTmp!AGE) = "", " ", Trim("" & rsTmp!AGE))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFP|" & IIf(Trim("" & rsTmp!AFP) = "", " ", Trim("" & rsTmp!AFP))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFPMOM|" & IIf(Trim("" & rsTmp!AFPMOM) = "", " ", Trim("" & rsTmp!AFPMOM))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HCG|" & IIf(Trim("" & rsTmp!HCG) = "", " ", Trim("" & rsTmp!HCG))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HCGMOM|" & IIf(Trim("" & rsTmp!HCGMOM) = "", " ", Trim("" & rsTmp!HCGMOM))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|T18|" & IIf(Trim("" & rsTmp!T18) = "", " ", Trim("" & rsTmp!T18))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|T21|" & IIf(Trim("" & rsTmp!T21) = "", " ", Trim("" & rsTmp!T21))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AGE_RISK|" & IIf(Trim("" & rsTmp!AGE_RISK) = "", " ", Trim("" & rsTmp!AGE_RISK))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|NTD|" & IIf(Trim("" & rsTmp!NTD) = "", " ", Trim("" & rsTmp!NTD))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
   
End Function

