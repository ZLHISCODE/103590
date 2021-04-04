VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:原子吸收分光光度计
'厂商:邦达诚科技(北京)有限公司
'作者:陈东
'用户:南桐总医院

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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str标本号 As String, str项目 As String
    
    Dim rsErs As ADODB.Recordset
    Dim str日期 As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT a.SampleDate as 日期, PatientID as 标本号,CaseID" & _
                " FROM PatientSend as A Where A.SampleDate >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & _
                "') And A.SampleDate<= Cdate('" & Format(DateAdd("d", 1, dtEnd), "yyyy-MM-dd") & "') " & _
                IIf(lngSampleNO = -1, "", " And val(PatientID)=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        
        Do While Not rsTmp.EOF
                
                str日期 = Format("" & rsTmp("日期"), "yyyy-MM-dd HH:mm:ss")
                str标本号 = Val("" & rsTmp("标本号"))
                
                strSQL = "Select Element as 通道码,Content as 结果  From PatientResult as A Where A.CaseID=" & rsTmp("CaseID")
                Set rsErs = cnAccess.Execute(strSQL)
                str项目 = ""
                Do Until rsErs.EOF
                    str项目 = str项目 & "|" & rsErs("通道码") & "|" & rsErs("结果")
                    rsErs.MoveNext
                Loop
                If str项目 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("标本号"))       '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '检验人
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"                          '标本
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '质控
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str项目                               '项目和结果
                End If
                               
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function



'0769-26208501
