VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_File_Hd2001A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public gcnExcel As New ADODB.Connection
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
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strID As String
    Dim strSheet As String
    Dim strItem As String
    Dim str日期 As String, str标本号 As String, str项目 As String
    strItem = ",CA19-9,NSE,CEA,CA242,Ferritin,Beta-HCG,AFP,Free-PSA,PSA,CA125,HGH,CA15-3,"
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    strSheet = OpenExcelFile(strFile)
    If strSheet = "" Then Exit Function
    
    With gcnExcel
        If .State = adStateClosed Then Exit Function
        Set rsTmp = OpenExcelSheet(Split(strSheet, "|")(0))
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        strID = "***"
        Do While Not rsTmp.EOF
        
            
            If Trim("" & rsTmp.Fields(6)) = "检验号：" Then
                
                If Val(strID) > 0 And str日期 <> "" And str项目 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strID)  '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '检验人
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血清"                      '标本
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0" & str项目                    '质控
                End If
                strID = rsTmp.Fields(7)
                str项目 = "": str日期 = ""
            ElseIf InStr(strItem, "," & Trim("" & rsTmp.Fields(3)) & ",") > 0 Then
                If Val(strID) > 0 Then
                    str项目 = str项目 & "|" & Trim("" & rsTmp.Fields(3)) & "|" & Trim("" & rsTmp.Fields(4))
                End If
            ElseIf Trim("" & rsTmp.Fields(2)) = "检验日期：" Then
                str日期 = Replace(Trim("" & rsTmp.Fields(3)), "/", "-")
                If Not IsDate(CDate(str日期)) Then str日期 = ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    
    If Val(strID) > 0 And str日期 <> "" And str项目 <> "" Then
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
        aResults(UBound(aResults)) = str日期
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strID)  '样本序号
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '检验人
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血清"                      '标本
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0" & str项目                    '质控
    End If
    ResultFromFile = aResults
End Function

Private Function OpenExcelFile(ByVal strFilename As String) As String
    '功能：打开Excel格式文件
    '入参：strFileName
    '出参：Sheet列表，以|分隔
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strSheet As String
    On Error GoTo errHandle
    OpenExcelFile = ""

    If gcnExcel.State = 1 Then     '如果以连接过，则关闭，初始化下次事务
        gcnExcel.Close
    End If
    
    gcnExcel.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilename & ";" & _
                              "Extended Properties=Excel 8.0;" & _
                              "Persist Security Info=False"
    gcnExcel.Open
    Set BiaoMing = gcnExcel.OpenSchema(adSchemaColumns)    '创建数据库记录集
    
    TableName = "": strSheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '列出所有表
            TableName = BiaoMing("table_name")
            strSheet = strSheet & "|" & TableName
        End If
        BiaoMing.MoveNext
    Loop
    
    Set BiaoMing = Nothing
    If strSheet <> "" Then
        OpenExcelFile = Mid(strSheet, 2)
    End If
    Exit Function
errHandle:
    OpenExcelFile = ""
    WriteLog "OpenExcelFile", strFilename, Err.Number & " " & Err.Description
    
End Function

Private Function OpenExcelSheet(ByVal strSheetName As String) As ADODB.Recordset
    '打开一个Sheet
    '入参: Sheet名
    '出参: ADO记录集
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSheet As String
    On Error GoTo errHandle
    
    If strSheetName = "" Then Exit Function
    
    strSheet = strSheetName
    If Right(strSheet, 1) <> "$" Then
        strSheet = strSheet & "$"
    End If
    
    rsTmp.Open strSheetName, gcnExcel, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    If Not rsTmp.EOF Then
        Set OpenExcelSheet = rsTmp
    End If

    Exit Function
errHandle:
    WriteLog "OpenExcelSheet", strSheetName, Err.Number & " " & Err.Description
End Function


