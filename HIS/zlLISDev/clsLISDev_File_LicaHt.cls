VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LicaHt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:LICA HT化学发光仪
'厂商:博阳生物科技
'作者:陈德荣
'用户:巴中市中医院
'日期:2012-10-25

Public gcnExcel As New ADODB.Connection

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
    Dim strsheet As String
    Dim strItem As String
    Dim str日期 As String, str标本号 As String, str项目 As String
    Dim colmin As Integer, colmax As Integer, colNum As Integer
    Dim ObjExcel As Object

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
     Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        WriteLog "OpenExcelFile", strFile, "EXCEL未正确安装，请正确安装EXCEL中文版后再运行！"
        Exit Function
    End If
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2: colNum = 2: colmax = 4
        ReDim aResults(0)
        Do While True
            If .Worksheets(1).Cells(lngRow, 1) <> "" Then
               If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                'Oct-24 11:46
                str日期 = Trim(.Worksheets(1).Cells(lngRow, 9))
                strItem = Trim(Replace(Mid(str日期, 1, InStr(str日期, "-") - 1), Chr(&H9), ""))
                strItem = Format(Now, "yyyy") & "-" & TrunMonth(strItem) & Trim(Mid(str日期, InStr(str日期, "-"))) & ":" & Format(Now, "SS")
                str日期 = Format(strItem, "yyyy-MM-dd HH:MM:SS")
                If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd HH:MM:SS")
                
                aResults(UBound(aResults)) = str日期
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(.Worksheets(1).Cells(lngRow, 1))  '样本序号
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '检验人
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血清"                      '标本
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '质控
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim(Replace(.Worksheets(1).Cells(lngRow, colNum), Chr(&H9), "")) '项目
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim(Replace(.Worksheets(1).Cells(lngRow, colmax), Chr(&H9), ""))   '结果
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '关闭EXCEL对象
    ObjExcel.quit
    Set ObjExcel = Nothing
    For lngRow = 0 To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LicaHt.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

Private Function OpenExcelFile(ByVal strFilename As String) As String
    '功能：打开Excel格式文件
    '入参：strFileName
    '出参：Sheet列表，以|分隔
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strsheet As String
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
    
    TableName = "": strsheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '列出所有表
            TableName = BiaoMing("table_name")
            strsheet = strsheet & "|" & TableName
        End If
        BiaoMing.MoveNext
    Loop
    
    Set BiaoMing = Nothing
    If strsheet <> "" Then
        OpenExcelFile = Mid(strsheet, 2)
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
    Dim strsheet As String
    On Error GoTo errHandle
    
    If strSheetName = "" Then Exit Function
    
    strsheet = strSheetName
    If Right(strsheet, 1) <> "$" Then
        strsheet = strsheet & "$"
    End If
    
    rsTmp.Open strSheetName, gcnExcel, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    If Not rsTmp.EOF Then
        Set OpenExcelSheet = rsTmp
    End If

    Exit Function
errHandle:
    WriteLog "OpenExcelSheet", strSheetName, Err.Number & " " & Err.Description
End Function
Public Function Read_Excel(ByVal sFile As String, ByVal strsheet As String) As ADODB.Recordset

      On Error GoTo fix_err
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim sconn As String

      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic

      sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
      rs.Open "SELECT * FROM [" & strsheet & "]", sconn
      Set Read_Excel = rs
      Exit Function
fix_err:
      WriteLog "OpenExcelSheet", strsheet, Err.Number & " " & Err.Description
End Function

Private Function TrunMonth(ByVal strMonth As String) As String
    '转换中文的月份为数字
    Select Case UCase(strMonth)
        Case "一月": TrunMonth = "01"
        Case "二月": TrunMonth = "02"
        Case "三月": TrunMonth = "03"
        Case "四月": TrunMonth = "04"
        Case "五月": TrunMonth = "05"
        Case "六月": TrunMonth = "06"
        Case "七月": TrunMonth = "07"
        Case "八月": TrunMonth = "08"
        Case "九月": TrunMonth = "09"
        Case "十月": TrunMonth = "10"
        Case "十一": TrunMonth = "11"
        Case "十二月": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function


