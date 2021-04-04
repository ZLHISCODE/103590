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
'����:LICA HT��ѧ������
'����:��������Ƽ�
'����:�µ���
'�û�:��������ҽԺ
'����:2012-10-25

Public gcnExcel As New ADODB.Connection

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strID As String
    Dim strsheet As String
    Dim strItem As String
    Dim str���� As String, str�걾�� As String, str��Ŀ As String
    Dim colmin As Integer, colmax As Integer, colNum As Integer
    Dim ObjExcel As Object

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
     Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        WriteLog "OpenExcelFile", strFile, "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�"
        Exit Function
    End If
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2: colNum = 2: colmax = 4
        ReDim aResults(0)
        Do While True
            If .Worksheets(1).Cells(lngRow, 1) <> "" Then
               If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                'Oct-24 11:46
                str���� = Trim(.Worksheets(1).Cells(lngRow, 9))
                strItem = Trim(Replace(Mid(str����, 1, InStr(str����, "-") - 1), Chr(&H9), ""))
                strItem = Format(Now, "yyyy") & "-" & TrunMonth(strItem) & Trim(Mid(str����, InStr(str����, "-"))) & ":" & Format(Now, "SS")
                str���� = Format(strItem, "yyyy-MM-dd HH:MM:SS")
                If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd HH:MM:SS")
                
                aResults(UBound(aResults)) = str����
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(.Worksheets(1).Cells(lngRow, 1))  '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "Ѫ��"                      '�걾
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '�ʿ�
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim(Replace(.Worksheets(1).Cells(lngRow, colNum), Chr(&H9), "")) '��Ŀ
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim(Replace(.Worksheets(1).Cells(lngRow, colmax), Chr(&H9), ""))   '���
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
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
    '���ܣ���Excel��ʽ�ļ�
    '��Σ�strFileName
    '���Σ�Sheet�б�����|�ָ�
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strsheet As String
    On Error GoTo errHandle
    OpenExcelFile = ""

    If gcnExcel.State = 1 Then     '��������ӹ�����رգ���ʼ���´�����
        gcnExcel.Close
    End If
    
    gcnExcel.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilename & ";" & _
                              "Extended Properties=Excel 8.0;" & _
                              "Persist Security Info=False"
    gcnExcel.Open
    Set BiaoMing = gcnExcel.OpenSchema(adSchemaColumns)    '�������ݿ��¼��
    
    TableName = "": strsheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '�г����б�
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
    '��һ��Sheet
    '���: Sheet��
    '����: ADO��¼��
    
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
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
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

