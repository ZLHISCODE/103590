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
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function
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
    Dim strSheet As String
    Dim strItem As String
    Dim str���� As String, str�걾�� As String, str��Ŀ As String
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
        
            
            If Trim("" & rsTmp.Fields(6)) = "����ţ�" Then
                
                If Val(strID) > 0 And str���� <> "" And str��Ŀ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strID)  '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "Ѫ��"                      '�걾
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0" & str��Ŀ                    '�ʿ�
                End If
                strID = rsTmp.Fields(7)
                str��Ŀ = "": str���� = ""
            ElseIf InStr(strItem, "," & Trim("" & rsTmp.Fields(3)) & ",") > 0 Then
                If Val(strID) > 0 Then
                    str��Ŀ = str��Ŀ & "|" & Trim("" & rsTmp.Fields(3)) & "|" & Trim("" & rsTmp.Fields(4))
                End If
            ElseIf Trim("" & rsTmp.Fields(2)) = "�������ڣ�" Then
                str���� = Replace(Trim("" & rsTmp.Fields(3)), "/", "-")
                If Not IsDate(CDate(str����)) Then str���� = ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    
    If Val(strID) > 0 And str���� <> "" And str��Ŀ <> "" Then
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
        aResults(UBound(aResults)) = str����
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strID)  '�������
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '������
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "Ѫ��"                      '�걾
        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0" & str��Ŀ                    '�ʿ�
    End If
    ResultFromFile = aResults
End Function

Private Function OpenExcelFile(ByVal strFilename As String) As String
    '���ܣ���Excel��ʽ�ļ�
    '��Σ�strFileName
    '���Σ�Sheet�б�����|�ָ�
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strSheet As String
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
    
    TableName = "": strSheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '�г����б�
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
    '��һ��Sheet
    '���: Sheet��
    '����: ADO��¼��
    
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

