VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FIle_DA7600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:DA7600 PCR��
'����:���ݴﰲ
'����:�콭
'�û�:�����еڶ�����ҽԺ
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo hErr '�򿪴�������
    '------------------------------------------------
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
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsNO As ADODB.Recordset, rsItem As ADODB.Recordset, rsImage As ADODB.Recordset
    Dim strBh As String, strAnti As String, str��� As String, strImage As String, str���� As String, str�걾��
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "SELECT Format(TestBill.M_TestDate,'yyyy-mm-dd') & '|' & TestBillItems.D_Code & '| |ѪҺ|0|' & TestBillItems.Item_Code & '|' &  TestBillItems.PatientItem_Result as result"
        strSQL = strSQL & " FROM TestBillItems INNER JOIN TestBill ON TestBill.M_Code = TestBillItems.M_Code"
        strSQL = strSQL & " Where  TestBill.M_TestDate between cdate('" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "') and cdate( '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') Order BY D_Code"
                      
        Set rsNO = cnAccess.Execute(strSQL)
        If rsNO.EOF Then
            .Close
            Exit Function
        End If
        
         ReDim aResults(0)
        lngResultCount = 0
        
        Do While (Not rsNO.EOF)
           
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = rsNO("result")

            rsNO.MoveNext
        Loop
    
    End With
    ResultFromFile = aResults
   
    '------------------------------------------------
    Exit Function
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LISErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_FIle_DA7600"
    Close #lngErrFileNo
End Function