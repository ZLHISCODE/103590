VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_tailai_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���Զ�ʱ��ֱ�ӫ�����߷�����
'����:�����зỪ�����������ι�˾
'����:������
'�û�:˷��ƽ³������ҽԺ
'����:2013-06-13

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
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBh As String, strAnti As String, str��� As String, strNO As String
    Dim lngResultCount As Long, str_��Ŀ As String

    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    strSQL = "select C.C4,C.C6,A.A2,A.A18,A.A16 from h_patient as A , h_result as C"
    strSQL = strSQL & "  where A.A1=C.C2"
    strSQL = strSQL & "  And Format(mid(A.A16,1,4) & '-' & mid(a.a16,5,2) & '-' & mid(a.a16,7,2),'yyyy-mm-dd 00:00:00')>='" & Format(dtStart, "yyyy-mm-dd 00:00:00") & "'"
    strSQL = strSQL & "  And Format(mid(A.A16,1,4) & '-' & mid(a.a16,5,2) & '-' & mid(a.a16,7,2),'yyyy-mm-dd 00:00:00')<'" & Format(dtEnd + 1, "yyyy-mm-dd 23:59:59") & "'"
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " And A.A2='" & lngSampleNO & "'")

    WriteLog "̩��I", strSQL, ""

    strBh = "***"
    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "̩��I", "�޼�¼��", ""
        Exit Function
    End If
    Do Until rsTmp.EOF
        strNO = rsTmp("A2")
        If strBh <> strNO Then  '�±걾
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)

            aResults(UBound(aResults)) = Format(Mid(rsTmp("A16"), 1, 4) & "-" & Mid(rsTmp("a16"), 5, 2) & "-" & Mid(rsTmp("a16"), 7, 2), "yyyy-mm-dd") '����ʱ��
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         Val(strNO)    '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         "Ѫ��"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            strBh = strNO
        End If

        If strBh = strNO Then
            str_��Ŀ = rsTmp("C4") & "|" & rsTmp("C6")
        End If
        If str_��Ŀ <> "" Then
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str_��Ŀ
            str_��Ŀ = ""
        End If

        WriteLog "̩��I", aResults(UBound(aResults)), ""
        rsTmp.MoveNext
    Loop

    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "̩��I", CStr(Erl()) & ":" & Err.Description, ""
End Function







