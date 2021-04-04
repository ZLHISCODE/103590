VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BioChip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��������о���Ķ���
'����:����
'����:lwf
'�û�:������ҽԺ

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '�򿪴�������
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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer

    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String
    Dim str�걾�� As String
    Dim strChar As String
    Dim intMaxValue As Integer

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    If Dir(strFile) = "" Then Exit Function

    With cnAccess

        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open

        If .State = adStateClosed Then Exit Function
        strSQL = "select A.ID�� as sampleid,b.������Ŀ as abrvname,b.������ as itemresult,A.���ʱ�� as testdate from �����Ϣ A,����� B Where A.����� = B.�����" & _
                 " and Format(A.���ʱ��,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(A.���ʱ��,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And val(a.sampleid)=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If Len(Trim("" & rsTmp!abrvname)) > 0 Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                str���� = Trim("" & Format(rsTmp("TestDate"), "yyyy-mm-dd"))
                str�걾�� = rsTmp("SampleID")
                If Len(str�걾��) > 9 Then
                    str�걾�� = Mid(str�걾��, 9)
                End If
                aResults(lngResultCount) = str���� '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         Val("" & str�걾��)  '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         "Ѫ��"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                         ' Val(rsTmp("�ʿ�")) '�ʿ�
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Trim("" & rsTmp!abrvname) & "|" & Trim("" & rsTmp!ItemResult)
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Function