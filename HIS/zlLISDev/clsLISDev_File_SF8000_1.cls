VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SF8000_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:SF-8000
'����:��������ϣ�¿Ƽ���չ���޹�˾ SFϵ��ȫ�Զ���Ѫ����
'����:����
'�û�:������ҽԺ

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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim intLoop As Integer

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False
        strSQL = "SELECT ���,max(�걾��1) as �걾��,max(��������1) as ��������,MIN(����ʱ��1) as ����ʱ��,max(����ҽʦ1) as ����ҽʦ,max(�걾����1) as �걾����,max(PT_sec1) as PT_sec,max(PT_INR1) as PT_INR ," _
                 & "max(PT_11) as PT_1 ,max(PT_PTR1) as PT_PTR ,max(APTT1) as APTT,max(FIB1) as FIB ,max(TT1) as TT FROM(select FORMAT(A.��������,'YYYYMMDD')+FORMAT(A.�걾��,'000') as ���,A.�걾�� as �걾��1,A.�������� as ��������1," _
                 & "format(A.����ʱ��,'HH:NN:SS') as ����ʱ��1,B.����ҽʦ AS ����ҽʦ1,B.�걾���� AS �걾����1," _
                 & "iif(A.��Ŀ����='PT' and A.��λ='S',A.���Խ��,'') as PT_sec1," _
                 & "iif(A.��Ŀ����='PT' and A.��λ='INR',A.���Խ��,'') as PT_INR1," _
                 & "iif(A.��Ŀ����='PT' and A.��λ='%',A.���Խ��,'') as PT_11," _
                 & "iif(A.��Ŀ����='PT' and A.��λ='PTR',A.���Խ��,'') as PT_PTR1," _
                 & "iif(A.��Ŀ����='APTT' ,A.���Խ��,'') as APTT1," _
                 & "iif(A.��Ŀ����='FIB' and A.��λ='g/L',A.���Խ��,'') as FIB1," _
                 & "iif(A.��Ŀ����='TT' ,A.���Խ��,'') as TT1 " _
                 & "from TB_TEST_RESULT A,TB_TEST_DATA B WHERE  " & _
                 IIf(lngSampleNO = -1, "  A.�������� >=#" & Format(dtStart, "yyyy-mm-dd") & "#" & _
                 " And A.�������� <=#" & Format(dtEnd, "yyyy-mm-dd") & "#", _
                 "  A.�������� =#" & Format(dtStart, "yyyy-mm-dd") & "# and  A.�걾�� =" & lngSampleNO) & " and A.�걾�� = B.�걾�� And A.�������� = B.��������) GROUP BY ���   ORDER BY ���+0"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("��������"), "YYYY-MM-DD ") & Format(rsTmp("����ʱ��"), "HH:NN:SS")
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("�걾��") & "|" & IIf(IsNull(rsTmp("����ҽʦ")), " ", rsTmp("����ҽʦ")) & "|" & IIf(IsNull(rsTmp("�걾����")), "Ѫ��", rsTmp("�걾����")) & "|0"
            aResults(lngResultCount) = aResults(lngResultCount) & "|PT-sec|" & Nvl(rsTmp("PT_sec"), " ") & "|PT-INR|" & Nvl(rsTmp("PT_INR"), " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|PT-%|" & Nvl(rsTmp("PT_1"), " ") & "|PT_PTR|" & Nvl(rsTmp("PT_PTR"), " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|APTT|" & Nvl(rsTmp("APTT"), " ") & "|FIB|" & Nvl(rsTmp("FIB"), " ") & "|TT|" & Nvl(rsTmp("TT"), " ")
            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_SF8000", strFile, Err.Number & " " & Err.Description
End Function