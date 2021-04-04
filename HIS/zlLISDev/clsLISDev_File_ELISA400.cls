VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ELISA400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:elisa 400 ø����
'����:����ŵ̩
'����:������
'ʱ��:2015-01-15
'�û�:�⳵�صڶ�����ҽԺ
Public MdbPass As String



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
    Dim lngResultCount As Long
    On Error GoTo ErrH
    MdbPass = "20100909JVD"
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function

    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & MdbPass & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "select B.TEST_DATE As ��������,C.SAMPLENBR As �걾��,'ѪҺ' As �걾,A.METHOD_ID As ͨ����,C.INTERPRETATION As ������ from" & _
            " TB_METHODDEFINE A,TB_RESULT_PLATE B,TB_SAMPLEDATA C where A.METHOD_ID=B.METHOD_ID And B.PLATE_ID=C.PLATE_ID" & _
            " And CDate(Format(B.TEST_DATE,'YYYY-MM-DD'))>=CDate('" & Format(dtStart, "yyyy-MM-dd") & "') And CDate(Format(B.TEST_DATE,'YYYY-MM-DD'))<=CDate('" & Format(dtEnd, "yyyy-MM-dd") & "')" & _
            IIf(lngSampleNO = -1, "", " And val(C.SAMPLENBR)=" & lngSampleNO) & _
            " order by B.TEST_DATE,C.SAMPLENBR "
        Set rsTmp = cnAccess.Execute(strSQL)
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("��������"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(rsTmp("�걾��")) '�걾��
            aResults(lngResultCount) = aResults(lngResultCount) & "| "  '������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("�걾") '�걾
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '�ʿ�
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ͨ����") & "|" & IIf(IsNull(rsTmp("������")), "*.**", rsTmp("������"))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
ErrH:
    If Error = 1 Then
        Resume
    End If
End Function



