VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ADC_CLIA_400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��ͼȫ�Զ���ѧ�������߷�����ADC CLIA 400
'����:��ͼʵ��������֣�ݣ����޹�˾
'MDB�ļ���ȡ
'�û���������ҽԺ
'����ʱ�䣺2014��2��14
Public PassWord As String

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

Public Function GetAnswerCmd() As String
    '��ʱ���͵�ָ��
    GetAnswerCmd = ""
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPaInfo As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim strData As String
    
    On Error GoTo errHandle
    '���ݿ�����
    PassWord = "20100909JVD"
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & PassWord & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        '��ȡͨ���뼰���
        strSQL = "SELECT A.SAMPLE_ID as �걾��, B.METHOD_ID as ͨ����, A.CALC_DATA as ���, C.TEST_DATE as ���� " & _
                 " FROM TB_SAMPLEDATA A,TB_METHODDEFINE B,TB_RESULT_PLATE C where A.PLATE_ID=C.PLATE_ID And C.METHOD_ID=B.METHOD_ID AND Format(C.TEST_DATE,'yyyy-MM-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' And Format(C.TEST_DATE,'yyyy-MM-dd') <= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                  IIf(lngSampleNO = -1, "", " And A.SAMPLE_ID =" & lngSampleNO) & _
                  " order by A.SAMPLE_ID,C.TEST_DATE"
                  

        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        lngResultCount = -1
        
        
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("����"), "yyyy-mm-dd") '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(rsTmp("�걾��")) '�걾��
            aResults(lngResultCount) = aResults(lngResultCount) & "| "  '������
            aResults(lngResultCount) = aResults(lngResultCount) & "|Ѫ��"  '�걾
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '�ʿ�
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ͨ����") & "|" & Format(Val(IIf(IsNull(rsTmp("���")), "0", rsTmp("���"))), "###0.000") '���
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
errHandle:
    
    WriteLog "clsLISDev_File_ADC_CLIA_400.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function
