VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_RT2100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�׶�ø����
'�ͺ�:RT-2100C
'����:Rayto(�׶�)
'����:�µ���
'�û�:��Դ������ҽԺ
'����:2012-04-17
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
                        ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
    'strFile������·�����ļ���
    'lngSampleNo���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    'dtStart����ʼʱ�䣬���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    'dtEnd������ʱ�䡣ֻ��ѡȡ����걾(lngSample=-1)ʱ�ò�����Ч�������ָ���ò�������ʱ�䷶ΧΪ>=dtStart
    '���صļ��������飨������������밴�����±�׼��֯�����
    'Ԫ��֮���� | �ָ�
    '�� 0 ��Ԫ�أ�����ʱ��
    '�� 1 ��Ԫ�أ��������
    '�� 2 ��Ԫ�أ�������
    '�� 3 ��Ԫ�أ��걾
    '�� 4 ��Ԫ�أ��Ƿ��ʿ�Ʒ
    '�ӵ� 5 ��Ԫ�ؿ�ʼΪ��������ÿ����Ԫ�ر�ʾһ��������Ŀ
    '�磺�� 5i ��Ԫ��Ϊ������Ŀ���� 5i+1 ��Ԫ�ر�ʾ������
 
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cnErs As New ADODB.Connection, strFileErs As String, rsErs As ADODB.Recordset, strImgErs As String
    Dim str_����ʱ�� As String, str_������Ŀ As String, str_������� As String, str_������ As String, str_�걾 As String
    Dim YYx As Integer, str_������ As String
    
        
    strFilePath = objFile.GetParentFolderName(strFile)  'ȡ·��
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
            .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
            .Open
            If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    ReDim aResults(0)
    strImgErs = ""
    
    If strFileErs <> "" Then
        strSQL = "SELECT Trim(CSTR(a.sample_date)) + ' '+ Trim(a.sample_time) As drq,a.sample_id As cybh,b.Item_Id As Item, " & _
                 "b.Result,b.Abs As OD,b.Result_Det As YYX From sample_info a,item_result b where format(a.test_date,'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' " & _
                 "and format(a.test_date,'yyyy-mm-dd')<= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNo = -1, "", " and VAL(a.sample_id)=" & lngSampleNo) & _
                 " And a.com_id=b.com_id and a.sample_ID=b.sample_ID and a.test_date=b.test_date order by a.sample_date,a.sample_id"
        Set rsErs = cnErs.Execute(strSQL)
        
        While Not rsErs.EOF
            '�����ļ��ṹ��ȡ������Ϣ
            If Val("" & rsErs!cybh) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                '�� 0 ��Ԫ�أ�����ʱ��
                '�� 1 ��Ԫ�أ��������
                '�� 2 ��Ԫ�أ�������
                '�� 3 ��Ԫ�أ��걾
                '�� 4 ��Ԫ�أ��Ƿ��ʿ�Ʒ
                '�� 5 ��Ԫ�ؿ�ʼΪ��������ÿ����Ԫ�ر�ʾһ��������Ŀ
                '�磺�� 5i ��Ԫ��Ϊ������Ŀ���� 5i+1 ��Ԫ�ر�ʾ������
                str_����ʱ�� = Format(Trim(rsErs!drq), "yyyy-MM-dd HH:mm:ss")                    '����ʱ��
                str_������� = Trim(rsErs!cybh)
                str_������ = " "
                str_�걾 = "Ѫ��|0"
                YYx = rsErs!YYx
                Select Case YYx
                    Case 0: str_������ = "��"
                    Case 1: str_������ = "����(+)"
                    Case 2: str_������ = "����(+)"
                    Case 3: str_������ = "?����"
                    Case 4: str_������ = "L"
                    Case 5: str_������ = "H"
                    Case 6: str_������ = "���ɡ�"
                    Case 7: str_������ = "����"
                    Case 8: str_������ = "���Է�Ӧ"
                    Case 9: str_������ = "���������"
                    Case 10: str_������ = "��CDC����"
                    Case 11: str_������ = "������"
                End Select
                
                '������Ŀ
                str_������Ŀ = rsErs!Item & "|" & str_������ & "^" & rsErs!OD & _
                                "^" & rsErs!result & "^0.00"
                aResults(UBound(aResults)) = str_����ʱ�� & "|" & str_������� & "|" & str_������ & "|" & str_�걾 & "|" & str_������Ŀ
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults
End Function
