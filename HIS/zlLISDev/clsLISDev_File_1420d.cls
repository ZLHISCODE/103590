VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_1420d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:VICTOR2 D (1420d) ���Ƿ�����(����������ɸ��)
'����:61-��𰣶�Ĭ(PerkinElmer)
'����:�¶�
'�û�:��������ҽԺ

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

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
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
    
    strFilePath = objFile.GetParentFolderName(strFile)
    ReDim aResults(0)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=amplly"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select  Pat_Id as �걾��,Entry_Date as ����,GEST ,""AGE@EDD"" as AGE,AFP,AFPMOM,HCG,HCGMOM,T18_RISK as T18,T21_RISK as T21,AGE_RISK,NTD "
        strSQL = strSQL & " From downs2t A Where A.entry_date >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
        strSQL = strSQL & " And A.entry_date< Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "')" & _
        IIf(lngSampleNO = -1, "", " And val(A.Pat_Id)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = Format(rsTmp("����"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("�걾��"))  '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ѫ��"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|GEST|" & IIf(Trim("" & rsTmp!GEST) = "", " ", Trim("" & rsTmp!GEST))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AGE|" & IIf(Trim("" & rsTmp!AGE) = "", " ", Trim("" & rsTmp!AGE))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFP|" & IIf(Trim("" & rsTmp!AFP) = "", " ", Trim("" & rsTmp!AFP))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFPMOM|" & IIf(Trim("" & rsTmp!AFPMOM) = "", " ", Trim("" & rsTmp!AFPMOM))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HCG|" & IIf(Trim("" & rsTmp!HCG) = "", " ", Trim("" & rsTmp!HCG))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HCGMOM|" & IIf(Trim("" & rsTmp!HCGMOM) = "", " ", Trim("" & rsTmp!HCGMOM))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|T18|" & IIf(Trim("" & rsTmp!T18) = "", " ", Trim("" & rsTmp!T18))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|T21|" & IIf(Trim("" & rsTmp!T21) = "", " ", Trim("" & rsTmp!T21))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AGE_RISK|" & IIf(Trim("" & rsTmp!AGE_RISK) = "", " ", Trim("" & rsTmp!AGE_RISK))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|NTD|" & IIf(Trim("" & rsTmp!NTD) = "", " ", Trim("" & rsTmp!NTD))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
   
End Function
