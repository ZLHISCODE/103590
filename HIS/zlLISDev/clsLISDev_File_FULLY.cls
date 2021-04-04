VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FULLY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FULLY ������
'����:
'����:�¶�
'�û�:��������������ҽԺ
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
    Dim lngID As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    ReDim aResults(0)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=amplly"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = " select a.QueryID as ID, a.sampledate as ʱ��,Sampleid as �걾��,Samplesource as �걾,ItemCode as ͨ����,Fvalue as ��� " & vbNewLine & _
                " from TB_RPRT_INFO a , TB_RPRT_RESULT b " & vbNewLine & _
                " Where a.QueryID = B.QueryID And a.sampledate >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')" & vbNewLine & _
                " And a.sampledate < Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "') " & vbNewLine & _
                IIf(lngSampleNO = -1, "", " And a.Sampleid ='" & lngSampleNO) & "' order by a.QueryID"
'        strSQL = "Select  Pat_Id as �걾��,Entry_Date as ����,GEST ,""AGE@EDD"" as AGE,AFP,AFPMOM,HCG,HCGMOM,T18_RISK as T18,T21_RISK as T21,AGE_RISK,NTD "
'        strSQL = strSQL & " From downs2t A Where A.entry_date >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
'        strSQL = strSQL & " And A.entry_date< Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "')" & _
        IIf(lngSampleNO = -1, "", " And val(A.Pat_Id)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If lngID <> Val(rsTmp("ID")) Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = Format(rsTmp("ʱ��"), "yyyy-mm-dd hh:mm:ss") & "|" & _
                                         rsTmp("�걾��") & "| |" & rsTmp("�걾") & "|0" & "|" & rsTmp("ͨ����") & _
                                         "|" & rsTmp("���")
            Else
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp("ͨ����") & "|" & rsTmp("���")
            End If
            lngID = Val(rsTmp("ID"))
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
   
End Function


