VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_KOFA_AM300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:KOFA AM300 ����������
'����:
'����:ŷ��ʤ
'�û�:��ݸݸ��ҽԺ

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
    Dim strDate As String, strNo As String
    Dim strItem As String, strResult As String, strValue As String
    On Error GoTo errH

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path


    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    If Format(dtStart, "yyyy-mm-dd") = Format(Now, "yyyy-mm-dd") Then
        strSQL = "SELECT  a.SampleID as �걾��,b.ItemID as ͨ����,a.ID,a.TestDate,b.ItemResult as ���,b.TestDate as ����ʱ�� " & _
                 "FROM SampleBaseInfor a, SampleTestData b " & _
                 "where a.ID=b.SampleID and Format(a.TestDate, 'yyyyMMdd')>=" & Format(dtStart, "yyyyMMdd") & _
               " and Format(a.TestDate, 'yyyyMMdd') <=" & Format(dtEnd, "yyyyMMdd") & _
                 IIf(lngSampleNO = -1, "", " And val(a.SampleID)=" & lngSampleNO) & _
               "  ORDER BY CStr(a.TestDate & '  ' & a.SampleID) "
    Else
        strSQL = "SELECT  a.SampleID as �걾��,b.ItemID as ͨ����,a.ID,a.TestDate,b.ItemResult as ���,b.TestDate as ����ʱ�� " & _
                 "FROM SampleBaseInforHistory a, SampleTestDataHistory b " & _
                 "where a.ID=b.SampleID and Format(a.TestDate, 'yyyyMMdd')>=" & Format(dtStart, "yyyyMMdd") & _
               " and Format(a.TestDate, 'yyyyMMdd') <=" & Format(dtEnd, "yyyyMMdd") & _
                 IIf(lngSampleNO = -1, "", " And val(a.SampleID)=" & lngSampleNO) & _
               "  ORDER BY CStr(a.TestDate & '  ' & a.SampleID) "
    End If


    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    strDate = "": strNo = ""
    strResult = ""
    Do Until rsTmp.EOF
        If strDate <> Trim("" & rsTmp.Fields("����ʱ��")) Or strNo <> Trim("" & rsTmp.Fields("�걾��")) Then
            If strResult <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = strResult

            End If
            strDate = Trim("" & rsTmp.Fields("����ʱ��"))
            strNo = Trim("" & rsTmp.Fields("�걾��"))
            strResult = strDate & "|" & strNo & "| |Ѫ��|0"


        End If
        strItem = Trim("" & rsTmp.Fields("ͨ����"))
        strValue = Trim("" & rsTmp.Fields("���"))

        If strItem <> "" And strValue <> "" Then
            strResult = strResult & "|" & strItem & "|" & strValue
        End If

        rsTmp.MoveNext
    Loop
    If strResult <> "" Then
        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
        aResults(UBound(aResults)) = strResult
    End If
    ResultFromFile = aResults
    Exit Function
errH:

End Function



