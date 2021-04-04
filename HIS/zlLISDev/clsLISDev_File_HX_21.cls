VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HX_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:HX-21ϸ������ҩ��������
'����:
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
        .ConnectionString = "Provider=MSDASQL.1;Password=weichang;Persist Security Info=True;User ID=admin;Extended Properties=" _
                            & """" & "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}" & """"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False

        strSQL = "SELECT Format(B.Repdate,'yyyy-mm-dd') as �ͼ�����,B.Bbh as �걾��,B.Bgys as  ������,B.jzname AS �걾����, " & _
                 "Y.ypmc as ������,iif(Y.mg='��ҩ', 'R',iif(Y.mg='����', 'S', 'I')) AS ҩ�����," & _
                 "'1' as ҩ������,Y.MIC AS ��� FROM Base B,Ypbase Y WHERE B.useid = Y.useid and " & _
                 IIf(lngSampleNO = -1, "  b.Repdate between #" & Format(dtStart, "yyyy-mm-dd") & " 00:00:00#" & _
                 " And #" & Format(dtEnd, "yyyy-mm-dd") & " 23:59:59#", _
                 " B.Bbh ='" & lngSampleNO & "'") & " ORDER BY B.Repdate,B.Bbh"
        'cydate  �ͼ�����         Repdate ��������
        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1

        ReDim Preserve aResults(0)
        rsTmp.MoveFirst
        Dim �ͼ�����, �걾��, �걾���� As String

        Do Until rsTmp.EOF

            If �ͼ����� <> Format(rsTmp("�ͼ�����"), "YYYY-MM-DD") Or �걾�� <> Nvl(rsTmp("�걾��")) Or �걾���� <> Nvl(rsTmp("�걾����")) Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                �ͼ����� = Format(rsTmp("�ͼ�����"), "YYYY-MM-DD")
                �걾�� = Nvl(rsTmp("�걾��"))
                �걾���� = Nvl(rsTmp("�걾����"))
                aResults(lngResultCount) = �ͼ�����
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("�걾��"), " ") & "|" & Nvl(rsTmp("������"), " ") & "|" & Nvl(rsTmp("�걾����"), " ") & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("������"), " ") & "|" _
                         & Nvl(rsTmp("ҩ�����"), " ") & "^" & Nvl(rsTmp("ҩ������"), " ") & "^" & Nvl(rsTmp("���"), " ")
            ElseIf Nvl(rsTmp("������")) <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("������"), " ") & "|" _
                         & Nvl(rsTmp("ҩ�����"), " ") & "^" & Nvl(rsTmp("ҩ������"), " ") & "^" & Nvl(rsTmp("���"), " ")
            End If
            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_HX_21", strFile, Err.Number & " " & Err.Description
End Function