VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ZL6000I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ����ZL6000I
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsVisc As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim intLoop As Integer
    Dim strͼ��  As String

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    Set cnAccess = New ADODB.Connection
    With cnAccess
        .ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;User ID=admin;Extended Properties=" _
                            & """" & "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}" & """"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False

        strSQL = "select * from testdata WHERE " & _
                 IIf(lngSampleNO = -1, "  ���� between #" & Format(dtStart, "yyyy-mm-dd") & "#" & _
                 " And #" & Format(dtEnd, "yyyy-mm-dd") & "#", _
                 " ��� ='" & lngSampleNO & "'") & " ORDER BY ����,���"
        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            ReDim Preserve aResults(0)
            ResultFromFile = aResults
            Exit Function
        End If
        lngResultCount = -1

        ReDim Preserve aResults(0)
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(Nvl(rsTmp!����) & " " & Nvl(rsTmp!ʱ��), "YYYY-MM-DD HH:NN:SS")    '����ʱ��
            '            aResults(lngResultCount) = Nvl(rsTmp!����) & Nvl(rsTmp!����) & " " & Nvl(rsTmp!ʱ��)    '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp!���)             '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp!����ҽ��)         '������
            aResults(lngResultCount) = aResults(lngResultCount) & "|ѪҺ|0"
            'ȫѪճ������
            strSQL = "select * from Visc where TestDataID =" & rsTmp!TestDataID & " order by ShearRate"
            Set rsVisc = .Execute(strSQL)
            Do Until rsVisc.EOF
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsVisc!ShearRate & "|" & Nvl(rsVisc!Visc)
                rsVisc.MoveNext
            Loop
            '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "|ȫѪճ��|" & Nvl(rsTmp!Ѫ��ճ��, " ") & "|Ѫ��|" & Nvl(rsTmp!Ѫ��, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|ѹ��|" & Nvl(rsTmp!ѹ��, " ") & "|ȫѪ�������ָ��|" & Nvl(rsTmp!ȫѪ�������ָ��, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|ȫѪ�������ָ��|" & Nvl(rsTmp!ȫѪ�������ָ��, " ") & "|Ѫ������Kֵ|" & Nvl(rsTmp!Ѫ������Kֵ, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|��ϸ���ۼ�ָ��|" & Nvl(rsTmp!��ϸ���ۼ�ָ��, " ") & "|��ϸ������ָ��|" & Nvl(rsTmp!��ϸ������ָ��, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|ȫѪ���л�ԭճ��|" & Nvl(rsTmp!ȫѪ���л�ԭճ��, " ") & "|ȫѪ���л�ԭճ��|" & Nvl(rsTmp!ȫѪ���л�ԭճ��, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|��ϸ������ָ��TK|" & Nvl(rsTmp!��ϸ������ָ��TK, " ") & "|��ϸ������ָ��|" & Nvl(rsTmp!��ϸ������ָ��, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|����ճ��|" & Nvl(rsTmp!����ճ��, " ") & "|��ά����ԭ|" & Nvl(rsTmp!��ά����ԭ, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|Ѫ���̴�|" & Nvl(rsTmp!Ѫ���̴�, " ") & "|������֬|" & Nvl(rsTmp!������֬, " ")

            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_ZL6000I", strFile, Err.Number & " " & Err.Description
End Function