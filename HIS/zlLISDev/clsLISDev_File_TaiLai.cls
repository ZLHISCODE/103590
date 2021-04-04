VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_TaiLai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: ̩��-1 ø����
'����: δ֪
'����: �¶�
'�û�: ��������ҽԺ
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
    Dim strID As String
    Dim rsErs As ADODB.Recordset
    Dim str��Ŀ��� As String, str���� As String, lng�걾�� As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "SELECT A2 as �걾��,A16 as ����,A1  " & _
                 " From h_patient As A " & _
                 "Where A.A16 >='" & Format(dtStart, "yyyyMMdd") & "'" & _
                 " And A.A16 <='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And val(A.A2)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        strID = "***"
        Do While Not rsTmp.EOF
            strID = "" & rsTmp("A1")
            strSQL = "Select C4,C6 From h_result As B Where B.C2='" & strID & "'"
            str��Ŀ��� = ""
            Set rsErs = cnAccess.Execute(strSQL)
            Do Until rsErs.EOF
                If Trim("" & rsErs!C4) <> "" And Trim("" & rsErs!C6) <> "" Then
                    str��Ŀ��� = str��Ŀ��� & "|" & Trim("" & rsErs!C4) & "|" & Trim("" & rsErs!C6) '���
                End If
                rsErs.MoveNext
            Loop
            
            str���� = Trim("" & rsTmp("����"))
            If Len(str����) = 8 And IsNumeric(str����) Then
                str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
                If Not IsDate(str����) Then str���� = ""
            Else
                str���� = ""
            End If
            lng�걾�� = Val("" & rsTmp("�걾��"))
            
            If str���� <> "" And lng�걾�� <> 0 And str��Ŀ��� <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = str����
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & lng�걾��   '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "Ѫ��"                      '�걾
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '�ʿ�
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & str��Ŀ���
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

