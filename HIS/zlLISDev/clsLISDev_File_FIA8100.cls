VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA8100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�û����⴨����ҽԺ
'���ߣ�ŷ��ʤ
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
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim str�걾�� As String, str��Ŀ As String
    Dim str���� As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
   
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function   'blood_para_table
        strsql = "SELECT data as ����, nCa as ͨ����,PCT as PCT���,result as CRP���,blood_para_table.[Ca++] as C��Ӧ���,blood_para_table.[No] as �걾��" & _
                " FROM blood_para_table Where data>= '" & Format(dtStart, "yyyy.MM.dd") & _
                "' And data<= '" & Format(DateAdd("d", 1, dtEnd), "yyyy.MM.dd") & "' " & _
                IIf(lngSampleNO = -1, "", " And No=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        Do While Not rsTmp.EOF
                str���� = Replace(Format("" & rsTmp("����"), "yyyy-MM-dd hh24:mi:ss"), ".", "-")
                str�걾�� = Val(rsTmp("�걾��"))
                str��Ŀ = ""
                
                    If rsTmp("ͨ����") = "PCT" Then
                      str��Ŀ = "|" & rsTmp("ͨ����") & "|" & rsTmp("PCT���")
                    
                    ElseIf rsTmp("ͨ����") = "U-CRP" And rsTmp("CRP���") <> "" And rsTmp("C��Ӧ���") <> "" Then
                      str��Ŀ = "|" & rsTmp("ͨ����") & "|" & rsTmp("CRP���") & "|" & "CRP" & "|" & rsTmp("C��Ӧ���")
                    ElseIf rsTmp("ͨ����") = "U-CRP" And rsTmp("CRP���") = "" And rsTmp("C��Ӧ���") <> "" Then
                      str��Ŀ = "|" & "CRP" & "|" & rsTmp("C��Ӧ���")
                    ElseIf rsTmp("ͨ����") = "U-CRP" And rsTmp("CRP���") <> "" And rsTmp("C��Ӧ���") = "" Then
                      str��Ŀ = "|" & rsTmp("ͨ����") & "|" & rsTmp("CRP���")
                    End If
               
                If str��Ŀ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||" & str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str�걾��                       '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "ѪҺ"                          '�걾
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '�ʿ�
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str��Ŀ                               '��Ŀ�ͽ��
                    
                    ElseIf aResults(UBound(aResults)) = "" Then
                    
                    aResults(UBound(aResults)) = str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str�걾��                       '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "ѪҺ"                          '�걾
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '�ʿ�
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str��Ŀ
                    End If
                    aResults(UBound(aResults)) = Replace(aResults(UBound(aResults)), "��", "")
                End If
                rsTmp.MoveNext
            Loop
  
        
    End With
    ResultFromFile = aResults
End Function




