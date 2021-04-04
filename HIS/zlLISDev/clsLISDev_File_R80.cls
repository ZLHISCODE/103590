VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_R80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫճ�ȷ����� LG-R-80F
'����:
'����:�¶�
'�û�:����202ҽԺ
'�޸ģ������� ---- ͼ������
'�û���������ҽԺ

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
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String
    Dim strChar As String
    Dim intMaxValue As Integer
    Dim str_�������� As String
    
    str_�������� = Trim(ReadIni("��������", "����", App.Path & "\LG-R-80F.ini"))
    If str_�������� = "" Then
        WriteIni "��������", "����", "LG-R-80F", App.Path & "\LG-R-80F.ini"
        str_�������� = "LG-R-80F"
    End If
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strsql = "SELECT mid(A.recordnum,1,8) as ����, mid(A.recordnum,9) as �걾��,sex as �Ա�, A.f1, A.f2, A.f3, A.f4, A.f5, A.f6, A.f7, A.T1, A.T2, A.T3, A.T4, A.T5, A.T6, A.T7, A.T8, A.C2, A.C3, A.C4, A.C5, Result " & _
                "FROM patient AS A where mid(A.recordnum,1,8)>='" & Format(dtStart, "yyyyMMdd") & "' And mid(A.recordnum,1,8)<='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                IIf(lngSampleNO = -1, "", " And val(Mid(A.recordnum,9))=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("�걾��") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                str���� = Trim("" & rsTmp("����"))
                
                aResults(lngResultCount) = Format(Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2), "yyyy-MM-dd") '����ʱ��
                                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("�걾��"))   '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("�ʿ�")) '�ʿ�
                    
                strBh = rsTmp("�걾��")
            End If
            
            For iField = 3 To 21
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    iField - 2 & "|" & Val(IIf(IsNull(rsTmp.Fields(iField).Value), "0", rsTmp.Fields(iField).Value)) '���
            Next
            If Trim("" & rsTmp!result) <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & Trim("" & rsTmp!result)
            End If
            '-- Ѫ�����ͼ������
            '-- ͼ������X-�ᣬY����
            '-- ��ʽ������;��������;���߼��������
            '        ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
            '              ���߼��������:����1��Y����1,����1��Y����2.....|����2��Y����1,����2��Y����2,...~����Y�ӱ�1,����Y�ӱ�2
            If str_�������� = "LG-R-80F" Then
                strImg = vbCrLf & "ճ����������;2;20,230|5-5,30-30,100-100,200-200|0-0,5-5,10-10,15-15,20-20;"
                
                If Trim("" & rsTmp("�Ա�")) = "Ů" Then
    '                ND0 = 9.25: QB0 = 10: ND1 = 4.4: QB1 = 150
    '                ND2 = 6.5: QB2 = 10: ND3 = 3.65: QB3 = 150
                    strImg = strImg & "9.25,10,4.4,150|6.5,10,3.65,150"
                Else
    '                ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
    '                ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                    strImg = strImg & "9.58,10,4.6,150|6.8,10,3.73,150"
                End If
                strImg = strImg & "~5-" & rsTmp!f3 & ",30-" & rsTmp!f2 & ",200-" & rsTmp!f1  '���
                strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),220,1" '������
                
                aResults(lngResultCount) = aResults(lngResultCount) & strImg
            ElseIf str_�������� = "LG-R-80F_WXX" Then
                strImg = vbCrLf & "ճ����������;2;35,250|1-1,40-3,80-10,120-30,180-100,230-300|0-0,5-5,10-10,15-15,20-20,25-25,30-30;"
                
                If Trim("" & rsTmp("�Ա�")) = "Ů" Then
                    strImg = strImg & "11.55,12,5.4,150"
                Else
                    strImg = strImg & "11.78,12,5.6,150"
                End If
                strImg = strImg & "~"
                strImg = strImg & ";VIS(mPa.s),35,35~SHR(1/S),225,2" '������
                aResults(lngResultCount) = aResults(lngResultCount) & strImg
            End If
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

