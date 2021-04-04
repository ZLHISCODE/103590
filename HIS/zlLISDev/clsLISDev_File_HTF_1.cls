VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HTF_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:HTF-1��ѧ�������߷�����     ����ɸ��
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
        .ConnectionString = "DBQ=" & strFile & ";User ID=admin;Password=weichang;DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = ""
        blnUnion = False

        strSQL = "SELECT ��� as �걾��,����ҽʦ,�걾 ,�������,AFP,HCG,PAPPA,AFPMOM,HCGMOM,PAPPAMOM,DS,ET,NTD,ASR,AFPASR,NAFP,NHCG,NPAPPA from T_dssign WHERE  " & _
                 IIf(lngSampleNO = -1, "  ������� between #" & Format$(dtStart, "yyyy-mm-dd") & " 00:00:00#" & _
                 " And #" & Format$(dtEnd, "yyyy-mm-dd") & " 23:59:59#", _
                 "  ������� between #" & Format$(dtStart, "yyyy-mm-dd") & " 00:00:00# and #" & Format$(dtStart, "yyyy-mm-dd") & " 23:59:59# and  ��� =" & lngSampleNO) & " ORDER BY ���+0"

        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format$(rsTmp("�������"), "YYYY-MM-DD HH:NN:SS")
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("�걾��") & "|" & IIf(IsNull(rsTmp("����ҽʦ")), " ", rsTmp("����ҽʦ")) & "|" & IIf(IsNull(rsTmp("�걾")), "Ѫ��", rsTmp("�걾")) & "|0"
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("AFP")), "", "|AFP|" & rsTmp("AFP")) & IIf(IsNull(rsTmp("HCG")), "", "|HCG|" & rsTmp("HCG"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("PAPPA")), "", "|PAPPA|" & rsTmp("PAPPA")) & IIf(IsNull(rsTmp("AFPMOM")), "", "|AFPMOM|" & rsTmp("AFPMOM"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("HCGMOM")), "", "|HCGMOM|" & rsTmp("HCGMOM")) & IIf(IsNull(rsTmp("PAPPAMOM")), "", "|PAPPAMOM|" & rsTmp("PAPPAMOM"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("DS")), "", "|DS|" & rsTmp("DS")) & IIf(IsNull(rsTmp("ET")), "", "|ET|" & rsTmp("ET"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("NTD")), "", "|NTD|" & rsTmp("NTD")) & IIf(IsNull(rsTmp("ASR")), "", "|ASR|" & rsTmp("ASR"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("AFPASR")), "", "|AFPASR|" & rsTmp("AFPASR")) & IIf(IsNull(rsTmp("NAFP")), "", "|NAFP|" & rsTmp("NAFP"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("NHCG")), "", "|NHCG|" & rsTmp("NHCG")) & IIf(IsNull(rsTmp("NPAPPA")), "", "|NPAPPA|" & rsTmp("NPAPPA"))
            aResults(lngResultCount) = Replace$(aResults(lngResultCount), "|.", "|0.")
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_HTF_1", strFile, Err.Number & " " & Err.Description
End Function