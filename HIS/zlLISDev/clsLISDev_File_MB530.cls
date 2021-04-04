VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MB530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø���� HEALES MB-530
'������:�����л��ɿƼ���չ���޹�˾
'�ͺ�:MB530
'����:���
'�û�:��ɽ����ҽԺ

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
    Dim str�걾�� As String, str��Ŀ As String
    
    Dim rsErs As ADODB.Recordset
    Dim str���� As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};PWD=huisong"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT distinct SampleNO as �걾��,SampleTestDT as ����,TestItemID as ͨ����,CalcRes as ���Խ��,ODRes as OD, A.[S/CORes] as SCO " & _
                " FROM TResult as A Where A.SampleTestDT >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & _
                "') And A.SampleTestDT< Cdate('" & Format(DateAdd("d", 1, dtEnd), "yyyy-MM-dd") & "') " & _
                IIf(lngSampleNo = -1, "", " And val(SampleNO)=" & lngSampleNo)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        ReDim aResults(0)
        'ø���� ͨ����|���Խ��[^OD^CutOff^sco]

        Do While Not rsTmp.EOF
                str���� = Format("" & rsTmp("����"), "yyyy-MM-dd")
                str�걾�� = Trim("" & rsTmp("�걾��"))
                
                If str�걾�� <> "" And Trim("" & rsTmp("���Խ��")) <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    str��Ŀ = Trim("" & rsTmp("ͨ����")) & "|" & Trim("" & rsTmp("���Խ��")) & "^" & Trim("" & rsTmp("OD")) & "^" & Trim("" & rsTmp("SCO")) & "^0"
                    aResults(UBound(aResults)) = str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp("�걾��")      '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "ѪҺ"                          '�걾
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '�ʿ�
                    
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str��Ŀ                           '��Ŀ�ͽ��
                End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


