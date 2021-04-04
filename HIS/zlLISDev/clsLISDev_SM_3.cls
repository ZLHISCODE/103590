VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SM_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsPaInfo As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        If dtStart < Date Then
            strSQL = "select checkDate as ʱ��,customername as �걾��,'Ѫ��' as �걾,'0' as �Ƿ��ʿ�,decision as ���,kindsymbol as ͨ���� " & _
                     " from platerecord " & _
                     " where checkDate >= #" & Format(dtStart, "yyyy/mm/dd") & "#" & _
                     " Order By customername ,checktime"
        Else
            strSQL = "select checkDate as ʱ��,customername as �걾��,'Ѫ��' as �걾,'0' as �Ƿ��ʿ�,decision as ���,kindsymbol as ͨ���� " & _
                     " from platerecord " & _
                     " where checkDate = #" & Format(dtStart, "yyyy/mm/dd") & "#" & _
                     " Order By customername,checktime "
            
        End If
        
        Set rsTmp = cnAccess.Execute(strSQL)
        lngSampleID = 0: lngResultCount = -1
        Do Until rsTmp.EOF
            If lngSampleID <> Val(rsTmp("�걾��")) Then
                '�½�ͷ��Ϣ
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = Format(rsTmp("ʱ��"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("�걾��") '�걾��
                aResults(lngResultCount) = aResults(lngResultCount) & "| "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("�걾") '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"  '�ʿ�
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ͨ����") & "|" & Replace(Replace(rsTmp("���"), "����", "���ԣ�-��"), "����", "���ԣ�+��")
            Else
                'ָ����Ϣ
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ͨ����") & "|" & Replace(Replace(rsTmp("���"), "����", "���ԣ�-��"), "����", "���ԣ�+��")
            End If
            lngSampleID = Val(rsTmp("�걾��"))
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

