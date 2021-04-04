VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DR6660"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����  DR6660-2
'����:
'����:������
'�û�:��կ������ҽԺ

Public Function GetStartCmd() As String
    '��һ������ʱ����
End Function

Public Function GetEndCmd() As String
    '�ӿڹرյ�ʱ�����
End Function

Public Function GetAnswerCmd() As String
    '��ʱ���͵�ָ��
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim lngResultCount As Long
    On Error GoTo ErrH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select series_itemname as ��Ŀ,series_value as ���,series_sampleID as �걾��,series_date as ���� From series Where series_date >= '" & Format(dtStart, "yyyy-mm-dd") & "'"
        strSQL = strSQL & " And series_date <= '" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
        IIf(lngSampleNO = -1, "", " And val(Series_sampleID) = " & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1

        Do Until rsTmp.EOF
            
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsTmp("����") & " " & "00:00:00", "yyyy-mm-dd hh:mm:ss") '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val(rsTmp("�걾��")) '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "| |ѪҺ|0|" & rsTmp("��Ŀ") & "|" & rsTmp("���")
                    
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults
    Exit Function
ErrH:
    If Error = 1 Then
        Resume
    End If
End Function





