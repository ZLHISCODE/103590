VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_G9800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ӫ�����������  Genelight9800
'����:���Ű��������﹤�����޹�˾
'����:�¶�
'�û�:������Ⱦ��ҽԺ
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
        Dim strFilePath As String, ObjFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
        Dim strBh As String, strAnti As String, str��� As String
        Dim lngResultCount As Long
        
        On Error GoTo hErr

100     strFilePath = ObjFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
104     With cnAccess
106         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
108         .Open
110         If .State = adStateClosed Then Exit Function
112         strSQL = "Select Mid(A.patient_id,9) as ���,A.report_Date as ����,item_name as ��Ŀ,A.result as ���1,A.value_illno as ��� "
114         strSQL = strSQL & " From analysis A Where Format(A.report_Date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
116         strSQL = strSQL & " And Format(A.report_Date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
            IIf(lngSampleNO = -1, "", " And val(mid(A.patient_id,9))=" & lngSampleNO)
118         Set rsTmp = cnAccess.Execute(strSQL)
120         If rsTmp.EOF Then
122             .Close
                Exit Function
            End If
        
124         lngResultCount = -1
126         strBh = "***"
            
            'rsTmp.MoveFirst
128         Do Until rsTmp.EOF
130             If strBh <> rsTmp("���") Then '�±걾

132                 lngResultCount = lngResultCount + 1
134                 ReDim Preserve aResults(lngResultCount)

136                 aResults(lngResultCount) = Format(rsTmp("����"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
138                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Val(rsTmp("���")) '�������
140                 aResults(lngResultCount) = aResults(lngResultCount) & "| "
142                 aResults(lngResultCount) = aResults(lngResultCount) & "|Ѫ��"
144                 aResults(lngResultCount) = aResults(lngResultCount) & "|0"
146                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp!��Ŀ

148                 str��� = CStr(rsTmp!���)
150                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & str���
152                 strBh = rsTmp("���")

                End If
154             rsTmp.MoveNext
            Loop
        End With
156     ResultFromFile = aResults
        Exit Function
hErr:
158     WriteLog "G9800", Erl() & "��," & Err.Description, ""
End Function



