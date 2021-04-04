VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ҽԺ���ƣ��������ҽԺ
'�������ƣ����߶���������
'�����ͺţ�FIA-4
'�������ң��Ừ
'��    �ߣ�������

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
 
        On Error GoTo ResultFromFile_Err

    
        Dim varTemp As Variant '�����������
        Dim cnAccess As New ADODB.Connection
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
        Dim aResults() As String, objFile As New Scripting.FileSystemObject
        Dim strFilePath As String, strFilename As String
        Dim strResult() As String
        Dim tempdt As String
        Dim objStream As textStream
        Dim lngResultCount As Long, strBh As String
        Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
        Dim strStartDt As String, strEndDt As String

        Dim strImgErs As String, str���� As String
    
        Dim lngSampleID As String
        '��������±����
100     ReDim aResults(0)
102     aResults(0) = ""
104     ResultFromFile = aResults

106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
114         .Open
116         If .State = adStateClosed Then Exit Function
            'Sample_Pos OR Par_Id
118         strStartDt = Format(dtStart, "yyyy-MM-dd")
120         strStartDt = Mid(strStartDt, 1, 4) & Mid(strStartDt, 6, 2) & Mid(strStartDt, 9, 2)
        
122         strEndDt = Format(dtEnd, "yyyy-MM-dd")
124         strEndDt = Mid(strEndDt, 1, 4) & Mid(strEndDt, 6, 2) & Mid(strEndDt, 9, 2)
        
            'SQL������AutoTRFIA-2�ͣ�FIA-4���˶��Խ��������������������ݿ��C15��
        
126         strSQL = "select A.A0 as ��¼ID, A.A2 as ��¼���, A.A17 as ����ʱ��,  B.C2  as ���˱��, B.C4 as ��Ŀ���� ,B.C6 as ��Ŀ���, b.C15 as ��Ŀ���v" & _
                         " from h_Patient a , h_Result b " & _
                         " where a.A1=b.C2  and a.A17>=" & Chr(39) & strStartDt & Chr(39) & " And a.A17<= " & Chr(39) & strEndDt & Chr(39) & _
                         " Order By A.A2"
128         Set rsTmp = cnAccess.Execute(strSQL)
130         If rsTmp.EOF Then
132             .Close
                Exit Function
            End If
        
134            lngResultCount = -1
136         strBh = "***"
138         Do While Not rsTmp.EOF
140             If strBh <> rsTmp("��¼���") Then '�±걾
142                 lngResultCount = lngResultCount + 1
144                 ReDim Preserve aResults(lngResultCount)
                
146                 aResults(lngResultCount) = Mid(rsTmp("����ʱ��"), 1, 4) & "-" & Mid(rsTmp("����ʱ��"), 5, 2) & "-" & Mid(rsTmp("����ʱ��"), 7, 2) '����ʱ��
                
148                 str���� = Mid(rsTmp("����ʱ��"), 1, 4) & "-" & Mid(rsTmp("����ʱ��"), 5, 2) & "-" & Mid(rsTmp("����ʱ��"), 7, 2) 'Format(rsTmp("����ʱ��"), "yyyy-mm-dd")
                
150                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Val("" & rsTmp("��¼���"))   '�������
152                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          " "  '������
154                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          "ѪҺ"  '�걾
156                 aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                       ' Val(rsTmp("�ʿ�")) '�ʿ�
                    
158                 strBh = rsTmp("��¼���")
                End If

160                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        rsTmp("��Ŀ����") & "|" & IIf(IsNull(rsTmp("��Ŀ���")) Or rsTmp("��Ŀ���") = "", rsTmp("��Ŀ���v"), rsTmp("��Ŀ���")) & " " '���

162             rsTmp.MoveNext
           Loop
164        ResultFromFile = aResults
        End With

        Exit Function

ResultFromFile_Err:
    If Error = 1 Then
        Resume
    End If
End Function







