VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EC5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����� EC-5000
'����:�����Ʒ�
'����:�¶�
'�û�:����������ҽԺ(̫ԭ)

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
        Dim strBh As String, strItem As String, i As Integer
        On Error GoTo hErr
    
100     ReDim aResults(0)
102     strFilePath = objFile.GetParentFolderName(strFile)
104     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
106     If Dir(strFile) = "" Then
108         ResultFromFile = aResults
            Exit Function
        End If
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
114         .Open
116         If .State = adStateClosed Then
118             ResultFromFile = aResults
                Exit Function
            End If
120         strSQL = "" & _
                    "SELECT Mid(A.�걾��,1,8) AS ����, Mid(A.�걾��,9) AS �걾��, A.��ϸ��, A.��ϸ��, A.���Ӽ���, A.��Ƥϸ��, A.�ᾧ, A.�γ�����, A.�������, A.����, A.ճҺ˿, A.��ɫ, A.������, A.��ԭ, A.������, A.ͪ��, A.ǱѪ, A.������, A.��������, A.[��ϸ��(�ɻ�)], A.������, A.����, A.PHֵ" & _
                    " FROM TestResult AS A Where Mid(A.�걾��,1,8)>='" & Format(dtStart, "yyyyMMdd") & "' And Mid(A.�걾��,1,8)<= '" & Format(dtEnd, "yyyyMMdd") & "' " & _
                    IIf(lngSampleNO = -1, "", " And val(Mid(A.�걾��,9))=" & lngSampleNO)
                    
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strItem = ""
132         Do While Not rsTmp.EOF
134             strItem = ""
136             strItem = Mid("" & rsTmp("����"), 1, 4) & "-" & Mid("" & rsTmp("����"), 5, 2) & "-" & Mid("" & rsTmp("����"), 7, 2) '����ʱ��
138             strItem = strItem & "|" & Val("" & rsTmp("�걾��"))   '�������
140             strItem = strItem & "| "  '������
142             strItem = strItem & "|ѪҺ"  '�걾
144             strItem = strItem & "|0"
                
146             For i = 2 To rsTmp.Fields.count - 1
148                 strBh = Trim("" & rsTmp.Fields(i).Value)
150                 If strBh <> "" Then
152                     If InStr(strBh, "+") > 0 Or InStr(strBh, "-") > 0 Then
154                         strBh = Trim(Mid(strBh, 1, 2))
                        End If
156                     strItem = strItem & "|1" & Format(i, "00") & "|" & strBh
                    End If
                Next
158             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
160             aResults(UBound(aResults)) = strItem
                
162             rsTmp.MoveNext
            Loop
                
        End With
164     ResultFromFile = aResults
        Exit Function
hErr:
166     WriteLog "EC5000", CStr(Erl()) & ":" & Err.Description, ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
End Sub





