VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ATB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:VITEK ATB ϸ��������
'����:÷�ﰣ VITEK  �����������ƿ���������
'����:�¶�
'�û�:��ú������ҽԺ

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
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
        Dim strBh As String, strAnti As String, str��� As String, strϸ�� As String
        Dim lngResultCount As Long
        Dim strType As String  '2009-07-19 ���� �»��ı��и��û��������Ǵ�Result������ȡ���ݣ���ԭ���Ǵ�ReqToExe����������,
                               'ATBFile.ini
        Dim strBZ As String '�Ƿ�ӱ�ע���ϰ��liscomm��֧�ּ��鱸ע��Ҫ������
        
        On Error GoTo errH
100     strType = ReadIni("����", "TYPE", App.Path & "\ATBFile.ini")
102     strBZ = ReadIni("����", "��ע", App.Path & "\ATBFile.ini")
104     If strType = "" Then strType = "REQ"
        
106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path

110     cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=scsl"
112     cnAccess.Open
114     If cnAccess.State = adStateClosed Then Exit Function
116     strSQL = "Select A.ReqNO as ���,A.ExeDate as ����,A.GermID as ϸ��,B.samNo as �걾�� "
118     If strType <> "Result" Then
120         strSQL = strSQL & " From ReqToExe A,RequistionBill B "
        Else
122         strSQL = strSQL & " From ResultExe A,ResultBill B "
        End If
124     strSQL = strSQL & " Where A.ReqNO=B.ReqNO And   A.ExeDate >=#" & Format(dtStart, "yyyy-MM-dd") & "#"
126     strSQL = strSQL & " And  A.ExeDate <=#" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
        IIf(lngSampleNO = -1, "", " And val(A.ReqNo)=" & lngSampleNO)
128     WriteLog "File_ATB", strSQL, ""
    
130     strBh = "***"
132     strϸ�� = ""
134     ReDim aResults(0)
    
136     Set rsTmp = cnAccess.Execute(strSQL)
138     If rsTmp.EOF Then
140         cnAccess.Close
142         WriteLog "File_ATB", "�޼�¼��", ""
            Exit Function
        End If
        
144     Do Until rsTmp.EOF
146         If strBh <> rsTmp("���") Or strϸ�� <> Trim("" & rsTmp("ϸ��")) Then '�±걾
148             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
150             aResults(UBound(aResults)) = Format(rsTmp("����"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
152             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    Val(rsTmp("�걾��")) '�������
154             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
156             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    IIf(IsNull(rsTmp("ϸ��")), " ", rsTmp("ϸ��")) '�걾
158             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
160             strϸ�� = Trim("" & rsTmp("ϸ��"))
162             strBh = rsTmp("���")
164             strAnti = ""
166             If strType <> "Result" Then
168                 strSQL = "Select AntiID,MicExplain,MicResult,ResShow From MicResult Where val(ReqNo)=" & strBh
                Else
170                 strSQL = "Select AntiID,MicExplain,MicResult,ResShow From ResultMic Where val(ReqNo)=" & strBh
                End If
172             Set rsAnti = cnAccess.Execute(strSQL)
174             Do Until rsAnti.EOF
176                 strAnti = strAnti & "|" & rsAnti!AntiID & "|" & rsAnti!MicExplain & "^1^" & rsAnti!ResShow
178                 rsAnti.MoveNext
                Loop
180             aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAnti
182             WriteLog "File_ATB", aResults(UBound(aResults)), ""
            End If
184         rsTmp.MoveNext
        Loop
        
186     ResultFromFile = aResults
        Exit Function
errH:
188     WriteLog "File_ATB", CStr(Erl()) & ":" & Err.Description, ""
End Function