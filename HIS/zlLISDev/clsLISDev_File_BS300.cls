VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BS300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BS300 ����������
'����:����
'����:�¶�
'�û�:�κ���ҽԺ

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
        Dim strUID As String, strPWD As String, strDate As String, strNo As String
        Dim strItem As String, strResult As String, strValue As String
        On Error GoTo errH

106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
        
        If Dir(App.Path & "\BS300.ini") <> "" Then
            strUID = ReadIni("BS300", "UID", App.Path & "\BS300.ini", "")
            strPWD = ReadIni("BS300", "PWD", App.Path & "\BS300.ini", "BS300")
        Else
            strUID = ""
            strPWD = "BS300"
            WriteIni "BS300", "UID", "", App.Path & "\BS300.ini"
            WriteIni "BS300", "PWD", "BS300", App.Path & "\BS300.ini"
        End If
110     cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=" & strUID & ";pwd=" & strPWD
112     cnAccess.Open
114     If cnAccess.State = adStateClosed Then Exit Function
        strSQL = "SELECT A.Sample_date, A.sample_no, A.Conc, B.name" & _
                 " FROM TESTDATAREC AS A, Item AS B " & _
                 " Where (((a.ITEM_ID) = [B].[ITEM_ID])) And A.Sample_date>=" & Format(dtStart, "yyyyMMdd") & _
                 " And A.Sample_date <=" & Format(dtEnd, "yyyyMMdd") & _
                  IIf(lngSampleNO = -1, "", " And val(A.sample_no)=" & lngSampleNO) & _
                 " ORDER BY CStr(A.sample_date & '  ' & a.sample_no) "

134     ReDim aResults(0)
    
136     Set rsTmp = cnAccess.Execute(strSQL)
138     If rsTmp.EOF Then
140         cnAccess.Close
142         WriteLog "BS300", "�޼�¼��", strSQL
            Exit Function
        End If
        strDate = "": strNo = ""
        strResult = ""
144     Do Until rsTmp.EOF
            
            If strDate <> Trim("" & rsTmp.Fields("Sample_date")) Or strNo <> Trim("" & rsTmp.Fields("Sample_NO")) Then
                
                If strResult <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strResult
                    
                End If
                

                strDate = Trim("" & rsTmp.Fields("Sample_date"))
                strNo = Trim("" & rsTmp.Fields("Sample_no"))
                
                strResult = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & "|" & strNo & "| |Boold|0"
                
                
            End If
            strItem = Trim("" & rsTmp.Fields("Name"))
            strValue = Trim("" & rsTmp.Fields("Conc"))
    
            If strItem <> "" And strValue <> "" Then
                strResult = strResult & "|" & strItem & "|" & strValue
            End If
            
184         rsTmp.MoveNext
        Loop
        If strResult <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strResult
        End If
186     ResultFromFile = aResults
        Exit Function
errH:
188     WriteLog "BS300", CStr(Erl()) & ":" & Err.Description, ""
End Function

