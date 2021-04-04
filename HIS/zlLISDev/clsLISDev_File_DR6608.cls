VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DR6608"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���߷����� DR6608
'����:��ɽ�ﰲҽ���豸���޹�˾
'����:�¶�
'�û�:�ɽ��ҽԺ

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
    Dim i As Integer, strDate As String, blnAdd As Boolean, strBh As String
    On Error GoTo errH
    
         
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "SELECT a.series_date,a.series_sampleid,a.series_itemname,a.series_value" & _
             " FROM  Series AS A Where CDate(a.series_date) >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "# And CDate(a.series_date) <=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#" & _
             " Order by a.series_date,a.series_sampleid"
'    strSQL = strSQL & "  Where cdate(A.sjrq) >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "#"
'    strSQL = strSQL & "   And  cdate(A.sjrq) <=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#" & _
'    IIf(lngSampleNo = -1, "", " And val(A.sxh)=" & lngSampleNo)
    WriteLog "File_DR6608", strSQL, ""
    
    ReDim aResults(0)
    
    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "File_DR6608", "�޼�¼��", ""
        Exit Function
    End If
    strBh = ""
    Do Until rsTmp.EOF
        If IsDate(Trim("" & rsTmp!series_date)) Then
            strDate = Trim("" & rsTmp!series_date)
            blnAdd = False
            If CDate(strDate) >= CDate(Format(dtStart, "yyyy-MM-dd 00:00:00")) And CDate(strDate) <= CDate(Format(dtEnd, "yyyy-MM-dd 23:59:59")) Then
                If lngSampleNo <> -1 Then
                     If Val(rsTmp!series_sampleid) = lngSampleNo Then
                        blnAdd = True
                     End If
                Else
                    blnAdd = True
                End If
 
            End If
            
            If blnAdd Then
                
                If strBh <> "" & rsTmp!series_sampleid Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        
                    aResults(UBound(aResults)) = rsTmp("series_date")  '����ʱ��
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("series_sampleid")) '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |Bold|0"
                End If
                
                If Trim("" & rsTmp!series_value) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp!series_itemname & "|" & rsTmp!series_value
                End If
                 
                strBh = "" & rsTmp!series_sampleid
            End If
        End If
        rsTmp.MoveNext
    Loop
        
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "File_DR6608", CStr(Erl()) & ":" & Err.Description, ""
End Function







