VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_BG10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ�������� BG-10
'����:÷�ݿ����߿Ƽ����޹�˾
'����:�¶�
'�û�:ɽ������ɽҽԺ

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
    Dim i As Integer, strDate As String, blnAdd As Boolean
    On Error GoTo errH
    
         
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "SELECT a.sxh , a.sjrq , a.tw ,a.hb, a.fio2, a.ph, a.pco2, a.po2 , a.phtc, a.pco2tc , a.po2tc , a.hco3, a.tco2 , a.beb, a.sbc ,a.beecf, a.so2 , a.p50, a.ri, a.[a-ado2]" & _
             " FROM jyjl AS a where a.sjrq is not null "
'    strSQL = strSQL & "  Where cdate(A.sjrq) >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "#"
'    strSQL = strSQL & "   And  cdate(A.sjrq) <=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#" & _
'    IIf(lngSampleNo = -1, "", " And val(A.sxh)=" & lngSampleNo)
    WriteLog "File_BG10", strSQL, ""
    
    ReDim aResults(0)
    
    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "File_BG10", "�޼�¼��", ""
        Exit Function
    End If
        
    Do Until rsTmp.EOF
        If IsDate(Trim("" & rsTmp!sjrq)) Then
            strDate = Trim("" & rsTmp!sjrq)
            blnAdd = False
            If CDate(strDate) >= CDate(Format(dtStart, "yyyy-MM-dd 00:00:00")) And CDate(strDate) <= CDate(Format(dtEnd, "yyyy-MM-dd 23:59:59")) Then
                If lngSampleNo <> -1 Then
                     If Val(rsTmp!sxh) = lngSampleNo Then
                        blnAdd = True
                     End If
                Else
                    blnAdd = True
                End If
 
            End If
            
            If blnAdd Then
                
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    
                aResults(UBound(aResults)) = Format(rsTmp("sjrq"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("sxh")) '�������
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |Bold|0"
                
                For i = 2 To rsTmp.Fields.count - 1
                    If Trim("" & rsTmp.Fields(i).Value) <> "" Then
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp.Fields(i).Name & "|" & rsTmp.Fields(i).Value
                    End If
                Next
            End If
        End If
        rsTmp.MoveNext
    Loop
        
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "File_BG10", CStr(Erl()) & ":" & Err.Description, ""
End Function



