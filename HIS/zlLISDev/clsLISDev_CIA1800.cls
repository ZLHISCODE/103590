VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CIA1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CIA-1800 ����
'����:���
'����:������
'�û�:������ҽҽԺ
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select index,testid,testdate From patientinfo Where cdate(testdate) >= Cdate('" & Format(dtStart, "yyyy-mm-dd") & "')"
        strSQL = strSQL & " And cdate(testdate) <= Cdate('" & Format(dtEnd, "yyyy-mm-dd") & "')" & _
        IIf(lngSampleNO = -1, "", " And val(TestID)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1

        Do Until rsTmp.EOF
            strSQL = "select item,concspec from assayresult where patientindex = " & rsTmp("index")
            Set rsItem = cnAccess.Execute(strSQL)
            If rsTmp.EOF Then GoTo NextDo:
            
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsTmp("testdate") & " " & ":00", "yyyy-mm-dd hh:mm:ss") '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val(rsTmp("testid")) '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "| |����Ѫ|0"
            Do Until rsItem.EOF
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsItem("item") & "|" & rsItem("concspec")
                rsItem.MoveNext
            Loop
                    
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults
   
End Function





