VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_IMMU1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: IMMULITE 1000 ȫ�Զ��������߷�����
'����: DPC
'����: �¶�
'�û�: ����3201ҽԺ
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
    Dim strID As String
    Dim rsErs As ADODB.Recordset
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=www.dpc.com"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "SELECT Rpt_no as �걾��,Format(Rpt_date,'yyyy-MM-dd') as ����,Rpt_ID  " & _
                 " From TBL_Rpt As A " & _
                 "Where Format(A.RPT_Date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                 " And Format(A.RPT_Date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And val(Rpt_no)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        strID = "***"
        Do While Not rsTmp.EOF
        
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                
            strID = "" & rsTmp("Rpt_ID")
            
            strSQL = "Select ITEM_CODE,ITEM_RESULT From TBL_Rpt_Item As B Where B.RPT_ID='" & strID & "'"
            Set rsErs = cnAccess.Execute(strSQL)
            Do Until rsErs.EOF
                If aResults(UBound(aResults)) = "" Then
                    aResults(UBound(aResults)) = Trim("" & rsTmp("����"))
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("�걾��"))   '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "Ѫ��"                      '�걾
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '�ʿ�
                End If
                If Trim("" & rsErs!ITEM_RESULT) <> "" And Trim("" & rsErs!ITEM_CODE) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim("" & rsErs!ITEM_CODE) & _
                         "|" & Trim("" & rsErs!ITEM_RESULT)  '���
                End If
                rsErs.MoveNext
            Loop
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function



