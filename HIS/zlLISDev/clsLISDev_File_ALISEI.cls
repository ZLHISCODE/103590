VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ALISEI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø����
'����:alisei
'�û�:����ҽԺ,�Ͼ���������ҽԺ
'by����ǿ

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
    Dim strsql As String, rsTmp As ADODB.Recordset, rs��Ŀ As ADODB.Recordset, rsʱ�� As ADODB.Recordset
    Dim lngResultCount As Long, strcutoff As String, str��Ŀ As String
    Dim str���� As String, str���Խ�� As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFile = strFilePath & "\EXPORT.MDB"
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        '���Ҽ�������
        strsql = "select SessionID as ����,SessionDate as ���� from Session" & _
                " where Format(SessionDate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(SessionDate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' "
        Set rsʱ�� = cnAccess.Execute(strsql)
        If rsʱ��.EOF Then
           .Close
           Exit Function
        End If
        lngResultCount = -1
        Do While Not rsʱ��.EOF
        
       '����test���е���Ŀ od cutoffֵ
           strsql = "select SessionID as ����,TestIndex as ���,Name as ��Ŀ,CutOff ,DcmScad as ����" & _
                     " from Test " & _
                     "where  Sessionid=" & rsʱ��("����")
            Set rsTmp = cnAccess.Execute(strsql)
            If rsTmp.EOF Then
                .Close
                Exit Function
            End If
            
            Do While Not rsTmp.EOF
            '��ȡ��Ŀ
            str��Ŀ = rsTmp("��Ŀ")
            '��ȡ��������
            str���� = Format(rsʱ��("����"), "yyyy-mm-dd hh:mm:ss")
            '��ȡ����Ŀ��cutoffֵ
            strcutoff = "" & rsTmp("cutoff")
            '׼����ʼ��ȡ���
              strsql = "select Type as ˳���,OD,Result as ���" & _
                       " from Pozzetti" & _
                       " where  SessionID=" & rsTmp("����") & " and TestIndex=" & rsTmp("���")
              Set rs��Ŀ = cnAccess.Execute(strsql)
                  Do While Not rs��Ŀ.EOF
                     If IsNumeric(rs��Ŀ("˳���")) Then
                        lngResultCount = lngResultCount + 1
                        ReDim Preserve aResults(lngResultCount)
                        If str��Ŀ = "HBeAB" Or str��Ŀ = "HBcAB" Then
                           If rs��Ŀ("OD") < strcutoff Then
                              str���Խ�� = "����(+)"
                           Else
                            str���Խ�� = "����(-)"
                           End If
                         Else
                           If rs��Ŀ("OD") > strcutoff Then
                              str���Խ�� = "����(+)"
                           Else
                            str���Խ�� = "����(-)"
                           End If
                         End If
                        aResults(lngResultCount) = str����  '����ʱ��
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rs��Ŀ("˳���"))   '�������
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                           " "  '������
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          "ѪҺ"  '�걾
                       aResults(lngResultCount) = aResults(lngResultCount) & "|0" '�ʿ�
                       aResults(lngResultCount) = aResults(lngResultCount) & "|" & str��Ŀ & "|" & str���Խ�� '���
                       aResults(lngResultCount) = aResults(lngResultCount) & _
                                     "^" & Val("" & rs��Ŀ("���")) & "^" & strcutoff & "^" & Val(rs��Ŀ("OD")) 'cutoff odֵ
    
                     End If
                     rs��Ŀ.MoveNext
                  Loop
              rsTmp.MoveNext
            Loop
            rsʱ��.MoveNext
      Loop
    End With
    ResultFromFile = aResults
End Function





