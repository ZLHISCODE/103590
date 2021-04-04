VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_gc5600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    Dim lngResultCount As Long, strBh As String, strDate As String
    Dim strChenel As String, strData As String
    Dim blnUnion As Boolean
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "": blnUnion = False
            strSQL = "Select shearrate,���,����,ʱ��,visc,Ѫ��ճ��,ѹ��,Ѫ��,ȫѪ�������ָ��,ȫѪ�������ָ��,Ѫ������Kֵ,��ϸ���ۼ�ָ��,��ϸ������ָ��," & _
                     " ȫѪ���л�ԭճ��,ȫѪ���л�ԭճ��,��ϸ������ָ��TK,��ϸ������ָ��,����ճ��,��ά����ԭ,Ѫ���̴�,������֬,����֬����,Ѫ��," & _
                    " ѪС��ճ����,����Ѫ˨����,��ϸ����Ӿ,ѪС��ۼ���,����Ѫ˨���� " & _
                        "FROM testdata  b,visc a  where  a.testdataid=b.testdataid " & _
                     IIf(lngSampleNO = -1, " And Format(b.����,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " And Format(����,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                    " And B.���=" & lngSampleNO) & _
                    " order by ���� desc,���"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            '��������±����
            ReDim aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        lngResultCount = -1
        Do While Not rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            strDate = Format(rsTmp("����"), "yyyy-mm-dd")
            strDate = strDate & " " & rsTmp("ʱ��")
            aResults(lngResultCount) = strDate '����ʱ��
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val(rsTmp("���")) '�������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                " " '������
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                "ѪҺ" '�걾
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                "0" '�ʿ�
                
            Select Case rsTmp("shearrate")
            Case 1
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.1|" & rsTmp("visc")
                If Not IsNull(rsTmp("Ѫ��ճ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|02|" & rsTmp("Ѫ��ճ��")
                If Not IsNull(rsTmp("ѹ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|03|" & rsTmp("ѹ��")
                If Not IsNull(rsTmp("Ѫ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|04|" & rsTmp("Ѫ��")
                If Not IsNull(rsTmp("ȫѪ�������ָ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|05|" & rsTmp("ȫѪ�������ָ��")
                If Not IsNull(rsTmp("ȫѪ�������ָ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|06|" & rsTmp("ȫѪ�������ָ��")
                If Not IsNull(rsTmp("Ѫ������Kֵ")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|07|" & rsTmp("Ѫ������Kֵ")
                If Not IsNull(rsTmp("��ϸ���ۼ�ָ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|08|" & rsTmp("��ϸ���ۼ�ָ��")
                If Not IsNull(rsTmp("��ϸ������ָ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|09|" & rsTmp("��ϸ������ָ��")
                If Not IsNull(rsTmp("ȫѪ���л�ԭճ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|10|" & rsTmp("ȫѪ���л�ԭճ��")
                If Not IsNull(rsTmp("ȫѪ���л�ԭճ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|11|" & rsTmp("ȫѪ���л�ԭճ��")
                If Not IsNull(rsTmp("��ϸ������ָ��TK")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|12|" & rsTmp("��ϸ������ָ��TK")
                If Not IsNull(rsTmp("��ϸ������ָ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|13|" & rsTmp("��ϸ������ָ��")
                If Not IsNull(rsTmp("����ճ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|14|" & rsTmp("����ճ��")
                If Not IsNull(rsTmp("��ά����ԭ")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|15|" & rsTmp("��ά����ԭ")
                If Not IsNull(rsTmp("Ѫ���̴�")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|16|" & rsTmp("Ѫ���̴�")
                If Not IsNull(rsTmp("������֬")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|17|" & rsTmp("������֬")
                If Not IsNull(rsTmp("����֬����")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|18|" & rsTmp("����֬����")
                If Not IsNull(rsTmp("Ѫ��")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|19|" & rsTmp("Ѫ��")
                If Not IsNull(rsTmp("ѪС��ճ����")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|20|" & rsTmp("ѪС��ճ����")
                If Not IsNull(rsTmp("����Ѫ˨����")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|21|" & rsTmp("����Ѫ˨����")
                If Not IsNull(rsTmp("��ϸ����Ӿ")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & rsTmp("��ϸ����Ӿ")
                If Not IsNull(rsTmp("ѪС��ۼ���")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|23|" & rsTmp("ѪС��ۼ���")
                If Not IsNull(rsTmp("����Ѫ˨����")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|24|" & rsTmp("����Ѫ˨����")
            Case 5
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.2|" & rsTmp("visc")
            Case 30
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.3|" & rsTmp("visc")
            Case 200
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.4|" & rsTmp("visc")
            End Select
           
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function