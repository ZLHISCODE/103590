Attribute VB_Name = "mdlRecipeAuditEx"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gbytClass As Byte
Public gblnInit As Boolean
Public grsCheckItems As ADODB.Recordset
Public gstrIDs As String

'�˱���Ҫ�롰���������Ŀ����ı���ֵ��ͬ�������޷����ö�Ӧ�Ĺ��ܺ���
Public Const GSTR_CODE_��ҩע���  As String = "D01"
'Public Const GSTR_CODE_�������Ŀ As String = "...."

Public Function F_��ҩע���(ByRef strMedicalID As String, ByRef strErr As String) As Boolean
'���ܣ���鴫���һ��ҽ���У��Ƿ�������֣������֣����ϵ���ҩע���
'������
'  strMedicalID��ʵ�Σ����ϸ�/���ϸ��ҽ��ID
'  strErr��ʵ�Σ����쳣��Ϣ
'���أ�True�ϸ�False���ϸ�

    Dim i As Integer, j As Integer
    Dim arrTmp As Variant
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strIDs As String
    Dim intCount As Integer
    
    strIDs = gstrIDs
    
    arrTmp = Split(strIDs, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Trim(arrTmp(i)) <> "" Then j = j + 1
    Next
    If j = 1 Then
        F_��ҩע��� = True
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '�����ҩע���ҽ��
    If Right(strIDs, 1) = "," Then strIDs = Left(strIDs, Len(strIDs) - 1)

    '��ǰ�ύ��ҩ���Ƿ�������ֺ���������
    strSQL = "Select a.ID " & _
             "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ���� C, Table(f_Num2list([1], ',')) D " & _
             "Where a.���Id = d.Column_Value And a.������Ŀid = b.Id And b.Id = c.ҩ��id And b.��� = '6' " & _
             "  And c.ҩƷ���� Like '%ע���%' And ROWNUM < 3 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��֤��ҩע���1", strIDs)
    If rsTemp.RecordCount >= 2 Then
        '��ҩע���ҽ������һ��
        rsTemp.MoveLast
        strMedicalID = CStr(rsTemp!ID)
        F_��ҩע��� = False
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    '�ټ��24Сʱ�ڵ�ҩ���Ƿ�������ֺ���������
'    strSQL = "Select a.id " & vbNewLine & _
'             "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ���� C," & vbNewLine & _
'             "     (Select a.����id " & vbNewLine & _
'             "      From ����ҽ����¼ A, Table(f_Num2list([1], ',') ) B " & vbNewLine & _
'             "      Where a.���id = b.Column_Value And Rownum < 2) D " & vbNewLine & _
'             "Where a.����id = d.����id And a.������Ŀid = b.Id And b.Id = c.ҩ��id And b.��� = '6' And c.ҩƷ���� Like '%ע���%' " & vbNewLine & _
'             "    And a.����ʱ�� >= Sysdate - 1 And Rownum < 3 "
    strSQL = "Select a.Id " & vbNewLine & _
             "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ���� C, ����ҽ����¼ D, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "Where a.����id = d.����id And d.���id = e.Column_Value And a.������Ŀid = b.Id And b.Id = c.ҩ��id And b.��� = '6' " & vbNewLine & _
             "    And c.ҩƷ���� Like '%ע���%' And a.����ʱ�� >= Sysdate - 1 And Rownum < 3 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��֤��ҩע���2", strIDs)
    If rsTemp.RecordCount >= 2 Then
        rsTemp.MoveLast
        strMedicalID = CStr(rsTemp!ID)
        F_��ҩע��� = False
    Else
        strMedicalID = ""
        F_��ҩע��� = True
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    strMedicalID = ""
    If zl9ComLib.ErrCenter() = 1 Then
        Resume
    Else
        strErr = Err.Description
    End If
End Function

'Public Function F_�·���(...) As Boolean
''���ܣ�
''������
''���أ�
'End Function
