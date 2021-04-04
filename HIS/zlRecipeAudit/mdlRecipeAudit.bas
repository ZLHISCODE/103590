Attribute VB_Name = "mdlRecipeAudit"
Option Explicit

Public gobjRecipeAuditEx As Object

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
'���أ�True�ɹ���Falseʧ��
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.���� = UserInfo.�û���
    Set rsTmp = SYS.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = zlCommFun.NVL(rsTmp!����ID, 0)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            UserInfo.�û��� = rsTmp!�û���
            GetUserInfo = True
        End If
        rsTmp.Close
    End If
End Function

Public Function AuditDrug(ByVal fldItem As Fields, ByVal lngPatientID As Long, _
    ByVal bytMedicalClass As Byte, ByVal lngBillID As Long, _
    ByVal strSubmitID As String, ByRef strMedicalID As String) As Byte
'���ܣ����ҩ��
'������
'  fldItem�����������Ŀ
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strSubmitID�����clsBusiness.AutoAudit������strSubmitID����˵������ҩ;��ҽ��ID��
'  strMedicalID(ʵ��)�����ϸ�ҽ��ID�������ֵ��ʾ����ҽ����ҩƷҽ��ID��
'���أ�0-�쳣/δ֪��1-�ϸ�2-���ϸ�
    
    Dim strID As String, strIDs As String, strReturn As String, strErr As String
    Dim objRecipeAuditEx As Object
    Dim l As Long

    On Error GoTo errHandle
    
    strIDs = GetMedicalID(strSubmitID)  '�����IDת����ҽ��ID
    strID = strSubmitID                 '
    
    Select Case UCase(fldItem!����)
        Case "A01"          'Ƥ��
            strMedicalID = RAI_AllergicTest(lngPatientID, bytMedicalClass, lngBillID, strIDs)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "A02"
        Case "A03"
        Case "A04"
        Case "A05", "2-7"  '�ظ���ҩ
            strMedicalID = RAI_RepeatDrug(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "A06"
        Case "A07"
        Case "1-4"         '��������Ӥ�׶�δд���ա�����
            If bytMedicalClass = 2 Then
                'ֻ��סԺ�����������
                AuditDrug = RAI_InfantAge(lngPatientID, lngBillID)
            Else
                AuditDrug = 1
            End If
            
        Case "1-9"         '�����޸�δǩ����ҩƷ����δע��ԭ��
            strMedicalID = RAI_OverloadExplain(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "1-10"         'δд�ٴ���ϻ���д��ȫ
            AuditDrug = RAI_Diagnosis(lngPatientID, bytMedicalClass, lngBillID, strID)
            
        Case "1-14"         'δ������ҩ�������
            strMedicalID = RAI_AntibiosisManage(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "C01"          'PASS���
            strMedicalID = RAI_PASS(lngPatientID, bytMedicalClass, lngBillID, NVL(fldItem!PASS���), strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case Else           '�Զ��������Ŀ
        
            If fldItem!��� = 4 Then
                If gobjRecipeAuditEx Is Nothing Then
                    On Error Resume Next
                    Set gobjRecipeAuditEx = CreateObject("zlRecipeAuditEx.clsRecipeAuditEx")
                    If gobjRecipeAuditEx Is Nothing Then
                        Err.Clear
                        gstrErrInfo = gstrErrInfo & vbCr & "������zlRecipeAuditEx������ʧ��"
                        Exit Function
                    End If
                    Err.Clear: On Error GoTo errHandle
                End If
                If gobjRecipeAuditEx.Init(gcnOracle, bytMedicalClass, strSubmitID) Then
                    If gobjRecipeAuditEx.Check(UCase(fldItem!����), strReturn, strErr) Then
                        '�ϸ�
                        AuditDrug = 1
                    Else
                        '�з���ҽ��IDΪ���ϸ񣬷�֮û�м����������δ���
                        AuditDrug = IIf(strReturn = "", 0, 2)
                    End If
                End If
            End If
            
    End Select
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "AuditDrug��" & vbCr & Err.Description
End Function

Private Function RAI_PASS(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strAuditPASS As String, ByVal strID As String) As String
'���ܣ�PASS���
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strAuditPASS��Ҫ����PASS�����ֵ
'  strID��ҽ��ID��     ��ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID]
'���أ����ϸ��ҽ��ID�� ��ʽ��ҽ��ID[;ҽ��ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    If strAuditPASS = "" Then
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select a.Id " & vbCr & _
             "From ����ҽ����¼ A, Table(f_Num2list([1], ',')) B, Table(f_Num2list([2], ';')) C " & vbCr & _
             "Where a.���Id = b.Column_Value And a.����� = c.Column_Value And a.����id = [3] And a.������� in ('5','6','7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����PASS�����ֵ", strID, strAuditPASS, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString("[1];", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
        .Close
    End With
    
    RAI_PASS = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_PASS��" & vbCr & Err.Descriptions
End Function

Private Function RAI_AntibiosisManage(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'���ܣ�����ҩƷ�Ƿ񰴹�����
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strID����ҩ;��ҽ��ID��   ��ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID]
'���أ����ϸ��ҽ��ID��      ��ʽ��ҽ��ID[;ҽ��ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '����Ƿ�������ʹ�õĿ���ҩ��
    strSQL = "Select a.Id " & vbCr & _
             "From ����ҽ����¼ A, ҩƷ���� B, Table(f_Num2list([1], ',')) C " & vbCr & _
             "Where a.������Ŀid = b.ҩ��id And a.���Id = c.Column_Value And b.������ > 1 " & vbCr & _
             "    And a.����id = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ҩ��", strID, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString(";[1]", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        .Close
    End With

    '������ʹ�õĿ���ҩҩ��
    If strReturn <> "" Then
        If Val(zlDatabase.GetPara("����ҩ��ּ�����", glngSys)) <> 1 Then
            'δ���ÿ���ҩ��ּ���������strReturnȷ���Ƿ�ϸ�
            RAI_AntibiosisManage = strReturn
        Else
            '�����ÿ���ҩ��ּ�������ʾ�ϸ�
            RAI_AntibiosisManage = ""
        End If
    End If
    
    Exit Function

errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_AntibiosisManage��" & vbCr & Err.Descriptions
End Function

Private Function RAI_Diagnosis(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As Byte
'���ܣ�δд�ٴ���ϻ���д��ȫ
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strID����ҩ;��ҽ��ID��  ��ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID]
'���أ�0-�쳣/δ֪��1-�ϸ�2-���ϸ�

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select Sum(Rec) Rec " & vbNewLine & _
             "From (Select Count(1) Rec " & vbNewLine & _
             "      From ������ϼ�¼ C, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "      Where c.ҽ��id = e.Column_Value And c.����id = [2] And c.������� Is Not Null And Rownum < 2 " & vbNewLine & _
             "      Union All " & vbNewLine & _
             "      Select Count(1) Rec" & vbNewLine & _
             "      From ������ϼ�¼ C, �������ҽ�� D, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "      Where e.Column_Value = d.ҽ��id And d.���id = c.Id And c.����id = [2] And c.������� Is Not Null And Rownum < 2 " & vbNewLine & _
             "      Union All " & vbNewLine & _
             "      Select Count(1) Rec From ������ϼ�¼ Where ������� Is Not Null And ����id = [2] And ��ҳid = [3] " & vbNewLine & _
             ") A "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����Ӧ�������", strID, lngPatientID, lngBillID)
    If rsTemp!Rec <= 0 Then
        RAI_Diagnosis = 2
    Else
        RAI_Diagnosis = 1
    End If
    rsTemp.Close
    
    Exit Function

errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_Diagnosis��" & vbCr & Err.Description
End Function

Private Function RAI_OverloadExplain(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'���ܣ�ҩƷ����˵��
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strID����ҩ;��ҽ��ID��  ��ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID]
'���أ����ϸ��ҽ��ID       ��ʽ��ҽ��ID[;ҽ��ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select a.ID " & vbCr & _
             "From ����ҽ����¼ A, ҩƷ���� B, Table(f_Num2list([1], ',')) C " & vbCr & _
             "Where a.������Ŀid = b.ҩ��id And a.���Id = c.Column_Value And a.������� In ('5', '6', '7') " & vbCr & _
             "  And Nvl(a.�ܸ�����, 0) > Nvl(b.��������, 0) And Nvl(b.��������, 0) > 0 " & vbCr & _
             "  And a.����˵�� is null And a.����ID = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҩƷ����", strID, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString("[1];", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
        .Close
    End With
    
    RAI_OverloadExplain = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_OverloadExplain��" & vbCr & Err.Descriptions
End Function

Private Function RAI_AllergicTest(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'���ܣ�ҩƷ��������
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strID��ҽ��ID��      ��ʽ��ҽ��ID[,ҽ��ID]
'���أ����ϸ��ҽ��ID    ��ʽ��ҽ��ID[;ҽ��ID]

    Dim arrID As Variant
    Dim l As Long
    Dim strReturn As String
    Dim intResult As Integer
    
    On Error GoTo errHandle
    
    If gobjPubAdvice Is Nothing Then Exit Function
    
    arrID = Split(strID, ",")
    For l = LBound(arrID) To UBound(arrID)
        '����zlPublicAdvice.CheckAdviceSkinResult��Ƥ�Ժ���
        intResult = gobjPubAdvice.CheckAdviceSkinResult(Val(arrID(l)))
        '-1��ʾ����Ƥ�Ի����ԣ�0��ʾ��δ���Ƥ�Խ����δ�´�Ƥ��ҽ����1��ʾ���ԣ�2��ʾ����
        If intResult = 0 Or intResult = 2 Then
            strReturn = strReturn & arrID(l) & ";"
        End If
    Next
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    
    RAI_AllergicTest = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_AllergicTest��" & vbCr & Err.Description
End Function

Private Function RAI_RepeatDrug(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'���ܣ�����ظ���ҩ
'������
'  lngPatientID������ID
'  bytMedicalClass�������ٴ����1-���2-סԺ
'  lngBillID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID
'  strID��ҽ��ID��      ��ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID]
'���أ����ϸ��ҽ��ID    ��ʽ��ҽ��ID[;ҽ��ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    
    '��鵱ǰ����ҽ����ҩ���Ƿ��ظ�
    strSQL = "Select Max(a.Id) ҽ��id, a.������Ŀid " & vbCr & _
             "From ����ҽ����¼ A, Table(f_Num2list([1], ',')) B" & vbCr & _
             "Where a.���Id = b.Column_Value And a.������� In ('5', '6', '7') " & vbCr & _
             "Group By a.������Ŀid " & vbCr & _
             "Having Count(a.������Ŀid) > 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ظ�ҩ��", strID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlStr.FormatString("[1];", rsTemp!ҽ��ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If strReturn <> "" Then
        strReturn = Left(strReturn, Len(strReturn) - 1)
        RAI_RepeatDrug = strReturn
        Exit Function
    End If
    
    strReturn = ""
    
    '�ټ�鵱ǰ����ҽ����ҩƷ�롰���͡�ͨ����ҩƷ��24Сʱ�ڣ��Ƿ��ظ�
    strSQL = "Select b.ҽ��id " & vbCr & _
             "From (Select Distinct a.������Ŀid " & vbCr & _
             "      From ����ҽ����¼ A, ����ҽ������ B " & IIf(bytMedicalClass = 1, ", ���˹Һż�¼ C ", " ") & vbCr & _
             "      Where a.Id = b.ҽ��id " & IIf(bytMedicalClass = 1, " And a.�Һŵ� = c.No ", " ") & "And a.������� In ('5', '6', '7') " & vbCr & _
             "          And a.����id = [1]" & IIf(bytMedicalClass = 1, " And c.Id = [2] ", " And a.��ҳID = [2] ") & _
             "          And b.����ʱ�� >= Sysdate - [3] / 24 ) A, " & vbCr & _
             "     (Select a.������Ŀid, a.Id ҽ��id " & vbCr & _
             "      From ����ҽ����¼ A, Table(f_Num2list([4], ',')) B " & vbCr & _
             "      Where a.���Id = b.Column_Value And a.������� In ('5', '6', '7') ) B " & vbCr & _
             "Where a.������Ŀid = b.������Ŀid "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ѷ��͵��ظ�ҩ��", lngPatientID, lngBillID, gintHoursRecipe, strID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlStr.FormatString("[1];", rsTemp!ҽ��ID)
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If strReturn <> "" Then
        strReturn = Left(strReturn, Len(strReturn) - 1)
        RAI_RepeatDrug = strReturn
        Exit Function
    End If
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_RepeatDrug��" & vbCr & Err.Description
End Function

Private Function RAI_InfantAge(ByVal lngPatientID As Long, ByVal lngMasterPageID As Long) As Byte
'���ܣ������������Ӥ�׶���д�ա����䣨סԺ�ࣩ
'������
'  lngPatientID������ID
'  lngMasterPageID��סԺ����Ϊ��ҳID
'���أ�0-�쳣/δ֪��1-�ϸ�2-���ϸ�

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select Count(1) Rec " & vbCr & _
             "From ������������¼ " & vbCr & _
             "Where ����id = [1] And ��ҳid = [2] And ����ʱ�� Is Null "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������Ӥ�׶�������", lngPatientID, lngMasterPageID)
    If rsTemp!Rec > 0 Then
        '�м�¼����������û����д����ʱ��
        RAI_InfantAge = 2
    Else
        '�޼�¼����û����������¼����������д����ʱ��
        RAI_InfantAge = 1
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_InfantAge��" & vbCr & Err.Description
End Function

'Public Function GetStoreID(ByVal lngID As Long) As Long
''���ܣ�ͨ��ҽ��ID��ȡҽ����ִ�п��ң���ҩҩ����ID
''������
''  lngID��ҽ��ID
''���أ���ҩҩ��ID
'
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    strSQL = "Select ִ�п���ID from ����ҽ����¼ where ID = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡִ�п���ID", lngID)
'    If rsTemp.RecordCount = 1 Then
'        GetStoreID = NVL(rsTemp!ִ�п���ID, 0)
'    End If
'    rsTemp.Close
'
'errHandle:
'    If ErrCenter = 1 Then Resume
'End Function

Public Function GetCalorie(ByVal lngPatientID As Long, ByVal lngRegisterID As Long, ByVal lngPageID As Long) As String
'���ܣ���ȡ����������Ҫ��
'������
'  lngPatientID������ID
'  lngRegisterID���Һŵ�ID
'  lngPageID����ҳID
'���أ�������Ҫ����ʽ��ֵ���磺66.5 + 13.8 * 61KG + 5.0 * 172CM - 6.8 * 30�� = 1564.30��

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lngPageID <= 0 Then
        '����
        strSQL = "Select zl_fun_pati_calorie([1], Null, [2]) ���� From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������Ҫ��", lngPatientID, lngRegisterID)
    Else
        'סԺ
        strSQL = "Select zl_fun_pati_calorie([1], [2], Null) ���� From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������Ҫ��", lngPatientID, lngPageID)
    End If
    
    If rsTemp.EOF = False Then
        GetCalorie = NVL(rsTemp!����)
    End If
    
    Exit Function

errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowReason(ByVal frmOwner As Form, ByVal strSQL As String, ByRef blnCancel As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ���������ѡ����
'������
'  frmOwner�������������
'  strSQL��SQL��ѯ
'  blnCancel��ʵ�Σ���Trueѡ��ȷ�ϣ�Falseѡ��ȡ��
'���أ���ѡ��ļ�¼

    Dim frmSelector As New frmReasonSelector

    Set ShowReason = frmSelector.ShowMe(frmOwner, strSQL, blnCancel, arrInput)
    
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte, ByVal vsfVar As VSFlexGrid, ByVal strTitle As String)
'-------------------------------------------------
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim lngRow As Long
    Dim lngColor As Long

    lngColor = vsfVar.GridColor
    vsfVar.GridColor = vbBlack

    lngRow = vsfVar.Row
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTitle
        
    objRow.Add strRange
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(SYS.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfVar
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    vsfVar.Row = lngRow
    vsfVar.GridColor = lngColor
End Sub

Public Function SendMessage(ByVal bytResult As Byte, ByVal lngAuditID As Long, _
    ByVal bytMode As Byte, ByRef objMIP As clsMipModule, _
    Optional ByVal blnSendBeforeAudit As Boolean = False) As Boolean
'���ܣ�������Ϣ֪ͨ����ҽ��
'������
'  bytResult��1-���ϸ�2-��鲻�ϸ�11-���ϸ��Զ�����ʧ��
'  lngAuditID����ID
'  bytMode��1-���2-סԺ
'  objMIP����Ϣƽ̨����
'  blnSendBeforeAudit��bytMode=1����Ч��True-ҽ������ǰ�󷽣�False-ҩ���䷢ҩǰ��
'���أ�True�ɹ���Falseʧ��

    Const STR_OUT_PATI As String = "ZLHIS_RECIPEAUDIT_001"
    Const STR_IN_PATI  As String = "ZLHIS_RECIPEAUDIT_002"
    
    Dim strOutIn As String, strXML As String, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim blnMIP As Boolean
    
    If objMIP Is Nothing Then
        blnMIP = False
    Else
        blnMIP = objMIP.IsConnect
    End If

    If bytMode = 1 Then
        strOutIn = STR_OUT_PATI     '����
    Else
        strOutIn = STR_IN_PATI      'סԺ
    End If
    
    On Error GoTo errHandle
    
    'XML�ṹ������
    strSQL = "Select c.����id, c.����, c.סԺ��, c.�����, b.������Դ, Decode(b.������Դ, 1, d.Id, 2, b.��ҳid, Null) ����id, " & _
             "    c.��ǰ����id, e1.���� ��ǰ����, c.��ǰ����id, e2.���� ��ǰ����, c.��ǰ����, b.����ҽ��, " & _
             "    to_char(a.���ʱ��, 'yyyy-mm-dd hh24:mi:ss') ���ʱ��, a.�����, a.�����, a.Ids " & vbNewLine & _
             "From (Select Max(a.ҽ��id) ҽ��id, b.���ʱ��, User �����id, b.�����, b.�����," & _
             "          f_List2str(Cast(Collect(Cast(a.ҽ��id As Varchar2(20))) As t_Strlist)) Ids " & vbNewLine & _
             "      From ���������ϸ A, ��������¼ B " & vbNewLine & _
             "      Where a.��id = b.Id And a.��id = [1] " & vbNewLine & _
             "      Group By b.���ʱ��, User, b.�����, b.�����" & vbNewLine & _
             "     ) A, ����ҽ����¼ B, ������Ϣ C, ���˹Һż�¼ D, ���ű� E1, ���ű� E2 " & vbNewLine & _
             "Where a.ҽ��id = b.Id And b.����id = c.����id And b.�Һŵ� = d.No(+) And c.��ǰ����id = E1.Id(+) And c.��ǰ����id = E2.Id(+) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ˡ�ҽ��������Ϣ", lngAuditID)
    If rsTemp.EOF = False Then
        '��װXML
        
        '������Ϣ
        objXML.AppendNode "patient_info"
        objXML.AppendData "patient_id", NVL(rsTemp!����ID)
        objXML.AppendData "patient_name", NVL(rsTemp!����)
        objXML.AppendData "in_number", NVL(rsTemp!סԺ��)
        objXML.AppendData "out_number", NVL(rsTemp!�����)
        objXML.AppendNode "patient_info", True
        
        'ҽ����Ϣ
        objXML.AppendNode "patient_clinic"
        objXML.AppendData "patient_source", NVL(rsTemp!������Դ)
        objXML.AppendData "clinic_id", NVL(rsTemp!����id)
        objXML.AppendData "clinic_area_id", NVL(rsTemp!��ǰ����id)
        objXML.AppendData "clinic_area_title", NVL(rsTemp!��ǰ����)
        objXML.AppendData "clinic_dept_id", NVL(rsTemp!��ǰ����id)
        objXML.AppendData "clinic_dept_title", NVL(rsTemp!��ǰ����)
        objXML.AppendData "clinic_room", ""
        objXML.AppendData "clinic_bed", NVL(rsTemp!��ǰ����)
        objXML.AppendNode "patient_clinic", True
        
        '����Ϣ
        objXML.AppendNode "recipe_audit_info"
        objXML.AppendData "create_doctor_name", NVL(rsTemp!����ҽ��)
        objXML.AppendData "ra_time", NVL(rsTemp!���ʱ��)
        objXML.AppendData "ra_chemist_id", UserInfo.ID
        objXML.AppendData "ra_chemist_name", NVL(rsTemp!�����)
        objXML.AppendData "ra_result", NVL(rsTemp!�����)
        objXML.AppendData "ra_sent", IIf(bytResult = 11, 1, 0)
        objXML.AppendData "order_ids", NVL(rsTemp!ids)
        objXML.AppendNode "recipe_audit_info", True
        
        strXML = objXML.XmlText
        
        objXML.ClearXmlText
        Set objXML = Nothing
    End If
    rsTemp.Close
    
    If strXML = "" Then
        MsgBox "��������ҽ��������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.����ҽ������վ�����͡�ǰ�����̣��ϸ��벻�ϸ���Ϣ����Ҫ���͸�ҽ��
    '2.����/סԺ�䷢ҩ�����̣���鲻�ϸ���Ϣ֪ͨҽ��
    If blnSendBeforeAudit Then
        'ͨ���ٴ�����Ϣ���Ʒ���
        If zlDatabase.SendMsg(strOutIn, strXML) = False Then
            MsgBox "��ҽ��������Ϣʧ�ܣ�", vbInformation, gstrSysName
        End If
        '��������ǰ��չ��
        If blnMIP Then
            'ͨ����Ϣƽ̨����
            If objMIP.CommitMessage(strOutIn, strXML) = False Then
                MsgBox "��ҽ��������Ϣʧ�ܣ�", vbInformation, gstrSysName
            End If
        End If
    Else
        '�䷢ҩǰ��չ��
        If bytResult = 2 Then
            'ͨ���ٴ�����Ϣ���Ʒ���
            If zlDatabase.SendMsg(strOutIn, strXML) = False Then
                MsgBox "��ҽ��������Ϣʧ�ܣ�", vbInformation, gstrSysName
            End If
            'ͨ����Ϣƽ̨����
            If blnMIP Then
                If objMIP.CommitMessage(strOutIn, strXML) = False Then
                    MsgBox "��ҽ��������Ϣʧ�ܣ�", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    
    SendMessage = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub PassResultView(ByRef objPASS As Object, ByVal blnOutPatient As Boolean, ByVal lngMedicalID As Long)
'���ܣ��鿴ָ��ҩ����PASS�����
'������
'  objPASS��PASS�ӿڶ���
'  blnOutPatient��True���ﲡ�ˣ�FalseסԺ����
'  lngMedicalID��ҩ��ID

    If objPASS Is Nothing Then Exit Sub
    
    Dim strSQL As String, strNO As String
    Dim lngPageID As Long
    Dim lngPatientID As Long
    Dim rsTemp As ADODB.Recordset
    
    '��ȡ������Ϣ
    On Error GoTo hErr
    strSQL = "Select ����id, ��ҳid, �Һŵ� From ����ҽ����¼ Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ͨ��ҽ��ID����ȡ������Ϣ", lngMedicalID)
    If rsTemp.EOF = False Then
        lngPatientID = zl9ComLib.NVL(rsTemp!����ID, 0)
        If blnOutPatient Then
            strNO = zl9ComLib.NVL(rsTemp!�Һŵ�)
        Else
            lngPageID = zl9ComLib.NVL(rsTemp!��ҳID, 0)
        End If
    End If
    rsTemp.Close
    
    '������ҩ���¼���󣬲��ܲ鿴�����
    On Error Resume Next: Err.Clear
    If blnOutPatient Then
        Call objPASS.zlPassRecipelCheck(lngPatientID, 0, strNO, CStr(lngMedicalID))
    Else
        Call objPASS.zlPassRecipelCheck(lngPatientID, lngPageID, "", CStr(lngMedicalID))
    End If
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    
    '�鿴������ҩ�ļ����
    On Error Resume Next: Err.Clear
    Call objPASS.zlPassShowWarn_YF(CStr(lngMedicalID))
    If Err.Number <> 0 Then Err.Clear
    
    Exit Sub

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Public Function GetRecipeAuditBills(ByVal bytType As Byte) As Boolean
'���ܣ������������סԺ�ġ���������¼���Ƿ����δ���ļ�¼
'������
'  bytType��0-�����������סԺ��1-���2-סԺ
'���أ�True����δ���ļ�¼��False������δ���ļ�¼

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If bytType = 1 Then
        '����
        strSQL = "Select ID From ��������¼ Where ״̬ = 0 And �ύʱ�� >= Trunc(Sysdate - [1]) And �Һ�Id Is Not Null And Rownum < 2 "
    ElseIf bytType = 2 Then
        'סԺ
        strSQL = "Select ID From ��������¼ Where ״̬ = 0 And �ύʱ�� >= Trunc(Sysdate - [1]) And ��ҳId Is Not Null And Rownum < 2 "
    Else
        '�����������סԺ
        strSQL = "Select ID From ��������¼ Where ״̬ = 0 And �ύʱ�� >= Trunc(Sysdate - [1]) And Rownum < 2 "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���δ���Ĵ�������¼", IIf(bytType = 1, 2, 4))
    GetRecipeAuditBills = rsTemp.EOF = False
    rsTemp.Close
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Function GetMedicalID(ByVal strParentID As String) As String
'���ܣ�����ҩ;��IDת����ҽ��ID
'������
'  strParentID����ҩ;��ID����ʽ��ҽ��ID1[,ҽ��ID2[,...]]
'���أ�ҽ��ID

    Dim strSQL As String, strResult As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strSQL = "Select a.Id From ����ҽ����¼ A, Table(Cast(f_Str2list([1]) As t_Strlist)) B Where a.���id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ҩ;��IDת����ҽ��ID", strParentID)
    With rsTemp
        Do While .EOF = False
            strResult = strResult & "," & !ID
            .MoveNext
        Loop
        .Close
        If strResult <> "" Then strResult = Mid(strResult, 2)
    End With
    
    GetMedicalID = strResult
    
    Exit Function
    
hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Public Function GetDiagnose(ByVal lngPatientID As Long, ByVal lngPageID As Long) As String
'���ܣ���ȡסԺ���˵����
'������
'   lngPatientID������ID
'   lngPageID����ҳID
'���أ��������

    Dim strSQL As String, str��� As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strSQL = "Select �Ƿ�����, ������� From ������ϼ�¼ Where ����id = [1] And ��ҳid = [2] And ������� Is Not Null Order By Rowid"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡסԺ���˵��������", lngPatientID, lngPageID)
    With rsTemp
        Do While .EOF = False
            str��� = str��� & "," & !������� & IIf(Val(!�Ƿ����� & "") = 1, "(��)", "")
            
            .MoveNext
        Loop
        .Close
    End With
    
    If str��� <> "" Then GetDiagnose = Mid(str���, 2)
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Public Sub DispCountNG(ByVal vsfCount As VSFlexGrid, ByVal lblDisp As Label)
'���ܣ�����ؼ��в��ϸ�������Ŀ����������ʾ��Label��
'������
'  vsfCount��Ҫ����Ŀؼ�
'  lblDisp��Ҫ��ʾ�Ŀؼ�

    Const STR_ITEMS As String = "�����Ŀ"
    Dim i As Integer, intCount As Integer
    
    With vsfCount
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ҩʦ���"))) = "���ϸ�" Then
                intCount = intCount + 1
            End If
        Next
    End With
    If intCount > 0 Then
        lblDisp = STR_ITEMS & zlStr.FormatString("������[1]��ϸ�", intCount)
    Else
        lblDisp = STR_ITEMS
    End If

End Sub

Public Function GetAuditResult(ByVal vsfVar As VSFlexGrid) As Boolean
'���ܣ����VSF�����Ŀ�С�ҩ����顱����Ƿ�������Ŀ���ϸ�
'���أ�True������Ŀ���ϸ�False�в��ϸ�

    Const STR_NAME As String = "ҩʦ���"
    Const STR_PASS As String = "�ϸ�"
    Dim i As Integer
    Dim blnAllPass As String
    
    With vsfVar
        blnAllPass = True
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex(STR_NAME)) <> STR_PASS Then
                blnAllPass = False
                Exit For
            End If
        Next
    End With
    
    GetAuditResult = blnAllPass
End Function

