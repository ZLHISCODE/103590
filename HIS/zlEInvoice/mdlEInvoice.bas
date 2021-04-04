Attribute VB_Name = "mdlEinvoice"
Option Explicit
Public gstrProductName As String, gstrSysName As String
Public gcnOracle As ADODB.Connection
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gobjEInvProviders As clsEInvProviders   '��ǰ�ṩ�߼�
Public gobjEinvProvider As clsEInvProvider  '��ǰʹ���ṩ��
Public glngInstanceCount As Long 'ʵ����
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '����ؼ�ʱ,ѡ����ʾ��ɫ
Public Const GRD_LOSTFOCUS_COLORSEL = &HE0E0E0  '&H80000010  '�뿪����ʱ,ѡ�����ʾ��ɫ

Public Sub InitEInvProviders()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿ��ṩ������
    '����:���˺�
    '����:2020-03-04 15:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim objEinvProvider As clsEInvProvider
    If Not gobjEInvProviders Is Nothing Then Exit Sub
    
    Set gobjEInvProviders = New clsEInvProviders
    strSQL = "Select ���,����,����,�Ƿ�����,����,������ From ����Ʊ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʊ�����")
    With rsTemp
        Do While Not .EOF
            Set objEinvProvider = gobjEInvProviders.Add(Val(Nvl(!���)), Nvl(!����), Val(Nvl(!�Ƿ�����)) = 1, Nvl(!����), True, "K" & zlStr.LPAD(Val(Nvl(!���)), 3, "0"))
            If objEinvProvider.�Ƿ����� Then Set gobjEinvProvider = objEinvProvider
            .MoveNext
        Loop
    End With
End Sub

Public Function GetPubEInvoiceObject(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    objPubEInvoice As Object, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ�ݹ����ӿڲ���
    '���:
    '   frmMain�����õ�������
    '   lngModule����ǰ����ģ���
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '   blnDeviceSet���豸���õ��õĳ�ʼ��
    '����:
    '����:��ʼ���ɹ�����true,���򷵻�False
    '˵��:
    '   1.ʹ�ñ�����ǰ,�����ȵ��ñ��ӿڽ��г�ʼ��
    '   2.��ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    If objPubEInvoice Is Nothing Then
        On Error Resume Next
        Set objPubEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "�����ڿ��õĵ���Ʊ�ݽӿڲ���(zlPublicExpense.clsPubEInvoice)������ϵͳ����Ա��ϵ����ϸ�Ĵ�����ϢΪ:" & vbCrLf & Err.Description, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If objPubEInvoice Is Nothing Then Exit Function
    
    GetPubEInvoiceObject = objPubEInvoice.zlInitialize(frmMain, byt����, gcnOracle, lngSys, lngModule, False, strExtend)
End Function

Public Function load��Ʊ��(cboControl As Object, ByRef rs��Ʊ�� As ADODB.Recordset, ByRef rs�շ�Ա As ADODB.Recordset)
    On Error GoTo ErrHandler
    cboControl.Clear
    If Get��Ʊ��(rs��Ʊ��) = False Then Exit Function
    If rs��Ʊ��.RecordCount > 0 Then
        Do While Not rs��Ʊ��.EOF
            cboControl.AddItem rs��Ʊ��!���� & "-" & rs��Ʊ��!����
            rs��Ʊ��.MoveNext
        Loop
        load��Ʊ�� = True: Exit Function
    End If
    
    If Load�շ�Ա(cboControl, rs�շ�Ա) = False Then Exit Function
    load��Ʊ�� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Load�շ�Ա(cboControl As Object, ByRef rs�շ�Ա As ADODB.Recordset)

    On Error GoTo ErrHandler
    cboControl.Clear
    If Get�շ�Ա(rs�շ�Ա) = False Then Exit Function
    
    Do While Not rs�շ�Ա.EOF
        cboControl.AddItem rs�շ�Ա!��� & "-" & rs�շ�Ա!����
        rs�շ�Ա.MoveNext
    Loop
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Select��Ʊ��(frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    cboControl As Object, rs��Ʊ�� As ADODB.Recordset) As Boolean
    'ģ�����ҿ�Ʊ��
    Dim lngCount As Long
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset, strAdded As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    
    '�ȸ��Ƽ�¼��
    On Error GoTo ErrHandler
    Set rsTemp = zlDatabase.zlCopyDataStructure(rs��Ʊ��)
    
    strText = cboControl.Text
    strCompents = strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    
    rs��Ʊ��.Filter = strFilter: lngCount = 0
    With rs��Ʊ��
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not rs��Ʊ��.EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.��������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01�����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ������������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strText Then strResult = Nvl(!����): lngCount = 0: Exit Do
                
                '1.��������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strText) Then
                    If lngCount = 0 Then strResult = Nvl(!����)
                    lngCount = lngCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(rs��Ʊ��!����) Like strText & "*" Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!����) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs��Ʊ��, rsTemp)
                        strAdded = strAdded & "," & Nvl(!����) & ","
                    End If
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ����
                    lngCount = lngCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!����) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs��Ʊ��, rsTemp)
                        strAdded = strAdded & "," & Nvl(!����) & ","
                    End If
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,�����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                    lngCount = lngCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(!����) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!����) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs��Ʊ��, rsTemp)
                        strAdded = strAdded & "," & Nvl(!����) & ","
                    End If
                End If
            End Select
            rs��Ʊ��.MoveNext
        Loop
    End With
    
    If lngCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
    'ֱ�Ӷ�λ
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckComBoxExists(cboControl, strResult, True) Then zlCommFun.PressKey vbKeyTab
        Select��Ʊ�� = True: Exit Function
    End If
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then
        'δ�ҵ�
        rsTemp.Close: Set rsTemp = Nothing
        Exit Function
    End If

    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "����"
    End Select
    
    '����ѡ����
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, lngSys, lngModule, cboControl, rsTemp, True, "", "ID", rsReturn) Then
        Call zlControl.ControlSetFocus(cboControl)
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '���ж�λ
                If CheckComBoxExists(cboControl, Nvl(rsReturn!����), True) Then zlCommFun.PressKey vbKeyTab
                rsTemp.Close: Set rsTemp = Nothing
                Select��Ʊ�� = True: Exit Function
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Select�շ�Ա(frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    cboControl As Object, rs�շ�Ա As ADODB.Recordset) As Boolean
    'ģ�������շ�Ա
    Dim lngCount As Long
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset, strAdded As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    
    '�ȸ��Ƽ�¼��
    On Error GoTo ErrHandler
    Set rsTemp = zlDatabase.zlCopyDataStructure(rs�շ�Ա)
    
    strText = cboControl.Text
    strCompents = strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    
    rs�շ�Ա.Filter = strFilter: lngCount = 0
    With rs�շ�Ա
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not rs�շ�Ա.EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strText Then strResult = Nvl(!����): lngCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strText) Then
                    If lngCount = 0 Then strResult = Nvl(!����)
                    lngCount = lngCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(rs�շ�Ա!���) Like strText & "*" Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs�շ�Ա, rsTemp)
                        strAdded = strAdded & "," & Nvl(!���) & ","
                    End If
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ����
                    lngCount = lngCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs�շ�Ա, rsTemp)
                        strAdded = strAdded & "," & Nvl(!���) & ","
                    End If
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                    lngCount = lngCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs�շ�Ա, rsTemp)
                        strAdded = strAdded & "," & Nvl(!���) & ","
                    End If
                End If
            End Select
            rs�շ�Ա.MoveNext
        Loop
    End With
    
    If lngCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
    'ֱ�Ӷ�λ
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckComBoxExists(cboControl, strResult, True) Then zlCommFun.PressKey vbKeyTab
        Select�շ�Ա = True: Exit Function
    End If
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then
        'δ�ҵ�
        rsTemp.Close: Set rsTemp = Nothing
        Exit Function
    End If

    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "����"
    End Select
    
    '����ѡ����
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, lngSys, lngModule, cboControl, rsTemp, True, "", "ID", rsReturn) Then
        Call zlControl.ControlSetFocus(cboControl)
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '���ж�λ
                If CheckComBoxExists(cboControl, Nvl(rsReturn!����), True) Then zlCommFun.PressKey vbKeyTab
                rsTemp.Close: Set rsTemp = Nothing
                Select�շ�Ա = True: Exit Function
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckComBoxExists(cboControl As Object, ByVal strText As String, _
    Optional ByVal blnLocateItem As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:���ڷ���true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    For i = 0 To cboControl.ListCount - 1
        If zlStr.NeedName(cboControl.List(i)) = strText Then
            If blnLocateItem Then cboControl.ListIndex = i
            CheckComBoxExists = True
            Exit Function
        End If
    Next
End Function

Private Function Get��Ʊ��(ByRef rs��Ʊ�� As ADODB.Recordset) As Boolean
    '���ؿ�Ʊ��
    Dim strSQL As String

    On Error GoTo ErrHandler
    strSQL = _
        " Select a.ID, a.����, a.����, a.����" & _
        " From ����Ʊ�ݿ�Ʊ�� A" & _
        " Where Nvl(A.����ʱ��, Sysdate + 1) > Sysdate And A.ĩ�� = 1" & _
        "           And (a.Ժ�� Is Null Or a.Ժ��='" & gstrNodeNo & "')"
    Set rs��Ʊ�� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ʊ������")
    Get��Ʊ�� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�շ�Ա(ByRef rs�շ�Ա As ADODB.Recordset) As Boolean
    '���ؿ�Ʊ��
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.Id, a.���, a.����, a.����" & _
        " From ��Ա�� A, ��Ա����˵�� B" & _
        " Where a.Id = b.��Աid And Nvl(a.����ʱ��, Sysdate + 1) > Sysdate" & _
        "           And b.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���','ҽ��','��ʿ')" & _
        "           And (a.վ�� Is Null Or a.վ��='" & gstrNodeNo & "')"
    Set rs�շ�Ա = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ�Ա����")
    Get�շ�Ա = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = (Forms.Count = 0)
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ
    '���:objPati-������Ϣ��
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-04 17:50:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'ʵ����Ϊ0ʱ���ŷ���Դ
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows   '�رմ���
    Err = 0: On Error Resume Next
    Set gobjEInvProviders = Nothing
    Set gobjEinvProvider = Nothing
    Set gcnOracle = Nothing
    
    zlReleaseResources = True
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False, Optional blnNotTran As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNotTran-����������
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    
    If blnNotTran = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function GetUserInfo(ByVal strDBUser As String) As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = strDBUser
    UserInfo.���� = strDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.�������� = "" & rsTmp!������
        UserInfo.���� = "" & rsTmp!����
        UserInfo.���� = "" & rsTmp!����
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEInvoiceData(ByVal bytҵ�񳡺� As Byte, ByVal dt��ʼʱ�� As Date, ByVal dt����ʱ�� As Date, _
    ByRef rsEInvoice As ADODB.Recordset, Optional ByVal bytƱ��״̬ As Byte, Optional ByVal bytʱ������ As Byte, _
    Optional ByVal bytQueryType As Byte, Optional ByVal varQueryValue As Variant, Optional ByVal str��Ʊ�� As String) As Boolean
    '��ȡ����Ʊ������
    '��Σ�
    '   bytҵ�񳡺� 0-���У�1-�շѣ�2-Ԥ����3-���ʣ�4-�Һţ�5-���￨
    '   bytƱ��״̬ 0-���У�1-������2-��죬3-��Ч
    '   bytʱ������ 0-Ʊ������ʱ�䣬1-����ʱ��
    '   bytQueryType ��ѯ���ͣ�0-���У�1-������ID��ѯ��2-�����õ��ݺŲ�ѯ��3-������Ʊ�ݺŲ�ѯ
    '   varQueryValue ��ѯ����ֵ���� bytQueryType ���ʹ��
    '   str��Ʊ�� ��Ʊ�����
    Dim strSQL As String, strWhere As String
    Dim strSqlSub As String
    
    On Error GoTo ErrHandler
    If bytʱ������ = 0 Then strWhere = strWhere & " And a.����ʱ�� Between [1] And [2]"
    
    Select Case bytƱ��״̬
    Case 0 '0-����
    Case 1 '1-����
        strWhere = strWhere & " And a.��¼״̬ In(1,3)"
    Case 2 '2-���
        strWhere = strWhere & " And a.��¼״̬ = 2"
    Case 3 '3-��Ч
        strWhere = strWhere & " And a.��¼״̬ = 1"
    End Select
    
    Select Case bytQueryType
    Case 0 '0-����
    Case 1 '1-������ID��ѯ
        strWhere = strWhere & " And a.����ID = [5]"
    Case 2 '2-�����õ��ݺŲ�ѯ
        strWhere = strWhere & " And b.NO = [5]"
    Case 3 '3-������Ʊ�ݺŲ�
        strWhere = strWhere & " And a.���� = [5] "
    End Select
    
    If str��Ʊ�� <> "" Then strWhere = strWhere & " And a.��Ʊ�� = [6]"
    
    '1)Ԥ����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 2 Then
        strSQL = _
            " Select a.ID, b.�տ�ʱ�� As �շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, b.No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��ID," & _
            "           To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As ��Ʊʱ��,a.�˿�ID,0 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,����Ԥ����¼ B" & _
            " Where a.����ID =b.ID And a.Ʊ��=2 And b.��¼����=1" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�տ�ʱ�� Between [3] And [4]")
        '����˿�
        strSQL = strSQL & " Union All " & _
            " Select a.ID, b.�տ�ʱ�� As �շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, b.No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��id," & _
            "           To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As ��Ʊʱ��,a.�˿�ID,0 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,����Ԥ����¼ B" & _
            " Where a.�˿�ID =b.ID And a.Ʊ��=2 And b.��¼����=11" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�տ�ʱ�� Between [3] And [4]")
    End If
    '2)���￨
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select a.ID, b.�Ǽ�ʱ�� As �շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, b.No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��ID," & _
            "           To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As ��Ʊʱ��,a.�˿�ID,0 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,סԺ���ü�¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=5 And b.��¼����=5 And b.��¼״̬ In(1,3)" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�Ǽ�ʱ�� Between [3] And [4]")
    End If
    '3)����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select a.ID, b.�շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, b.No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��ID," & _
            "           To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As ��Ʊʱ��,a.�˿�ID,0 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,���˽��ʼ�¼ B" & _
            " Where a.����ID =b.ID And a.Ʊ��=3 And b.��¼״̬ In(1,3)" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�շ�ʱ�� Between [3] And [4]")
    End If
    '4)�Һš��շ�
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Or bytҵ�񳡺� = 4 Then
        strSqlSub = _
            " Select a.ID, b.�Ǽ�ʱ�� As �շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, Min(b.No) As No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��ID," & _
            "           Max(To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss')) As ��Ʊʱ��,a.�˿�ID,0 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,������ü�¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=[Ʊ��] And b.��¼����=[��¼����] And b.��¼״̬ In(1,3)" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�Ǽ�ʱ�� Between [3] And [4]") & _
            " Group By b.�Ǽ�ʱ��,a.ID,a.Ʊ��,a.��¼״̬, a.����, a.����,a.������, a.Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��,a.ԭƱ��id,a.�˿�ID"
        
        '���ղ������
        strSqlSub = strSqlSub & " Union All " & _
            " Select a.ID, b.�Ǽ�ʱ�� As �շ�ʱ��,a.Ʊ��,a.��¼״̬ As Ʊ��״̬, b.No, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���,a.������, Decode(a.��¼״̬, 2, -1, 1) * a.Ʊ�ݽ�� As Ʊ�ݽ��," & _
            "           a.����ID,a.����id,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.�Ƿ񻻿�,a.ֽ�ʷ�Ʊ��,a.��Ʊ��, a.ԭƱ��ID," & _
            "           To_Char(To_Date(Substr(a.����ʱ��, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As ��Ʊʱ��,a.�˿�ID,1 As ������" & _
            " From ����Ʊ��ʹ�ü�¼ A,���ò����¼ B" & _
            " Where a.����ID =b.����ID And a.Ʊ��=[Ʊ��]  And b.��¼����=[��¼����] And Nvl(b.���ӱ�־,0)=[���ӱ�־] And b.��¼״̬ In(1,3)" & strWhere & _
                        IIf(bytʱ������ = 0, "", " And b.�Ǽ�ʱ�� Between [3] And [4]")
                    
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(strSqlSub, "[��¼����]", 1), "[���ӱ�־]", 0), "[Ʊ��]", 1)
        End If
        
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(strSqlSub, "[��¼����]", 4), "[���ӱ�־]", 1), "[Ʊ��]", 4)
        End If
    End If
    strSQL = strSQL & " Order By �շ�ʱ��"
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʊ������", _
        Format(dt��ʼʱ��, "yyyyMMddHHmmss"), Format(dt����ʱ��, "yyyyMMddHHmmss"), dt��ʼʱ��, dt����ʱ��, varQueryValue, str��Ʊ��)
    GetEInvoiceData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEInvoiceExse(ByVal bytҵ�񳡺� As Byte, ByVal lng����ID As Long, ByRef rsExse As ADODB.Recordset) As Boolean
    '��ȡ����Ʊ�ݷ�������
    '��Σ�
    '   bytҵ�񳡺� 1-�շѣ�2-Ԥ����3-���ʣ�4-�Һţ�5-���￨
    '   lng����ID bytҵ�񳡺�=2������Ԥ����¼.ID������������ID
    '���Σ�
    '   rsExse ��¼����NO,���,��������,������,�ѱ�,���,����,��Ʒ��,���,��λ,ִ�п���,����,����,Ӧ�ս��,ʵ�ս��,���ʽ��
    Dim strSQL As String, strWhere As String
    
    On Error GoTo ErrHandler
    Select Case bytҵ�񳡺�
    '1)���￨������
    Case 3, 5
        strSQL = _
            " Select a.No, Nvl(a.���, a.�۸񸸺�) As ���, a.��������id, a.������, a.�ѱ�," & _
            "        a.�շ����, a.�շ�ϸĿid, a.���㵥λ As ��λ, a.ִ�в���id," & _
            "        Avg(Nvl(a.����, 1) * a.����) As ����, Sum(a.��׼����) As ����," & _
            "        Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��,Sum(a.���ʽ��) As ���ʽ��" & _
            " From סԺ���ü�¼ A, סԺ���ü�¼ A1" & _
            " Where a.��¼���� = a1.��¼���� And a.No = a1.No And a.��� = a1.��� And a1.����id = [1]" & _
            " Group By a.No, a.��¼����, a.��¼״̬, Nvl(a.���, a.�۸񸸺�), a.��������id, a.������, a.�ѱ�," & _
            "       a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.ִ�в���id"
    '2)�Һţ��շ�
    Case 1, 4
        strSQL = _
        " Select a.No, Nvl(a.���, a.�۸񸸺�) As ���, a.��������id, a.������, a.�ѱ�," & _
        "        a.�շ����, a.�շ�ϸĿid, a.���㵥λ As ��λ, a.ִ�в���id," & _
        "        Avg(Nvl(a.����, 1) * a.����) As ����, Sum(a.��׼����) As ����," & _
        "        Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��" & _
        " From ������ü�¼ A," & _
        "      (Select a.��¼����, a.No, a.���" & _
        "        From ������ü�¼ A" & _
        "        Where a.����id = [1] And Not Exists (Select 1 From ���ò����¼ Where �շѽ���id = a.����id)" & _
        "        Union All" & _
        "        Select a.��¼����, a.No, a.���" & _
        "        From ������ü�¼ A, ���ò����¼ B" & _
        "        Where a.����id = b.�շѽ���id And b.����id = [1]) A1" & _
        " Where Mod(a.��¼����, 10) = a1.��¼���� And a.No = A1.No And a.��� = A1.���" & _
        " Group By a.No, a.��¼����, a.��¼״̬, Nvl(a.���, a.�۸񸸺�), a.��������id, a.������, a.�ѱ�," & _
        "       a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.ִ�в���id"
    Case Else
        Exit Function
    End Select
    
    strSQL = _
        " Select a.No, a.���, b.���� As ��������, a.������, a.�ѱ�, c.���� As ���, e.����, f.���� As ��Ʒ��," & _
        "        e.���, a.��λ, B1.���� As ִ�п���, Sum(a.����) As ����, Avg(a.����) As ����," & _
        "        Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��" & _
        " From (" & strSQL & ") A, ���ű� B, ���ű� B1, �շ���Ŀ��� C, �շ���ĿĿ¼ E, �շ���Ŀ���� F" & _
        " Where a.��������id = b.Id And a.ִ�в���id = B1.Id And a.�շ���� = c.���� And a.�շ�ϸĿid = e.Id" & _
        "       And e.Id = f.�շ�ϸĿid(+) And f.����(+) = 1 And f.����(+) = 3" & _
        " Group By a.No, a.���, b.����, a.������, a.�ѱ�, c.����, e.����, f.����, e.���, a.��λ, B1.����" & _
        " Having Nvl(Sum(a.����), 0) <> 0" & _
        " Order By a.No, a.���"

    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID)
    GetEInvoiceExse = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetExseData(ByVal bytҵ�񳡺� As Byte, ByVal str�շ�Ա As String, _
    ByVal dt��ʼʱ�� As Date, ByVal dt����ʱ�� As Date, ByRef rsExse As ADODB.Recordset) As Boolean
    '��ȡ����Ʊ�ݷ�������
    '��Σ�
    '   bytҵ�񳡺� 0-���У�1-�շѣ�2-Ԥ����3-���ʣ�4-�Һţ�5-���￨
    Dim strSQL As String, strWhere As String, strSqlSub As String
    
    On Error GoTo ErrHandler
    strWhere = " And a.�տ�ʱ�� Between [1] And [2]"
    If Trim(str�շ�Ա) <> "" Then strWhere = strWhere & " And a.����Ա����=[3]"
    
    '1)Ԥ����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 2 Then
        strSQL = _
            " Select 2 As ҵ������, a.Id As ����ID, a.No, a.���, a.����Ա����, a.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, Null As ����, Null As �Ա�, Null As ����, a.Ԥ�����, Null As ����ID, Null As ��������, Null As ������" & _
            " From ����Ԥ����¼ A" & _
            " Where a.��¼���� = 1 And a.��¼״̬ = 1 And a.Ԥ������Ʊ�� = 1" & strWhere & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.Id And Ʊ�� = 2 And ��¼״̬ = 1)"
        '����˿�
        strSQL = strSQL & " Union All" & _
            " Select 12 As ҵ������, b.ID As ����ID, a.No, a.���, a.����Ա����, a.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, Null As ����, Null As �Ա�, Null As ����, a.Ԥ�����,a.Id As ����ID, Null As ��������, Null As ������" & _
            " From ����Ԥ����¼ A,����Ԥ����¼ B" & _
            " Where a.��¼���� = 11 And a.��¼״̬ = 1 And a.Ԥ������Ʊ�� = 1" & strWhere & _
            "       And Exists(Select 1 From ����Ԥ����¼ Where ��¼���� = 1 And ���ӱ�־ = 1 And ����id = a.����id)" & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.Id And Ʊ�� = 2 And ��¼״̬ = 1)" & _
            "       And a.No=b.No And b.��¼����=1 And b.��¼״̬ In(1,3)"
    End If
    '2)���￨
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select 5 As ҵ������, b.����id As ����ID, b.No, Sum(a.���ʽ��) As ���, b.����Ա����, b.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, a.����, a.�Ա�, a.����, Null As Ԥ�����, Null As ����ID, Null As ��������, Null As ������" & _
            " From סԺ���ü�¼ A, סԺ���ü�¼ A1," & _
            "      (Select Distinct a.����id, a.����Ա����, a.�տ�ʱ��, b.No" & _
            "       From ����Ԥ����¼ A, סԺ���ü�¼ B" & _
            "       Where a.����id = b.����ID And b.��¼���� = 5 And b.��¼״̬ In(1,3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "             And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����Id And Ʊ�� = 5 And ��¼״̬ = 1)) B" & _
            " Where a.No = a1.No And a.��� = a1.��� And a1.����id = b.����ID And a.��¼���� = 5" & _
            " Group By b.����id, b.No, b.����Ա����, b.�տ�ʱ��, a.����id, a.��ҳid, a.����, a.�Ա�, a.����" & _
            " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
    End If
    '3)����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct 3 As ҵ������, a.����id As ����ID, b.No, b.���ʽ�� As ���, a.����Ա����, a.�տ�ʱ��," & _
            "           b.����id, b.��ҳid, Null As ����, Null As �Ա�, Null As ����, Null As Ԥ�����, Null As ����ID, b.��������, Null As ������" & _
            " From ����Ԥ����¼ A, ���˽��ʼ�¼ B" & _
            " Where a.����id = b.ID And b.��¼״̬ = 1 And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����Id And Ʊ�� = 3 And ��¼״̬ = 1)"
    End If
    '4)�Һš��շ�
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Or bytҵ�񳡺� = 4 Then
        strSqlSub = _
            " Select a.����id, a.����Ա����, a.�տ�ʱ��, Min(b.No) As No, 0 As ������, Null As ����ID" & _
            " From ����Ԥ����¼ A, ������ü�¼ B" & _
            " Where a.����id = b.����ID And b.��¼���� = [��¼����] And b.��¼״̬ In(1,3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "             And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����id And Ʊ�� = [Ʊ��] And ��¼״̬ = 1)" & _
            "             And Not Exists(Select 1 From ���ò����¼ Where �շѽ���id = a.����id And ��¼���� = 1 And Nvl(���ӱ�־,0) = [���ӱ�־])" & _
            " Group By a.����id, a.����Ա����, a.�տ�ʱ��"
        '�������
        strSqlSub = strSqlSub & " Union All " & _
            " Select ����id, ����Ա����, �տ�ʱ��, No, 1 As ������, ����ID" & _
            " From (Select Distinct b.�շѽ���ID As ����id, a.����Ա����, a.�տ�ʱ��,b.No As No, b.����ID," & _
            "                    Row_Number() Over(Partition By b.��¼����, b.No Order By b.�Ǽ�ʱ��) As ���" & _
            "            From ����Ԥ����¼ A, ���ò����¼ B" & _
            "            Where a.����ID=b.����ID And b.��¼���� = 1 And Nvl(b.���ӱ�־,0) = [���ӱ�־] And b.��¼״̬ In(1, 3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "                       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����id And Ʊ�� = [Ʊ��] And ��¼״̬ = 1))" & _
            " Where ��� = 1"
            
        strSqlSub = _
            " Select [ҵ������] As ҵ������, Nvl(b.����ID, b.����id) As ����ID, Min(b.No) As No, Sum(a.���ʽ��) As ���, b.����Ա����, b.�տ�ʱ��," & _
            "        a.����id, a.��ҳid, a.����, a.�Ա�, a.����, Null As Ԥ�����, Null As ����ID, Null As ��������, b.������" & _
            " From ������ü�¼ A, ������ü�¼ A1,(" & strSqlSub & ") B" & _
            " Where a.No = a1.No And a.��� = a1.��� And a1.����id = b.����ID And Mod(a.��¼����,10)=[��¼����]" & _
            " Group By Nvl(b.����ID, b.����id), b.����Ա����, b.�տ�ʱ��, a.����id, a.��ҳid, b.������, a.����, a.�Ա�, a.����" & _
            " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
        
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[ҵ������]", 1), "[��¼����]", 1), "[Ʊ��]", 1), "[���ӱ�־]", 0)
        End If
        
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[ҵ������]", 4), "[��¼����]", 4), "[Ʊ��]", 4), "[���ӱ�־]", 1)
        End If
    End If
    
    strSQL = _
        " Select Nvl(n.����,Nvl(m.����,a.����)) As ����,Nvl(n.�Ա�,Nvl(m.�Ա�,a.�Ա�)) As �Ա�,Nvl(n.����,Nvl(m.����,a.����)) As ����," & _
        "           m.����� As �����, Nvl(n.סԺ��,m.סԺ��) As סԺ��, a.ҵ������, a.����id, a.No, a.���, a.����Ա����, a.�տ�ʱ��, " & _
        "           a.����id, a.��ҳid, a.Ԥ�����, a.����id, a.��������, a.������" & _
        " From (" & strSQL & ") A, ������Ϣ M, ������ҳ N" & _
        " Where a.����ID=m.����ID(+) And a.����ID=n.����ID(+) And a.��ҳID=n.��ҳID(+)" & _
        " Order By �տ�ʱ��"
    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʊ������", dt��ʼʱ��, dt����ʱ��, str�շ�Ա)
    GetExseData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Grid_SelAllRecord(vsfGrid As VSFlexGrid, ByVal blnSel As Boolean, Optional ByVal strColName As String = "ѡ��")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȫѡ/ȫ���¼
    '���:
    '   blnSel-ѡ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsfGrid
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex(strColName)) = blnSel
        Next
    End With
End Sub

Public Function GetSwapCollectFromBalanceID(ByVal byt���� As Byte, ByVal lngԭ����ID As Long, _
    ByRef cllSwapData_Out As Collection, Optional ByVal bln������ As Boolean, _
    Optional ByVal lng����ID As Long, Optional ByVal blnShowMsg As Boolean = True, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ���׽�����Ϣ
    '���:
    '    byt����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '   lngԭ����ID byt����=2������Ԥ����¼.ID������������ID
    '����:
    '   cllSwapData_Out-���ؽ�����Ϣ
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����),key(_�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    Select Case byt����
    Case 1, 4
        If bln������ Then
            strWhere = " And b.����id In(Select �շѽ���ID From ���ò����¼ Where ����ID=[1])"
        Else
            strWhere = " And b.����id = [1]"
        End If
    
        strSQL = _
            " Select Max(a.����id) As ����ID, Max(a.��ҳid) As ��ҳID, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����," & _
            "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.���ʽ��) As ���ʽ��, Max(a.�Ǽ�ʱ��) As �շ�ʱ��" & _
            " From (Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���, Sum(a.���ʽ��) As ���ʽ��, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��" & _
            "        From ������ü�¼ A, ������ü�¼ B" & _
            "        Where Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.No = b.No And a.��� = b.���" & strWhere & _
            "        Group By a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���" & _
            "        Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) A"
        
        strInsureSql = "Select Max(����) As ���� From ���ս����¼ Where ���� = 1 And ��¼id = [1]"
        
        strSQL = _
            " Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, m.�����, Nvl(n.סԺ��, m.סԺ��) As סԺ��," & _
            "           a.No, a.���ʽ��, a.�շ�ʱ��, b.����" & _
            " From (" & strSQL & ") A, (" & strInsureSql & ") B, ������Ϣ M, ������ҳ N" & _
            " Where a.����id = m.����id(+) And a.����id = n.����id(+) And a.��ҳid = n.��ҳid(+) And a.No Is Not Null"
    Case 2
        strSQL = _
            "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(A.���) As ���ʽ��, Max(A.Ԥ������Ʊ��) As �Ƿ����Ʊ��, " & _
            "          Max(Nvl(d.����, c.����)) As ����, " & _
            "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
            "          max(M.����) as ����,to_char(max(A.�տ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(a.Ԥ�����) as Ԥ�����" & _
            "   From  ����Ԥ����¼ A, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=3  and ��¼ID=[1] ) M" & _
            "   Where a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
            "   Group By a.Id, a.No, a.����id, a.��ҳid"
    Case 3
        strSQL = _
            "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(b.��Ԥ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, " & _
            "          Max(decode(nvl(A.����ID,0),0,A.ԭ��,Nvl(d.����, c.����))) As ����, " & _
            "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
            "          max(M.����) as ����,to_char(max(A.�շ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(A.��������) as ��������" & _
            "   From ���˽��ʼ�¼ A, ����Ԥ����¼ B, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=2  and ��¼ID=[1] ) M" & _
            "   Where a.id=b.����ID and  a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
            "   Group By a.Id, a.No, a.����id, a.��ҳid"
    Case 5
        strSQL = _
            "   Select a.����id As ID, b.No, a.����id, a.��ҳid, Sum(a.��Ԥ��) As ���ʽ��, Max(a.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, " & _
            "          Max(c.����) As ����, Max(c.סԺ��) As סԺ��, Max(c.�����) As �����, 0 As ����, " & _
            "          To_Char(Max(a.�տ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ�� " & _
            "   From ����Ԥ����¼ A, (Select  ����id,No From סԺ���ü�¼ Where ����id = [1]) B, ������Ϣ C  " & _
            "   Where a.����id = b.����id And a.����id = c.����id(+)  And a.����id = [1] " & _
            "   Group By a.����id, b.No, a.����id, a.��ҳid"
    Case Else
        strErrMsg_Out = "���볡�ϡ�" & byt���� & "��������Ч��": Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��������Ʊ����Ϣ", IIf(byt���� = 2 And lng����ID <> 0, lng����ID, lngԭ����ID))
    If rsTemp.EOF Then strErrMsg_Out = "��ʣ��δ�˷������ݡ�": Exit Function
    
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!����ID)), "_����ID"
    cllPati.Add Val(Nvl(rsTemp!��ҳid)), "_��ҳID"
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�Ա�), "_�Ա�"
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�����), "_�����"
    cllPati.Add Nvl(rsTemp!סԺ��), "_סԺ��"
    cllPati.Add Val(Nvl(rsTemp!����)), "_����"

    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add "", "_��Ʊ��"
    cllBalanceInfo.Add lngԭ����ID, "_����ID"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add Nvl(rsTemp!No), "_���ݺ�"
    cllBalanceInfo.Add Format(Nvl(rsTemp!�շ�ʱ��), "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    If byt���� = 1 Or byt���� = 4 Then
        cllBalanceInfo.Add IIf(bln������, 1, 0), "_�Ƿ񲹽���"
    Else
        cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
    End If
    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add Val(Nvl(rsTemp!���ʽ��)), "_������"
    cllBalanceInfo.Add 0, "_����ID"
    Select Case byt����
    Case 2
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!Ԥ�����)) = 0, 3, Val(Nvl(rsTemp!Ԥ�����))), "_��������" 'Ԥ�����:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add 0, "_��Լ��λ����"
    Case 3
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!��������)) = 0, 3, Val(Nvl(rsTemp!��������))), "_��������"  '��������:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add IIf(Val(Nvl(rsTemp!����ID)) = 0, 1, 0), "_��Լ��λ����"
    Case Else
        cllBalanceInfo.Add 1, "_��������"
        cllBalanceInfo.Add 0, "_��Լ��λ����"
    End Select
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    If Not blnShowMsg Then strErrMsg_Out = Err.Description: Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Getҽ�ƿ����ʽ����(ByVal strҽ�ƿ����ʽ���� As String) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ����ҽ�ƿ����ʽ�����ȡ����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Max(����) as ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Getҽ�ƿ����ʽ����", strҽ�ƿ����ʽ����)
    If rsTmp.RecordCount > 0 Then
        Getҽ�ƿ����ʽ���� = Nvl(rsTmp!����)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetԤ�������ܶ�(ByVal strNO As String) As Double
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����Ԥ�����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Sum(���) As Ʊ���ܽ��" & vbNewLine & _
            "  From (Select Sum(���) As ���" & vbNewLine & _
            "         From ����Ԥ����¼" & vbNewLine & _
            "         Where NO = r_Deposit_Rec.No And ��¼���� = 1" & vbNewLine & _
            "         Union All" & vbNewLine & _
            "         Select Sum(��Ԥ��) As ���" & vbNewLine & _
            "         From ����Ԥ����¼" & vbNewLine & _
            "         Where ����id In (Select Distinct ����id From ����Ԥ����¼ Where NO = [1] And Mod(��¼����, 10) = 1) And" & vbNewLine & _
            "               Nvl(���, 0) < 0 And Mod(��¼����, 10) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetԤ�����", strNO)
    If rsTmp.RecordCount > 0 Then
        GetԤ�������ܶ� = Val(Nvl(rsTmp!Ʊ���ܽ��))
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetԤ�����(ByVal lng����ID As Long, ByVal intԤ������ As Integer) As Double
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ����Ԥ�����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Max(Ԥ�����) As Ԥ����� From ������� " & _
            " Where ����id = [1] And ���� = 1 And ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetԤ�����", lng����ID, intԤ������)
    If rsTmp.RecordCount > 0 Then
        GetԤ����� = Val(Nvl(rsTmp!Ԥ�����))
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetEInvoiceInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = _
    " Select ID, Ʊ��, ��¼״̬, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, ���� As Ʊ�ݴ���, ���� As Ʊ�ݺ���, ������ As Ʊ��У����, Ʊ�ݽ��," & _
    "           ����ʱ��, Url����, ԭƱ��id, �Ƿ񻻿�, ֽ�ʷ�Ʊ��, ��ӡid, ��ע, ����Ա���, ����Ա����, �Ǽ�ʱ��, ��Ʊ��, ϵͳ��Դ, Url���� " & _
    " From ����Ʊ��ʹ�ü�¼ Where Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ�����Ʊ��ʹ�ü�¼�����顣": Exit Function
    End If
    Set GetEInvoiceInfo = rsTmp
    Exit Function
errHand:
    strErrMsg_Out = Err.Description
End Function

Public Function GetEInvoiceWithPatiInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select a.Ʊ��, a.���� As Ʊ�ݴ���, a.���� As Ʊ�ݺ���, a.������ As Ʊ��У����, b.�ֻ���, b.email," & vbNewLine & _
            "        a.�Ƿ񻻿� " & vbNewLine & _
            "From ����Ʊ��ʹ�ü�¼ a, ������Ϣ b" & vbNewLine & _
            "Where a.Id =[] And a.����id = b.����id(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "δ�ҵ�����Ʊ��ʹ�ü�¼�����顣": Exit Function
    End If
    Set GetEInvoiceWithPatiInfo = rsTmp
    Exit Function
errHand:
    strErrMsg_Out = Err.Description
End Function

Public Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Զ���ɫ
    '���ƣ����˺�
    '���ڣ�2010-03-23 10:52:23
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '����ؼ�
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             .BackColorSel = vbBlue
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '���ѡ����ɫ
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
              
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Public Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngOldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����иı�ʱ,������ص���ɫ
    '��Σ�CustomColor-�Զ�����ɫ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:22:38
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '�иı�ʱ
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Public Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1, Optional ForeColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '���ܣ��뿪����ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Ƿ����Զ�����ɫ������(BackColor)�ķ�ʽ������)
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:03:05
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
            If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            If ForeColor = -1 Then .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
        If ForeColor <> -1 Then
            .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = ForeColor
        End If
        .ForeColorSel = .ForeColor
    End With
End Sub
