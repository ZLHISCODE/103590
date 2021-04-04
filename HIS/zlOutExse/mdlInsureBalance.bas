Attribute VB_Name = "mdlInsureBalance"
Option Explicit


Public Function initInsurePara(ByVal intInsure As Integer, ByVal lng����ID As Long, _
    ByVal bln����Ʊ�� As Boolean, _
    Optional ByVal lng����ID As Long) As Ty_InsurePatiPara
    '��ʼ��ҽ������
    Dim tyInsurePara As Ty_InsurePatiPara
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then Exit Function
    If gclsInsure Is Nothing Then Exit Function
    
    tyInsurePara.��������ҽ����Ŀ = gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, intInsure)
    tyInsurePara.�����շѴ�Ϊ���۵� = gclsInsure.GetCapability(support�����շѴ�Ϊ���۵�, lng����ID, intInsure)
    tyInsurePara.�����ѽɿ���� = gclsInsure.GetCapability(support�����ѽɿ����, lng����ID, intInsure)
    tyInsurePara.������봫����ϸ = gclsInsure.GetCapability(support������봫����ϸ, lng����ID, intInsure)
    
    If bln����Ʊ�� = False Then
        tyInsurePara.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
        tyInsurePara.���������շ� = gclsInsure.GetCapability(support���������շ�, lng����ID, intInsure)
    End If
    
    tyInsurePara.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, lng����ID, intInsure)
    tyInsurePara.�൥��һ�ν��� = gclsInsure.GetCapability(support�൥��һ�ν���, lng����ID, intInsure)
    tyInsurePara.�൥�ݷֵ��ݽ��� = gclsInsure.GetCapability(support�൥�ݷֵ��ݽ���, lng����ID, intInsure)
    tyInsurePara.һ�ν���ֵ����˷� = gclsInsure.GetCapability(supportһ�ν���ֵ����˷�, lng����ID, intInsure)
    
    tyInsurePara.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, lng����ID, intInsure)
    tyInsurePara.�൥���շ� = gclsInsure.GetCapability(support�൥���շ�, lng����ID, intInsure)
    tyInsurePara.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, intInsure)
    tyInsurePara.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, lng����ID, intInsure)
    tyInsurePara.���Ը� = gclsInsure.GetCapability(support�շ��ʻ������Ը�, lng����ID, intInsure)
    
    tyInsurePara.ȫ�Ը� = gclsInsure.GetCapability(support�շ��ʻ�ȫ�Է�, lng����ID, intInsure)
    tyInsurePara.blnOnlyBjYb = False
    tyInsurePara.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, lng����ID, intInsure)
    tyInsurePara.ҽ������Ʊ�� = False
    tyInsurePara.����������� = gclsInsure.GetCapability(support�����������, lng����ID, intInsure)
    
    tyInsurePara.������ȫ�� = gclsInsure.GetCapability(support������ȫ��, lng����ID, intInsure)
    initInsurePara = tyInsurePara
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlExecuteInsurePreSwap(ByVal bytMode As Byte, objBill As ExpenseBill, _
    ByVal strInvoice As String, ByVal str�ѱ� As String, ByVal bln���� As Boolean, _
    ByVal intInsure As Integer, ByRef colBalance As Collection, _
    ByRef bln����Ʊ�� As Boolean, ByRef strInsureInvoice As String, _
    ByVal strErrMsg As String, Optional ByVal blnReBalance As Boolean, _
    Optional ByVal lng����ID As Long, Optional ByVal bln������ As Boolean, _
    Optional ByVal blnErrBill As Boolean) As Boolean
    '����Ԥ����
    '��Σ�
    '   bytMode ҽ������ģʽ��0-�൥��һ�ν���,1-�൥��һ�ν���ֵ����˷�,2-�൥�ݷֵ��ݽ���
    '   objBill ��������,ExpenseBill����
    '   strInvoice ��ǰ��Ʊ��
    '   blnReBalance �Ƿ����½���,ΪTrueʱ�����ݸ��ݽ���ID�����ݿ��л�ȡ����Ҫ�ǲ�����Ͷ൥��һ�ν�������
    '���Σ�
    '   colBalance Ԥ��������(ÿ�ŵ��ݶ�Ӧһ��BalanceMoneys����Ԫ��),�൥��һ�ν���ʱ���ڵ�һ�ŵ�����
    '   bln����Ʊ�� �Ƿ�ҽ������Ʊ��
    '   strInsureInvoice ҽ�����ط�Ʊ��
    '   strErrMsg ������Ϣ,Falseʱ����
    Dim strDate As String, rsBalance As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim strBalance As String, strAdvance As String
    Dim varBalance As Variant, varItem As Variant, str���㷽ʽ As String
    Dim p As Long, i As Long
    
    On Error GoTo ErrHandler
    bln����Ʊ�� = False: strInsureInvoice = ""
    strErrMsg = ""
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:mm:ss")
    
    '2-�൥�ݷֵ��ݽ���
    If bytMode = 2 Then
        If blnErrBill Then Set rsBalance = zlGetBalanceDetail(0, lng����ID, 1)
        For p = 1 To objBill.Pages.Count
            
            strBalance = ""
            If blnErrBill Then
                '�����ŵ����Ƿ��ѳɹ�ҽ������
                rsBalance.Filter = "No='" & objBill.Pages(p).NO & "'"
                Do While Not rsBalance.EOF
                    strBalance = strBalance & IIf(strBalance = "", "", "||")
                    strBalance = strBalance & Nvl(rsBalance!���㷽ʽ) & "|" & Val(Nvl(rsBalance!���))
                    rsBalance.MoveNext
                Loop
            End If
            
            If strBalance <> "" Then
                Call SetBalanceVal(colBalance, p, strBalance)
            Else
                If blnReBalance Then
                    Set rsRecord = MakePreSwapDataFromDB(strInvoice, lng����ID, objBill.Pages(p).NO, bln������)
                Else
                    Set rsRecord = MakeBillRecord(objBill, bln����, p, strDate, str�ѱ�, strInvoice)
                End If
                strBalance = "": strAdvance = IIf(blnReBalance, "1", "")
                
                If Not gclsInsure.ClinicPreSwap(rsRecord, strBalance, intInsure, strAdvance) Then
                    strErrMsg = "�� " & p & " �ŵ���Ԥ����ʧ�ܡ�"
                    Exit Function
                End If
                
                'ֻҪ��һ�ŵ����Զ���Ʊ�ţ���Ҫ��Ʊ��
                If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then    'ҽ��Ʊ�ݺ� Then
                    '38821,��ʽ:Ʊ�ݺ�;�Ƿ���Ʊ��(1-����Ʊ��;0-�Զ���Ʊ��)
                    varItem = Split(strAdvance & ";", ";")
                    strInsureInvoice = varItem(0)
                    bln����Ʊ�� = bln����Ʊ�� And Val(varItem(1)) = 1
                End If
                
                '������ʽ;���;�Ƿ������޸�|....
                If strBalance <> "" Then
                    strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
                    Call SetBalanceVal(colBalance, p, strBalance)
                End If
            End If
        Next
        ZlExecuteInsurePreSwap = True: Exit Function
    End If
    
    
    '0-�൥��һ�ν���,1-�൥��һ�ν���ֵ����˷�
    If blnReBalance Then
        Set rsRecord = MakePreSwapDataFromDB(strInvoice, lng����ID, "", bln������)
    Else
        Set rsRecord = MakeBillRecord(objBill, bln����, 0, strDate, str�ѱ�, strInvoice)
    End If
    strBalance = ""
    If bln������ Then
        strAdvance = IIf(blnReBalance, "3", "2")
    Else
        strAdvance = IIf(blnReBalance, "1", "")
    End If
    
    If Not gclsInsure.ClinicPreSwap(rsRecord, strBalance, intInsure, strAdvance) Then
        strErrMsg = "����Ԥ����ʧ�ܡ�"
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then
        '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
        varItem = Split(strAdvance & ";", ";")
        strInsureInvoice = varItem(0)
        bln����Ʊ�� = Val(varItem(1)) = 1
    End If
    
    '������ʽ;���;�Ƿ������޸�|....
    If strBalance <> "" Then
        If bytMode = 0 Then
            '0-�൥��һ�ν���
            strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
            Call SetBalanceVal(colBalance, 1, strBalance)
        
        Else
            '1-�൥��һ�ν���ֵ����˷�
            '�������:���㷽ʽ;���;�Ƿ������޸�|...||�������:���㷽ʽ;���;�Ƿ������޸�|...||...
            varBalance = Split(strBalance, "||")
            For i = 0 To UBound(varBalance)
                If InStr(varBalance(i), ":") = 0 Then
                    strErrMsg = "����Ԥ���㷵�ؽ�������ʽ����ȷ��"
                    Exit Function
                End If
                
                varItem = Split(varBalance(i), ":")
                p = Val(varItem(0)): str���㷽ʽ = varItem(1)
                If p < 1 Or p > colBalance.Count Then
                    strErrMsg = "����Ԥ���㷵�ؽ�������ʽ����ȷ��"
                    Exit Function
                End If
                
                str���㷽ʽ = Replace(Replace(str���㷽ʽ, "|", "||"), ";", "|")
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            Next
        End If
    End If
    
    ZlExecuteInsurePreSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckInsureBalanceValid(ByRef rs���㷽ʽ As ADODB.Recordset, _
    ByVal colBalance As Collection) As String
    '���ҽ���е�����û�еĽ��㷽ʽ�����ر���û�еĽ��㷽ʽ
    '��Σ�
    '   colBalance BalanceMoneys����
    Dim i As Integer, strNone As String
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHandler
    If colBalance Is Nothing Then Exit Function
    
    For i = 1 To colBalance.Count
        For Each objItem In colBalance(i)
            If rs���㷽ʽ Is Nothing Then
                If InStrEx(strNone, objItem.���㷽ʽ) = False Then
                    strNone = strNone & "," & objItem.���㷽ʽ
                End If
            Else
                rs���㷽ʽ.Filter = "(����='" & objItem.���㷽ʽ & "' And ����=3) Or (����='" & objItem.���㷽ʽ & "' And ����=4)"
                If rs���㷽ʽ.EOF Then
                    If InStrEx(strNone, objItem.���㷽ʽ) = False Then
                        strNone = strNone & "," & objItem.���㷽ʽ
                    End If
                End If
            End If
        Next
    Next
    If strNone <> "" Then strNone = Mid(strNone, 2)
    
    CheckInsureBalanceValid = strNone
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InsureBalanced(ByVal intInsure As Integer, ByVal lng����ID As Long) As Boolean
    '�ж��Ƿ��ѳɹ�������ҽ������
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then Exit Function
    'У�Ա�־����2���ѳɹ�����
    strSql = _
        "Select 1" & vbNewLine & _
        "From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
        "Where a.���㷽ʽ = b.���� And b.���� In (3, 4)  And Nvl(У�Ա�־, 0) = 2" & vbNewLine & _
        "      And a.��¼���� = 3 And a.����id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "mdlCliniBalance", lng����ID)
    InsureBalanced = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMedicareSum(colBalance As Collection, Optional ByVal strItem As String, Optional ByVal intPage As Integer, _
    Optional ByVal blnOrig As Boolean, Optional ByVal intPageCount As Integer) As Currency
    '���ܣ���ȡ���ս���Ľ��
    '������strItem=�Ƿ�ָ�����㷽ʽ,����Ϊ���н��㷽ʽ
    '      blnOrig=�Ƿ�ȡԭʼ(���)������,����ȡ����(�޸ĺ�)��Ч���
    '      intPage=�Ƿ�ָ������,����Ϊ���е���
    '      intPageCount=���㵥������
    '˵�����ú�����colBalanceΪ׼����,����ҽ�������շ�Ҳ��
    Dim curMoney As Currency, p As Integer
    Dim intPageStart As Integer, intPageEnd As Integer
    Dim objItem As BalanceMoney
    
    intPageStart = IIf(intPage = 0, 1, intPage)
    intPageEnd = IIf(intPage = 0, IIf(intPageCount = 0, colBalance.Count, intPageCount), intPage)
    For p = intPageStart To intPageEnd
        For Each objItem In colBalance(p)
            If strItem = "" Or objItem.���㷽ʽ = strItem Then
                If blnOrig Then
                    curMoney = curMoney + objItem.ԭʼ���
                Else
                    curMoney = curMoney + objItem.��Ч���
                End If
            End If
        Next
    Next
    GetMedicareSum = curMoney
End Function

Public Function GetMedicareStr(colBalance As Collection, Optional ByVal intPage As Integer, _
    Optional ByVal intPageCount As Integer) As String
    '���ܣ����ر��ս��㷽ʽ��,"���㷽ʽ|���||...."
    '������intPage=�Ƿ�ָ������,����Ϊ���е���
    '      intPageCount=���㵥��������
    '˵�����ú�����colBalanceΪ׼����,����ҽ�������շ�Ҳ��
    Dim p As Integer
    Dim rsTemp As New ADODB.Recordset, strBalance As String
    Dim intPageStart As Integer, intPageEnd As Integer
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    rsTemp.Fields.Append "���㷽ʽ", adVarChar, 20, adFldIsNullable
    rsTemp.Fields.Append "���", adCurrency, , adFldIsNullable
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    intPageStart = IIf(intPage = 0, 1, intPage)
    intPageEnd = IIf(intPage = 0, IIf(intPageCount = 0, colBalance.Count, intPageCount), intPage)
    For p = intPageStart To intPageEnd
        For Each objItem In colBalance(p)
            rsTemp.Find "���㷽ʽ='" & objItem.���㷽ʽ & "'", , adSearchForward, 1
            If rsTemp.EOF Then rsTemp.AddNew
            rsTemp!���㷽ʽ = objItem.���㷽ʽ
            rsTemp!��� = Val(Nvl(rsTemp!���)) + objItem.��Ч���
            rsTemp.Update
        Next
    Next
    
    strBalance = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strBalance = strBalance & "||" & Nvl(rsTemp!���㷽ʽ) & "|" & Nvl(rsTemp!���)
        rsTemp.MoveNext
    Loop
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetMedicareStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInsureBalanceSum(objBalanceMoneys As BalanceMoneys, _
    Optional ByVal strItem As String, Optional ByVal blnOrig As Boolean) As Currency
    '��ȡ���ս���Ľ��
    '��Σ�
    '   strItem �Ƿ�ָ�����㷽ʽ,����Ϊ���н��㷽ʽ
    '   blnOrig �Ƿ�ȡԭʼ(���)������,����ȡ����(�޸ĺ�)��Ч���
    Dim curMoney As Currency
    Dim objItem As BalanceMoney
     
    If objBalanceMoneys Is Nothing Then Exit Function
    For Each objItem In objBalanceMoneys
        If strItem = "" Or objItem.���㷽ʽ = strItem Then
            If blnOrig Then
                curMoney = curMoney + objItem.ԭʼ���
            Else
                curMoney = curMoney + objItem.��Ч���
            End If
        End If
    Next
    GetInsureBalanceSum = curMoney
End Function

Private Function GetInsureBalanceStrAll(objBalanceBills As BalanceBills) As String
    '��ȡ���е��ݵ�Ԥ������,"���㷽ʽ|���||...."
    Dim i As Integer
    Dim colBalance As New Collection
    
    If objBalanceBills Is Nothing Then Exit Function
    For i = 1 To objBalanceBills.Count
        colBalance.Add objBalanceBills(i).Ԥ����
    Next
    GetInsureBalanceStrAll = GetMedicareStr(colBalance)
End Function

Public Function GetInsureBalanceStr(objBalanceMoneys As BalanceMoneys) As String
    '��ȡ���ս��㴮,"���㷽ʽ|���||...."
    Dim strBalance As String
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    If objBalanceMoneys Is Nothing Then Exit Function
    For Each objItem In objBalanceMoneys
        strBalance = strBalance & "||" & objItem.���㷽ʽ & "|" & objItem.��Ч���
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetInsureBalanceStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
    ByVal strBalance As String)
    '���ܣ�����ָ����ŵ���ָ�����ս��㷽ʽ����Чֵ
    '������
    '       strBalance-���ݽ��㷽ʽ�ַ������ý��㷽ʽ��¼������ʽ�����㷽ʽ1|���1||���㷽ʽ2|���2||...
    '˵�����ú�����colBalanceΪ׼����,����ҽ�������շ�Ҳ��
    '˵������������ҽ���շ��޸ı��ս���������۵�ҽ���շ����ø����ʻ��Ƚ�����
    Dim i As Long
    Dim varBalance As Variant, varTemp As Variant
    Dim blnFind As Boolean
    Dim objItem As BalanceMoney, objBalanceMoneys As BalanceMoneys
    
    If strBalance = "" Then Exit Sub
    
    Set objBalanceMoneys = colBalance(intPage)
    
    '��ʽ�����㷽ʽ1|���1||���㷽ʽ2|���2||...
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(varBalance(i) & "|||", "|")
        blnFind = False
        For Each objItem In objBalanceMoneys
            If objItem.���㷽ʽ = varTemp(0) Then
                objItem.��Ч��� = varTemp(1)
                blnFind = True: Exit For
            End If
        Next
            
        If Not blnFind Then
            Set objItem = New BalanceMoney
            objItem.���㷽ʽ = varTemp(0)
            objItem.ԭʼ��� = varTemp(1)
            objItem.�����޸� = Val(varTemp(2)) = 1
            objItem.��Ч��� = varTemp(1)
            objBalanceMoneys.AddItem objItem
        End If
    Next

    colBalance.Remove intPage '����Ԫ�ز���ֱ���޸�
    If colBalance.Count >= intPage Then
        colBalance.Add objBalanceMoneys, , intPage
    Else
        colBalance.Add objBalanceMoneys
    End If
End Sub

Public Function zlInsureCheck(ByVal strԤ���� As String, ByVal strAdvance As String) As Boolean
    '��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:
    '   strԤ����-���ս���
    '   strAdvance-ҽ�����صĽ���
    '˵����
    '   ��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    Dim blnFind  As Boolean, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    If strAdvance = "" Or strԤ���� = strAdvance Then Exit Function
    
    zlInsureCheck = True
    
    varData = Split(strԤ����, "||")
    varData1 = Split(strAdvance, "||")
    If UBound(varData) <> UBound(varData1) Then Exit Function
    
    For i = 0 To UBound(varData)
        blnFind = False
        varTemp = Split(varData(i), "|")
        For j = 0 To UBound(varData1)
            varTemp1 = Split(varData1(j), "|")
            If varTemp(0) = varTemp1(0) Then
                blnFind = True
                If Val(varTemp(1)) <> Val(varTemp1(1)) Then Exit Function
            End If
        Next
        If Not blnFind Then Exit Function
    Next
    zlInsureCheck = False
End Function

Public Function zlExecuteInsureSwap(ByVal bytMode As Byte, ByVal objPati As clsPatientInfo, _
    ByVal intInsure As Integer, ByVal str�������� As String, ByVal blnOnlyBalanceSuccessedNo As Boolean, _
    ByVal lng����ID As Long, ByVal lng������� As Long, objBalanceBills As BalanceBills, _
    ByRef blnCommit As Boolean, Optional ByRef strSavedNos As String, Optional ByRef lngSavedBillCount As Long, _
    Optional ByRef blnYbBalanced As Boolean, Optional ByRef strErrMsg As String) As Boolean
    'ҽ������
    '��Σ�
    '   bytMode ҽ������ģʽ��0-�൥��һ�ν���,1-�൥��һ�ν���ֵ����˷�,2-�൥�ݷֵ��ݽ���
    '   blnOnlyBalanceSuccessedNo �൥�ݷֵ��ݽ���ʱ�Ƿ�ֻ��ҽ������ɹ������շ�
    '   strSavedNos,lngSavedBillCount �൥�ݷֵ��ݽ���ʱ�ѽ���ɹ��ĵ������
    '   blnYbBalanced �൥�ݷֵ��ݽ���ʱ��ҽ������ɹ������շ�
    '˵��:��Ҫ�������������,�����˷Ѻ�,�ù������ύ,����Ҫ�������ύ
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    Dim colBalance As Collection, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAdvance As String, strAdvanceOld  As String
    Dim cur����֧�� As Currency, curҽ������ As Currency
    Dim curȫ�Ը� As Currency, cur���Ը� As Currency
    Dim strAllԤ���� As String, strԤ���� As String, str���㷽ʽ As String
    Dim rsBalance As ADODB.Recordset, objBill As BalanceBill
    Dim p As Long, i As Long, blnFind As Boolean
    Dim varAdvance As Variant, varItem As Variant
    Dim blnCurrentCommit As Boolean
    
    On Error GoTo ErrHandler
    blnCommit = False: strSavedNos = ""
    blnYbBalanced = False: strErrMsg = ""
    If intInsure = 0 Then gcnOracle.RollbackTrans: Exit Function
    
    blnTrans = True
    strAllԤ���� = GetInsureBalanceStrAll(objBalanceBills)
    '�ȱ���Ԥ������
    Call SaveInsureBalance(objPati, lng����ID, strAllԤ����)
    
    '2-�൥�ݷֵ��ݽ���
    If bytMode = 2 Then
        Set colBalance = New Collection
        Set rsBalance = zlGetBalanceDetail(0, lng����ID, 1)
        
        For p = 1 To objBalanceBills.Count
            colBalance.Add New BalanceMoneys
            Set objBill = objBalanceBills(p)
            
            '�����ŵ����Ƿ��ѳɹ�ҽ������
            str���㷽ʽ = GetYBBalanceNo(rsBalance, objBill.NO)
            
            If str���㷽ʽ <> "" Then
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
                strSavedNos = strSavedNos & "," & objBill.NO
            Else
                strAdvance = lng������� & "|" & objBill.NO
                strAdvanceOld = strAdvance
                
                strԤ���� = GetInsureBalanceStr(objBill.Ԥ����)
                Call SaveInsureBalanceDetail(lng����ID, objBill.NO, strԤ����)
                
                cur����֧�� = GetInsureBalanceSum(objBill.Ԥ����, str��������)
                curҽ������ = GetInsureBalanceSum(objBill.Ԥ����, "ҽ������")
                curȫ�Ը� = objBill.ȫ�Ը�
                cur���Ը� = objBill.���Ը�
                
                If Not gclsInsure.ClinicSwap(lng����ID, cur����֧��, curҽ������, curȫ�Ը�, cur���Ը�, _
                    intInsure, strAdvance) Then
                    If blnOnlyBalanceSuccessedNo Then GoTo ErrHandler:
                    gcnOracle.RollbackTrans
                    If blnCurrentCommit Then Call CorrectInsureErrBalance(objPati, lng����ID)  'ҽ������У��
                    Exit Function
                End If
                If strAdvance = strAdvanceOld Then strAdvance = ""
                blnTransMedicare = True
                
                If zlInsureCheck(strԤ����, strAdvance) Then
                    Call SaveInsureBalanceDetail(lng����ID, objBill.NO, strAdvance)
                    strԤ���� = strAdvance
                End If
                
                Call SetBalanceVal(colBalance, p, strԤ����)
                strSavedNos = strSavedNos & "," & objBill.NO
                
                gcnOracle.CommitTrans: blnTrans = False
                blnCommit = True: blnCurrentCommit = True
                
                Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
                blnTransMedicare = False
                
                gcnOracle.BeginTrans: blnTrans = True
            End If
        Next
        strAdvance = GetMedicareStr(colBalance)
        
    '1-�൥��һ�ν���ֵ����˷�
    ElseIf bytMode = 1 Then
        Set colBalance = New Collection
        strAdvance = lng�������
        strAdvanceOld = strAdvance
        
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            strԤ���� = GetInsureBalanceStr(objBill.Ԥ����)
            Call SaveInsureBalanceDetail(lng����ID, objBill.NO, strԤ����)
            
            cur����֧�� = cur����֧�� + GetInsureBalanceSum(objBill.Ԥ����, str��������)
            curҽ������ = curҽ������ + GetInsureBalanceSum(objBill.Ԥ����, "ҽ������")
            curȫ�Ը� = curȫ�Ը� + objBill.ȫ�Ը�
            cur���Ը� = cur���Ը� + objBill.���Ը�
        Next
        
        If Not gclsInsure.ClinicSwap(lng����ID, cur����֧��, curҽ������, curȫ�Ը�, cur���Ը�, _
            intInsure, strAdvance) Then gcnOracle.RollbackTrans: Exit Function
        If strAdvance = strAdvanceOld Then strAdvance = ""
        blnTransMedicare = True
        
        'NO:���㷽ʽ,���|���㷽ʽ,���|...||NO:���㷽ʽ,���|���㷽ʽ,���|...||...
        varAdvance = Split(strAdvance, "||")
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            '�������ĳһ�ŵ���û�з��ض�Ӧ������Ϣ���Ͱ�Ԥ����������
            blnFind = False
            For i = 0 To UBound(varAdvance)
                If InStr(varAdvance(i), ":") = 0 Then
                    strErrMsg = "ҽ����������ʽ����ȷ��"
                    Exit Function
                End If
                
                varItem = Split(varAdvance(i), ":")
                If objBill.NO = varItem(0) Then
                    str���㷽ʽ = Replace(Replace(varItem(1), "|", "||"), ",", "|")
                    blnFind = True: Exit For
                End If
            Next
            
            If blnFind Then
                'ֱ������ҽ�������������Ƿ���ҪУ��
                Call SaveInsureBalanceDetail(lng����ID, objBill.NO, str���㷽ʽ)
            Else
                str���㷽ʽ = GetInsureBalanceStr(objBill.Ԥ����)
            End If
            
            colBalance.Add New BalanceMoneys
            SetBalanceVal colBalance, p, str���㷽ʽ
        Next
        strAdvance = GetMedicareStr(colBalance)
    
    '0-�൥��һ�ν���
    Else
        strAdvance = lng�������
        strAdvanceOld = strAdvance
        
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            cur����֧�� = cur����֧�� + GetInsureBalanceSum(objBill.Ԥ����, str��������)
            curҽ������ = curҽ������ + GetInsureBalanceSum(objBill.Ԥ����, "ҽ������")
            curȫ�Ը� = curȫ�Ը� + objBill.ȫ�Ը�
            cur���Ը� = cur���Ը� + objBill.���Ը�
        Next
        
        If Not gclsInsure.ClinicSwap(lng����ID, cur����֧��, curҽ������, _
            curȫ�Ը�, cur���Ը�, intInsure, strAdvance) Then gcnOracle.RollbackTrans: Exit Function
        If strAdvance = strAdvanceOld Then strAdvance = ""
        blnTransMedicare = True
    End If
    
    'У������Ľ�����
    If zlInsureCheck(strAllԤ����, strAdvance) Then
        Call SaveInsureBalance(objPati, lng����ID, strAdvance)
    End If
    Call InsureBalanceOver(lng����ID)
    gcnOracle.CommitTrans: blnTrans = False
    
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
    End If
    zlExecuteInsureSwap = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, intInsure)
    End If
    
    If bytMode = 2 And strSavedNos <> "" Then
        '105338:���ֽ���ɹ���ֻ�Խ���ɹ��ⲿ�ֵ����շ�
        If blnOnlyBalanceSuccessedNo Then
            On Error GoTo LastErrHandler
            strSavedNos = Mid(strSavedNos, 2)
            lngSavedBillCount = p - 1
            
            strAdvance = GetMedicareStr(colBalance)
            gcnOracle.BeginTrans: blnTrans = True
            '1.ɾ��δ�ɹ��ķ��õ��ݣ��ָ�Ϊ���۵�
            For i = objBalanceBills.Count To p Step -1
                Set objBill = objBalanceBills(i)
                If InStr("," & strSavedNos & ",", "," & objBill.NO & ",") = 0 Then
                    Call CancelBillBalance(lng����ID, objBill.NO)
                End If
            Next
            
            '2.У��ҽ������
            Call SaveInsureBalance(objPati, lng����ID, strAdvance)
            Call InsureBalanceOver(lng����ID)
            gcnOracle.CommitTrans: blnTrans = False
            blnYbBalanced = True: Exit Function
        ElseIf blnCurrentCommit Then
            Call CorrectInsureErrBalance(objPati, lng����ID) 'ҽ������У��
        End If
    ElseIf Err <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Exit Function
LastErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExecuteInsureDel(ByVal bytMode As Byte, objPati As clsPatientInfo, _
    ByVal intInsure As Integer, ByVal strSavedAdvance As String, ByVal lng����ID As Long, _
    ByVal lng����ID As Long, Optional ByRef blnCommit As Boolean, _
    Optional rsBillDelRecord As ADODB.Recordset, _
    Optional ByVal bln����������� As Boolean, Optional ByVal str�����ʻ� As String) As Boolean
    'ҽ������
    '��Σ�
    '   bytMode ҽ���˷�ģʽ��0-�൥��һ�ν���,1-�൥�ݷֵ����˷�
    '   strSavedAdvance ��Ҫ������ԭʼҽ��������Ϣ
    '   rsBillDelRecord��intInsure��bln����������ϡ�str�����ʻ� bytMode=1ʱ����
    '˵��:��Ҫ�������������,�����˷Ѻ�,�ù������ύ,����Ҫ�������ύ
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    Dim colBalance As Collection, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAdvance As String, strAdvanceOld  As String
    Dim str���㷽ʽ As String, p As Long
    Dim rsBalance As ADODB.Recordset, rsOldBalance As ADODB.Recordset
    Dim strKey As String, cllNo As Collection
    Dim strNo As String
    Dim blnCurrentCommit As Boolean
    
    On Error GoTo ErrHandler
    blnTrans = True
    strAdvance = lng����ID & "|" & "0"
    
    If bytMode = 0 Then '0-�൥��һ�ν���
        If Not gclsInsure.ClinicDelSwap(lng����ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        If strAdvance = lng����ID & "|" & "0" Then strAdvance = ""
        blnTransMedicare = True
        
    ElseIf bytMode = 1 Then '1-�൥�ݷֵ����˷�
        Set colBalance = New Collection
        Set rsOldBalance = zlGetBalanceDetail(0, lng����ID, 1)
        Set rsBalance = zlGetBalanceDetail(0, lng����ID, 1)
        
        strAdvanceOld = strAdvance
        
        '93337,�˷�ʱ�����ݺŵ�����нӿڵ���
        Set cllNo = New Collection
        With rsBillDelRecord
            .Filter = ""
            .Sort = "No Desc"
            Do While Not .EOF
                strKey = "_" & Nvl(!NO)
                If CollectionExitsValue(cllNo, strKey) = False Then
                    cllNo.Add Nvl(!NO), strKey
                End If
                rsBillDelRecord.MoveNext
            Loop
        End With
        
        For p = 1 To cllNo.Count
            colBalance.Add New BalanceMoneys
            strNo = cllNo(p)
            
            '�����ŵ����Ƿ��ѳɹ�ҽ������
            str���㷽ʽ = GetYBBalanceNo(rsBalance, strNo)
            
            If str���㷽ʽ <> "" Then
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            Else
                str���㷽ʽ = GetYBBalanceNo(rsOldBalance, strNo, True, objPati.����ID, intInsure, True, bln�����������, str�����ʻ�)
                'str���㷽ʽ Ϊ�գ���ʾҽ����֧��ҽ������
                If str���㷽ʽ <> "" Then
                    Call SaveInsureBalanceDetail(lng����ID, strNo, str���㷽ʽ)
                    
                    strAdvance = strAdvanceOld & "|" & strNo
                    '��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
                    If Not gclsInsure.ClinicDelSwap(lng����ID, True, intInsure, strAdvance) Then
                        gcnOracle.RollbackTrans
                        If blnCurrentCommit Then Call CorrectInsureErrBalance(objPati, lng����ID, True)  'ҽ������У��
                        Exit Function
                    End If
                    If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
                    blnTransMedicare = True
                    
                    If zlInsureCheck(str���㷽ʽ, strAdvance) Then
                        Call SaveInsureBalanceDetail(lng����ID, strNo, strAdvance)
                        str���㷽ʽ = strAdvance
                    End If
                    Call SetBalanceVal(colBalance, p, str���㷽ʽ)
                    gcnOracle.CommitTrans: blnTrans = False
                    blnCommit = True
                    blnCurrentCommit = True
                    
                    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
                    blnTransMedicare = False
                    
                    gcnOracle.BeginTrans: blnTrans = True
                End If
            End If
        Next
        
        'ȫ���ɹ��������ܵĽ��㷽ʽ
        strAdvance = GetMedicareStr(colBalance)
    End If
    
    'У������Ľ�����
    If zlInsureCheck(strSavedAdvance, strAdvance) Then
        Call SaveInsureBalance(objPati, lng����ID, strAdvance, True)
    End If
    Call InsureBalanceOver(lng����ID)
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    End If
    zlExecuteInsureDel = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    If bytMode = 1 And blnCurrentCommit Then
        Call CorrectInsureErrBalance(objPati, lng����ID, True)   'ҽ������У��
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveInsureBalanceDetail(ByVal lng����ID As Long, ByVal strNo As String, _
    ByVal strBalance As String, Optional cllPro As Collection)
    '����ҽ��������ϸ
    Dim strSql As String
    
    'Zl_ҽ��������ϸ_Insert(
    strSql = "Zl_ҽ��������ϸ_Insert( "
    '  ����id_In   ҽ��������ϸ.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  No_In       ҽ��������ϸ.No%Type,
    strSql = strSql & "'" & strNo & "',"
    '  ���㷽ʽ_In Varchar2,
    strSql = strSql & "'" & strBalance & "',"
    '  ��ע_In     ҽ��������ϸ.��ע%Type := Null,
    strSql = strSql & "" & "NULL" & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
End Sub

Public Sub SaveInsureBalance(ByVal objPati As clsPatientInfo, ByVal lng����ID As Long, _
    ByVal strBalance As String, Optional ByVal blnDel As Boolean, _
    Optional ByVal lng��������ID As Long, Optional cllPro As Collection)
    '����ҽ����������
    Dim strSql As String
    
    If blnDel Then
        'Zl_�����˷ѽ���_Modify_S(
        strSql = "Zl_�����˷ѽ���_Modify_S("
        '  ��������_In      Number,
        strSql = strSql & "" & 3 & ","
        '  ����id_In        ������ü�¼.����id%Type,
        strSql = strSql & "" & objPati.����ID & ","
        '  ����_In          ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "',"
        '  �Ա�_In          ����Ԥ����¼.�Ա�%Type,
        strSql = strSql & "'" & objPati.�Ա� & "',"
        '  ����_In          ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "',"
        '  �����_In        ����Ԥ����¼.�����%Type,
        strSql = strSql & "'" & objPati.����� & "',"
        '  סԺ��_In        ����Ԥ����¼.סԺ��%Type,
        strSql = strSql & "'" & objPati.סԺ�� & "',"
        '  ���ʽ����_In  ����Ԥ����¼.���ʽ����%Type,
        strSql = strSql & "'" & objPati.ҽ�Ƹ��ʽ & "',"
        '  ����id_In        ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng����ID & ","
        '  ���㷽ʽ_In      Varchar2
        strSql = strSql & "'" & strBalance & "',"
        '  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  �����id_In      ����Ԥ����¼.�����id%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ����_In          ����Ԥ����¼.����%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ����˷�_In      Number := 0,
        strSql = strSql & "" & "0" & ","
        '  ԭ����id_In      ����Ԥ����¼.����id%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ʣ��תԤ��_In    Number := 0,
        strSql = strSql & "" & "0" & ","
        '  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ��Ԥ������ids_In Varchar2 := Null,
        strSql = strSql & "" & "NULL" & ","
        '  ��������id_In    ����Ԥ����¼.��������id%Type := Null,
        strSql = strSql & "" & IIf(lng��������ID = 0, "NULL", lng��������ID) & ")"
    Else
        'Zl_�����շѽ���_Modify_S
        strSql = "Zl_�����շѽ���_Modify_S("
        '  ��������_In   Number,
        strSql = strSql & "" & 2 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSql = strSql & "" & objPati.����ID & ","
        '  ����_In          ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "',"
        '  �Ա�_In          ����Ԥ����¼.�Ա�%Type,
        strSql = strSql & "'" & objPati.�Ա� & "',"
        '  ����_In          ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "',"
        '  �����_In        ����Ԥ����¼.�����%Type,
        strSql = strSql & "'" & objPati.����� & "',"
        '  סԺ��_In        ����Ԥ����¼.סԺ��%Type,
        strSql = strSql & "'" & objPati.סԺ�� & "',"
        '  ���ʽ����_In  ����Ԥ����¼.���ʽ����%Type,
        strSql = strSql & "'" & objPati.ҽ�Ƹ��ʽ & "',"
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSql = strSql & "'" & strBalance & "')"
    End If
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
End Sub

Public Sub InsureBalanceOver(ByVal lng����ID As Long, _
    Optional cllPro As Collection)
    'ҽ����ɽ��㣬����У�Ա�־
    Dim strSql As String
    On Error GoTo errHandle
    
    'Zl_���������շ�_ҽ������(
    strSql = "Zl_���������շ�_ҽ������( "
    '  ����id_In   ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    strSql = strSql & "" & "NULL" & ","
    '  ���ս���_In Varchar2
    strSql = strSql & "" & "NULL" & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function Get���ʱ������(ByVal curʵ�պϼ� As Currency, ByVal cur����Ԥ�� As Currency, _
    ByVal dbl������� As Double, ByVal dbl����͸֧ As Double) As Currency
    '��������ʻ�֧�����
    If RoundEx(curʵ�պϼ�, 6) <= 0 Then Get���ʱ������ = 0: Exit Function
    
    If RoundEx(dbl������� + dbl����͸֧, 6) <= 0 Then '��ǰ�������(��͸֧)
        Get���ʱ������ = 0
    Else
        If RoundEx(dbl������� + dbl����͸֧, 6) >= RoundEx(cur����Ԥ��, 6) Then '������֧����Χ���㹻(��͸֧)
            Get���ʱ������ = cur����Ԥ��
        Else
            Get���ʱ������ = dbl������� + dbl����͸֧
        End If
    End If
End Function

Public Function GetYBBalance(rsBalance As ADODB.Recordset, ByVal lng����ID As Long, _
    Optional ByVal blnDelCheck As Boolean, Optional ByVal lng����ID As Long, Optional ByVal intInsure As Integer, _
    Optional ByVal bln����������� As Boolean, Optional ByVal str�����ʻ� As String, _
    Optional ByVal blnDel As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ԭ���㷽ʽ�ͽ�����
    '����:���ؽ�����Ϣ,��ʽ:���㷽ʽ|������||...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    
    On Error GoTo errHandle
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    rsBalance.Filter = "����=2 and ����ID=" & lng����ID
    If rsBalance.RecordCount = 0 Then Exit Function
    With rsBalance
        Do While Not .EOF
            If blnDelCheck Then
                '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                If bln����������� Then
                    If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, Nvl(!���㷽ʽ)) Then
                        str���㷽ʽ = str���㷽ʽ & "||" & Nvl(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!��Ԥ��))
                    End If
                Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                    If !���㷽ʽ <> str�����ʻ� Then
                        str���㷽ʽ = str���㷽ʽ & "||" & Nvl(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!��Ԥ��))
                    End If
                End If
            Else
                str���㷽ʽ = str���㷽ʽ & "||" & Nvl(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!��Ԥ��))
            End If
            
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 3)
    GetYBBalance = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetYBBalanceNo(rsBalance As ADODB.Recordset, Optional ByVal strNos As String, _
    Optional ByVal blnDelCheck As Boolean, Optional ByVal lng����ID As Long, _
    Optional ByVal intInsure As Integer, Optional ByVal blnDel As Boolean, _
    Optional ByVal bln����������� As Boolean, Optional ByVal str�����ʻ� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݻ�ȡҽ��ԭ���㷽ʽ�ͽ�����
    '������
    '   strNOs - ���ݺ�,����ö��Ÿ�����A0001,A0002,...
    '   blnDelCheck - �Ƿ������������������
    '����:���ؽ�����Ϣ,��ʽ:���㷽ʽ|������||...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, varNos As Variant, strFilter As String
    Dim i As Integer, p As Integer
    Dim colBalance As Collection, strPreNo As String
    
    If blnDelCheck And intInsure = 0 Then Exit Function
    If rsBalance Is Nothing Then Exit Function
    
    varNos = Split(strNos, ",")
    For i = 0 To UBound(varNos)
        strFilter = strFilter & " Or No='" & varNos(i) & "'"
    Next
    If strFilter <> "" Then strFilter = Mid(strFilter, 4)
    rsBalance.Filter = strFilter
    If rsBalance.RecordCount = 0 Then Exit Function
    
    '�ֶ�:����id,NO,���㷽ʽ,���,�����id,��������id,������ˮ��,����˵��,ҽ��,��������
    rsBalance.Sort = "No"
    Set colBalance = New Collection
    p = 1: colBalance.Add New BalanceMoneys
    With rsBalance
        strPreNo = Nvl(!NO)
        Do While Not .EOF
            If strPreNo <> Nvl(!NO) Then
                p = p + 1: colBalance.Add New BalanceMoneys
                strPreNo = Nvl(!NO)
            End If
            If blnDelCheck Then
                '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                If bln����������� Then
                    If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, !���㷽ʽ) Then
                        str���㷽ʽ = Nvl(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!���))
                    End If
                Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                    If !���㷽ʽ <> str�����ʻ� Then
                        str���㷽ʽ = Nvl(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!���))
                    End If
                End If
            Else
                str���㷽ʽ = Nvl(!���㷽ʽ) & "|" & Val(Nvl(!���))
            End If
            
            Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            .MoveNext
        Loop
    End With
    GetYBBalanceNo = GetMedicareStr(colBalance)
End Function

Public Function GetYBTOCash(rsBalance As ADODB.Recordset, _
    ByVal lng����ID As Long, ByVal intInsure As Integer, _
    ByVal bln����������� As Boolean, ByVal str�����ʻ� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ʹ���ֽ�֧���Ľ��㷽ʽ(����ö��ŷָ�)
    '����:���ؽ��㷽ʽ,����ö��ŷָ�:�����ʻ�,ҽ������...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    rsBalance.Filter = "����=" & Enum_BalanceType.ҽ��
    If rsBalance.RecordCount = 0 Then Exit Function
    With rsBalance
        Do While Not .EOF
            '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
            If bln����������� Then
                If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, Nvl(!���㷽ʽ)) Then
                    str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
                End If
            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                If !���㷽ʽ = str�����ʻ� Then
                    str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
                End If
            End If
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    GetYBTOCash = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CorrectInsureErrBalance(ByVal objPati As clsPatientInfo, _
    ByVal lng����ID As Long, Optional ByVal blnDel As Boolean) As Boolean
    'ҽ������У��
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset, strBalance As String
    Dim rsBalanceSaved As ADODB.Recordset, strBalanceSaved As String
    
    On Error GoTo ErrHandler
    strSql = "Select 1" & _
            " From ����Ԥ����¼ A, ���㷽ʽ B" & _
            " Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And ����id = [1] And a.�����ID Is Null " & _
            "       And Nvl(a.У�Ա�־, 0) = 1 And Rownum < 2"
    strSql = strSql & "Union All" & _
            " Select 1" & _
            " From ���ս����¼" & _
            " Where ��¼id = [1] " & _
            "       And Not Exists(Select 1 " & _
            "                      From ����Ԥ����¼ A, ���㷽ʽ B" & _
            "                      Where a.���㷽ʽ = b.���� And a.����id = ��¼id " & _
            "                            And b.���� In (3, 4) And a.�����ID Is Null)" & _
            "       And �����ID Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ������ҪУ�Ե�ҽ������", lng����ID)
    If rsTemp.EOF Then CorrectInsureErrBalance = True: Exit Function
    
    '��ͨ����ҽ��������ϸ������У��
    Set rsBalance = zlGetBalanceDetail(0, lng����ID, 1)
    strBalance = GetYBBalanceNo(rsBalance)
    
    If strBalance = "" Then
        strSql = "Select a.����ID,a.���㷽ʽ,a.���" & _
            " From ���ս�����ϸ A ,���㷽ʽ C" & _
            " Where a.���㷽ʽ=c.���� And c.���� in (3,4) And a.����id =[1] And a.��־=1 " & _
            " Order by ���㷽ʽ"
        'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�",�����ſ���ҽ����Ľ��㷽ʽ
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
        Do While Not rsTemp.EOF
            strBalance = strBalance & "||" & Nvl(rsTemp!���㷽ʽ) & "|" & Val(Nvl(rsTemp!���))
            rsTemp.MoveNext
        Loop
        If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    End If
    'û�к˶�����,ֱ�ӷ���
    If strBalance = "" Then CorrectInsureErrBalance = True: Exit Function
    
    '����Ƿ���ҪУ��
    Set rsBalanceSaved = zlFromIDGetChargeBalance(0, lng����ID)
    strBalanceSaved = GetYBBalance(rsBalanceSaved, lng����ID)
    If zlInsureCheck(strBalanceSaved, strBalance) Then
        Call SaveInsureBalance(objPati, lng����ID, strBalance, blnDel)
    End If
    
    CorrectInsureErrBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakeBillRecord(objBill As ExpenseBill, ByVal bln���� As Boolean, ByVal intPage As Integer, _
    ByVal strDate As String, ByVal str�ѱ� As String, ByVal strInvoice As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)
    '���:intPage=�൥���շ�ģʽʱ��ָ���ĵ���,���Ϊ��,��ʾȫ������
    '        strDate=����ʱ��,
    '        strInvoice=Ʊ�ݺ�
    '����:
    '����:ҽ��������ݵ����ݼ�(�������(1--n),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��)
    '����:���˺�
    '����:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSql As String
    Dim dbl���� As Double, curʵ�� As Currency, curͳ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand:
    Set rsTmp = MakePreSwapData()
    
    intStartPage = IIf(intPage <= 0, 1, intPage)
    intPages = IIf(intPage <= 0, objBill.Pages.Count, intPage)
    For p = intStartPage To intPages
         If objBill.Pages(p).NO <> "" Then       '��ȡ���ǻ��۵�
            '��ȡ�Ļ��۵�(�ۼ۵�λ)
            strSql = _
            "Select '" & strInvoice & "' as ʵ��Ʊ��,NO,Nvl( �۸񸸺�, ���) as ���,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
                    objBill.����ID & " As ����ID,'" & str�ѱ� & "' As �ѱ�,�շ����,�վݷ�Ŀ,���㵥λ,������," & _
            "       �շ�ϸĿID,���մ���ID As ����֧������ID,Nvl(������Ŀ��,0) As �Ƿ�ҽ��,���ձ���," & _
            "       Avg(Nvl(����,0)*����) As ����,Avg(��׼����) As ����," & _
            "       Sum(ʵ�ս��) As ʵ�ս��,Sum(ͳ����) As ͳ����,ժҪ," & _
                    IIf(bln����, "1", "0") & " as �Ƿ���,��������ID,ִ�в���ID" & _
            " From ������ü�¼" & _
            " Where ��¼����=1 And NO=[1]" & _
            " Group By Nvl(�۸񸸺�,���),�շ����,�վݷ�Ŀ,���㵥λ,������," & _
            "       �շ�ϸĿID,���մ���ID,Nvl(������Ŀ��,0),���ձ���,ժҪ,��������ID,ִ�в���ID,NO" & _
            " Order by  ��� "
            Set rsNo = zlDatabase.OpenSQLRecord(strSql, "��ȡ���۵�����-ҽ��", objBill.Pages(p).NO)
            If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
            Do While Not rsNo.EOF
                rsTmp.AddNew
                rsTmp!������� = p
                rsTmp!�ѱ� = str�ѱ�
                rsTmp!NO = Nvl(rsNo!NO)   '����ȡ���۵�ʱ����ֵ
                rsTmp!��� = Val(Nvl(rsNo!���))   '����ȡ���۵�ʱ����ֵ
                rsTmp!ʵ��Ʊ�� = strInvoice
                rsTmp!����ʱ�� = CDate(strDate)
                rsTmp!����ID = IIf(objBill.����ID = 0, Null, objBill.����ID)
                rsTmp!�շ���� = Nvl(rsNo!�շ����)
                rsTmp!�վݷ�Ŀ = Nvl(rsNo!�վݷ�Ŀ)
                rsTmp!������ = Nvl(rsNo!������)
                rsTmp!�շ�ϸĿID = Val(Nvl(rsNo!�շ�ϸĿID))
                rsTmp!���㵥λ = Nvl(rsNo!���㵥λ)
                rsTmp!���� = Val(Nvl(rsNo!����))
                rsTmp!���� = Val(Nvl(rsNo!����))
                rsTmp!ʵ�ս�� = Val(Nvl(rsNo!ʵ�ս��))
                rsTmp!ͳ���� = Val(Nvl(rsNo!ͳ����))
                rsTmp!����֧������ID = IIf(Val(Nvl(rsNo!����֧������ID)) = 0, Null, Val(Nvl(rsNo!����֧������ID)))
                rsTmp!�Ƿ�ҽ�� = Val(Nvl(rsNo!�Ƿ�ҽ��))
                rsTmp!���ձ��� = Nvl(rsNo!���ձ���)
                rsTmp!ժҪ = Nvl(rsNo!ժҪ)
                rsTmp!�Ƿ��� = IIf(bln����, 1, 0)
                rsTmp!��������ID = Val(Nvl(rsNo!��������ID))
                rsTmp!ִ�в���ID = Val(Nvl(rsNo!ִ�в���ID))
                rsTmp.Update
                rsNo.MoveNext
            Loop
         Else
            For i = 1 To objBill.Pages(p).Details.Count
                dbl���� = 0: curʵ�� = 0: curͳ�� = 0
                With objBill.Pages(p).Details(i)
                    For j = 1 To .InComes.Count
                        dbl���� = dbl���� + .InComes(j).��׼����
                        curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                        curͳ�� = curͳ�� + .InComes(j).ͳ����
                    Next
                    rsTmp.AddNew
                    rsTmp!������� = p
                    rsTmp!�ѱ� = str�ѱ�
                    rsTmp!NO = ""   '����ȡ���۵�ʱ����ֵ
                    rsTmp!��� = i
                    rsTmp!ʵ��Ʊ�� = strInvoice
                    rsTmp!����ʱ�� = CDate(strDate)
                    rsTmp!����ID = IIf(objBill.����ID = 0, Null, objBill.����ID)
                    rsTmp!�շ���� = .�շ����
                    If .InComes.Count > 0 Then
                        rsTmp!�վݷ�Ŀ = .InComes(1).�վݷ�Ŀ
                    Else
                        rsTmp!�վݷ�Ŀ = Null
                    End If
                    rsTmp!������ = objBill.Pages(p).������
                    
                    rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                    
                    rsTmp!���㵥λ = .���㵥λ
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                        '��ҩ����λת��Ϊ�ۼ۵�λ
                        rsTmp!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                        rsTmp!���� = Format(dbl���� / .Detail.ҩ����װ, gstrFeePrecisionFmt)
                    Else
                        rsTmp!���� = IIf(.���� = 0, 1, .����) * .����
                        rsTmp!���� = Format(dbl����, gstrFeePrecisionFmt)
                    End If
                    rsTmp!ʵ�ս�� = Format(curʵ��, gstrDec)
                    rsTmp!ͳ���� = Format(curͳ��, gstrDec)
                    rsTmp!����֧������ID = IIf(.���մ���ID = 0, Null, .���մ���ID)
                    rsTmp!�Ƿ�ҽ�� = IIf(.������Ŀ��, 1, 0)
                    rsTmp!���ձ��� = .���ձ���
                    rsTmp!ժҪ = .ժҪ
                    rsTmp!�Ƿ��� = IIf(bln����, 1, 0)
                    rsTmp!��������ID = objBill.Pages(p).��������ID
                    rsTmp!ִ�в���ID = .ִ�в���ID
                    rsTmp.Update
                End With
            Next
        End If
    Next
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeBillRecord = rsTmp
    Exit Function
Errhand::
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function MakePreSwapData() As ADODB.Recordset
    '����һ��Ԥ�����¼�ṹ
    '����:ҽ��������ݵ����ݼ��ṹ
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    With rsTmp.Fields
        .Append "�������", adBigInt, 50, adFldIsNullable
        .Append "�ѱ�", adVarChar, 50, adFldIsNullable
        .Append "NO", adVarChar, 8, adFldIsNullable
        .Append "���", adBigInt, , adFldIsNullable '����:42961
        .Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
        .Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
        .Append "����ID", adBigInt, , adFldIsNullable
        .Append "�շ����", adVarChar, 2, adFldIsNullable
        .Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        .Append "���㵥λ", adVarChar, 50, adFldIsNullable
        .Append "������", adVarChar, 100, adFldIsNullable
        .Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
        .Append "����", adDouble, , adFldIsNullable
        .Append "����", adDouble, , adFldIsNullable
        .Append "ʵ�ս��", adCurrency, , adFldIsNullable
        .Append "ͳ����", adCurrency, , adFldIsNullable
        .Append "����֧������ID", adBigInt, , adFldIsNullable
        .Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
        .Append "���ձ���", adVarChar, 50, adFldIsNullable
        .Append "ժҪ", adVarChar, 2000, adFldIsNullable
        .Append "�Ƿ���", adBigInt, , adFldIsNullable
        .Append "��������ID", adBigInt, , adFldIsNullable
        .Append "ִ�в���ID", adBigInt, , adFldIsNullable
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set MakePreSwapData = rsTmp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePreSwapDataFromDB(ByVal strInvoice As String, _
    ByVal lng����ID As Long, Optional ByVal strNo As String, _
    Optional ByVal bln������ As Boolean) As ADODB.Recordset
    '���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)����Ҫ���ȫ���ؽ�Ͳ�����
    '���:
    '   strInvoice-��ʼ��Ʊ��
    '����:
    '����:ҽ��������ݵ����ݼ�(�������(1--n),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��)
    Dim p As Integer, strSql As String
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim strWhere As String, strNos As String
    
    Err = 0: On Error GoTo Errhand:
    Set rsTmp = MakePreSwapData()
    If strNo <> "" Then strWhere = " And a.No=[2]"

    If bln������ Then
        strSql = _
            "Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, To_Char(Max(a.�Ǽ�ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��," & vbNewLine & _
            "       a.����id, a.�ѱ�, a.�շ����, a.�վݷ�Ŀ, a.���㵥λ, a.������, a.�շ�ϸĿid, a.���մ���id As ����֧������id," & vbNewLine & _
            "       Nvl(a.������Ŀ��, 0) As �Ƿ�ҽ��, a.���ձ���, Avg(Nvl(a.����, 0) * a.����) As ����, Avg(a.��׼����) As ����," & vbNewLine & _
            "       Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.ͳ����) As ͳ����, Max(a.ժҪ) As ժҪ," & vbNewLine & _
            "       Nvl(a.�Ӱ��־, 0) As �Ƿ���, a.��������id, a.ִ�в���id, a.����id" & vbNewLine & _
            "From ������ü�¼ A, (Select Distinct �շѽ���id From ���ò����¼ Where ����id = [1]) B" & vbNewLine & _
            "Where Mod(a.��¼����, 10) = 1 And a.����id = b.�շѽ���id" & strWhere & vbNewLine & _
            "Group By a.No, Nvl(a.�۸񸸺�, a.���), a.����id, a.�ѱ�, a.�շ����, a.�վݷ�Ŀ, a.���㵥λ, a.������," & vbNewLine & _
            "      a.�շ�ϸĿid, a.���մ���id, Nvl(a.������Ŀ��, 0), a.���ձ���, a.ժҪ, Nvl(a.�Ӱ��־, 0)," & vbNewLine & _
            "      a.��������id, a.ִ�в���id, a.����id"
    Else
        strSql = _
            "Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, To_Char(a.�Ǽ�ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��," & vbNewLine & _
            "       a.����id, a.�ѱ�, a.�շ����, a.�վݷ�Ŀ, a.���㵥λ, a.������, a.�շ�ϸĿid, a.���մ���id As ����֧������id," & vbNewLine & _
            "       Nvl(a.������Ŀ��, 0) As �Ƿ�ҽ��, a.���ձ���, Nvl(a.����, 0) * a.���� As ����, a.��׼���� As ����," & vbNewLine & _
            "       a.ʵ�ս��, a.ͳ����, a.ժҪ As ժҪ,Nvl(a.�Ӱ��־, 0) As �Ƿ���, a.��������id, a.ִ�в���id, a.����id" & vbNewLine & _
            "From ������ü�¼ A" & vbNewLine & _
            "Where Mod(a.��¼����, 10) = 1 And a.����id = [1]"
    End If
    
    strSql = _
        "Select '" & strInvoice & "' As ʵ��Ʊ��, a.No, a.���, Max(a.����ʱ��) As ����ʱ��, a.����id, a.�ѱ�, a.�շ����, a.�վݷ�Ŀ," & vbNewLine & _
        "       a.���㵥λ, a.������, a.�շ�ϸĿid, a.����֧������id, a.�Ƿ�ҽ��, a.���ձ���, Sum(a.����) As ����," & vbNewLine & _
        "       Max(a.����) As ����, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.ͳ����) As ͳ����, Max(a.ժҪ) As ժҪ," & vbNewLine & _
        "       Max(a.�Ƿ���) As �Ƿ���, Max(a.��������id) As ��������id, Max(a.ִ�в���id) As ִ�в���id" & vbNewLine & _
        "From (" & strSql & ") A" & vbNewLine & _
        "Group By a.No, a.���, a.����id, a.�ѱ�, a.�շ����, a.�վݷ�Ŀ, a.���㵥λ, a.������," & vbNewLine & _
        "      a.�շ�ϸĿid, a.����֧������id, a.�Ƿ�ҽ��, a.���ձ���" & vbNewLine & _
        "Having Nvl(Sum(a.����), 0) <> 0" & vbNewLine & _
        "Order By NO, ���"
    Set rsNo = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����շ�����-ҽ��", lng����ID, strNo)
    
    With rsNo
        p = 0: strNos = ""
        Do While Not rsNo.EOF
            If InStrEx(strNos, Nvl(!NO)) = False Then
                strNos = strNos & "," & Nvl(!NO)
                p = p + 1
            End If
            
            rsTmp.AddNew
            rsTmp!������� = p
            rsTmp!�ѱ� = !�ѱ�
            rsTmp!NO = Nvl(!NO)
            rsTmp!��� = Val(Nvl(!���))
            rsTmp!ʵ��Ʊ�� = Nvl(!ʵ��Ʊ��)
            rsTmp!����ʱ�� = !����ʱ��
            rsTmp!����ID = Val(Nvl(!����ID))
            rsTmp!�շ���� = Nvl(!�շ����)
            rsTmp!�վݷ�Ŀ = Nvl(!�վݷ�Ŀ)
            rsTmp!������ = Nvl(!������)
            rsTmp!�շ�ϸĿID = Val(Nvl(!�շ�ϸĿID))
            rsTmp!���㵥λ = Nvl(!���㵥λ)
            rsTmp!���� = Val(Nvl(!����))
            rsTmp!���� = Val(Nvl(!����))
            rsTmp!ʵ�ս�� = Val(Nvl(!ʵ�ս��))
            rsTmp!ͳ���� = Val(Nvl(!ͳ����))
            rsTmp!����֧������ID = IIf(Val(Nvl(!����֧������ID)) = 0, Null, Val(Nvl(!����֧������ID)))
            rsTmp!�Ƿ�ҽ�� = Val(Nvl(!�Ƿ�ҽ��))
            rsTmp!���ձ��� = Nvl(!���ձ���)
            rsTmp!ժҪ = Nvl(!ժҪ)
            rsTmp!�Ƿ��� = Val(Nvl(!�Ƿ���))
            rsTmp!��������ID = Val(Nvl(!��������ID))
            rsTmp!ִ�в���ID = Val(Nvl(!ִ�в���ID))
            rsTmp.Update
            
            rsNo.MoveNext
        Loop
    End With
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakePreSwapDataFromDB = rsTmp
    
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteInsureInfoUpdate(ByVal lng����ID As Long, ByVal intInsure As Integer, _
    ByRef objBalanceBills As BalanceBills) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ռ�¼�ı�����Ϣ
    '����:
    '   str���ս��-"ʵ�պϼ�;����ͳ��;ȫ�Ը�;���Ը�"
    '����:�������ռ�¼�ı�����Ϣ���³ɹ�����True�����򷵻�False
    '����:Ƚ����
    '����:2014-9-16
    '����:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsReCharge As ADODB.Recordset
    Dim strBXInfo As String, strPreNo As String
    Dim curʵ�ս�� As Currency, curͳ���� As Currency, bln������Ŀ As Boolean
    Dim blnTrans As Boolean, cllReChargePro As Collection
    Dim objBalanceBill As BalanceBill
    
    On Error GoTo Errhand
    Set objBalanceBills = New BalanceBills
    strSql = _
        "Select a.Id, a.No, a.���, a.����id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����, Nvl(a.ʵ�ս��, 0) As ʵ�ս��, a.ժҪ," & vbNewLine & _
        "       Nvl(a.������Ŀ��, 0) As ������Ŀ��, a.���մ���id, Nvl(a.ͳ����, 0) As ͳ����, a.���ձ���, a.��������" & vbNewLine & _
        "From ������ü�¼ A" & vbNewLine & _
        "Where Mod(a.��¼����, 10) = 1 And a.����id = [1]"
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSql, "��ȡ���շ��ü�¼", lng����ID)
    With rsReCharge
        If .RecordCount > 0 Then
            Set cllReChargePro = New Collection
            .Sort = "NO,���"
            Do While Not .EOF
                If strPreNo <> Nvl(!NO) Then
                    If strPreNo <> "" Then
                        objBalanceBills.AddItem objBalanceBill
                    End If
                    
                    Set objBalanceBill = New BalanceBill
                    objBalanceBill.NO = Nvl(!NO)
                    strPreNo = Nvl(!NO)
                End If
                
                '������Ŀ��(0/1);���մ���ID;����ͳ����;������Ŀ����;ժҪ;��������
                strBXInfo = gclsInsure.GetItemInsure(Nvl(!����ID), Nvl(!�շ�ϸĿID), Val(Nvl(!ʵ�ս��)), True, intInsure, _
                        Nvl(!ժҪ) & "||" & Val(Nvl(!����)))
                If strBXInfo <> "" Then
                    '  Zl_�����շѼ�¼_Update
                    strSql = "Zl_�����շѼ�¼_Update("
                    '  Id_In         In ������ü�¼.Id%Type,
                    strSql = strSql & Nvl(!ID) & ","
                    '  ���մ���id_In In ������ü�¼.���մ���id%Type,
                    strSql = strSql & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  ������Ŀ��_In In ������ü�¼.������Ŀ��%Type,
                    strSql = strSql & Val(Split(strBXInfo, ";")(0)) & ","
                    '  ���ձ���_In   In ������ü�¼.���ձ���%Type,
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  ��������_In   In ������ü�¼.��������%Type,
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  ͳ����_In   In ������ü�¼.ͳ����%Type,
                    strSql = strSql & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  ժҪ_In       In ������ü�¼.ժҪ%Type
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllReChargePro, strSql
                    
                    curͳ���� = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln������Ŀ = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    curͳ���� = Val(Nvl(!ͳ����))
                    bln������Ŀ = Val(Nvl(!������Ŀ��)) = 1
                End If
                
                'ͳ�Ʊ��ս��
                curʵ�ս�� = Val(Nvl(!ʵ�ս��))
                If curͳ���� = 0 Or Not bln������Ŀ Then
                    '��ԭʼ���Ϊ׼,���ֱܷҴ���
                    objBalanceBill.ȫ�Ը� = objBalanceBill.ȫ�Ը� + curʵ�ս��
                Else
                    objBalanceBill.����ͳ�� = objBalanceBill.����ͳ�� + curͳ����
                    '��ԭʼ���Ϊ׼,���ֱܷҴ���
                    objBalanceBill.���Ը� = objBalanceBill.���Ը� + curʵ�ս�� - curͳ����
                End If
                objBalanceBill.ʵ�պϼ� = objBalanceBill.ʵ�պϼ� + CCur(Val(Nvl(!ʵ�ս��)))
                
                rsReCharge.MoveNext
            Loop
            If strPreNo <> "" Then
                objBalanceBills.AddItem objBalanceBill
            End If
            
            'ִ�й���
            blnTrans = True
            zlExecuteProcedureArrAy cllReChargePro, "ִ�б�����Ϣ����", True, True
            blnTrans = False
        End If
    End With
    ExecuteInsureInfoUpdate = True
    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function CheckInsureCanCancel(ByVal intInsure As Integer, ByVal bln����������� As Boolean, ByVal lng������� As Long, ByVal lng����ID As Long) As Boolean
    '���ҽ�������Ƿ���������
    Dim strSql As String, rsBalance As ADODB.Recordset
    Dim strBalance As String
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then CheckInsureCanCancel = True: Exit Function
    If bln����������� = False Then
        ShowMsgbox "ҽ����֧������������ϣ��������ϱ��ν��㣡"
        Exit Function
    End If
        
    strSql = _
        "Select a.���㷽ʽ" & vbNewLine & _
        "From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
        "Where a.���㷽ʽ = b.���� And b.���� In(3,4) And a.�����ID Is Null And a.������� = [1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, "��ȡҽ����������", lng�������)
    Do While Not rsBalance.EOF
        If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, Nvl(rsBalance!���㷽ʽ)) Then
            strBalance = strBalance & "," & Nvl(rsBalance!���㷽ʽ)
        End If
        rsBalance.MoveNext
    Loop
    
    If strBalance <> "" Then
        ShowMsgbox "ҽ�����㷽ʽ��" & Mid(strBalance, 2) & "����֧�����ϣ��������ϱ��ν��㣡"
        Exit Function
    End If
    CheckInsureCanCancel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InsureSwapSuccess(ByVal intInsure As Integer, ByVal lng����ID As Long) As Boolean
    'ҽ�������Ƿ�ɹ�
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then InsureSwapSuccess = True: Exit Function
    strSql = _
        "Select 1" & vbNewLine & _
        "From ����Ԥ����¼ A, ���ս����¼ B, ���㷽ʽ C" & vbNewLine & _
        "Where a.����id = b.��¼id And a.���㷽ʽ = c.���� And c.���� In (3, 4) And Nvl(a.У�Ա�־, 0) = 2" & vbNewLine & _
        "      And a.����id = [2] And a.�����id Is Null And b.���� = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�ж�ҽ�������Ƿ�ɹ�", intInsure, lng����ID)
    InsureSwapSuccess = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistsInsureSwap(ByVal lng����ID As Long, ByRef intInsure As Integer) As Boolean
    '�Ƿ����ҽ������
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        "Select b.����" & vbNewLine & _
        "From ����Ԥ����¼ A, ���ս����¼ B, ���㷽ʽ C" & vbNewLine & _
        "Where a.����id = b.��¼id And a.���㷽ʽ = c.���� And c.���� In (3, 4)" & vbNewLine & _
        "      And a.����id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�ж�ҽ�������Ƿ�ɹ�", lng����ID)
    If rsTemp.EOF Then Exit Function
        
    intInsure = Val(Nvl(rsTemp!����))
    ExistsInsureSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetCustomPatiInsure(ByVal lng����ID As Long) As Integer
    '��ȡ�������࣬�ڲ���ʶ��ɹ�����ã�����������Զ�����ҽ�����ʶ��
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lng����ID = 0 Then Exit Function
    '���������֧��ҽ���򲻵����Զ������
    If GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", "") = "" Then Exit Function
    
    strSql = "Select Zl_Custom_Getpatiinsure([1]) As ���� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", lng����ID)
    If rsTemp.EOF Then Exit Function
        
    GetCustomPatiInsure = Val(Nvl(rsTemp!����))
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
