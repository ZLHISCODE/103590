VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsureSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstr�������� As String
Private mclsInsure As clsInsure

Private mlng����ID As Long
Private mlng���� As Long

Private mbln���ý������ As Boolean

Private mblnҽ������Ʊ�� As Boolean
Private mbln����������� As Boolean
Private mbln�൥�ݷֵ��ݽ��� As Boolean
Private mblnһ�ν���ֵ����˷� As Boolean

Public Property Set Insure(ByVal vData As clsInsure)
    Set mclsInsure = vData
End Property

Public Property Get Insure() As clsInsure
    Set Insure = mclsInsure
End Property

Public Property Let ����ID(ByVal vData As Long)
    mlng����ID = vData
End Property

Public Property Get ����ID() As Long
    ����ID = mlng����ID
End Property

Public Property Let ��������(ByVal vData As String)
    mstr�������� = vData
End Property

Public Property Get ��������() As String
    �������� = mstr��������
End Property

Public Property Let ����(ByVal vData As Long)
    mlng���� = vData
End Property

Public Property Get ����() As Long
    ���� = mlng����
End Property

Public Function MakePreSwapData() As ADODB.Recordset
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

Public Function MakePreSwapDataFromDB(ByVal bytMode As Byte, ByVal varValue As Variant, _
    ByVal strInvoice As String, Optional ByVal bln������ As Boolean) As ADODB.Recordset
    '���ݵ��ݺŻ����ID����һ����¼��Ϣ(���ۼ۵�λ)
    '���:
    '   bytMode ���ݻ�ȡ��ʽ��0-�����ݺŻ�ȡ��1-������ID��ȡ
    '   varValue bytMode=0ʱ�����ݺţ�bytMode=1ʱ������ID
    '   strInvoice ��ʼ��Ʊ��
    '   bln������ �Ƿ񲹽���
    '����:
    '����:ҽ��������ݵ����ݼ�
    Dim strTable  As String, strWhere As String
    Dim p As Integer, strSQL As String, strNos As String
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim lng����ID As Long, strNo As String
    
    On Error GoTo ErrHandler
    Set rsTmp = MakePreSwapData()
    
    If bytMode = 0 Then
        strWhere = strWhere & " And a.No=[2]"
        strNo = CStr(varValue)
    Else
        If bln������ Then
           strTable = ",(Select distinct �շѽ���ID From ���ò����¼ Where ����ID=[1]) B"
           strWhere = " And A.����ID=b.�շѽ���ID"
        Else
            strWhere = " And A.����ID=[1]"
        End If
        lng����ID = CLng(varValue)
    End If

    strSQL = _
        " Select A.NO,Nvl( A.�۸񸸺�, A.���) as ���,To_char(max(A.�Ǽ�ʱ��),'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
        "       A.����ID,A.�ѱ�,A.�շ����,A.�վݷ�Ŀ,A.���㵥λ,A.������," & _
        "       A.�շ�ϸĿID,A.���մ���ID As ����֧������ID,Nvl(A.������Ŀ��,0) As �Ƿ�ҽ��,A.���ձ���," & _
        "       Avg(Nvl(A.����,0)*A.����) As ����,Avg(A.��׼����) As ����," & _
        "       Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.ͳ����) As ͳ����,max(A.ժҪ) as ժҪ," & _
        "       nvl(A.�Ӱ��־,0) as �Ƿ���,A.��������ID,A.ִ�в���ID,A.����ID" & _
        " From ������ü�¼ A" & strTable & _
        " Where Mod(A.��¼����,10)=1 " & strWhere & _
        " Group By A.NO, Nvl(A.�۸񸸺�, A.���),A.����id, A.�ѱ�, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ," & _
        "       A.������, A.�շ�ϸĿid, A.���մ���id, Nvl(A.������Ŀ��, 0), A.���ձ���, A.ժҪ, Nvl(A.�Ӱ��־, 0)," & _
        "       A.��������id, A.ִ�в���id,A.����ID"
    
    strSQL = _
        " Select '" & strInvoice & "' as ʵ��Ʊ��,A.NO,A.���,max(A.����ʱ��) as ����ʱ��," & _
        "       A.����ID,A.�ѱ�,A.�շ����,A.�վݷ�Ŀ,A.���㵥λ,A.������," & _
        "       A.�շ�ϸĿID,A.����֧������ID,A.�Ƿ�ҽ��,A.���ձ���," & _
        "       sum(A.����) as ����,max(A.����) As ����, Sum(A.ʵ�ս��) As ʵ�ս��, " & _
        "       Sum(A.ͳ����) As ͳ����,max(A.ժҪ) as ժҪ," & _
        "       Max(A.�Ƿ���) as �Ƿ���,max(A.��������ID) as ��������ID,max(A.ִ�в���ID ) as ִ�в���ID " & _
        " From (" & strSQL & ") A" & _
        " Group By A.NO,A.���,A.����id, A.�ѱ�, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, A.������, A.�շ�ϸĿid, A.����֧������ID, " & _
        "       A.�Ƿ�ҽ��, A.���ձ���" & _
        " Order by NO,���"
    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����շ�����-ҽ��", lng����ID, strNo)
    
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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecutePreSwap( _
    rsRecord As ADODB.Recordset, _
    ByRef colBalance As Collection, _
    ByRef strInsureInvoice As String, _
    Optional ByVal blnReBalance As Boolean, _
    Optional ByVal bln������ As Boolean) As Boolean
    '����Ԥ����
    '��Σ�
    '   blnReBalance �Ƿ����½���
    '���Σ�
    '   colBalance Ԥ��������(ÿ�ŵ��ݶ�Ӧһ��BalanceMoneys����Ԫ��),�൥��һ�ν���ʱ���ڵ�һ�ŵ�����
    '   strInsureInvoice ҽ�����ط�Ʊ��
    Dim rsRecord As ADODB.Recordset
    Dim strBalance As String, strAdvance As String
    Dim varBalance As Variant, varItem As Variant, str���㷽ʽ As String
    Dim p As Long, i As Long
    
    On Error GoTo ErrHandler
    mblnҽ������Ʊ�� = False: strInsureInvoice = ""
    
    '2-�൥�ݷֵ��ݽ���
    If mbln�൥�ݷֵ��ݽ��� Then
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
                strBalance = ""
                strAdvance = IIf(blnReBalance, "1", "")
                
                If Not mclsInsure.ClinicPreSwap(rsRecord, strBalance, mlng����, strAdvance) Then
                    ShowMsgbox "�� " & p & " �ŵ���Ԥ����ʧ�ܡ�"
                    Exit Function
                End If
                
                'ֻҪ��һ�ŵ����Զ���Ʊ�ţ���Ҫ��Ʊ��
                If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then    'ҽ��Ʊ�ݺ� Then
                    '38821,��ʽ:Ʊ�ݺ�;�Ƿ���Ʊ��(1-����Ʊ��;0-�Զ���Ʊ��)
                    varItem = Split(strAdvance & ";", ";")
                    strInsureInvoice = varItem(0)
                    mblnҽ������Ʊ�� = mblnҽ������Ʊ�� And Val(varItem(1)) = 1
                End If
                
                '������ʽ;���;�Ƿ������޸�|....
                If strBalance <> "" Then
                    strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
                    Call SetBalanceVal(colBalance, p, strBalance)
                End If
            End If
        Next
        ExecutePreSwap = True: Exit Function
    End If
    
    '0-�൥��һ�ν���,1-�൥��һ�ν���ֵ����˷�
    strBalance = ""
    strAdvance = IIf(bln������ And blnReBalance, "3", IIf(bln������, "2", "1"))
    
    If Not mclsInsure.ClinicPreSwap(rsRecord, strBalance, mlng����, strAdvance) Then
        ShowMsgbox "����Ԥ����ʧ�ܡ�"
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then
        '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
        varItem = Split(strAdvance & ";", ";")
        strInsureInvoice = varItem(0)
        mblnҽ������Ʊ�� = Val(varItem(1)) = 1
    End If
    
    '������ʽ;���;�Ƿ������޸�|....
    If strBalance <> "" Then
        If mblnһ�ν���ֵ����˷� Then
            '1-�൥��һ�ν���ֵ����˷�
            '�������:���㷽ʽ;���;�Ƿ������޸�|...||�������:���㷽ʽ;���;�Ƿ������޸�|...||...
            varBalance = Split(strBalance, "||")
            For i = 0 To UBound(varBalance)
                If InStr(varBalance(i), ":") = 0 Then
                    ShowMsgbox "����Ԥ���㷵�ؽ�������ʽ����ȷ��"
                    Exit Function
                End If
                
                varItem = Split(varBalance(i), ":")
                p = Val(varItem(0)): str���㷽ʽ = varItem(1)
                If p < 1 Or p > colBalance.Count Then
                    ShowMsgbox "����Ԥ���㷵�ؽ�������ʽ����ȷ��"
                    Exit Function
                End If
                
                str���㷽ʽ = Replace(Replace(str���㷽ʽ, "|", "||"), ";", "|")
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            Next
        Else
            '0-�൥��һ�ν���
            strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
            Call SetBalanceVal(colBalance, 1, strBalance)
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

Public Function IsCanCancel(ByVal lng������� As Long) As Boolean
    '���ҽ�������Ƿ���������
    Dim strSQL As String, rsBalance As ADODB.Recordset
    Dim strBalance As String

    On Error GoTo ErrHandler
    If mbln����������� = False Then
        ShowMsgbox "ҽ����֧������������ϣ��������ϱ��ν��㣡"
        Exit Function
    End If

    strSQL = _
        "Select a.���㷽ʽ" & vbNewLine & _
        "From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
        "Where a.���㷽ʽ = b.���� And b.���� In(3,4) And a.�����ID Is Null And a.������� = [1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����������", lng�������)
    Do While Not rsBalance.EOF
        If Not gclsInsure.GetCapability(support�����������, mlng����ID, mlng����, Nvl(rsBalance!���㷽ʽ)) Then
            strBalance = strBalance & "," & Nvl(rsBalance!���㷽ʽ)
        End If
        rsBalance.MoveNext
    Loop

    If strBalance <> "" Then
        ShowMsgbox "ҽ�����㷽ʽ��" & Mid(strBalance, 2) & "����֧�����ϣ��������ϱ��ν��㣡"
        Exit Function
    End If
    IsCanCancel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteInsureInfoUpdate(ByVal lng����ID As Long, ByRef objBalanceBills As BalanceBills) As Boolean
    '�������ռ�¼�ı�����Ϣ
    '����:
    '����:�������ռ�¼�ı�����Ϣ���³ɹ�����True�����򷵻�False
    Dim strSQL As String, rsReCharge As ADODB.Recordset
    Dim strBXInfo As String, strPreNo As String
    Dim curʵ�ս�� As Currency, curͳ���� As Currency, bln������Ŀ As Boolean
    Dim blnTrans As Boolean, cllReChargePro As Collection
    Dim objBalanceBill As BalanceBill
    
    On Error GoTo ErrHandler
    Set objBalanceBills = New BalanceBills
    If mbln���ý������ = False Then
        strSQL = _
            "Select a.Id, a.No, a.���, a.����id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(a.����, 0) As ����," & vbNewLine & _
            "       Nvl(a.ʵ�ս��, 0) As ʵ�ս��, a.ժҪ, Nvl(a.������Ŀ��, 0) As ������Ŀ��, " & vbNewLine & _
            "       a.���մ���id, Nvl(a.ͳ����, 0) As ͳ����, a.���ձ���, a.��������" & vbNewLine & _
            "From ������ü�¼ A" & vbNewLine & _
            "Where a.��¼���� = 11 And a.����id = [1]"
    Else
        strSQL = _
            "Select a.No, a.���, a.����id, a.�շ�ϸĿid, Sum(Nvl(a.����, 1) * a.����) As ����," & vbNewLine & _
            "       Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս��, Max(Decode(a.��¼״̬,0,null,a.ժҪ)) As ժҪ," & vbNewLine & _
            "       Max(Decode(a.��¼״̬,0,null,Nvl(a.������Ŀ��, 0))) As ������Ŀ��, " & vbNewLine & _
            "       Max(Decode(a.��¼״̬,0,null,a.���մ���id)) As ���մ���id," & vbNewLine & _
            "       Nvl(Sum(a.ͳ����), 0) As ͳ����," & vbNewLine & _
            "       Max(Decode(a.��¼״̬,0,null,a.���ձ���)) As ���ձ���," & vbNewLine & _
            "       Max(Decode(a.��¼״̬,0,null,a.��������)) As ��������," & vbNewLine & _
            "       Max(Decode(a.��¼״̬,0,null,a.Id)) As ID" & vbNewLine & _
            "From ������ü�¼ A, ���ý������ B" & vbNewLine & _
            "Where a.ID = b.����ID And b.�Ƿ����� = 1 And b.����id = [1]" & vbNewLine & _
            "Group By a.No,a.���,a.����id,a.�շ�ϸĿid" & vbNewLine & _
            "Having Nvl(Sum(Nvl(a.����, 1) * a.����),0) <> 0"
    End If
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���շ��ü�¼", lng����ID)
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
                strBXInfo = mclsInsure.GetItemInsure(Nvl(!����ID), Nvl(!�շ�ϸĿID), Val(Nvl(!ʵ�ս��)), _
                    True, mlng����, Nvl(!ժҪ) & "||" & Val(Nvl(!����)))
                If strBXInfo <> "" Then
                    '  Zl_�����շѼ�¼_Update
                    strSQL = "Zl_�����շѼ�¼_Update("
                    '  Id_In         In ������ü�¼.Id%Type,
                    strSQL = strSQL & Nvl(!ID) & ","
                    '  ���մ���id_In In ������ü�¼.���մ���id%Type,
                    strSQL = strSQL & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  ������Ŀ��_In In ������ü�¼.������Ŀ��%Type,
                    strSQL = strSQL & Val(Split(strBXInfo, ";")(0)) & ","
                    '  ���ձ���_In   In ������ü�¼.���ձ���%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  ��������_In   In ������ü�¼.��������%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  ͳ����_In   In ������ü�¼.ͳ����%Type,
                    strSQL = strSQL & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  ժҪ_In       In ������ü�¼.ժҪ%Type
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllReChargePro, strSQL
                    
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
            
            '���ý������ʱ������ԭʼ���ü�¼
            If mbln���ý������ = False Then
                'ִ�й���
                blnTrans = True
                zlExecuteProcedureArrAy cllReChargePro, "ִ�б�����Ϣ����", True, True
                blnTrans = False
            End If
        End If
    End With
    ExecuteInsureInfoUpdate = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function