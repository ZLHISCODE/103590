VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicInsureDelSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mlng����ID As Long '�ֲ�����
Private mstr���� As String '�ֲ�����
Private mstr�Ա� As String '�ֲ�����
Private mstr���� As String '�ֲ�����
Private mstr����� As String '�ֲ�����
Private mstrסԺ�� As String '�ֲ�����
Private mstrҽ�Ƹ��ʽ As String '�ֲ�����
Private mstr�����ʻ� As String '�ֲ�����
Private mint���� As Integer '�ֲ�����
Private mbln����������� As Boolean '�ֲ�����
Private mbln�˷Ѻ��ӡ�ص� As Boolean '�ֲ�����
Private mbln������� As Boolean '�ֲ�����
Private mlng����ID As Long '�ֲ�����
Private mblnҽ���ӿڴ�ӡƱ�� As Boolean '�ֲ�����
Private mlng����ID As Long '�ֲ�����
'�ֶ�:����,NO,���㷽ʽ,��������,��Ԥ��,�����ID,���㿨���,����,���ѿ�ID,������ˮ��,����˵��,ԭ����ID,��������ID
Private mrsBalance As ADODB.Recordset '�ֲ�����
Private mbln�ֵ����˷� As Boolean '�ֲ�����

Public Property Let ҽ�Ƹ��ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ�Ƹ��ʽ = 5
    mstrҽ�Ƹ��ʽ = vData
End Property


Public Property Get ҽ�Ƹ��ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ�Ƹ��ʽ
    ҽ�Ƹ��ʽ = mstrҽ�Ƹ��ʽ
End Property



Public Property Let סԺ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ�� = 5
    mstrסԺ�� = vData
End Property


Public Property Get סԺ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.סԺ��
    סԺ�� = mstrסԺ��
End Property



Public Property Let �����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mstr����� = vData
End Property


Public Property Get �����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����
    ����� = mstr�����
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mstr���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mstr����
End Property



Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mstr�Ա� = vData
End Property


Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ա�
    �Ա� = mstr�Ա�
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mstr���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mstr����
End Property




Public Function InitData(ByVal str�����ʻ� As String) As Boolean
    '��ʼ��������������
    mstr�����ʻ� = str�����ʻ�
    InitData = True
End Function

Public Property Get �ֵ����˷�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֵ����˷�
    �ֵ����˷� = mbln�ֵ����˷�
End Property

Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.rsBalance = Form1
    Set mrsBalance = vData
    
    If mrsBalance Is Nothing Then Exit Property
    If mrsBalance.State <> 1 Then Exit Property
    
    '�ֶ�:����,NO,���㷽ʽ,��������,��Ԥ��,�����ID,���㿨���,����,���ѿ�ID,������ˮ��,����˵��,ԭ����ID,��������ID
    mrsBalance.Filter = "����=" & gBalanceType.gEM_ҽ�� & " And NO<>null"
    mbln�ֵ����˷� = Not mrsBalance.EOF
End Property

Public Property Get rsBalance() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rsBalance
    Set rsBalance = mrsBalance
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mlng����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mlng����ID
End Property

Public Property Get ҽ���ӿڴ�ӡƱ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ���ӿڴ�ӡƱ��
    ҽ���ӿڴ�ӡƱ�� = mblnҽ���ӿڴ�ӡƱ��
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mlng����ID = vData
    
    If gclsInsure Is Nothing Then Exit Property
    mblnҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, mlng����ID, mint����, CStr(mlng����ID))
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mlng����ID
End Property

Public Property Let �������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mbln������� = vData
End Property

Public Property Get �������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mbln�������
End Property

Public Property Get �˷Ѻ��ӡ�ص�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷Ѻ��ӡ�ص�
    �˷Ѻ��ӡ�ص� = mbln�˷Ѻ��ӡ�ص�
End Property

Public Property Get �����������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����������
    ����������� = mbln�����������
End Property

Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mint���� = vData
    
    If gclsInsure Is Nothing Then Exit Property
    mbln�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, mlng����ID, mint����)
    mbln����������� = gclsInsure.GetCapability(support�����������, mlng����ID, mint����)
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mint����
End Property


Public Property Get �����ʻ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����ʻ�
    �����ʻ� = mstr�����ʻ�
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mlng����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mlng����ID
End Property

Public Function ExecuteClinicDelSwap(cllBillPro As Collection, ByRef blnCommit As Boolean) As Boolean
    'ִ���˷�
    '˵����
    '   �ɵ����߿�������ִ�гɹ����ύ���ݣ�ִ��ʧ�ܣ��ع�����
    
    blnCommit = False
    If mint���� = 0 Then
        MsgBox "������Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    If rsBalance Is Nothing Then
        MsgBox "ԭʼ����������Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    If rsBalance.State <> 1 Then
        MsgBox "ԭʼ����������Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln������� Then
        If ExecuteReplenishDel(cllBillPro) Then
            blnCommit = True
            ExecuteClinicDelSwap = True
        End If
    Else
        ExecuteClinicDelSwap = ExecuteClinicDel(cllBillPro, blnCommit)
    End If
End Function

Private Function GetAllNos() As String
    '��ȡ�൥�ݷֵ��ݽ���ʱ�����е��ݺţ����ݺŽ�������
    Dim strNos As String
    
    '�ֶ�:����,NO,���㷽ʽ,��������,��Ԥ��,�����ID,���㿨���,����,���ѿ�ID,������ˮ��,����˵��,ԭ����ID,��������ID
    mrsBalance.Filter = "����=" & gBalanceType.gEM_ҽ�� & " And NO<>null"
    mrsBalance.Sort = "No Desc"
    Do While Not mrsBalance.EOF
        If InStr("," & strNos & ",", "," & NVL(mrsBalance!NO) & ",") = 0 Then
            strNos = strNos & "," & NVL(mrsBalance!NO)
        End If
        mrsBalance.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetAllNos = strNos
End Function

Private Function ExecuteClinicDel(cllBillPro As Collection, ByRef blnCommit As Boolean) As Boolean
    '�����շ��˷�
    '��Σ�
    '   bytMode ҽ���˷�ģʽ��0-�൥��һ�ν���,1-�൥�ݷֵ����˷�
    '   strSavedAdvance ��Ҫ������ԭʼҽ��������Ϣ
    '   strNos��lng����ID��intInsure��bln����������ϡ�str�����ʻ� bytMode=1ʱ����
    '   strNos ���������еĵ��ݺţ���ʽ��A001,A002,...
    '˵��:��Ҫ�������������,�����˷Ѻ�,�ù������ύ,����Ҫ�������ύ
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAllAdvance As String, strAdvance As String
    Dim colBalance As Collection, varNos As Variant
    Dim p As Integer, strNO As String, str���㷽ʽ As String
    Dim blnCurrentCommit As Boolean
    Dim rsDelBalance As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "ҽ�������˷�", True, False
    
    strAllAdvance = GetYBBalance(mrsBalance, "", True, True)
    Call SaveInsureBalance(mlng����ID, mlng����ID, strAllAdvance)
    
    If Not mbln�ֵ����˷� Then
        strAdvance = mlng����ID & "|" & "0"
        If Not gclsInsure.ClinicDelSwap(mlng����ID, True, mint����, strAdvance) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        blnTransMedicare = True
    Else
        Set colBalance = New Collection
        Set rsDelBalance = GetBalanceDetail(0, mlng����ID, 1)
        
        '93337,�˷�ʱ�����ݺŵ�����нӿڵ���
        varNos = Split(GetAllNos(), ",")
        For p = UBound(varNos) To LBound(varNos) Step -1
            colBalance.Add New BalanceMoneys
            strNO = varNos(p)
            
            '�����ŵ����Ƿ��ѳɹ�ҽ������
            str���㷽ʽ = GetYBBalance(rsDelBalance, strNO)
            
            If str���㷽ʽ <> "" Then
                Call SetBalanceVal(colBalance, colBalance.Count, str���㷽ʽ)
            Else
                str���㷽ʽ = GetYBBalance(mrsBalance, strNO, True, True)
                'str���㷽ʽ Ϊ�գ���ʾҽ����֧��ҽ������
                If str���㷽ʽ <> "" Then
                    Call SaveInsureBalanceDetail(mlng����ID, strNO, str���㷽ʽ)
                    
                    strAdvance = mlng����ID & "|" & "0" & "|" & strNO
                    '��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
                    If Not gclsInsure.ClinicDelSwap(mlng����ID, True, mint����, strAdvance) Then
                        gcnOracle.RollbackTrans
                        If blnCurrentCommit Then Call CorrectInsureErrBalance(mlng����ID, mlng����ID)  'ҽ������У��
                        Exit Function
                    End If
                    blnTransMedicare = True
                    
                    If InsureCheck(str���㷽ʽ, strAdvance) Then
                        Call SaveInsureBalanceDetail(mlng����ID, strNO, strAdvance)
                        str���㷽ʽ = strAdvance
                    End If
                    Call SetBalanceVal(colBalance, colBalance.Count, str���㷽ʽ)
                    gcnOracle.CommitTrans: blnTrans = False
                    blnCommit = True
                    blnCurrentCommit = True
                    
                    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mint����)
                    blnTransMedicare = False
                    
                    gcnOracle.BeginTrans: blnTrans = True
                End If
            End If
        Next
        
        'ȫ���ɹ��������ܵĽ��㷽ʽ
        strAdvance = GetBalanceStr(colBalance)
    End If
    
    'У�Խ�����
    If InsureCheck(strAllAdvance, strAdvance) Then
        Call SaveInsureBalance(mlng����ID, mlng����ID, strAdvance)
    End If
    
    Call BalanceOver(mlng����ID)
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mint����)
    ExecuteClinicDel = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
    
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mint����)
    If blnCurrentCommit Then Call CorrectInsureErrBalance(mlng����ID, mlng����ID) 'ҽ������У��
End Function

Private Function ExecuteReplenishDel(cllBillPro As Collection) As Boolean
    '��������˷�
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAllAdvance As String, strAdvance As String
    
    On Error GoTo ErrHandler
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "ҽ����������˷�", True, False
    
    strAllAdvance = GetYBBalance(mrsBalance, "", True, True)
    Call SaveInsureBalance(mlng����ID, mlng����ID, strAllAdvance, True)
    
    'ClinicDelSwap(lngStlID As Long, Optional ByVal bln�˷� As Boolean = True, _
        Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
    '����:�������˷ѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ��
    '���:lngStlID-��Ҫ�˵ķѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '     bln�˷� -�������˷ѽ��׻��Ǹķѽ����ڵ��ñ��ӿ�
    '     strAdvance:��ʽ:����ID|��������־|��,ÿλ|�ָ�
    '           ��һλ:�������ID,ҽ�����Ը��ݳ���ID������ȡ��
    '           �ڶ�λ:��������־,1-����������;0�ǲ���������
    '           ����λ:NO:��ǰ�����NO
    '           ����λ��: ���Ժ���չ
    '     ע�⣺
    '           strAdvance��10.34.0��ǰ(�������ʽ���)
    '               �൥��һ�ν���ʱ,�������ԭ����IDs:����ID1,����ID2,...
    '               �����������ʽΪ:�˷ѵ���������|��ǰ�˵ڼ��ŵ���
    '����:strAdvance:1.ԭ���˻�ʱ�����ؿ�
    '                2.�˷ѽ��㷽ʽ���շѽ��㷽ʽ��һ��ʱ�����ظ�ʽΪ�����㷽ʽ|���||���㷽ʽ|���||�������У����Ϊ����
    strAdvance = mlng����ID & "|1"
    If Not gclsInsure.ClinicDelSwap(mlng����ID, True, mint����, strAdvance) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    blnTransMedicare = True
    
    If Not InsureCheck(strAllAdvance, strAdvance) Then
        gcnOracle.CommitTrans
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mint����)
        ExecuteReplenishDel = True: Exit Function
    End If
    gcnOracle.CommitTrans: gcnOracle.BeginTrans
    
    '�˷Ѻ��շѲ�һ��ʱ,��ҪЧ��
    Call SaveInsureBalance(mlng����ID, mlng����ID, strAdvance, True)
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mint����)
    ExecuteReplenishDel = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
    
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mint����)
End Function

Private Function GetYBBalance(rsBalance As ADODB.Recordset, Optional ByVal strNos As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnDelCheck As Boolean) As String
    '�����ݻ�ȡҽ��ԭ���㷽ʽ�ͽ�����
    '������
    '   strNOs - ���ݺ�,����ö��Ÿ�����A0001,A0002,...
    '   blnDel - ����Ƿ�ȡ��
    '   blnDelCheck - �Ƿ������������������
    '����:������Ϣ,��ʽ:���㷽ʽ|������||...
    Dim str���㷽ʽ As String, varNos As Variant, strFilter As String
    Dim i As Integer, p As Integer
    Dim colBalance As Collection, strPreNo As String
    
    On Error GoTo errHandle
    If blnDelCheck And mint���� = 0 Then Exit Function
    If rsBalance Is Nothing Then Exit Function
    
    '�ֶ�:����,NO,���㷽ʽ,��������,��Ԥ��,�����ID,���㿨���,����,���ѿ�ID,������ˮ��,����˵��,ԭ����ID,��������ID
    If strNos = "" Then
        strFilter = "����=" & gBalanceType.gEM_ҽ��
    Else
        varNos = Split(strNos, ",")
        For i = 0 To UBound(varNos)
            strFilter = strFilter & " Or (����=" & gBalanceType.gEM_ҽ�� & " And No='" & varNos(i) & "')"
        Next
        If strFilter <> "" Then strFilter = Mid(strFilter, 5)
    End If
    rsBalance.Filter = strFilter
    If rsBalance.RecordCount = 0 Then Exit Function
    
    rsBalance.Sort = "No"
    Set colBalance = New Collection
    p = 1: colBalance.Add New BalanceMoneys
    With rsBalance
        strPreNo = NVL(!NO)
        Do While Not .EOF
            If strPreNo <> NVL(!NO) Then
                p = p + 1: colBalance.Add New BalanceMoneys
                strPreNo = NVL(!NO)
            End If
            If blnDelCheck Then
                '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                If mbln����������� Then
                    If gclsInsure.GetCapability(support�����������, mlng����ID, mint����, !���㷽ʽ) Then
                        str���㷽ʽ = NVL(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(NVL(!��Ԥ��))
                    End If
                Else     '��֧�������������ʱ,ֻ����������Ϊ�ֽ�,����ԭ����,������ҽ������
                    If !���㷽ʽ <> mstr�����ʻ� Then
                        str���㷽ʽ = NVL(!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(NVL(!��Ԥ��))
                    End If
                End If
            Else
                str���㷽ʽ = NVL(!���㷽ʽ) & "|" & Val(NVL(!��Ԥ��))
            End If
            
            Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            .MoveNext
        Loop
    End With
    GetYBBalance = GetBalanceStr(colBalance)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureCheck(ByVal strԤ���� As String, ByVal strAdvance As String) As Boolean
    '��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:
    '   strԤ����-���ս���
    '   strAdvance-ҽ�����صĽ���
    '˵����
    '   ��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    Dim blnFind  As Boolean, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo ErrHandler
    If strAdvance = "" Or strԤ���� = strAdvance Then Exit Function
    
    InsureCheck = True
    
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
    InsureCheck = False
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function GetBalanceStr(colBalance As Collection, Optional ByVal intPage As Integer, _
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
    rsTemp.Fields.Append "���㷽ʽ", adVarChar, 50, adFldIsNullable
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
            rsTemp!��� = Val(NVL(rsTemp!���)) + objItem.��Ч���
            rsTemp.Update
        Next
    Next
    
    strBalance = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strBalance = strBalance & "||" & NVL(rsTemp!���㷽ʽ) & "|" & NVL(rsTemp!���)
        rsTemp.MoveNext
    Loop
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetBalanceStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
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

Private Sub SaveInsureBalanceDetail(ByVal lng����ID As Long, ByVal strNO As String, _
    ByVal strBalance As String)
    '����ҽ��������ϸ
    Dim strSQL As String
    'Zl_ҽ��������ϸ_Insert(
    strSQL = "Zl_ҽ��������ϸ_Insert( "
    '  ����id_In   ҽ��������ϸ.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  No_In       ҽ��������ϸ.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  ���㷽ʽ_In Varchar2,
    strSQL = strSQL & "'" & strBalance & "',"
    '  ��ע_In     ҽ��������ϸ.��ע%Type := Null,
    strSQL = strSQL & "" & "NULL" & ")"
    zlDatabase.ExecuteProcedure strSQL, "����ҽ��������ϸ"
End Sub

Private Sub SaveInsureBalance(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal strBalance As String, Optional ByVal bln������� As Boolean)
    '����ҽ����������
    Dim strSQL As String
    If bln������� Then
        strSQL = "Zl_���ò������_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & mlng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strBalance & "')"
    Else
        'Zl_�����˷ѽ���_Modify(
        strSQL = "Zl_�����˷ѽ���_Modify_S("
        '  ��������_In      Number,
        strSQL = strSQL & "" & 3 & ","
        '  ����id_In        ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����_In          ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & mstr���� & "',"
        '  �Ա�_In          ����Ԥ����¼.�Ա�%Type,
        strSQL = strSQL & "'" & mstr�Ա� & "',"
        '  ����_In          ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & mstr���� & "',"
        '  �����_In        ����Ԥ����¼.�����%Type,
        strSQL = strSQL & "'" & mstr����� & "',"
        '  סԺ��_In        ����Ԥ����¼.סԺ��%Type,
        strSQL = strSQL & "'" & mstrסԺ�� & "',"
        '  ���ʽ����_In  ����Ԥ����¼.���ʽ����%Type,
        strSQL = strSQL & "'" & mstrҽ�Ƹ��ʽ & "',"
        '  ����id_In        ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In      Varchar2
        strSQL = strSQL & "'" & strBalance & "')"
    End If
    zlDatabase.ExecuteProcedure strSQL, "����ҽ����������"
End Sub

Private Sub BalanceOver(ByVal lng����ID As Long)
    'ҽ����ɽ��㣬����У�Ա�־
    Dim strSQL As String
    'Zl_���������շ�_ҽ������(
    strSQL = "Zl_���������շ�_ҽ������( "
    '  ����id_In   ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���ս���_In Varchar2
    strSQL = strSQL & "" & "NULL" & ")"
    zlDatabase.ExecuteProcedure strSQL, "���ҽ������"
End Sub

Private Function CorrectInsureErrBalance(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
    '�൥�ݷֵ��ݽ��㣬��;ʧ��ʱУ�Խ�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset, strBalance As String
    Dim rsBalanceSaved As ADODB.Recordset, strBalanceSaved As String
    
    On Error GoTo ErrHandler
    strSQL = "Select 1" & _
            " From ����Ԥ����¼ A, ���㷽ʽ B" & _
            " Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And ����id = [1] And a.�����ID Is Null " & _
            "       And Nvl(a.У�Ա�־, 0) = 1 And Rownum < 2"
    strSQL = strSQL & "Union All" & _
            " Select 1" & _
            " From ���ս����¼" & _
            " Where ��¼id = [1] " & _
            "       And Not Exists(Select 1 " & _
            "                      From ����Ԥ����¼ A, ���㷽ʽ B" & _
            "                      Where a.���㷽ʽ = b.���� And a.����id = ��¼id " & _
            "                            And b.���� In (3, 4) And a.�����ID Is Null)" & _
            "       And �����ID Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ������ҪУ�Ե�ҽ������", lng����ID)
    If rsTemp.EOF Then CorrectInsureErrBalance = True: Exit Function
    
    '��ͨ����ҽ��������ϸ������У��
    Set rsBalance = GetBalanceDetail(0, lng����ID, 1)
    strBalance = GetYBBalance(rsBalance)
    
    If strBalance = "" Then
        strSQL = "Select a.����ID,a.���㷽ʽ,a.���" & _
            " From ���ս�����ϸ A ,���㷽ʽ C" & _
            " Where a.���㷽ʽ=c.���� And c.���� in (3,4) And a.����id =[1] And a.��־=1 " & _
            " Order by ���㷽ʽ"
        'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�",�����ſ���ҽ����Ľ��㷽ʽ
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", lng����ID)
        Do While Not rsTemp.EOF
            strBalance = strBalance & "||" & NVL(rsTemp!���㷽ʽ) & "|" & Val(NVL(rsTemp!���))
            rsTemp.MoveNext
        Loop
        If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    End If
    'û�к˶�����,ֱ�ӷ���
    If strBalance = "" Then CorrectInsureErrBalance = True: Exit Function
    
    '����Ƿ���ҪУ��
    Set rsBalanceSaved = GetChargeBalance(0, lng����ID)
    strBalanceSaved = GetYBBalance(rsBalanceSaved, lng����ID)
    If InsureCheck(strBalanceSaved, strBalance) Then
        Call SaveInsureBalance(lng����ID, lng����ID, strBalance)
    End If
    
    CorrectInsureErrBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function