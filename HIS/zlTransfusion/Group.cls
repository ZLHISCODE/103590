VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Bills"
Option Explicit

'��������ֵ�ľֲ�����
'- ����ҽ������
Private mvarִ��ҽ��ID As Long '�ֲ�����
Private mvar���ͺ� As Long '�ֲ�����
Private mvar�������� As Integer
Private mvar��ִ������  As Integer

Private mvarִ�з��� As String '1-��Һ', 2, '2-ע��', 3, '3-Ƥ��', '0-����
Private mvarƤ�Խ�� As String  '(+) ���� (-)����
Private mvarִ��Ƶ�� As String
Private mvar�÷� As String
Private mvarƵ�ʴ��� As Integer
Private mvar���� As Integer
Private mvarִ��״̬ As Integer  '0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��
Private mvar��� As Integer
Private mvarҺ���� As Long
Private mvar��ʱ As Long
Private mvar˵�� As String
Private mvarҽ������ As String
Private mvarִ���� As String
Private mvar�Ʒ�״̬ As Integer  '-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷ�(���շ��࣬ʵ���Ƿ���ִ�У���Ҫ���ݶ�Ӧ���õ����Ƿ�����շ��ж�)
'Ҫ��ʱ��,��������,ִ��ժҪ,ִ��ʱ��
Private mvarҪ��ʱ�� As Date
Private mvar�������� As Integer
Private mvar�ϴ�ִ��ʱ�� As Date
Private mvar�շѽ�� As Currency '0 δ�շ� >0��ʾ���շ�
Private mvar��¼���� As Integer  '���һ��ͨ�޸� 1-�շѼ�¼,2-���ʼ�¼
Private mvar�ϴ���� As Integer     '�ϴ����
Private mvar�˶��� As String
'--
' ����ҽ����¼
Private mColBills As Collection '-- ҽ��ִ�еľ������ݼ���,�� �羲����Һ��ҩƷ��ϸ


Public Property Let �˶���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�˶��� = vData
End Property

Public Property Get �˶���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �˶��� = mvar�˶���
End Property


Public Property Let ��¼����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��¼���� = 5
    mvar��¼���� = vData
End Property


Public Property Get ��¼����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��¼����
    ��¼���� = mvar��¼����
End Property

Public Property Let Ҫ��ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarҪ��ʱ�� = vData
End Property

Public Property Get Ҫ��ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     Ҫ��ʱ�� = mvarҪ��ʱ��
End Property

Public Property Let �շѽ��(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�շѽ�� = vData
End Property

Public Property Get �շѽ��() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �շѽ�� = mvar�շѽ��
End Property

Public Property Let �ϴ�ִ��ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�ϴ�ִ��ʱ�� = vData
End Property

Public Property Get �ϴ�ִ��ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �ϴ�ִ��ʱ�� = mvar�ϴ�ִ��ʱ��
End Property

Public Function AddBills(ByVal objBIlls As Bills, Optional sKey As String) As Bills
    Dim objNewMember As Bills
    Set objNewMember = New Bills
    
    Set objNewMember = objBIlls
    If Len(sKey) = 0 Then
        mColBills.Add objNewMember
    Else
        mColBills.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set AddBills = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get BillsItem(vntIndexKey As Variant) As Bills
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set BillsItem = mColBills(vntIndexKey)
End Property

Public Property Get BillsCount() As Long
    ' x.Count
    BillsCount = mColBills.Count
End Property


Public Sub BillsRemove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mColBills.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mColBills.[_NewEnum]
End Property

Public Property Let ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar��� = vData
End Property

Public Property Get ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ��� = mvar���
End Property
Public Property Let �ϴ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�ϴ���� = vData
End Property

Public Property Get �ϴ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �ϴ���� = mvar�ϴ����
End Property
Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�������� = vData
End Property

Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �������� = mvar��������
End Property

Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�������� = vData
End Property

Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �������� = mvar��������
End Property

Public Property Let ��ִ������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar��ִ������ = vData
End Property

Public Property Get ��ִ������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ��ִ������ = mvar��ִ������
End Property

Public Property Let ִ�з���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarִ�з��� = vData
End Property

Public Property Get ִ�з���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ִ�з��� = mvarִ�з���
End Property

Public Property Let ִ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarִ���� = vData
End Property

Public Property Get ִ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ִ���� = mvarִ����
End Property

Public Property Let Ƥ�Խ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarƤ�Խ�� = vData
End Property

Public Property Get Ƥ�Խ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     Ƥ�Խ�� = mvarƤ�Խ��
End Property

Public Property Let ִ��Ƶ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarִ��Ƶ�� = vData
End Property

Public Property Get ִ��Ƶ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ִ��Ƶ�� = mvarִ��Ƶ��
End Property

Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar���� = vData
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     ���� = mvar����
End Property

Public Property Let �Ʒ�״̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvar�Ʒ�״̬ = vData
End Property

Public Property Get �Ʒ�״̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     �Ʒ�״̬ = mvar�Ʒ�״̬
End Property

Public Property Let Ƶ�ʴ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ƶ�ʴ���  = 1
    mvarƵ�ʴ��� = vData
End Property

Public Property Get Ƶ�ʴ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.Ƶ�ʴ���
     Ƶ�ʴ��� = mvarƵ�ʴ���
End Property

Public Property Let �÷�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��״̬  = "0-δִ��"
    mvar�÷� = vData
End Property

Public Property Get �÷�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��״̬
     �÷� = mvar�÷�
End Property

Public Property Let ִ��״̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��״̬  = "0-δִ��"
    mvarִ��״̬ = vData
End Property

Public Property Get ִ��״̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��״̬
     ִ��״̬ = mvarִ��״̬
End Property

Public Property Let ���ͺ�(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ͺ� = 5
    mvar���ͺ� = vData
End Property


Public Property Get ���ͺ�() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���ͺ�
    ���ͺ� = mvar���ͺ�
End Property

Public Property Let ִ��ҽ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ��ҽ��ID = 5
    mvarִ��ҽ��ID = vData
End Property

Public Property Get ִ��ҽ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    ִ��ҽ��ID = mvarִ��ҽ��ID
End Property

Public Property Get Һ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    Dim objBIll As Bill
    mvarҺ���� = 0
    For Each objBIll In Me.BillsItem(Me.ִ��ҽ��ID & "_" & Me.���ͺ�)
        mvarҺ���� = mvarҺ���� + objBIll.����
    Next
    Һ���� = mvarҺ����
End Property

Public Property Get ��ʱ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    Dim objBIll As Bill
    mvar��ʱ = 0
    For Each objBIll In Me.BillsItem(Me.ִ��ҽ��ID & "_" & Me.���ͺ�)
        mvar��ʱ = mvar��ʱ + objBIll.ʱ��
    Next
    ��ʱ = mvar��ʱ
End Property

Public Property Get ˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    Dim objBIll As Bill
    If mvar˵�� = "" Then
        mvarҽ������ = ""
        For Each objBIll In Me.BillsItem(Me.ִ��ҽ��ID & "_" & Me.���ͺ�)
            mvarҽ������ = mvarҽ������ & objBIll.ҽ������
        Next
        ˵�� = mvarҽ������
    Else
        ˵�� = mvar˵��
    End If
End Property

Public Property Let ˵��(ByVal vDate As String)
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ��ҽ��ID
    mvar˵�� = MidUni(vDate, 1, 200)
End Property

Private Sub Class_Terminate()
    Set mColBills = Nothing
End Sub

Private Sub Class_Initialize()
    '������󴴽�����
    Set mColBills = New Collection
End Sub

Public Function FuncExecRefuse() As Boolean
    '���ܣ��ӵ�ʱ�ܾ�ִ��
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String
    
    '����ִ�л���ִ�в������ܾ�
    
    If Me.ִ��״̬ = 2 Then
        MsgBox "��ִ����Ŀ��ǰ�Ѿ��ܾ�ִ�С�", vbInformation, gstrSysName
        Exit Function
    End If
    If Me.ִ��״̬ = 3 Then
        MsgBox "��ִ����Ŀ��ǰ����ִ�У����ܾܾ���", vbInformation, gstrSysName
        Exit Function
    End If
    If Me.ִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ��ǰ�Ѿ�ִ�У����ܾܾ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("ȷʵҪ�ܾ�ִ�и���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    lngҽ��ID = Me.ִ��ҽ��ID
    lng���ͺ� = Me.���ͺ�
    
    strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & lngҽ��ID & "," & lng���ͺ� & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRefuse")
    On Error GoTo 0
    Me.ִ��״̬ = 2
    
    FuncExecRefuse = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncExecRestore() As Boolean
'���ܣ�ȡ���ܾ�ִ��
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, i As Long
    
    If Me.ִ��״̬ <> 2 Then
        MsgBox "��ִ����Ŀû�б��ܾ�ִ�С�", vbInformation, gstrSysName
        Exit Function
    End If
        
    If MsgBox("ȷʵҪȡ���ܾ�ִ�и���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    lngҽ��ID = Me.ִ��ҽ��ID
    lng���ͺ� = Me.���ͺ�
    
    strSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & lngҽ��ID & "," & lng���ͺ� & ")"
    
    On Error GoTo errH
 
    Call zlDatabase.ExecuteProcedure(strSQL, "ExecutiveGroup.FuncExecRestore")
 
    On Error GoTo 0
    Me.ִ��״̬ = 0
    
    FuncExecRestore = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function