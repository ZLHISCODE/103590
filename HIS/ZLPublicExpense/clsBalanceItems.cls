VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Member0" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Tag As String

'�ֲ����������漯��
Private mCol As Collection
Public Enum gBalanceType
    gEM_��ͨ���� = 0
    gEM_Ԥ���� = 1
    gEM_ҽ�� = 2
    gEM_һ��ͨ = 3
    gEM_һ��ͨ_�� = 4
    gEM_���ѿ� = 5
    gEM_���ʵ� = 6
    gEM_���۵� = 7
End Enum
Private mvarobjTag As Object '�ֲ�����
Private mvar������ As Double '�ֲ�����
Private mvarԭʼ��� As Double '�ֲ�����
Private mvarδ�˽�� As Double '�ֲ�����
Private mvar���� As Double
Private mvar�Ƿ�ת�� As Boolean '�ֲ�����
Private mvar�˷ѽ���IDs As String '�ֲ�����
Private mvar�Ƿ��쳣�˷� As Boolean '�ֲ�����
Private mvar���� As gBalanceType '�ֲ�����
Private mvar�շ����� As Byte '�ֲ�����
Private mvar���ݺ� As String '�ֲ�����
Private mvar�Ƿ񱣴� As Boolean '�ֲ�����
Private mvarͬ��״̬ As Integer '�ֲ�����
Private mvar�쳣ID As Long '�ֲ�����
Private mvarҵ��ID As Long '�ֲ�����
Private mvar������� As Boolean '�ֲ�����
Private mvar����ʱ�� As String '�ֲ�����
Public Property Set objTag(ByVal vData As Object)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objTag = Form1
    Set mvarobjTag = vData
End Property


Public Property Get objTag() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objTag
    Set objTag = mvarobjTag
End Property


Public Property Let ����ʱ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar����ʱ�� = vData
End Property
Public Property Get ����ʱ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʱ��
    ����ʱ�� = mvar����ʱ��
End Property

Public Property Let �������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ������� = 5
    mvar������� = vData
End Property
Public Property Get �������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�������
    ������� = mvar�������
End Property


Public Property Let ҵ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҵ��ID = 5
    mvarҵ��ID = vData
End Property
Public Property Get ҵ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҵ��ID
    ҵ��ID = mvarҵ��ID
End Property


Public Property Let ͬ��״̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͬ��״̬ = 5
    mvarͬ��״̬ = vData
End Property
Public Property Get ͬ��״̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͬ��״̬
    ͬ��״̬ = mvarͬ��״̬
End Property

Public Property Let �쳣ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�쳣ID = 5
    mvar�쳣ID = vData
End Property
Public Property Get �쳣ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�쳣ID
    �쳣ID = mvar�쳣ID
End Property

Public Property Let �Ƿ񱣴�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ񱣴� = 5
    mvar�Ƿ񱣴� = vData
End Property
Public Property Get �Ƿ񱣴�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ񱣴�
    �Ƿ񱣴� = mvar�Ƿ񱣴�
End Property



Public Property Let ���ݺ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ݺ� = 5
    mvar���ݺ� = vData
End Property
Public Property Get ���ݺ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ݺ�
    ���ݺ� = mvar���ݺ�
End Property


Public Property Let �շ�����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ����� = 5
    mvar�շ����� = vData
End Property
Public Property Get �շ�����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�շ�����
    �շ����� = mvar�շ�����
End Property

Public Property Let ����(ByVal vData As gBalanceType)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As gBalanceType
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property



Public Property Let �Ƿ��쳣�˷�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��쳣�˷� = 5
    mvar�Ƿ��쳣�˷� = vData
End Property


Public Property Get �Ƿ��쳣�˷�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��쳣�˷�
    �Ƿ��쳣�˷� = mvar�Ƿ��쳣�˷�
End Property



Public Property Let �˷ѽ���IDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˷ѽ���IDs = 5
    mvar�˷ѽ���IDs = vData
End Property


Public Property Get �˷ѽ���IDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷ѽ���IDs
    �˷ѽ���IDs = mvar�˷ѽ���IDs
End Property



Public Property Let �Ƿ�ת��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�ת�� = 5
    mvar�Ƿ�ת�� = vData
End Property


Public Property Get �Ƿ�ת��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�ת��
    �Ƿ�ת�� = mvar�Ƿ�ת��
End Property

Public Property Let ԭʼ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭʼ��� = 5
    mvarԭʼ��� = vData
End Property
Public Property Get ԭʼ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭʼ���
    ԭʼ��� = mvarԭʼ���
End Property
Public Property Let δ�˽��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ�˽�� = 5
    mvarδ�˽�� = vData
End Property
Public Property Get δ�˽��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ�˽��
    δ�˽�� = mvarδ�˽��
End Property

Public Property Let ������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property
Public Property Get ������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property



Public Property Let ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property
Public Property Get ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Sub AddItem(ByVal objItem As clsBalanceItem, Optional sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objItem
    Else
        mCol.Add objItem, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsBalanceItem
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub


Public Function Clone(Optional blnSubItem As Boolean = True) As clsBalanceItems
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��¡һ���µĽ�����Ϣ��
    '���:blnSubItem-�Ƿ���Ҫ��������
    '����:���˺�
    '����:2019-11-13 10:51:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItems As New clsBalanceItems
    Dim objItem As clsBalanceItem, i As Long
    With objItems
        .���� = Me.����
        .������ = Me.������
        .δ�˽�� = Me.δ�˽��
        
        .�Ƿ��쳣�˷� = Me.�Ƿ��쳣�˷�
        .�Ƿ�ת�� = Me.�Ƿ�ת��
        .�շ����� = Me.�շ�����
        .�˷ѽ���IDs = Me.�˷ѽ���IDs
        .���� = Me.����
        .���ݺ� = Me.���ݺ�
        .ҵ��ID = Me.ҵ��ID
        .�쳣ID = Me.�쳣ID
        .ͬ��״̬ = Me.ͬ��״̬
        .������� = Me.�������
        .�Ƿ񱣴� = Me.�Ƿ񱣴�
        .����ʱ�� = Me.����ʱ��
        .ԭʼ��� = Me.ԭʼ���
        
        .Tag = Me.Tag
        Set .objTag = Me.objTag
    End With
    If blnSubItem = False Then Set Clone = objItems: Exit Function
    For i = 1 To Me.Count
        Set objItem = Me(i).Clone
        objItems.AddItem objItem
    Next
    Set Clone = objItems
End Function
Public Sub CloneItemsPropertyByItems(ByVal objSourceItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��¡Items����ֵ,������������仯
    '���:
    '����:���˺�
    '����:2019-11-29 20:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If objSourceItems Is Nothing Then Exit Sub
    With Me
        .���� = objSourceItems.����
        .������ = 0
        .�Ƿ��쳣�˷� = objSourceItems.�Ƿ��쳣�˷�
        .�Ƿ�ת�� = objSourceItems.�Ƿ�ת��
        .�շ����� = objSourceItems.�շ�����
        .�˷ѽ���IDs = objSourceItems.�˷ѽ���IDs
        .���� = objSourceItems.����
        .���ݺ� = objSourceItems.���ݺ�
        .ҵ��ID = objSourceItems.ҵ��ID
        .�쳣ID = objSourceItems.�쳣ID
        .ͬ��״̬ = objSourceItems.ͬ��״̬
        .������� = objSourceItems.�������
        .�Ƿ񱣴� = objSourceItems.�Ƿ񱣴�
        .����ʱ�� = objSourceItems.����ʱ��
        .δ�˽�� = objSourceItems.δ�˽��
        
        .Tag = objSourceItems.Tag
        Set objTag = objSourceItems.objTag
        For i = 1 To Me.Count
            .������ = .������ + Me(i).������
            .δ�˽�� = .������ + Me(i).δ�˽��
            .ԭʼ��� = .ԭʼ��� + Me(i).ԭʼ���
        Next
    End With
End Sub