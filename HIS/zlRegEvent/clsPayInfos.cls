VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPayInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsPayInfo"
Option Explicit
'��ǰ�շ��˷ѽ���
'�ֲ����������漯��
Private mCol As Collection
'��������ֵ�ľֲ�����
Private mvarԤ���� As Double

Private mvarReg_���ݺ� As String '�ֲ�����
Private mvarReg_������ As Double '�ֲ�����
Private mvarReg_����ID As Long '�ֲ�����
Private mvarReg_Hav������ As Boolean '�ҺŷѵĽ��㷽ʽ�Ƿ�����������������ǰ�жϣ��������Ӽ��в���
Private mvarReg_Hav���ѿ� As Boolean
Private mvarReg_����Ʊ�� As Boolean '�ֲ�����

'���ѵ�֧����ʽΨһ�������������г���
Private mvarCard_���ݺ� As String '�ֲ�����
Private mvarCard_������ As Double '�ֲ�����
Private mvarCard_����ID As Long '�ֲ�����
Private mvarCard_���㷽ʽ As String '�ֲ�����
Private mvarCard_�����ID As Long '�ֲ�����
Private mvarCard_���ѿ� As Boolean '�ֲ�����
Private mvarCard_���ѿ�ID As Long '�ֲ�����
Private mvarCard_֧������ As String '�ֲ�����
Private mvarCard_֧������ As String '�ֲ�����
Private mvarCard_������ˮ�� As String '�ֲ�����
Private mvarCard_����˵�� As String '�ֲ�����
Private mvarCard_��������ID As Long '�ֲ�����
Private mvarCard_CardPro As Card '�ֲ�����
Private mvarCard_У�Ա�־ As Integer '�ֲ�����
Private mvarCard_����ɹ� As Boolean '�ֲ�����
Private mvarCard_������� As String '�ֲ�����
Private mvarCard_����Ʊ�� As Boolean '�ֲ�����
Private mvarCard_PayRow As Long '�ֲ�����

Public Property Let Card_����Ʊ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_����Ʊ�� = 5
    mvarCard_����Ʊ�� = vData
End Property


Public Property Get Card_����Ʊ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_����Ʊ��
    Card_����Ʊ�� = mvarCard_����Ʊ��
End Property



Public Property Let Reg_����Ʊ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_����Ʊ�� = 5
    mvarReg_����Ʊ�� = vData
End Property


Public Property Get Reg_����Ʊ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_����Ʊ��
    Reg_����Ʊ�� = mvarReg_����Ʊ��
End Property
Public Property Let Card_PayRow(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_PayRow = 5
    mvarCard_PayRow = vData
End Property


Public Property Get Card_PayRow() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_PayRow
    Card_PayRow = mvarCard_PayRow
End Property




Public Property Let Card_�������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_������� = 5
    mvarCard_������� = vData
End Property


Public Property Get Card_�������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_�������
    Card_������� = mvarCard_�������
End Property




Public Property Let Reg_Hav���ѿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_Hav���ѿ� = 5
    mvarReg_Hav���ѿ� = vData
End Property


Public Property Get Reg_Hav���ѿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_Hav���ѿ�
    Reg_Hav���ѿ� = mvarReg_Hav���ѿ�
End Property




Public Property Let Card_���ѿ�ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���ѿ�ID = 5
    mvarCard_���ѿ�ID = vData
End Property


Public Property Get Card_���ѿ�ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_���ѿ�ID
    Card_���ѿ�ID = mvarCard_���ѿ�ID
End Property


Public Property Let Card_����ɹ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_����ɹ� = 5
    mvarCard_����ɹ� = vData
End Property


Public Property Get Card_����ɹ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_����ɹ�
    Card_����ɹ� = mvarCard_����ɹ�
End Property



Public Property Let Card_У�Ա�־(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_У�Ա�־ = 5
    mvarCard_У�Ա�־ = vData
End Property


Public Property Get Card_У�Ա�־() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_У�Ա�־
    Card_У�Ա�־ = mvarCard_У�Ա�־
End Property





Public Property Let Reg_Hav������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_Hav������ = 5
    mvarReg_Hav������ = vData
End Property


Public Property Get Reg_Hav������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_Hav������
    Reg_Hav������ = mvarReg_Hav������
End Property



Public Property Set Card_CardPro(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Card_CardPro = Form1
    Set mvarCard_CardPro = vData
End Property


Public Property Get Card_CardPro() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_CardPro
    Set Card_CardPro = mvarCard_CardPro
End Property



Public Property Let Card_��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_��������ID = 5
    mvarCard_��������ID = vData
End Property


Public Property Get Card_��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_��������ID
    Card_��������ID = mvarCard_��������ID
End Property



Public Property Let Card_����˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_����˵�� = 5
    mvarCard_����˵�� = vData
End Property


Public Property Get Card_����˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_����˵��
    Card_����˵�� = mvarCard_����˵��
End Property



Public Property Let Card_������ˮ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_������ˮ�� = 5
    mvarCard_������ˮ�� = vData
End Property


Public Property Get Card_������ˮ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_������ˮ��
    Card_������ˮ�� = mvarCard_������ˮ��
End Property



Public Property Let Card_֧������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_֧������ = 5
    mvarCard_֧������ = vData
End Property


Public Property Get Card_֧������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_֧������
    Card_֧������ = mvarCard_֧������
End Property



Public Property Let Card_֧������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_֧������ = 5
    mvarCard_֧������ = vData
End Property


Public Property Get Card_֧������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_֧������
    Card_֧������ = mvarCard_֧������
End Property



Public Property Let Card_���ѿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���ѿ� = 5
    mvarCard_���ѿ� = vData
End Property


Public Property Get Card_���ѿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_���ѿ�
    Card_���ѿ� = mvarCard_���ѿ�
End Property



Public Property Let Card_�����ID(ByVal vData As Long)
Attribute Card_�����ID.VB_Description = "����֧�������ID"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_�����ID = 5
    mvarCard_�����ID = vData
End Property


Public Property Get Card_�����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_�����ID
    Card_�����ID = mvarCard_�����ID
End Property




Public Property Let Ԥ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ���� = 5
    mvarԤ���� = vData
End Property


Public Property Get Ԥ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ����
    Ԥ���� = mvarԤ����
End Property


Public Property Let Card_���㷽ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���㷽ʽ = 5
    mvarCard_���㷽ʽ = vData
End Property


Public Property Get Card_���㷽ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_���㷽ʽ
    Card_���㷽ʽ = mvarCard_���㷽ʽ
End Property


Public Property Let Card_����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_����ID = 5
    mvarCard_����ID = vData
End Property


Public Property Get Card_����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_����ID
    Card_����ID = mvarCard_����ID
End Property



Public Property Let Card_������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_������ = 5
    mvarCard_������ = vData
End Property


Public Property Get Card_������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_������
    Card_������ = mvarCard_������
End Property



Public Property Let Card_���ݺ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���ݺ� = 5
    mvarCard_���ݺ� = vData
End Property


Public Property Get Card_���ݺ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_���ݺ�
    Card_���ݺ� = mvarCard_���ݺ�
End Property


Public Property Let Reg_����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_����ID = 5
    mvarReg_����ID = vData
End Property


Public Property Get Reg_����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_����ID
    Reg_����ID = mvarReg_����ID
End Property



Public Property Let Reg_������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_������ = 5
    mvarReg_������ = vData
End Property


Public Property Get Reg_������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_������
    Reg_������ = mvarReg_������
End Property



Public Property Let Reg_���ݺ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reg_���ݺ� = 5
    mvarReg_���ݺ� = vData
End Property


Public Property Get Reg_���ݺ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reg_���ݺ�
    Reg_���ݺ� = mvarReg_���ݺ�
End Property


Public Sub AddItem(ByVal objPayInfo As clsPayInfo, Optional ByVal sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objPayInfo
    Else
        mCol.Add objPayInfo, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsPayInfo
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Dim i As Integer
    If UCase(Left(vntIndexKey, 3)) = "PAY" Then
        For i = 1 To mCol.Count
            If mCol(i).֧������ = Val(Mid(vntIndexKey, 4)) Then
                Set Item = mCol(i): Exit For
            End If
        Next
    Else
        Set Item = mCol(vntIndexKey)
    End If
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

Public Sub ReMoveAll()
    'ɾ�������е�����Ԫ��ʱʹ�á�

    Set mCol = New Collection
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

