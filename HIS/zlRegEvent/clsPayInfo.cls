VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsPayRegInfo"
Attribute VB_Ext_KEY = "Member0" ,"clsPayRegInfo"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection
'���Һ�֧����Ϣ
Private mvar���� As String '�ֲ�����
Private mvar�ӿ���� As Long '�ֲ�����
Private mvar���ѿ� As Boolean '�ֲ�����
Private mvarУ�Ա�־ As Integer '�ֲ�����
Private mvar֧������ As gPagePay '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar֧������ As String '�ֲ�����
Private mvar�������� As Integer '�ֲ�����
Private mvar���ѿ�ID As Long '�ֲ�����
Private mvar����ɹ� As Boolean '�ֲ�����
Private mvar��������ID As Long '�ֲ�����
Private mvar�̶� As Boolean '�ֲ�����
Private mvar�������� As Boolean '�ֲ�����
Private mvarQRCode As String '�ֲ�����
Private mvarCardPro As Card '�ֲ�����

'������Ϣ���Ӽ��ظ���ֻΪ�����������㷽ʽȡֵ
Private mvar���㷽ʽ As String '�ֲ�����
Private mvar֧����� As Double '�ֲ�����
Private mvar������� As String '�ֲ�����
Private mvar������ˮ�� As String '�ֲ�����
Private mvar����˵�� As String '�ֲ�����
Private mvarPayRow As Long '�ֲ�����

Public Property Let ��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property




Public Property Let �̶�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�̶� = 5
    mvar�̶� = vData
End Property


Public Property Get �̶�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�̶�
    �̶� = mvar�̶�
End Property




Public Property Let QRCode(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.QRCode = 5
    mvarQRCode = vData
End Property


Public Property Get QRCode() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.QRCode
    QRCode = mvarQRCode
End Property






Public Property Let �������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property


Public Property Get �������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mvar�������
End Property






Public Property Let PayRow(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PayRow = 5
    mvarPayRow = vData
End Property


Public Property Get PayRow() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PayRow
    PayRow = mvarPayRow
End Property


Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������ID = 5
    mvar��������ID = vData
End Property


Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������ID
    ��������ID = mvar��������ID
End Property


Public Property Let ����˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����˵�� = 5
    mvar����˵�� = vData
End Property


Public Property Get ����˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����˵��
    ����˵�� = mvar����˵��
End Property



Public Property Let ������ˮ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ˮ�� = 5
    mvar������ˮ�� = vData
End Property


Public Property Get ������ˮ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������ˮ��
    ������ˮ�� = mvar������ˮ��
End Property



Public Property Let ֧�����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.֧����� = 5
    mvar֧����� = vData
End Property


Public Property Get ֧�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.֧�����
    If mCol.Count > 1 Then
        Dim i As Integer, dbl��� As Double, objSubPay As clsSubPayInfo
        For i = 1 To mCol.Count
            Set objSubPay = mCol(i)
            dbl��� = dbl��� + objSubPay.������
        Next
        mvar֧����� = dbl���
    End If
    ֧����� = mvar֧�����
End Property



Public Property Let ���㷽ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���㷽ʽ = 5
    mvar���㷽ʽ = vData
End Property


Public Property Get ���㷽ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���㷽ʽ
    ���㷽ʽ = mvar���㷽ʽ
End Property




Public Property Let ����ɹ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ɹ� = 5
    mvar����ɹ� = vData
End Property


Public Property Get ����ɹ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ɹ�
    ����ɹ� = mvar����ɹ�
End Property



Public Property Let ���ѿ�ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ѿ�ID = 5
    mvar���ѿ�ID = vData
End Property


Public Property Get ���ѿ�ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ѿ�ID
    ���ѿ�ID = mvar���ѿ�ID
End Property



Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property



Public Property Let ֧������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.֧������ = 5
    mvar֧������ = vData
End Property


Public Property Get ֧������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.֧������
    ֧������ = mvar֧������
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property



Public Property Let ֧������(ByVal vData As gPagePay)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.֧������ = Form1
    mvar֧������ = vData
End Property


Public Property Get ֧������() As gPagePay
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.֧������
    ֧������ = mvar֧������
End Property



Public Property Let У�Ա�־(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.У�Ա�־ = 5
    mvarУ�Ա�־ = vData
End Property


Public Property Get У�Ա�־() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.У�Ա�־
    У�Ա�־ = mvarУ�Ա�־
End Property



Public Property Set CardPro(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.CardPro = Form1
    Set mvarCardPro = vData
End Property


Public Property Get CardPro() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CardPro
    Set CardPro = mvarCardPro
End Property



Public Property Let ���ѿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ѿ� = 5
    mvar���ѿ� = vData
End Property


Public Property Get ���ѿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ѿ�
    ���ѿ� = mvar���ѿ�
End Property



Public Property Let �ӿ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ӿ���� = 5
    mvar�ӿ���� = vData
End Property


Public Property Get �ӿ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ӿ����
    �ӿ���� = mvar�ӿ����
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property


Public Sub AddItem(ByVal objPayRegInfo As clsSubPayInfo, Optional ByVal sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objPayRegInfo
    Else
        mCol.Add objPayRegInfo, sKey
    End If
End Sub


Public Property Get Item(vntIndexKey As Variant) As clsSubPayInfo
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

Public Sub ReMoveAll()
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
