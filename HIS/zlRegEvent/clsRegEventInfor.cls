VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEventInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsPayInfos"
Option Explicit
'*****************************************************************************************************
'����������д���򼰷ּ�
'clsRegEventInfor - |�Һ���Ϣ���Ǽ�ʱ�䡢����id�Ȼ���ֵ
'                   |������Ϣ�����˻�����Ϣ����Ҫ�ǽӿڵ���ʱ�õ�
'                   |����������Ϣ�����š�������Լ���������(�������˿����󿨡�NUll)
'   �NzlOneCardComLib.clsPatientInfo -|���˻�����Ϣ
'   �NclsPayInfos - |�Һŵ�����Ϣ���Һŵ��ݺš����Լ��Ƿ���������֧��(��ǰ�ж��Ƿ���Ҫ���ýӿ�)
'                   |����������Ϣ���������ݺ��Լ�����
'                   |����������Ϣ��֧����ʽ�Լ��ӿڽ�����Ϣ
'       �NclsPayInfo - |�ҺŽ��㼯��(֧����ʽ��֧�����͡�������Ϣ���ɹ���־)
'                      |������Ϣ���Ӽ������ظ���ֻΪ�������������ҽ�����㷽ʽȡֵ
'           �NclsSubPayInfo - |��������ҽ��������Ϣ(���㷽ʽ�������������Ϣ)
'                             |ע�⣺�ڸ����Ӽ���ʱ��clsPayInfo��֧������Զ�����
'1.�Һŷ��뿨�ѷֿ��洢�����������ţ��������˷�ʱҪ���ҺŷѺͿ�����ͬһ��������֧�������
'2.ȡ�ҺŽ�����Ϣ����ж��ܽ���Ƿ�Ϊ0��ֻ�����ڶ�������������飬����ͨ�������ID�ͽ�����������
'3.ȡ���ѽ�����ϢҪ�жϲ�������
'4.Ϊ֧����������ҽ�����ض��ֽ��㷽ʽ���ҺŽ������ݷּ�����
'*****************************************************************************************************

Public Enum gPagePay
    Pay_CashPay = 0 '�ֽ�֧��
    Pay_ThreePay = 1 '������֧��
    Pay_SquarePay = 2 '���ѿ�֧��
    Pay_AccountPay = 3 '�˻�֧��
    Pay_InsurePay = 4 'ҽ��֧��
    Pay_OneCard = 5 '��һ��ͨ
End Enum

Public Enum gCardPay
    CP_NoCard = 0 'û�з���
    CP_���˺�
    CP_����
    CP_��
    CP_�˿�
    CP_���
End Enum

'��������ֵ�ľֲ�����
Private mvar��ʽ���� As String '�ֲ�����
Private mvarԤ���� As String '�ֲ�����
Private mvarobjPayInfos As clsPayInfos '�ֲ�����
Private mvarobjPati As zlOneCardComLib.clsPatientInfo  '�ֲ�����
'������Ϣ
Private mvarstrYBPati As String '�ֲ�����
Private mvarintInsure As Integer '�ֲ�����
Private mvarʵ����֤ As Boolean '�ֲ�����
Private mvar��ʱ������ As Boolean '�ֲ�����
Private mvarcllDrug As Collection '���˹���ҩ��
Private mvarcllImmune As Collection '�������߼�¼
Private mvarcllCert As Collection '����֤����Ϣ
Private mvarcllOther As Collection '���˴ӱ���Ϣ
Private mvarcllContact As Collection '������ϵ����Ϣ
'�Һ���Ϣ
Private mvar�Ʒѷ�ʽ As Integer '�ֲ�����
Private mvar��ǰ֧�� As Double '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar�Ǽ�ʱ�� As Date '�ֲ�����
Private mvar����ʱ�� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As Long '�ֲ�����
Private mvar�Ƿ��� As Boolean '�ֲ�����
'������Ϣ
Private mvarCard_���ƿ� As Boolean '�ֲ�����
Private mvarCard_�����ID As Long '�ֲ�����
Private mvarCard_�䶯���� As gCardPay '�ֲ�����
Private mvarCard_���� As String '�ֲ�����
Private mvarCard_��ĿID As Long '�ֲ�����

Public Property Set cllContact(ByVal vData As Collection)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.cllContact = Form1
    Set mvarcllContact = vData
End Property


Public Property Get cllContact() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.cllContact
    Set cllContact = mvarcllContact
End Property



Public Property Set cllOther(ByVal vData As Collection)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.cllOther = Form1
    Set mvarcllOther = vData
End Property


Public Property Get cllOther() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.cllOther
    Set cllOther = mvarcllOther
End Property



Public Property Set cllCert(ByVal vData As Collection)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.cllCert = Form1
    Set mvarcllCert = vData
End Property


Public Property Get cllCert() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.cllCert
    Set cllCert = mvarcllCert
End Property


Public Property Set cllImmune(ByVal vData As Collection)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.cllImmune = Form1
    Set mvarcllImmune = vData
End Property


Public Property Get cllImmune() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.cllImmune
    Set cllImmune = mvarcllImmune
End Property



Public Property Set cllDrug(ByVal vData As Collection)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.cllDrug = Form1
    Set mvarcllDrug = vData
End Property


Public Property Get cllDrug() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.cllDrug
    Set cllDrug = mvarcllDrug
End Property



Public Property Let Card_��ĿID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_��ĿID = 5
    mvarCard_��ĿID = vData
End Property


Public Property Get Card_��ĿID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_��ĿID
    Card_��ĿID = mvarCard_��ĿID
End Property




Public Property Let �Ʒѷ�ʽ(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ʒѷ�ʽ = 5
    mvar�Ʒѷ�ʽ = vData
End Property


Public Property Get �Ʒѷ�ʽ() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ʒѷ�ʽ
    �Ʒѷ�ʽ = mvar�Ʒѷ�ʽ
End Property




Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
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

Public Sub Clear(Optional ByVal blnClearPati As Boolean = True, Optional ByVal blnClearInsure As Boolean = True)
    If blnClearPati Then
        mvarʵ����֤ = False
        mvar��ʱ������ = False
    End If
        
    If blnClearPati Or blnClearInsure Then
        mvarstrYBPati = ""
        mvarintInsure = 0
    End If
    mvar��ǰ֧�� = 0
    mvar����ID = 0
    mvar�Ǽ�ʱ�� = 0
    mvar�Ƿ��� = False
    mvarCard_���ƿ� = False
    mvarCard_�����ID = 0
    mvarCard_�䶯���� = CP_NoCard
    mvarCard_���� = ""
    mvar��ʽ���� = ""
    mvarԤ���� = ""
    Set mvarobjPayInfos = New clsPayInfos
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
End Sub

Public Sub ClearPati()
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
    Set mvarcllDrug = Nothing  '���˹���ҩ��
    Set mvarcllImmune = Nothing '�������߼�¼
    Set mvarcllCert = Nothing  '����֤����Ϣ
    Set mvarcllOther = Nothing '���˴ӱ���Ϣ
    Set mvarcllContact = Nothing '������ϵ����Ϣ
    mvarstrYBPati = ""
    mvarintInsure = 0
    mvarʵ����֤ = False
    mvar��ʱ������ = False
End Sub


Public Property Let �Ǽ�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar�Ǽ�ʱ�� = vData
End Property


Public Property Get �Ǽ�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ǽ�ʱ��
    �Ǽ�ʱ�� = mvar�Ǽ�ʱ��
End Property


Public Property Let ��ǰ֧��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ֧�� = 5
    mvar��ǰ֧�� = vData
End Property


Public Property Get ��ǰ֧��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ֧��
    ��ǰ֧�� = mvar��ǰ֧��
End Property


Public Property Let �Ƿ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��� = 5
    mvar�Ƿ��� = vData
End Property


Public Property Get �Ƿ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ���
    �Ƿ��� = mvar�Ƿ���
End Property



Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property


Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mvar����ID
End Property


Public Property Let ��ʱ������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ʱ������ = 5
    mvar��ʱ������ = vData
End Property


Public Property Get ��ʱ������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ʱ������
    ��ʱ������ = mvar��ʱ������
End Property



Public Property Let ʵ����֤(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ����֤ = 5
    mvarʵ����֤ = vData
End Property


Public Property Get ʵ����֤() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ����֤
    ʵ����֤ = mvarʵ����֤
End Property


Public Property Let intInsure(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intInsure = 5
    mvarintInsure = vData
End Property


Public Property Get intInsure() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intInsure
    intInsure = mvarintInsure
End Property



Public Property Let strYBPati(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strYBPati = 5
    mvarstrYBPati = vData
End Property


Public Property Get strYBPati() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strYBPati
    strYBPati = mvarstrYBPati
End Property


'
'Public Property Let ����ID(ByVal vData As Long)
''������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
''Syntax: X.����ID = 5
'    mvar����ID = vData
'End Property
'
'
'Public Property Get ����ID() As Long
''��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
''Syntax: Debug.Print X.����ID
'    ����ID = mvar����ID
'End Property
'
'
'
'Public Property Let �ѱ�(ByVal vData As String)
''������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
''Syntax: X.�ѱ� = 5
'    mvar�ѱ� = vData
'End Property
'
'
'Public Property Get �ѱ�() As String
''��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
''Syntax: Debug.Print X.�ѱ�
'    �ѱ� = mvar�ѱ�
'End Property
'
'
'
'Public Property Let ����(ByVal vData As String)
''������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
''Syntax: X.���� = 5
'    mvar���� = vData
'End Property
'
'
'Public Property Get ����() As String
''��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
''Syntax: Debug.Print X.����
'    ���� = mvar����
'End Property
'
'
'
'Public Property Let �Ա�(ByVal vData As String)
''������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
''Syntax: X.�Ա� = 5
'    mvar�Ա� = vData
'End Property
'
'
'Public Property Get �Ա�() As String
''��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
''Syntax: Debug.Print X.�Ա�
'    �Ա� = mvar�Ա�
'End Property
'
'
'
'Public Property Let ����(ByVal vData As String)
''������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
''Syntax: X.���� = 5
'    mvar���� = vData
'End Property
'
'
'Public Property Get ����() As String
''��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
''Syntax: Debug.Print X.����
'    ���� = mvar����
'End Property



Public Property Let Card_�����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_�����ID = 5
    mvarCard_�����ID = vData
End Property


Public Property Get Card_�����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_�����ID
    Card_�����ID = mvarCard_�����ID
End Property



Public Property Let Card_�䶯����(ByVal vData As gCardPay)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_�䶯���� = 5
    mvarCard_�䶯���� = vData
End Property


Public Property Get Card_�䶯����() As gCardPay
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_�䶯����
    Card_�䶯���� = mvarCard_�䶯����
End Property



Public Property Let Card_����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���� = 5
    mvarCard_���� = vData
End Property


Public Property Get Card_����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_����
    Card_���� = mvarCard_����
End Property



Public Property Let Card_���ƿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Card_���ƿ� = 5
    mvarCard_���ƿ� = vData
End Property


Public Property Get Card_���ƿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Card_���ƿ�
    Card_���ƿ� = mvarCard_���ƿ�
End Property



Public Property Set objPayInfos(ByVal vData As clsPayInfos)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objPayInfos = Form1
    Set mvarobjPayInfos = vData
End Property


Public Property Get objPayInfos() As clsPayInfos
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objPayInfos
    Set objPayInfos = mvarobjPayInfos
End Property


Public Property Set objPati(ByVal vData As zlOneCardComLib.clsPatientInfo)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objPati = Form1
    Set mvarobjPati = vData
End Property


Public Property Get objPati() As zlOneCardComLib.clsPatientInfo
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objPati
    Set objPati = mvarobjPati
End Property



Public Property Let Ԥ����(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Ԥ���� = Form1
    mvarԤ���� = vData
End Property


Public Property Get Ԥ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ����
    Ԥ���� = mvarԤ����
End Property



Public Property Let ��ʽ����(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.��ʽ���� = Form1
    mvar��ʽ���� = vData
End Property


Public Property Get ��ʽ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ʽ����
    ��ʽ���� = mvar��ʽ����
End Property

Private Sub Class_Initialize()
    Set mvarobjPayInfos = New clsPayInfos
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
End Sub

Private Sub Class_Terminate()
    Set mvarobjPayInfos = Nothing
    Set mvarobjPati = Nothing
End Sub