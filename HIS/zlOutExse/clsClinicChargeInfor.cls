VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicChargeInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'��������ֵ�ľֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarShareUserID As Long '�ֲ�����
Private mvarPatiUseType As String '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar�շѽ��� As String '�ֲ�����
Private mvarԤ����� As String '�ֲ�����
Private mvar��ǰ��Ʊ�� As String '�ֲ�����
Private mvarintInsure As Integer '�ֲ�����
Private mvarӦ�ս�� As Double '�ֲ�����
Private mvarʵ�ս�� As Double '�ֲ�����
Private mvar����Ӧ�� As Double  '�ֲ�����
Private mvarҽ��Ԥ���� As Double '�ֲ�����
Private mvarҽ�������� As Double '�ֲ�����
Private mvar���Ѻϼ� As Double '�ֲ�����
Private mvarҽ������Ʊ�� As Boolean   '�ֲ�����
Private mvarӦ���ۼ� As Double '�ֲ�����
Private mvar�ɿ� As Double '�ֲ�����
Private mvar�Ҳ� As Double '�ֲ�����
Private mvarNos As String '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����ID As Long  '�ֲ�����
Private mvar������� As Long '�ֲ�����
Private mvarrsBalance As ADODB.Recordset   '�ֲ�����
Private mvarrs���㷽ʽ As ADODB.Recordset   '�ֲ�����
Private mvar����IDs As String '�ֲ�����
Private mvarȱʡ���㷽ʽ As String '�ֲ�����
Private mvar������Դ As String

Public Property Let ������Դ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������Դ = 5
    mvar������Դ = vData
End Property


Public Property Get ������Դ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������Դ
    ������Դ = mvar������Դ
End Property


Public Property Let ȱʡ���㷽ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ȱʡ���㷽ʽ = 5
    mvarȱʡ���㷽ʽ = vData
End Property


Public Property Get ȱʡ���㷽ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ȱʡ���㷽ʽ
    ȱʡ���㷽ʽ = mvarȱʡ���㷽ʽ
End Property

Public Property Let ����IDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����IDs = 5
    mvar����IDs = vData
End Property


Public Property Get ����IDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����IDs
    ����IDs = mvar����IDs
End Property


Public Property Let Nos(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Nos = 5
    mvarNos = vData
End Property


Public Property Get Nos() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Nos
    Nos = mvarNos
End Property



Public Property Let �Ҳ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ҳ� = 5
    mvar�Ҳ� = vData
End Property


Public Property Get �Ҳ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ҳ�
    �Ҳ� = mvar�Ҳ�
End Property

Public Property Let �ɿ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ɿ� = 5
    mvar�ɿ� = vData
End Property


Public Property Get �ɿ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ɿ�
    �ɿ� = mvar�ɿ�
End Property



Public Property Let Ӧ���ۼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ӧ���ۼ� = 5
    mvarӦ���ۼ� = vData
End Property


Public Property Get Ӧ���ۼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ӧ���ۼ�
    Ӧ���ۼ� = mvarӦ���ۼ�
End Property

Public Property Let ����Ӧ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����Ӧ�� = 5
    mvar����Ӧ�� = vData
End Property


Public Property Get ����Ӧ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����Ӧ��
    ����Ӧ�� = mvar����Ӧ��
End Property
 


Public Property Let ҽ��������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ�������� = 5
    mvarҽ�������� = vData
End Property


Public Property Get ҽ��������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ��������
    ҽ�������� = mvarҽ��������
End Property




Public Property Let ҽ��Ԥ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ��Ԥ���� = 5
    mvarҽ��Ԥ���� = vData
End Property


Public Property Get ҽ��Ԥ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ��Ԥ����
    ҽ��Ԥ���� = mvarҽ��Ԥ����
End Property



Public Property Let ʵ�ս��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ�ս�� = 5
    mvarʵ�ս�� = vData
End Property


Public Property Get ʵ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ�ս��
    ʵ�ս�� = mvarʵ�ս��
End Property



Public Property Let Ӧ�ս��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ӧ�ս�� = 5
    mvarӦ�ս�� = vData
End Property


Public Property Get Ӧ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ӧ�ս��
    Ӧ�ս�� = mvarӦ�ս��
End Property

Public Property Let ���Ѻϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���Ѻϼ� = 5
    mvar���Ѻϼ� = vData
End Property


Public Property Get ���Ѻϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���Ѻϼ�
    ���Ѻϼ� = mvar���Ѻϼ�
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

Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property


Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ա�
    �Ա� = mvar�Ա�
End Property



Public Property Let �ѱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ѱ� = 5
    mvar�ѱ� = vData
End Property


Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ѱ�
    �ѱ� = mvar�ѱ�
End Property
Public Property Let �շѽ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շѽ��� = 5
    mvar�շѽ��� = vData
End Property
Public Property Get �շѽ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�շѽ���
    �շѽ��� = mvar�շѽ���
End Property
Public Property Let Ԥ�����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ����� = 5
    mvarԤ����� = vData
End Property
Public Property Get Ԥ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����
    Ԥ����� = mvarԤ�����
End Property
 
 Public Property Let ��ǰ��Ʊ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ��Ʊ�� = 5
    mvar��ǰ��Ʊ�� = vData
End Property
Public Property Get ��ǰ��Ʊ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ��Ʊ��
    ��ǰ��Ʊ�� = mvar��ǰ��Ʊ��
End Property
 
Public Property Let PatiUseType(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatiUseType = 5
    mvarPatiUseType = vData
End Property


Public Property Get PatiUseType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatiUseType
    PatiUseType = mvarPatiUseType
End Property



Public Property Let ShareUserID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ShareUserID = 5
    mvarShareUserID = vData
End Property


Public Property Get ShareUserID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ShareUserID
    ShareUserID = mvarShareUserID
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

Public Property Get ����ID() As Long
    ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
    mvar����ID = vData
End Property

Public Property Get �������() As Long
    ������� = mvar�������
End Property

Public Property Let �������(ByVal vData As Long)
    mvar������� = vData
End Property

Public Property Let ҽ������Ʊ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ������Ʊ�� = 5
    mvarҽ������Ʊ�� = vData
End Property

Public Property Get ҽ������Ʊ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ������Ʊ��
    ҽ������Ʊ�� = mvarҽ������Ʊ��
End Property
Public Property Get ����ID() As Long
   ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
   mvar����ID = vData
End Property
 
Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.rsBalance = 5
    Set mvarrsBalance = vData
End Property
Public Property Get rsBalance() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rsBalance
   Set rsBalance = mvarrsBalance
End Property
Public Property Set rs���㷽ʽ(ByVal vData As ADODB.Recordset)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.rs���㷽ʽ = 5
    Set mvarrs���㷽ʽ = vData
End Property
Public Property Get rs���㷽ʽ() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rs���㷽ʽ
   Set rs���㷽ʽ = mvarrs���㷽ʽ
End Property

 