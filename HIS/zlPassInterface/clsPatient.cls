VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mlng����ID As Long '�ֲ�����
Private mlng��ҳID As Long '�ֲ�����
Private mlng�Һ�ID As Long '�ֲ�����
Private mdbl��ʶ�� As Double '�ֲ�����
Private mstr���� As String '�ֲ�����
Private mstr�Ա� As String '�ֲ�����
Private mstr���� As String '�ֲ�����
Private mintӤ�� As Integer '�ֲ�����
Private mstr�Һŵ� As String '�ֲ�����
Private mdat�������� As Date '�ֲ�����
Private mlngDeptID As Long  '�ֲ�����

Public Property Let dat��������(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.dat�������� = 5
    mdat�������� = vData
End Property

Public Property Get dat��������() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.dat��������
     dat�������� = mdat��������
End Property

Public Property Let str�Һŵ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.str�Һŵ� = 5
    mstr�Һŵ� = vData
End Property

Public Property Get str�Һŵ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.str�Һŵ�
     str�Һŵ� = mstr�Һŵ�
End Property


Public Property Let intӤ��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intӤ�� = 5
    mintӤ�� = vData
End Property


Public Property Get intӤ��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intӤ��
     intӤ�� = mintӤ��
End Property

Public Property Let str����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.str���� = 5
    mstr���� = vData
End Property

Public Property Get str����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.str����
     str���� = mstr����
End Property

Public Property Let str�Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.str�Ա� = 5
    mstr�Ա� = vData
End Property

Public Property Get str�Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.str�Ա�
     str�Ա� = mstr�Ա�
End Property

Public Property Let str����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.str���� = 5
    mstr���� = vData
End Property

Public Property Get str����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.str����
    str���� = mstr����
End Property

Public Property Let dbl��ʶ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.dbl��ʶ�� = 5
    mdbl��ʶ�� = vData
End Property

Public Property Get dbl��ʶ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.dbl��ʶ��
    dbl��ʶ�� = mdbl��ʶ��
End Property

Public Property Let lng�Һ�ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.lng�Һ�ID = 5
    mlng�Һ�ID = vData
End Property

Public Property Get lng�Һ�ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.lng�Һ�ID
    lng�Һ�ID = mlng�Һ�ID
End Property

Public Property Let lng��ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.lng��ҳID = 5
    mlng��ҳID = vData
End Property

Public Property Get lng��ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.lng��ҳID
    lng��ҳID = mlng��ҳID
End Property

Public Property Let lng����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.lng����ID = 5
    mlng����ID = vData
End Property

Public Property Get lng����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.lng����ID
    lng����ID = mlng����ID
End Property

Public Property Let lngDeptID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.lngDeptID = 5
    mlngDeptID = vData
End Property

Public Property Get lngDeptID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.lngDeptID
    lngDeptID = mlngDeptID
End Property
