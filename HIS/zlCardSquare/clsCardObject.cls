VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'��������ֵ�ľֲ�����
Private mvarCardObject As Object '�ֲ�����
Private mvar���ƿ� As Boolean '�ֲ�����
Private mvar�ӿ���� As Long '�ֲ�����
Private mvarCardPreporty As clsCard '�ֲ�����
Private mvar���ѿ� As Boolean '�ֲ�����
Private mvarInitCompents As Boolean '�ֲ�����
Public Property Let InitCompents(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ѿ� = 5
    mvarInitCompents = vData
End Property
Public Property Get InitCompents() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ѿ�
    InitCompents = mvarInitCompents
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



Public Property Set CardPreporty(ByVal vData As clsCard)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.CardPreporty = Form1
    Set mvarCardPreporty = vData
End Property


Public Property Get CardPreporty() As clsCard
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CardPreporty
    Set CardPreporty = mvarCardPreporty
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



Public Property Let ���ƿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ƿ� = 5
    mvar���ƿ� = vData
End Property


Public Property Get ���ƿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ƿ�
    ���ƿ� = mvar���ƿ�
End Property



Public Property Set CardObject(ByVal vData As Object)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.CardObject = Form1
    Set mvarCardObject = vData
End Property


Public Property Get CardObject() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CardObject
    Set CardObject = mvarCardObject
End Property


