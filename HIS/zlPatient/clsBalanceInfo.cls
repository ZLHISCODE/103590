VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��������ֵ�ľֲ�����
Private mvar����δ�� As Double '�ֲ�����
Private mvar��ǰ���� As Double '�ֲ�����
Private mvar�Ѹ��ϼ� As Double '�ֲ�����
Private mvarδ���ϼ� As Double '�ֲ�����
Private mvarԤ�����ܶ� As Double '�ֲ�����
Private mvarҽ��֧���ϼ� As Double '�ֲ�����
Private mvar��Ԥ���ϼ� As Double '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar�ɿ� As Double '�ֲ�����
Private mvar�Ҳ� As Double '�ֲ�����
Private mvar���� As Double '�ֲ�����
Private mvar�ֽ�֧�� As Double '�ֲ�����
Private mvar��Ԥ��ID As Long '�ֲ�����
Private mvar��Ԥ�����ݺ� As String '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����ʱ�� As Date '�ֲ�����
Private mvar�Ƿ񱣴���ʵ� As Boolean '�ֲ�����
Private mvar���ʵ��ݺ� As String '�ֲ�����
Private mvar�Ƿ�Ԥ���鿨 As Boolean '�ֲ�����
Private mvar����ԭ�� As String '�ֲ�����
Private mvarobjInsure As clsInsureInfor '�ֲ�����
Private mvar�������� As Byte '�ֲ�����
Private mvar�Ƿ���;���� As Boolean '�ֲ�����
Private mvar�Ƿ񱣴�Ԥ�� As Boolean '�ֲ�����
Private mvar�Ƿ���Ԥ���˿� As Boolean '�ֲ�����
Private mvar��֧Ʊ�� As Double '�ֲ�����
Public Property Let ��֧Ʊ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��֧Ʊ�� = 5
    mvar��֧Ʊ�� = vData
End Property


Public Property Get ��֧Ʊ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��֧Ʊ��
    ��֧Ʊ�� = mvar��֧Ʊ��
End Property



Public Property Let �Ƿ���Ԥ���˿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ���Ԥ���˿� = 5
    mvar�Ƿ���Ԥ���˿� = vData
End Property


Public Property Get �Ƿ���Ԥ���˿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ���Ԥ���˿�
    �Ƿ���Ԥ���˿� = mvar�Ƿ���Ԥ���˿�
End Property




Public Property Let �Ƿ񱣴�Ԥ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ񱣴�Ԥ�� = 5
    mvar�Ƿ񱣴�Ԥ�� = vData
End Property


Public Property Get �Ƿ񱣴�Ԥ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ񱣴�Ԥ��
    �Ƿ񱣴�Ԥ�� = mvar�Ƿ񱣴�Ԥ��
End Property



Public Property Let �Ƿ���;����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ���;���� = 5
    mvar�Ƿ���;���� = vData
End Property


Public Property Get �Ƿ���;����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ���;����
    �Ƿ���;���� = mvar�Ƿ���;����
End Property



Public Property Let ��������(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property



Public Property Set objInsure(ByVal vData As clsInsureInfor)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objInsure = Form1
    Set mvarobjInsure = vData
End Property


Public Property Get objInsure() As clsInsureInfor
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objInsure
    Set objInsure = mvarobjInsure
End Property




Public Property Let ����ԭ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ԭ�� = 5
    mvar����ԭ�� = vData
End Property


Public Property Get ����ԭ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ԭ��
    ����ԭ�� = mvar����ԭ��
End Property



Public Property Let �Ƿ�Ԥ���鿨(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�Ԥ���鿨 = 5
    mvar�Ƿ�Ԥ���鿨 = vData
End Property


Public Property Get �Ƿ�Ԥ���鿨() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�Ԥ���鿨
    �Ƿ�Ԥ���鿨 = mvar�Ƿ�Ԥ���鿨
End Property



Public Property Let ���ʵ��ݺ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ʵ��ݺ� = 5
    mvar���ʵ��ݺ� = vData
End Property


Public Property Get ���ʵ��ݺ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ʵ��ݺ�
    ���ʵ��ݺ� = mvar���ʵ��ݺ�
End Property



Public Property Let �Ƿ񱣴���ʵ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ񱣴���ʵ� = 5
    mvar�Ƿ񱣴���ʵ� = vData
End Property


Public Property Get �Ƿ񱣴���ʵ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ񱣴���ʵ�
    �Ƿ񱣴���ʵ� = mvar�Ƿ񱣴���ʵ�
End Property



Public Property Let ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar����ʱ�� = vData
End Property


Public Property Get ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʱ��
    ����ʱ�� = mvar����ʱ��
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



Public Property Let ��Ԥ�����ݺ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ԥ�����ݺ� = 5
    mvar��Ԥ�����ݺ� = vData
End Property


Public Property Get ��Ԥ�����ݺ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ԥ�����ݺ�
    ��Ԥ�����ݺ� = mvar��Ԥ�����ݺ�
End Property



Public Property Let ��Ԥ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ԥ��ID = 5
    mvar��Ԥ��ID = vData
End Property


Public Property Get ��Ԥ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ԥ��ID
    ��Ԥ��ID = mvar��Ԥ��ID
End Property



Public Property Let �ֽ�֧��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֽ�֧�� = 5
    mvar�ֽ�֧�� = vData
End Property


Public Property Get �ֽ�֧��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֽ�֧��
    �ֽ�֧�� = mvar�ֽ�֧��
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



Public Property Let ��Ԥ���ϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ԥ���ϼ� = 5
    mvar��Ԥ���ϼ� = vData
End Property


Public Property Get ��Ԥ���ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ԥ���ϼ�
    ��Ԥ���ϼ� = mvar��Ԥ���ϼ�
End Property



Public Property Let ҽ��֧���ϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ��֧���ϼ� = 5
    mvarҽ��֧���ϼ� = vData
End Property


Public Property Get ҽ��֧���ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ��֧���ϼ�
    ҽ��֧���ϼ� = mvarҽ��֧���ϼ�
End Property



Public Property Let Ԥ�����ܶ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ�����ܶ� = 5
    mvarԤ�����ܶ� = vData
End Property


Public Property Get Ԥ�����ܶ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����ܶ�
    Ԥ�����ܶ� = mvarԤ�����ܶ�
End Property



Public Property Let δ���ϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ���ϼ� = 5
    mvarδ���ϼ� = vData
End Property


Public Property Get δ���ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ���ϼ�
    δ���ϼ� = mvarδ���ϼ�
End Property



Public Property Let �Ѹ��ϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ѹ��ϼ� = 5
    mvar�Ѹ��ϼ� = vData
End Property


Public Property Get �Ѹ��ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ѹ��ϼ�
    �Ѹ��ϼ� = mvar�Ѹ��ϼ�
End Property



Public Property Let ��ǰ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ���� = 5
    mvar��ǰ���� = vData
End Property


Public Property Get ��ǰ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ����
    ��ǰ���� = mvar��ǰ����
End Property



Public Property Let ����δ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����δ�� = 5
    mvar����δ�� = vData
End Property


Public Property Get ����δ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����δ��
     ����δ�� = mvar����δ��
End Property

Private Sub Class_Initialize()
    Set mvarobjInsure = New clsInsureInfor
End Sub

Private Sub Class_Terminate()
    Set mvarobjInsure = Nothing
End Sub