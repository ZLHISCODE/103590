VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvarobjCard As Card '�ֲ�����
Private mvar�����ID As Long '�ֲ�����
Private mvar�Ƿ�ת�� As Boolean '�ֲ�����
Private mvarԤ��ID As Long '�ֲ�����
Private mvar������ˮ�� As String '�ֲ�����
Private mvar����˵�� As String '�ֲ�����
Private mvar������ As Double '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvarTag As String '�ֲ�����
Private mvar�Ƿ��˿�ֽ��� As Boolean '�ֲ�����
Private mvar�к� As Long '�ֲ�����
Private mvar���㷽ʽ As String '�ֲ�����
Private mvarУ�Ա�־ As Byte '�ֲ�����
Private mvar������� As Boolean '�ֲ�����
Private mvar����ʱ�� As Date '�ֲ�����
Private mvarʣ���� As Double '�ֲ�����
Private mvarobjTag As Object '�ֲ�����
Private mvar������� As String '�ֲ�����
Private mvar����ժҪ As String '�ֲ�����
Private mvar�ɿ��� As Double '�ֲ�����
Private mvar�Ҳ� As Double '�ֲ�����
Private mvar���� As Double '�ֲ�����
Private mvar���ѿ� As Boolean '�ֲ�����
Private mvar���ѿ�ID As Long '�ֲ�����
Private mvar�Ƿ����� As Boolean '�ֲ�����
Private mvar������� As String '�ֲ�����
Private mvar�ʻ���� As Double '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ƿ������༭ As Boolean '�ֲ�����
Private mvar�Ƿ�����ɾ�� As Boolean '�ֲ�����
Private mvar����IDs As String '�ֲ�����
Private mvar��������ID As Long '�ֲ�����
Private mvar�Ƿ���� As Boolean '�ֲ�����
Private mvar�������� As Integer '�ֲ�����
Private mvar�Ƿ�Ԥ�� As Boolean '�ֲ�����
Private mvarԭʼ��� As Double '�ֲ�����
Private mvar�������� As Integer '�ֲ�����
Private mvarδ�˽�� As Double '�ֲ�����
Private mvar�Ƿ��˿� As Boolean '�ֲ�����
Private mvar�Ƿ�ȱʡ As Boolean '�ֲ�����
Private mvar�Ƿ��������� As Boolean '�ֲ�����
Private mvar�Ƿ�ǿ������ As Boolean '�ֲ�����
Private mvar�Ƿ񱣴� As Boolean '�ֲ�����
Private mvar�˿����ˮ�� As String '�ֲ�����
Private mvar�˿��˵�� As String '�ֲ�����
Private mvar�Ƿ��ѻ�ҽ�� As Boolean '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvarQRCode As String '�ֲ�����
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



Public Property Let �Ƿ��ѻ�ҽ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��ѻ�ҽ�� = 5
    mvar�Ƿ��ѻ�ҽ�� = vData
End Property


Public Property Get �Ƿ��ѻ�ҽ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��ѻ�ҽ��
    �Ƿ��ѻ�ҽ�� = mvar�Ƿ��ѻ�ҽ��
End Property



Public Property Let �˿��˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˿��˵�� = 5
    mvar�˿��˵�� = vData
End Property


Public Property Get �˿��˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˿��˵��
    �˿��˵�� = mvar�˿��˵��
End Property



Public Property Let �˿����ˮ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˿����ˮ�� = 5
    mvar�˿����ˮ�� = vData
End Property


Public Property Get �˿����ˮ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˿����ˮ��
    �˿����ˮ�� = mvar�˿����ˮ��
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



Public Property Let �Ƿ�ǿ������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�ǿ������ = 5
    mvar�Ƿ�ǿ������ = vData
End Property


Public Property Get �Ƿ�ǿ������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�ǿ������
    �Ƿ�ǿ������ = mvar�Ƿ�ǿ������
End Property



Public Property Let �Ƿ���������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��������� = 5
    mvar�Ƿ��������� = vData
End Property


Public Property Get �Ƿ���������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ���������
    �Ƿ��������� = mvar�Ƿ���������
End Property



Public Property Let �Ƿ�ȱʡ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�ȱʡ = 5
    mvar�Ƿ�ȱʡ = vData
End Property


Public Property Get �Ƿ�ȱʡ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�ȱʡ
    �Ƿ�ȱʡ = mvar�Ƿ�ȱʡ
End Property



Public Property Let �Ƿ��˿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��˿� = 5
    mvar�Ƿ��˿� = vData
End Property


Public Property Get �Ƿ��˿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��˿�
    �Ƿ��˿� = mvar�Ƿ��˿�
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



Public Property Let �Ƿ�Ԥ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�Ԥ�� = 5
    mvar�Ƿ�Ԥ�� = vData
End Property


Public Property Get �Ƿ�Ԥ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�Ԥ��
    �Ƿ�Ԥ�� = mvar�Ƿ�Ԥ��
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
Public Property Let �Ƿ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ���� = 5
    mvar�Ƿ���� = vData
End Property
Public Property Get �Ƿ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ����
    �Ƿ���� = mvar�Ƿ����
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



Public Property Let �Ƿ�����ɾ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�����ɾ�� = 5
    mvar�Ƿ�����ɾ�� = vData
End Property


Public Property Get �Ƿ�����ɾ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�����ɾ��
    �Ƿ�����ɾ�� = mvar�Ƿ�����ɾ��
End Property



Public Property Let �Ƿ������༭(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ������༭ = 5
    mvar�Ƿ������༭ = vData
End Property


Public Property Get �Ƿ������༭() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ������༭
    �Ƿ������༭ = mvar�Ƿ������༭
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



Public Property Let �ʻ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ʻ���� = 5
    mvar�ʻ���� = vData
End Property


Public Property Get �ʻ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ʻ����
    �ʻ���� = mvar�ʻ����
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



Public Property Let �Ƿ�����(ByVal vData As Boolean)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.�Ƿ����� = Form1
      mvar�Ƿ����� = vData
End Property


Public Property Get �Ƿ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�����
      �Ƿ����� = mvar�Ƿ�����
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



Public Property Let �ɿ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ɿ��� = 5
    mvar�ɿ��� = vData
End Property


Public Property Get �ɿ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ɿ���
     �ɿ��� = mvar�ɿ���
End Property



Public Property Let ����ժҪ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ժҪ = 5
    mvar����ժҪ = vData
End Property


Public Property Get ����ժҪ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ժҪ
    ����ժҪ = mvar����ժҪ
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



Public Property Let ʣ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʣ���� = 5
    mvarʣ���� = vData
End Property


Public Property Get ʣ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʣ����
    ʣ���� = mvarʣ����
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



Public Property Let �������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property


Public Property Get �������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mvar�������
End Property



Public Property Let У�Ա�־(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.У�Ա�־ = 5
    mvarУ�Ա�־ = vData
End Property


Public Property Get У�Ա�־() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.У�Ա�־
    У�Ա�־ = mvarУ�Ա�־
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



Public Property Let �к�(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�к� = 5
    mvar�к� = vData
End Property


Public Property Get �к�() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�к�
    �к� = mvar�к�
End Property



Public Property Let �Ƿ��˿�ֽ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��˿�ֽ��� = 5
    mvar�Ƿ��˿�ֽ��� = vData
End Property


Public Property Get �Ƿ��˿�ֽ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��˿�ֽ���
    �Ƿ��˿�ֽ��� = mvar�Ƿ��˿�ֽ���
End Property



Public Property Let Tag(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
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



Public Property Let Ԥ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ��ID = 5
    mvarԤ��ID = vData
End Property


Public Property Get Ԥ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ��ID
    Ԥ��ID = mvarԤ��ID
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



Public Property Let �����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����ID = 5
    mvar�����ID = vData
End Property


Public Property Get �����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����ID
    �����ID = mvar�����ID
End Property



Public Property Set objCard(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objCard = Form1
    Set mvarobjCard = vData
End Property


Public Property Get objCard() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objCard
    Set objCard = mvarobjCard
End Property


