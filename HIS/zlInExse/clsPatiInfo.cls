VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvarPatient As clsPatientInfo '�ֲ�����

Private mvarԤ����� As Double '�ֲ�����
Private mvar������� As Double '�ֲ�����
Private mvarԤ��ʣ��ϼ� As Double '�ֲ�����
Private mvarʵ����� As Double '�ֲ�����

Private mvarδ���ۼ� As Double '�ֲ�����
Private mvar�������� As Integer '�ֲ�����
Private mvarbln�˿� As Boolean '�ֲ�����
Private mvarbln�������� As Boolean '�ֲ�����
Private mvar�Ƿ��Լ��λ As Boolean '�ֲ�����

Public Property Set Patient(ByVal vData As clsPatientInfo)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Patient = Form1
    Set mvarPatient = vData
End Property


Public Property Get Patient() As clsPatientInfo
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Patient
    Set Patient = mvarPatient
End Property




Public Property Let �Ƿ��Լ��λ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��Լ��λ = 5
    mvar�Ƿ��Լ��λ = vData
End Property


Public Property Get �Ƿ��Լ��λ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��Լ��λ
    �Ƿ��Լ��λ = mvar�Ƿ��Լ��λ
End Property
Public Property Let bln��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln�������� = 5
    mvarbln�������� = vData
End Property


Public Property Get bln��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln��������
    bln�������� = mvarbln��������
End Property



Public Property Let bln�˿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln�˿� = 5
    mvarbln�˿� = vData
End Property


Public Property Get bln�˿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln�˿�
    bln�˿� = mvarbln�˿�
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



Public Property Let δ���ۼ�(ByVal vData As Double)
Attribute δ���ۼ�.VB_Description = "�ϴ�δ���ۼƽ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ���ۼ� = 5
    mvarδ���ۼ� = vData
End Property


Public Property Get δ���ۼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ���ۼ�
      δ���ۼ� = mvarδ���ۼ�
End Property



Public Property Let ʵ�����(ByVal vData As Double)
Attribute ʵ�����.VB_Description = "����Ԥ����ϸ�е�ʵ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ����� = 5
    mvarʵ����� = vData
End Property


Public Property Get ʵ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ�����
    ʵ����� = mvarʵ�����
End Property



Public Property Let Ԥ��ʣ��ϼ�(ByVal vData As Double)
Attribute Ԥ��ʣ��ϼ�.VB_Description = "Ԥ�����-δ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ��ʣ��ϼ� = 5
    mvarԤ��ʣ��ϼ� = vData
End Property


Public Property Get Ԥ��ʣ��ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ��ʣ��ϼ�
    Ԥ��ʣ��ϼ� = mvarԤ��ʣ��ϼ�
End Property



Public Property Let �������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property


Public Property Get �������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mvar�������
End Property



Public Property Let Ԥ�����(ByVal vData As Double)
Attribute Ԥ�����.VB_Description = "Ԥ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ����� = 5
    mvarԤ����� = vData
End Property


Public Property Get Ԥ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����
    Ԥ����� = mvarԤ�����
End Property