VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsurePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvar�ֱҴ��� As Boolean '�ֲ�����
Private mvarҽ���ӿڴ�ӡƱ�� As Boolean '�ֲ�����
Private mvar���ﲡ�˽������� As Boolean '�ֲ�����
Private mvar����Ԥ���� As Boolean '�ֲ�����
Private mvar������봫����ϸ As Boolean '�ֲ�����
Private mvar�������_�������� As Boolean '�ֲ�����
Private mvarδ�����Ժ As Boolean '�ֲ�����
Private mvar����ʹ�ø����ʻ� As Boolean '�ֲ�����
Private mvar��Ժ��������Ժ As Boolean '�ֲ�����
Private mvar��Ժ���˽������� As Boolean '�ֲ�����
Private mvar��;���ʽ��������ϴ����� As Boolean '�ֲ�����
Private mvar�������ú���ýӿ� As Boolean '�ֲ�����
Private mvar�������Ϻ��ӡ�ص� As Boolean '�ֲ�����
Private mvarסԺ�������� As Boolean '�ֲ�����
Private mvar��������סԺ���� As Boolean '�ֲ�����
Public Property Let ��������סԺ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������סԺ���� = 5
    mvar��������סԺ���� = vData
End Property


Public Property Get ��������סԺ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������סԺ����
    ��������סԺ���� = mvar��������סԺ����
End Property



Public Property Let סԺ��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ�������� = 5
    mvarסԺ�������� = vData
End Property


Public Property Get סԺ��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.סԺ��������
    סԺ�������� = mvarסԺ��������
End Property



Public Property Let �������Ϻ��ӡ�ص�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������Ϻ��ӡ�ص� = 5
    mvar�������Ϻ��ӡ�ص� = vData
End Property


Public Property Get �������Ϻ��ӡ�ص�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������Ϻ��ӡ�ص�
    �������Ϻ��ӡ�ص� = mvar�������Ϻ��ӡ�ص�
End Property



Public Property Let �������ú���ýӿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������ú���ýӿ� = 5
    mvar�������ú���ýӿ� = vData
End Property


Public Property Get �������ú���ýӿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������ú���ýӿ�
    �������ú���ýӿ� = mvar�������ú���ýӿ�
End Property



Public Property Let ��;���ʽ��������ϴ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��;���ʽ��������ϴ����� = 5
    mvar��;���ʽ��������ϴ����� = vData
End Property


Public Property Get ��;���ʽ��������ϴ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��;���ʽ��������ϴ�����
    ��;���ʽ��������ϴ����� = mvar��;���ʽ��������ϴ�����
End Property



Public Property Let ��Ժ���˽�������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ժ���˽������� = 5
    mvar��Ժ���˽������� = vData
End Property


Public Property Get ��Ժ���˽�������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ժ���˽�������
    ��Ժ���˽������� = mvar��Ժ���˽�������
End Property



Public Property Let ��Ժ��������Ժ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ժ��������Ժ = 5
    mvar��Ժ��������Ժ = vData
End Property


Public Property Get ��Ժ��������Ժ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ժ��������Ժ
    ��Ժ��������Ժ = mvar��Ժ��������Ժ
End Property



Public Property Let ����ʹ�ø����ʻ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʹ�ø����ʻ� = 5
    mvar����ʹ�ø����ʻ� = vData
End Property


Public Property Get ����ʹ�ø����ʻ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʹ�ø����ʻ�
    ����ʹ�ø����ʻ� = mvar����ʹ�ø����ʻ�
End Property



Public Property Let δ�����Ժ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ�����Ժ = 5
    mvarδ�����Ժ = vData
End Property


Public Property Get δ�����Ժ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ�����Ժ
    δ�����Ժ = mvarδ�����Ժ
End Property



Public Property Let �������_��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������_�������� = 5
    mvar�������_�������� = vData
End Property


Public Property Get �������_��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������_��������
    �������_�������� = mvar�������_��������
End Property



Public Property Let ������봫����ϸ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������봫����ϸ = 5
    mvar������봫����ϸ = vData
End Property


Public Property Get ������봫����ϸ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������봫����ϸ
    ������봫����ϸ = mvar������봫����ϸ
End Property



Public Property Let ����Ԥ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����Ԥ���� = 5
    mvar����Ԥ���� = vData
End Property


Public Property Get ����Ԥ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����Ԥ����
    ����Ԥ���� = mvar����Ԥ����
End Property



Public Property Let ���ﲡ�˽�������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ﲡ�˽������� = 5
    mvar���ﲡ�˽������� = vData
End Property


Public Property Get ���ﲡ�˽�������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ﲡ�˽�������
    ���ﲡ�˽������� = mvar���ﲡ�˽�������
End Property



Public Property Let ҽ���ӿڴ�ӡƱ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ���ӿڴ�ӡƱ�� = 5
    mvarҽ���ӿڴ�ӡƱ�� = vData
End Property


Public Property Get ҽ���ӿڴ�ӡƱ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ���ӿڴ�ӡƱ��
    ҽ���ӿڴ�ӡƱ�� = mvarҽ���ӿڴ�ӡƱ��
End Property



Public Property Let �ֱҴ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֱҴ��� = 5
    mvar�ֱҴ��� = vData
End Property


Public Property Get �ֱҴ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֱҴ���
    �ֱҴ��� = mvar�ֱҴ���
End Property


