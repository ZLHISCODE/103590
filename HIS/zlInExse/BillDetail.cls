VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Detail"
Option Explicit 'Ҫ���������

Public Key As String
Private mvar�շ�ϸĿID As Long '�ֲ�����
Private mvar��� As Integer  '�ֲ�����
Private mvar�������� As Integer  '�ֲ�����
Private mvar�������� As Integer  '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar��ҳID As Long '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarסԺ�� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar������ As Double   '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvarҽ�Ƹ��� As String '�ֲ�����
Private mvar�շ���� As String '�ֲ�����
Private mvar���㵥λ As String '�ֲ�����
Private mvar��ҩ���� As String '�ֲ�����
Private mvar���� As Integer  '�ֲ�����
Private mvar���� As Double '�ֲ�����
Private mvar���ӱ�־ As Byte '�ֲ�����
Private mvar���￨�� As String '�ֲ�����
Private mvarִ�в���ID As Long '�ֲ�����
Private mvarInComes As BillInComes '�ֲ�����
Private mvarDetail As Detail '�ֲ�����
Private mvar���մ���ID As Long '�ֲ�����
Private mvar������Ŀ�� As Boolean '�ֲ�����
Private mvar���ձ��� As String '�ֲ�����
Private mvarժҪ As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarԭʼ���� As Long '�ֲ�����
Private mvarԭʼִ�в���ID As Long '�ֲ�����
Private mvarӤ���� As Long '�ֲ�����
Private mvarTag As Variant '�ֲ�����
Public Property Let Tag(ByVal vData As Variant)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Tag = 5
    mvarTag = vData
End Property
Public Property Set Tag(ByVal vData As Variant)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Tag = Form1
    Set mvarTag = vData
End Property
Public Property Get Tag() As Variant
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Tag
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property


Public Property Let Ӥ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭʼִ�в���ID = 5
    mvarӤ���� = vData
End Property


Public Property Get Ӥ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭʼִ�в���ID
    Ӥ���� = mvarӤ����
End Property



Public Property Let ԭʼִ�в���ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭʼִ�в���ID = 5
    mvarԭʼִ�в���ID = vData
End Property


Public Property Get ԭʼִ�в���ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭʼִ�в���ID
    ԭʼִ�в���ID = mvarԭʼִ�в���ID
End Property



Public Property Let ԭʼ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭʼ���� = 5
    mvarԭʼ���� = vData
End Property


Public Property Get ԭʼ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭʼ����
    ԭʼ���� = mvarԭʼ����
End Property




Public Property Let ���ձ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ձ��� = 5
    mvar���ձ��� = vData
End Property

Public Property Get ���ձ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���ձ���
    ���ձ��� = mvar���ձ���
End Property

Public Property Let ������Ŀ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���￨�� = 5
    mvar������Ŀ�� = vData
End Property

Public Property Get ������Ŀ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���￨��
    ������Ŀ�� = mvar������Ŀ��
End Property

Public Property Let ���մ���ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���￨�� = 5
    mvar���մ���ID = vData
End Property

Public Property Get ���մ���ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���￨��
    ���մ���ID = mvar���մ���ID
End Property

Public Property Set Detail(ByVal vData As Detail)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Detail = Form1
    Set mvarDetail = vData
End Property

Public Property Get Detail() As Detail
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Detail
    Set Detail = mvarDetail
End Property

Public Property Set InComes(ByVal vData As BillInComes)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.InComes = Form1
    Set mvarInComes = vData
End Property

Public Property Get InComes() As BillInComes
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.InComes
    Set InComes = mvarInComes
End Property

Public Property Let ִ�в���ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ�в���ID = 5
    mvarִ�в���ID = vData
End Property

Public Property Get ִ�в���ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ִ�в���ID
    ִ�в���ID = mvarִ�в���ID
End Property

Public Property Let ���￨��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���￨�� = 5
    mvar���￨�� = vData
End Property

Public Property Get ���￨��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���￨��
    ���￨�� = mvar���￨��
End Property

Public Property Let ���ӱ�־(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ӱ�־ = 5
    mvar���ӱ�־ = vData
End Property

Public Property Get ���ӱ�־() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���ӱ�־
    ���ӱ�־ = mvar���ӱ�־
End Property

Public Property Let ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ��ҩ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ҩ���� = 5
    mvar��ҩ���� = vData
End Property

Public Property Get ��ҩ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��ҩ����
    ��ҩ���� = mvar��ҩ����
End Property

Public Property Let ���㵥λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���㵥λ = 5
    mvar���㵥λ = vData
End Property

Public Property Get ���㵥λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���㵥λ
    ���㵥λ = mvar���㵥λ
End Property

Public Property Let �շ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ���� = 5
    mvar�շ���� = vData
End Property

Public Property Get �շ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�շ����
    �շ���� = mvar�շ����
End Property

Public Property Let �ѱ�(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.�ѱ� = Form1
    mvar�ѱ� = vData
End Property

Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�ѱ�
    �ѱ� = mvar�ѱ�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let סԺ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ�� = 5
    mvarסԺ�� = vData
End Property

Public Property Get סԺ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.סԺ��
    סԺ�� = mvarסԺ��
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property

Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�Ա�
    �Ա� = mvar�Ա�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ��ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ҳID = 5
    mvar��ҳID = vData
End Property

Public Property Get ��ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��ҳID
    ��ҳID = mvar��ҳID
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��������
    �������� = mvar��������
End Property

Public Property Let ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��� = mvar���
End Property

Public Property Let �շ�ϸĿID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvar�շ�ϸĿID = vData
End Property

Public Property Get �շ�ϸĿID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�շ�ϸĿID
    �շ�ϸĿID = mvar�շ�ϸĿID
End Property

Private Sub Class_Initialize()
    Set Me.InComes = New BillInComes
    Set Me.Detail = New Detail
End Sub

Private Sub Class_Terminate()
    Set Me.InComes = Nothing
    Set Me.Detail = Nothing
End Sub

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar������ = vData
End Property

Public Property Get ������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ������ = mvar������
End Property

Public Property Let ҽ�Ƹ���(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.�ѱ� = Form1
    mvarҽ�Ƹ��� = vData
End Property

Public Property Get ҽ�Ƹ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�ѱ�
    ҽ�Ƹ��� = mvarҽ�Ƹ���
End Property

Public Property Let ժҪ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���￨�� = 5
    mvarժҪ = vData
End Property

Public Property Get ժҪ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���￨��
    ժҪ = mvarժҪ
End Property

Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��������
    �������� = mvar��������
End Property