VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "һ��Ԫ�ؾ��Ǽ��ʵ��ϵ�һ���ؼ�"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarKey As String '�ֲ�����
Private mvarControl As Object '�ֲ�����
Private mvarTabIndex As Long '�ֲ�����
Private mvarValue As String '�ֲ�����
Private mvarVisible As Boolean '�ֲ�����

Public Property Let Visible(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Visible = 5
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property

Public Property Let Value(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Value = 5
    mvarValue = vData
End Property

Public Property Get Value() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property

Public Property Let TabIndex(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.TabIndex = 5
    mvarTabIndex = vData
End Property

Public Property Get TabIndex() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.TabIndex
    TabIndex = mvarTabIndex
End Property

Public Property Set Control(ByVal vData As Object)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Control = Form1
    Set mvarControl = vData
End Property


Public Property Get Control() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Control
    Set Control = mvarControl
End Property

Public Property Let Key(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property