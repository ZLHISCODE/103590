VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlTFPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'�ֲ�������������ֵ
Private mvarDeviceName As String '���ظ���
Private mvarPaperSizes As String '���ظ���
Private mvarPaperBins As String '���ظ���
Private mvarPort As String '���ظ���
Private mvarOrientations As String '���ظ���
Private mvarCurrent As Boolean '���ظ���

Public Property Let Current(ByVal vData As Boolean)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Current = 5
    mvarCurrent = vData
End Property

Public Property Get Current() As Boolean
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Current
    Current = mvarCurrent
End Property

Public Property Let Orientations(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Orientation = 5
    mvarOrientations = vData
End Property

Public Property Get Orientations() As String
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Orientation
    Orientations = mvarOrientations
End Property



Public Property Let Port(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Port = 5
    mvarPort = vData
End Property


Public Property Get Port() As String
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Port
    Port = mvarPort
End Property


Public Property Let PaperBins(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.PaperBin = 5
    mvarPaperBins = vData
End Property


Public Property Get PaperBins() As String
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.PaperBin
    PaperBins = mvarPaperBins
End Property

Public Property Let PaperSizes(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.PaperSize = 5
    mvarPaperSizes = vData
End Property

Public Property Get PaperSizes() As String
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.PaperSize
    PaperSizes = mvarPaperSizes
End Property


Public Property Let DeviceName(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.DeviceName = 5
    mvarDeviceName = vData
End Property

Public Property Get DeviceName() As String
Attribute DeviceName.VB_UserMemId = 0
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.DeviceName
    DeviceName = mvarDeviceName
End Property


