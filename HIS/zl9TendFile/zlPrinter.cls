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

'局部变量保存属性值
Private mvarDeviceName As String '本地复本
Private mvarPaperSizes As String '本地复本
Private mvarPaperBins As String '本地复本
Private mvarPort As String '本地复本
Private mvarOrientations As String '本地复本
Private mvarCurrent As Boolean '本地复本

Public Property Let Current(ByVal vData As Boolean)
'当给属性赋值时在参数左边使用。
'Syntax: X.Current = 5
    mvarCurrent = vData
End Property

Public Property Get Current() As Boolean
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Current
    Current = mvarCurrent
End Property

Public Property Let Orientations(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.Orientation = 5
    mvarOrientations = vData
End Property

Public Property Get Orientations() As String
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Orientation
    Orientations = mvarOrientations
End Property



Public Property Let Port(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.Port = 5
    mvarPort = vData
End Property


Public Property Get Port() As String
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Port
    Port = mvarPort
End Property


Public Property Let PaperBins(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.PaperBin = 5
    mvarPaperBins = vData
End Property


Public Property Get PaperBins() As String
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.PaperBin
    PaperBins = mvarPaperBins
End Property

Public Property Let PaperSizes(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.PaperSize = 5
    mvarPaperSizes = vData
End Property

Public Property Get PaperSizes() As String
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.PaperSize
    PaperSizes = mvarPaperSizes
End Property


Public Property Let DeviceName(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.DeviceName = 5
    mvarDeviceName = vData
End Property

Public Property Get DeviceName() As String
Attribute DeviceName.VB_UserMemId = 0
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.DeviceName
    DeviceName = mvarDeviceName
End Property



