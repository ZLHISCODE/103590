VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'################################################################################################################
'## 内部变量
'################################################################################################################

Private mvarKey As String               '关键字
Private mvarRow As Long                 '行
Private mvarStartY As Long              'Y坐标
Private mvarHeight As Long              '行高
Private mvarMultiLine As Boolean        '是否多行显示，默认为True
Private mvarAutoHeight As Boolean       '自动高度，默认为True
Private mvarMaxHeight As Long           '最大行高
Private mvarVisibled As Boolean         '是否可见
Private mvarTitle As String             '行标题
Private mvarTag As String               '扩展标记
Private mvarHAlignment As HAlignEnum    '水平对齐方式
Private mvarVAlignment As VAlignEnum    '垂直对齐方式
Private mvarTextFormat As String        '文本格式
Private mvarProtected As Boolean        '是否保护

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let StartY(ByVal vData As Long)
    mvarStartY = vData
End Property

Public Property Get StartY() As Long
    StartY = mvarStartY
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let MultiLine(ByVal vData As Boolean)
    mvarMultiLine = vData
End Property

Public Property Get MultiLine() As Boolean
    MultiLine = mvarMultiLine
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let MaxHeight(ByVal vData As Long)
    mvarMaxHeight = vData
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = mvarMaxHeight
End Property

Public Property Let Visibled(ByVal vData As Boolean)
    mvarVisibled = vData
End Property

Public Property Get Visibled() As Boolean
    Visibled = mvarVisibled
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let TextFormat(ByVal vData As String)
    mvarTextFormat = vData
End Property

Public Property Get TextFormat() As String
    TextFormat = mvarTextFormat
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

'################################################################################################################
'## 功能：  复制对象
'##
'## 参数：  blnWithKey  :是否连同Key值一起复制！
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cRow
    Dim cNode As New cRow
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Row = Me.Row
    cNode.StartY = Me.StartY
    cNode.Height = Me.Height
    cNode.MultiLine = Me.MultiLine
    cNode.AutoHeight = Me.AutoHeight
    cNode.MaxHeight = Me.MaxHeight
    cNode.Visibled = Me.Visibled
    cNode.Title = Me.Title
    cNode.Tag = Me.Tag
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.TextFormat = Me.TextFormat
    cNode.Protected = Me.Protected
    
    Set Clone = cNode
End Function

Private Sub Class_Initialize()
    '默认属性初始化
    Me.Height = 300             '默认行高300
    Me.MultiLine = True         '默认多行显示
    Me.AutoHeight = True        '默认自动行高
    Me.MaxHeight = 999999       '默认最大行高999999
    Me.Visibled = True          '默认显示
    Me.Title = ""               '默认行标题为空
    Me.Tag = ""                 '默认行标记为空
    Me.HAlignment = HALignLeft  '默认靠左
    Me.VAlignment = VALignTop   '默认靠上
    Me.TextFormat = ""          '默认格式为空
    Me.Protected = False        '默认不保护
End Sub



















