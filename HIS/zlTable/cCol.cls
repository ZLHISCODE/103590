VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCol"
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
Private mvarCol As Long                 '列
Private mvarStartX As Long              'X坐标
Private mvarWidth As Long               '列宽
Private mvarFixedWidth As Boolean       '是否锁定宽度，默认为False
Private mvarMaxWidth As Long            '最大列宽
Private mvarVisibled As Boolean         '是否可见
Private mvarTitle As String             '列标题
Private mvarTag As String               '扩展标记
Private mvarHAlignment As HAlignEnum    '水平对齐方式
Private mvarVAlignment As VAlignEnum    '垂直对齐方式
Private mvarTextFormat As String        '文本格式
Private mvarIndent As Long              '缩进值
Private mvarProtected As Boolean        '是否保护

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let StartX(ByVal vData As Long)
    mvarStartX = vData
End Property

Public Property Get StartX() As Long
    StartX = mvarStartX
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let MaxWidth(ByVal vData As Long)
    mvarMaxWidth = vData
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = mvarMaxWidth
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

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
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
Public Function Clone(Optional blnWithKey As Boolean = False) As cCol
    Dim cNode As New cCol
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Col = Me.Col
    cNode.StartX = Me.StartX
    cNode.Width = Me.Width
    cNode.FixedWidth = Me.FixedWidth
    cNode.MaxWidth = Me.MaxWidth
    cNode.Visibled = Me.Visibled
    cNode.Title = Me.Title
    cNode.Tag = Me.Tag
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.TextFormat = Me.TextFormat
    cNode.Indent = Me.Indent
    cNode.Protected = Me.Protected
    
    Set Clone = cNode
End Function

Private Sub Class_Initialize()
    '默认属性初始化
    Me.Width = 1600             '默认列宽1600
    Me.FixedWidth = False       '默认可以调整列宽
    Me.MaxWidth = 999999        '默认最大宽度不超过999999
    Me.Visibled = True          '默认显示
    Me.Title = ""               '默认列标题为空
    Me.Tag = ""                 '默认列标记为空
    Me.HAlignment = HALignLeft  '默认靠左
    Me.VAlignment = VALignTop   '默认靠上
    Me.TextFormat = ""          '默认格式为空
    Me.Indent = 0               '默认缩进为0
    Me.Protected = False        '默认不保护
End Sub

















