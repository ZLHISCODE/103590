VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarKey As String           '关键字，用于选择后删除该对象！
Private mvar类型 As Byte            '0-文本,1-线条,2,折线,3-矩形,4-多边形,5-圆(椭圆), 6-顺序编号
Private mvar内容 As String          '文本的内容
Private mvar点集 As String          '折线或矩形的点集:2,3;34,25;...
Private mvarX1 As Long              '左上角坐标
Private mvarY1 As Long              '
Private mvarX2 As Long              '右下角坐标
Private mvarY2 As Long              '
Private mvar填充色 As OLE_COLOR     '
Private mvar填充方式 As Integer     '-1-不填充,-2-实心,0-水平线,1-垂直线,2-左斜线,3-右斜线,4-水平和垂直线,5-交叉线
Private mvar线条色 As OLE_COLOR     '
Private mvar字体色 As OLE_COLOR     '
Private mvar线型 As Byte            '0-实线,1-划线,2-点线,3-点划线,4-双点划线
Private mvar线宽 As Byte            '
Private mvar字体 As New StdFont     '字体,字号,字色,0000；四位分别表示粗体,斜体,下线,删除线
Private mvarSelected As Boolean     '是否选中，用于平移该对象！

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Selected(ByVal vData As Boolean)
    mvarSelected = vData
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Set 字体(ByVal vData As StdFont)
    Set mvar字体 = vData
End Property

Public Property Let 字体(ByVal vData As StdFont)
    Set mvar字体 = vData
End Property

Public Property Get 字体() As StdFont
    Set 字体 = mvar字体
End Property

Public Property Let 线宽(ByVal vData As Byte)
    mvar线宽 = vData
End Property

Public Property Get 线宽() As Byte
    线宽 = mvar线宽
End Property

Public Property Let 线型(ByVal vData As Byte)
    mvar线型 = vData
End Property

Public Property Get 线型() As Byte
    线型 = mvar线型
End Property

Public Property Let 线条色(ByVal vData As Long)
    mvar线条色 = vData
End Property

Public Property Get 线条色() As Long
    线条色 = mvar线条色
End Property

Public Property Let 字体色(ByVal vData As Long)
    mvar字体色 = vData
End Property

Public Property Get 字体色() As Long
    字体色 = mvar字体色
End Property

Public Property Let 填充方式(ByVal vData As Integer)
    mvar填充方式 = vData
End Property

Public Property Get 填充方式() As Integer
    填充方式 = mvar填充方式
End Property

Public Property Let 填充色(ByVal vData As Long)
    mvar填充色 = vData
End Property

Public Property Get 填充色() As Long
    填充色 = mvar填充色
End Property

Public Property Let Y2(ByVal vData As Long)
    mvarY2 = vData
End Property

Public Property Get Y2() As Long
    Y2 = mvarY2
End Property

Public Property Let X2(ByVal vData As Long)
    mvarX2 = vData
End Property

Public Property Get X2() As Long
    X2 = mvarX2
End Property

Public Property Let Y1(ByVal vData As Long)
    mvarY1 = vData
End Property

Public Property Get Y1() As Long
    Y1 = mvarY1
End Property

Public Property Let X1(ByVal vData As Long)
    mvarX1 = vData
End Property

Public Property Get X1() As Long
    X1 = mvarX1
End Property

Public Property Let 点集(ByVal vData As String)
    mvar点集 = vData
End Property

Public Property Get 点集() As String
    点集 = mvar点集
End Property

Public Property Let 内容(ByVal vData As String)
    '替换特殊字符，用于存储数据字符串的连接符
    mvar内容 = Replace(vData, "|", "↑")
End Property

Public Property Get 内容() As String
    内容 = mvar内容
End Property

Public Property Let 类型(ByVal vData As Byte)
    mvar类型 = vData
End Property

Public Property Get 类型() As Byte
    类型 = mvar类型
End Property

'################################################################################################################
'## 功能：  复制表格对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cPicMark
    Dim cNode As New cPicMark
    If blnWithKeys Then cNode.Key = Me.Key
    cNode.Selected = Me.Selected
    cNode.X1 = Me.X1
    cNode.Y1 = Me.Y1
    cNode.X2 = Me.X2
    cNode.Y2 = Me.Y2
    cNode.点集 = Me.点集
    cNode.类型 = Me.类型
    cNode.内容 = Me.内容
    cNode.填充方式 = Me.填充方式
    cNode.填充色 = Me.填充色
    cNode.线宽 = Me.线宽
    cNode.线条色 = Me.线条色
    cNode.线型 = Me.线型
    cNode.字体 = Me.字体
    cNode.字体色 = Me.字体色
    Set Clone = cNode
End Function

'################################################################################################################
'## 功能：  获取标记的属性字符串
'################################################################################################################
Public Property Get 对象属性() As String
    Dim strR As String
    strR = strR & IIf(Me.Selected, "1", "0") & "|"
    strR = strR & Me.类型 & "|"
    strR = strR & IIf(Me.内容 = "", " ", Me.内容) & "|"
    strR = strR & IIf(Me.点集 = "", " ", Me.点集) & "|"
    strR = strR & Me.X1 & "|"
    strR = strR & Me.Y1 & "|"
    strR = strR & Me.X2 & "|"
    strR = strR & Me.Y2 & "|"
    strR = strR & Me.填充色 & "|"
    strR = strR & Me.填充方式 & "|"
    strR = strR & Me.线条色 & "|"
    strR = strR & Me.字体色 & "|"
    strR = strR & Me.线型 & "|"
    strR = strR & Me.线宽 & "|"
    strR = strR & IIf(Me.字体 = "", " ", Me.字体)
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置标记内容
'################################################################################################################
Public Property Let 对象属性(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    If UBound(T) > 0 Then
        Me.Selected = IIf(T(0) = 1, True, False)
        Me.类型 = T(1)
        Me.内容 = Trim(T(2))
        Me.点集 = Trim(T(3))
        Me.X1 = T(4)
        Me.Y1 = T(5)
        Me.X2 = T(6)
        Me.Y2 = T(7)
        Me.填充色 = T(8)
        Me.填充方式 = T(9)
        Me.线条色 = T(10)
        Me.字体色 = T(11)
        Me.线型 = T(12)
        Me.线宽 = T(13)
        Me.字体.Name = Trim(T(14))
    End If
End Property

Private Sub Class_Initialize()
    Set mvar字体 = New StdFont
End Sub

Private Sub Class_Terminate()
    Set mvar字体 = Nothing
End Sub
