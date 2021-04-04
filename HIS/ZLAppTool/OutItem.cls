VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCaption As String '局部复制
Private mvarIcon As String '局部复制
Private mvarGroupName As String '局部复制
Private mvarKey As String '局部复制
Private mvarTag As String '局部复制

Public Property Let Tag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Tag = 5
    mvarTag = vData
End Property

Public Property Get Tag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property

Public Property Let Key(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property

Public Property Let GroupName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.GroupName = 5
    mvarGroupName = vData
End Property

Public Property Get GroupName() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.GroupName
    GroupName = mvarGroupName
End Property

Public Property Let Icon(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property

Public Property Get Icon() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property

Public Property Let Caption(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property
