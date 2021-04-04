VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutNurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"OutNurses"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar姓名 As String '局部复制

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
' X.姓名
    姓名 = mvar姓名
End Property



