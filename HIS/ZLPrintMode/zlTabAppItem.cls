VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlTabAppItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarText As String '本地复本
Public Property Let Text(ByVal vData As String)
'当给属性赋值时在参数左边使用。
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



