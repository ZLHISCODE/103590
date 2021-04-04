VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasicWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'保持属性值的局部变量
Private mvarCodes As String '局部复制
Public Property Let Codes(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Codes = 5
    mvarCodes = vData
End Property


Public Property Get Codes() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Codes
    Codes = mvarCodes
End Property



