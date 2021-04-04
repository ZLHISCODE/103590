VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvarDataSet As ADODB.Recordset '局部复制
Private mvarDataName As String '局部复制

Public Property Let DataName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.DataName = 5
    mvarDataName = vData
End Property

Public Property Get DataName() As String
Attribute DataName.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.DataName
    DataName = mvarDataName
End Property

Public Property Set DataSet(ByVal vData As ADODB.Recordset)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.DataSet = Form1
    Set mvarDataSet = vData
End Property

Public Property Get DataSet() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.DataSet
    Set DataSet = mvarDataSet
End Property



