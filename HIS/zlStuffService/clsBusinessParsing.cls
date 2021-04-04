VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessParsing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'保持属性值的局部变量
Private mvaobjReturnData As clsBusinessDatas     '局部复制
Private mvarobjReturnRecord As clsBusinessDatas '局部复制
 
Public Property Set objReturnRecord(ByVal vData As clsBusinessDatas)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
   Set mvarobjReturnRecord = vData
End Property

Public Property Get objReturnRecord() As clsBusinessDatas
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
   Set objReturnRecord = mvarobjReturnRecord
End Property
 


Public Property Set objReturnData(ByVal vData As clsBusinessDatas)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.objReturnData = 5
    Set mvaobjReturnData = vData
End Property
Public Property Get objReturnData() As clsBusinessDatas
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objReturnData
    Set objReturnData = mvaobjReturnData
End Property




