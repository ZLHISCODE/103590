VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'如 id=[1] * {%0#1}  {%0#1}=(select Num from t1 where 号码=([2]+[3]))
Public Parent As clsSqlItem
Public Context As String
Public BracketTag As String

Public HasLinkWhere As Boolean

Public HasPar As Boolean
Public HasSubPar As Boolean

Public HasSubSql As Boolean

Private maryParNo()  As Long

Private mobjSubSqls() As clsSqlStruct    '存储当前context中包含的查询, 如F1 between {%0#1} and {%0#2}，  {%0#1}和{%0#2}分别对应一个查询
Private mobjSubItems() As clsSqlItem '存储当前context中包含的非select结构，如 decode{%0#1} as 字段1，{%0#1}可以是函数或运算，如([1]+[2], 0, 'a', 1, 'b', 'c')

'pars
Property Get ParsCount() As Long
    ParsCount = UBound(maryParNo)
End Property

Property Get ParNo(ByVal lngIndex As Long) As Long
    ParNo = maryParNo(lngIndex)
End Property

Property Let ParNo(ByVal lngIndex As Long, ByVal lngValue As Long)
    maryParNo(lngIndex) = lngValue
End Property

Public Sub AddParNo(ByVal lngParNo As Long)
    Dim lngUbound As Long
    
    lngUbound = UBound(maryParNo) + 1
    
    ReDim Preserve maryParNo(lngUbound)
    maryParNo(lngUbound) = lngParNo
End Sub

Public Sub AddParNoAry(ByRef aryParNo() As Long)
    maryParNo = aryParNo
End Sub

Public Sub CopyParNoTo(ByRef aryTag() As Long)
    Call CopyLngArray(maryParNo, aryTag)
End Sub

'SubSqls
Property Get SubSqlsCount() As Long
    SubSqlsCount = UBound(mobjSubSqls)
End Property

Property Get SubSqls(ByVal lngIndex As Long) As clsSqlStruct
    Set SubSqls = mobjSubSqls(lngIndex)
End Property

Property Set SubSqls(ByVal lngIndex As Long, ByRef objSqlStruct As clsSqlStruct)
    Set mobjSubSqls(lngIndex) = objSqlStruct
End Property

Public Sub AddSubSql(ByRef objSubSql As clsSqlStruct)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjSubSqls) + 1
    
    ReDim Preserve mobjSubSqls(lngUbound)
    Set mobjSubSqls(lngUbound) = objSubSql
End Sub

Public Sub AddSubSqlAry(ByRef objSubSql() As clsSqlStruct)
    mobjSubSqls = objSubSql
End Sub

'subItems
Property Get SubItemsCount() As Long
    SubItemsCount = UBound(mobjSubItems)
End Property

Property Get SubItems(ByVal lngIndex As Long) As clsSqlItem
    Set SubItems = mobjSubItems(lngIndex)
End Property

Property Set SubItems(ByVal lngIndex As Long, ByRef objSqlItem As clsSqlItem)
    Set mobjSubItems(lngIndex) = objSqlItem
End Property

Public Sub AddSubItem(ByRef objSubItem As clsSqlItem)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjSubItems) + 1
    
    ReDim Preserve mobjSubItems(lngUbound)
    Set mobjSubItems(lngUbound) = objSubItem
End Sub

Public Sub AddSubItemAry(ByRef objSubItem() As clsSqlItem)
    mobjSubItems = objSubItem
End Sub






Private Sub Class_Initialize()
    ReDim maryParNo(0)
    ReDim mobjSubSqls(0)
    ReDim mobjSubItems(0)
End Sub
