VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScShowCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Name As String
Public Icon As String   '列图标
Public HiddenCol As Boolean '隐藏列
Public HiddenTitle As Boolean   '隐藏列标题
Public HiddenData As Boolean    '隐藏列数据显示
Public IsTotal As Boolean       '是否统计列
Public UseListLocate As Boolean '是否可列表定位
Public SortContrastCol As String    '排序对照列
Public DataConvert As String    '数据转换
Public IsNumerSort As Boolean    '数字排序


Private mobjRowRelation() As clsScRowRelation



Property Get RowRelationCount()
    RowRelationCount = UBound(mobjRowRelation)
End Property


Property Get RowRelation(ByVal lngIndex As Long) As clsScRowRelation
    Set RowRelation = mobjRowRelation(lngIndex)
End Property


Public Sub AddRowRelation(objRowRelation As clsScRowRelation)
'添加RowRelation
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjRowRelation) + 1
    ReDim Preserve mobjRowRelation(lngUbound)
    
    Set mobjRowRelation(lngUbound) = objRowRelation
End Sub

Private Sub Class_Initialize()
    ReDim mobjRowRelation(0)
End Sub

Private Sub Class_Terminate()
    ReDim mobjRowRelation(0)
    Set mobjRowRelation(0) = Nothing
End Sub
