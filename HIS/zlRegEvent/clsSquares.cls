VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSquare"
Attribute VB_Ext_KEY = "Member0" ,"clsSquare"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'局部变量，保存集合
Private mCol As Collection
Public Function AddItem(NO As String, 结算性质 As Integer, 结算方式 As String, 结算金额 As Double, _
    结算号码 As String, 备注 As String, 允许修改 As Boolean, 误差费 As Double, _
    刷卡类别ID As Long, 刷卡卡号 As String, 交易流水号 As String, 交易说明 As String, _
    其他交易信息 As String, Optional sKey As String) As clsSquare
    '创建新对象
    Dim objNewMember As clsSquare
    Set objNewMember = New clsSquare
    
    '设置传入方法的属性
    objNewMember.NO = NO
    objNewMember.结算性质 = 结算性质
    objNewMember.结算方式 = 结算方式
    objNewMember.结算金额 = 结算金额
    objNewMember.结算性质 = 结算性质
    objNewMember.结算号码 = 结算号码
    objNewMember.备注 = 备注
    objNewMember.允许修改 = 允许修改
    objNewMember.误差费 = 误差费
    objNewMember.刷卡类别ID = 刷卡类别ID
    objNewMember.刷卡卡号 = 刷卡卡号
    objNewMember.交易流水号 = 交易流水号
    objNewMember.交易说明 = 交易说明
    objNewMember.其他交易信息 = 其他交易信息
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '返回已创建的对象
    Set AddItem = objNewMember
    Set objNewMember = Nothing
End Function
Public Function Add(objSquare As clsSquare, Optional sKey As String) As clsSquare
    If Len(sKey) = 0 Then
       mCol.Add objSquare
    Else
        mCol.Add objSquare, sKey
    End If
    Set Add = objSquare
End Function

Public Property Get Item(vntIndexKey As Variant) As clsSquare
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

