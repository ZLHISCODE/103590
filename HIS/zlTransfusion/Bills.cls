VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Bill"
Attribute VB_Ext_KEY = "Member0" ,"Bill"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(病人ID As Long, 医嘱ID As Long, 医嘱内容 As String, 单量 As Currency, 单位 As String, 容量 As Long, 时间 As Long, 医生嘱托 As String, 序号 As Integer, 金额 As Currency, 明细计费状态 As Integer, Optional sKey As String) As Bill
    '创建新对象
    Dim objNewMember As Bill
    Set objNewMember = New Bill
    
    On Error GoTo errH
    '设置传入方法的属性
    objNewMember.病人ID = 病人ID
    objNewMember.医嘱ID = 医嘱ID
    objNewMember.医嘱内容 = DelInvalidChar(医嘱内容, "'")
    objNewMember.单量 = 单量
    objNewMember.单位 = 单位
    objNewMember.容量 = 容量
    objNewMember.时间 = 时间
    objNewMember.医生嘱托 = DelInvalidChar(医生嘱托, "'")
    objNewMember.序号 = 序号
    objNewMember.金额 = 金额
    objNewMember.明细计费状态 = 明细计费状态
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
errH:
    If Err.Number = 457 Then
        Exit Function
    ElseIf ErrCenter = 1 Then
        Resume
    End If

End Function

Public Property Get Item(vntIndexKey As Variant) As Bill
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
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

