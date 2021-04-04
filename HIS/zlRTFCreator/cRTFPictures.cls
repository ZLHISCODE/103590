VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFPictures"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFPicture"
Attribute VB_Ext_KEY = "Member0" ,"cRTFPicture"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function LastPicture() As cRTFPicture
Attribute LastPicture.VB_Description = "获取最后添加的图片对象。"
    Set LastPicture = mCol(mCol.Count)
End Function

Public Function Add(FileName As String, _
Width As Long, _
Height As Long, _
Optional Alignment As AlignTypeEnum = alm0_默认对齐, _
Optional sKey As String, _
Optional ScaleX As Long = 100, _
Optional ScaleY As Long = 100, _
Optional CropLeft As Long = 0, _
Optional CropRight As Long = 0, _
Optional CropTop As Long = 0, _
Optional CropBottom As Long = 0, _
Optional WidthOld As Long = 100, _
Optional HeightOld As Long = 100, _
Optional PicType As PicTypeEnum = pic_JPG, _
Optional SingleLine As Boolean = True) As Long
Attribute Add.VB_Description = "添加一个图片元素。"
    '创建新对象
    Dim objNewMember As cRTFPicture
    Set objNewMember = New cRTFPicture

    '设置传入方法的属性
    objNewMember.FileName = FileName
    objNewMember.ScaleX = ScaleX
    objNewMember.ScaleY = ScaleY
    objNewMember.CropLeft = CropLeft
    objNewMember.CropRight = CropRight
    objNewMember.CropTop = CropTop
    objNewMember.CropBottom = CropBottom
    objNewMember.Width = Width
    objNewMember.Height = Height
    objNewMember.WidthOld = WidthOld
    objNewMember.HeightOld = HeightOld
    objNewMember.PicType = PicType
    objNewMember.SingleLine = SingleLine
    objNewMember.Alignment = Alignment

    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Add = mCol.Count
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFPicture
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
Attribute Remove.VB_Description = "删除一个图片元素对象。"
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

