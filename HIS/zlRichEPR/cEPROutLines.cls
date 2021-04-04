VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPROutLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cCPROutLine"
Attribute VB_Ext_KEY = "Member0" ,"cCPROutLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlng流水号 As Long

Public Function LastOutline() As cCPROutLine
    Set LastOutline = mCol(mCol.Count)
End Function

Public Function Add() As cCPROutLine
    '创建新对象
    Dim objNewMember As cCPROutLine
    Dim sKey As String
    
    Set objNewMember = New cCPROutLine
    
    mlng流水号 = mlng流水号 + 1
    sKey = "K" & mlng流水号
    objNewMember.流水号 = mlng流水号
    
    mCol.Add objNewMember, sKey
    
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As cCPROutLine
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
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

Public Sub FillTree(ByVal objTree As Object)
    '将节点填充到树中
    Dim i As Long, j As Long, k As Long
    objTree.Nodes.Clear
    
    If mCol.Count = 0 Then
        Exit Sub
    End If
    LockWindowUpdate objTree.hWnd
    mCol(1).层次 = 1        '文章标题一定是标题1
    
    Dim Root As Node    '根节点
    Dim NodeOutline(1 To 9) As Node    '子节点
    Dim Node As Node
    
    Set Root = objTree.Nodes.Add(Text:="文件提纲", Image:=1)
    Root.Bold = True
    Root.Expanded = True
    
    For k = 1 To mCol.Count
        For i = 1 To mCol.Count
            If mCol(i).流水号 = mCol(k).流水号 Then
                If mCol(i).层次 > j + 1 Then mCol(i).层次 = j + 1
                Select Case mCol(i).层次
                Case 1
                    Set Node = objTree.Nodes.Add(Root, tvwChild, "K" & mCol(i).流水号, IIf(Len(mCol(i).名称) > 18, Left(mCol(i).名称, 18) & "...", mCol(i).名称), 2)
                Case Else
                    Set Node = objTree.Nodes.Add(NodeOutline(mCol(i).层次 - 1), tvwChild, "K" & mCol(i).流水号, IIf(Len(mCol(i).名称) > 18, Left(mCol(i).名称, 18) & "...", mCol(i).名称), 2)
                End Select
                Set NodeOutline(mCol(i).层次) = Node
                j = mCol(i).层次
                Node.Expanded = True
                Exit For
            End If
        Next i
    Next k
    
    LockWindowUpdate 0
End Sub







