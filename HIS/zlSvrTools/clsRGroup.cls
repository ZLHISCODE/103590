VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsRelas"
Attribute VB_Ext_KEY = "Member0" ,"clsRelas"
'局部变量，保存集合
Private mCol As Collection

Public Function Add(objRelas As clsRelas, Optional sKey As String) As clsRelas
    '创建新对象
    Dim objNewMember As clsRelas
    Set objNewMember = objRelas
    mCol.Add objNewMember, sKey
    '返回已创建的对象
    
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsRelas
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

Public Sub RelasCheck(ByVal strKey As String, ByVal blnChecked As Boolean)
    '检查权限之间的关系，并做出相应的处理
    
    Dim objRelas As clsRelas
    Dim int分组 As Integer, int关系 As Integer, int主功能 As Integer, bln主功能之一有选中 As Boolean
    int分组 = Me.Item(strKey).分组
    int关系 = Me.Item(strKey).关系
    int主功能 = Me.Item(strKey).主功能
    Me.Item(strKey).Checked = blnChecked
    
    If int关系 = 1 Then
        For Each objRelas In Me
            If objRelas.分组 = int分组 And objRelas.Key <> strKey Then
                If blnChecked Then
                    Me.Item(objRelas.Key).Checked = False
                End If
            End If
        Next
    ElseIf int关系 = 2 Then
        If int主功能 = 1 Then
            '本身是主功能, 选中的情况下，子功能不用做处理
            If Not blnChecked Then
                '未选中的情况下，检查其他主功能是否有选中，如果都未选中，则取消子功能
                bln主功能之一有选中 = False
                
                '先处理主项关系为OR的情况
                For Each objRelas In Me
                    If objRelas.分组 = int分组 And objRelas.主功能 = 1 Then
                        If objRelas.主功能关系 = 0 Then
                            If objRelas.Checked Then bln主功能之一有选中 = True
                        End If
                    End If
                Next
                
                '再处理主项关系为AND的情况
                For Each objRelas In Me
                    If objRelas.分组 = int分组 And objRelas.主功能 = 1 Then
                        If objRelas.主功能关系 = 1 Then
                            If objRelas.Checked And bln主功能之一有选中 Then
                                bln主功能之一有选中 = True
                            Else
                                bln主功能之一有选中 = False
                            End If
                        End If
                    End If
                Next
                
                If Not bln主功能之一有选中 Then
                    For Each objRelas In Me
                        If objRelas.分组 = int分组 And objRelas.主功能 = 0 Then
                            Me.Item(objRelas.Key).Checked = False
                        End If
                    Next
                End If
            End If
        Else
            '本身不是主功能
            If blnChecked Then
                '主功能选中，才能为选中
                bln主功能之一有选中 = False
                For Each objRelas In Me
                    If objRelas.分组 = int分组 And objRelas.主功能 = 1 Then
                        If objRelas.主功能关系 = 0 Then
                            If objRelas.Checked Then bln主功能之一有选中 = True
                        End If
                    End If
                Next
                
                '再处理主项关系为AND的情况
                For Each objRelas In Me
                    If objRelas.分组 = int分组 And objRelas.主功能 = 1 Then
                        If objRelas.主功能关系 = 1 Then
                            If objRelas.Checked And bln主功能之一有选中 Then
                                bln主功能之一有选中 = True
                            Else
                                bln主功能之一有选中 = False
                            End If
                        End If
                    End If
                Next
                
                If Not bln主功能之一有选中 Then
                    Me.Item(strKey).Checked = False
                End If
            End If
        End If
    Else
        For Each objRelas In Me
            If objRelas.分组 = int分组 And objRelas.Key <> strKey Then
                Me.Item(objRelas.Key).Checked = blnChecked
            End If
        Next
    End If

End Sub
