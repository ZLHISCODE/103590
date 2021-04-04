VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'不用菜单设计器，用API自已创建简单的弹出菜单。
'在工程中新增一个类，名称为cPopupMenu

Option Explicit
'

Public Event MenuClick(ByVal Key As Long, ByVal Caption As String)
Public Event MenuBeforeShow(ByRef Cancel As Boolean)
Public Event MenuAfterClick()

Private mblnCreated As Boolean                      '是否已经创建菜单
Private mlngMainMenu As Long                        '弹出菜单ID
Private mrsMenu As New ADODB.Recordset

Public Function Add(ByVal Key As Long, ByVal Caption As String, Optional ByVal ParentKey As Long = 0, Optional ByVal Style As Byte = 1, Optional ByVal Enabled As Boolean = True, Optional ByVal Index As Long = 0, Optional Checked As Boolean = False) As Boolean
    With mrsMenu
        .AddNew
        .Fields("标题").Value = Caption
        .Fields("类型").Value = Style
        .Fields("可用").Value = Enabled
        .Fields("复选").Value = Checked
        .Fields("序号").Value = Index
        .Fields("Key").Value = Key
        .Fields("上级Key").Value = ParentKey
    End With
End Function

Public Property Get ItemValue(vntIndexKey As Variant, vntField As String) As Variant
    '引用集合中的一个元素时使用。
    mrsMenu.Filter = ""
    mrsMenu.Filter = "Key=" & vntIndexKey
    ItemValue = mrsMenu.Fields(vntField).Value
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    mrsMenu.Filter = ""
    Count = mrsMenu.RecordCount
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    mrsMenu.Filter = ""
    mrsMenu.Filter = "Key=" & vntIndexKey
    mrsMenu.Delete adAffectCurrent
End Sub

Private Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    Dim lngLoop As Long
    Dim lngCol As Long
    
    '创建rsTo结构
    If rsTo.State = adStateOpen Then rsTo.Close
    With rsTo
        For lngCol = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields.Append rsFrom.Fields(lngCol).Name, rsFrom.Fields(lngCol).Type, rsFrom.Fields(lngCol).DefinedSize
        Next
        rsTo.Open
    End With
    
    '复制数据
    rsFrom.Filter = ""
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    For lngLoop = 1 To rsFrom.RecordCount
        rsTo.AddNew
        For lngCol = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngCol).Value = rsFrom.Fields(lngCol).Value
        Next
        rsFrom.MoveNext
    Next
End Sub

Private Function CreateSubMenu(Optional ByVal strUpKey As Long = 0) As Long
    Dim lngLoop As Long
    Dim lngMenuHand As Long
    Dim lngSubMenuHand As Long
    Dim rs As New ADODB.Recordset
    Dim lngFlags As Long
    Dim lngMenu As Long
    
    '复制记录集
    Call CopyRecord(mrsMenu, rs)
        
    rs.Filter = ""
    rs.Filter = "上级Key=" & strUpKey
    If rs.RecordCount = 0 Then Exit Function
    
    rs.MoveFirst
    
    lngMenuHand = CreatePopupMenu()
    If lngMenuHand <= 0 Then Exit Function
    
    For lngLoop = 1 To rs.RecordCount
        Select Case Val(rs.Fields("类型").Value)
        Case 2
            lngFlags = MF_SEPARATOR
        Case Else
            lngFlags = MF_STRING
            If rs.Fields("可用").Value = False Then lngFlags = lngFlags + MF_GRAYED
            If rs.Fields("复选").Value Then lngFlags = lngFlags + MF_CHECKED
            
            
        End Select
        
        lngMenu = AppendMenu(lngMenuHand, lngFlags, Val(rs.Fields("Key").Value), rs.Fields("标题").Value)
        
        '检查是否有子菜单
        lngSubMenuHand = 0
        lngSubMenuHand = CreateSubMenu(rs.Fields("Key").Value)
        
        '修改子菜单的所属关系
        If lngSubMenuHand > 0 Then Call ModifyMenu(lngMenuHand, lngLoop, lngFlags + MF_POPUP, lngSubMenuHand, CStr(rs.Fields("标题").Value))
                
        rs.MoveNext
    Next
    
    CreateSubMenu = lngMenuHand
End Function

Public Function CreatePopMenu() As Boolean
    '功能:创建弹出菜单
            
    If mlngMainMenu > 0 Then
        Call DestroyMenu(mlngMainMenu)
        mlngMainMenu = 0
        mblnCreated = False
    End If

    mlngMainMenu = CreateSubMenu
    
    mblnCreated = True
    
    CreatePopMenu = True
End Function

Public Sub ShowPopupMenuByCursor()
    '功能:在当前位置显示弹出菜单
    Dim p As POINTAPI
    
    Call GetCursorPos(p)
    Call ShowPopupMenu(p.X * 15, p.Y * 15)

End Sub

Public Sub ShowPopupMenu(ByVal X As Single, ByVal Y As Single)
    '功能:在指定位置显示弹出菜单
    
    Dim lngMenuPostion As Long
    Dim strMenuCaption As String
    Dim blnCancel As Boolean
    
    RaiseEvent MenuBeforeShow(blnCancel)
    If blnCancel Then Exit Sub
    
    If mblnCreated = False Then
        '如果没有创建,则先创建菜单
        Call CreatePopMenu
    End If
    
    If mlngMainMenu <= 0 Then Exit Sub
    
    lngMenuPostion = TrackPopupMenu(mlngMainMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, X / 15, Y / 15, 0, GetForegroundWindow(), 0)
    
    RaiseEvent MenuAfterClick
    
    If lngMenuPostion <= 0 Then Exit Sub
    
    strMenuCaption = Space(255)
    Call GetMenuString(mlngMainMenu, lngMenuPostion, strMenuCaption, Len(strMenuCaption), MF_BYCOMMAND)
    
    RaiseEvent MenuClick(lngMenuPostion, strMenuCaption)
    
End Sub

Private Sub Class_Initialize()
    With mrsMenu
        .Fields.Append "标题", adVarChar, 30
        .Fields.Append "类型", adTinyInt
        .Fields.Append "可用", adBoolean
        .Fields.Append "复选", adBoolean
        .Fields.Append "序号", adInteger
        .Fields.Append "Key", adBigInt
        .Fields.Append "上级Key", adBigInt
        .Open
    End With
End Sub

Private Sub Class_Terminate()
    If mlngMainMenu > 0 Then Call DestroyMenu(mlngMainMenu)
    
    If mrsMenu.State = adStateOpen Then mrsMenu.Close
    Set mrsMenu = Nothing
End Sub
