VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   ControlBox      =   0   'False
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6810
      TabIndex        =   9
      Top             =   0
      Width           =   6810
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   120
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6810
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3810
      Width           =   6810
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   750
         MaxLength       =   6
         TabIndex        =   7
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5265
         TabIndex        =   5
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   4
         Top             =   105
         Width           =   1100
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   1
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4725
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":014A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2400
      ScaleHeight     =   1110
      ScaleWidth      =   2220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strKey As String
'入口参数
Private mstrTitle As String
Private mstrNote As String
Private mbytStyle As Byte
Private mstrSeek As String
Private mbln末级 As Boolean
'出口参数
Private rsSel As ADODB.Recordset

Public Function ShowSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    Optional ByVal strTitle As String, Optional bln末级 As Boolean, _
    Optional ByVal strSeek As String, Optional ByVal strNote As String, Optional ByVal blnMessage As Boolean = True, Optional ByVal blnOne As Boolean = False, Optional gcnConnect As ADODB.Connection) As ADODB.Recordset
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     strSQL=数据来源
'     strTitle=选择器类型命名
'     strNote=选择说明
'     bytStyle=选择器风格
'       为0时:ID,…
'       为1时:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'     bln末级=当bytStyle=1时,是否只能选择末级为1的项目
'     strSeek=缺省定位项,当bytStyle<>2时有效
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值

    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mbln末级 = bln末级
    mstrSeek = strSeek
    
    strKey = ""
    
    If strSQL <> "" Then
        On Error GoTo ErrH
        
        Set rsSel = New ADODB.Recordset
        rsSel.CursorLocation = adUseClient
        
        Screen.MousePointer = 11
        If Not frmParent Is Nothing Then
            frmParent.Refresh
        End If
        
        If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
        Call OpenRecordset_OtherBase(rsSel, mstrTitle & "选择", strSQL, gcnConnect)
        
        Screen.MousePointer = 0
        
        '没有数据则返回
        If rsSel.EOF Then
            If Not strSQL Like "*%*" Then
                '如果不是输入匹配(是全选择器)则提示
                If blnMessage = True Then
                    MsgBox "没有" & mstrTitle & "数据,请先初始化" & mstrTitle & "数据！", vbInformation, gstrSysName
                End If
            End If
            Unload Me: Exit Function
        End If
         
        '只有一行数据
        If rsSel.RecordCount = 1 Then
            If strSQL Like "*%*" Or blnOne = True Then
                '如果是输入匹配，就直接返回(否则让用户选择)
                Set ShowSelect = rsSel
                Unload Me: Exit Function
            End If
        End If
        
        '用户选择器
        Me.Show 1, frmParent
        
        Set ShowSelect = rsSel
        
        Unload Me
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub cmdCancel_Click()
    Set rsSel = Nothing '取消标志]
    Call SaveWinState(Me)
    Hide
End Sub

Private Sub cmdOK_Click()
    If rsSel.RecordCount <> 1 Then Exit Sub
    If mbln末级 And mbytStyle = 1 Then
        If rsSel!末级 <> 1 Then Exit Sub
    End If
    Call SaveWinState(Me)
    Hide
End Sub

Private Sub Form_Activate()
    If lvw.Visible Then
        lvw.SetFocus
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOK.Enabled Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim objNode As Node
    
    '缺省宽度
    If mbytStyle <> 2 Then Me.Width = 4500
    Call RestoreWinState(Me)
    
    If mstrTitle <> "" Then Me.Caption = mstrTitle & "选择"
    If mstrNote <> "" Then lblInfo.Caption = mstrNote
    
    '设置可见状态
    Select Case mbytStyle
        Case 0
            lvw.Visible = True
            tvw_s.Visible = False
            pic.Visible = False
        Case 1
            lvw.Visible = False
            tvw_s.Visible = True
            pic.Visible = False
            
            lblFind.Visible = False
            txtFind.Visible = False
        Case 2
            lvw.Visible = True
            tvw_s.Visible = True
            pic.Visible = True
    End Select
    
    '装入数据
    Select Case mbytStyle
        Case 0
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To rsSel.Fields.Count - 1
                If Not rsSel.Fields(i).Name Like "*ID" And rsSel.Fields(i).Name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & rsSel.Fields(i).Name, rsSel.Fields(i).Name
                End If
            Next
            'Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrTitle, "")
            
            lvw.ListItems.Clear
            Call FillList
        Case 1
            '所有树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            
            If Not rsSel.EOF Then
                For i = 1 To rsSel.RecordCount
                    If IsNull(rsSel!上级ID) Then
                        Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & rsSel!ID, IIf(IsNull(rsSel!编码), "", "[" & rsSel!编码 & "]") & rsSel!名称, 1)
                    Else
                        Set objNode = tvw_s.Nodes.Add("_" & rsSel!上级ID, 4, "_" & rsSel!ID, IIf(IsNull(rsSel!编码), "", "[" & rsSel!编码 & "]") & rsSel!名称, 1)
                    End If
                    If objNode.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then
                        objNode.Selected = True
                        objNode.Parent.Expanded = True
                    End If
                    rsSel.MoveNext
                Next
                If tvw_s.SelectedItem.Index = 1 Then tvw_s.Nodes(1).Child.Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Case 2
            '非末级树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            
            If Not rsSel.EOF Then
                rsSel.Filter = "末级=0"
                For i = 1 To rsSel.RecordCount
                    If IsNull(rsSel!上级ID) Then
                        Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & rsSel!ID, IIf(IsNull(rsSel!编码), "", "[" & rsSel!编码 & "]") & rsSel!名称, 1)
                    Else
                        Set objNode = tvw_s.Nodes.Add("_" & rsSel!上级ID, 4, "_" & rsSel!ID, IIf(IsNull(rsSel!编码), "", "[" & rsSel!编码 & "]") & rsSel!名称, 1)
                    End If
                    rsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To rsSel.Fields.Count - 1
                If Not rsSel.Fields(i).Name Like "*ID" And rsSel.Fields(i).Name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & rsSel.Fields(i).Name, rsSel.Fields(i).Name
                End If
            Next
            'Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrTitle, "")
            
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = picInfo.Height
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height
        Case 1
            tvw_s.Top = picInfo.Height
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = picInfo.Height
            tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height
            
            pic.Top = tvw_s.Top
            pic.Left = tvw_s.Width
            pic.Height = tvw_s.Height
            
            lvw.Top = tvw_s.Top
            lvw.Left = tvw_s.Width + pic.Width
            lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
            lvw.Height = tvw_s.Height
    End Select
    
    picBack.Left = lvw.Left
    picBack.Top = lvw.Top
    picBack.Width = lvw.Width
    picBack.Height = lvw.Height
    
    'If Me.ScaleWidth - cmdCancel.Width * 1.3 >= cmdHelp.Left + cmdHelp.Width * 2 + 300 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.1
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me)
End Sub

Private Sub lvw_DblClick()
    If cmdOK.Enabled And Not lvw.SelectedItem Is Nothing Then cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strFilter As String
    
    If rsSel.Fields("ID").Type = adVarChar Then
        strFilter = "ID='" & Mid(Item.Key, 2) & "'"
    Else
        strFilter = "ID=" & Mid(Item.Key, 2)
    End If
    If mbytStyle = 2 Then strFilter = strFilter & " And 末级=1"
    
    rsSel.Filter = strFilter
    cmdOK.Enabled = (rsSel.RecordCount = 1)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or lvw.Width - x < 1000 Then Exit Sub
        pic.Left = pic.Left + x
        tvw_s.Width = tvw_s.Width + x
        lvw.Left = lvw.Left + x
        lvw.Width = lvw.Width - x
        picBack.Left = picBack.Left + x
        picBack.Width = picBack.Width - x
        Me.Refresh
    End If
End Sub

Private Sub FillList()
'功能：装入ListView数据
    Dim i As Integer, j As Integer
    Dim objItem As ListItem
        
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To rsSel.RecordCount
        For j = 0 To rsSel.Fields.Count - 1
            If Not rsSel.Fields(j).Name Like "*ID" And rsSel.Fields(j).Name <> "末级" Then
                If lvw.ColumnHeaders("_" & rsSel.Fields(j).Name).Index = 1 Then
                    Set objItem = lvw.ListItems.Add(, "_" & rsSel!ID, IIf(IsNull(rsSel.Fields(j).Value), "", rsSel.Fields(j).Value), , 1)
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & rsSel.Fields(j).Name).Index - 1) = IIf(IsNull(rsSel.Fields(j).Value), "", rsSel.Fields(j).Value)
                End If
            End If
        Next
        rsSel.MoveNext
    Next
    
    Call zlControl.LvwSetColWidth(lvw)
    
    If Not lvw.SelectedItem Is Nothing Then
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub



Private Sub tvw_s_DblClick()
    If cmdOK.Enabled And mbytStyle = 1 Then cmdOK_Click
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strKeys As String, i As Integer
    Dim strFilter As String
    
    If strKey = Node.Key Then Exit Sub
    strKey = Node.Key
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If rsSel.Fields("ID").Type = adVarChar Then
                rsSel.Filter = "ID='" & Mid(Node.Key, 2) & "'"
            Else
                rsSel.Filter = "ID=" & Mid(Node.Key, 2)
            End If
            If mbln末级 Then
                cmdOK.Enabled = (rsSel!末级 = 1)
            Else
                cmdOK.Enabled = True
            End If
        Else
            cmdOK.Enabled = False
        End If
    ElseIf mbytStyle = 2 Then
        lvw.ListItems.Clear
        If Node.Key = "Root" Then
            rsSel.Filter = "末级=1"
            If Visible Then lvw.SetFocus
        Else
            strKeys = GetSubTree(Node)
            For i = 0 To UBound(Split(strKeys, ","))
                If rsSel.Fields("上级ID").Type = adVarChar Then
                    strFilter = strFilter & " Or (末级=1 And 上级ID='" & Split(strKeys, ",")(i) & "')"
                Else
                    strFilter = strFilter & " Or (末级=1 And 上级ID=" & Split(strKeys, ",")(i) & ")"
                End If
            Next
            strFilter = Mid(strFilter, 5)
            rsSel.Filter = strFilter
            
'            If rsSel.Fields("上级ID").Type = adVarChar Then
'                rsSel.Filter = "末级=1 And 上级ID='" & Mid(Node.Key, 2) & "'"
'            Else
'                rsSel.Filter = "末级=1 And 上级ID=" & Mid(Node.Key, 2)
'            End If
        End If
        If Not rsSel.EOF Then Call FillList
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'功能：返回一个结点的子树结点的Key(含该结点)
    Dim strKeys As String
    Dim objTmp As Node
    
    strKeys = "," & Mid(objNode.Key, 2) & strKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.children > 0 Then
            strKeys = "," & GetSubTree(objTmp) & strKeys
        Else
            strKeys = "," & Mid(objTmp.Key, 2) & strKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(strKeys, 2)
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub txtFind_Change()
'功能：根据用户输入的内容查找匹配的内容
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    Set lst = lvw.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '非文本不能做到部分匹配
        lngSubItems = lvw.ColumnHeaders.Count - 1
        For Each lst In lvw.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub


