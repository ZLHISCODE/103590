VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRole 
   Caption         =   "人员角色分配"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "frmPresRole.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   9000
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   5640
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4455
      ScaleWidth      =   45
      TabIndex        =   16
      Top             =   3480
      Width           =   50
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   7560
      TabIndex        =   15
      ToolTipText     =   "请输入完整的角色名进行查找！"
      Top             =   540
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfUnGrantedFuncs 
      Height          =   2055
      Left            =   5760
      TabIndex        =   13
      Top             =   5880
      Width           =   3135
      _cx             =   5530
      _cy             =   3625
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPresRole.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGrantedFuncs 
      Height          =   1815
      Left            =   5760
      TabIndex        =   12
      Top             =   3720
      Width           =   3135
      _cx             =   5530
      _cy             =   3201
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPresRole.frx":0053
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   5280
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRole.frx":009A
            Key             =   "Role"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7800
      TabIndex        =   8
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6600
      TabIndex        =   7
      Top             =   8040
      Width           =   1100
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   4245
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "所属部门"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "缺省"
         Object.Width           =   970
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3889
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "所属部门"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "查找(&F)"
      Height          =   195
      Left            =   6720
      TabIndex        =   14
      Top             =   593
      Width           =   705
   End
   Begin VB.Label lblGrantedFuncs 
      Caption         =   "已授权功能"
      Height          =   200
      Left            =   5760
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblUnGrantedFuncs 
      Caption         =   "未授权功能"
      Height          =   200
      Left            =   5760
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblModule 
      Caption         =   "模块清单"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      Caption         =   "授权内容(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3165
      Width           =   990
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      Caption         =   "人员角色(&R)"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   990
   End
End
Attribute VB_Name = "frmPresRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPersonID    As Long        '人员ID
Private mstrUser        As String      '人员对应用户
Private mblnOk          As Boolean
Private mlngSysIdx      As Long         '当前选择的系统Index
Private mlngRoleIdx     As String       '当前选择的角色Index
Private mlngModuleIdx   As Long         '当前选择的模块Index
Private mblnLoad        As Boolean
Private Enum FuncCols
    Col_序号 = 0
    Col_功能 = 1
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal lngPersonID As Long) As Boolean
'功能：入口函数
    mblnOk = False
    mlngPersonID = lngPersonID
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
End Function

Private Sub cboSystem_Click()
    If mblnLoad Then Exit Sub
    If mlngSysIdx = cboSystem.ListIndex Then Exit Sub
    mlngSysIdx = cboSystem.ListIndex
    mlngModuleIdx = -1
    Call FillModule
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strRole As String
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To lvwRole.ListItems.Count
        If lvwRole.ListItems(i).Checked = True Then
            strRole = strRole & "ZL_" & lvwRole.ListItems(i).Text & ","
        End If
    Next

    If strRole <> "" Then
        gstrSQL = "Grant " & Mid(strRole, 1, Len(strRole) - 1) & " to " & mstrUser
        gcnOracle.Execute gstrSQL, , adCmdText
        If Err <> 0 Then
            MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
            If lvwRole.Enabled = True Then lvwRole.SetFocus
            Exit Sub
        End If
    End If
    '需要收回权限
    For i = 1 To lvwRole.ListItems.Count
        If lvwRole.ListItems(i).Checked = False Then
            gstrSQL = "revoke ZL_" & lvwRole.ListItems(i).Text & " from " & mstrUser
            gcnOracle.Execute gstrSQL, , adCmdText
        End If
    Next
    Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Add('" & mstrUser & "')", Me.Caption)
    If Err.Number <> 0 Then Err.Clear
    Unload Me
    mblnOk = True
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Select 姓名, b.用户名" & vbNewLine & _
            "From 人员表 a, 上机人员表 b" & vbNewLine & _
            "Where Id = [1]" & vbNewLine & _
            "And a.Id = b.人员id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-人员姓名", mlngPersonID)
    If Not rsTmp.EOF Then
        lblName.Caption = "姓名： " & rsTmp!姓名
        mstrUser = rsTmp!用户名 & ""
    End If
    '初始化
    mlngModuleIdx = -1
    mlngSysIdx = -1
    mlngModuleIdx = -1
    mblnLoad = True
    lvwRole.Icons = ils32
    Call FillRoleAndSystem
    mblnLoad = False
    Call cboSystem_Click
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillRoleAndSystem()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    Dim lstTmp As ListItem
    
    On Error GoTo ErrH
    strSQL = "Select a.角色, Decode(d.角色, Null, 0, 1) As 已授权" & vbNewLine & _
            "From (Select Substr(Granted_Role, 4) 角色" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And Admin_Option = 'YES'" & vbNewLine & _
            "       And Grantee = User) a," & vbNewLine & _
            "     (Select Distinct Substr(b.Granted_Role, 4) 角色" & vbNewLine & _
            "       From Dba_Role_Privs b, 上机人员表 c" & vbNewLine & _
            "       Where b.Grantee = c.用户名" & vbNewLine & _
            "       And b.Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And c.人员id = [1]) d" & vbNewLine & _
            "Where a.角色 = d.角色(+)" & vbNewLine & _
            "Order By a.角色"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-用户角色", mlngPersonID)
    With lvwRole
        .ListItems.Clear
        For i = 1 To rsTmp.RecordCount
            Set lstTmp = .ListItems.Add(, "R" & Format(i, "00000"), rsTmp!角色, "Role")
            lstTmp.Checked = rsTmp!已授权 = 1
            rsTmp.MoveNext
        Next
        cmdOK.Enabled = .ListItems.Count > 0
        If .ListItems.Count > 0 Then
            .ListItems(1).Selected = True
            mlngRoleIdx = .SelectedItem.Index
        End If
        
    End With
    strSQL = "Select Distinct m.编号, m.名称, m.共享号, m.所有者, m.安装日期, m.正常安装, m.版本号" & vbNewLine & _
            "From (Select Distinct 角色, 系统 From Zlrolegrant Where 序号 >= 100) r, Zlsystems m," & vbNewLine & _
            "     (Select Granted_Role" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And Admin_Option = 'YES'" & vbNewLine & _
            "       And Grantee = User) n" & vbNewLine & _
            "Where r.系统 = m.编号" & vbNewLine & _
            "And r.角色 = n.Granted_Role" & vbNewLine & _
            "Order By m.编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-已安装系统")
    
    With cboSystem
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称 & " v" & rsTmp!版本号 & "（" & rsTmp!编号 & "）"
            .ItemData(cboSystem.NewIndex) = rsTmp!编号
            If rsTmp!所有者 = UCase(gstrUserName) And .ListIndex < 0 Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        '有两种系统是程序固定的
        If (zlRegTool And 2) = 2 Then .AddItem "自定义报表"
        .AddItem "基础工具"
        .AddItem "取数函数"
        .AddItem "基础编码"
        If .ListIndex < 0 Then .ListIndex = 0
        mlngSysIdx = .ListIndex
    End With
 
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    txtFind.Move Me.ScaleWidth - 200 - txtFind.Width, lblRole.Top
    lblFind.Move txtFind.Left - lblFind.Width - 50, txtFind.Top + 30
    lvwRole.Move 50, txtFind.Top + txtFind.Height + 30, Me.ScaleWidth - 200, Me.ScaleHeight / 4
    lblSystem.Move 50, lvwRole.Top + lvwRole.Height + 150
    cboSystem.Move lblSystem.Width + lblSystem.Left + 50, lblSystem.Top - 50
    lblModule.Move lblSystem.Left, cboSystem.Top + cboSystem.Height + 100
    
    lvwModule.Move 50, lblModule.Top + lblModule.Height, (Me.ScaleWidth / 3) * 2, Me.ScaleHeight - lblModule.Top - lblModule.Height - 100 - cmdOK.Height - 100
    picSplit.Move lvwModule.Left + lvwModule.Width, lvwModule.Top, 50, lvwModule.Height
    vsfGrantedFuncs.Move lvwModule.Left + lvwModule.Width + 50, lvwModule.Top, Me.ScaleWidth - lvwModule.Width - 200, lvwModule.Height / 2 - 200
    lblUnGrantedFuncs.Move vsfGrantedFuncs.Left + 100, vsfGrantedFuncs.Top + vsfGrantedFuncs.Height + 100
    vsfUnGrantedFuncs.Move vsfGrantedFuncs.Left, lblUnGrantedFuncs.Top + lblUnGrantedFuncs.Height + 50, vsfGrantedFuncs.Width, (lvwModule.Height / 2) - 150
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 200, vsfUnGrantedFuncs.Top + vsfUnGrantedFuncs.Height + 120
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 200, cmdCancel.Top
    lblGrantedFuncs.Move lblUnGrantedFuncs.Left, lblModule.Top
End Sub

Private Sub lvwModule_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngModule As Long, strRole As String, lngSys As Long
    Dim vsTmp As VSFlexGrid
    
    If mblnLoad Then Exit Sub
    If mlngModuleIdx = Item.Index Then Exit Sub
    mlngModuleIdx = Item.Index
    On Error GoTo ErrH
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = "ZL_" & lvwRole.SelectedItem.Text
    End If
    If strRole = "" Then Exit Sub
    If cboSystem.Text = "基础编码" Or cboSystem.Text = "取数函数" Then Exit Sub
    lngModule = Val(Mid(Item.Key, 2))
    lngSys = cboSystem.ItemData(cboSystem.ListIndex)
    
    strSQL = "Select a.功能, Decode(b.功能, Null, 0, 1) 授权" & vbNewLine & _
            "From (Select 功能" & vbNewLine & _
            "       From Zlprogfuncs" & vbNewLine & _
            "       Where Nvl(系统,0) = [1]" & vbNewLine & _
            "       And 序号 = [2]) a," & vbNewLine & _
            "     (Select 功能" & vbNewLine & _
            "       From Zlrolegrant" & vbNewLine & _
            "       Where Nvl(系统,0) = [1]" & vbNewLine & _
            "       And 序号 = [2]" & vbNewLine & _
            "       And 角色 = [3]) b" & vbNewLine & _
            "Where a.功能 = b.功能(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSys, lngModule, strRole)
    vsfGrantedFuncs.Rows = vsfGrantedFuncs.FixedRows
    vsfUnGrantedFuncs.Rows = vsfUnGrantedFuncs.FixedRows
    Do While Not rsTmp.EOF
        If rsTmp!授权 = 1 Then
            Set vsTmp = vsfGrantedFuncs
        Else
            Set vsTmp = vsfUnGrantedFuncs
        End If
        vsTmp.Rows = vsTmp.Rows + 1
        vsTmp.TextMatrix(vsTmp.Rows - 1, Col_序号) = vsTmp.Rows - 1
        vsTmp.TextMatrix(vsTmp.Rows - 1, Col_功能) = rsTmp!功能 & ""
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwRole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mblnLoad Then Exit Sub
    If mlngRoleIdx = Item.Index Then Exit Sub
    mlngRoleIdx = Item.Index
    mlngModuleIdx = -1
    Call FillModule
End Sub

Private Sub FillModule()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lst As ListItem
    Dim strRole As String, strPre As String
    
    On Error GoTo ErrH
    lvwModule.ListItems.Clear
    vsfGrantedFuncs.Rows = vsfGrantedFuncs.FixedRows
    vsfUnGrantedFuncs.Rows = vsfUnGrantedFuncs.FixedRows
    '更新列表项
    With lvwModule.ColumnHeaders
        .Clear
        If cboSystem.Text = "基础编码" Then
            .Add , , "编码表", "1200"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cboSystem.Text = "取数函数" Then
            .Add , , "函数名", "1200"
            .Add , , "中文名", "1500"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cboSystem.Text = "基础工具" Then
            .Add , , "序号", "600"
            .Add , , "标题", "1800"
            .Add , , "说明", "3000"
        Else
            .Add , , "序号", "600"
            .Add , , "标题", "1800"
            .Add , , "说明", "3000"
        End If
    End With
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = "ZL_" & lvwRole.SelectedItem.Text
    End If
    If strRole = "" Then Exit Sub
    If cboSystem.Text = "基础编码" Then '显示该角色能访问的基础表
        strSQL = "Select t.系统, t.表名, t.说明" & vbNewLine & _
                    "From (Select s.名称 || '（' || s.编号 || '）' As 系统, s.所有者, b.表名, b.说明" & vbNewLine & _
                    "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
                    "       Where b.系统 = s.编号) t, User_Tab_Privs r" & vbNewLine & _
                    "Where t.所有者 = r.Owner" & vbNewLine & _
                    "And t.表名 = r.Table_Name" & vbNewLine & _
                    "And r.Grantee =[1]" & vbNewLine & _
                    "And r.Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
                    "Group By t.系统, t.表名, t.说明" & vbNewLine & _
                    "Having Count(r.Privilege) = 4"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!表名)
            lst.SubItems(1) = rsTmp!系统
            lst.SubItems(2) = rsTmp!说明 & ""
            rsTmp.MoveNext
        Loop
    ElseIf cboSystem.Text = "取数函数" Then '显示该角色能访问的取数函数
        strSQL = "Select s.名称 || '（' || s.编号 || '）' As 系统, s.所有者, f.函数名, f.中文名, f.说明" & vbNewLine & _
                "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
                "Where f.系统 = s.编号" & vbNewLine & _
                "And s.所有者 = r.Owner" & vbNewLine & _
                "And Upper(f.函数名) = r.Table_Name" & vbNewLine & _
                "And r.Grantee =[1]" & vbNewLine & _
                "And r.Privilege = 'EXECUTE'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!函数名)
            lst.SubItems(1) = rsTmp!中文名
            lst.SubItems(2) = rsTmp!系统
            lst.SubItems(3) = rsTmp!说明 & ""
            rsTmp.MoveNext
        Loop
    Else
        If cboSystem.Text = "基础工具" Then '显示该角色能访问的基础工具
            strSQL = "Select p.序号, p.标题, p.说明, r.功能" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where r.系统 Is Null" & vbNewLine & _
                    "And p.序号 = r.序号" & vbNewLine & _
                    "And r.角色 =[1]" & vbNewLine & _
                    "And p.系统 Is Null" & vbNewLine & _
                    "And p.序号 < 100" & vbNewLine & _
                    "Order By p.序号"
            
        Else '显示该角色能访问的模块
            strSQL = "Select p.序号, p.标题, p.说明, r.功能" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where Nvl(r.系统, 0) = Nvl(p.系统, 0)" & vbNewLine & _
                    "And p.序号 = r.序号" & vbNewLine & _
                    "And p.序号 >= 100" & vbNewLine & _
                    "And r.角色 = [1] And " & vbNewLine & _
                    IIF(cboSystem.Text = "自定义报表", " P.系统 is null", " (P.系统=" & cboSystem.ItemData(cboSystem.ListIndex) & " OR P.序号 Between 10000 And 19999)") & vbNewLine & _
                    "Order By p.序号"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            If strPre <> rsTmp!序号 & "" Then
                Set lst = lvwModule.ListItems.Add(, "K" & rsTmp!序号, rsTmp!序号)
                lst.SubItems(1) = rsTmp!标题
                lst.SubItems(2) = rsTmp!说明 & ""
                strPre = rsTmp!序号
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then  '鼠标左键
        If lvwModule.Width + X > 300 And picSplit.Left + X < Me.Width - 300 Then
            lvwModule.Move lvwModule.Left, lvwModule.Top, lvwModule.Width + X
            picSplit.Left = picSplit.Left + X
            lblGrantedFuncs.Left = picSplit.Left + picSplit.Width + 100
            vsfGrantedFuncs.Left = picSplit.Left + picSplit.Width
            vsfGrantedFuncs.Width = vsfGrantedFuncs.Width - X
            lblUnGrantedFuncs.Left = lblGrantedFuncs.Left
            vsfUnGrantedFuncs.Left = picSplit.Left + picSplit.Width
            vsfUnGrantedFuncs.Width = vsfUnGrantedFuncs.Width - X
        End If
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim litem As ListItem
    Dim lsItem As ListItem
    
    If KeyAscii = vbKeyReturn Then
        For Each litem In lvwRole.ListItems
            If litem.Text = txtFind.Text Or litem.Text = UCase(txtFind.Text) Then
                Set lsItem = litem
            Else
                litem.Selected = False
            End If
        Next
        If Not lsItem Is Nothing Then
            lsItem.Selected = True
            txtFind.SetFocus
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
        Else
            MsgBox "未查询到你想要的角色，请重新输入！", vbInformation, gstrSysName
            txtFind.SetFocus
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
        End If
    End If
End Sub
