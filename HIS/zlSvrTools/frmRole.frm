VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRole 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "角色授权管理"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmRole.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   9735
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRolesReset 
      Caption         =   "重整所有角色"
      Height          =   350
      Left            =   7575
      TabIndex        =   24
      Top             =   3120
      Width           =   1875
   End
   Begin VB.CheckBox chkOnlyShowNOSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "只显示未分系统角色"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7575
      TabIndex        =   23
      Top             =   225
      Width           =   2040
   End
   Begin VB.CommandButton cmdSystemMove 
      Caption         =   "移到系统(&S)"
      Height          =   350
      Left            =   7575
      TabIndex        =   22
      Top             =   2565
      Width           =   1875
   End
   Begin VB.CommandButton cmdRoleMove 
      Caption         =   "移到分组(&M)"
      Height          =   350
      Left            =   7575
      TabIndex        =   21
      Top             =   2205
      Width           =   1875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除角色(&D)"
      Height          =   350
      Left            =   7575
      TabIndex        =   20
      Top             =   1560
      Width           =   1875
   End
   Begin VB.CommandButton cmdDeleteGroup 
      Caption         =   "删除组(&R)"
      Height          =   350
      Left            =   2325
      TabIndex        =   19
      Top             =   3900
      Width           =   1100
   End
   Begin VB.CommandButton cmdModifyGroup 
      Caption         =   "修改组(&E)"
      Height          =   350
      Left            =   1230
      TabIndex        =   18
      Top             =   3900
      Width           =   1100
   End
   Begin VB.PictureBox picHLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   2070
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5835
      TabIndex        =   16
      Top             =   4350
      Width           =   5835
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "新建组(&N)"
      Height          =   350
      Left            =   135
      TabIndex        =   5
      Top             =   3900
      Width           =   1100
   End
   Begin VB.CommandButton cmdGrantAll 
      Caption         =   "恢复所有角色及权限"
      Height          =   350
      Left            =   7575
      TabIndex        =   8
      Top             =   3465
      Width           =   1875
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "角色授权(&G)"
      Height          =   350
      Left            =   7575
      TabIndex        =   7
      Top             =   1215
      Width           =   1875
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制角色(&C)"
      Height          =   350
      Left            =   7575
      TabIndex        =   9
      Top             =   1875
      Width           =   1875
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加角色(&A)"
      Height          =   350
      Left            =   7575
      TabIndex        =   6
      Top             =   870
      Width           =   1875
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7560
      TabIndex        =   3
      Top             =   555
      Width           =   1875
   End
   Begin MSComctlLib.TreeView tvwGroups 
      Height          =   2955
      Left            =   150
      TabIndex        =   1
      Top             =   870
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5212
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   295
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5880
      Top             =   3915
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
            Picture         =   "frmRole.frx":803A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUser 
      Cancel          =   -1  'True
      Caption         =   "修改角色的授权用户"
      Height          =   350
      Left            =   7575
      TabIndex        =   14
      Top             =   4155
      Width           =   1875
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改模块的使用权限"
      Height          =   350
      Left            =   7575
      TabIndex        =   13
      Top             =   3810
      Width           =   1875
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   4755
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   555
      Width           =   3915
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4620
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":85D4
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":92AE
            Key             =   "Role_Moved"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   2970
      Left            =   3510
      TabIndex        =   4
      Top             =   870
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5239
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_角色"
         Object.Tag             =   "角色"
         Text            =   "角色"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Grantee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin_Option"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "System"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SystemName"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1605
      Left            =   135
      TabIndex        =   15
      Top             =   4575
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "角色授权管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   17
      Top             =   150
      Width           =   1440
   End
   Begin VB.Label lblRoleGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "角色组信息"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label lblSearch 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查找"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7110
      TabIndex        =   2
      Top             =   615
      Width           =   360
   End
   Begin VB.Label lblModule 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已授权对象或功能"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "角色所属系统"
      Height          =   180
      Left            =   3510
      TabIndex        =   11
      Top             =   615
      Width           =   1200
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAddGroups 
         Caption         =   "新建组(&N)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "修改组(&M)"
      End
      Begin VB.Menu mnuPopuDeleteGroups 
         Caption         =   "删除组(&D)"
      End
   End
   Begin VB.Menu mnuPopuRole 
      Caption         =   "弹出菜单角色"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuRoleAdd 
         Caption         =   "增加角色(&N)"
      End
      Begin VB.Menu mnuPopuRoleDelete 
         Caption         =   "删除角色(&M)"
      End
      Begin VB.Menu mnuPopuRoleMove1 
         Caption         =   "角色移到分组(&M)…"
         Begin VB.Menu mnuPopuRoleMoveGroups 
            Caption         =   "组1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuPopuRoleMove2 
         Caption         =   "角色移到系统(&S)…"
         Begin VB.Menu mnuPopuRoleMoveSystems 
            Caption         =   "组2"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsRole As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnMoveTop As Boolean
Private msngPreHeigt As Single
Private mobjTip  As clsTipSwap           '悬浮提示框对象
Private mfrmGrant As frmRoleGrant
Private mstrSystemsName As String        '角色移动系统时的目标系统，用于展示提示文字

Private Enum lvwModuleHeader
    LH_系统 = 0
    LH_序号 = 1
    LH_功能或对象 = 2
    LH_说明 = 3
    LH_授权功能 = 4
    LH_系统号 = 5
End Enum

Private Enum LvwRoleHeader
    LRH_角色 = 0
    LRH_Grantee = 1
    LRH_Admin_Option = 2
    LRH_Group = 3
    LRH_System = 4
    LRH_SystemName = 5
End Enum

Private Sub chkOnlyShowNOSystem_Click()
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    tvwGroups_NodeClick tvwGroups.SelectedItem
End Sub

Private Sub cmdAdd_Click()
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    Dim lst As ListItem
    Dim str所有者() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, lngCount As Long
    Dim strUserName As String
    
    On Error GoTo ErrHandle
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 And glngSysNo = -1 Then
        MsgBox "若在本系统下创建了角色，" & vbNewLine & "那么该角色将被本系统单独管理。", vbInformation, gstrSysName
    End If
    '当用户拥有的角色数量达到148个时，用户登录时会提示错误
    gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!数量, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
            Exit Sub
        Else
            '所有者角色数量达到限制时，借助Systeme用户创建
            'SYSTEM所创建的角色，不授予所有者
            gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!数量, 0) >= 148 Then
                MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strRoleName = frmNameEdit.GetName(name角色)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
    strUserName = gstrUserName
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        strUserName = "SYSTEM"
        Set cnTemp = gcnSystem
    End If
    
    On Error Resume Next
    cnTemp.Execute "Create Role " & strRoleName & " Not Identified"
    
    If err <> 0 Then
        MsgBox "由于重名或命名错误或者角色数超过了数据库的参数限制" & vbCrLf & _
                "(可修改数据库启动参数调整角色数目)，导致角色增加失败。", vbExclamation, gstrSysName
        Call SetEnable
    Else
        On Error GoTo ErrHandle
        '将角色信息同步插入到Zlroles表中
        gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & strRoleName & "'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", "," & cmbSystem.ItemData(cmbSystem.ListIndex)) & ")"
        ExecuteProcedure gstrSQL, "新增角色与系统的对应关系"
        
        '插入重要操作日志
        Call SaveAuditLog(1, "增加角色", Split(strRoleName, "_")(1))
        
        strSQL = "Select Distinct s.所有者 From All_Tables t, Zlsystems s Where t.Table_Name = '部门表' And t.Owner = s.所有者"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
        ReDim str所有者(0 To rsTmp.RecordCount)
        Do While Not rsTmp.EOF
            str所有者(lngCount) = rsTmp!所有者
            lngCount = lngCount + 1
            rsTmp.MoveNext
        Loop
        Call GrantSpecialToRole(cnTemp, strRoleName, False, str所有者, True)
        If tvwGroups.SelectedItem Is Nothing Then
        ElseIf tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
            '刘兴宏:20070615加入
            '过程参数:zlTools.b_Rolegroupmgr.RoletoRolegroup
            '        组名_In In ZlRolegroups.组名%Type,
            '        角色_In In ZlRolegroups.角色%Type := Null
            gstrSQL = "zlTools.b_Rolegroupmgr.RoleToRoleGroup("
            gstrSQL = gstrSQL & "'" & Mid(tvwGroups.SelectedItem.Key, 2) & "',"
            gstrSQL = gstrSQL & "'" & strRoleName & "')"
            ExecuteProcedure gstrSQL, Me.Caption
        End If
        Set lst = lvwRole.ListItems.Add(, strRoleName, Mid(strRoleName, 4), "Role", "Role")
        If Not lst Is Nothing Then
            lst.SubItems(LRH_Grantee) = strUserName
            lst.SubItems(LRH_Admin_Option) = "YES"
        End If
        lst.Selected = True
        Call InitRoleData
        Call lvwRole_ItemClick(lst)
    End If
    Exit Sub
ErrHandle:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub DeleteRole()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除角色
    '编制:刘兴宏
    '日期:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRoleName As String
    Dim intIndex As Integer
    Dim strSQL As String, strUserList As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strRemarks As String
    
    On Error GoTo ErrHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    strRoleName = lvwRole.SelectedItem.Key
    intIndex = lvwRole.SelectedItem.Index
    
    If MsgBox("真的要删除角色“" & lvwRole.SelectedItem.Text & "”吗？", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    '判断该角色是否正在被使用
    strSQL = "Select Grantee 用户名 From Dba_Role_Privs Where Granted_Role = [1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strRoleName)
    If rsTemp.RecordCount > 1 Then
        '说明该角色正在被使用，则不能被删除
        For i = 1 To rsTemp.RecordCount
            If i > 3 Then Exit For
            strUserList = strUserList & "“" & rsTemp!用户名 & "”" & vbNewLine
            rsTemp.MoveNext
        Next
        If rsTemp.RecordCount > 3 Then
            MsgBox "该角色正在被" & vbNewLine & strUserList & "等" & rsTemp.RecordCount & _
            "个用户使用，要删除该角色请先修改以上用户的的角色！", vbInformation, gstrSysName
        Else
            MsgBox "该角色正在被用户" & vbNewLine & strUserList & _
            "使用，要删除该角色请先修改以上用户的的角色！", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    '验证身份并输入操作说明
    strRemarks = "删除角色：" & lvwRole.SelectedItem.Text
    If Not CheckAuditStatus("0401", "删除角色", strRemarks) Then Exit Sub
    Screen.MousePointer = 11
    If lvwRole.SelectedItem.SubItems(LRH_Grantee) = UCase(gstrUserName) _
        And lvwRole.SelectedItem.SubItems(LRH_Admin_Option) = "YES" Then
        gcnOracle.Execute "Drop Role " & strRoleName
    Else
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        gcnSystem.Execute "Drop Role " & strRoleName
    End If
    
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Delete('" & UCase(strRoleName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    gstrSQL = "zltools.Zl_Zlroles_Edit(3,'" & strRoleName & "')"
    ExecuteProcedure gstrSQL, "删除角色与系统的对应关系"
    
    '插入重要操作日志
    Call SaveAuditLog(3, "删除角色", lvwRole.SelectedItem.Text, strRemarks)
    lvwRole.ListItems.Remove intIndex
    If lvwRole.ListItems.Count > 0 Then
        If intIndex > lvwRole.ListItems.Count Then
            intIndex = lvwRole.ListItems.Count
        End If
        lvwRole.ListItems(intIndex).Selected = True
    End If
    Call InitRoleData
    Call FillModule
    Call SetEnable
    Screen.MousePointer = 0
    Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCopy_Click()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:复制角色
    '编制:刘兴宏
    '日期:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSourceRole As String
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    Dim objItem As ListItem
    Dim strUserName As String
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 And glngSysNo = -1 Then
        MsgBox "若在本系统下创建了角色，" & vbNewLine & "那么该角色将被本系统单独管理。", vbInformation, gstrSysName
    End If
    '当用户拥有的角色数量达到148个时，用户登录时会提示错误
    gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!数量, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not CheckRushHours("0401", "复制角色") Then
                Exit Sub
            End If
            '所有者角色数量达到限制时，借助Systeme用户创建
            'SYSTEM所创建的角色，不授予所有者
            gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!数量, 0) >= 148 Then
                MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strSourceRole = lvwRole.SelectedItem.Key
    strRoleName = frmNameEdit.GetName(name角色)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
    strUserName = gstrUserName
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        strUserName = "SYSTEM"
        Set cnTemp = gcnSystem
    End If
 
    If Not CopyRole(cnTemp, strSourceRole, strRoleName) Then Exit Sub
    
    '重新授权
    Call RoleGrant(strRoleName)
    '插入重要操作日志
    Call SaveAuditLog(1, "复制角色", "由" & Split(strSourceRole, "_")(1) & "复制得到" & Split(strRoleName, "_")(1))
    Set objItem = lvwRole.ListItems.Add(, strRoleName, Mid(strRoleName, 4), "Role", "Role")
    If Not objItem Is Nothing Then
        objItem.SubItems(LRH_Grantee) = strUserName
        objItem.SubItems(LRH_Admin_Option) = "YES"
    End If
    Call InitRoleData
    
    strKey = lvwRole.SelectedItem.Key
    err = 0: On Error Resume Next
    objItem.Selected = True
    If err = 0 Then
        lvwRole.ListItems(strKey).Selected = False
    End If
    Call lvwRole_ItemClick(lvwRole.SelectedItem)
    
    Exit Sub
ErrHandle:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Function CopyRole(cnTemp As ADODB.Connection, ByVal strSourceRole As String, ByVal strTargetRole As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拷贝角色权限成新的角色权限
    '参数:strSourceRole-源角色
    '     strTargetRole-目标角色
    '返回:复制成功,返回true,否则False
    '编制:刘兴宏
    '日期:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim str所有者() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, lngCount As Long
    
    err = 0: On Error Resume Next
    cnTemp.Execute "Create Role " & strTargetRole & " Not Identified"
    If err <> 0 Then
        MsgBox "由于重名或命名错误或者目标角色数超过了数据库的参数限制" & vbCrLf & _
                "(可修改数据库启动参数调整角色数目)，导致角色增加失败。", vbExclamation, gstrSysName
        Exit Function
    End If
    err = 0: On Error GoTo errHand:
    strSQL = "Select Distinct s.所有者 From All_Tables t, Zlsystems s Where t.Table_Name = '部门表' And t.Owner = s.所有者"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    ReDim str所有者(0 To rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        str所有者(lngCount) = rsTmp!所有者
        lngCount = lngCount + 1
        rsTmp.MoveNext
    Loop
    Call GrantSpecialToRole(cnTemp, strTargetRole, False, str所有者, True)
    '过程:zlTools.b_Rolegroupmgr.Role_Copy
    '    源角色_In   In zlRoleGrant.角色%Type,
    '    目标角色_In In zlRoleGrant.角色%Type
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Copy("
    gstrSQL = gstrSQL & "'" & UCase(strSourceRole) & "',"
    gstrSQL = gstrSQL & "'" & UCase(strTargetRole) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    '将角色信息同步插入到Zlroles表中
    gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & strTargetRole & "'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", "," & cmbSystem.ItemData(cmbSystem.ListIndex)) & ")"
    ExecuteProcedure gstrSQL, "新增角色与系统的对应关系"
    CopyRole = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Function RoleGrant(ByVal str角色 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定的角色重新授权
    '参数:str角色-角色
    '返回:复制成功,返回true,否则False
    '编制:刘兴宏
    '日期:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsUser As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str所有者() As String
    Dim lngCount As Long
    
    Me.MousePointer = vbHourglass
    On Error GoTo ErrHandle
    
    
    '得到可以所有进行授权的所有者名
    rsUser.CursorLocation = adUseClient
    gstrSQL = "select distinct S.所有者 from all_tables T,zlsystems S where T.table_name='部门表' And T.OWNER=S.所有者"
    rsUser.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    ReDim str所有者(0 To rsUser.RecordCount)
    Do Until rsUser.EOF
        str所有者(lngCount) = rsUser("所有者")
        lngCount = lngCount + 1
        rsUser.MoveNext
    Loop
    rsUser.Close

    '授予权限表中填写的权限
    Dim objclsPrivilege As New clsPrivilege
    Call objclsPrivilege.InitOracle(gcnOracle)
    Call objclsPrivilege.ReGrantPrivilege(str角色, str所有者)
    Set objclsPrivilege = Nothing
    
    MousePointer = 0
    RoleGrant = True
    Exit Function
ErrHandle:
    MousePointer = 0
    MsgBox "当前用户的权限不能完成本操作。", vbInformation, gstrSysName
End Function

Private Sub cmdDel_Click()
    '删除角色
    If cmdAdd.Enabled = False Then Exit Sub
    Call DeleteRole
End Sub

Private Sub cmdDeleteGroup_Click()
    '删除组
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
        Call DeleteRoleGroups
        Call LoadMenus
    End If
End Sub

Private Sub cmdGrant_Click()
    If Not CheckRushHours("0401", "角色授权") Then
        Exit Sub
    End If
    If mfrmGrant Is Nothing Then
        Set mfrmGrant = New frmRoleGrant
    End If
    If mfrmGrant.GrantToRole(lvwRole.SelectedItem.Key) = True Then
        Call FillModule
    End If
End Sub

Private Sub cmdGrantAll_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lst As ListItem, lngCount As Long
    Dim str所有者() As String
    
    If MsgBox("本功能根据应用系统保存的所有角色，在数据库中检查并补充创建角色，授予应用系统的公共基础对象权限,以及授予相关数据库对象的访问权限。" & vbCrLf & _
          "当在数据库中删除了角色，或者按用户模式恢复数据时，执行此操作可修正不一致的数据，以及对应用系统角色的对象重新授权。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    If Not CheckRushHours("0401", "恢复所有角色及权限") Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    On Error GoTo errH
    strSQL = "Select Distinct s.所有者 From All_Tables t, Zlsystems s Where t.Table_Name = '部门表' And t.Owner = s.所有者"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    ReDim str所有者(0 To rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        str所有者(lngCount) = rsTmp!所有者
        lngCount = lngCount + 1
        rsTmp.MoveNext
    Loop
    On Error Resume Next
    '首先创建不存在的角色
    strSQL = "Select Distinct r.角色" & vbNewLine & _
            "From Zlsystems s, Zlrolegrant r" & vbNewLine & _
            "Where s.编号 = r.系统 And s.所有者 = User And r.角色 Not In (Select Granted_Role From User_Role_Privs)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then err.Clear
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            '创建不存在的角色
            gcnOracle.Execute "Create Role " & rsTmp!角色 & " Not Identified"
            If err.Number = 0 Then
                '创建成功，将其增加到列表中
                Set lst = lvwRole.ListItems.Add(, rsTmp!角色 & "", Mid(rsTmp!角色 & "", 4), "Role", "Role")
                Call GrantSpecialToRole(gcnOracle, rsTmp!角色, False, str所有者, True)
                '将角色信息同步插入到Zlroles表中
                gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & rsTmp!角色 & "')"
                ExecuteProcedure gstrSQL, "新增角色与系统的对应关系"
            Else
                err.Clear
            End If
            rsTmp.MoveNext
        Loop
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    '开始授权
    Call ReGrantToRole(gcnOracle, "", True, str所有者)
    '插入重要操作日志
    Call SaveAuditLog(2, "恢复所有角色及权限", "恢复所有角色及权限")
    '显示授权清单
    If Not lst Is Nothing Then
        Call InitRoleData
        lst.Selected = True
        Call lvwRole_ItemClick(lst)
    End If
    MsgBox "所有角色重新授权完成！", vbInformation, gstrSysName
    MousePointer = 0
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MousePointer = 0
    MsgBox "当前用户的权限不能完成本操作。", vbInformation, gstrSysName
End Sub

Private Sub cmdModify_Click()
    frmProgPriv.ProgPriv
End Sub

Private Sub cmdModifyGroup_Click()
    '更名
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
        tvwGroups.SelectedItem.Text = Mid(tvwGroups.SelectedItem.Text, 1, Len(tvwGroups.SelectedItem.Text) - Len(tvwGroups.SelectedItem.Tag))
        tvwGroups.Tag = 1
        Call tvwGroups.StartLabelEdit
    End If
    tvwGroups.Tag = 0
    tvwGroups.SelectedItem.Text = tvwGroups.SelectedItem.Text & tvwGroups.SelectedItem.Tag
End Sub

Private Sub cmdNewGroup_Click()
    Dim strGroupsName As String
    Dim lst As ListItem
    Dim objNode As Node
ReDo:
    strGroupsName = frmNameEdit.GetName(name组名)
    
    If strGroupsName = "" Then Exit Sub
    If strGroupsName = "未分组" Then
        MsgBox "“未分组”为特殊分组,不能重复添加,请检查", vbInformation, gstrSysName
        Exit Sub
    End If
    If ActualLen(strGroupsName) > 30 Then
        MsgBox "你输入的角色名称不能大于30个字符或15个汉字,请检查!", vbDefaultButton1 + vbInformation, gstrSysName
        GoTo ReDo:
    End If
    strGroupsName = UCase(Trim(strGroupsName))
    
    '刘兴宏:20070615加入
    '过程参数:zlTools.b_Rolegroupmgr.Rolegroup_Add(组名_In In ZlRolegroups.组名%Type)
    err = 0: On Error GoTo errHand:
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Add("
    gstrSQL = gstrSQL & "'" & UCase(strGroupsName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
        
    Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & strGroupsName, strGroupsName & "(0)", 1, 1)
    objNode.Tag = "(0)"
    'objNode.Selected = True
    Call LoadMenu(strGroupsName, strGroupsName)
    
    Call FillModule
    Call SetEnable

'     '新增组
'     '方法是:在Tvw中增加一个新的角色名称
'     Dim objNode As Node
'     Dim int序号 As Integer
'     Dim strTargetGroup As String
'ReDo:
'    Err = 0: On Error Resume Next
'    int序号 = int序号 + 1
'    strTargetGroup = "新建组:" & int序号
'     Set objNode = tvwGroups.Nodes.Add(, "Root", strTargetGroup, strTargetGroup)
'     If Err <> 0 Then
'        Err.Clear: On Error GoTo 0
'        GoTo ReDo
'     End If
'     Err = 0
'    objNode.Tag = "1"
'    objNode.Selected = True
'    tvwGroups.SetFocus
'    tvwGroups.LabelEdit
    Exit Sub
errHand:
        Call ShowErrHand
End Sub
Private Sub ShowErrHand()
    '------------------------------------------------------------------------------------------
    '功能:获取错误信息
    '------------------------------------------------------------------------------------------
    Dim strNote As String, lngErrNum As Long
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '日志变量
            lngErrNum = gcnOracle.Errors(0).NativeError
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Sub
        Else
            MsgBox "注意:" & vbCrLf & "操作发生如下错误:" & vbCrLf & err.Description, vbExclamation, App.Title
        End If
    Else
        MsgBox "注意:" & vbCrLf & "操作发生如下错误:" & vbCrLf & err.Description, vbExclamation, App.Title
    End If
End Sub

Private Sub cmdRoleMove_Click()
    '将一个角色由一个系统移动到另一个分组
    Dim i As Long
    
    '根据当前选择角色所属系统或分组，将对于弹出菜单置灰
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveGroups.UBound
            If lvwRole.SelectedItem.SubItems(LRH_Group) = mnuPopuRoleMoveGroups(i).Tag Then
                mnuPopuRoleMoveGroups(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRoleMove1
    mnuPopuRoleMove1.Visible = True
End Sub

Private Sub cmdRolesReset_Click()
    On Error GoTo errH
    If MsgBox("本功能将清除本产品保存的所有角色，根据用户在数据库中实际拥有的角色重新产生所有角色数据。" & vbCrLf & _
                "当用户在应用系统中的角色与数据库中实际的角色不一致时，执行此操作可修正不一致的数据。" & vbCrLf & _
                "你确定要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call ExecuteProcedure("Zltools.Zl_Checkrolesdiff()", "检查并更新zlroles")
    Call FillRoleData(tvwGroups.SelectedItem.Key, True)
    '插入重要操作日志
    Call SaveAuditLog(2, "重整所有角色", "重整所有角色，操作完成。")
    MsgBox "重整所有角色，操作完成。", vbInformation, gstrSysName
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdSystemMove_Click()
    '将一个系统下的角色移动到另一个系统下
    Dim i As Long
    
    '根据当前选择角色所属系统或分组，将对于弹出菜单置灰
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveSystems.UBound
            If lvwRole.SelectedItem.SubItems(LRH_System) = mnuPopuRoleMoveSystems(i).Tag Then
                mnuPopuRoleMoveSystems(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRoleMove2
    mnuPopuRoleMove2.Visible = True
End Sub

Private Sub cmdUser_Click()
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Call frmRoleUser.ShowEdit(lvwRole.SelectedItem.Text)
End Sub


Private Sub Form_Activate()
    Dim lngTop As Long
    mblnMoveTop = False
    If mblnFirst = False Then Exit Sub
    '个性化设置
    lngTop = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具\角色管理", "PicHLine_TOP", "4500"))
    '如果老用户原来用过这个功能，那么注册表中一定能查到数据，如果比4500小，界面显示就会很难看
    If lngTop < 4500 Then lngTop = 4500
    picHLine.Top = lngTop
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
    mblnFirst = False
End Sub

Private Sub Form_Load()
   Dim rsTemp As New ADODB.Recordset
   Dim lngTop As Long
   
    '判断该用户能否创建角色
    gstrSQL = _
        " Select 1 From User_Sys_Privs Where Privilege='CREATE ROLE'" & _
        " Union" & _
        " Select 1 From Role_Sys_Privs Where Privilege='CREATE ROLE'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    cmdAdd.Enabled = rsTemp.RecordCount > 0
    
    If glngSysNo <> -1 Then
        cmdRolesReset.Visible = False
        cmdGrantAll.Visible = False
        cmdSystemMove.Visible = False
    End If
    Call Get注册码
    Call FillRollGroup
    Call FillSystem
    mblnFirst = True
End Sub

Private Sub cmbSystem_Click()
    cmbSystem.Tag = cmbSystem.ItemData(cmbSystem.ListIndex)
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 Then
        chkOnlyShowNOSystem.Visible = True
    Else
        chkOnlyShowNOSystem.Visible = False
    End If
    lvwRole.Tag = 1
    Call tvwGroups_NodeClick(tvwGroups.SelectedItem)
    lvwRole.Tag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '个性化设置
    SaveSetting "ZLSOFT", "公共模块\服务器管理工具\角色管理", "PicHLine_TOP", picHLine.Top
    If Not mfrmGrant Is Nothing Then
        Set mfrmGrant = Nothing
    End If
    Set mrsRole = Nothing
    Set mobjTip = Nothing
End Sub

Private Sub lvwModule_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwModule.SortKey = ColumnHeader.Index - 1
    lvwModule.SortOrder = Abs(Not lvwModule.SortOrder = 1)
    lvwModule.Sorted = True
End Sub

Private Sub lvwRole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillModule
    Call SetEnable
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    On Error Resume Next
    
    Me.lvwRole.Height = Me.picHLine.Top - Me.lvwRole.Top
    Me.tvwGroups.Height = Me.lvwRole.Height - cmdNewGroup.Height - 20
    Me.lblSystem.Left = lvwRole.Left
    lblModule.Top = picHLine.Top + picHLine.Height + 50
    
    With Me.lvwModule
        .Top = lblModule.Top + lblModule.Height + 50 ' cmdModify.Top + cmdModify.Height + 50
        If Me.ScaleHeight - .Top < 0 Then
             .Height = 0
        Else
            .Height = Me.ScaleHeight - .Top - 50
        End If
        .Width = ScaleWidth - 50 - .Left
        .ColumnHeaders(5).Width = .Width - .ColumnHeaders(1).Width - _
                                .ColumnHeaders(2).Width - .ColumnHeaders(3).Width - _
                                .ColumnHeaders(4).Width - .ColumnHeaders(6).Width
    End With
    
    With cmdAdd
        .Left = ScaleWidth - .Width - 50
    End With
    
    cmdGrant.Left = cmdAdd.Left
    cmdCopy.Left = cmdAdd.Left
    cmdDel.Left = cmdAdd.Left
    cmdRoleMove.Left = cmdAdd.Left
    cmdModify.Left = cmdAdd.Left
    cmdSystemMove.Left = cmdAdd.Left
    chkOnlyShowNOSystem.Left = cmdAdd.Left
    chkOnlyShowNOSystem.Top = txtSearch.Top
    
    With cmdUser
        .Top = lvwRole.Top + lvwRole.Height - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdModify
        .Top = cmdUser.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdGrantAll
        .Top = cmdModify.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdRolesReset
        .Top = cmdGrantAll.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With lvwRole
        If cmdAdd.Left - 50 - .Left < 0 Then
            .Width = 0
        Else
            .Width = cmdAdd.Left - 50 - .Left
        End If
    End With
    
    With txtSearch
        .Left = lvwRole.Width + lvwRole.Left - .Width
    End With
    
    With lblSearch
        .Left = txtSearch.Left - 50 - .Width
    End With
    
    cmdNewGroup.Top = tvwGroups.Top + tvwGroups.Height + 20
    cmdModifyGroup.Top = cmdNewGroup.Top
    cmdDeleteGroup.Top = cmdNewGroup.Top
    Me.picHLine.Left = 0: Me.picHLine.Width = Me.ScaleWidth
    msngPreHeigt = Me.ScaleHeight - picHLine.Top
End Sub

'Private Sub FillRole()
'    Dim rsTemp As New ADODB.Recordset
'
'    rsTemp.CursorLocation = adUseClient
'
'    '判断该用户能否创建角色
'    gstrSQL = "Select 1 from User_Sys_privs where privilege='CREATE ROLE' " & _
'        "union Select 1 from Role_Sys_privs where privilege='CREATE ROLE'"
'
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    cmdAdd.Enabled = rsTemp.RecordCount > 0
'    cmdDelete.Enabled = cmdAdd.Enabled
'    rsTemp.Close
'
'
'
'    '显示可以进行授权的角色
'    If gblnDBA = True Then
'        gstrSQL = "select * from DBA_Roles where Upper(Role) Like 'ZL_%'"
'    Else
'        gstrSQL = "select Granted_Role as Role from user_Role_privs " & _
'            "where Granted_Role Like 'ZL_%'" 'ADMIN_OPTION='YES'选项可以不加
'    End If
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    lvwRole.ListItems.Clear
'    Do Until rsTemp.EOF
'        lvwRole.ListItems.Add , rsTemp("Role"), Mid(rsTemp("Role"), 4), "Role", "Role"
'        rsTemp.MoveNext
'    Loop
'    If lvwRole.ListItems.Count > 0 Then
'        lvwRole.ListItems(1).Selected = True
'    Else
'        cmdGrant.Enabled = False
'    End If
'    rsTemp.Close
'    Call SetEnable
'End Sub

Private Sub FillSystem()
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    '若为单系统登录，则不显示系统
    If glngSysNo <> -1 Then
        lblSystem.Visible = False
        cmbSystem.Visible = False
        cmbSystem.addItem "单系统登录"
        cmbSystem.ItemData(cmbSystem.NewIndex) = glngSysNo
        cmbSystem.ListIndex = 0
        chkOnlyShowNOSystem.Visible = False
    Else
        '显示可以所有的系统
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
        Set rsTemp = zlGetRegSystems
        cmbSystem.Clear
        cmbSystem.addItem "所有系统"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
        Call LoadMenu("所有系统", "-1", False)
        Do Until rsTemp.EOF
            cmbSystem.addItem RPAD(rsTemp("名称") & "（" & rsTemp("编号") & "）", 25) & " v" & rsTemp("版本号")
            cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp("编号")
            Call LoadMenu(rsTemp("名称"), rsTemp("编号"), False)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
    End If
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillModule()
    Dim rsTemp As ADODB.Recordset
    Dim lst As ListItem
    Dim strRole As String
    Dim strSystem As String
    
    If cmbSystem.ListCount = 0 And glngSysNo = -1 Then Exit Sub
    
    '若选择的系统为所有系统,则显示该角色在所有系统下的权限
    '若选择的系统为具体的系统，则仅显示该角色在该系统下的权限
    LockWindowUpdate lvwModule.hwnd
    
    lvwModule.ColumnHeaders.Clear
    lvwModule.ListItems.Clear
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = lvwRole.SelectedItem.Key
    End If
    '更新列表项
    With lvwModule.ColumnHeaders
        .Add , , "系统", "2000"
        .Add , , "序号", "800"
        .Add , , "功能或对象", "1500"
        .Add , , "说明", "5000"
        .Add , , "授权功能", "1500"
        .Add , , "系统号", "0"
    End With
    
    If strRole = "" Then
        '角色为空，退出
        LockWindowUpdate 0
        Exit Sub
    End If
    
    '显示该角色能访问的基础表
    gstrSQL = "Select t.系统, t.编号, t.表名, t.说明" & vbNewLine & _
            "From (Select s.名称 || '（' || s.编号 || '）' As 系统, s.编号, s.所有者, b.表名, b.说明" & vbNewLine & _
            "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
            "       Where b.系统 = s.编号" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And s.编号 = [2]") & ") t, User_Tab_Privs r" & vbNewLine & _
            "Where t.所有者 = r.Owner And t.表名 = r.Table_Name And r.Grantee = [1] And" & vbNewLine & _
            "      r.Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
            "Group By t.系统, t.编号, t.表名, t.说明" & vbNewLine & _
            "Having Count(r.Privilege) = 4"

    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "获取基础编码信息", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, , rsTemp("系统"))
        lst.SubItems(LH_功能或对象) = rsTemp("表名")
        lst.SubItems(LH_说明) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        lst.SubItems(LH_系统号) = Lpad(rsTemp("编号"), 4)
        rsTemp.MoveNext
    Loop
    '显示该角色能访问的基础表
    gstrSQL = "Select s.名称 || '（' || s.编号 || '）' As 系统, s.编号, s.所有者, f.函数名, f.中文名, f.说明" & vbNewLine & _
            "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
            "Where f.系统 = s.编号 And s.所有者 = r.Owner And Upper(f.函数名) = r.Table_Name And r.Grantee = [1] And" & vbNewLine & _
            "      r.Privilege = 'EXECUTE'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And s.编号 = [2]")

    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "获取取数函数信息", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, , rsTemp("系统"))
        lst.SubItems(LH_功能或对象) = rsTemp("函数名") & "(" & rsTemp("中文名") & ")"
        lst.SubItems(LH_说明) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        lst.SubItems(LH_系统号) = Lpad(rsTemp("编号"), 4)
        rsTemp.MoveNext
    Loop
    '显示该角色能访问的基础工具
    gstrSQL = "select P.序号,P.标题,P.说明,R.功能 from " & _
            "zlRoleGrant R,zlPrograms P " & _
            "where R.系统 is Null And P.序号=R.序号 And R.角色=[1]" & _
            " And P.系统 is Null And P.序号<100 And P.部件 is Null " & _
            " Order By P.序号"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "获取基础工具信息", strRole)
    On Error Resume Next
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), "基础工具")
        lst.SubItems(LH_系统号) = "9999"
        If IsNull(rsTemp("序号")) Then
            If rsTemp("功能") <> "基本" Then
                Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
                lst.SubItems(LH_授权功能) = IIf(lst.SubItems(LH_授权功能) = "", "", lst.SubItems(LH_授权功能) & ",") & rsTemp("功能")
            End If
        Else
            lst.SubItems(LH_序号) = Lpad(rsTemp("序号"), 6)
            lst.SubItems(LH_功能或对象) = rsTemp("标题")
            lst.SubItems(LH_说明) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            If rsTemp("功能") <> "基本" Then
                lst.SubItems(LH_授权功能) = rsTemp("功能")
            End If
        End If
        rsTemp.MoveNext
    Loop
        
    '显示该角色能访问的模块以及自定义报表
    gstrSQL = "select S.名称||'（'||S.编号||'）' 系统,S.编号,P.序号,P.标题,P.说明,R.功能 from " & _
            "zlRoleGrant R,zlPrograms P ,zlSystems S " & _
            "where nvl(R.系统,0)=nvl(P.系统,0) And P.序号=R.序号 And P.序号>=100 And R.角色=[1] And P.系统 = S.编号(+)" & _
            IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And Nvl(P.系统, [2]) = [2]") & _
            " Order By P.序号"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "获取各个系统的信息", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    On Error Resume Next
    Do Until rsTemp.EOF
        strSystem = rsTemp("系统")
        If rsTemp("系统") = "（）" Then
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), "自定义报表")
            lst.SubItems(LH_系统号) = "8888"
        Else
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), rsTemp("系统"))
            lst.SubItems(LH_系统号) = Lpad(rsTemp("编号"), 4)
        End If
        If err <> 0 Then
            err.Clear
            If rsTemp("功能") <> "基本" Then
                Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
                lst.SubItems(LH_授权功能) = IIf(lst.SubItems(LH_授权功能) = "", "", lst.SubItems(LH_授权功能) & ",") & rsTemp("功能")
            End If
        Else
            lst.SubItems(LH_序号) = Lpad(rsTemp("序号"), 6)
            lst.SubItems(LH_功能或对象) = rsTemp("标题")
            lst.SubItems(LH_说明) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            If rsTemp("功能") <> "基本" Then
                lst.SubItems(LH_授权功能) = rsTemp("功能")
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    lvwModule.SortKey = LH_系统号
    lvwModule.SortOrder = 0
    lvwModule.Sorted = True
    
    LockWindowUpdate 0
End Sub

Private Sub SetEnable()
'设置各个按钮的Enable属性
    Dim blnHave As Boolean
    Dim i As Long
    Dim lstItem As ListItem
    blnHave = Not lvwRole.SelectedItem Is Nothing
    mnuPopuModify.Enabled = tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup"
    mnuPopuDeleteGroups.Enabled = mnuPopuModify.Enabled
    cmdModifyGroup.Enabled = mnuPopuModify.Enabled
    cmdDeleteGroup.Enabled = mnuPopuModify.Enabled
    'cmdDelete.Enabled = cmdAdd.Enabled And blnHave
    cmdGrant.Enabled = blnHave
    cmdUser.Enabled = blnHave
    cmdCopy.Enabled = blnHave
    cmdDel.Enabled = blnHave
    cmdRoleMove.Enabled = blnHave
    cmdGrantAll.Enabled = (gblnOwner = True)
    mnuPopuRoleDelete.Enabled = blnHave
    blnHave = False
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected = True Then
            blnHave = True
            Exit For
        End If
    Next
    For i = 1 To mnuPopuRoleMoveGroups.UBound
        mnuPopuRoleMoveGroups(i).Enabled = blnHave
        If UCase(mnuPopuRoleMoveGroups(i).Tag) = UCase(tvwGroups.SelectedItem.Key) Then
            mnuPopuRoleMoveGroups(i).Enabled = False
        End If
    Next
    For i = 1 To mnuPopuRoleMoveSystems.UBound
        mnuPopuRoleMoveSystems(i).Enabled = blnHave
        If mnuPopuRoleMoveSystems(i).Tag = cmbSystem.ItemData(cmbSystem.ListIndex) Then
            mnuPopuRoleMoveSystems(i).Enabled = False
        End If
    Next
    If glngSysNo <> -1 Then
        mnuPopuRoleMove2.Visible = False
    End If
End Sub


Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    Dim rsTmp As ADODB.Recordset
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "角色权限"
    Set objPrint.Body.objData = lvwModule
    objPrint.UnderAppItems.Add "角色：" & lvwRole.SelectedItem.Text
    If glngSysNo <> -1 Then
        gstrSQL = "Select 名称 From Zlsystems Where 编号 = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "查询系统名称", cmbSystem.ItemData(cmbSystem.ListIndex))
        objPrint.UnderAppItems.Add "授权系统：" & rsTmp!名称
    Else
        objPrint.UnderAppItems.Add "授权系统：" & cmbSystem.Text
    End If
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub lvwRole_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        Call mnuPopuRoleDelete_Click
    End Select
End Sub

Private Sub lvwRole_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    Dim strTip As String, strTitle As String
    
    If Button = 1 Then
        '设置图标
        If lvwRole.SelectedItem Is Nothing Then Exit Sub
        Set lvwRole.DragIcon = lvwRole.SelectedItem.CreateDragImage
        lvwRole.Drag 1
    End If
    
    '只有当非单系统登录并且当前选择系统为所有系统时，才弹出提示框
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 And glngSysNo = -1 And Button = 0 Then
        If mobjTip Is Nothing Then
            Call InitTips
        End If
        Set objItem = lvwRole.HitTest(x, y)
        If Not objItem Is Nothing Then
            If objItem.SubItems(LRH_SystemName) <> "所有系统" Then
                strTip = objItem.SubItems(LRH_SystemName)
                strTitle = "所属系统"
            Else
                strTip = ""
                strTitle = ""
            End If
            mobjTip.TipText = strTip
            mobjTip.Title = strTitle
        Else
            mobjTip.TipText = ""
            mobjTip.Title = ""
        End If
    End If
End Sub

Private Sub InitTips()
    Set mobjTip = New clsTipSwap
    Set mobjTip.ParentControl = lvwRole
    mobjTip.Icon = TTIconInfo
    mobjTip.Style = TTBalloon
    mobjTip.Create
End Sub

Private Sub lvwRole_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long

    If Button = 1 Then Exit Sub
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveGroups.UBound
            mnuPopuRoleMoveGroups(i).Enabled = True
            If lvwRole.SelectedItem.SubItems(LRH_Group) = mnuPopuRoleMoveGroups(i).Tag Then
                mnuPopuRoleMoveGroups(i).Enabled = False
            End If
        Next
        For i = 1 To mnuPopuRoleMoveSystems.UBound
            mnuPopuRoleMoveSystems(i).Enabled = True
            If lvwRole.SelectedItem.SubItems(LRH_System) = mnuPopuRoleMoveSystems(i).Tag Then
                mnuPopuRoleMoveSystems(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRole
End Sub

Private Sub mnuPopuAddGroups_Click()
    Call cmdNewGroup_Click
End Sub

Private Sub mnuPopuDeleteGroups_Click()
    Call cmdDeleteGroup_Click
End Sub

Private Sub mnuPopuModify_Click()
    Call cmdModifyGroup_Click
End Sub

Private Sub mnuPopuRoleAdd_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuPopuRoleDelete_Click()
    Call cmdDel_Click
End Sub

Private Sub mnuPopuRoleMoveGroups_Click(Index As Integer)
    Dim strTargetGroup As String
    If mnuPopuRoleMoveGroups(Index).Tag = "" Then Exit Sub
    strTargetGroup = Mid(mnuPopuRoleMoveGroups(Index).Tag, 2)
    
    If strTargetGroup = "NGROUP" Then
        If strTargetGroup = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移出该组吗?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
        strTargetGroup = ""
    Else
        If strTargetGroup = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移动到组“" & strTargetGroup & "”里吗?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(strTargetGroup, True) = False Then Exit Sub
    Call SetEnable
    
End Sub

Private Sub mnuPopuRoleMoveSystems_Click(Index As Integer)
    Dim strSystemsNo As String
    
    If mnuPopuRoleMoveSystems(Index).Tag = "" Then Exit Sub
    mstrSystemsName = mnuPopuRoleMoveSystems(Index).Caption
    strSystemsNo = mnuPopuRoleMoveSystems(Index).Tag
    
    If mstrSystemsName = "所有系统" Then
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移出该系统吗?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
        mstrSystemsName = ""
    Else
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移动到系统“" & mstrSystemsName & "”里吗?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(strSystemsNo, False) = False Then Exit Sub
    Call SetEnable
End Sub

Private Sub picHLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHLine.BackColor = &H8000000F: Me.picHLine.Top = Me.picHLine.Top + y
End Sub

Private Sub picHLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picHLine.BackColor = Me.BackColor
    If Me.picHLine.Top < 4500 Then Me.picHLine.Top = 4500
    If Me.picHLine.Top > Me.ScaleHeight - 1500 Then Me.picHLine.Top = Me.ScaleHeight - 1500
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
End Sub
 
Private Sub tvwGroups_AfterLabelEdit(Cancel As Integer, NewString As String)
    err = 0: On Error GoTo errHand:
    Dim strKey As String
    Dim strTag As String
    
    strKey = UCase(Mid(tvwGroups.SelectedItem.Key, 2))
    If strKey = NewString Then
        NewString = UCase(NewString) & tvwGroups.SelectedItem.Tag
        Exit Sub
    End If
    If NewString = "未分组" Then
        MsgBox "组名为:未分组已经存在,不能增加此组,请检查", vbExclamation, "服务器管理工具"
        NewString = tvwGroups.SelectedItem.Text
    Else
        '过程:zlTools.b_Rolegroupmgr.Rolegroup_Delete(
        '    组名_Old_In In ZlRolegroups.组名%Type,
        '    组名_New_In In ZlRolegroups.组名%Type
        gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Rename("
        gstrSQL = gstrSQL & "'" & strKey & "',"
        gstrSQL = gstrSQL & "'" & UCase(NewString) & "')"
        ExecuteProcedure gstrSQL, Me.Caption
        tvwGroups.SelectedItem.Key = "K" & NewString
        NewString = UCase(NewString) & tvwGroups.SelectedItem.Tag
        strTag = tvwGroups.SelectedItem.Tag
        Call InitRoleData
        tvwGroups.SelectedItem.Tag = strTag
    End If
    Exit Sub
errHand:
    Cancel = True
    Call ShowErrHand
End Sub
Private Sub tvwGroups_BeforeLabelEdit(Cancel As Integer)
    'tvwGroups.Tag = 0 表示仅允许点击按钮或双击才可修改组名，非此类方法均不可修改
    If Me.tvwGroups.SelectedItem.Key = "Root" Or Me.tvwGroups.SelectedItem.Key = "unGroup" Or Val(tvwGroups.Tag) = 0 Then
        Cancel = True
    End If
End Sub
Private Sub DeleteRoleGroups()
    '---------------------------------------------------------------------------------------------------------
    '功能:删除组
    '编制:刘兴宏
    '日期:2007/06/15
    '---------------------------------------------------------------------------------------------------------
    Dim strRoleGroupName As String
    Dim intIndex As Integer
    Dim rsTmp As ADODB.Recordset
    
    strRoleGroupName = Mid(tvwGroups.SelectedItem.Text, 1, Len(tvwGroups.SelectedItem.Text) - Len(tvwGroups.SelectedItem.Tag))
    intIndex = tvwGroups.SelectedItem.Index
    
    If MsgBox("真的要删除“" & strRoleGroupName & "”的角色组吗？", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    '若当前选择的系统为“所有系统”，则不作判断。
    '否则判断其它系统中该分组下是否还有角色，若有，则不能删除
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> "-1" Then
        '先判断当前系统中该分组下是否还有角色，即先检查lvwRole中有没有项目
        If lvwRole.ListItems.Count > 0 Then
            MsgBox "该分组下还有角色，故不能删除该分组。" & vbNewLine & "若一定要删除该分组，可将系统切换为“所有系统”，再进行删除分组操作！", vbInformation, gstrSysName
            Exit Sub
        Else
            gstrSQL = "Select Count(1) 数量 From Zlrolegroups b Where 组名 = [1] And 角色 Is Not Null"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "查找其它系统中该分组下的角色", strRoleGroupName)
            If rsTmp!数量 > 0 Then
                MsgBox "该分组下还有角色，故不能删除该分组。" & vbNewLine & "若一定要删除该分组，可将系统切换为“所有系统”，再进行删除分组操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    err = 0: On Error GoTo errHand:
    '过程:zlTools.b_Rolegroupmgr.Rolegroup_Delete(组名_In In ZlRolegroups.组名%Type)
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Delete("
    gstrSQL = gstrSQL & "'" & UCase(strRoleGroupName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    tvwGroups.Nodes.Remove intIndex
    If tvwGroups.Nodes.Count > 0 Then
        If intIndex > tvwGroups.Nodes.Count Then intIndex = tvwGroups.Nodes.Count
        tvwGroups.Nodes(intIndex).Selected = True
        tvwGroups.Nodes(intIndex).EnsureVisible
    End If
    Call FillRoleData(tvwGroups.SelectedItem.Key, IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, True, False))
    Exit Sub
errHand:
    Call ShowErrHand
End Sub

Private Sub tvwGroups_DblClick()
    Call cmdModifyGroup_Click
End Sub

Private Sub tvwGroups_DragDrop(Source As Control, x As Single, y As Single)
    Dim strTargetGroup As String, str角色 As String, intIndex As Integer
    Dim lstItem As ListItem
    Dim strKeys As String
    Dim arrVar As Variant
    Dim i As Long
    
    err = 0: On Error GoTo errHand:
    If Source Is lvwRole And Not tvwGroups.DropHighlight Is Nothing Then
        intIndex = -1
        strTargetGroup = Mid(tvwGroups.DropHighlight.Key, 2)
        Set tvwGroups.DropHighlight = Nothing
        tvwGroups.DropHighlight = tvwGroups.SelectedItem

        If strTargetGroup = "oot" Or strTargetGroup = "nGroup" Then
            If MsgBox("你真要将角色“" & Source.SelectedItem.Text & "...” 移出该组吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            strTargetGroup = ""
        Else
            If MsgBox("你真要将角色“" & Source.SelectedItem.Text & "...” 移动到组“" & strTargetGroup & "”里吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If

        gcnOracle.BeginTrans
        strKeys = ""
        For Each lstItem In lvwRole.ListItems
            If lstItem.Selected Then
                If intIndex < 0 Then
                    intIndex = lstItem.Index
                End If
                str角色 = lstItem.Key
                strKeys = strKeys & "'" & lstItem.Key

                If MoveToRoleGroup(strTargetGroup, str角色) = False Then
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            End If
        Next
        gcnOracle.CommitTrans
        If strKeys <> "" Then strKeys = Mid(strKeys, 2)
        '若当前节点是在所有角色中，则无需调整
        Call InitRoleData
        If tvwGroups.SelectedItem.Key <> "Root" Then
            arrVar = Split(strKeys, "'")
            For i = 0 To UBound(arrVar)
                lvwRole.ListItems.Remove arrVar(i)
            Next

            If lvwRole.ListItems.Count > 0 Then
                If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
                lvwRole.ListItems(intIndex).Selected = True
            End If
            Call FillModule
        Else
            lvwRole.SelectedItem.SubItems(LRH_Group) = IIf(strTargetGroup = "", "UNGROUP", "K" & strTargetGroup)
        End If
    End If
    Call SetEnable
    tvwGroups.Refresh
     
    Set tvwGroups.DropHighlight = Nothing
    Exit Sub
errHand:
    Set tvwGroups.DropHighlight = Nothing
    Call ShowErrHand
End Sub
Private Function MoveToGroups(ByVal strTargetGroup As String, Optional ByVal blnType As Boolean = True) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '功能:将指定角色移动到组中
    '参数:
    '     strTargetGroup-移到目标组的组名，若是移动到系统，则为系统号，若是移动到分组，则为组名
    '     blnType-分组类型
    '         blnType = True  按角色分组
    '         blnType = False 按系统分组
    '移动成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strKeys  As String
    Dim lstItem As ListItem
    Dim str角色 As String
    Dim arrVar As Variant
    Dim i As Long
    MoveToGroups = False
    intIndex = -1
    gcnOracle.BeginTrans
    strKeys = ""
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected Then
            If intIndex < 0 Then
                intIndex = lstItem.Index
            End If
            str角色 = lstItem.Key
            strKeys = strKeys & "'" & lstItem.Key
            If blnType = True Then
                If MoveToRoleGroup(strTargetGroup, str角色) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            Else
                If MoveToSystemGroup(strTargetGroup, str角色) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
        End If
    Next
    gcnOracle.CommitTrans
    If strKeys <> "" Then strKeys = Mid(strKeys, 2)
    '若当前节点是在所有角色或所有系统中，则无需调整，但要弹出移动成功的提示
    Call InitRoleData
    If (blnType = True And tvwGroups.SelectedItem.Key <> "Root") Or (blnType = False And (cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 Or chkOnlyShowNOSystem.value = 1)) Then
        arrVar = Split(strKeys, "'")
        For i = 0 To UBound(arrVar)
            lvwRole.ListItems.Remove arrVar(i)
        Next
        If lvwRole.ListItems.Count > 0 Then
            If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
            lvwRole.ListItems(intIndex).Selected = True
        End If
        FillModule
    Else
        If blnType Then
            '按角色分组
            lvwRole.SelectedItem.SubItems(LRH_Group) = IIf(strTargetGroup = "", "UNGROUP", "K" & strTargetGroup)
        Else
            '按系统分组
            lvwRole.SelectedItem.SubItems(LRH_System) = strTargetGroup
            lvwRole.SelectedItem.SubItems(LRH_SystemName) = mstrSystemsName
            If strTargetGroup = -1 Then
                lvwRole.SelectedItem.Icon = "Role"
            Else
                lvwRole.SelectedItem.Icon = "Role_Moved"
        End If
        End If
    End If
    MoveToGroups = True
End Function
Private Function MoveToRoleGroup(ByVal str组 As String, str角色 As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '功能:将指定的角色移到组中
    '参数:str组-移到组的组名
    '     str角色-指定的角色
    '移动成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    
    '判断该角色是否在该组中已存在
    If str组 = "" Then
        mrsRole.Filter = "组名 = null And Role = '" & str角色 & "'"
    Else
        mrsRole.Filter = "组名 = '" & str组 & "' And Role = '" & str角色 & "'"
    End If
    If mrsRole.RecordCount = 1 Then
        MsgBox "角色“" & str角色 & "”已存在于" & vbNewLine & "分组“" & IIf(str组 = "", "未分组", str组) & "”中，无需再次移动！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '刘兴宏:20070615加入
    '过程参数:zlTools.b_Rolegroupmgr.RoletoRolegroup
    '        组名_In In ZlRolegroups.组名%Type,
    '        角色_In In ZlRolegroups.角色%Type := Null
    gstrSQL = "zlTools.b_Rolegroupmgr.RoletoRolegroup("
    gstrSQL = gstrSQL & IIf(str组 = "", "Null", "'" & UCase(str组) & "'") & ","
    gstrSQL = gstrSQL & "'" & UCase(str角色) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    MoveToRoleGroup = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Function MoveToSystemGroup(ByVal lngGroup As Long, ByVal strRoleName As String) As Boolean
    '功能：将指定角色移动到系统分组中
    '参数：
    '     lngGroup-需要移到的系统的编号，若为-1，表示要将角色移动到“所有系统”
    '     strRoleName-需要移动的角色的名称
    
    On Error GoTo errHand:
    
    '判断该角色是否在该系统中已存在
    mrsRole.Filter = "系统 = " & lngGroup & " And Role = '" & strRoleName & "'"
    If mrsRole.RecordCount = 1 Then
        MsgBox "角色“" & strRoleName & "”已存在于" & vbNewLine & "该系统中，无需再次移动！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zltools.Zl_Zlroles_Edit(2,'" & strRoleName & "'" & IIf(lngGroup = -1, "", "," & lngGroup) & ")"
    ExecuteProcedure gstrSQL, "修改角色所在系统"
    MoveToSystemGroup = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Sub tvwGroups_DragOver(Source As Control, x As Single, y As Single, State As Integer)
      Dim objOver As Node
      If Source Is lvwRole Then
           Set objOver = tvwGroups.HitTest(x, y)
            If Not objOver Is Nothing Then
                '判断当前所选角色是否已在当前所选分组里
                If lvwRole.SelectedItem.SubItems(LRH_Group) = UCase(objOver.Key) Or objOver.Key = "Root" Then
                    Set tvwGroups.DropHighlight = Nothing
                    lvwRole.DragIcon = Nothing
                Else
                    Set tvwGroups.DropHighlight = objOver
                    tvwGroups.DropHighlight.EnsureVisible
                    lvwRole.DragIcon = ils32.ListImages(lvwRole.SelectedItem.Icon).Picture
                    
                End If
            Else
                Set tvwGroups.DropHighlight = Nothing
            End If
      End If
End Sub

Private Sub tvwGroups_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        mnuPopuDeleteGroups_Click
    End Select
End Sub

Private Sub tvwGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    PopupMenu mnuPopu
End Sub

Private Sub tvwGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    '---------------------------------------------------------------------------------------------------------
    '获取相应的角色权限
    '---------------------------------------------------------------------------------------------------------
    Call FillRoleData(Node.Key, IIf(Val(lvwRole.Tag) = 1, True, False))
End Sub

Private Function FillRoleData(ByVal strTargetGroup As String, Optional ByVal blnRefreshData As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:根据组名,获取角色信息,并填充到lvw中
    '参数:
    '     strTargetGroup:<>""指定组名,=""表示所有角色
    '     blnRefreshData:标记是否需要刷新角色数据
    '返回:加载成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------
    Dim rsGroups As ADODB.Recordset
    Dim objItem As ListItem
    Dim strSearch  As String, strFilter As String
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    strSearch = UCase(Trim(txtSearch.Text))
    If strTargetGroup <> "Root" And strTargetGroup <> "unGroup" Then
        strTargetGroup = Mid(strTargetGroup, 2)
    End If
     
    '当查看数据时，不对数据源进行刷新
    '当进行了数据修改后，才重新刷新数据源
    If blnRefreshData Then Call InitRoleData
    If mrsRole Is Nothing Then Call InitRoleData
    
    If strTargetGroup = "Root" Then
        strFilter = ""
    ElseIf strTargetGroup = "unGroup" Then
        strFilter = "组名 = null"
    Else
        strFilter = "组名 = '" & strTargetGroup & "'"
    End If
    
    '判断是否仅显示未分系统的角色
    If chkOnlyShowNOSystem.Visible = True And chkOnlyShowNOSystem.value = 1 Then
        If strFilter = "" Then
            strFilter = "系统 = null"
        Else
            strFilter = strFilter & " And 系统 = null"
        End If
    End If
    
    If strSearch <> "" Then
        If strFilter = "" Then
            strFilter = "RoleName Like '" & strSearch & "%' or 简码 Like '" & strSearch & "%'"
        Else
            strFilter = "(" & strFilter & " And RoleName Like '" & strSearch & "%') or (" & strFilter & " And 简码 Like '" & strSearch & "%')"
        End If
    End If
    mrsRole.Filter = strFilter
    lvwRole.ListItems.Clear
    
    With mrsRole
        Do Until .EOF
            If IsNull(!系统) Or cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 Then
            Set objItem = lvwRole.ListItems.Add(, Nvl(!Role), Nvl(!RoleName), "Role", "Role")
            Else
                Set objItem = lvwRole.ListItems.Add(, Nvl(!Role), Nvl(!RoleName), "Role_Moved", "Role_Moved")
            End If
            If Not objItem Is Nothing Then
                objItem.SubItems(LRH_Grantee) = Nvl(!Grantee)
                objItem.SubItems(LRH_Admin_Option) = Nvl(!Admin_Option)
                objItem.SubItems(LRH_Group) = IIf(IsNull(!组名), "UNGROUP", "K" & !组名)
                objItem.SubItems(LRH_System) = IIf(IsNull(!系统), "-1", !系统)
                objItem.SubItems(LRH_SystemName) = IIf(IsNull(!系统名称), "所有系统", !系统名称)
            End If
            mrsRole.MoveNext
        Loop
    End With
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(LRH_Grantee).Selected = True
        Call lvwRole_ItemClick(lvwRole.SelectedItem)
    Else
        cmdGrant.Enabled = False
        Call SetEnable
    End If
    mrsRole.Filter = 0
    
    FillRoleData = True
    
    Exit Function
ErrHandle:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Function

Private Sub InitRoleData()
    '初始化记录集并填充每个分组中的角色个数
    Dim Node As Node
    
    On Error GoTo errH
    gstrSQL = _
        " Select User as Grantee,'YES' as Admin_Option,Substr(A.名称,4) as RoleName," & _
        " A.名称 Role,zlSpellCode(Substr(A.名称,4)) as 简码, b.组名, a.系统, c.名称 系统名称" & _
        " From zlTools.Zlroles A,zlTools.zlRoleGroups B, zlSystems C" & _
        " Where A.名称=B.角色(+) and A.系统 = C.编号(+)" & _
        IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And A.系统 = [1]") & _
        " Order by A.名称"
    Set mrsRole = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption, cmbSystem.ItemData(cmbSystem.ListIndex))
    
    For Each Node In tvwGroups.Nodes
        If Node.Key <> "Root" Then
            mrsRole.Filter = IIf(Node.Key = "unGroup", "组名 = null", "组名 = '" & Mid(Node.Text, 1, Len(Node.Text) - Len(Node.Tag)) & "'") & ""
        Else
            mrsRole.Filter = ""
        End If
        If Node.Tag = "" Then
            Node.Text = Node.Text & "(" & mrsRole.RecordCount & ")"
        Else
            Node.Text = Mid(Node.Text, 1, Len(Node.Text) - Len(Node.Tag)) & "(" & mrsRole.RecordCount & ")"
        End If
        Node.Tag = "(" & mrsRole.RecordCount & ")"
    Next
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function SearchRole(ByVal strFilter As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:过虑出相应的角色
    '参数:strFilter-过滤串
    '返回:成功,返回ture,否则返回False
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    SearchRole = True
    If mrsRole Is Nothing Then Exit Function
    If mrsRole.State <> 1 Then Exit Function
    If mrsRole.RecordCount = 0 Then Exit Function
    
    strFilter = UCase(strFilter)
    SearchRole = False
    If strFilter = "" Then
    Else
        mrsRole.Filter = "RoleName Like '" & strFilter & "%' or 简码 Like '" & strFilter & "%'"
    End If
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        lvwRole.ListItems.Add , Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role"
        mrsRole.MoveNext
    Loop
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(LRH_Grantee).Selected = True
    Else
        cmdGrant.Enabled = False
    End If
    Call SetEnable
    mrsRole.Filter = 0
    SearchRole = True
End Function

Private Sub txtSearch_Change()
    Call SearchRole(Trim(txtSearch.Text))
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub
Private Sub FillRollGroup()
    '--------------------------------------------------------------------------------------------
    '功能:加载角色组
    '编制:刘兴宏
    '日期:2007/06/15
    '--------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Dim objNode As Node
    gstrSQL = "Select distinct 组名 From zlRoleGroups"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With tvwGroups
        .Nodes.Clear
        Set objNode = tvwGroups.Nodes.Add(, 4, "Root", "所有角色", 1, 1)
        objNode.Selected = True
        objNode.Expanded = True
        Set objNode = tvwGroups.Nodes.Add("Root", tvwChild, "unGroup", "未分组", 1, 1)
        objNode.Sorted = True
        Call LoadMenu("未分组", "unGroup")
        Do While Not rsTemp.EOF
            Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & Nvl(rsTemp!组名), Nvl(rsTemp!组名), 1, 1)
            objNode.Sorted = True
            Call LoadMenu(Nvl(rsTemp!组名), Nvl(rsTemp!组名))
            rsTemp.MoveNext
        Loop
    End With
End Sub
Private Sub LoadMenu(ByVal strTittle As String, ByVal strTag As String, Optional ByVal blnType As Boolean = True)
'加载“角色移动到分组”“角色移动到菜单”弹出列表
'blnType = true : 加载分组菜单
'blnType = false : 加载系统菜单

        Dim varMenu As Variant
        Dim intCount As Integer
        
        If blnType = True Then
            Set varMenu = mnuPopuRoleMoveGroups
            intCount = varMenu.Count
            Load varMenu(intCount)
            varMenu(intCount).Caption = strTittle
            If strTag = "unGroup" Then
                varMenu(intCount).Tag = UCase(strTag)
            Else
                varMenu(intCount).Tag = UCase("K" & strTag)
            End If
            varMenu(intCount).Visible = True
            mnuPopuRoleMove1.Visible = True
            varMenu(0).Visible = False
        Else
            Set varMenu = mnuPopuRoleMoveSystems
            intCount = varMenu.Count
            Load varMenu(intCount)
            varMenu(intCount).Caption = strTittle
            varMenu(intCount).Tag = strTag
            varMenu(intCount).Visible = True
            mnuPopuRoleMove2.Visible = True
            varMenu(0).Visible = False
        End If
End Sub
Private Sub LoadMenus()
    Dim objNode As Node
    Call UnLoadMenus
    For Each objNode In tvwGroups.Nodes
        If objNode.Key <> "Root" Then
            Call LoadMenu(Mid(objNode.Text, 1, Len(objNode.Text) - Len(objNode.Tag)), IIf(objNode.Key = "unGroup", objNode.Key, Mid(objNode.Key, 2)))
        End If
    Next
End Sub
Private Sub UnLoadMenus()
    '功能:拆卸菜单
        Dim varMenu As Variant
        Dim intCount As Integer
        Set varMenu = mnuPopuRoleMoveGroups
        mnuPopuRoleMoveGroups(0).Visible = True
        mnuPopuRoleMove1.Visible = True
        For intCount = 1 To mnuPopuRoleMoveGroups.UBound
            Unload varMenu(intCount)
        Next
        
End Sub

