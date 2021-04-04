VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRoleBat 
   Caption         =   "批量角色分配"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmPresRoleBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7710
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "回收权限"
      Height          =   350
      Left            =   3600
      TabIndex        =   7
      Top             =   4680
      Width           =   900
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "授予权限"
      Height          =   350
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   900
   End
   Begin VB.ListBox lstRole 
      Height          =   1680
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "科室："
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
         TabIndex        =   4
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   4710
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   600
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "B.名称"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   7800
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   120
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":6852
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":752C
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":DD8E
            Key             =   "YES"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1560
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":145F0
            Key             =   "YES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresRoleBat.frx":1AE52
            Key             =   "NO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1725
      Left            =   4710
      TabIndex        =   9
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3043
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "系统"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "模块"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUnGrantedPres 
      Height          =   4005
      Left            =   600
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "姓名"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "用户名"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSComctlLib.ListView lvwGrantedPres 
      Height          =   4005
      Left            =   4710
      TabIndex        =   11
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "姓名"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "用户名"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label lblRole 
      Caption         =   "角色"
      Height          =   255
      Left            =   98
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblUnGrantedPres 
      Caption         =   "未授予该角色的人员"
      Height          =   315
      Left            =   570
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblGrantedPres 
      Caption         =   "已授予该角色的人员"
      Height          =   195
      Left            =   4710
      TabIndex        =   15
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lblModule 
      Caption         =   "模块清单"
      Height          =   255
      Left            =   3825
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      Caption         =   "授权内容"
      Height          =   180
      Left            =   3825
      TabIndex        =   13
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "查找"
      Height          =   180
      Left            =   98
      TabIndex        =   12
      Top             =   660
      Width           =   360
   End
End
Attribute VB_Name = "frmPresRoleBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDeptId      As Long     '当前科室ID
Private mstrDeptName    As String   '当前科室名称
Private mintOld         As Integer
Public Sub ShowMe(ByVal frmParent As Object, ByVal lngDept As Long, ByVal strDeptName As String)
    mlngDeptId = lngDept
    mstrDeptName = strDeptName
    Me.Show vbModal, frmParent
End Sub

Private Sub cboSystem_Click()
    FillModule
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGrant_Click()
    Dim i As Integer
    Dim strErr As String
    
    On Error Resume Next
    For i = 1 To lvwUnGrantedPres.ListItems.Count
        If lvwUnGrantedPres.ListItems(i).Checked Then
            gstrSQL = "Grant ZL_" & lstRole.Text & "  to " & lvwUnGrantedPres.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL, , adCmdText
            If Err <> 0 Then
                strErr = strErr & vbCrLf & Err.Description
                Err.Clear
            Else
                Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Add('" & lvwUnGrantedPres.ListItems(i).SubItems(1) & "','ZL_" & lstRole.Text & "')", Me.Caption)
            End If
        End If
    Next
    If strErr <> "" Then
        MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    Dim strErr As String
    On Error Resume Next
    
    For i = 1 To lvwGrantedPres.ListItems.Count
        If lvwGrantedPres.ListItems(i).Checked Then
            gstrSQL = "revoke ZL_" & lstRole.Text & " from " & lvwGrantedPres.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL, , adCmdText
            If Err <> 0 Then
                strErr = strErr & vbCrLf & Err.Description
                Err.Clear
            Else
                Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Del('" & lvwGrantedPres.ListItems(i).SubItems(1) & "','ZL_" & lstRole.Text & "')", Me.Caption)
            End If
        End If
    Next
    
    If strErr <> "" Then
        MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub Form_Load()
    lblName.Caption = "科室：" & mstrDeptName
    Call FillRoleAndSystem
End Sub

Private Sub FillRoleAndSystem()
'加载当前登录用户具有管理权限的角色和系统
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    Dim lstTmp As ListItem
    
    On Error GoTo errH
    strSQL = "Select Substr(Granted_Role, 4) 角色" & vbNewLine & _
            "From Dba_Role_Privs" & vbNewLine & _
            "Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "And Admin_Option = 'YES'" & vbNewLine & _
            "And Grantee = User" & vbNewLine & _
            "Order By Substr(Granted_Role, 4)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-用户角色")
    With lstRole
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!角色
            rsTmp.MoveNext
        Next
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
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillModule()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lst As ListItem
    Dim strRole As String, strPre As String
    
    On Error GoTo errH
    lvwModule.ListItems.Clear
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
    If lstRole.ListIndex <> -1 Then
        strRole = "ZL_" & lstRole.List(lstRole.ListIndex)
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
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lstRole_Click()
    FillModule
    LoadPreson
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intOld As Integer
    Dim blnEnd As Boolean
    
    If KeyAscii <> 13 Then Exit Sub
    If txtFind.Text = "" Then Exit Sub
    
    zlControl.TxtSelAll txtFind
    
    If txtFind.Tag <> txtFind.Text Then
        mintOld = 0
        txtFind.Tag = txtFind.Text
    Else
        If mintOld + 1 >= lstRole.ListCount Then
            mintOld = 0
            txtFind.Tag = ""
        End If
    End If
    
RowX:
    For i = mintOld To lstRole.ListCount
        If InStr(1, lstRole.List(i), txtFind.Text) > 0 Or InStr(1, zlStr.GetCodeByVB(lstRole.List(i)), UCase(txtFind.Text)) > 0 Then
            lstRole.Selected(i) = True
            mintOld = i + 1
            blnEnd = True
            Exit Sub
        End If
    Next
    
    If Not blnEnd And mintOld <> 0 Then
        mintOld = 0
        GoTo RowX
    End If
End Sub

Private Sub LoadPreson()
'加载人员
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lstTmp As ListItem
    
    On Error GoTo errH
    strSQL = "Select a.用户名, d.姓名, Decode(e.Granted_Role, Null, 0, 1) 权限" & vbNewLine & _
            "From 上机人员表 a, 部门人员 b, 部门表 c, 人员表 d, (Select Grantee, Granted_Role From Dba_Role_Privs Where Granted_Role = [2]) e" & vbNewLine & _
            "Where a.人员id = b.人员id" & vbNewLine & _
            "And b.部门id = c.Id" & vbNewLine & _
            "And d.Id = a.人员id" & vbNewLine & _
            "And a.用户名 = e.Grantee(+)" & vbNewLine & _
            "And b.部门id = [1]" & vbNewLine & _
            "And a.用户名 <> [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-用户角色", mlngDeptId, "ZL_" & lstRole.Text, UCase(gstrDbUser))
    
    lvwGrantedPres.ListItems.Clear
    lvwUnGrantedPres.ListItems.Clear
    Do While Not rsTmp.EOF
        If rsTmp!权限 = 1 Then
            Set lstTmp = lvwGrantedPres.ListItems.Add(, rsTmp!用户名, rsTmp!姓名, , "YES")
            lstTmp.SubItems(1) = rsTmp!用户名
        Else
            Set lstTmp = lvwUnGrantedPres.ListItems.Add(, rsTmp!用户名, rsTmp!姓名, , "NO")
            lstTmp.SubItems(1) = rsTmp!用户名
        End If
        lstTmp.Checked = True
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

