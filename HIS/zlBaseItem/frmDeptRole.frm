VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRoleBat 
   Caption         =   "批量角色分配"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmDeptRole.frx":0000
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
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   570
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
            Picture         =   "frmDeptRole.frx":6852
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":752C
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":DD8E
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
            Picture         =   "frmDeptRole.frx":145F0
            Key             =   "YES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":1AE52
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
   Begin MSComctlLib.ListView lvwGrant 
      Height          =   4005
      Left            =   570
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
   Begin MSComctlLib.ListView lvwRemove 
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
   Begin VB.Label lblNO人员 
      Caption         =   "未授予该角色的人员"
      Height          =   315
      Left            =   570
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lbl人员 
      Caption         =   "已授予该角色的人员"
      Height          =   195
      Left            =   4710
      TabIndex        =   15
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lbl已授权 
      Caption         =   "模块清单"
      Height          =   255
      Left            =   3825
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblContent 
      AutoSize        =   -1  'True
      Caption         =   "授权内容"
      Height          =   180
      Left            =   3825
      TabIndex        =   13
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "查找"
      Height          =   180
      Index           =   2
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
Private mlngDept As Long

Private Sub FillRole()
    Const STR_ICON = "Role"
    Dim rsTmp As ADODB.Recordset
    Dim lstTmp As ListItem
    Dim i As Long

    On Error GoTo errHandle
    gstrSQL = "Select a.角色, Decode(b.角色, Null, 0, 1) As 应用" & vbNewLine & _
            "  From (Select Substr(Granted_Role, 4) 角色" & vbNewLine & _
            "           From Dba_Role_Privs" & vbNewLine & _
            "          Where Granted_Role Like 'ZL_%' And Admin_Option = 'YES' And Grantee = User) a, (Select Distinct Substr(B1.Granted_Role, 4) 角色" & vbNewLine & _
            "           From Dba_Role_Privs B1" & vbNewLine & _
            "          Where B1.Granted_Role Like" & vbNewLine & _
            "                'ZL_%') b" & vbNewLine & _
            " Where a.角色 = b.角色(+)" & vbNewLine & _
            " Order By a.角色"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-用户角色")
    With Me.lstRole
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!角色
            rsTmp.MoveNext
        Next
        rsTmp.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal lngDept As Long, ByVal strDeptName As String)
    mlngDept = lngDept
    lblName.Caption = "科室：" & strDeptName
    Me.Show
End Sub

Private Sub cboRole_Click()
'    LoadInfo
    LoadPreson
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
    
    For i = 1 To Me.lvwGrant.ListItems.Count
        If Me.lvwGrant.ListItems(i).Checked Then
            gstrSQL = "Grant ZL_" & Me.lstRole.Text & "  to " & Me.lvwGrant.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL
            If Err <> 0 Then
                strErr = vbCrLf & Err.Description
'                MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
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
    
    For i = 1 To Me.lvwRemove.ListItems.Count
        If Me.lvwRemove.ListItems(i).Checked Then
            gstrSQL = "revoke ZL_" & Me.lstRole.Text & " from " & Me.lvwRemove.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL
            If Err <> 0 Then
                strErr = vbCrLf & Err.Description
'                MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
            End If
        End If
    Next
    
    If strErr <> "" Then
        MsgBox "权限不足，授予角色失败。" & vbCrLf & "错误信息如下:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub Form_Load()
    FillRole
    
    FillSystem
    
End Sub

'Private Sub LoadInfo()
'    Dim rsTemp As Recordset
'    Dim lstTmp As ListItem
'
'    On Error GoTo errHandle
'    gstrSQL = "Select Distinct a.名称, b.标题,B.序号" & vbNewLine & _
'    "From Zlsystems a, Zlprograms b, Zlrolegrant c" & vbNewLine & _
'    "Where a.编号 = c.系统 And b.系统 = c.系统 And b.序号 = c.序号 And c.角色=[1]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadInfo", "ZL_" & Me.lstRole.Text)
'
'    With Me.lvwRole
'        .ListItems.Clear
'        Do While Not rsTemp.EOF
'            Set lstTmp = .ListItems.Add(, "A" & rsTemp!序号, rsTemp!名称)
'            lstTmp.SubItems(1) = rsTemp!标题
'            rsTemp.MoveNext
'        Loop
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub LoadPreson()
    Dim rsTmp As Recordset
    Dim lstTmp As ListItem
    
    On Error GoTo errHandle
    gstrSQL = "Select a.用户名, d.姓名,decode(nvl(e.Granted_Role,''),'',0,1) 权限" & vbNewLine & _
            "From 上机人员表 a, 部门人员 b, 部门表 c, 人员表 d, Dba_Role_Privs e" & vbNewLine & _
            "Where a.人员id = b.人员id And b.部门id = c.Id And d.Id = a.人员id And A.用户名=e.Grantee(+) And b.部门id =[1] And e.Granted_Role(+) = [2] " & _
            "And a.用户名<>[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-用户角色", mlngDept, "ZL_" & Me.lstRole.Text, gstrDbUser)
    
    Me.lvwGrant.ListItems.Clear
    Me.lvwRemove.ListItems.Clear
    Do While Not rsTmp.EOF
        If rsTmp!权限 = 1 Then
            Set lstTmp = lvwRemove.ListItems.Add(, rsTmp!用户名, rsTmp!姓名, , "YES")
            lstTmp.SubItems(1) = rsTmp!用户名
        Else
            Set lstTmp = lvwGrant.ListItems.Add(, rsTmp!用户名, rsTmp!姓名, , "NO")
            lstTmp.SubItems(1) = rsTmp!用户名
        End If
        lstTmp.Checked = True
        rsTmp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub FillSystem()
    Dim rsTemp As New ADODB.Recordset
    Dim strSystem As String
    Dim i As Integer
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    '显示可以所有的系统
'    gstrSQL = "Select 编号, 名称, 共享号, 所有者, 安装日期, 正常安装, 版本号 From zlSystems order by 编号"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-已安装系统")
    cboSystem.Clear
    
    For i = 1 To lstRole.ListCount
        gstrSQL = "Select distinct M.编号, M.名称, M.共享号, M.所有者, M.安装日期,M.正常安装, M.版本号 " & _
                  "  From zlRoleGrant R, zlPrograms P,zlsystems M" & _
                  "  Where Nvl(r.系统, 0) = Nvl(p.系统, 0) And p.序号 = r.序号 And p.序号 >= 100 And Substr(r.角色, 4) = [1] and m.编号=p.系统" & _
                    "  Order By m.编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-已安装系统", lstRole.List(i))
        Do Until rsTemp.EOF
            If InStr(1, strSystem, rsTemp!编号 & "|") = 0 Then
                cboSystem.AddItem rsTemp("名称") & " v" & rsTemp("版本号") & "（" & rsTemp("编号") & "）"
                strSystem = strSystem & "|" & rsTemp!编号 & "|"
                cboSystem.ItemData(cboSystem.NewIndex) = rsTemp("编号")
                If rsTemp("所有者") = UCase(gstrUserName) And cboSystem.ListIndex < 0 Then
                    cboSystem.ListIndex = cboSystem.NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
    Next
    
    '有两种系统是程序固定的
    If (zlRegTool And 2) = 2 Then cboSystem.AddItem "自定义报表"
    cboSystem.AddItem "基础工具"
    cboSystem.AddItem "取数函数"
    cboSystem.AddItem "基础编码"
    If cboSystem.ListIndex < 0 Then cboSystem.ListIndex = 0
    Exit Sub

errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
'    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Sub lstRole_Click()
    LoadPreson
    FillModule
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then Exit Sub
    
    If Me.txtEdit.Text = "" Then Exit Sub
    
    zlControl.TxtSelAll txtEdit
    
    For i = 0 To Me.lstRole.ListCount
        If InStr(1, Me.lstRole.List(i), Me.txtEdit.Text) > 0 Then
            Me.lstRole.Selected(i) = True
            Exit Sub
        End If
    Next
End Sub

Private Sub FillModule()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strRole As String
    
'    LockWindowUpdate lvwModule.hwnd
    On Error GoTo errHandle
    lvwModule.ColumnHeaders.Clear
    lvwModule.ListItems.Clear
    If lstRole.ListIndex <> -1 Then
        strRole = lstRole.List(lstRole.ListIndex)
    End If
    '更新列表项
    With lvwModule.ColumnHeaders
        If cboSystem.Text = "基础编码" Then
'            lblModule.Caption = "可管理的编码表"
            .Add , , "编码表", "1200"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cboSystem.Text = "取数函数" Then
'            lblModule.Caption = "可调用的函数"
            .Add , , "函数名", "1200"
            .Add , , "中文名", "1500"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cboSystem.Text = "基础工具" Then
'            lblModule.Caption = "已授权的基础工具"
            .Add , , "序号", "600"
            .Add , , "标题", "1800"
            .Add , , "说明", "3000"
'            .Add , , "授权功能", "5000"
        Else
'            lblModule.Caption = "已授权模块"
            .Add , , "序号", "600"
            .Add , , "标题", "1800"
            .Add , , "说明", "3000"
'            .Add , , "授权功能", "5000"
        End If
    End With
'    lnModuel.X1 = lblModule.Left + lblModule.Width
    
    If strRole = "" Then
        '角色为空，退出
'        LockWindowUpdate 0
        Exit Sub
    End If

    If cboSystem.Text = "基础编码" Then
        '显示该角色能访问的基础表
        gstrSQL = "select T.系统,T.表名,T.说明 from " & _
                "(SELECT S.名称||'（'||S.编号||'）' as 系统,S.所有者,B.表名,B.说明 FROM zlSystems S,zlBaseCode B where B.系统=S.编号) T,USER_TAB_PRIVS R " & _
                "where T.所有者=R.OWNER AND T.表名=R.TABLE_NAME AND R.GRANTEE='" & strRole & _
                "' and R.PRIVILEGE in ('SELECT','INSERT','UPDATE','DELETE') " & _
                "GROUP BY T.系统,T.表名,T.说明 " & _
                "Having Count(R.PRIVILEGE) = 4"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("表名"))
            lst.SubItems(1) = rsTemp("系统")
            lst.SubItems(2) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            rsTemp.MoveNext
        Loop
    ElseIf cboSystem.Text = "取数函数" Then
        '显示该角色能访问的基础表
        gstrSQL = "select S.名称||'（'||S.编号||'）' as 系统,S.所有者,F.函数名,F.中文名,F.说明 " & _
                  " from zlSystems S,zlFunctions F,USER_TAB_PRIVS R " & _
                  " where  F.系统=S.编号 and S.所有者=R.OWNER AND upper(F.函数名)=R.TABLE_NAME AND R.GRANTEE='" & strRole & "' and R.PRIVILEGE ='EXECUTE'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("函数名"))
            lst.SubItems(1) = rsTemp("中文名")
            lst.SubItems(2) = rsTemp("系统")
            lst.SubItems(3) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            rsTemp.MoveNext
        Loop
    ElseIf cboSystem.Text = "基础工具" Then
        '显示该角色能访问的基础工具
        gstrSQL = "select P.序号,P.标题,P.说明,R.功能 from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where R.系统 is null and P.序号=R.序号 AND substr(R.角色,4)='" & strRole & _
                "'  AND P.系统 is null and P.序号<100 and P.部件 is null " & _
                " Order By P.序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), rsTemp("序号"))
            If Err <> 0 Then
                Err.Clear
                If rsTemp("功能") <> "基本" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
'                    lst.SubItems(3) = IIF(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("功能")
                End If
            Else
                lst.SubItems(1) = rsTemp("标题")
                lst.SubItems(2) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
'                If rsTemp("功能") <> "基本" Then
'                    lst.SubItems(3) = rsTemp("功能")
'                End If
            End If
            rsTemp.MoveNext
        Loop
    Else
        '显示该角色能访问的模块
        gstrSQL = "select P.序号,P.标题,P.说明,R.功能 from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where nvl(R.系统,0)=nvl(P.系统,0) and P.序号=R.序号 and P.序号>=100 AND substr(R.角色,4)='" & strRole & "'  AND " & _
                IIF(cboSystem.Text = "自定义报表", " P.系统 is null ", " P.系统=" & cboSystem.ItemData(cboSystem.ListIndex)) & _
                " Order By P.序号"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), rsTemp("序号"))
            If Err <> 0 Then
                Err.Clear
                If rsTemp("功能") <> "基本" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
'                    lst.SubItems(3) = IIF(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("功能")
                End If
            Else
                lst.SubItems(1) = rsTemp("标题")
                lst.SubItems(2) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
'                If rsTemp("功能") <> "基本" Then
'                    lst.SubItems(3) = rsTemp("功能")
'                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
'    LockWindowUpdate 0
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
