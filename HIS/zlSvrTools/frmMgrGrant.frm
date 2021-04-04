VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMgrGrant 
   BackColor       =   &H80000005&
   Caption         =   "管理工具授权"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmMgrGrant.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   7410
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新用户授权(&A)"
      Height          =   350
      Left            =   5880
      TabIndex        =   1
      ToolTipText     =   "搜索框内输入内容后可查找到没有权限的人员。"
      Top             =   720
      Width           =   1365
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "调整授权&G)"
      Height          =   350
      Left            =   5880
      TabIndex        =   2
      Top             =   1080
      Width           =   1365
   End
   Begin MSComctlLib.ListView lvwProg 
      Height          =   2145
      Left            =   945
      TabIndex        =   4
      Top             =   3540
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   3784
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImgBig"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "模块"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "授权功能"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   2070
      Left            =   945
      TabIndex        =   0
      Top             =   1110
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3651
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "ImgSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Settlement"
         Text            =   "用户名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "人员编号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "人员姓名"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "所属部门"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "功能"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBig 
      Left            =   5610
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":803A
            Key             =   "自动提醒"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":8914
            Key             =   "系统装卸管理"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":91EE
            Key             =   "数据转移"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":9AC8
            Key             =   "用户注册管理"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":A3A2
            Key             =   "系统升迁管理"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":AC7C
            Key             =   "系统参数管理"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":B556
            Key             =   "运行日志管理"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":BE30
            Key             =   "错误日志管理"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":C70A
            Key             =   "系统运行选项"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":CFE4
            Key             =   "对象检查修复"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":D8BE
            Key             =   "数据导出"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":E198
            Key             =   "站点文件收集"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":EA72
            Key             =   "编译无效对象"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":F34C
            Key             =   "后台作业管理"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":FC26
            Key             =   "数据导入"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":10500
            Key             =   "数据调入"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":10DDA
            Key             =   "数据清除"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":116B4
            Key             =   "数据调出"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":11F8E
            Key             =   "运行状态监控"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":12868
            Key             =   "置换安装脚本"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":13142
            Key             =   "站点部件升级"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":13A1C
            Key             =   "报表管理"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":142F6
            Key             =   "函数管理"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":14BD0
            Key             =   "管理工具授权"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":154AA
            Key             =   "用户授权管理"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":15D84
            Key             =   "角色授权管理"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":1C5E6
            Key             =   "菜单重组规划"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":1CEC0
            Key             =   "站点运行控制"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":1D79A
            Key             =   "空间管理"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":23FFC
            Key             =   "外键索引"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":2A85E
            Key             =   "会话解锁"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":310C0
            Key             =   "数据库性能"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":37922
            Key             =   "SQL跟踪"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":3E184
            Key             =   "DBA工具"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":449E6
            Key             =   "SQL性能"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgSmall 
      Left            =   6120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":4B248
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":4CFDA
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":4DCB4
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":54516
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":5AD78
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   6720
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":615DA
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":6336C
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":64046
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":6A8A8
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrGrant.frx":7110A
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000005&
      Caption         =   "已授权用户："
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "管理工具授权"
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
      Left            =   175
      TabIndex        =   5
      Top             =   125
      Width           =   1440
   End
   Begin VB.Label lblProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "授权功能"
      Height          =   180
      Left            =   945
      TabIndex        =   3
      Top             =   3315
      Width           =   720
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmMgrGrant.frx":7796C
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmMgrGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstr所有者 As String '保存当前系统的所有者名

Private mrsUsers As ADODB.Recordset
Private mLastIndex As Long
Private Enum LvwProgList
    LPL_序号 = 0
    LPL_模块 = 1
    LPL_授权功能 = 2
End Enum


Private Sub cmdAdd_Click()
    If frmMgrUserGrant.GrantToProg("", "", "") = True Then
        Call FillUser
    End If
End Sub

Private Sub cmdGrant_Click()
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    If frmMgrUserGrant.GrantToProg(lvwUser.SelectedItem.Text, lvwUser.SelectedItem.SubItems(2), lvwUser.SelectedItem.SubItems(4)) = True Then
        Call FillUser
    End If
End Sub

Private Sub Form_Load()
   
    If gblnDBA Then
        Me.lvwUser.ColumnHeaders(4).Width = 1000
    Else
        Me.lvwUser.ColumnHeaders(4).Width = 0
    End If
    Call FillUser
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    
    err = 0: On Error Resume Next
    Me.cmdGrant.Left = Me.ScaleWidth - 200 - Me.cmdGrant.Width
    Me.cmdAdd.Left = Me.ScaleWidth - 200 - Me.cmdGrant.Width
    Me.lvwUser.Width = Me.cmdGrant.Left - 90 - Me.lvwUser.Left
    Me.lvwProg.Width = Me.ScaleWidth - Me.lvwProg.Left - 200
    If lvwProg.ListItems.Count > 9 Then
        Me.lvwProg.ColumnHeaders(3).Width = Me.lvwProg.Width - Me.lvwProg.ColumnHeaders(1).Width - _
                                            Me.lvwProg.ColumnHeaders(2).Width - 250
    Else
        Me.lvwProg.ColumnHeaders(3).Width = Me.lvwProg.Width - Me.lvwProg.ColumnHeaders(1).Width - _
                                            Me.lvwProg.ColumnHeaders(2).Width
    End If
    lngTemp = (Me.ScaleHeight - lvwProg.Height - lblProg.Height - 800) - lvwUser.Top
    lvwUser.Height = IIf(lngTemp > 400, lngTemp, 400)
    lblProg.Top = lvwUser.Top + lvwUser.Height + 100
    lvwProg.Top = lblProg.Top + lblProg.Height + 100
End Sub

Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mLastIndex = Item.Index
    Call FillProg
    If lvwProg.ListItems.Count > 9 Then
        Me.lvwProg.ColumnHeaders(3).Width = Me.lvwProg.Width - Me.lvwProg.ColumnHeaders(1).Width - _
                                            Me.lvwProg.ColumnHeaders(2).Width - 250
    Else
        Me.lvwProg.ColumnHeaders(3).Width = Me.lvwProg.Width - Me.lvwProg.ColumnHeaders(1).Width - _
                                            Me.lvwProg.ColumnHeaders(2).Width
    End If
End Sub

Private Sub FillUser()
'功能：填充用户
'参数：blnFilter是否根据已有的记录集查找,blnNoHave查找没有权限的人员
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strIco As String
    Dim blnOwner As Boolean     '所有者
    Dim str人员姓名 As String
    '显示可以进行当前系统的用户与对应的人员
    Dim strSearch As String
    Dim strSource() As Byte
    Dim strDest() As Byte
    Dim StrJiemi() As Byte
    
    On Error GoTo errHandle
    gstrSQL = "Select a.用户名 As Username, c.编号, c.姓名, c.简码 As 人员简码, d.编码 As 部门编码, d.名称 As 部门名称, d.简码 As 部门简码,A.功能" & vbNewLine & _
            "From Zlmgrgrant A, 上机人员表 B, 人员表 C, 部门表 D, 部门人员 E" & vbNewLine & _
            "Where a.用户名 = b.用户名 And b.人员id = c.Id And c.Id = e.人员id And d.Id = e.部门id And e.缺省 = 1" & _
            " Order By c.姓名"
    Set rsTemp = New ADODB.Recordset
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    Set mrsUsers = rsTemp
        
    lvwUser.ListItems.Clear
    
    Do Until rsTemp.EOF
        str人员姓名 = Nvl(rsTemp!姓名)
        If rsTemp("功能") & "" <> "" Then
            strIco = "User"
        Else
            strIco = "UserInfor"
        End If
        Set lst = lvwUser.ListItems.Add(, "C" & rsTemp("USERNAME"), rsTemp("USERNAME"), strIco, strIco)
        lst.SubItems(1) = IIf(IsNull(rsTemp("编号")), "", rsTemp("编号"))
        lst.SubItems(2) = str人员姓名
        lst.SubItems(3) = IIf(IsNull(rsTemp("部门名称")), "", rsTemp("部门名称"))
        ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
        Call Func16CodeToByte(rsTemp("功能") & "", strDest)
        If rsTemp("功能") & "" <> "" Then Call DES_Decode(strDest, StrJiemi, gobjRegister.zlRegInfo("单位名称", False, 0))
        lst.SubItems(4) = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
        rsTemp.MoveNext
    Loop
    If lvwUser.ListItems.Count > 0 Then
        If mLastIndex > 0 And mLastIndex < lvwUser.ListItems.Count Then
            lvwUser.ListItems(mLastIndex).Selected = True
        Else
            lvwUser.ListItems(1).Selected = True
        End If
        Call FillProg
    End If
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub FillProg()
'功能：填充功能
    Dim rsTemp As New ADODB.Recordset
    Dim strProg As String, objItem As ListItem
    Dim strFunc As String
    Dim arrProg() As String
    Dim i As Long
    
    On Error GoTo errHandle
    lvwProg.ListItems.Clear
    If lvwUser.SelectedItem Is Nothing Then
        Exit Sub
    Else
        strProg = lvwUser.SelectedItem.SubItems(4)
    End If
    '先把模块中的功能剔除掉，方便查询模块信息
    arrProg = Split(strProg, ",")
    strProg = ""
    For i = 0 To UBound(arrProg)
        strProg = strProg & "," & Mid(arrProg(i), 1, InStr(arrProg(i) & ":", ":") - 1)
    Next
    strProg = Mid(strProg, 2)
    
    '显示该用户具有的角色
    gstrSQL = "Select a.编号, a.标题, b.功能" & vbNewLine & _
            "From Zlsvrtools a, Zlsvrfuncs b," & vbNewLine & _
            "     (Select Column_Value From Table(Cast(f_Str2list('" & strProg & "') As Zltools.t_Strlist))) c" & vbNewLine & _
            "Where a.编号 = b.序号 And a.编号 = c.Column_Value" & vbNewLine & _
            "Order By a.编号, b.排列"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    For i = 0 To UBound(arrProg)
        rsTemp.Filter = "编号 = '" & Split(arrProg(i), ":")(0) & "'"
        If rsTemp.RecordCount > 0 Then
            Set objItem = lvwProg.ListItems.Add(, , rsTemp("编号"))
            objItem.SubItems(LPL_模块) = rsTemp("标题")
            strFunc = Split(arrProg(i) & ":", ":")(1)
            '为了兼容以前的用户，若功能字符串为空，则表示其有所有功能的权限
            If strFunc = "" Then
                Do While Not rsTemp.EOF
                    strFunc = strFunc & "," & rsTemp!功能
                    rsTemp.MoveNext
                Loop
                objItem.SubItems(LPL_授权功能) = Mid(strFunc, 2)
            Else
                objItem.SubItems(LPL_授权功能) = Replace(strFunc, "|", ",")
            End If
        End If
    Next
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function
