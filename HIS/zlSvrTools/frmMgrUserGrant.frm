VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMgrUserGrant 
   Caption         =   "管理工具授权"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "frmMgrUserGrant.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   8790
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   3930
      Picture         =   "frmMgrUserGrant.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   4500
      Picture         =   "frmMgrUserGrant.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwGranted 
      Height          =   3800
      Left            =   5040
      TabIndex        =   9
      Top             =   960
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   6694
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwNoGrant 
      Height          =   3800
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   6694
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找用户(&F)"
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   65
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4800
      TabIndex        =   1
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6150
      TabIndex        =   3
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7425
      TabIndex        =   4
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -360
      TabIndex        =   0
      Top             =   525
      Width           =   10110
   End
   Begin MSComctlLib.ImageList Img16 
      Left            =   3975
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1A5E
            Key             =   "自动提醒"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":82C0
            Key             =   "数据连接"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EB22
            Key             =   "操作日志管理"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":15384
            Key             =   "功能限时管理"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1BBE6
            Key             =   "系统装卸管理"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":22448
            Key             =   "数据转移"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":28CAA
            Key             =   "用户注册管理"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":2F50C
            Key             =   "系统升迁管理"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":35D6E
            Key             =   "系统参数管理"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":3C5D0
            Key             =   "运行日志管理"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":42E32
            Key             =   "错误日志管理"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":49694
            Key             =   "系统运行选项"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":4FEF6
            Key             =   "对象检查修复"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":56758
            Key             =   "数据导出"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":5CFBA
            Key             =   "站点文件收集"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6381C
            Key             =   "编译无效对象"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6A07E
            Key             =   "后台作业管理"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":708E0
            Key             =   "数据导入"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":77142
            Key             =   "数据调入"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":7D9A4
            Key             =   "数据清除"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":84206
            Key             =   "数据调出"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":8AA68
            Key             =   "运行状态监控"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":912CA
            Key             =   "置换安装脚本"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":97B2C
            Key             =   "站点部件升级"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":9E38E
            Key             =   "报表管理"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":A4BF0
            Key             =   "函数管理"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":AB452
            Key             =   "用户授权管理"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B1CB4
            Key             =   "角色授权管理"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B8516
            Key             =   "菜单重组规划"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":BED78
            Key             =   "客户端运行控制"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C55DA
            Key             =   "权限管理"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C5EB4
            Key             =   "装卸管理"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C678E
            Key             =   "数据管理"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7068
            Key             =   "运行管理"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7602
            Key             =   "专项工具"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7EDC
            Key             =   "DBA工具"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":CE73E
            Key             =   "空间管理"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":D4FA0
            Key             =   "SQL性能"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":DB802
            Key             =   "会话解锁"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":E2064
            Key             =   "外键索引"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":E88C6
            Key             =   "SQL跟踪"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EF128
            Key             =   "数据库性能"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFunc 
      Height          =   1710
      Left            =   225
      TabIndex        =   12
      Top             =   4875
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   3016
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "功能"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "说明"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "缺省"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "对“梁唐彬”进行授权处理。"
      Height          =   180
      Left            =   960
      TabIndex        =   7
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgOne 
      Height          =   480
      Left            =   300
      Picture         =   "frmMgrUserGrant.frx":F598A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblModul 
      AutoSize        =   -1  'True
      Caption         =   "可授权模块(&A)"
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   660
      Width           =   1170
   End
   Begin VB.Label lblGranted 
      AutoSize        =   -1  'True
      Caption         =   "已授权模块(&G)"
      Height          =   180
      Left            =   4935
      TabIndex        =   5
      Top             =   660
      Width           =   1170
   End
End
Attribute VB_Name = "frmMgrUserGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrProg As String
Private mstrAccount As String '为空表示新用户授权
Private mblnOk As Boolean
Private mrsProgFuncs As ADODB.Recordset
Private mblnIsChange As Boolean '记录界面是否发生了修改

Private Enum LvwFuncList
    LFL_功能 = 0
    LFL_说明 = 1
    LFL_缺省 = 2
End Enum

Public Function GrantToProg(ByVal strAccount As String, ByVal strUser As String, ByVal strProg As String) As Boolean
    mstrUser = strUser
    mstrAccount = strAccount
    mstrProg = strProg
    mblnOk = False
    Me.Show 1
    GrantToProg = mblnOk
End Function

Private Sub cmdCancel_Click()
    If mblnIsChange Then
        If MsgBox("该人员功能权限信息已被更改，确定要放弃更改并退出吗？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
    Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click()
    Call FindPersonnel
End Sub

Private Sub MoveProg(objMoveIn As TreeView, objMoveOut As TreeView)
    Dim i As Long, y As Long
    Dim strDel As String, Node As Node
    
    For i = objMoveOut.Nodes.Count To 1 Step -1
        err = 0
        On Error Resume Next
        If objMoveOut.Nodes(i).Checked And Not objMoveOut.Nodes(i).Parent Is Nothing Then
            mblnIsChange = True
            If err = 0 Then
                err = 0
                If objMoveIn.Nodes(objMoveOut.Nodes(i).Parent.Key).Key <> "" Then
                    If err <> 0 Then
                        '新增父项
                        Set Node = objMoveIn.Nodes.Add(, , objMoveOut.Nodes(i).Parent.Key, objMoveOut.Nodes(i).Parent.Text, objMoveOut.Nodes(i).Parent.Image, objMoveOut.Nodes(i).Parent.SelectedImage)
                        Node.Expanded = objMoveOut.Nodes(i).Parent.Expanded
                        Node.Checked = objMoveOut.Nodes(i).Parent.Checked
                        Node.ForeColor = objMoveOut.Nodes(i).Parent.ForeColor
                    End If
                     '新增子项
                    Set Node = objMoveIn.Nodes.Add(objMoveOut.Nodes(i).Parent.Key, tvwChild, objMoveOut.Nodes(i).Key, objMoveOut.Nodes(i).Text, objMoveOut.Nodes(i).Image, objMoveOut.Nodes(i).SelectedImage)
                    Node.Expanded = objMoveOut.Nodes(i).Expanded
                    Node.Checked = objMoveOut.Nodes(i).Checked
                    Node.ForeColor = objMoveOut.Nodes(i).ForeColor
                    '删除子项
                    If objMoveOut.Nodes(i).Parent.Children = 1 Then
                        objMoveOut.Nodes.Remove objMoveOut.Nodes(i).Parent.Index
                    Else
                        objMoveOut.Nodes.Remove i
                    End If
                    
                End If
                On Error GoTo 0
            End If
        End If
    Next
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        Call MoveProg(tvwGranted, tvwNoGrant)
    ElseIf Index = 1 Then
        Call MoveProg(tvwNoGrant, tvwGranted)
    End If
End Sub

Private Sub cmdOK_Click()
'功能：授权
    Dim i As Integer, j As Integer
    Dim strProg As String, strFunc As String, strKey As String
    Dim StrJiami() As Byte
    Dim strPwText As String
    Dim rsTemp As New ADODB.Recordset
    
    If mstrAccount = "" Then
        MsgBox "请先查找需要授权的用户。", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    '组装功能字符串
    For i = 1 To tvwGranted.Nodes.Count
        If Not tvwGranted.Nodes(i).Parent Is Nothing Then
            strKey = Mid(tvwGranted.Nodes(i).Key, 2)
            mrsProgFuncs.Filter = "编号 = '" & strKey & "' And 权限 = 1"
            strFunc = ""
            Do While Not mrsProgFuncs.EOF
                strFunc = strFunc & "|" & mrsProgFuncs!功能
                mrsProgFuncs.MoveNext
            Loop
            If strFunc <> "" Then
                strFunc = ":" & "基本" & "|" & Mid(strFunc, 2)
            Else
                strFunc = ":" & "基本"
            End If
            strProg = strProg & "," & strKey & strFunc
        End If
    Next
    strProg = Mid(strProg, 2)
    '功能加密
    If strProg <> "" Then
        Call DES_Encode(StrConv(strProg, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("单位名称", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    On Error GoTo errHandle
    gstrSQL = "Select 1 From zlMgrGrant Where 用户名='" & mstrAccount & "'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount > 0 Then
        If strPwText = "" Then
            gstrSQL = "Delete zlMgrGrant Where 用户名='" & mstrAccount & "'"
        Else
            gstrSQL = "Update zlMgrGrant Set 功能='" & strPwText & "' Where 用户名='" & mstrAccount & "'"
        End If
    Else
        gstrSQL = "Insert into zlMgrGrant(用户名,功能) values('" & mstrAccount & "','" & strPwText & "')"
    End If
    gcnOracle.Execute gstrSQL
    '更新管理员账户信息
    rsTemp.Close
    '未授权程序不更新管理员信息
    If Not gstrPassword Like "未授权的程序:*" Then
        gstrSQL = "Select 1 From zlRegInfo where 项目='管理员'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
         If rsTemp.RecordCount > 0 Then
            gstrSQL = "Update zlRegInfo Set 内容='" & gstrUserName & "' Where 项目='管理员'"
        Else
            gstrSQL = "Insert into zlRegInfo(项目,内容) values('管理员','" & gstrUserName & "')"
        End If
        gcnOracle.Execute gstrSQL
        '验证码
        strPwText = ""
        ReDim Preserve StrJiami(0)
        If gstrPassword <> "" Then
            Call DES_Encode(StrConv(gstrPassword, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("单位名称", False, 0))
            strPwText = FuncByteTo16Code(StrJiami)
        End If
        rsTemp.Close
        gstrSQL = "Select 1 From zlRegInfo where 项目='验证码'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
         If rsTemp.RecordCount > 0 Then
            gstrSQL = "Update zlRegInfo Set 内容='" & strPwText & "' Where 项目='验证码'"
        Else
            gstrSQL = "Insert into zlRegInfo(项目,内容) values('验证码','" & strPwText & "')"
        End If
        gcnOracle.Execute gstrSQL
    End If
    mblnOk = True
    Unload Me
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Load()
    If mstrAccount = "" Then
        lblNote.Caption = "请先输入用户名、人员姓名或简码。"
        txtFind.Visible = True
        cmdFind.Visible = True
    Else
        lblNote.Caption = "正在对""" & mstrUser & """进行管理工具授权。"
        txtFind.Visible = False
        cmdFind.Visible = False
    End If
    Call InitProgFuncData
    Call FillProg
End Sub

'初始化模块功能信息到一个记录集里
Private Sub InitProgFuncData()
    Dim rsTemp As ADODB.Recordset
    Dim strProg As String
    Dim arrProg() As String
    Dim arrFunc() As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errh
    '查询出所有可以授权的模块及功能
    gstrSQL = "Select a.编号, a.标题, a.上级, b.功能, b.缺省, b.排列, 0 权限, b.说明" & vbNewLine & _
            "From Zlsvrtools a, Zlsvrfuncs b" & vbNewLine & _
            "Where a.编号 = b.序号(+) And a.编号 <> '0404'" & vbNewLine & _
            "Order By a.编号, b.排列"
    Set mrsProgFuncs = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption))
    '更新权限字段
    '解析mstrProg，从而获取选中人员所拥有的模块及功能的授权情况
    arrProg = Split(mstrProg, ",")
    For i = 0 To UBound(arrProg)
        strProg = Split(arrProg(i), ":")(0)
        Call RecUpdate(mrsProgFuncs, "编号 = '" & strProg & "' And 功能 = '基本'", "权限", 1)
        arrFunc = Split(Split(arrProg(i) & ":", ":")(1), "|")
        If UBound(arrFunc) = -1 Then
            '对于原来的用户，只对模块进行了授权，并未对功能进行授权，故功能字符串肯定为空，这时默认勾选所有功能
            Call RecUpdate(mrsProgFuncs, "编号 = '" & strProg & "'", "权限", 1)
        Else
        For j = 0 To UBound(arrFunc)
            Call RecUpdate(mrsProgFuncs, "编号 = '" & strProg & "' And 功能 = '" & arrFunc(j) & "'", "权限", 1)
        Next
        End If
    Next
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillProg()
'功能：填充功能
    Dim strProg As String, Node As Node
    Dim i As Long
    
    On Error GoTo errHandle
    '显示该用户具有的角色
    mrsProgFuncs.Filter = "功能 = '基本' Or 功能 = Null"

    Do Until mrsProgFuncs.EOF
        With IIf(mrsProgFuncs!权限 = 0, tvwNoGrant, tvwGranted)
            '上级两边都加
            If IsNull(mrsProgFuncs("上级")) Then
                Set Node = tvwNoGrant.Nodes.Add(, , "D" & mrsProgFuncs("编号"), "【" & mrsProgFuncs("编号") & "】" & mrsProgFuncs("标题"))
                tvwNoGrant.Nodes("D" & mrsProgFuncs("编号")).Sorted = True
                tvwNoGrant.Nodes("D" & mrsProgFuncs("编号")).Expanded = True
                tvwNoGrant.Nodes("D" & mrsProgFuncs("编号")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!标题 & "").Index
                err.Clear: On Error GoTo errHandle
                Set Node = tvwGranted.Nodes.Add(, , "D" & mrsProgFuncs("编号"), "【" & mrsProgFuncs("编号") & "】" & mrsProgFuncs("标题"))
                tvwGranted.Nodes("D" & mrsProgFuncs("编号")).Sorted = True
                tvwGranted.Nodes("D" & mrsProgFuncs("编号")).Expanded = True
                tvwGranted.Nodes("D" & mrsProgFuncs("编号")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!标题 & "").Index
                err.Clear: On Error GoTo errHandle
            Else
                Set Node = .Nodes.Add("D" & mrsProgFuncs("上级"), tvwChild, "C" & mrsProgFuncs("编号"), mrsProgFuncs("标题"))
                .Nodes("C" & mrsProgFuncs("编号")).Sorted = True
                Node.Checked = False
                mrsProgFuncs.Update "权限", 0
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!标题 & "").Index
                err.Clear: On Error GoTo errHandle
            End If
        End With
        mrsProgFuncs.MoveNext
    Loop
    '删除没有子项的分类
    For i = tvwNoGrant.Nodes.Count To 1 Step -1
        If tvwNoGrant.Nodes(i).Children = 0 And tvwNoGrant.Nodes(i).Parent Is Nothing Then
            tvwNoGrant.Nodes.Remove i
        End If
    Next
    For i = tvwGranted.Nodes.Count To 1 Step -1
        If tvwGranted.Nodes(i).Children = 0 And tvwGranted.Nodes(i).Parent Is Nothing Then
            tvwGranted.Nodes.Remove i
        End If
    Next
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

'填充模块对应的功能，基本功能不再显示，为默认授予
Private Sub FillFunction(ByVal strPorgNo As String)
    Dim lst As ListItem
    
    On Error GoTo errh
    lvwFunc.ListItems.Clear
    mrsProgFuncs.Filter = "编号 = '" & strPorgNo & "' And 功能 <> '基本'"
    '填充功能信息及勾选情况
    With mrsProgFuncs
        Do While Not .EOF
            Set lst = lvwFunc.ListItems.Add(, !编号 & "_" & !排列, !功能)
            lst.SubItems(LFL_说明) = !说明 & ""
            lst.SubItems(LFL_缺省) = !缺省 & ""
            lst.Checked = IIf(!权限 = 1, True, False)
            .MoveNext
        Loop
    End With
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If mblnIsChange Then
            If MsgBox("该人员功能权限信息已被更改，确定要放弃更改并退出吗？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                Cancel = 1
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraLine.Width = Me.Width
    cmdOK.Move Me.Width - cmdCancel.Width - cmdOK.Width - 400, Me.Height - cmdOK.Height - 650
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 100, cmdOK.Top
    lvwFunc.Width = Me.Width - 600
    lvwFunc.Top = cmdOK.Top - lvwFunc.Height - 100
    If lvwFunc.ListItems.Count > 5 Then
        lvwFunc.ColumnHeaders(2).Width = lvwFunc.Width - lvwFunc.ColumnHeaders(1).Width - 250
    Else
        lvwFunc.ColumnHeaders(2).Width = lvwFunc.Width - lvwFunc.ColumnHeaders(1).Width
    End If
    tvwNoGrant.Width = Me.Width \ 2 - 885
    tvwGranted.Width = Me.Width \ 2 - 885
    tvwNoGrant.Height = lvwFunc.Top - tvwNoGrant.Top - 100
    tvwGranted.Height = tvwNoGrant.Height
    tvwGranted.Left = tvwNoGrant.Left + tvwNoGrant.Width + 1185
    
    cmdMove(1).Left = tvwNoGrant.Left + tvwNoGrant.Width + 150
    cmdMove(0).Left = cmdMove(1).Left + cmdMove(1).Width + 150
    lblGranted.Left = tvwGranted.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsProgFuncs = Nothing
            mblnIsChange = False
End Sub

Private Sub lvwFunc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '若修改了一项，则更新记录集
    mblnIsChange = True
    Call RecUpdate(mrsProgFuncs, "编号 = '" & Split(Item.Key, "_")(0) & "' And 功能 = '" & Item.Text & "'", "权限", IIf(Item.Checked, 1, 0))
End Sub

Private Sub tvwGranted_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Call tvwGranted_NodeClick(Node)
    Call NodeCheckMode(Node, tvwGranted)
End Sub

Private Sub tvwGranted_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillFunction(Mid(Node.Key, 2))
    If Node = tvwGranted.SelectedItem Then
        lvwFunc.Enabled = True
        lvwFunc.BackColor = &H80000005
    End If
End Sub

Private Sub tvwNoGrant_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Call tvwNoGrant_NodeClick(Node)
    Call NodeCheckMode(Node, tvwNoGrant)
    If Node = tvwNoGrant.SelectedItem Then
        If Node.Checked = False Then
           lvwFunc.Enabled = False
           lvwFunc.BackColor = &H8000000F
        Else
           lvwFunc.Enabled = True
           lvwFunc.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub NodeCheckMode(ByRef Node As MSComctlLib.Node, ByRef objtvwThis As TreeView)
'功能：让树表选中父节点，自动选中所有子节点，选中所有子节点，父节点也选中
    Dim i As Long
    Dim blnIsNothing As Boolean
    
    LockWindowUpdate objtvwThis.hwnd
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Key Then
                    objtvwThis.Nodes(i).Checked = Node.Checked
                End If
            End If
        Next
    Else
        For i = Node.Parent.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Parent.Key Then
                    If Not objtvwThis.Nodes(i).Checked = Node.Checked Then blnIsNothing = True
                End If
            End If
        Next
        '若勾选的报表是当前选择项，则把lvwFunc中的为缺省的项也勾选上
        If Node = objtvwThis.SelectedItem Then
            For i = 1 To lvwFunc.ListItems.Count
                If Node.Checked = False Then
                    lvwFunc.ListItems.Item(i).Checked = False
                    Call lvwFunc_ItemCheck(lvwFunc.ListItems.Item(i))
                ElseIf lvwFunc.ListItems.Item(i).SubItems(LFL_缺省) = "1" Then
                    lvwFunc.ListItems.Item(i).Checked = True
                    Call lvwFunc_ItemCheck(lvwFunc.ListItems.Item(i))
                End If
            Next
        End If
        If blnIsNothing Then
            Node.Parent.Checked = False
        Else
            Node.Parent.Checked = Node.Checked
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub tvwNoGrant_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Parent Is Nothing Then Exit Sub
    Call FillFunction(Mid(Node.Key, 2))
    If Node.Checked = False Then
        lvwFunc.Enabled = False
        lvwFunc.BackColor = &H8000000F
    Else
        lvwFunc.Enabled = True
        lvwFunc.BackColor = &H80000005
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindPersonnel
    End If
End Sub

Private Sub FindPersonnel()
'功能：查找人员
    Dim rsTemp As New Recordset
    Dim objPoint As POINTAPI
    
    If txtFind.Text = "" Then Exit Sub
    gstrSQL = "Select b.用户名, c.姓名, c.简码, d.名称 As 部门名称" & vbNewLine & _
            "From  Zlmgrgrant A,上机人员表 B, 人员表 C, 部门表 D, 部门人员 E" & vbNewLine & _
            "Where a.用户名(+) = b.用户名 And b.人员id = c.Id And c.Id = e.人员id And d.Id = e.部门id And A.用户名 is null And e.缺省 = 1 And B.用户名 <> '" & gstrUserName & "'" & _
            " And(b.用户名 like '" & UCase(Trim(txtFind.Text)) & "%' Or c.姓名 Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.简码 Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.编号=' & UCase(Trim(txtFind.Text)) & ')" & _
            " Order By c.姓名"
    Set rsTemp = New ADODB.Recordset
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "您查找的用户不存在，或是已经拥有了权限，请检查。", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        Exit Sub
    End If
    Call ClientToScreen(txtFind.hwnd, objPoint)
    
    If frmSelectList.ShowSelect(Nothing, rsTemp, "用户名,900,0,1;姓名,900,0,1;简码,650,0,0;部门名称,1500,0,1", objPoint.x * 15 - 30, objPoint.y * 15 + cmdFind.Height - 30, txtFind.Width + cmdFind.Width + 1300, 3000, "", "查找人员", , , True) = False Then
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        rsTemp.Filter = 0
        Exit Sub
    Else
        txtFind.Text = rsTemp!姓名 & ""
        mstrAccount = rsTemp!用户名 & ""
        mstrUser = rsTemp!姓名 & ""
    End If
End Sub
