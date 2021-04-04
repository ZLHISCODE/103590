VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户信息编辑"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmUserEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.TreeView tvwPerson 
      Height          =   6210
      Left            =   0
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10954
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "Img小图标"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox picProcess 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   12600
      TabIndex        =   34
      Top             =   7155
      Visible         =   0   'False
      Width           =   12600
      Begin MSComctlLib.ProgressBar prg 
         Height          =   240
         Left            =   60
         TabIndex        =   35
         Top             =   285
         Width           =   12270
         _ExtentX        =   21643
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblProssCaption 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   180
         Left            =   180
         TabIndex        =   37
         Top             =   60
         Width           =   90
      End
      Begin VB.Label lblStep 
         Alignment       =   1  'Right Justify
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10695
         TabIndex        =   36
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.ComboBox cboWorkRange 
      Height          =   300
      Left            =   6315
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3975
      Width           =   1320
   End
   Begin VB.ComboBox cboManPros 
      Height          =   300
      Left            =   7725
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Frame fraApplyTo 
      Caption         =   "应用于尚未创建用户的人员性质为                                  的"
      Height          =   870
      Left            =   3450
      TabIndex        =   23
      Top             =   4035
      Width           =   7890
      Begin VB.OptionButton optApplyTo 
         Caption         =   "当前人员(&3)"
         Height          =   240
         Index           =   2
         Left            =   675
         TabIndex        =   26
         Top             =   405
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "所有人员(&5)"
         Height          =   240
         Index           =   1
         Left            =   3885
         TabIndex        =   28
         Top             =   405
         Width           =   1395
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "当前部门人员(&4)"
         Height          =   240
         Index           =   0
         Left            =   2145
         TabIndex        =   27
         Top             =   405
         Width           =   1785
      End
   End
   Begin VB.OptionButton optMode 
      Caption         =   "按人员编号设置用户名(&2)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   15
      ToolTipText     =   "如果编码中每一个为字符的,则直接用编码构建用户,否则为U+编码形式构建用户"
      Top             =   4500
      Width           =   2595
   End
   Begin VB.OptionButton optMode 
      Caption         =   "按姓名简码设置用户名(&1)"
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   14
      ToolTipText     =   "当应用于其他人员时,如果有存在相同的简码,则以简码+编号的形式构建用户"
      Top             =   4200
      Value           =   -1  'True
      Width           =   2595
   End
   Begin MSComctlLib.ImageList Img小图标 
      Left            =   5295
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":000C
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0326
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0640
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0F1A
            Key             =   "Module"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   11400
      TabIndex        =   31
      Top             =   4410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11400
      TabIndex        =   30
      Top             =   705
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   11400
      TabIndex        =   29
      Top             =   330
      Width           =   1100
   End
   Begin VB.Frame fraPerson 
      Caption         =   "对应人员"
      Height          =   2850
      Left            =   135
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2070
      Width           =   3210
      Begin VB.TextBox txtno 
         Height          =   300
         Left            =   675
         MaxLength       =   20
         TabIndex        =   38
         Top             =   292
         Width           =   1125
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Width           =   315
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   675
         TabIndex        =   13
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   675
         TabIndex        =   11
         Top             =   690
         Width           =   2175
      End
      Begin VB.Label lblDept 
         Caption         =   "部门"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1155
         Width           =   585
      End
      Begin VB.Label lblName 
         Caption         =   "姓名"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblNo 
         Caption         =   "编码"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   345
         Width           =   540
      End
   End
   Begin VB.Frame fraUser 
      Caption         =   "用户信息"
      Height          =   1755
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   3210
      Begin VB.TextBox txtVerify 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   1860
      End
      Begin VB.TextBox txtPasswd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   795
         Width           =   1860
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   2
         Top             =   405
         Width           =   1860
      End
      Begin VB.Label lblExamPwd 
         Caption         =   "确认密码"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label lblPasswd 
         AutoSize        =   -1  'True
         Caption         =   "密码"
         Height          =   180
         Left            =   585
         TabIndex        =   3
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblUserName 
         Caption         =   "用户名"
         Height          =   180
         Left            =   405
         TabIndex        =   1
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame fraLine 
      Caption         =   "  授予的相关权限  "
      Height          =   3030
      Left            =   -15
      TabIndex        =   32
      Top             =   5085
      Width           =   12525
      Begin VSFlex8Ctl.VSFlexGrid vsfModule 
         Height          =   2370
         Left            =   150
         TabIndex        =   40
         Top             =   210
         Width           =   12270
         _cx             =   21643
         _cy             =   4180
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
         BackColorBkg    =   -2147483643
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmUserEdit.frx":14B4
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
         ExplorerBar     =   5
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
   End
   Begin VB.Frame fraRole 
      Caption         =   "角色授权:(选择该用户可充当的角色)"
      Height          =   3705
      Left            =   3450
      TabIndex        =   16
      Top             =   105
      Width           =   7875
      Begin VB.CheckBox chkGranted 
         Caption         =   "只显已授权角色(&O)"
         Height          =   285
         Left            =   4185
         TabIndex        =   19
         Top             =   278
         Width           =   1860
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   990
         TabIndex        =   21
         Top             =   645
         Width           =   2805
      End
      Begin VB.ComboBox cboRoleGroups 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   270
         Width           =   2805
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRole 
         Height          =   2580
         Left            =   105
         TabIndex        =   22
         Top             =   1020
         Width           =   7740
         _cx             =   13652
         _cy             =   4551
         Appearance      =   2
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmUserEdit.frx":1545
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Editable        =   2
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "过  滤(&S)"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lblRoleGroups 
         AutoSize        =   -1  'True
         Caption         =   "角色组(&R)"
         Height          =   180
         Left            =   150
         TabIndex        =   17
         Top             =   345
         Width           =   810
      End
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4620
      TabIndex        =   33
      Top             =   1320
      Width           =   90
   End
End
Attribute VB_Name = "frmUserEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
Private Enum ApplyToEnum
    ATE_当前部门 = 0
    ATE_所有人员 = 1
    ATE_当前人员 = 2
End Enum

Private Enum VsfModuleTitle
    VMT_系统 = 0
    VMT_序号 = 1
    VMT_标题 = 2
    VMT_功能 = 3
    VMT_说明 = 4
End Enum

Private mstr所有者 As String
Private mrsModule As New ADODB.Recordset '保存角色的功能细节
Private mrsRole As New ADODB.Recordset   '保存角色
Private mstrUser As String
Private mblnSucceed As Boolean
Private mstrItem As String
Private mblnLoad As Boolean
Private mblnRISMsg As Boolean
Private mstrCreateUserList As String     '记录新增的用户名列表

'==============================================================
'==公共接口
'==============================================================
Public Function UserEdit(ByVal strOwner As String, Optional ByVal strUser As String, Optional ByRef strItem As String) As Boolean
'参数: strOwner 正在编辑的系统的所有者名
'      strUser  当前编辑的用户名，如果为空，表示新增
'出参:strItem  修改后的返回值，用于更新界面显示。
'返回：如果增加或修改成功,返回true,否则False
    mstr所有者 = strOwner
    mstrUser = strUser
    mblnSucceed = False: mstrItem = ""
    frmUserEdit.Show vbModal, frmMDIMain
    UserEdit = mblnSucceed
    strItem = mstrItem
End Function

'==============================================================
'==控件事件
'==============================================================
Private Sub cboRoleGroups_Click()
    Call FillRole
End Sub

Private Sub chkGranted_Click()
    Call FillRole
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "ZL9Svrtools\" & Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim blnChangeMen As Boolean, strUser As String
    Dim rsPerson As ADODB.Recordset
    Dim strPre简码 As String
    Dim objPercent As clsPercent
    Dim blnHaveRis As Boolean
    Dim strNote As String, strCheck As String, strNoCheck As String
    
    '数据校验
    If Not VilidateData() Then Exit Sub
    mblnRISMsg = False
    If UCase(gstrSTOwner) = UCase(mstr所有者) And gblnMustRIS Then  '是标准版的所有者
        blnHaveRis = gblnRIS
    End If
    On Error Resume Next
    '生成数据库用户以及对应应用系统人员
    strUser = UCase(Trim(txtUserName.Text))
    '修改用户且对应人员发生变化
    blnChangeMen = txtName.Tag <> txtNO.Tag And txtNO.Tag <> ""
    If blnChangeMen And mstrUser = mstr所有者 Then '只能修改对应人员，则直接退出
        gcnOracle.Execute "Delete From " & mstr所有者 & ".上机人员表 Where 用户名='" & strUser & "'"
        gcnOracle.Execute "Insert Into " & mstr所有者 & ".上机人员表(用户名,人员id) Values ('" & strUser & "'," & txtNO.Tag & ")"
        If err.Number <> 0 Then
            MsgBox "由于共享性原因，用户身份信息(部门、姓名)更改失败。" & vbNewLine & err.Description, vbExclamation, gstrSysName
            err.Clear
        ElseIf blnHaveRis Then '通知RIS该用户被修改
            If Not gobjRIS.UserEdit(2, strUser) Then
                mblnRISMsg = True
            End If
        End If
    Else
        If blnChangeMen Then '创建应用系统用户失败则退出
            If Not CreateApptionUser(Val(txtNO.Tag), strUser, txtPasswd.Text, mstrUser = "", True) Then
                txtUserName = "": txtPasswd = "": txtVerify = ""
                txtUserName.SetFocus: Exit Sub
            ElseIf blnHaveRis Then '通知RIS该用户被修改
                If Not gobjRIS.UserEdit(IIf(mstrUser = "", 1, 2), strUser) Then
                    mblnRISMsg = True
                End If
            End If
        End If
        If mstrUser <> mstr所有者 Then '可以进行授权调整
            If Not gblnDBA Then '当前用户非DBA，且存在其他用户创建的角色，则需要使用System连接
                mrsRole.Filter = "Grantee <> '" & gstrUserName & "'"
                If mrsRole.RecordCount > 0 Then '获取SysTem用户连接
                    Set gcnSystem = GetConnection("SYSTEM")
                    If gcnSystem Is Nothing Then Exit Sub '失败就退出
                End If
            End If
            mrsRole.Filter = ""
            Do While Not mrsRole.EOF
                If mrsRole!Granted <> mrsRole!勾选 Or mrsRole!Admin <> mrsRole!转授 Then
                    mrsRole.Update "改变", 1
                End If
                mrsRole.MoveNext
            Loop
            mrsRole.Filter = "改变=1"
            If Not optApplyTo(ATE_当前人员).value Then Set rsPerson = GetOtherPerson
            '若只应用与当前人员，且角色没有发生变化，则自动退出.或者其他的应用者没有查询到
            If Not (mrsRole.RecordCount = 0 And optApplyTo(ATE_当前人员).value Or Not optApplyTo(ATE_当前人员).value And rsPerson Is Nothing And mrsRole.RecordCount = 0) Then
                If mrsRole.RecordCount <> 0 Then
                    mrsRole.Filter = "改变=1"
                    If Not ApplyOnePerson(strUser, True, True) Then
                        If vsRole.Enabled Then vsRole.SetFocus
                        Exit Sub
                    End If
                    mrsRole.Filter = "勾选=1": mrsRole.Sort = "Role" '取消过滤，供授权使用
                Else
                    mrsRole.Filter = "勾选=1": mrsRole.Sort = "Role" '取消过滤，供授权使用
                End If
                If Not rsPerson Is Nothing Then '应用到其他人员
                    Set objPercent = New clsPercent
                    Call objPercent.InitPercent(prg, rsPerson.RecordCount)
                    lblProssCaption.Caption = "": lblStep.Caption = "": picProcess.Visible = True
                    Do While Not rsPerson.EOF
                        strUser = GetUserName(rsPerson!编号 & "", rsPerson!简码 & "", strPre简码)
                        lblProssCaption.Caption = "正在处理用户:" & strUser & "(" & rsPerson!姓名 & ")"
                        If CreateApptionUser(Val(rsPerson!id & ""), strUser, strUser, True) Then
                            If blnHaveRis Then '通知RIS用户被创建
                                If Not gobjRIS.UserEdit(1, strUser) Then
                                    mblnRISMsg = True
                                End If
                            End If
                            Call ApplyOnePerson(strUser)
                        End If
                        strPre简码 = rsPerson!简码 & ""
                        objPercent.LoopPercent
                        lblStep.Caption = prg.value & "%"
                        rsPerson.MoveNext
                    Loop
                    lblProssCaption.Caption = "": lblStep.Caption = "": picProcess.Visible = False
                End If
            End If
        End If
    End If
    If mblnRISMsg Then
        MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(UserEdit)未调用成功，请联系管理员！", vbInformation, gstrSysName
    End If
    mblnSucceed = True
    If txtUserName.Enabled = False Then '修改用户
        strNote = "": strCheck = "": strNoCheck = ""
        If txtName.Tag <> txtNO.Tag And txtNO.Tag <> "" Then
            strNote = "对应人员由“" & txtUserName.Tag & "”修改为“" & txtName.Text & "”；"
            txtUserName.Tag = ""
        End If
        mrsRole.Filter = "改变=1"
        If mrsRole.RecordCount > 0 Then
            Do While Not mrsRole.EOF
                If mrsRole!勾选 = 0 Then  '取消勾选的角色
                    strNoCheck = IIf(strNoCheck = "", "取消勾选的角色有：", strNoCheck & "，") & mrsRole!RoleName
                Else  '添加勾选的角色
                    strCheck = IIf(strCheck = "", "添加勾选的角色有：", strCheck & "，") & mrsRole!RoleName
                End If
                mrsRole.MoveNext
            Loop
        End If
        strNote = strNote & IIf(strCheck = "", "", strCheck & "；") & IIf(strNoCheck = "", "", strNoCheck & "；")
        If mstrCreateUserList <> "" Then
            strNote = strNote & "关联添加的用户有：" & mstrCreateUserList
            mstrCreateUserList = ""
        End If
                
        '插入重要操作日志
        If strNote <> "" Then
            Call SaveAuditLog(2, "修改用户", txtUserName.Text & "；" & strNote)
        End If
        If optApplyTo(ATE_当前人员).value Then
            mstrItem = txtNO.Text & "|" & txtName.Text & "|" & txtDept.Text
        End If
        Unload Me
    Else '新增
        mrsRole.Filter = "勾选=1"
        If mrsRole.RecordCount > 0 Then
            Do While Not mrsRole.EOF
                strNote = IIf(strNote = "", "", strNote & "，") & mrsRole!RoleName
                mrsRole.MoveNext
            Loop
        End If
        '插入重要操作日志
        If mstrCreateUserList <> "" Then
            Call SaveAuditLog(1, "新增用户", mstrCreateUserList & "；添加的角色为：" & strNote)
        End If
        
        txtUserName.Text = "": txtPasswd.Text = "": txtVerify.Text = ""
        txtNO.Text = "": txtNO.Tag = "": txtName.Text = ""
        txtDept.Text = "": mstrCreateUserList = ""
        cmdSelect.SetFocus
    End If
End Sub

Private Sub cmdSelect_Click()
    If Not LoadPerson Then
        MsgBox "未找到" & IIf(mstrUser = "", "尚未创建用户的人员！", "部门人员！"), vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNO Is ActiveControl Then
            KeyAscii = 0
        Else
            PressKey vbKeyTab
        End If
    ElseIf KeyAscii = Asc("'") Or Chr(KeyAscii) = "@" Or Chr(KeyAscii) = " " Or Chr(KeyAscii) = "\" Or Chr(KeyAscii) = """" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If Not InitBaseData Then GoTo errEnd
    '打开所有角色的对应的模块
    If Not GetModuleAndRole Then GoTo errEnd
    If Not InitUser() Then GoTo errEnd
    '选择所有角色
    cboRoleGroups.ListIndex = 0
    Exit Sub
errEnd: '强制关闭窗体
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsModule Is Nothing Then Set mrsModule = Nothing
    Call SaveSetting("ZLSOFT", "用户设置", "方式", IIf(optMode(0).value, "0", "1"))
End Sub

Private Sub tvwPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call tvwPerson_DblClick
End Sub

Private Sub tvwPerson_LostFocus()
    DoEvents
    If cmdSelect Is ActiveControl Or txtNO Is ActiveControl Or tvwPerson Is ActiveControl Then Exit Sub
    If txtNO.Tag = "" And lblNO.Tag <> "" Then
        txtNO.Tag = Split(lblNO.Tag, ",")(0)
        txtNO.Text = Split(lblNO.Tag, ",")(1)
        txtName.Text = lblName.Tag
    End If
    tvwPerson.Visible = False
    tvwPerson.Nodes.Clear
End Sub

Private Sub tvwPerson_DblClick()
    If tvwPerson.SelectedItem Is Nothing Then Exit Sub
    If tvwPerson.SelectedItem.Tag <> 2 Then Exit Sub
    '选择的是人员节点
    Call ChangePerson(tvwPerson.SelectedItem)
End Sub

Private Sub txtNO_GotFocus()
    SelAll txtNO
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    txtNO.Tag = ""
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lblNO.Tag <> "" Then
        If txtNO.Text = Split(lblNO.Tag, ",")(1) Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            Exit Sub
        End If
    End If
    LoadPerson (txtNO.Text)
    If txtNO.Tag = "" And txtNO.Text <> "" And Not tvwPerson.Visible Then
        MsgBox "没有找到任何" & IIf(mstrUser = "", "尚未创建用户的", "") & "人员信息，请检查后重新输入！"
        If lblNO.Tag <> "" Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            txtNO.Text = Split(lblNO.Tag, ",")(1)
        End If
        SelAll txtNO
    End If
End Sub

Private Sub txtno_LostFocus()
    If tvwPerson.Visible And Not cmdSelect Is ActiveControl And Not txtNO Is ActiveControl And Not tvwPerson Is ActiveControl Then
        tvwPerson.Visible = False
        txtNO.SetFocus
        txtNO.SelStart = 0
        Me.txtNO.SelLength = Len(txtNO.Text)
    ElseIf tvwPerson.Visible Then
        tvwPerson.SetFocus
    End If
End Sub

Private Sub txtno_Validate(Cancel As Boolean)
    If cmdSelect Is ActiveControl Or tvwPerson.Visible Then Exit Sub
     '列出部门表和对应人员
    If txtNO.Tag = "" Then
        If lblNO.Tag <> "" Then
            If txtNO.Text = Split(lblNO.Tag, ",")(1) Then
                Me.txtNO.Tag = Split(lblNO.Tag, ",")(0)
                Exit Sub
            End If
        End If
        LoadPerson (txtNO.Text)
    End If
    
    If txtNO.Tag = "" And txtNO.Text <> "" And Not tvwPerson.Visible Then
        MsgBox "没有找到任何" & IIf(mstrUser = "", "尚未创建用户的", "") & "人员信息，请检查后重新输入！"
        If lblNO.Tag <> "" Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            txtNO.Text = Split(lblNO.Tag, ",")(1)
        End If
        Cancel = True
        SelAll txtNO
    End If
End Sub

Private Sub txtPasswd_GotFocus()
    SelAll txtPasswd
End Sub

Private Sub txtSearch_Change()
    Call FillRole
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserName_Change()
    txtUserName = UCase(txtUserName)
    txtUserName.SelStart = Len(txtUserName)
End Sub

Private Sub txtUserName_GotFocus()
    SelAll txtUserName
End Sub

Private Sub txtVerify_GotFocus()
    SelAll txtVerify
End Sub

Private Sub vsRole_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngNameCol As Long
    With vsRole
        lngNameCol = (Col \ 2) * 2
        If Col = lngNameCol Then
            If .Cell(flexcpChecked, Row, Col) = flexUnchecked Then
                .Cell(flexcpChecked, Row, Col + 1) = flexUnchecked
            End If
        Else
            If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                .Cell(flexcpChecked, Row, Col - 1) = flexChecked
            End If
        End If
        Call RecUpdate(mrsRole, "RoleName='" & .TextMatrix(Row, lngNameCol) & "'", "勾选", IIf(.Cell(flexcpChecked, Row, lngNameCol) = flexUnchecked, 0, 1), "转授", IIf(.Cell(flexcpChecked, Row, lngNameCol + 1) = flexUnchecked, 0, 1))
        '勾选转授，以及勾选授权，重新调整模块展示
        If Col = lngNameCol Or Col <> lngNameCol And .Cell(flexcpChecked, Row, lngNameCol + 1) = flexChecked Then
            Call FillModule
        End If
    End With
End Sub

Private Sub vsRole_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    If vsRole.TextMatrix(NewRowSel, NewColSel - (NewColSel Mod 2)) = "" Then Cancel = True
End Sub

Private Sub vsRole_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsRole.BackColor = vsRole.BackColorFixed Then
        Cancel = True
    ElseIf vsRole.TextMatrix(Row, Col - (Col Mod 2)) = "" Then
        Cancel = True
    End If
End Sub

'==============================================================
'==私有方法
'==============================================================
Private Function InitBaseData() As Boolean
'功能：初始化工作范围，人员性质,用户设置方式
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strInfo As String
    '用户设置方式
    optMode(0).value = Val(GetSetting("ZLSOFT", "用户设置", "方式", "0")) = 0
    optMode(1).value = Not Me.optMode(0).value
    '初始化工作范围
    With cboWorkRange
        .addItem "      门诊": .ItemData(.NewIndex) = 1
        .addItem "      住院": .ItemData(.NewIndex) = 2
        .addItem "门诊与住院": .ItemData(.NewIndex) = 3
    End With
    '初始化人员性质
    On Error GoTo errh:
    strInfo = "人员性质"
RUMMan:
    strSQL = "Select 编码, 名称 From " & mstr所有者 & ".人员性质分类"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboManPros.Clear
    With rsTmp
        Do While Not .EOF
            cboManPros.addItem !名称
            .MoveNext
        Loop
    End With
    strInfo = "角色分组"
RUMGroup:
    strSQL = "Select '所有分组' 组名, Count(1) 数量, 2 标识" & vbNewLine & _
            "From Zlroles" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Nvl(b.组名,'未分组') 组名, Count(1) 数量, Decode(b.组名, Null, 1, 0) 标识" & vbNewLine & _
            "From Zltools.Zlroles a, Zltools.Zlrolegroups b" & vbNewLine & _
            "Where a.名称 = b.角色(+)" & vbNewLine & _
            "Group By b.组名" & vbNewLine & _
            "Order By 组名"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With cboRoleGroups
        .Clear
        rsTmp.Filter = "组名 = '所有分组' And 标识 = 2"
        .addItem "所有角色" & "(" & rsTmp!数量 & ")"
        rsTmp.Filter = "组名 = '未分组'"
        If rsTmp.RecordCount <> 0 Then
            .addItem "未分组" & "(" & rsTmp!数量 & ")"
        End If
        rsTmp.Filter = "标识 = 0"
        Do While Not rsTmp.EOF
            .addItem rsTmp!组名 & "(" & rsTmp!数量 & ")"
            rsTmp.MoveNext
        Loop
    End With
    InitBaseData = True
    Exit Function
errh:
    If strInfo <> "" Then
        If MsgBox("装入" & strInfo & "时出现以下错误：" & vbCrLf & vbCrLf & _
                    err.Description & vbCrLf & vbCrLf & "需要再试一次吗？", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            err.Clear
            If strInfo = "人员性质" Then
                GoTo RUMMan
            Else
                GoTo RUMGroup
            End If
        End If
    Else
        MsgBox "  错误号:" & err.Number & "  错误描述:" & err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function InitUser() As Boolean
'功能：初始化用户数据
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    '既不是DBA，也不是当前系统的所有者，不能修改对应人员
    If Not (gblnDBA Or mstr所有者 = UCase(gstrUserName)) Then
        cmdSelect.Enabled = False
    End If
    
    If mstrUser <> "" Then
        '当前用户是所有者，就不允许对角色进行修改
        If mstr所有者 = mstrUser Then
            cboRoleGroups.Enabled = False
            txtSearch.Enabled = False
            chkGranted.Enabled = False
            vsRole.BackColor = vsRole.BackColorFixed: vsRole.BackColorBkg = vsRole.BackColorFixed
            cboWorkRange.Enabled = False
            cboManPros.Enabled = False
            optApplyTo(ATE_当前部门).Enabled = False: optApplyTo(ATE_所有人员).Enabled = False: optApplyTo(ATE_当前人员).Enabled = False
        End If
        '加载用户相关信息
        txtUserName.Text = mstrUser: txtPasswd.Text = "12345678": txtVerify.Text = "12345678"
        txtUserName.Enabled = False: txtPasswd.Enabled = False: txtVerify.Enabled = False
        optMode(0).Enabled = False: optMode(1).Enabled = False
        On Error GoTo errh
ReLoad:
        strSQL = "Select c.Id, c.编号, c.姓名, a.名称, a.Id As 部门id" & vbNewLine & _
                    "From " & mstr所有者 & ".部门表 a, " & mstr所有者 & ".部门人员 b, " & mstr所有者 & ".人员表 c, " & mstr所有者 & ".上机人员表 d" & vbNewLine & _
                    "Where a.Id = b.部门id And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And b.缺省 = 1 And b.人员id = c.Id And" & vbNewLine & _
                    "      c.Id = d.人员id And d.用户名 = '" & mstrUser & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        
        If rsTmp.RecordCount > 0 Then
            txtNO.Text = rsTmp!编号
            txtName.Text = rsTmp!姓名
            txtUserName.Tag = rsTmp!姓名
            txtDept.Text = rsTmp!名称
            txtDept.Tag = rsTmp!部门id
            txtNO.Tag = rsTmp!id: txtName.Tag = rsTmp!id
            Call LocateManPros(rsTmp!id)
            Call LocateWorkRange(rsTmp!部门id)
        End If
    End If
    InitUser = True
    Exit Function
errh:
    If MsgBox("装入人员时出现以下错误：" & vbCrLf & vbCrLf & _
                err.Description & vbCrLf & vbCrLf & "需要再试一次吗？", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        err.Clear
        GoTo ReLoad
    End If
End Function


Private Function GetModuleAndRole() As Boolean
'功能，获取角色数据以及角色对应的模块数据
    Dim strInfo As String, strSQL As String
    
    strSQL = "Select User As Grantee, g.名称 Role, g.系统, Decode(n.Granted_Role, Null, 0, 1) As Granted," & vbNewLine & _
            "       Decode(n.Admin_Option, 'YES', 1, 0) As Admin, b.组名, Zlspellcode(Substr(g.名称, 3)) As 简码, Substr(g.名称, 4) Rolename," & vbNewLine & _
            "       Decode(n.Granted_Role, Null, 0, 1) 勾选, Decode(n.Admin_Option, 'YES', 1, 0) 转授, 1 展示, 0 改变" & vbNewLine & _
            "From (Select 名称, 系统 From Zlroles) g," & vbNewLine & _
            "     (Select Granted_Role, Admin_Option From Dba_Role_Privs Where Grantee = [1] And Granted_Role Like 'ZL_%') n," & vbNewLine & _
            "     Zlrolegroups b" & vbNewLine & _
            "Where g.名称 = n.Granted_Role(+) And g.名称 = b.角色(+)" & vbNewLine & _
            "Order By g.名称"

    strInfo = "角色数据"
    On Error GoTo errh
RUMRole:
    Set mrsRole = Nothing
    Set mrsRole = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrUser))
    strSQL = "Select *" & vbNewLine & _
            "From (" & vbNewLine & _
            "       Select c.名称, c.编号, a.角色, a.序号, a.功能, b.标题, b.说明" & vbNewLine & _
            "From Zlrolegrant a, Zlprograms b, Zlsystems c" & vbNewLine & _
            "Where a.序号 = b.序号 And Nvl(a.系统, 0) = Nvl(b.系统, 0) And b.系统 = c.编号(+)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select t.名称, t.编号, r.Grantee 角色, Null 序号, Null 功能, t.表名 标题, t.说明" & vbNewLine & _
            "From (Select s.名称, s.编号, s.所有者, b.表名, b.说明 From Zlsystems s, Zlbasecode b Where b.系统 = s.编号) t," & vbNewLine & _
            "     (Select Grantee, Owner, Table_Name" & vbNewLine & _
            "       From User_Tab_Privs" & vbNewLine & _
            "       Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
            "       Group By Grantee, Owner, Table_Name" & vbNewLine & _
            "       Having Count(Privilege) = 4) r" & vbNewLine & _
            "Where t.所有者 = r.Owner And t.表名 = r.Table_Name" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select s.名称, s.编号, r.Grantee 角色, Null 序号, Null 功能, f.函数名 || '(' || f.中文名 || ')' 标题, f.说明" & vbNewLine & _
            "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
            "       Where f.系统 = s.编号 And s.所有者 = r.Owner And Upper(f.函数名) = r.Table_Name And r.Privilege = 'EXECUTE')" & vbNewLine & _
            "       Order By 编号, 序号"

    strInfo = "角色模块数据"
RUMModule:
    Set mrsModule = Nothing
    Set mrsModule = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    GetModuleAndRole = True
    Exit Function
errh:
    If MsgBox("装入" & strInfo & "时出现以下错误：" & vbCrLf & vbCrLf & _
                err.Description & vbCrLf & vbCrLf & "需要再试一次吗？", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        err.Clear
        If strInfo = "角色数据" Then
            GoTo RUMRole
        Else
            GoTo RUMModule
        End If
    End If
End Function

Private Sub LocateManPros(ByVal lng人员id As Long)
'功能：定位当前人员的人员性质
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    
    If mstr所有者 = mstrUser Then Exit Sub
    strSQL = "Select 人员性质 From 人员性质说明 Where 人员id= " & lng人员id
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)

    With cboManPros
        .ListIndex = -1
        If rsTmp.EOF Then Exit Sub
        For i = 0 To .ListCount
            If .List(i) = Nvl(rsTmp!人员性质) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
End Sub

Private Sub LocateWorkRange(ByVal lng部门ID As Long)
'功能：定位工作范围
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnLock As Boolean
    If mstr所有者 = mstrUser Then Exit Sub
    If lng部门ID <> 0 Then
        strSQL = "Select 1 From " & mstr所有者 & ".部门性质说明 Where 部门id = " & lng部门ID & " And 工作性质 = '临床'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        blnLock = rsTmp.EOF
    Else
        blnLock = True
    End If
    If blnLock Then
        cboWorkRange.ListIndex = -1
        cboWorkRange.Enabled = False
    Else
        cboWorkRange.ListIndex = 2
        cboWorkRange.Enabled = True
    End If
End Sub

Private Function FillRole() As Boolean
'功能:将数据填充到Role列表中
'参数:strFilter-过滤条件
'返回:加载成功,返回true,否则返回False
    Dim lngRow As Long, lngCol As Long, i As Long
    
    Call RecUpdate(mrsRole, "", "展示", 0)
    mrsRole.Filter = GetRoleFilter
    mrsRole.Sort = "Rolename"
    With vsRole
        .Redraw = flexRDNone: .Rows = .FixedRows
        .Rows = -1 * Int(-1 * Val(mrsRole.RecordCount / 3)) + .FixedRows
        .Row = 0: .Col = 0
        If .Rows > .FixedRows Then
            .CellBorderRange 0, 1, .Rows - 1, 1, vbBlack, 0, 0, 1, 0, 0, 0
            .CellBorderRange 0, 3, .Rows - 1, 3, vbBlack, 0, 0, 1, 0, 0, 0
            .Cell(flexcpPictureAlignment, .FixedRows, 1, .Rows - 1, 1) = 4
            .Cell(flexcpPictureAlignment, .FixedRows, 3, .Rows - 1, 3) = 4
            .Cell(flexcpPictureAlignment, .FixedRows, 5, .Rows - 1, 5) = 4
            For i = 0 To mrsRole.RecordCount - 1
                lngRow = i \ 3 + .FixedRows: lngCol = (i Mod 3) * 2
                .TextMatrix(lngRow, lngCol) = mrsRole!RoleName
                .Cell(flexcpChecked, lngRow, lngCol) = IIf(mrsRole!勾选 = 1, flexChecked, flexUnchecked)
                .Cell(flexcpChecked, lngRow, lngCol + 1) = IIf(mrsRole!Admin = 1, flexChecked, flexUnchecked)
                mrsRole.Update "展示", 1
                mrsRole.MoveNext
            Next
            .Row = .FixedRows: .Col = 0
        End If
        .Redraw = flexRDDirect
    End With
    Call FillModule
    FillRole = True
End Function

Private Function GetRoleFilter() As String
    Dim strSearChar As String, strGroup As String
    Dim strFilter As String

    strSearChar = Replace(UCase(Trim(txtSearch.Text)), "'", "")
    strGroup = Mid(cboRoleGroups.Text, 1, InStrRev(cboRoleGroups.Text, "(") - 1)
    If strGroup = "所有角色" Then strGroup = ""
    If strGroup = "未分组" Then
        strGroup = ""
        strFilter = "组名=null"
    End If
    '进行组过滤
    strFilter = IIf(strGroup = "", strFilter, "组名='" & strGroup & "'")
    strFilter = strFilter & IIf(glngSysNo = -1, "", IIf(strFilter = "", "", " And ") & "系统=" & glngSysNo)
    strFilter = strFilter & IIf(chkGranted.value = 1, IIf(strFilter <> "", " And ", "") & "勾选=1", "")
    If strSearChar <> "" Then
        If strFilter = "" Then
            strFilter = "RoleName Like '" & strSearChar & "%' OR 简码 Like '" & strSearChar & "%'"
        Else
            strFilter = "(" & strFilter & " And RoleName Like '" & strSearChar & "%' ) OR (" & strFilter & " And 简码 Like '" & strSearChar & "%' )"
        End If
    End If
    GetRoleFilter = strFilter
End Function

Private Sub FillModule()
'按所选取的角色列出所有的模块及功能
    Dim strModule As String, strFun As String
    Dim strAllFun As String, strRole As String
    Dim lngSerialNumber As Long
    
    '获取界面上展示并勾选的角色
    mrsRole.Filter = "展示=1 And 勾选=1"
    vsfModule.Rows = 1
    If mrsRole.RecordCount = 0 Then Exit Sub
    Do While Not mrsRole.EOF
        strRole = strRole & " OR (角色='" & mrsRole!Role & "'" & IIf(glngSysNo = -1, "", " And 编号=" & glngSysNo) & ")"
        mrsRole.MoveNext
    Loop
    strRole = Mid(strRole, Len(" OR "))
    '过滤界面上展示并勾选的角色对应的模块，并以序号功能排序
    mrsModule.Filter = strRole
    Do While Not mrsModule.EOF
        If lngSerialNumber <> Nvl(mrsModule!序号, -1) Or IsNull(mrsModule!序号) Then
            vsfModule.Rows = vsfModule.Rows + 1
            If IsNull(mrsModule!名称) Then
                If mrsModule!序号 < 100 Then
                    vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_系统) = "基础工具"
                Else
                    vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_系统) = "自定义报表"
                End If
            Else
                vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_系统) = mrsModule!名称 & "(" & mrsModule!编号 & ")"
            End If
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_序号) = mrsModule!序号 & ""
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_标题) = mrsModule!标题
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_功能) = IIf(mrsModule!功能 & "" = "基本", "", mrsModule!功能 & "")
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_说明) = mrsModule!说明 & ""
            lngSerialNumber = Nvl(mrsModule!序号, -1)
        Else
            If mrsModule!功能 & "" <> "基本" Then
                vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_功能) = vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_功能) & IIf(IsNull(mrsModule!功能), "", ",") & mrsModule!功能
            End If
        End If
        mrsModule.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox "加载授权模块数据时发生错误,详细的错误信息如下:" & vbCrLf & "  错误号:" & err.Number & "  错误描述:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Function LoadPerson(Optional ByVal strMenInfo As String) As Boolean
'功能：根据条件展示人员选择器
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim rsMen As ADODB.Recordset
    Dim objNode As Node, objParent As Node
    Dim strKey As String, i As Long
    
    On Error GoTo errh
    tvwPerson.Nodes.Clear: tvwPerson.Visible = False: tvwPerson.Tag = ""
     If txtNO.Tag <> "" And strMenInfo = "" Then
        strKey = "P" & txtNO.Tag
     End If
    '读取匹配人员，若未读取到则退出,新增用户时mstrUser = ""，只查询未创建的用户
    strSQL = "Select a.Id, a.编号, a.姓名, b.部门id" & vbNewLine & _
                "From " & mstr所有者 & ".人员表 a, " & mstr所有者 & ".部门表 c, " & mstr所有者 & ".部门人员 b" & IIf(mstrUser = "", "," & mstr所有者 & ".上机人员表 D", "") & vbNewLine & _
                "Where (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And a.Id = b.人员id And b.缺省 = 1 And b.部门id = c.Id And" & vbNewLine & _
                "      (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) " & IIf(mstrUser = "", " And B.人员id = d.人员id(+) And D.人员id is null ", "")
    If strMenInfo <> "" Then
        strMenInfo = UCase(strMenInfo)
        If IsNumeric(strMenInfo) Then
            strSQL = strSQL & "And a.编号 Like '" & strMenInfo & "%'"
        ElseIf IsCharAlpha(strMenInfo) Then
            strSQL = strSQL & "And a.简码 Like '" & strMenInfo & "%'"
        ElseIf IsCharChinese(strMenInfo) Then
            strSQL = strSQL & "And a.姓名 Like '" & strMenInfo & "%'"
        Else
            strSQL = gstrSQL & "And (a.简码 Like '" & strMenInfo & "%' OR a.姓名 Like '" & strMenInfo & "%')"
        End If
    End If
    Set rsMen = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If rsMen.EOF Then Exit Function '没有查询到人员则不展示
    '读取并加载部门树形列表
    strSQL = "Select Id, 编码, 名称, 上级id" & vbNewLine & _
                "From " & mstr所有者 & ".部门表" & vbNewLine & _
                "Where 编码 <> '-' And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & vbNewLine & _
                "Start With 上级id Is Null" & vbNewLine & _
                "Connect By Prior Id = 上级id"
    Set rsDept = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Do While Not rsDept.EOF
        If IsNull(rsDept!上级id) Then
            Set objNode = tvwPerson.Nodes.Add(, , "K" & rsDept!id, "【" & rsDept!编码 & "】" & rsDept!名称, "Dept", "Dept")
        Else
            Set objNode = tvwPerson.Nodes.Add("K" & rsDept!上级id, tvwChild, "K" & rsDept!id, "【" & rsDept!编码 & "】" & rsDept!名称, "Dept", "Dept")
        End If
        objNode.Tag = 0
        rsDept.MoveNext
    Loop
    '加载部门的下属人员
    Do While Not rsMen.EOF
        Set objNode = tvwPerson.Nodes.Add("K" & rsMen!部门id, tvwChild, "P" & rsMen!id, "【" & rsMen!编号 & "】" & rsMen!姓名, "User", "User")
        objNode.ForeColor = RGB(0, 0, 255)
        If strKey = "" Then
            strKey = objNode.Key
        End If
        '标记父级，声明已经存在子级或子级的子级,在查询模式下，并展开节点
        If objNode.Parent.Tag = 0 Then
            Set objParent = objNode.Parent
            Do While Not objParent Is Nothing
                If objParent.Tag = 1 Then Exit Do
                objParent.Tag = 1 '标记
                If strMenInfo <> "" Then objParent.Expanded = True
                Set objParent = objParent.Parent
            Loop
        End If
        objNode.Tag = 2
        rsMen.MoveNext
    Loop
    '移除没有下级的父节点
    i = 1
    Do
        If tvwPerson.Nodes(i).Tag = 0 Then
            tvwPerson.Nodes.Remove i
        Else
             i = i + 1
        End If
    Loop While (i <= tvwPerson.Nodes.Count)
    On Error Resume Next
    If tvwPerson.Nodes.Count = 0 Then Exit Function '没有节点则不展示
    LoadPerson = True
    DoEvents
    tvwPerson.Visible = True
    If strKey <> "" Then
        tvwPerson.Nodes(strKey).Selected = True
        tvwPerson.SelectedItem.EnsureVisible
    End If
    tvwPerson.SetFocus
    With tvwPerson
        .Top = Me.ScaleTop
        .Left = cmdSelect.Left + cmdSelect.Width + fraPerson.Left
        .Height = Me.ScaleHeight
        .ZOrder
    End With
    Exit Function
errh:
    If MsgBox("装入人员时出现以下错误：" & vbCrLf & vbCrLf & _
        err.Description & vbCrLf & vbCrLf & "需要再试一次吗？", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Function

Private Function VilidateData() As Boolean
'功能：对数据合法性进行校验
    Dim strUser As String
    
    strUser = UCase(Trim(txtUserName.Text))
    If strUser = "SYS" Or strUser = "SYSTEM" Then
        MsgBox "不能使用SYS，SYSTEM用户。", vbInformation, gstrSysName
        txtUserName.Text = "": Exit Function
    End If
    
    If txtUserName.Enabled Then      '新建用户
        If Len(strUser) = 0 Then
            MsgBox "请输入用户名。", vbExclamation, gstrSysName
            txtUserName.SetFocus: Exit Function
        End If
        If Len(Trim(txtPasswd)) < 2 Then
            MsgBox "必须分配两位字符以上的用户密码。", vbExclamation, gstrSysName
            txtPasswd.SetFocus: Exit Function
        End If
        If StrIsValid(strUser, txtUserName.MaxLength) = False Then
            txtUserName.SetFocus
            Exit Function
        End If
        If StrIsValid(Trim(txtPasswd.Text), txtPasswd.MaxLength) = False Then
            txtPasswd.SetFocus: Exit Function
        End If
        If txtPasswd <> txtVerify Then
            MsgBox "用户密码与验证密码不一致", vbExclamation, gstrSysName
            txtPasswd = "": txtVerify = ""
            txtPasswd.SetFocus: Exit Function
        End If
    End If
    
    VilidateData = True
End Function

Private Function GetUserName(ByVal str编码 As String, ByVal str简码 As String, ByVal strPre简码 As String) As String
'功能：生成一个数据库用户名
    If optMode(0).value = True Then
        If strPre简码 <> str简码 Then
            GetUserName = str简码
        Else
            GetUserName = str简码 & str编码
        End If
    Else
        '按U+编号处理
        If IsNumeric(Left(str编码, 1)) Then
            GetUserName = "U" & str编码
        Else
            GetUserName = str编码
        End If
    End If
End Function

Private Function GetOtherPerson() As ADODB.Recordset
'功能：根据条件，获取到可以应用的其他人
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strWorkRange As String, strManPros As String
    '为什么增加临床性质限定？
    '没有临床信息，则进行部门性质过滤
    On Error GoTo errh
    If cboWorkRange.ListIndex > -1 Then
        strWorkRange = " And e.工作性质 = '临床'"
        If cboWorkRange.ItemData(cboWorkRange.ListIndex) < 2 Then
            strWorkRange = strWorkRange & " And e.服务对象 in (" & cboWorkRange.ItemData(cboWorkRange.ListIndex) & ",3) "
        End If
    End If
    strManPros = cboManPros.Text
    strSQL = "Select Distinct a.Id, a.编号, Decode(a.简码, Null, Zlspellcode(a.姓名), a.简码) As 简码, a.姓名" & vbNewLine & _
                    "From " & mstr所有者 & ".人员表 a, " & mstr所有者 & ".部门人员 b, " & mstr所有者 & ".上机人员表 c" & vbNewLine
    '增加人员性质的读取
    strSQL = strSQL & IIf(strManPros = "", "", ", " & mstr所有者 & ".人员性质说明 d") & vbNewLine
    '增加工作范围读取
    strSQL = strSQL & IIf(strWorkRange = "", "", ", " & mstr所有者 & ".部门性质说明 e") & vbNewLine
    '过滤条件处理,增加部门过滤条件
     strSQL = strSQL & _
                    "Where a.Id = b.人员id And a.Id = c.人员id(+) And" & vbNewLine & _
                    "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.缺省 = 1 And a.Id <> " & Val(txtNO.Tag) & vbNewLine & _
                    IIf(optApplyTo(ATE_当前部门).value, "  And B.部门id = " & Val(txtDept.Tag), "")
     '增加人员性质的过滤
     strSQL = strSQL & _
                   IIf(cboManPros <> "", " And A.id=D.人员id And D.人员性质='" & cboManPros & "'", "")
    '增加工作范围的过滤
     strSQL = strSQL & _
                   IIf(strWorkRange <> "", " And b.部门id = e.部门id " & strWorkRange, "")
    strSQL = strSQL & "And c.用户名 is Null  Order By 简码"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If rsTmp.RecordCount = 0 Then Exit Function '没有查询到可授权人员则退出
    Set GetOtherPerson = rsTmp
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetGrantRole(ByRef strNormalRoles As String, ByRef strAdminRoles As String) As Boolean
'功能：获取应该授权的角色
    '获取普通授权与转授授权
    strNormalRoles = "": strAdminRoles = ""
    mrsRole.Filter = "勾选=1": mrsRole.Sort = "转授"
    Do While Not mrsRole.EOF
        If mrsRole!转授 = 1 Then
            strAdminRoles = strAdminRoles & "," & mrsRole!Role
        Else
            strNormalRoles = strNormalRoles & "," & mrsRole!Role
        End If
        mrsRole.MoveNext
    Loop
    strNormalRoles = Mid(strNormalRoles, 2): strAdminRoles = Mid(strAdminRoles, 2)
    '取消过滤，供授权使用
    mrsRole.Filter = "": mrsRole.Sort = "Role"
End Function

Private Function SpellCode(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '功能：返回指定字符串的拼音简码
    '参数：（SSC编制）
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    '氽
    aryStard = Split("八;擦;哒;讹;发;噶;哈;击;;咔;垃;妈;拿;噢;啪;七;热;撒;他;挖;;挖;西;鸭;匝", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbProperCase)         '将全角转换为半角，首位转换为大写
    
    strCode = ""
    For intBit = 1 To Len(strAsk)
        If Mid(strAsk, intBit, 1) = "她" Then
            blnCan = True
            strCode = strCode & "T"
        ElseIf Asc(Mid(strAsk, intBit, 1)) < 0 Then
            blnCan = True
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(Mid(strAsk, intBit, 1), aryStard(iCount), vbTextCompare) = -1 Then
                        strCode = strCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        strCode = strCode & "Z"
                    End If
                End If
            Next
        Else
            If Mid(strAsk, intBit, 1) >= "A" And Mid(strAsk, intBit, 1) <= "Z" Then
                strCode = strCode & Mid(strAsk, intBit, 1)
            End If
        End If
        If Len(strCode) >= 10 Then Exit For
    Next
    SpellCode = strCode
    
End Function

Public Function ApplyOnePerson(ByVal strUser As String, Optional ByVal blnMsg As Boolean, Optional ByVal blnCurPerson As Boolean) As Boolean
'功能：讲角色授权应用给一个用户
'参数：
'         strUser=用户名
'         strPwd=密码
'         blnMsg=是否消息提示
'返回：是否成功
    Dim strRSQL As String, strGSQL As String
    On Error Resume Next
    '先收回权限，然后授权
    mrsRole.MoveFirst
    Do While Not mrsRole.EOF
        If blnCurPerson Then strRSQL = "Revoke " & mrsRole!Role & " From " & strUser
        If mrsRole!勾选 = 1 Then
            strGSQL = "Grant " & mrsRole!Role & " to " & strUser & IIf(mrsRole!转授 = 1, " With Admin Option", "")
        Else
            strGSQL = ""
        End If
        If Not gblnDBA And mrsRole!Grantee <> gstrUserName Then
            If blnCurPerson Then Call gclsBase.ExecuteCmdText(strRSQL, Me.Caption, gcnSystem, True)
            If err.Number <> 0 Then err.Clear
            If strGSQL <> "" Then Call gclsBase.ExecuteCmdText(strGSQL, Me.Caption, gcnSystem, True)
        Else
            If blnCurPerson Then Call gclsBase.ExecuteCmdText(strRSQL, Me.Caption, , True)
            If err.Number <> 0 Then err.Clear
            If strGSQL <> "" Then Call gclsBase.ExecuteCmdText(strGSQL, Me.Caption, , True)
        End If
        If err.Number <> 0 Then
            If blnMsg Then MsgBox "角色授予失败,错误信息如下:" & vbCrLf & err.Description, vbExclamation, gstrSysName
            err.Clear
            If blnMsg Then Exit Function
        End If
        mrsRole.MoveNext
    Loop
    '记录角色授权信息
    If err.Number <> 0 Then err.Clear
    Call ExecuteProcedure("Zl_Zluserroles_Add('" & strUser & "')", Me.Caption)
    If err.Number <> 0 Then
        If blnMsg Then MsgBox "角色授予失败,错误信息如下:" & vbCrLf & err.Description, vbExclamation, gstrSysName
        err.Clear
        Exit Function
    End If
    ApplyOnePerson = True
End Function


Private Function CreateApptionUser(ByVal lngID As Long, ByVal strUser As String, ByVal strPwd As String, Optional ByVal blnNew As Boolean, Optional ByVal blnMsg As Boolean) As Boolean
 '功能：创建用户并授权，或修改用户并授权
 'lngID=人员ID
 'strUser=用户
 'strPwd=密码
 'blnNew=是否是新用户
 'blnMsg=是否消息提示
    Dim strError As String
    
    If Not blnNew Then
        gcnOracle.Execute "Delete From " & mstr所有者 & ".上机人员表 Where 用户名='" & strUser & "'"
    Else
        Call gobjRegister.CreateUser(gcnOracle, strUser, strPwd, strError)
        If strError = "" Then
            Call gclsBase.ExecuteCmdText("Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & strUser, Me.Caption, , True)
            Call AlterUserTableSpaces(gcnOracle, strUser)
        Else
            If blnMsg Then MsgBox "用户名或密码不符合要求。请检查该用户是否已经存在!" & vbCrLf & "错误信息如下:" & vbCrLf & strError, vbExclamation, gstrSysName
            If blnMsg Then Exit Function
        End If
    End If
    gcnOracle.Execute "Insert Into " & mstr所有者 & ".上机人员表(用户名,人员id) Values ('" & strUser & "'," & lngID & ")"
     If err.Number <> 0 Then
        err.Clear
        MsgBox "由于共享性原因，用户身份信息(部门、姓名)更改失败。" & vbNewLine & err.Description, vbExclamation, gstrSysName
    End If
    mstrCreateUserList = IIf(mstrCreateUserList = "", "", mstrCreateUserList & "，") & strUser
    CreateApptionUser = True
End Function

Private Sub ChangePerson(ByVal objNode As Node)
'功能：修改人员时界面数据处理
    Dim arrTmp As Variant
    
    txtNO.Tag = Val(Mid(objNode.Key, 2))
    arrTmp = Split(objNode.Text, "】")
    txtName.Text = Mid(objNode.Text, Len(arrTmp(0)) + 2)
    txtNO.Text = Mid(arrTmp(0), 2)
    lblNO.Tag = txtNO.Tag & "," & txtNO.Text
    lblName.Tag = txtName.Text
    arrTmp = Split(objNode.Parent.Text, "】")
    txtDept.Text = Mid(objNode.Parent.Text, Len(arrTmp(0)) + 2)
    txtDept.Tag = Mid(objNode.Parent.Key, 2)
    tvwPerson.Visible = False
    If txtUserName.Enabled Then
        If optMode(0).value Then
            txtUserName.Text = SpellCode(txtName.Text)
        Else
            If UCase(Left(txtNO.Text, 1)) >= "A" And UCase(Left(txtNO.Text, 1)) <= "Z" Then
                txtUserName.Text = txtNO.Text
            Else
                txtUserName.Text = "U" & txtNO.Text
            End If
        End If
        txtUserName.SetFocus
    Else
        If vsRole.Enabled Then vsRole.SetFocus
    End If
    Call LocateManPros(Val(txtNO.Tag))
    Call LocateWorkRange(Val(txtDept.Tag))
End Sub

