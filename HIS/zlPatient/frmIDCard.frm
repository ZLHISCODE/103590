VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.2#0"; "zlIDKind.ocx"
Begin VB.Form frmIDCard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "就诊卡发放"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIDCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt就诊卡 
      Height          =   360
      Left            =   4800
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   210
      TabIndex        =   16
      Top             =   4320
      Width           =   1500
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   3300
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   825
      Width           =   3300
      Begin VB.ComboBox cboNO 
         Height          =   360
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "热键:F12"
         Top             =   30
         Width           =   1560
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2685
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "热键：F8"
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2715
         TabIndex        =   31
         Top             =   30
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.PictureBox picFace 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   90
      ScaleHeight     =   2895
      ScaleWidth      =   7290
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1275
      Width           =   7290
      Begin VB.TextBox txtAudi 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4335
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1800
         Width           =   1665
      End
      Begin VB.TextBox txt操作员 
         Height          =   360
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2355
         Width           =   1575
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1259
         Width           =   1590
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1185
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   712
         Width           =   1575
      End
      Begin VB.CheckBox chkBilling 
         Caption         =   "记帐"
         Height          =   240
         Left            =   3495
         TabIndex        =   6
         Top             =   765
         Width           =   810
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   360
         Left            =   4335
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   705
         Width           =   1695
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "热键：F11"
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         Height          =   360
         Left            =   4350
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtOld 
         Height          =   360
         Left            =   5775
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1300
      End
      Begin VB.TextBox txtCardNO 
         BackColor       =   &H00EBFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4335
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox txtPass 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1185
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1806
         Width           =   1590
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   4335
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2355
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   360
         Left            =   1185
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "快捷键F4"
         Top             =   165
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;就|就诊卡|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl结算方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算方式"
         Height          =   240
         Left            =   3315
         TabIndex        =   35
         Top             =   765
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证密码"
         Height          =   240
         Left            =   3315
         TabIndex        =   34
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发卡人"
         Height          =   240
         Left            =   390
         TabIndex        =   33
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方式"
         Height          =   240
         Left            =   630
         TabIndex        =   28
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   240
         Left            =   630
         TabIndex        =   27
         Top             =   765
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发卡时间"
         Height          =   240
         Left            =   3315
         TabIndex        =   26
         Top             =   2415
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   240
         Left            =   630
         TabIndex        =   25
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   240
         Left            =   3810
         TabIndex        =   24
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   240
         Left            =   5280
         TabIndex        =   23
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblCardNO 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "卡号"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3720
         TabIndex        =   22
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   240
         Left            =   630
         TabIndex        =   21
         Top             =   1860
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   5775
      TabIndex        =   15
      ToolTipText     =   "热键:Esc"
      Top             =   4320
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   4155
      TabIndex        =   14
      Top             =   4320
      Width           =   1500
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   32
      Top             =   4815
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmIDCard.frx":0442
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl刷卡 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊卡"
      Height          =   240
      Left            =   3840
      TabIndex        =   36
      Top             =   880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "就诊卡发放单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   90
      TabIndex        =   19
      Top             =   210
      Width           =   7170
   End
End
Attribute VB_Name = "frmIDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit '要求变量声明
'''''说明：在发卡状态换切到退卡,可以刷卡(病人栏)或输入单据号确定
''''Public mbytInState As Byte '入：0=发卡(补卡、换卡),1-查看记录,2-退卡(单据号、卡号),只支持单据号确定
''''Public mblnViewCancel As Boolean '入：是否查看已退单据
''''Public mstrInNO As String '入：要查看或要退的单据号,从病人信息登记中调用退卡时为空
''''Public mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
''''
''''Private mblnUnLoad As Boolean
'''''(收费类别,收费细目ID,计算单位,收入项目ID,收入项目,收据费目,原价,现价,是否变价,科室标志)
''''Private mrs就诊卡 As ADODB.Recordset
''''Private mrsInfo As New ADODB.Recordset '保存病人信息
''''Private mlng磁卡ID As Long '从共用及自用就诊卡批次中选择的领用ID
''''Private mblnICCard As Boolean 'IC卡发卡,要同时填写病人信息的IC卡字段
''''
''''Private WithEvents mobjIDCard As clsIDCard
''''Private mobjICCard As Object
''''Private Enum IDKinds
''''    C0姓名 = 0
''''    C1医保号 = 1
''''    C2身份证号 = 2
''''    C3IC卡号 = 3
''''    C4门诊号 = 4
''''    C5就诊卡 = 5
''''End Enum
''''Private mint退卡模式 As Integer
''''Private mstr退卡验证 As String
''''
''''Private Sub IDKind_Click()
''''    If IDKind.IDKind = IDKinds.C3IC卡号 Then
''''        If mobjICCard Is Nothing Then
''''            Set mobjICCard = CreateObject("zlICCard.clsICCard")
''''            Set mobjICCard.gcnOracle = gcnOracle
''''        End If
''''        If Not mobjICCard Is Nothing Then
''''            txtPatient.Text = mobjICCard.Read_Card()
''''            If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub lblCardNO_Click()
''''    If txtCardNO.Enabled = False Or txtCardNO.Locked Then Exit Sub
''''    If mobjICCard Is Nothing Then
''''        Set mobjICCard = CreateObject("zlICCard.clsICCard")
''''        Set mobjICCard.gcnOracle = gcnOracle
''''    End If
''''    If Not mobjICCard Is Nothing Then
''''        txtCardNO.Text = mobjICCard.Read_Card()
''''        If txtCardNO.Text <> "" Then mblnICCard = True
''''    End If
''''End Sub
''''
''''Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
''''                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
''''    Dim lngPreIDKind As Long
''''
''''    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
''''        lngPreIDKind = IDKind.IDKind
''''        IDKind.IDKind = IDKinds.C2身份证号
''''        txtPatient.Text = strID
''''        Call txtPatient_KeyPress(vbKeyReturn)
''''        IDKind.IDKind = lngPreIDKind
''''    End If
''''End Sub
''''
''''Private Sub cboStyle_Click()
''''    If Me.Visible Then
''''        If cboStyle.ListIndex = 2 Then '如果换卡,则不收取病人就诊卡费用
''''            txtMoney.Text = "0.00"
''''            txtMoney.Locked = True
''''        ElseIf Val(txtMoney.Text) = 0 Then
''''            If mrs就诊卡!是否变价 = 1 And cboStyle.ListIndex = 0 Then  '发卡按最低限价
''''               txtMoney.Text = Format(mrs就诊卡!缺省价格, "0.00")
''''            Else
''''                txtMoney.Text = Format(mrs就诊卡!现价, "0.00")
''''            End If
''''            txtMoney.Locked = Not (mrs就诊卡!是否变价 = 1)
''''            txtMoney.TabStop = (mrs就诊卡!是否变价 = 1)
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub cboStyle_KeyPress(KeyAscii As Integer)
''''    Dim lngIdx As Long
''''    If KeyAscii = 13 And cboStyle.ListIndex <> -1 Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    ElseIf KeyAscii = 13 And cboStyle.ListIndex = -1 Then
''''        Beep
''''    End If
''''    If cboStyle.Locked Then Exit Sub
''''    If SendMessage(cboStyle.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
''''    lngIdx = MatchIndex(cboStyle.hwnd, KeyAscii)
''''    If lngIdx <> -2 Then cboStyle.ListIndex = lngIdx
''''End Sub
''''
''''Private Sub cbo结算方式_KeyPress(KeyAscii As Integer)
''''    Dim lngIdx As Long
''''    If KeyAscii = 13 And cbo结算方式.ListIndex <> -1 Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    ElseIf KeyAscii = 13 And cbo结算方式.ListIndex = -1 Then
''''        Beep
''''    End If
''''    If cbo结算方式.Locked Then Exit Sub
''''    If SendMessage(cbo结算方式.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
''''    lngIdx = MatchIndex(cbo结算方式.hwnd, KeyAscii)
''''    If lngIdx <> -2 Then cbo结算方式.ListIndex = lngIdx
''''End Sub
''''
''''Private Sub chkBilling_Click()
''''    If chkBilling.Value = Checked Then
''''        cbo结算方式.Enabled = False
''''    ElseIf cbo结算方式.ListCount = 0 Then
''''        chkBilling.Value = Checked
''''        cbo结算方式.Enabled = False
''''    Else
''''        cbo结算方式.Enabled = True
''''    End If
''''    If Visible And txtPatient.Text <> "" And txtCardNO.Enabled Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub chkBilling_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub chkCancel_Click()
''''    sta.Panels(2).Text = ""
''''    If chkCancel.Value = Checked Then
''''        '按下
''''        chkCancel.ForeColor = &HFF&
''''        If mbytInState = 0 Then
''''            Set txtPatient.Container = Me
''''            txtPatient.Top = picFace.Top + txtSex.Top
''''            txtPatient.Left = picFace.Left + txtMoney.Left + IDKind.Width
''''            txtPatient.PasswordChar = "*"
''''            IDKind.Enabled = False
''''            txtPatient.Locked = True
''''        End If
''''        picFace.Enabled = False
''''        '清除相关界面和数据
''''        Call NewCard
''''        txtPatient.Text = ""
''''        '待输入退款的单据号
''''        cboNO.Text = "": cboNO.Tag = ""
''''        cboNO.Locked = False
''''        If cboNO.Visible Then cboNO.SetFocus
''''
''''        '问题28130、27929 by lesfeng 2010-02-26 发卡中退卡 显示
''''        If mbytInState = 0 Then
''''            If mint退卡模式 = 1 Or mint退卡模式 = 3 Then
''''                lbl刷卡.Visible = True
''''                txt就诊卡.Visible = True
''''                If txt就诊卡.Visible Then txt就诊卡.SetFocus
''''            End If
''''        End If
''''    Else
''''        '弹起
''''        chkCancel.ForeColor = 0
''''        If mbytInState = 0 Then
''''            Set txtPatient.Container = picFace
''''            txtPatient.Top = txtSex.Top
''''            txtPatient.Left = txtMoney.Left + IDKind.Width
''''            txtPatient.PasswordChar = ""
''''            IDKind.Enabled = True
''''            txtPatient.Locked = False
''''        End If
''''        picFace.Enabled = True
''''        '清除相关界面和数据
''''        Call NewCard
''''        txtPatient.Text = ""
''''        txtMoney.Text = Format(IIf(mrs就诊卡!是否变价 = 1, mrs就诊卡!缺省价格, mrs就诊卡!现价), "0.00")
''''        txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''        '新的一张结帐单
''''        cboNO.Locked = True
''''        txtPatient.SetFocus
''''
''''        '问题28130、27929 by lesfeng 2010-02-26 发卡中退卡 隐藏
''''        If mbytInState = 0 Then
''''            If mint退卡模式 = 1 Or mint退卡模式 = 3 Then
''''                lbl刷卡.Visible = False
''''                txt就诊卡.Visible = False
''''                txt就诊卡.Text = ""
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub cmdCancel_Click()
''''    If mbytInState = 0 And gblnOK Then
''''        If chkCancel.Value = Checked Then
''''            If MsgBox("确实要放弃退卡退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''        Else
''''            If mrsInfo.State = adStateOpen Then
''''                If glngSys Like "8??" Then
''''                    If MsgBox("该客户的会员卡尚未发放,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''                Else
''''                    If MsgBox("该病人的就诊卡尚未发放,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''                End If
''''            Else
''''                If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
''''            End If
''''        End If
''''    End If
''''    Unload Me
''''End Sub
''''
''''Private Sub cmdHelp_Click()
''''ShowHelp App.ProductName, Me.hwnd, Me.Name
''''End Sub
''''
''''Private Sub cmdOK_Click()
''''    Dim strNO As String, strSQL As String, strCard As String, strICCard As String
''''    Dim i As Integer, strTmp As String
''''    Dim str验证卡号 As String
''''    Dim blnTrans As Boolean
''''
''''    If chkCancel.Value = Checked Then
''''        '退款
''''        If cboNO.Tag = "" Then
''''            If glngSys Like "8??" Then
''''                MsgBox "该会员卡发放记录未正确读取,不能退卡！", vbExclamation, gstrSysName
''''                '问题31345 by lesfeng 2010-07-08
''''                Exit Sub
''''            Else
''''                MsgBox "该就诊卡发放记录未正确读取,不能退卡！", vbExclamation, gstrSysName
''''                '问题31345 by lesfeng 2010-07-08
''''                Exit Sub
''''            End If
''''            '问题28130、27929 by lesfeng 2010-02-26 退卡验证姓名
''''            If mint退卡模式 = 0 Then
''''                '问题31345 by lesfeng 2010-07-08
''''                If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus: Exit Sub
''''            Else
''''                '问题31345 by lesfeng 2010-07-08
''''                If txt就诊卡.Enabled And txt就诊卡.Visible Then txt就诊卡.SetFocus: Exit Sub
''''            End If
''''        End If
''''        '问题28130、27929 by lesfeng 2010-02-26 退卡验证姓名
'''''        If mint退卡模式 <> 0 Then
''''        If (mint退卡模式 = 2 Or mint退卡模式 = 3) And txt就诊卡.Visible Then
''''            str验证卡号 = Trim(txtCardNO.Text)
''''            If mstr退卡验证 = "" Or str验证卡号 <> mstr退卡验证 Then
''''                MsgBox "退卡验证失败，请核对实际卡号与当前单据卡号是否一致！", vbExclamation, gstrSysName
''''                If txt就诊卡.Enabled And txt就诊卡.Visible Then txt就诊卡.SetFocus
''''                Exit Sub
''''            End If
''''        End If
''''
''''        If Not CancelBill(cboNO.Tag) Then '退款(以cboNO.Tag=NO)
''''            MsgBox "退卡操作失败，请重试该操作！", vbExclamation, gstrSysName
''''            Exit Sub
''''        End If
''''
''''        mstr退卡验证 = ""
''''
''''        If mbytInState <> 2 Then
''''            chkCancel.Value = Unchecked '(并激活事件)
''''        Else
''''            gblnOK = True
''''            Unload Me: Exit Sub '退卡模式操作后退出
''''        End If
''''    Else
''''        '新就诊卡发放记录
''''        If mrsInfo.State = adStateClosed Then
''''            If glngSys Like "8??" Then
''''                MsgBox "没有确定要发放会员卡的客户,该操作不能继续！", vbExclamation, gstrSysName
''''            Else
''''                MsgBox "没有确定要发放就诊卡的病人,该操作不能继续！", vbExclamation, gstrSysName
''''            End If
''''            txtPatient.SetFocus: Exit Sub
''''        End If
''''
''''        If chkBilling.Value = 0 Then
''''            If cbo结算方式.ListCount = 0 Then
''''                MsgBox "没有可选结算方式,请使用记帐发卡或先到结算方式管理中进行设置！", vbExclamation, gstrSysName
''''                cbo结算方式.SetFocus: Exit Sub
''''            ElseIf cbo结算方式.ListIndex = -1 Then
''''                MsgBox "不记帐发卡必须确定结算方式！", vbExclamation, gstrSysName
''''                cbo结算方式.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        '金额检查:   cboStyle.ListIndex <> 2:问题:25930
''''        If mrs就诊卡!是否变价 = 1 And cboStyle.ListIndex <> 2 Then
''''            If mrs就诊卡!现价 <> 0 And Abs(CCur(txtMoney.Text)) > Abs(mrs就诊卡!现价) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "会员卡金额绝对值不能高于最高限价:" & Format(Abs(mrs就诊卡!现价), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "就诊卡金额绝对值不能高于最高限价:" & Format(Abs(mrs就诊卡!现价), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                txtMoney.SetFocus: Exit Sub
''''            End If
''''            If mrs就诊卡!原价 <> 0 And Abs(CCur(txtMoney.Text)) < Abs(mrs就诊卡!原价) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "会员卡金额绝对值不能低于最低限价:" & Format(Abs(mrs就诊卡!原价), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "就诊卡金额绝对值不能低于最低限价:" & Format(Abs(mrs就诊卡!原价), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                txtMoney.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        '发卡类型检查
''''        If cboStyle.ListIndex = -1 Then
''''            MsgBox "请确定发卡类型！", vbExclamation, gstrSysName
''''            cboStyle.SetFocus: Exit Sub
''''        End If
''''        If IsNull(mrsInfo!就诊卡号) Then
''''            If cboStyle.ListIndex <> 0 Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "无会员卡的客户只能选择发卡！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "无卡病人只能选择发卡！", vbExclamation, gstrSysName
''''                End If
''''                cboStyle.SetFocus: Exit Sub
''''            End If
''''        Else
''''            If cboStyle.ListIndex = 0 Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "持会员卡的客户只能选择补卡或换卡！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "持卡病人只能选择补卡或换卡！", vbExclamation, gstrSysName
''''                End If
''''                cboStyle.SetFocus: Exit Sub
''''            End If
''''        End If
''''
''''        If Not IsDate(txtDate.Text) Then
''''            MsgBox "请输入正确的发卡时间！", vbExclamation, gstrSysName
''''            txtDate.SetFocus: Exit Sub
''''        End If
''''        If txtCardNO.Text = "" Then
''''           MsgBox "请刷卡确定卡号！", vbExclamation, gstrSysName
''''           txtCardNO.SetFocus: Exit Sub
''''        End If
''''
''''        If txtPass.Text <> txtAudi.Text Then
''''            MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
''''            txtPass.Text = "": txtAudi.Text = ""
''''            txtPass.SetFocus: Exit Sub
''''        End If
''''
''''        '保存前检查就诊卡是否有，是否在范围内
''''        If gblnBill磁卡 Then
''''            mlng磁卡ID = CheckUsedBill(5, IIf(mlng磁卡ID > 0, mlng磁卡ID, glng磁卡ID), UCase(txtCardNO.Text))
''''            If mlng磁卡ID <= 0 Then
''''                Select Case mlng磁卡ID
''''                    Case 0 '操作失败
''''                    Case -1
''''                        If glngSys Like "8??" Then
''''                            MsgBox "你已没有自用及共用的会员卡,请先在本地设置共用批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "你已没有自用及共用的就诊卡,请先在本地设置共用批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                        End If
''''                    Case -2
''''                        If glngSys Like "8??" Then
''''                            MsgBox "本地共用的会员卡已用完,请重新设置本地共用会员卡批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "本地共用的就诊卡已用完,请重新设置本地共用就诊卡批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                        End If
''''                    Case -3
''''                        If glngSys Like "8??" Then
''''                            MsgBox "该张会员卡号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
''''                        Else
''''                            MsgBox "该张就诊卡号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
''''                        End If
''''                        txtCardNO.SetFocus
''''                End Select
''''                Exit Sub
''''            End If
''''        End If
''''
''''        '存盘
''''        If CByte(cboStyle.ListIndex) = 2 Then
''''            '补卡,相当于重打,不产生费用,但要获取原单据号
''''            strNO = GetNOFromCard(mrsInfo!就诊卡号)
''''            If strNO = "" Then
''''                MsgBox "没有发现该病人以前的发卡记录，不能补卡！", vbExclamation, gstrSysName
''''                Exit Sub
''''            End If
''''        Else
''''            '发卡,换卡新产生费用
''''            strNO = zldatabase.GetNextNo(16)
''''        End If
''''
''''        strCard = UCase(txtCardNO.Text)
''''        strICCard = IIf(mblnICCard, strCard, "")
''''
''''        strSQL = SaveIDCard(CByte(cboStyle.ListIndex), strNO, mrsInfo!病人ID, mrsInfo!主页ID, _
''''            IIf(mrsInfo!病区ID = 0, UserInfo.部门ID, mrsInfo!病区ID), _
''''            IIf(mrsInfo!科室ID = 0, UserInfo.部门ID, mrsInfo!科室ID), _
''''            IIf(mrsInfo!住院号 = 0, mrsInfo!门诊号, mrsInfo!住院号), IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别), IIf(IsNull(mrsInfo!就诊卡号), "", mrsInfo!就诊卡号), _
''''            mrsInfo!姓名, IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别), IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄), _
''''            strCard, txtPass.Text, IIf(mrs就诊卡!是否变价 = 0, mrs就诊卡!现价, CCur(txtMoney.Text)), CCur(txtMoney.Text), IIf(Not cbo结算方式.Enabled, "", NeedName(cbo结算方式.Text)), _
''''            CDate(txtDate.Text), mlng磁卡ID, mrs就诊卡, strICCard)
''''
''''        On Error GoTo errH
''''        gcnOracle.BeginTrans: blnTrans = True
'''''        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'''''        gcnOracle.Execute strSQL, , adCmdStoredProc
'''''        Call SQLTest
''''        zldatabase.ExecuteProcedure strSQL, Me.Caption
''''        gcnOracle.CommitTrans: blnTrans = False
''''
''''
''''        On Error GoTo 0
''''        '刘兴洪:24662
''''        Dim strOutPut As String
''''        'If Not mobjICCard Is Nothing Then
''''        Call zlExcuteUploadSwap(Val(Nvl(mrsInfo!病人ID)), strOutPut, mobjICCard)
''''        'End If
''''
''''        '打印(非记帐,非换卡时)
''''
''''        '加入单据历史记录(所有类型单据)
''''        For i = 0 To cboNO.ListCount - 1
''''            strTmp = strTmp & "," & cboNO.List(i)
''''        Next
''''        strTmp = strNO & strTmp
''''        cboNO.Clear
''''        For i = 0 To UBound(Split(strTmp, ","))
''''            cboNO.AddItem Split(strTmp, ",")(i)
''''            If i = 9 Then Exit For '只显示10个
''''        Next
''''    End If
''''
''''    gblnOK = True
''''
''''    '清除相关界面和数据
''''    Call NewCard(False)
''''    mblnICCard = False  '不能放在newcard中,因为可能先读卡再读病人
''''    txtPatient.Text = ""
''''    If Val(txtMoney.Text) = 0 Then txtMoney.Text = Format(mrs就诊卡!现价, "0.00")
''''    txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''
''''    '保存后检查就诊卡是否有
''''    '不自动产生新卡号(不同于票据)
''''    If gblnBill磁卡 Then
''''        mlng磁卡ID = CheckUsedBill(5, IIf(mlng磁卡ID > 0, mlng磁卡ID, glng磁卡ID))
''''        If mlng磁卡ID <= 0 Then
''''            Select Case mlng磁卡ID
''''                Case 0 '操作失败
''''                Case -1
''''                    If glngSys Like "8??" Then
''''                        MsgBox "你已没有自用及共用的会员卡,请先在本地设置共用批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "你已没有自用及共用的就诊卡,请先在本地设置共用批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                    End If
''''                Case -2
''''                    If glngSys Like "8??" Then
''''                        MsgBox "本地共用的会员卡已用完,请重新设置本地共用会员卡批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "本地共用的就诊卡已用完,请重新设置本地共用就诊卡批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                    End If
''''            End Select
''''            Exit Sub
''''        End If
''''    End If
''''
''''    txtPatient.SetFocus
''''    Exit Sub
''''errH:
''''    If blnTrans Then gcnOracle.RollbackTrans
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Sub
''''
''''Private Sub Form_Activate()
''''    If mbytInState = 2 Then
''''        '问题28130、27929 by lesfeng 2010-02-26 发卡中退卡 显示
''''        If mint退卡模式 = 2 Or mint退卡模式 = 3 Then
''''            lbl刷卡.Visible = True
''''            txt就诊卡.Visible = True
''''            If txt就诊卡.Visible Then txt就诊卡.SetFocus
''''        Else
''''            cmdOK.SetFocus
''''        End If
''''    ElseIf mbytInState = 1 Then
''''        cmdCancel.SetFocus
''''    End If
''''End Sub
''''
''''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''    On Error Resume Next
''''    Select Case KeyCode
''''        Case vbKeyF4
''''            If Shift = vbCtrlMask Then
''''                If IDKind.Enabled Then IDKind.IDKind = IDKinds.C3IC卡号: Call IDKind_Click
''''            ElseIf Me.ActiveControl Is txtPatient Then
''''                If IDKind.Enabled Then
''''                    If Shift = vbShiftMask Then
''''                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
''''                    Else
''''                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
''''                    End If
''''                End If
''''            End If
''''        Case vbKeyF8
''''            If chkCancel.Visible And picNo.Enabled Then
''''                chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
''''            End If
''''        Case vbKeyF11
''''            If Not txtPatient.Locked Then txtPatient.SetFocus
''''        Case vbKeyF12
''''            If Not cboNO.Locked And picNo.Enabled Then cboNO.SetFocus
''''        Case vbKeyEscape
''''            cmdCancel_Click
''''    End Select
''''End Sub
''''
''''Private Sub Form_KeyPress(KeyAscii As Integer)
''''    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''End Sub
''''
''''Private Sub Form_Load()
''''    If glngSys Like "8??" Then
''''        Caption = "会员卡发放"
''''        lblPatient.Caption = "客户"
''''        lblTitle.Caption = gstrUnitName & "会员卡发放单"
''''        chkBilling.Visible = False
''''        lbl结算方式.Visible = True
''''    Else
''''        lblTitle.Caption = gstrUnitName & "就诊卡发放单"
''''    End If
''''
''''    mblnUnLoad = False
''''    '问题28130、27929 by lesfeng 2010-02-26
''''    mint退卡模式 = Val(zldatabase.GetPara("退卡刷卡", glngSys, glngModul))
'''''    mint退卡模式 = 3
''''    mstr退卡验证 = ""
''''    '就诊卡领用检查
''''    If mbytInState = 0 And gblnBill磁卡 Then
''''        mlng磁卡ID = CheckUsedBill(5, IIf(mlng磁卡ID > 0, mlng磁卡ID, glng磁卡ID))
''''        If mlng磁卡ID <= 0 Then
''''            Select Case mlng磁卡ID
''''                Case 0 '操作失败
''''                Case -1
''''                    If glngSys Like "8??" Then
''''                        MsgBox "你已没有自用及共用的会员卡,请先在本地设置共用批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "你已没有自用及共用的就诊卡,请先在本地设置共用批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                    End If
''''                Case -2
''''                    If glngSys Like "8??" Then
''''                        MsgBox "本地共用的会员卡已用完,请重新设置本地共用会员卡批次或领用一批会员卡！", vbExclamation, gstrSysName
''''                    Else
''''                        MsgBox "本地共用的就诊卡已用完,请重新设置本地共用就诊卡批次或领用一批就诊卡！", vbExclamation, gstrSysName
''''                    End If
''''            End Select
''''            mblnUnLoad = True: Unload Me: Exit Sub
''''        End If
''''    End If
''''
''''    RestoreWinState Me
''''
''''    Call InitFace
''''    If mblnUnLoad Then: Unload Me: Exit Sub
''''
''''    Call RaisEffect(picFace, -1)
''''End Sub
''''
''''Private Sub Form_Unload(Cancel As Integer)
''''    mblnICCard = False
''''    mbytInState = 0
''''    mblnViewCancel = False
''''    mstrInNO = ""
''''    mlng磁卡ID = 0
''''    mblnUnLoad = False
''''    mblnNOMoved = False
''''    Set mobjICCard = Nothing
''''    If Not mobjIDCard Is Nothing Then
''''        Call mobjIDCard.SetEnabled(False)
''''        Set mobjIDCard = Nothing
''''    End If
''''End Sub
''''
''''Private Sub InitFace()
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim i As Integer, strSQL As String
''''    gblnOK = False
''''
''''    If gblnShowCard Then txtCardNO.PasswordChar = ""
''''
''''    IDKind.Enabled = mbytInState = 0
''''
''''    Select Case mbytInState
''''        Case 0 '发卡
''''            Set mrsInfo = New ADODB.Recordset
''''            Set mobjIDCard = New clsIDCard
''''
''''            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''
''''            cboStyle.AddItem "1-发卡"
''''            cboStyle.AddItem "2-补卡"
''''            cboStyle.AddItem "3-换卡" '不收取金额,不处理费用,仅处理票据
''''            cboStyle.ListIndex = 0
''''            chkBilling.Value = IIf(gbln记账 = True, 1, 0)
''''            txt操作员.Text = UserInfo.姓名
''''
''''            '结算方式
''''            strSQL = _
''''                "Select B.编码,B.名称,Nvl(A.缺省标志,0) as 缺省" & _
''''                " From 结算方式应用 A,结算方式 B" & _
''''                " Where A.应用场合='就诊卡' And B.名称=A.结算方式 And Nvl(B.性质,1) IN(1,2)" & _
''''                " Order by B.编码"
''''            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
''''
''''            If Not rsTmp.EOF Then
''''                For i = 1 To rsTmp.RecordCount
''''                    cbo结算方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
''''                    If rsTmp!缺省 = 1 Then
''''                        cbo结算方式.ListIndex = cbo结算方式.NewIndex
''''                        cbo结算方式.ItemData(cbo结算方式.NewIndex) = 1
''''                    End If
''''                    rsTmp.MoveNext
''''                Next
''''                If cbo结算方式.ListIndex = -1 Then cbo结算方式.ListIndex = 0
''''            Else
''''                '无结算方式只能记帐发卡
''''                If glngSys Like "8??" Then
''''                    MsgBox "会员卡场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
''''                    mblnUnLoad = True: Exit Sub
''''                Else
''''                    MsgBox "就诊卡场合没有可用的结算方式，只能使用记帐方式发卡。", vbInformation, gstrSysName
''''                End If
''''                chkBilling.Value = 1
''''                chkBilling.Enabled = False
''''                cbo结算方式.Enabled = False
''''            End If
''''
''''            '就诊卡费用
''''            Set mrs就诊卡 = GetSpecialInfo("就诊卡")
''''            If Not mrs就诊卡 Is Nothing Then
''''                If Not mrs就诊卡.EOF Then
''''                    txtMoney.Locked = Not (mrs就诊卡!是否变价 = 1)
''''                    txtMoney.TabStop = (mrs就诊卡!是否变价 = 1)
''''                    If mrs就诊卡!是否变价 = 1 Then
''''                        txtMoney.Text = Format(mrs就诊卡!缺省价格, "0.00")
''''                    Else
''''                        txtMoney.Text = Format(mrs就诊卡!现价, "0.00")
''''                    End If
''''                End If
''''            Else
''''                If glngSys Like "8??" Then
''''                    MsgBox "尚未设置会员卡收费信息，请先到药店运行参数中设置！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "尚未设置就诊卡收费信息，请先到基础参数设置中处理！", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''        Case 1 '预览
''''            chkCancel.Visible = False
''''            If mblnViewCancel Then lblFlag.Visible = True
''''            picNo.Enabled = False
''''            picFace.Enabled = False
''''            cmdOK.Visible = False
''''
''''            cmdCancel.Caption = "退出(&X)"
''''
''''            If Not ReadBill(mstrInNO) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "不能正确读取该会员卡发放记录！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "不能正确读取该就诊卡发放记录！", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''        Case 2  '退卡
''''            chkCancel.Value = Checked '同时激活事件
''''            picFace.Enabled = False
''''            picNo.Enabled = False
''''
''''            If Not ReadBill(mstrInNO) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "不能正确读取该会员卡发放记录！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "不能正确读取该就诊卡发放记录！", vbExclamation, gstrSysName
''''                End If
''''                mblnUnLoad = True: Exit Sub
''''            End If
''''            '问题28130、27929 by lesfeng 2010-02-26
''''            If mint退卡模式 = 2 Or mint退卡模式 = 3 Then
''''                lbl刷卡.Visible = True
''''                txt就诊卡.Visible = True
''''                If txt就诊卡.Visible Then txt就诊卡.SetFocus
''''            End If
''''
''''    End Select
''''End Sub
''''
''''Private Sub txtAudi_GotFocus()
''''    SelAll txtAudi
''''    If glngSys Like "8??" Then
''''        sta.Panels(2) = "请客户再次输入相同的的密码！"
''''    Else
''''        sta.Panels(2) = "请病人再次输入相同的的密码！"
''''    End If
''''End Sub
''''
''''Private Sub txtAudi_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then
''''        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''    Else
''''        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''    End If
''''End Sub
''''
''''Private Sub txtCardNO_GotFocus()
''''    SelAll txtCardNO
''''    sta.Panels(2) = "请将磁卡从刷卡器上轻轻划过！"
''''    Call Beep: Beep
''''End Sub
''''
''''Private Sub txtCardNO_KeyPress(KeyAscii As Integer)
'''''    If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
''''    KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
''''
''''    If KeyAscii <> 13 Then
''''        If Len(txtCardNO.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Then
''''            txtCardNO.Text = txtCardNO.Text & Chr(KeyAscii)
''''            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''        End If
''''    Else
''''        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
''''    End If
''''End Sub
''''
''''Private Sub txtCardNO_LostFocus()
''''    sta.Panels(2) = ""
''''End Sub
''''
''''Private Sub txtDate_GotFocus()
''''    txtDate.SelStart = 8
''''    txtDate.SelLength = Len(txtDate.Text) - 8
''''End Sub
''''
''''Private Sub txtDate_KeyPress(KeyAscii As Integer)
''''    If IsDate(txtDate.Text) And KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
''''End Sub
''''
''''Private Sub txtMoney_GotFocus()
''''    If Not txtMoney.Locked Then SelAll txtMoney
''''End Sub
''''
''''Private Sub txtMoney_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 And txtMoney.Text <> "" Then
''''        Call zlCommFun.PressKey(vbKeyTab)
''''    Else
''''        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
''''    End If
''''End Sub
''''
''''Private Sub txtMoney_Validate(Cancel As Boolean)
''''    If IsNumeric(txtMoney.Text) Then txtMoney.Text = Format(txtMoney.Text, "0.00")
''''    '金额检查
''''    If mbytInState = 0 Then
''''        If cboStyle.ListIndex = 2 Then Exit Sub
''''        If mrs就诊卡!是否变价 = 1 Then
''''            If mrs就诊卡!现价 <> 0 And Abs(CCur(txtMoney.Text)) > Abs(mrs就诊卡!现价) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "会员卡金额绝对值不能高于最高限价:" & Format(Abs(mrs就诊卡!现价), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "就诊卡金额绝对值不能高于最高限价:" & Format(Abs(mrs就诊卡!现价), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                SelAll txtMoney: Cancel = True: Exit Sub
''''            End If
''''
''''            If mrs就诊卡!原价 <> 0 And Abs(CCur(txtMoney.Text)) < Abs(mrs就诊卡!原价) Then
''''                If glngSys Like "8??" Then
''''                    MsgBox "会员卡金额绝对值不能低于最低限价:" & Format(Abs(mrs就诊卡!原价), "0.00"), vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "就诊卡金额绝对值不能低于最低限价:" & Format(Abs(mrs就诊卡!原价), "0.00"), vbExclamation, gstrSysName
''''                End If
''''                SelAll txtMoney: Cancel = True: Exit Sub
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub txtPass_GotFocus()
''''    SelAll txtPass
''''    If glngSys Like "8??" Then
''''        sta.Panels(2) = "请客户输入10位以内的密码！"
''''    Else
''''        sta.Panels(2) = "请病人输入10位以内的密码！"
''''    End If
''''End Sub
''''
''''Private Sub txtPass_KeyPress(KeyAscii As Integer)
''''    If KeyAscii = 13 Then
''''        KeyAscii = 0
''''        If txtPass.Text = "" And txtAudi.Text = "" Then
''''            cmdOK.SetFocus
''''        Else
''''            Call zlCommFun.PressKey(vbKeyTab)
''''        End If
''''    Else
''''        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
''''    End If
''''End Sub
''''
''''Private Sub txtPass_LostFocus()
''''    sta.Panels(2) = ""
''''End Sub
''''
''''Private Sub cboNO_GotFocus()
''''    If Not cboNO.Locked Then SelAll cboNO
''''End Sub
''''
''''Private Sub cboNO_KeyPress(KeyAscii As Integer)
''''    Dim strOper As String, vDate As Date
''''
''''    If cboNO.Locked Then Exit Sub
''''
''''    '转换成大写(汉字不可处理)
''''    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''
''''    '第一位可以输入字母,其它位不行
''''    If KeyAscii <> 13 Then
''''        Call SetNOInputLimit(cboNO, KeyAscii)
''''    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
''''        cboNO.Text = GetFullNO(cboNO.Text, 16)
''''
''''        '是否已转入后备数据表中
''''        If zldatabase.NOMoved("住院费用记录", cboNO.Text, , "5") Then
''''            If Not ReturnMovedExes(cboNO.Text, 5, Me.Caption) Then Exit Sub
''''            mblnNOMoved = False
''''        End If
''''
''''        '单据权限
''''        If Not ReadBillInfo(2, cboNO.Text, 5, strOper, vDate) Then
''''            txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
''''        End If
''''        If Not BillOperCheck(8, strOper, vDate, "退卡") Then
''''            txtPatient.Text = "": cboNO.Text = "": cboNO.SetFocus: Exit Sub
''''        End If
''''
''''        '读取要退卡的记录(由NO)
''''        Select Case ReadBill(cboNO.Text)
''''            Case -1
''''                '问题28130、27929 by lesfeng 2010-02-26 发卡中退卡 显示
''''                If mint退卡模式 = 1 Or mint退卡模式 = 3 Then
''''                    If txt就诊卡.Visible Then txt就诊卡.SetFocus
''''                Else
''''                    cmdOK.SetFocus
''''                End If
''''            Case 0
''''                If glngSys Like "8??" Then
''''                    MsgBox "读取该会员卡发放记录失败！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "读取该就诊卡发放记录失败！", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''            Case 1
''''                If glngSys Like "8??" Then
''''                    MsgBox "该会员卡发放记录不存在！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "该就诊卡发放记录不存在！", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''            Case 2
''''                If glngSys Like "8??" Then
''''                    MsgBox "该会员卡发放记录已经退除！", vbExclamation, gstrSysName
''''                Else
''''                    MsgBox "该就诊卡发放记录已经退除！", vbExclamation, gstrSysName
''''                End If
''''                txtPatient.Text = "": cboNO.SetFocus
''''        End Select
''''    End If
''''End Sub
''''
''''Private Sub txtPatient_Change()
''''    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
''''End Sub
''''
''''Private Sub txtPatient_GotFocus()
''''    SelAll txtPatient
''''    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
''''End Sub
''''
''''Private Sub txtPatient_KeyPress(KeyAscii As Integer)
''''    Dim strNO As String
''''    Dim blnCard As Boolean
''''
''''    If txtPatient.Locked Then Exit Sub
''''    '特殊字符过滤在Form_KeyPress中进行
''''    '刷卡显示处理
''''    If chkCancel.Value = Checked Then txtPatient.PasswordChar = IIf(gblnShowCard, "", "*")
''''
''''    '-010或+010时IsNumeric(txtPatient.Text)仍是true,此时txtPatient.Text的长度不包括正在输入的字符
''''    If IDKind.IDKind = IDKinds.C0姓名 Then
''''        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)
''''    ElseIf IDKind.IDKind = IDKinds.C4门诊号 Then
''''        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
''''            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
''''        End If
''''    End If
''''
''''    If blnCard And Len(txtPatient.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtPatient.Text <> "" Then
''''        If KeyAscii <> 13 Then
''''            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
''''            txtPatient.SelStart = Len(txtPatient.Text)
''''        End If
''''        KeyAscii = 0
''''        sta.Panels(2) = ""
''''
''''        If chkCancel.Value = 1 Then
''''            If blnCard = True Then
''''                '1.读取要退卡的记录(由卡号)
''''                strNO = GetNOFromCard(txtPatient.Text)
''''                If strNO = "" Then
''''                    If glngSys Like "8??" Then
''''                        sta.Panels(2) = "不能读取客户发卡信息,请确定是否正确刷卡！"
''''                    Else
''''                        sta.Panels(2) = "不能读取病人发卡信息,请确定是否正确刷卡！"
''''                    End If
''''                    txtPatient.Text = "": txtPatient.SetFocus
''''                    Call Beep: Exit Sub
''''                End If
''''                '处理退卡
''''                cboNO.Text = strNO
''''                Call cboNO_KeyPress(13): Exit Sub
''''            Else
''''                sta.Panels(2) = "不能读取就卡信息,请确定是否正确刷卡！"
''''                txtPatient.Text = "": txtPatient.SetFocus
''''                Call Beep: Exit Sub
''''            End If
''''        Else
''''            '2.输入卡号,或住院号等病人标识,发,换,补卡
''''            Call NewCard(False)  '清除信息
''''            '读取病人信息
''''            If Not GetPatient(txtPatient.Text, blnCard) Then
''''                If glngSys Like "8??" Then
''''                    sta.Panels(2) = "没有发现该客户信息,可能未建档,请检查！"
''''                Else
''''                    sta.Panels(2) = "没有发现该病人信息,可能未建档,请检查！"
''''                End If
''''                txtPatient.Text = "": txtPatient.SetFocus
''''                Call Beep: Exit Sub
''''            End If
''''            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm")
''''            If cboStyle.ListIndex = 0 And mrs就诊卡!是否变价 = 1 Then
''''                txtMoney.Text = Format(mrs就诊卡!缺省价格, "0.00")
''''            Else
''''                txtMoney.Text = Format(mrs就诊卡!现价, "0.00")
''''
''''                If mrs就诊卡!是否变价 = 0 Then
''''                    txtMoney.Text = Format(GetActualMoney(Nvl(mrsInfo!费别), mrs就诊卡!收入项目ID, mrs就诊卡!现价, mrs就诊卡!收费细目ID), "0.00")
''''                End If
''''            End If
''''
''''            If Not IsNull(mrsInfo!就诊卡号) Then
''''                If MsgBox(IIf(glngSys Like "8??", "客户", "病人") & "已经持有" & IIf(glngSys Like "8??", "会员", "就诊") & "卡，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
''''                    Set mrsInfo = New ADODB.Recordset
''''                    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
''''                    txtPatient.SetFocus: Exit Sub
''''                Else
''''                    If cboStyle.ListCount > 2 And blnCard = True Then  '第三个
''''                        cboStyle.ListIndex = 2 '如果是划卡,则自动设置为换卡
''''                    Else
''''                        cboStyle.ListIndex = 1 '如果是输病人标识,则自动设为补卡
''''                    End If
''''                End If
''''            End If
''''            '处理新的发放记录
''''            txtPatient.Text = mrsInfo!姓名
''''            txtSex.Text = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
''''            txtOld.Text = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
''''            Call zlCommFun.PressKey(vbKeyTab) '定位到金额或记帐
''''            Exit Sub
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub NewCard(Optional blnMoney As Boolean = True)
''''    Set mrsInfo = New ADODB.Recordset
''''    cboNO.Text = ""
''''    If blnMoney Then txtMoney.Text = ""
''''    txtSex.Text = ""
''''    txtOld.Text = ""
''''    chkBilling.Value = IIf(gbln记账 = True, 1, 0)
''''    txtCardNO.Text = ""
''''    txtPass.Text = ""
''''    txtAudi.Text = ""
''''    txtDate.Text = "____-__-__ __:__"
''''    If cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0
''''End Sub
''''
''''Private Function GetPatient(ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'''''功能：读取病人信息
'''''参数：strInput=病人标识码(A-病人ID,B+住院号,C/床号,D*门诊号,G.挂号单号)
'''''返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close
''''    Dim strSQL As String, objRect As RECT
''''    On Error GoTo errH
''''
''''    '病人在院时用住院费别,否则用门诊费别
''''    strSQL = _
''''        " Select Rownum id,A.就诊卡号,A.病人ID,Nvl(B.主页ID,0) as 主页ID," & _
''''        " Nvl(A.当前病区ID,0) as 病区ID,Nvl(A.当前科室ID,0) as 科室ID," & _
''''        " A.姓名,A.性别,A.年龄,Nvl(A.住院号,0) as 住院号," & _
''''        " Nvl(A.门诊号,0) as 门诊号,Nvl(A.当前床号,0) as 床号," & _
''''        " Decode(A.当前科室ID,NULL,A.费别,B.费别) as 费别,A.家庭地址" & _
''''        " From 病人信息 A,病案主页 B" & _
''''        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+)" & _
''''        " And Nvl(A.住院次数,0)=B.主页ID(+) And Nvl(B.主页ID(+),0)<>0"
''''    If Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
''''        strSQL = strSQL & " And A.病人ID=[1] and '%'='%'"
''''    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(仅对住院病人)
''''        strSQL = strSQL & " And A.当前科室ID is Not NULL And A.住院号=[1] and '%'='%'"
''''    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
''''        strSQL = strSQL & " And A.当前科室ID is NULL And A.门诊号=[1] and '%'='%'"
''''    ElseIf blnCard = True Or IDKind.IDKind = IDKinds.C5就诊卡 Then
''''        strInput = UCase(strInput)
''''        strSQL = strSQL & " And A.就诊卡号=[2] and '%'='%'"  '就诊卡号
''''    Else '当作姓名
''''        Select Case IDKind.IDKind
''''            Case IDKinds.C0姓名
''''                strSQL = strSQL & " And A.姓名=[2] and '%'='%'"
''''            Case IDKinds.C1医保号
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.医保号=[2] and '%'='%'"
''''            Case IDKinds.C2身份证号
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.身份证号=[2] and '%'='%'"
''''            Case IDKinds.C3IC卡号
''''                strInput = UCase(strInput)
''''                strSQL = strSQL & " And A.IC卡号=[2] and '%'='%'"
''''            Case IDKinds.C4门诊号
''''                If Not IsNumeric(strInput) Then strInput = "0"
''''                strSQL = strSQL & " And A.当前科室ID is NULL And A.门诊号=[2] and '%'='%'"
''''        End Select
''''    End If
''''
''''    'Set mrsInfo = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
''''    objRect = GetControlRect(txtPatient.hwnd)
''''    Set mrsInfo = zldatabase.ShowSQLSelect(Me, strSQL, 0, "查病人信息", False, "病人ID", "", False, False, True, objRect.Left, objRect.Top, txtPatient.Height, False, True, False, Mid(strInput, 2), strInput)
''''
''''    If mrsInfo.State = adStateClosed Then
''''        Set mrsInfo = New ADODB.Recordset
''''    Else
''''        GetPatient = True
''''    End If
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''    Set mrsInfo = New ADODB.Recordset
''''End Function
''''
''''Private Function ReadBill(strNO As String) As Integer
'''''功能：由单据号读取并显示就诊卡发放记录
'''''返回：
'''''     -1:成功
'''''      0:失败
'''''      1:该记录不存在
'''''      2:该记录已经作废(当mblnViewCancel=False时有效)
''''    On Error GoTo errH
''''
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim strFullNO As String
''''
''''    strFullNO = GetFullNO(strNO, 16)
''''    '因为就诊卡费用的结帐ID可能是记帐发卡后结帐时产生的ID,
''''    '所以与预交记录联接时一定要加记录性质=5限制
''''    'by lesfeng 2010-03-08 处理A.* 问题
''''    gstrSQL = _
''''        " Select A.NO,A.姓名,A.性别,A.年龄,A.实际票号,A.附加标志,A.记录状态,A.实收金额,A.操作员姓名,A.发生时间,B.卡验证码,C.结算方式 " & _
''''        " From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A,病人信息 B," & _
''''        " (Select 结算方式,结帐ID From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 Where 记录性质=5 And NO=[1]) C" & _
''''        " Where A.结帐ID=C.结帐ID(+) And A.记录性质=5 And A.病人ID=B.病人ID And A.NO=[1] " & _
''''         IIf(mblnViewCancel, "And A.记录状态=3 ", "")
''''
''''    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
''''    If rsTmp.EOF Then ReadBill = 1: Exit Function
''''
''''    If Not mblnViewCancel And (rsTmp!记录状态 = 3 Or rsTmp!记录状态 = 2) Then
''''        ReadBill = 2: Exit Function
''''    End If
''''
''''    cboNO.Text = rsTmp!NO
''''    cboNO.Tag = rsTmp!NO
''''    txtPatient.Text = rsTmp!姓名
''''    txtPatient.PasswordChar = ""
''''    txtSex.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
''''    txtOld.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
''''
''''    If IsNull(rsTmp!结算方式) Then
''''        chkBilling.Value = Checked
''''    Else
''''        chkBilling.Value = Unchecked
''''        If cbo结算方式.ListCount = 0 Then
''''            cbo结算方式.AddItem rsTmp!结算方式
''''            cbo结算方式.ListIndex = 0
''''        Else
''''            cbo结算方式.ListIndex = GetCboIndex(cbo结算方式, rsTmp!结算方式)
''''        End If
''''    End If
''''    txtCardNO.Text = IIf(IsNull(rsTmp!实际票号), "", rsTmp!实际票号)
''''    txtPass.Text = IIf(IsNull(rsTmp!卡验证码), "", rsTmp!卡验证码)
''''    txtAudi.Text = txtPass.Text
''''
''''    If cboStyle.ListCount = 0 Then
''''        Select Case rsTmp!附加标志
''''            Case 0
''''                cboStyle.AddItem "发卡"
''''            Case 1
''''                cboStyle.AddItem "补卡"
''''            Case 2
''''                cboStyle.AddItem "换卡"
''''        End Select
''''        cboStyle.ListIndex = 0
''''    Else
''''        cboStyle.ListIndex = rsTmp!附加标志
''''    End If
''''    txtMoney.Text = Format(rsTmp!实收金额, "0.00")
''''    txt操作员.Text = rsTmp!操作员姓名
''''    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm")
''''    ReadBill = -1
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Function CancelBill(strNO As String) As Boolean
'''''功能：退除病人就诊卡费用记录
''''    Dim strSQL As String
''''    Dim blnTrans As Boolean
''''
''''    On Error GoTo errH
''''
''''    '调用过程"zl_就诊卡记录_Delete"
''''    strSQL = "zl_就诊卡记录_DELETE('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
''''
''''    gcnOracle.BeginTrans: blnTrans = True
''''
'''''    Call SQLTest(App.ProductName, Me.Caption, strSQL)   'SQLTest
'''''    gcnOracle.Execute strSQL, , adCmdStoredProc
'''''    Call SQLTest
''''    zldatabase.ExecuteProcedure strSQL, Me.Caption
''''
''''    gcnOracle.CommitTrans: blnTrans = False
''''
''''    CancelBill = True
''''    Exit Function
''''errH:
''''    If blnTrans Then gcnOracle.RollbackTrans
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Sub txtPatient_LostFocus()
''''    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
''''End Sub
''''
''''Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then
''''        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
''''        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
''''    End If
''''End Sub
''''
''''Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then
''''        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
''''    End If
''''End Sub
''''
''''Private Sub txt就诊卡_GotFocus()
''''    Call zlControl.TxtSelAll(txt就诊卡)
'''''    With txtApplyman
'''''        .SelStart = 0
'''''        .SelLength = Len(.Text)
'''''    End With
''''End Sub
''''
'''''问题28130、27929 by lesfeng 2010-02-26
''''Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
''''    Dim strCardNO As String
''''    KeyAscii = Asc(UCase(Chr(KeyAscii)))
''''    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
''''    If KeyAscii = 13 Then
''''        strCardNO = Trim(txt就诊卡)
''''        If mbytInState = 0 Then
''''            Call NewCard
''''            If ReadCardNo(strCardNO, 2) = -1 Then
''''                cmdOK.SetFocus
''''            Else
''''                Call zlControl.TxtSelAll(txt就诊卡)
'''''                txt就诊卡.SetFocus
''''                sta.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
''''            End If
''''        Else
''''            If ReadCardNo(strCardNO, 1) = -1 Then
''''                cmdOK.SetFocus
''''            Else
''''                Call zlControl.TxtSelAll(txt就诊卡)
'''''                txt就诊卡.SetFocus
''''                sta.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
''''            End If
''''        End If
''''    End If
''''End Sub
''''
''''Private Function ReadCardNo(ByVal strNO As String, ByVal intFlag As Integer) As Integer
'''''功能：刷卡验证就诊卡退卡姓名一致性及刷卡取数
'''''输入：strNO 卡号
'''''      intFlag 标志 1 验证 2 取数
'''''返回：
'''''     -1:成功
'''''      0:失败
'''''      1:该记录不存在
''''    On Error GoTo errH
''''
''''    Dim rsTmp As New ADODB.Recordset
''''    Dim strSQL As String
''''    Dim lng病人ID As Long
''''    Dim str单据号 As String
''''
''''    ReadCardNo = 0
''''
''''    strSQL = "select 就诊卡号,姓名,病人ID from 病人信息 where 就诊卡号 = [1]"
''''
''''    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
''''    If rsTmp.EOF Then ReadCardNo = 1: Exit Function
''''
''''    mstr退卡验证 = IIf(IsNull(rsTmp!就诊卡号), "", rsTmp!就诊卡号)
''''    lng病人ID = IIf(IsNull(rsTmp!病人ID), 0, rsTmp!病人ID)
''''    If intFlag = 1 Then
''''        ReadCardNo = -1
''''        rsTmp.Close
''''        Exit Function
''''    Else
''''        rsTmp.Close
''''        '获取就诊卡在费用中的No
''''        strSQL = _
''''        " Select A.NO,B.卡验证码 " & _
''''        " From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A,病人信息 B" & _
''''        " Where A.记录性质=5 And A.病人ID=B.病人ID And A.实际票号=[1] and A.病人ID = [2]" & _
''''         IIf(mblnViewCancel, " And A.记录状态=3 ", " And A.记录状态=1 ")
''''         '问题31841:" And A.记录状态=1 "
''''         Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng病人ID)
''''        If rsTmp.EOF Then ReadCardNo = 1: Exit Function
''''
''''        str单据号 = IIf(IsNull(rsTmp!NO), "", rsTmp!NO)
''''        If ReadBill(str单据号) = -1 Then
''''            ReadCardNo = -1
''''            rsTmp.Close
''''            Exit Function
''''        End If
''''    End If
''''    rsTmp.Close
''''    Exit Function
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Function
''''
''''Private Sub txt就诊卡_LostFocus()
''''    Dim strCardNO As String
''''
''''    strCardNO = Trim(txt就诊卡)
''''    If mbytInState = 0 Then
''''        Call NewCard
''''        If ReadCardNo(strCardNO, 2) = -1 Then
''''            cmdOK.SetFocus
''''        Else
'''''            txt就诊卡.SetFocus
''''            sta.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
''''        End If
''''    Else
''''        If ReadCardNo(strCardNO, 1) = -1 Then
''''            cmdOK.SetFocus
''''        Else
'''''            txt就诊卡.SetFocus
''''            sta.Panels(2) = "没有发现该就诊卡的信息,可能未建档,请检查！"
''''        End If
''''    End If
''''End Sub
