VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应用系统安装"
   ClientHeight    =   5085
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmAppCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   1
      Left            =   1305
      TabIndex        =   9
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.ComboBox cmbEnjoy 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1440
         Width           =   2610
      End
      Begin VB.Frame fraOwner 
         Caption         =   "新建所有者"
         Height          =   1755
         Left            =   585
         TabIndex        =   29
         Top             =   1935
         Width           =   4530
         Begin VB.CheckBox chkDBA 
            Caption         =   "授予DBA角色"
            Height          =   255
            Left            =   3030
            TabIndex        =   50
            Top             =   1215
            Width           =   1320
         End
         Begin VB.TextBox txtOwnerUsr 
            Height          =   300
            Left            =   810
            TabIndex        =   30
            Top             =   360
            Width           =   1890
         End
         Begin VB.TextBox txtOwnerPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   810
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   780
            Width           =   1890
         End
         Begin VB.TextBox txtOwnerLab 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   810
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   32
            Top             =   1200
            Width           =   1890
         End
         Begin VB.Label lblDBA 
            Caption         =   "可以根据管理习惯决定是否授予DBA角色"
            Height          =   660
            Left            =   3030
            TabIndex        =   49
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label lblNewUser 
            AutoSize        =   -1  'True
            Caption         =   "用户名"
            Height          =   180
            Left            =   210
            TabIndex        =   35
            Top             =   420
            Width           =   540
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "密码"
            Height          =   180
            Left            =   390
            TabIndex        =   34
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lblNewLab 
            AutoSize        =   -1  'True
            Caption         =   "验证"
            Height          =   180
            Left            =   390
            TabIndex        =   33
            Top             =   1260
            Width           =   360
         End
      End
      Begin VB.CheckBox chkEnjoy 
         Caption         =   "选择需共享系统(&S)"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   1500
         Width           =   1830
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第二步 设置本系统所有者"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   225
         Width           =   2595
      End
      Begin VB.Label lblNote 
         Caption         =   "    在系统中已经安装了其他产品，可选择现有系统共享(但必须知道所有者的密码)；也可以新建所有者，不与任何产品共享。"
         Height          =   585
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   5250
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   0
      Left            =   1305
      TabIndex        =   4
      Top             =   -120
      Width           =   6075
      Begin VB.Frame fraSys 
         Height          =   1005
         Left            =   570
         TabIndex        =   46
         Top             =   2340
         Width           =   4545
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "版本号："
            Height          =   180
            Left            =   210
            TabIndex        =   48
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblSysName 
            AutoSize        =   -1  'True
            Caption         =   "系统名："
            Height          =   180
            Left            =   210
            TabIndex        =   47
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdSetupFile 
         Caption         =   "选择(&S)…"
         Height          =   350
         Left            =   570
         TabIndex        =   5
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblSetupFile 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   570
         TabIndex        =   28
         Top             =   1650
         Width           =   4545
      End
      Begin VB.Label lbliniFile 
         AutoSize        =   -1  'True
         Caption         =   "应用安装配置文件"
         Height          =   180
         Left            =   570
         TabIndex        =   27
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblNote 
         Caption         =   "    应用系统的安装依赖于配置文件和与之相关的服务器创建脚本文件，请正确指定安装配置文件。"
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第一步 指定安装配置文件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   2
      Left            =   1305
      TabIndex        =   14
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraSpace 
         Height          =   2070
         Left            =   495
         TabIndex        =   55
         Top             =   1770
         Width           =   5055
         Begin VB.CheckBox chkLogin 
            Caption         =   "索引表空间记录日志"
            Height          =   270
            Index           =   0
            Left            =   3000
            TabIndex        =   61
            ToolTipText     =   "索引表空间是否产生日志，默认不产生日志"
            Top             =   1155
            Width           =   1920
         End
         Begin VB.ComboBox cboSpaceExtentType 
            Height          =   300
            Index           =   0
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   60
            ToolTipText     =   "AUTOALLOCATE 或 UNIFORM Size nM"
            Top             =   1620
            Width           =   1815
         End
         Begin VB.CheckBox chkSpaceExtd 
            Caption         =   "自动扩展"
            Height          =   270
            Index           =   0
            Left            =   1815
            TabIndex        =   59
            ToolTipText     =   "AUTOEXTEND ON Next (表空间大小/10)M"
            Top             =   1155
            Width           =   1065
         End
         Begin VB.TextBox txtSpaceExtentSize 
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   58
            Text            =   "1"
            Top             =   1620
            Width           =   255
         End
         Begin VB.TextBox txtSpaceFile 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   705
            TabIndex        =   57
            Top             =   675
            Width           =   4005
         End
         Begin VB.TextBox txtSpaceSize 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   705
            MaxLength       =   10
            TabIndex        =   56
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblSpaceExtend 
            AutoSize        =   -1  'True
            Caption         =   "区尺寸"
            Height          =   180
            Left            =   105
            TabIndex        =   67
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label lblTBS 
            Caption         =   "M"
            Height          =   255
            Left            =   2970
            TabIndex        =   66
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label txtSpaceName 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   0
            Left            =   705
            TabIndex        =   65
            Top             =   225
            Width           =   2145
         End
         Begin VB.Label lblSpaceName 
            AutoSize        =   -1  'True
            Caption         =   "名称"
            Height          =   180
            Left            =   285
            TabIndex        =   64
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblSpaceFile 
            AutoSize        =   -1  'True
            Caption         =   "文件"
            Height          =   180
            Left            =   285
            TabIndex        =   63
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lblSpaceSize 
            AutoSize        =   -1  'True
            Caption         =   "大小          M"
            Height          =   180
            Left            =   285
            TabIndex        =   62
            Top             =   1185
            Width           =   1350
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   465
         Width           =   5800
      End
      Begin MSComctlLib.TabStrip tbsSpace 
         Height          =   2520
         Left            =   405
         TabIndex        =   54
         ToolTipText     =   "创建的表空间类型为本地管理表空间(非ASSM)"
         Top             =   1440
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   4445
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.Label lblNote 
         Caption         =   "    系统需要建立以下表空间，为降低磁盘I/O冲突，最好根据服务器磁盘情况，将表空间分别建立在不同的磁盘上。"
         Height          =   405
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   675
         Width           =   5610
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第三步 数据存储空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5775
      TabIndex        =   0
      Top             =   4260
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4704
      Left            =   0
      ScaleHeight     =   4650
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.Image imgSetup 
         Height          =   3315
         Left            =   60
         Picture         =   "frmAppCreate.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   1545
      TabIndex        =   44
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   2490
      TabIndex        =   43
      Top             =   4860
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4695
      TabIndex        =   3
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   42
      Top             =   4704
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppCreate.frx":5B70
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9340
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "13:59"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   3
      Left            =   1305
      TabIndex        =   23
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CheckBox chkSelData 
         Caption         =   "选择安装数据(&S)"
         Height          =   210
         Left            =   765
         TabIndex        =   40
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.Frame fraSelData 
         Height          =   3090
         Left            =   585
         TabIndex        =   37
         Top             =   1140
         Width           =   4845
         Begin VB.PictureBox picXp 
            BorderStyle     =   0  'None
            Height          =   2760
            Left            =   75
            ScaleHeight     =   2760
            ScaleWidth      =   1845
            TabIndex        =   51
            Top             =   255
            Width           =   1845
            Begin VB.OptionButton optData 
               Caption         =   "可选数据分组0"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   52
               Top             =   15
               Width           =   1665
            End
            Begin VB.Label lblNoData 
               Caption         =   "该数据组不再细分为可选的数据项"
               Height          =   2730
               Left            =   60
               TabIndex        =   53
               Top             =   0
               Visible         =   0   'False
               Width           =   4410
            End
         End
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "全清"
            Height          =   315
            Left            =   3090
            TabIndex        =   39
            Top             =   210
            Width           =   1100
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "全选"
            Height          =   315
            Left            =   1965
            TabIndex        =   38
            Top             =   210
            Width           =   1100
         End
         Begin VB.ListBox lstData 
            Height          =   1950
            Index           =   0
            ItemData        =   "frmAppCreate.frx":6402
            Left            =   1950
            List            =   "frmAppCreate.frx":6404
            Style           =   1  'Checkbox
            TabIndex        =   41
            Top             =   525
            Visible         =   0   'False
            Width           =   2670
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblNote 
         Caption         =   "    为能更快使用，系统准备了部分应用数据，根据不同的使用情况，可以选择安装不同的数据组。"
         Height          =   405
         Index           =   3
         Left            =   225
         TabIndex        =   26
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第四步 安装数据选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   25
         Top             =   225
         Width           =   2145
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   4
      Left            =   1305
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblNextDo 
         AutoSize        =   -1  'True
         Caption         =   "    点击""完成""开始自动装载系统，或者""取消""终止系统装载，或""上一步""重新调整应用系统装载配置。"
         Height          =   360
         Left            =   225
         TabIndex        =   45
         Top             =   2025
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRegAudit 
         AutoSize        =   -1  'True
         Caption         =   "    由于还不具备该系统应用授权，虽然可以继续装载，但无法正常使用。"
         Height          =   360
         Left            =   225
         TabIndex        =   22
         Top             =   1335
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第五步 完成"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   165
         TabIndex        =   21
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    已经完成了对该系统装载的全部设置。"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   20
         Top             =   720
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmAppCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum 数据段
    sec表名 = 0             '含有大对象字段的数据表
    sec字段名 = 1           '大对象字段
    sec字段类型 = 2         'Long 或 Raw
    sec主键 = 3             '包括主键的字段名与字段值，由|分隔。如果主键是由多个字段组成，那依次列出
    sec操作方法 = 4         'Insert 或 Update
    sec文件名 = 5           '含有大对象内容的文件，不包含路径
End Enum
Private mstrIniPath      As String                 '安装配置文件目录
Private intDefSysCode   As String                 '系统编号
Private strDefSysName   As String                 '系统名称
Private strDefVersion   As String                 '版本号
Private strDefSpace   As String                   '表空间定义串
Private strDefUser      As String                 '新的缺省用户名
Private strDefData      As String                 '用户可选的数据

Private mstrExtSysCode  As String                  '要进行扩展的主系统的编号
Private mstrExtVersion  As String                  '要进行扩展的主系统的版本
Private mstrTbsPath As String                        '缺省表空间路径名称，根据历史表空间产生

Private objText As TextStream
Private mstrLogFile As String
Private mclsRunScript As clsRunScript  '脚本解析执行类
Private mfrmUpSys As frmAppUpgradeNew
Private intStep As Integer

Private mbln帐套 As Boolean    '本次安装是否是属于帐套安装
Private mlng帐套 As Long       '帐套号
Private mlst标准 As ListItem   '相对于要安装的帐套，这是提供标准管理数据的系统

Private mcnOwner As New ADODB.Connection
Private intCount As Integer, intItems As Integer
        
Private aryRow() As String
Private aryVal() As String


Private Sub cboSpaceExtentType_Click(Index As Integer)
    txtSpaceExtentSize(Index).Enabled = (cboSpaceExtentType(Index).ListIndex = 1)
    If txtSpaceExtentSize(Index).Enabled Then
        If MsgBox("本参数建议采用“自动分配区尺寸”选项，是否原还为默认值？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            cboSpaceExtentType(Index).ListIndex = 0
            txtSpaceExtentSize(Index).Enabled = (cboSpaceExtentType(Index).ListIndex = 1)
        End If
    End If
End Sub

Private Sub chkEnjoy_Click()
'是否共享安装
    Dim blnEnjoy As Boolean
    blnEnjoy = chkEnjoy.value = 1
    cmbEnjoy.Enabled = blnEnjoy
    If blnEnjoy Then
        fraOwner.Caption = "所有者"
        txtOwnerUsr.Text = cmbEnjoy.Tag
        txtOwnerUsr.Enabled = False
        txtOwnerLab.Enabled = False
        chkDBA.Enabled = False
    Else
        fraOwner.Caption = "新建所有者"
        txtOwnerUsr.Text = strDefUser
        txtOwnerUsr.Enabled = True
        txtOwnerLab.Enabled = True
        chkDBA.Enabled = True
        
        If fraSetup(1).Visible = False Then Exit Sub
        txtOwnerUsr.SetFocus
    End If
    '设置控件位置以及状态
    lblNewLab.Visible = Not blnEnjoy
    txtOwnerLab.Visible = Not blnEnjoy
    chkDBA.Visible = Not blnEnjoy
    lblDBA.Visible = Not blnEnjoy
    '设置控件位置
    txtOwnerUsr.Left = IIf(blnEnjoy, 1200, 810)
    txtOwnerPwd.Left = txtOwnerUsr.Left
    lblNewUser.Left = txtOwnerUsr.Left - lblNewUser.Width - 60
    lblNewPwd.Left = txtOwnerUsr.Left - lblNewPwd.Width - 60
    txtOwnerUsr.Top = IIf(blnEnjoy, 540, 360)
    lblNewUser.Top = txtOwnerUsr.Top + (txtOwnerUsr.Height - lblNewUser.Height) / 2
    txtOwnerPwd.Top = txtOwnerUsr.Top + txtOwnerUsr.Height + IIf(blnEnjoy, 240, 120)
    lblNewPwd.Top = txtOwnerPwd.Top + (txtOwnerPwd.Height - lblNewPwd.Height) / 2
    If mstrExtSysCode = "" Then txtOwnerPwd.SetFocus
End Sub

Private Sub chkSelData_Click()
'是否安装可选择数据
    Dim i As Integer
    Dim blnEnable As Boolean
    If chkSelData.value = 0 Then
        blnEnable = False
    Else
        blnEnable = True
    End If
    
    fraSelData.Enabled = blnEnable
    For i = optData.LBound To optData.UBound
        optData(i).Enabled = blnEnable
    Next
    For i = lstData.LBound To lstData.UBound
        lstData(i).Enabled = blnEnable
    Next
    
    cmdSelectAll.Enabled = blnEnable
    cmdClearAll.Enabled = blnEnable
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("安装未完成，真的取消吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    '如果独立安装时已创建用户(点了第二步)，则删除用户
    If chkEnjoy.value = 0 And txtOwnerUsr.Tag = "已创建用户" Then
    
        On Error Resume Next
        gstrSQL = "drop user " & txtOwnerUsr.Text
        gcnOracle.Execute gstrSQL
    End If
    
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdSelectAll_Click()
    Dim lngIndex As Long, lngCount As Long
    
    For lngIndex = lstData.LBound To lstData.UBound
        With lstData(lngIndex)
            If .Visible = True Then
                For lngCount = 0 To .ListCount - 1
                    .Selected(lngCount) = True
                Next
                
                .Refresh
            End If
        End With
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim lngIndex As Long, lngCount As Long
    
    For lngIndex = lstData.LBound To lstData.UBound
        With lstData(lngIndex)
            If .Visible = True Then
                For lngCount = 0 To .ListCount - 1
                    .Selected(lngCount) = False
                Next
                
                .Refresh
            End If
        End With
    Next
End Sub

Private Sub cmdSetupFile_Click()
    With frmMDIMain.dlgMain
        .InitDir = App.Path
        .FileName = lblSetupFile.Caption
        .DialogTitle = "选择应用安装配置文件"
        .Filter = "(应用安装配置文件)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblSetupFile.Caption = .FileName
        End If
    End With
    If ChkSetupFile(True) = False Then
        mbln帐套 = False
        lblSetupFile.Caption = ""
        cmdSetupFile.SetFocus
    End If

End Sub

Private Sub cmdNext_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    Dim strError As String
    
    SetPromptText ""
    If fraSetup(0).Visible Then
        '------------------------------------------------------------
        '第一步：
        '------------------------------------------------------------
        If Trim(lblSetupFile.Caption) = "" Then
            MsgBox "未正确选择服务器安装配置文件，不能继续。", vbExclamation, gstrSysName
            cmdSetupFile.SetFocus
            Exit Sub
        End If
        
        '------------------------------
        fraSetup(0).Visible = False
        fraSetup(1).Visible = True
        cmdPrevious.Enabled = True
        If cmbEnjoy.ListCount > 0 Then cmbEnjoy.ListIndex = 0
    
    ElseIf fraSetup(1).Visible Then
        '------------------------------------------------------------
        '第二步：
        '------------------------------------------------------------
        If chkEnjoy.value = 1 Then       '共享则必须正确输入所有者
            Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(txtOwnerUsr.Text), Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
            If mcnOwner.State = adStateClosed Then
                MsgBox "所有者密码错误，不能继续。", vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            
            Call SetSQLTrace(gstrServer, Trim(txtOwnerUsr.Text), mcnOwner)
        Else
            '不与其他系统共享，必须建新用户
            If Len(Trim(txtOwnerUsr.Text)) = 0 Then
                MsgBox "请正确指定新用户名。", vbExclamation, gstrSysName
                txtOwnerUsr.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtOwnerPwd.Text)) = 0 Then
                MsgBox "本系统规定，必须指定新用户密码。", vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            If txtOwnerPwd.Text <> txtOwnerLab.Text Then
                MsgBox "密码及其验证的不符合。", vbExclamation, gstrSysName
                txtOwnerPwd.Text = ""
                txtOwnerLab.Text = ""
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            
            Call gobjRegister.CreateUser(gcnOracle, txtOwnerUsr.Text, Trim(txtOwnerPwd.Text), strError)
            If strError <> "" Then
                MsgBox "用户名或密码不符合数据库要求，请重新定义。" & vbCrLf & strError, vbExclamation, gstrSysName
                txtOwnerUsr.SetFocus
                Exit Sub
            End If
            txtOwnerUsr.Tag = "已创建用户"
            
        End If
        
        '------------------------------
        fraSetup(1).Visible = False
        
        If mbln帐套 = False Then
            fraSetup(2).Visible = True
        Else
            '如果是安装帐套，那跳过表空间的设置
            fraSetup(3).Visible = True
        End If
        
    ElseIf fraSetup(2).Visible Then
        '------------------------------------------------------------
        '第三步：
        '------------------------------------------------------------
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If Len(Trim(txtSpaceFile(intCount).Text)) = 0 Then
                MsgBox "请定义" & txtSpaceName(intCount).Caption & "表空间的数据文件。", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Val(txtSpaceSize(intCount).Text) < Val(txtSpaceSize(intCount).Tag) Then
                MsgBox "表空间" & txtSpaceName(intCount).Caption & "必须大于" & txtSpaceSize(intCount).Tag & "M。", vbExclamation, gstrSysName
                txtSpaceSize(intCount).Text = txtSpaceSize(intCount).Tag
                Exit Sub
            End If
            
            If Val(txtSpaceSize(intCount).Text) > 10000 Then
                MsgBox "表空间" & txtSpaceName(intCount).Caption & "超过10G了。", vbExclamation, gstrSysName
                Exit Sub
            End If
        Next
        Call tbsSpace_Click
        
        fraSetup(2).Visible = False
        fraSetup(3).Visible = True
        If optData(0).Visible = False Then
            fraSetup(3).Visible = False
            fraSetup(4).Visible = True
            'cmdNext.Caption = "完成(&F)"
            lblStep(4).Caption = "第四步 产品授权验证"
        End If
    
    ElseIf fraSetup(3).Visible Then
        '------------------------------------------------------------
        '第四步：
        '------------------------------------------------------------
        fraSetup(3).Visible = False
        fraSetup(4).Visible = True
        cmdNext.Caption = "完成(&F)"
        lblStep(4).Caption = "第五步 完成"
    
    ElseIf fraSetup(4).Visible Then
        '------------------------------------------------------------
        '第五步：
        '------------------------------------------------------------
        If chkEnjoy.value = 0 Then
            Set gcnTools = GetConnection("ZLTOOLS")
            If gcnTools Is Nothing Then Exit Sub
        End If
        
        gstrSQL = "    已经完成了所有的安装设置，系统将进入自动安装过程。" & vbCr & vbCr _
                & "    安装过程可能运行较长时间，请不要随意强行中断；否则，" & vbCr _
                & "将可能产生数据垃圾，影响系统运行。" & vbCr & vbCr _
                & "   继续安装吗？"
        If MsgBox(gstrSQL, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        cmdCancel.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        fraSetup(4).Enabled = False
        
        '升级管理工具
        mstrLogFile = GetLogPath(LT_安装, intDefSysCode * 100 + mlng帐套)
        Set objfrmUpSys = New frmAppUpgradeNew '用来清除模块变量
        If Not objfrmUpSys.ToolsInstallUp(Me, stbThis.Panels(2), intDefSysCode * 100 + mlng帐套, lblSetupFile.Caption, mstrLogFile) Then
            cmdNext.Enabled = True
            Unload Me
            Exit Sub
        End If
        
        If SysInstall() Then
            MsgBox "安装成功，可以在完成应用程序安装后正常使用该系统。", vbInformation, gstrSysName
            On Error Resume Next
            Shell "notepad " & mstrLogFile
            err.Clear: On Error GoTo 0
        Else
            MsgBox "安装失败，系统将自动清除已经安装的内容…", vbInformation, gstrSysName
            lblStep(4).Caption = "正在撤卸已经安装的内容…"
            DoEvents
            Call UnInstall
        End If
        cmdNext.Enabled = True
        Unload Me
    End If

End Sub

Private Sub cmdPrevious_Click()
    If fraSetup(4).Visible Then
        cmdNext.Caption = "下一步(&N)"
        fraSetup(4).Visible = False
        fraSetup(3).Visible = True
        If lstData.Count = 1 Then
            fraSetup(3).Visible = False
            fraSetup(2).Visible = True
        End If
    ElseIf fraSetup(3).Visible Then
        fraSetup(3).Visible = False
        
        If mbln帐套 = False Then
            fraSetup(2).Visible = True
        Else
            '帐套安装时跳过表空间的设置
            fraSetup(1).Visible = True
        End If
    ElseIf fraSetup(2).Visible Then
        fraSetup(2).Visible = False
        fraSetup(1).Visible = True
    ElseIf fraSetup(1).Visible Then
        fraSetup(1).Visible = False
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    mbln帐套 = False '缺省认为不是
    Call ApplyOEM(stbThis)
    With imgSetup
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    '根据当前系统的数据文件确定缺省的表空间文件路径
    With rsTemp
        gstrSQL = "select NAME from V$DATAFILE where ROWNUM<2 order by CREATION_TIME"
        .Open gstrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then
            mstrTbsPath = "C:\"
        Else
            If InStr(1, StrReverse(!name), "\") > 0 Then
                mstrTbsPath = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "\") + 1)
            ElseIf InStr(1, StrReverse(!name), "/") > 0 Then
                mstrTbsPath = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "/") + 1)
            Else
                mstrTbsPath = "C:\"
            End If
        End If
    End With
    
    '如果发现当前目录存在安装培植文件，则直接填写
    mstrIniPath = GetSetupPath(App.Path)
    If Dir(mstrIniPath & "\zlSetup.ini") <> "" Then
        lblSetupFile.Caption = mstrIniPath & "\zlSetup.ini"
        If ChkSetupFile() = False Then
            mbln帐套 = False
            mstrIniPath = ""
            lblSetupFile.Caption = ""
        End If
    End If
    
End Sub

Private Function GetSetupPath(ByVal strAppPath As String) As String
'得到缺省的安装路径
    Dim strPath() As String
    Dim strTemp As String
    
    ReDim strPath(0 To 0) As String
    
    strTemp = Dir(strAppPath & "\", vbDirectory)
    Do While strTemp <> ""
        strTemp = UCase(strTemp)
        If InStr(strTemp, ".") = 0 Then
            If strTemp <> "APPLY" And strTemp <> "TOOLS" And strTemp <> "附加文件" Then
                ReDim Preserve strPath(0 To UBound(strPath) + 1)
                strPath(UBound(strPath)) = strTemp
            End If
        End If
        strTemp = Dir(, vbDirectory)
    Loop
    If UBound(strPath) = 1 Then
        GetSetupPath = strAppPath & "\" & strPath(1) & "\应用脚本"
    Else
        GetSetupPath = strAppPath
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdNext.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    Set mclsRunScript = Nothing

    Set objText = Nothing
    Set mcnOwner = Nothing
End Sub

Private Sub cmbEnjoy_Click()
    Dim rsTemp As New ADODB.Recordset
    
    With rsTemp
        gstrSQL = "select 所有者 from zlSystems where 编号=" & cmbEnjoy.ItemData(cmbEnjoy.ListIndex)
        .Open gstrSQL, gcnOracle, adOpenKeyset
        cmbEnjoy.Tag = !所有者
        If txtOwnerUsr.Enabled = False Then
            txtOwnerUsr.Text = cmbEnjoy.Tag
        End If
    End With
End Sub


Private Sub lstData_Click(Index As Integer)
    SetPromptText Split(lstData(Index).Tag, "=")(lstData(Index).ListIndex + 1)
End Sub

Private Sub optData_Click(Index As Integer)
    SetPromptText optData(Index).ToolTipText
    For intCount = 0 To optData.UBound
        If intCount = Index And lstData(Index).ListCount > 0 Then
            lstData(intCount).Visible = True
            lblNoData.Visible = False
        Else
            lstData(intCount).Visible = False
            lblNoData.Visible = True
        End If
    Next
    
    cmdSelectAll.Visible = lstData(Index).ListCount > 0
    cmdClearAll.Visible = lstData(Index).ListCount > 0
    
    With lblNoData
        .Left = lstData(0).Left
        .Width = lstData(0).Width
        .Top = lstData(0).Top
        .Height = lstData(0).Height
        .Caption = vbCrLf & "     " & optData(Index).Caption & "数据组不包含细分的可选数据项。"
    End With
End Sub


Private Sub tbsSpace_Click()
    For intCount = 0 To tbsSpace.Tabs.Count - 1
        txtSpaceName(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceFile(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceSize(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        chkSpaceExtd(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        cboSpaceExtentType(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceExtentSize(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceExtentSize(intCount).Enabled = (cboSpaceExtentType(intCount).ListIndex = 1)
        '索引表空间的日志属性
        If tbsSpace.Tabs(intCount + 1).Selected Then
            chkLogin(intCount).Visible = UCase(txtSpaceName(intCount).Caption) Like "ZL9INDEX*"
        Else
            chkLogin(intCount).Visible = False
        End If
       
        If tbsSpace.Tabs(intCount + 1).Selected Then
            txtSpaceFile(intCount).SetFocus
        End If
    Next
End Sub

Private Sub txtOwnerUsr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpaceExtentSize_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub SetProgressVisible(ByVal blnVisible As Boolean)
    If blnVisible = True Then
        If stbThis.Panels.Count = 3 Then
            '增加一个窗格
            stbThis.Panels.Add 3
            stbThis.Panels(3).AutoSize = sbrSpring
            stbThis.Panels(2).AutoSize = sbrNoAutoSize
            stbThis.Panels(2).MinWidth = 1440
        End If
        pgbState.Left = stbThis.Panels(3).Left + 30
        pgbState.Width = stbThis.Panels(4).Left - pgbState.Left - 150
        pgbState.Top = stbThis.Top + stbThis.Height / 3
        pgbState.Visible = True
    Else
        If stbThis.Panels.Count = 4 Then
            stbThis.Panels(2).AutoSize = sbrSpring
            stbThis.Panels.Remove 3
        End If
        pgbState.Visible = False
    End If
    
End Sub

Private Function ChkSetupFile(Optional blnMsg As Boolean) As Boolean
    Dim strTemp As String
    '-------------------------------------
    '检查解释安装配置文件的正确性
    '-------------------------------------
    mstrIniPath = Mid(lblSetupFile.Caption, 1, Len(lblSetupFile.Caption) - 11)
    '相关文件匹配性检查
    strTemp = ""
    If Dir(mstrIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "序列文件" & mstrIniPath & "zlSequence.sql"
    If Dir(mstrIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "数据表文件" & mstrIniPath & "zlTable.sql"
    If Dir(mstrIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "约束文件" & mstrIniPath & "zlConstraint.sql"
    If Dir(mstrIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "索引文件" & mstrIniPath & "zlIndex.sql"
    If Dir(mstrIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "视图文件" & mstrIniPath & "zlView.sql"
    If Dir(mstrIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "函数过程文件" & mstrIniPath & "zlProgram.sql"
    
    '不检查,因为9系统没有此文件
    'If Dir(mstrIniPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "包文件" & mstrIniPath & "zlPackage.sql"
    
    If Dir(mstrIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "管理数据文件" & mstrIniPath & "zlManData.sql"
    If Dir(mstrIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "应用数据文件" & mstrIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '安装配置文件解释
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(lblSetupFile.Caption)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        intDefSysCode = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统名]" Then
        strDefSysName = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        strDefVersion = Trim(Mid(strTemp, 6))
    
        '判断是否应该把本次安装作为帐套安装
        Dim lngTemp As Long
        Dim lngMax As Long        '最大的帐套号
        Dim blnHase  As Boolean   '是否有同系统存在
        Dim lngMin As Long
        Dim lstTemp As ListItem
        
        
        mbln帐套 = False
        mlng帐套 = 0
        lngMin = 99
        For Each lstTemp In frmAppStart.lvwSys.ListItems
            lngTemp = Mid(lstTemp.Key, 2)
            If lngTemp \ 100 = intDefSysCode Then
                '系统相同
                blnHase = True
                If lngMax < lngTemp Mod 100 Then
                    lngMax = lngTemp Mod 100 '保存最大的帐套号
                End If
                
                If strDefVersion = lstTemp.SubItems(1) Then
                    '版本也相同，那就可以了
                    mbln帐套 = True
                    If lngMin > lngTemp Mod 100 Then
                        lngMin = lngTemp Mod 100 '保存最小的帐套号
                        Set mlst标准 = lstTemp '取最小帐套号作为标准帐套
                    ElseIf lngMin = 99 Then
                        Set mlst标准 = lstTemp '初始先任意获取一个作为标准帐套
                    End If
                End If
            End If
        Next
        If blnHase = True Then
            '有同系统的安装
            If mbln帐套 = False Then
                If blnMsg Then MsgBox "当前数据库中也有相同类型的系统存在，但由于版本不符，不能新增。", vbInformation, gstrSysName
                Exit Function
            Else
                If blnMsg = False Then
                    Exit Function
                Else
                    If lngMax >= 99 Then
                        MsgBox "当前数据库中也有相同类型的系统存在，且数量足够多，不能新增。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If MsgBox("当前数据库中已有" & strDefSysName & "系统存在，你是否要再新增一个？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    mlng帐套 = lngMax + 1
                End If
            End If
        End If
    Else
        err.Raise 10
    End If
    Caption = "应用系统安装" & " - " & strDefSysName & " V" & strDefVersion
    lblSysName.Caption = "系统名：" & strDefSysName
    lblVersion.Caption = "版本号：" & strDefVersion
        
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[表空间]" Then
        strDefSpace = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[用户名]" Then
        strDefUser = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[数据组]" Then
        strDefData = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    mstrExtSysCode = ""
    mstrExtVersion = ""
    If Not objText.AtEndOfStream Then
        '还有扩展系统的设置
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 5) = "[主系统]" Then
            mstrExtSysCode = Trim(Mid(strTemp, 6))
            
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 5) = "[主版本]" Then
                mstrExtVersion = Trim(Mid(strTemp, 6))
            Else
                mstrExtSysCode = ""
            End If
        End If
    End If
    Call FillShare  '得到共享清单
    If mstrExtSysCode <> "" And cmbEnjoy.ListCount = 0 Then
        If blnMsg Then MsgBox "该扩展系统没找到可依赖的主系统。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If err <> 0 Then
        If blnMsg Then MsgBox "安装配置文件丢失或不正确。", vbExclamation, gstrSysName
        Exit Function
    End If
    objText.Close
    
    '表空间正确性检查
    intItems = tbsSpace.Tabs.Count
    For intCount = 0 To intItems - 2
        tbsSpace.Tabs.Remove 1
    Next
    
    err = 0
    On Error Resume Next
    aryRow = Split(strDefSpace, "||")
    For intCount = 0 To UBound(aryRow)
        aryVal = Split(aryRow(intCount), "|")
        If intCount = 0 Then
            tbsSpace.Tabs(1).Caption = aryVal(0)
            tbsSpace.Tabs(1).Key = aryVal(1)
        Else
            tbsSpace.Tabs.Add , aryVal(1), aryVal(0)
        End If
        If intCount > txtSpaceName.Count - 1 Then Load txtSpaceName(intCount)
        If intCount > txtSpaceFile.Count - 1 Then Load txtSpaceFile(intCount)
        If intCount > txtSpaceSize.Count - 1 Then Load txtSpaceSize(intCount)
        If intCount > chkSpaceExtd.Count - 1 Then Load chkSpaceExtd(intCount)
        If intCount > cboSpaceExtentType.Count - 1 Then Load cboSpaceExtentType(intCount)
        If intCount > txtSpaceExtentSize.Count - 1 Then Load txtSpaceExtentSize(intCount)
        '索引表空间的日志属性
        If intCount > chkLogin.Count - 1 Then Load chkLogin(intCount)
        
        txtSpaceName(intCount).Caption = aryVal(1)
        If UCase(aryVal(1)) Like "ZL9INDEX*" Then
            chkLogin(intCount).value = 0
        Else
            chkLogin(intCount).value = 1
        End If
        chkLogin(intCount).Visible = False
        
        txtSpaceFile(intCount).Tag = aryVal(1)
        txtSpaceFile(intCount).Text = mstrTbsPath & txtSpaceFile(intCount).Tag & ".DBF"
        txtSpaceSize(intCount).Text = aryVal(2)
        txtSpaceSize(intCount).Tag = aryVal(3)
        
        If aryVal(4) = "T" Then
            chkSpaceExtd(intCount).value = 1
        Else
            chkSpaceExtd(intCount).value = 0
        End If
        
        
        '表空间区分配类型
        cboSpaceExtentType(intCount).Clear
        cboSpaceExtentType(intCount).AddItem "自动分配区尺寸"
        cboSpaceExtentType(intCount).AddItem "统一分配区尺寸"
        cboSpaceExtentType(intCount).ListIndex = 0
        txtSpaceExtentSize(intCount).Text = 1
        txtSpaceExtentSize(intCount).Enabled = (cboSpaceExtentType(intCount).ListIndex = 1)
    Next
    If err <> 0 Then
        If blnMsg Then MsgBox "安装配置文件表空间设置错误，不能继续安装。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If mstrExtSysCode = "" Then
        '非扩展系统，处理不变
        If mbln帐套 = False Then
            '没有多帐套，可能共享
            If cmbEnjoy.ListCount = 0 Then
                chkEnjoy.value = 0
                chkEnjoy.Enabled = False
            Else
                chkEnjoy.Enabled = True
            End If
            If chkEnjoy.value <> 1 Then
                txtOwnerUsr.Text = strDefUser
            End If
            
        Else
            chkEnjoy.Enabled = False '不能再选择共享，只能新增
            chkEnjoy.value = 0
            txtOwnerUsr.Text = strDefUser & mlng帐套
            '也不用于检查表空间的设置，因为用以前的
        End If
    Else
        '根据合并与否判断共享情况
        chkEnjoy.Enabled = False '如果该系统不是扩展系统，那么允许不选择共享性
        chkEnjoy.value = 1
    End If
    
    '数据分组可选文件匹配性检查
    Dim intOptions As Integer       '选项框的数目
    Dim lngHeight As Long           '控件的行高度
    
    For intCount = 0 To optData.UBound
        optData(intCount).Visible = False
    Next
    For intCount = 0 To lstData.UBound
        lstData(intCount).Visible = False
    Next
    
    intOptions = 0
    err = 0
    aryRow = Split(strDefData, "||")
    For intCount = 0 To UBound(aryRow)
        If Dir(mstrIniPath & "zlSelData" & intCount & ".sql") <> "" Then
            If intOptions > optData.Count - 1 Then Load optData(intOptions)
            optData(intOptions).Tag = intCount
            optData(intOptions).Left = optData(0).Left
            optData(intOptions).ToolTipText = ""
            optData(intOptions).Visible = True
            
            If intOptions > lstData.Count - 1 Then Load lstData(intOptions)
            lstData(intOptions).Left = fraSelData.Width / 2 - 300
            lstData(intOptions).Width = fraSelData.Width / 2 - optData(0).Left + 300
            lstData(intOptions).Top = lstData(0).Top
            lstData(intOptions).Tag = ""
            lstData(intOptions).Clear
            intItems = InStr(1, aryRow(intCount), ">")
            If intItems = 0 Then
                If InStr(1, aryRow(intCount), "=") = 0 Then
                    optData(intOptions).Caption = Trim(aryRow(intCount))
                Else
                    optData(intOptions).Caption = Trim(Left(aryRow(intCount), InStr(1, aryRow(intCount), "=") - 1))
                    optData(intOptions).ToolTipText = Trim(Mid(aryRow(intCount), InStr(1, aryRow(intCount), "=") + 1))
                End If
            Else
                optData(intOptions).Caption = Trim(Mid(aryRow(intCount), 1, intItems - 1))
                If InStr(1, optData(intOptions).Caption, "=") > 0 Then
                    optData(intOptions).ToolTipText = Trim(Mid(optData(intOptions).Caption, InStr(1, optData(intOptions).Caption, "=") + 1))
                    optData(intOptions).Caption = Trim(Left(optData(intOptions).Caption, InStr(1, optData(intOptions).Caption, "=") - 1))
                End If
                strTemp = Mid(aryRow(intCount), intItems + 1)
                aryVal = Split(strTemp, "|")
                For intItems = 0 To UBound(aryVal)
                    If Dir(mstrIniPath & "zlSelData" & intCount & intItems & ".sql") <> "" Then
                        If InStr(1, aryVal(intItems), "=") = 0 Then
                            lstData(intOptions).AddItem Trim(aryVal(intItems))
                            lstData(intOptions).Tag = lstData(intOptions).Tag & "="
                        Else
                            lstData(intOptions).AddItem Trim(Left(aryVal(intItems), InStr(1, aryVal(intItems), "=") - 1))
                            lstData(intOptions).Tag = lstData(intOptions).Tag & Mid(aryVal(intItems), InStr(1, aryVal(intItems), "="))
                        End If
                        lstData(intOptions).ItemData(lstData(intOptions).NewIndex) = Val(intItems)
                    End If
                Next
            End If
            intOptions = intOptions + 1
        End If
    Next
    cmdSelectAll.Left = lstData(0).Left
    cmdClearAll.Left = lstData(0).Left + lstData(0).Width - cmdClearAll.Width
    
    
    If err <> 0 Then
        If blnMsg Then MsgBox "安装配置文件数据分组设置错误，不能继续安装。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If intOptions = 1 Then
        optData(0).Top = lstData(0).Top
        optData(0).Height = lstData(0).Height
        lstData(0).Left = optData(0).Left
        lstData(0).Width = fraSelData.Width - optData(0).Left * 2
        lstData(0).ZOrder
    ElseIf intOptions <> 0 Then
        lngHeight = lstData(0).Height / intOptions
        For intCount = 0 To intOptions - 1
            optData(intCount).Top = lstData(0).Top + intCount * lngHeight
            optData(intCount).Height = lngHeight
        Next
    Else
        lblNoData.Visible = True
    End If
    If intOptions <> 0 Then
        optData(0).value = True
        Call optData_Click(0)
    End If

    '顺便把注册文件也一并检查了
    Call ChkRegFile
    SetPromptText ""
    
    ChkSetupFile = True
End Function

Private Sub ChkRegFile()
    '判断系统授权
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo errHand
    gstrSQL = "Select Count(*) From zltools.Zlregfunc f, zltools.Zlreginfo r, zltools.zlRegAudit t Where r.项目 = '授权证章' And f.系统 = " & intDefSysCode
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.Fields(0).value > 0 Then
        Me.lblRegAudit.Caption = "    已经具备该系统应用授权，可以在装载后正常授权使用。"
        Exit Sub
    End If
errHand:
    Me.lblRegAudit.Caption = "    由于还不具备该系统应用授权，虽然可以继续装载，但无法正常授权使用！"
End Sub

Private Sub FillShare()
'读出可用的共享清单
    Dim rsTemp As New ADODB.Recordset
    Dim varVersion As Variant, varExtVersin As Variant
    Dim i As Long, bln满足 As Boolean
    
    cmbEnjoy.Clear
    If mstrExtSysCode = "" Then
        '本系统不是扩展系统，可共享任意系统
        gstrSQL = "select 编号,名称 from zlsystems order by 编号"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            cmbEnjoy.AddItem rsTemp("名称") & "(" & rsTemp("编号") & ")"
            cmbEnjoy.ItemData(cmbEnjoy.NewIndex) = rsTemp("编号")
            rsTemp.MoveNext
        Loop
    Else
        '是扩展系统，那必须要完成三方面的判断
        '1)系统号相符
        '2)没被其它的相同系统扩展
        '3)版本不能低于要求
        gstrSQL = "select A.编号,A.名称,A.版本号 from zlsystems A " & _
                  "  Where floor(A.编号 / 100) = " & mstrExtSysCode & _
                  "        and not exists (select B.编号 from zlsystems B where B.共享号=A.编号 and floor(B.编号/100)=" & intDefSysCode & ")"
        
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        
        varExtVersin = Split(mstrExtVersion, ".")
        Do Until rsTemp.EOF
            '判断版本
            bln满足 = True
            varVersion = Split(rsTemp("版本号"), ".")
            
            For i = LBound(varExtVersin) To UBound(varExtVersin)
                If Val(varExtVersin(i)) > Val(varVersion(i)) Then
                    '脚本中的版本号高于实际数据库的，不能满足
                    bln满足 = False
                    Exit For
                ElseIf Val(varExtVersin(i)) < Val(varVersion(i)) Then
                    '已经满足，不要再比较下一位
                    Exit For
                End If
            Next
            
            If bln满足 = True Then
                '符合条件
                cmbEnjoy.AddItem rsTemp("名称") & "(" & rsTemp("编号") & ")"
                cmbEnjoy.ItemData(cmbEnjoy.NewIndex) = rsTemp("编号")
            End If
            rsTemp.MoveNext
        Loop
        
    End If
End Sub
Private Function SysInstall() As Boolean
    '----------------------------------
    '功能：完成系统的安装处理
    '---------安装算法-----------------
    '    创建本系统数据表空间
    '    If not 共享已经安装的系统 Then
    '        创建本系统所有者
    '        由工具所有者授予必要的工具数据对象权限
    '    End If
    '    创建本系统数据对象
    '    必须数据及可选数据安装
    '----------------------------------
    Dim rsTemp As New ADODB.Recordset, cnCtxsys As New ADODB.Connection
    Dim strTmpSpace As String
    Dim strTemp As String, strError As String
    Dim intVer As Integer
    Dim blnIgnoreErr As Boolean     '忽略错误
    Dim strPassword As String, strUserName As String, lngAbort As Long, cllRoles As Collection
    
    strUserName = UCase(Trim(txtOwnerUsr.Text))
    On Error GoTo errHand
    intVer = GetOracleVersion
    gcnOracle.Execute "Grant Select on sys.v_$session to Public"
    gcnOracle.Execute "Grant Select on sys.v_$parameter to Public"
        
    With rsTemp
        gstrSQL = "SELECT TEMPORARY_TABLESPACE FROM DBA_USERS WHERE USERNAME='ZLTOOLS'"
        .Open gstrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then SysInstall = False: Exit Function
        strTmpSpace = .Fields(0).value
    End With
    
    If mbln帐套 = False Then
        '创建本系统数据表空间
        SetPromptText "创建表空间"
        pgbState.value = 0
        
        '
        SetProgressVisible True
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If CreateTbs(txtSpaceName(intCount).Caption, _
                            txtSpaceFile(intCount).Text, _
                            txtSpaceSize(intCount).Text, _
                            chkSpaceExtd(intCount).value, _
                            False, _
                            cboSpaceExtentType(intCount).ListIndex = 0, _
                            Val(txtSpaceExtentSize(intCount).Text), _
                            chkLogin(intCount).value = 0 _
                            ) <> 1 Then
                GoTo errHand
            End If
            pgbState.value = (intCount + 1) / tbsSpace.Tabs.Count * 100
            DoEvents
        Next
        pgbState.value = 0
        SetProgressVisible False
    End If
    
    '如果不共享已经安装的系统
    If chkEnjoy.value = 0 Then
        '创建本系统所有者
        SetPromptText "授权给所有者" & strUserName
        pgbState.value = 0
        SetProgressVisible True
                
        '在第二步时已创建
        
        gstrSQL = "alter user " & strUserName & _
                " DEFAULT TABLESPACE " & txtSpaceName(0).Caption & _
                " TEMPORARY TABLESPACE " & strTmpSpace
        gcnOracle.Execute gstrSQL
        
        '12c的resource角色缺省没有UNLIMITED TABLESPACE权限
        '增加CREATE TRIGGER权限，用于历史数据转出的存储过程中为禁用外键的表临时创建触发器（转完后会删除）
        '存储过程中execute immediate执行动态SQL时，需显示授权，即使所属角色有权限（例如：即使是DBA，仍然需要授权）
        gstrSQL = "Grant Connect,Resource," & IIf(chkDBA.value = 1, "DBA,", "") & _
                " UNLIMITED TABLESPACE,Create Table,Create Sequence,Create Role,Create User,Drop User,Alter User,Create Public Synonym,Drop Public Synonym," & _
                " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster," & _
                " Create Materialized View, Alter Any Materialized View, Drop Any Materialized View,CREATE TRIGGER" & _
                " to " & strUserName & " With Admin Option"
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Select on sys.dba_role_privs to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Select on sys.dba_roles to " & strUserName
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Execute on sys.dbms_sql to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
        
        gstrSQL = "Grant Select on sys.gv_$session to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
     
        On Error Resume Next '创建全文检索的参数，有可能没有该用户，所以把错误屏蔽
        gstrSQL = "Grant CTXAPP to " & strUserName & " With Admin Option"
        gcnOracle.Execute gstrSQL
        gcnOracle.Execute "alter user ctxsys identified by ctxsys"
        gcnOracle.Execute "alter user  ctxsys account Unlock"
        cnCtxsys.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ctxsys", "ctxsys"
        cnCtxsys.Execute "Grant Execute on ctx_ddl to " & strUserName & " With Grant Option" '为了在过程中执行包函数
        
        err = 0: On Error GoTo errHand

        '由工具所有者授予必要的工具数据对象权限
        SetPromptText "管理工具对象权限授予" & strUserName
        SetProgressVisible False
        Call ReGrantForTools(gcnTools, strUserName)
    End If
    
    '填写安装系统清单
    gstrSQL = "insert into zlSystems(编号,共享号,名称,所有者,安装日期,正常安装,版本号)" & _
            " values(" & intDefSysCode * 100 + mlng帐套 '编号的前三位是系统号，后两位是帐套号
    If chkEnjoy.value = 1 Then
        gstrSQL = gstrSQL & "," & cmbEnjoy.ItemData(cmbEnjoy.ListIndex)
    Else
        gstrSQL = gstrSQL & ",null"
    End If
    gstrSQL = gstrSQL & ",'" & strDefSysName & "'"
    gstrSQL = gstrSQL & ",'" & strUserName & "'"
    gstrSQL = gstrSQL & ",sysdate,0,'" & strDefVersion & "')"
    gcnOracle.Execute gstrSQL
    
    
    '创建本系统数据对象
    Set mcnOwner = gobjRegister.GetConnection(gstrServer, strUserName, Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
    strPassword = gobjRegister.GetPassword
    Call SetSQLTrace(gstrServer, strUserName, mcnOwner)
    
    Set cllRoles = New Collection
    
    '是否忽略错语
    blnIgnoreErr = chkEnjoy.value <> 0
    If gblnInIDE Then blnIgnoreErr = False
    Set mclsRunScript = New clsRunScript
    Call mclsRunScript.InitGlobalPara(Me, intDefSysCode * 100 + mlng帐套, blnIgnoreErr, mstrLogFile)
    Call mclsRunScript.InitUserList(strUserName, strPassword)
    Set mclsRunScript.Connection = mcnOwner: mclsRunScript.ConnectType = 0
    mclsRunScript.IsRoleCollect = True '收集角色
    
    SetProgressVisible True
    SetPromptText "创建序列"
    If RunSQLScript(mstrIniPath & "zlSequence.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "创建数据表"
    If RunSQLScript(mstrIniPath & "zlTable.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "创建约束"
    If RunSQLScript(mstrIniPath & "zlConstraint.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    SetPromptText "创建索引"
    If RunSQLScript(mstrIniPath & "zlIndex.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    SetPromptText "创建视图"
    If RunSQLScript(mstrIniPath & "zlView.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "函数与过程"
    If RunSQLScript(mstrIniPath & "zlProgram.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    If Dir(mstrIniPath & "zlPackage.sql") <> "" Then
        SetPromptText "创建包"
        If RunSQLScript(mstrIniPath & "zlPackage.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    End If
    Set cllRoles = mclsRunScript.Roles
    
    If cllRoles.Count <> 0 Then
        '需要对相关的角色进行授权
        SetPromptText "角色授权处理"
        If GrantToRole(mcnOwner, cllRoles, strUserName) = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    End If
    
    
    If chkEnjoy.value <> 0 Then
        '共享安装时，需要重新编译（因为可能有同名的存储过程被重新创建了）
        SetPromptText "编译对象"
        Call ReCompileProcedure(mcnOwner)
    End If
    
   
    '必须数据
    SetPromptText "管理数据安装"
    If mbln帐套 = False Then
        If RunSQLScript(mstrIniPath & "zlManData.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    Else
        '通过数据库中拷贝得到
        If CopyManageData(mcnOwner) = False Then GoTo errHand
    End If
    SetPromptText "应用数据安装"
    If RunSQLScript(mstrIniPath & "zlAppData.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    '安装报表
    SetPromptText "固定报表安装"
    If mbln帐套 = False Then
        If RunSQLScript(mstrIniPath & "zlReport.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    Else
        '通过数据库中拷贝得到
        If CopyReport(mcnOwner, Mid(mlst标准.Key, 2), intDefSysCode * 100 + mlng帐套) = False Then GoTo errHand
    End If
    
    '可选数据安装
    If chkSelData.value = 1 Then
        For intCount = 0 To optData.UBound
            If optData(intCount).value = True Then
                SetPromptText optData(intCount).Caption
                If RunSQLScript(mstrIniPath & "zlSelData" & optData(intCount).Tag & ".sql") = False Then
                    SetProgressVisible False: GoTo errHand:
                End If
                
                For intItems = 0 To lstData(intCount).ListCount - 1
                    If lstData(intCount).Selected(intItems) = True Then
                        SetPromptText lstData(intCount).List(intItems)
                        If RunSQLScript(mstrIniPath & "zlSelData" & optData(intCount).Tag & lstData(intCount).ItemData(intItems) & ".sql") = False Then
                            SetProgressVisible False: GoTo errHand:
                        End If
                    End If
                Next
            End If
        Next
    End If
    
    '调整安装导致的序列与实际数值的匹配
    SetPromptText "序列检查"
    DoEvents
    Call ChkSequence
    
    '填写安装记录为正常安装
    gstrSQL = "update zlSystems set 正常安装=1 where 编号=" & intDefSysCode * 100 + mlng帐套
    gcnOracle.Execute gstrSQL
    gstrSQL = "insert into zlSysFiles(系统,操作,文件名,日期,操作人)" & _
            " values (" & intDefSysCode * 100 + mlng帐套 & ",1,'" & lblSetupFile.Caption & "',sysdate,user)"
    gcnOracle.Execute gstrSQL
    
    If CheckHavHistory(intDefSysCode * 100 + mlng帐套) Then
        '刘兴洪：加入创建历史数据空间
        If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, strUserName, _
            strPassword, intDefSysCode * 100 + mlng帐套, 0, 0) = False Then
            If mcnOwner.State = adStateOpen Then mcnOwner.Close
            Set mcnOwner = Nothing
            Exit Function
        End If
    End If
    
    '创建当前所有者的全部对象的公共同义词('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
    mcnOwner.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
    
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    Set mcnOwner = Nothing
    Set mclsRunScript = Nothing
    SysInstall = True
    Exit Function

errHand:
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    
    Set mcnOwner = Nothing
    SetProgressVisible False
    SysInstall = False
End Function

Private Sub SetPromptText(ByVal strText As String)
    stbThis.Panels(2).Text = strText
    stbThis.Panels(2).ToolTipText = strText
End Sub

Private Function UnInstall() As Boolean
    '----------------------------------
    '功能：删除已经的安装处理
    '----------------------------------
    Dim rsTemp As New ADODB.Recordset, rsSys As New ADODB.Recordset, blnCanRemoveMSGData As Boolean
    Dim strSpaces As String, strFiles As String, aryFile() As String, strErrInfo As String
    Dim lngRowH As Long
    
    
    If mbln帐套 = False Then
        '搜索数据文件
        strSpaces = ""
        If intDefSysCode = 1 Or intDefSysCode = 25 Then
            'ZLMSGDATA表空间标准版也存在，LIS也存在，当只存在一个系统时可以直接卸载
            gstrSQL = "Select Count(1) 计数" & vbNewLine & _
                        "From Zlsystems" & vbNewLine & _
                        "Where Floor(编号 / 100) In (1, 25)"
            rsSys.Open gstrSQL, gcnOracle
            blnCanRemoveMSGData = rsSys!计数 = 1
        Else
            blnCanRemoveMSGData = True '没有ZLMSGDATA表空间，为了简化逻辑
        End If
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If blnCanRemoveMSGData Then
                strSpaces = strSpaces & ",'" & UCase(Trim(txtSpaceName(intCount).Caption)) & "'"
            ElseIf UCase(Trim(txtSpaceName(intCount).Caption)) <> "ZLMSGDATA" Then
                strSpaces = strSpaces & ",'" & UCase(Trim(txtSpaceName(intCount).Caption)) & "'"
            End If
            DoEvents
        Next
        strFiles = ""
        With rsTemp
            gstrSQL = "select F.NAME from V$TABLESPACE T,V$DATAFILE F where T.TS#=F.TS#  and T.NAME in(" & Mid(strSpaces, 2) & ")"
            .Open gstrSQL, gcnOracle
            Do While Not .EOF
                strFiles = strFiles & ";" & .Fields(0).value
                DoEvents
                .MoveNext
            Loop
        End With
    End If
    strErrInfo = ""
    err = 0
    On Error Resume Next
    
    SetPromptText "正在清除已安装的数据…"
    '删除安装记录
    gstrSQL = "delete from zlSystems where 编号=" & intDefSysCode * 100 + mlng帐套
    gcnOracle.Execute gstrSQL
    
    '清理无效菜单
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            .Open gstrSQL, gcnOracle
            If .EOF Then Exit Do
            gstrSQL = "delete from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            gcnOracle.Execute gstrSQL
        Loop
    End With
    
    '删除本系统所有者
    If chkEnjoy.value = 0 Then
        SetPromptText "正在删除已创建的用户…"
        intCount = 0
        Do
            gcnOracle.Execute "drop user " & txtOwnerUsr.Text & " cascade"
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open "select * from all_users where username='" & UCase(txtOwnerUsr.Text) & "'", gcnOracle
                If .EOF Then Exit Do
            End With
            intCount = intCount + 1
            DoEvents
            If intCount > 10000 Then
                strErrInfo = strErrInfo & vbCr & "用户:" & txtOwnerUsr
                Exit Do
            End If
        Loop
    End If
    
    If mbln帐套 = False Then
        '删除本系统数据表空间
        SetPromptText "正在删除已创建的表空间和数据文件…"
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If CheckSpaceIsUse("表空间", txtSpaceName(intCount).Caption, txtOwnerUsr.Text) = False Then
                '没有其他用户使用，可以删除
                gcnOracle.Execute "alter tablespace " & txtSpaceName(intCount).Caption & " offline"
                
                DoEvents
                gcnOracle.Execute "drop tablespace " & txtSpaceName(intCount).Caption & " including contents and datafiles cascade constraints"
            End If
        Next
        
        '取消直接删除文件的语句，因为本机的文件不一定是数据库服务器文件
    End If
    
    SetPromptText ""
    If strErrInfo <> "" Then
        MsgBox "请重启动Oracle后,手工删除以下内容：" & strErrInfo, vbExclamation, gstrSysName
    Else
        MsgBox "请检查硬盘空间和数据库系统，确认无误后重新安装。", vbExclamation, gstrSysName
    End If
End Function


Private Function CreateTbs(TbsName As String, TbsFile As String, TbsSize As Integer, Optional AutoExtend As Boolean, _
     Optional Temp As Boolean, Optional AutoAllocate As Boolean, Optional ExtentSize As Integer, Optional Nologging As Boolean) As Byte
    '----------------------------------------------
    '功能：系统用户,根据参数创建表空间,固定为本地管理类型(8i以前不支持,那时只能创建字典管理类型)
    '       因可能涉及LOB字段等原因,不创建ASSM表空间(仅9i以上支持,SEGMENT SPACE MANAGEMENT AUTO)
    '参数：
    '   TbsName:表空间名称
    '   TbsFile:表空间文件
    '   TbsSize:表空间大小(M为单位)
    '   Extend:是否自动管理区,否则统一范围尺寸
    '   ExtentSize:统一区尺寸,临时表空间必须指定尺寸(Oracle缺省为1M)
    '   Temp:是否为临时表空间
    '返回：1-创建成功；2-表空间已经存在；3-创建失败
    '----------------------------------------------
    DoEvents
    If Temp Then
        gstrSQL = "CREATE TEMPORARY TABLESPACE " & TbsName & " TEMPFILE '" & TbsFile & "'"
    Else
        gstrSQL = "CREATE TABLESPACE " & TbsName & " DATAFILE '" & TbsFile & "'"
    End If
    gstrSQL = gstrSQL & _
            " SIZE " & TbsSize & "M REUSE " & _
             IIf(AutoExtend, "AUTOEXTEND ON NEXT " & IIf(TbsSize \ 10 = 0, 1, TbsSize \ 10) & "M", "") & _
            " EXTENT MANAGEMENT LOCAL " & _
                IIf(AutoAllocate And Not Temp, " AUTOALLOCATE", " UNIFORM SIZE " & IIf(ExtentSize = 0, "1", ExtentSize) & "M") & _
                IIf(Nologging And Not Temp, " Nologging", "")
            
    err = 0
    On Error Resume Next
    gcnOracle.Execute gstrSQL
    DoEvents
    If err = 0 Then
        CreateTbs = 1
    ElseIf gcnOracle.Errors.Count > 0 Then
        'ORA-01543: 表空间'XXX'已经存在
        If UCase(gcnOracle.Errors(0).Description) Like "ORA-01543: *'ZLMSGDATA'*" Then
            err.Clear
            CreateTbs = 1
        Else
            If MsgBox("出现下述错误，是否跳过继续？" & vbCrLf & vbTab & gcnOracle.Errors(0).Description, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                CreateTbs = 2
            Else
                CreateTbs = 1
            End If
        End If
    Else
        MsgBox "表空间" & TbsName & "无法创建，请检查磁盘大小等。", vbExclamation, gstrSysName
        CreateTbs = 2
    End If

End Function

Private Function GrantToRole(ByVal cnThis As ADODB.Connection, ByVal cllRoles As Collection, ByVal strOwnerName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将重新对相关的角色进行授权
    '入参:cllRoles-角色集
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-09 16:42:54
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, strRoleName As String
    Dim lngCount As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strOwner() As String
    ReDim strOwner(0)
    strOwner(0) = strOwnerName
    '该文件处理完之后，再处理角色的授权
    lngCount = cllRoles.Count
    pgbState.value = 0
    If lngCount = 0 Then Exit Function
    SetProgressVisible True
    '由于系统号不同，数据可能没增加
    Dim lngNewSystem As Long, lngOldSystem  As Long
    lngNewSystem = intDefSysCode * 100 + mlng帐套
    lngOldSystem = Mid(mlst标准.Key, 2)
            
    For i = 0 To cllRoles.Count
        strRoleName = cllRoles(i)
        If mbln帐套 = True Then
            gstrSQL = "insert into zlRoleGrant(系统,角色,序号,功能) " & _
                   " select " & lngNewSystem & ",角色,序号,功能 from zlRoleGrant where 角色='" & strRoleName & "' and 系统=" & lngOldSystem
            cnThis.Execute gstrSQL
        End If
        gstrSQL = "select B.对象,B.权限 from zlrolegrant A,zlprogprivs B " & _
                    " where A.角色='" & strRoleName & "' and B.所有者='" & UCase(Trim(txtOwnerUsr.Text)) & "' and A.系统=B.系统 and A.序号=B.序号 and A.功能=B.功能 "
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, cnThis, adOpenStatic, adLockReadOnly
        
        Do Until rsTemp.EOF
            gstrSQL = "GRANT " & rsTemp("权限") & " ON " & rsTemp("对象") & " TO " & strRoleName
            cnThis.Execute gstrSQL
            rsTemp.MoveNext
        Loop
        
        Call GrantSpecialToRole(cnThis, strRoleName, False, strOwner, True)
        pgbState.value = Int(pgbState.value / lngCount * 100)
        
    Next
    SetProgressVisible False
    GrantToRole = True
End Function

Private Function CopyManageData(ByVal cnExecuter As ADODB.Connection) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngNewSystem As Long
    Dim lngOldSystem As Long
    Dim strOldOwner As String
    
    pgbState.value = 0
    SetProgressVisible True
    
    lngNewSystem = intDefSysCode * 100 + mlng帐套
    lngOldSystem = Mid(mlst标准.Key, 2)
    
    strOldOwner = GetOwnerName(lngOldSystem, gcnOracle)
    On Error GoTo errHandle
    'zlComponent数据
    gstrSQL = "insert into zlComponent(部件,名称,主版本,次版本,附版本,系统) " & _
                "select 部件,名称,主版本,次版本,附版本," & lngNewSystem & " from zlComponent where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 5
    
    'zlPrograms数据
    gstrSQL = "insert into zlPrograms(序号,标题,说明,部件,系统) " & _
                "select 序号,标题,说明,部件," & lngNewSystem & " from zlPrograms where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 20
    
    'zlProgFuncs数据
    gstrSQL = "insert into zlProgFuncs(序号,功能,系统) " & _
                "select 序号,功能," & lngNewSystem & " from zlProgFuncs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 35
    
    'zlProgPrivs数据
    gstrSQL = "insert into zlProgPrivs(序号,功能,所有者,对象,权限,系统) " & _
                "select 序号,功能,decode(所有者,'" & strOldOwner & "',user,所有者),对象,权限," & lngNewSystem & " from zlProgPrivs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 65
    
    'zlMenus数据
    '清理无效菜单
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            .Open gstrSQL, cnExecuter
            If .EOF Then Exit Do
            gstrSQL = "delete from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            cnExecuter.Execute gstrSQL
        Loop
    End With
    CopyMenu gcnOracle, lngOldSystem, lngNewSystem
    pgbState.value = 85
    
    'zlBaseCode数据
    gstrSQL = "insert into zlBaseCode(表名,固定,说明,分类,系统) " & _
                "select 表名,固定,说明,分类," & lngNewSystem & " from zlBaseCode where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 90
    
    'zlbaktables数据
    gstrSQL = "Insert Into zltools.zlbaktables (系统, 表名, 组号, 序号, 直接转出) select " & lngNewSystem & ", 表名, 组号, 序号, 直接转出 from zltools.zlbaktables where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    
    'zlDataMove数据
    gstrSQL = "insert into zlDataMove(组号,组名,说明,日期字段,转出描述,上次日期,系统,状态) " & _
                "select 组号,组名,说明,日期字段,转出描述,上次日期," & lngNewSystem & ",状态 from zlDataMove where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 95
    
    'zlAutoJobs数据
    gstrSQL = "insert into zlAutoJobs(类型,序号,名称,说明,内容,参数,执行时间,间隔时间,系统) " & _
                "select 类型,序号,名称,说明,内容,参数,执行时间,间隔时间," & lngNewSystem & " from zlAutoJobs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 97
    
    'zlParameters数据
    gstrSQL = "Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明) " & _
            " Select zlParameters_ID.Nextval," & lngNewSystem & ",模块,私有,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 99
    
    pgbState.value = 0
    pgbState.Visible = True
    CopyManageData = True
    Exit Function
errHandle:
    If MsgBox("出现下列错误，是否继续？" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    pgbState.value = 0
    pgbState.Visible = True
    
End Function

Private Sub ChkSequence()
    '----------------------------------------------
    '功能：整理序列的当前号码
    '----------------------------------------------
    Dim rsLst As ADODB.Recordset
    
    pgbState.value = 0
    SetProgressVisible True
    
    Set rsLst = GetSequence("", mcnOwner)
    With rsLst
        Do Until .EOF
            DoEvents
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            Call AdjustNameSequece(!Owner & "." & !Table_Name, mcnOwner, !Column_Name)
            .MoveNext
        Loop
    End With
    Call Adjust结帐ID(mcnOwner)
    
    pgbState.value = 0
    SetProgressVisible False
End Sub

Private Function RunSQLScript(ByVal strFile As String) As Boolean
'功能：执行SQL脚本
'      strFile=SQL脚本名
'返回：RunSQLScript=文件是否执行成功
    Dim strTmp As String
    Dim strTmpPath As String
    Dim strCaprion As String
    
    With mclsRunScript
        .ProcMode = 0
        pgbState.value = 0
        If .OpenFile(strFile) Then
            Do While Not .EOF
                pgbState.value = .ProcessValue
                Call .CollectRoles
                If .ExecuteSQL = False Then Exit Function
                Call .ReadNextSQL
            Loop
            RunSQLScript = True
        Else
            RunSQLScript = False
        End If
    End With
End Function

