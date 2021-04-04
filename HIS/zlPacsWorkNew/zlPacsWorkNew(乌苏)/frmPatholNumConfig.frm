VERSION 5.00
Begin VB.Form frmPatholNumConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病理号配置"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150
   Icon            =   "frmPatholNumConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton opt2 
      Caption         =   "2位"
      Height          =   255
      Left            =   1800
      TabIndex        =   52
      Top             =   4650
      Width           =   615
   End
   Begin VB.OptionButton opt1 
      Caption         =   "4位"
      Height          =   255
      Left            =   1200
      TabIndex        =   51
      Top             =   4650
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   7680
      TabIndex        =   49
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   6240
      TabIndex        =   48
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.ComboBox cbxShiJianGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":000C
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   3720
         Width           =   975
      End
      Begin VB.ComboBox cbxChangGuiGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":004B
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1120
         Width           =   975
      End
      Begin VB.ComboBox cbxKuaiPianGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":008A
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":00A0
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1640
         Width           =   975
      End
      Begin VB.ComboBox cbxBingDongGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":00C9
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":00DF
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox cbxXiBaoGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":0108
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":011E
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2680
         Width           =   975
      End
      Begin VB.ComboBox cbxHuiZhenGuiZe 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":0147
         Left            =   7440
         List            =   "frmPatholNumConfig.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   3200
         Width           =   975
      End
      Begin VB.TextBox txtKuaiPianPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   58
         Top             =   1640
         Width           =   615
      End
      Begin VB.CheckBox chkKuaiPianYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   57
         Top             =   1650
         Width           =   255
      End
      Begin VB.CheckBox chkKuaiPianMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   56
         Top             =   1650
         Width           =   255
      End
      Begin VB.CheckBox chkKuaiPianDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   55
         Top             =   1650
         Width           =   255
      End
      Begin VB.ComboBox cbxKuaiPianNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":0186
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":01A5
         TabIndex        =   54
         Text            =   "5"
         Top             =   1640
         Width           =   855
      End
      Begin VB.TextBox txtKuaiPianStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   53
         Text            =   "1"
         Top             =   1640
         Width           =   1095
      End
      Begin VB.TextBox txtChangGuiStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   47
         Text            =   "1"
         Top             =   1120
         Width           =   1095
      End
      Begin VB.TextBox txtBingDongStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   46
         Text            =   "1"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtXiBaoStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   45
         Text            =   "1"
         Top             =   2680
         Width           =   1095
      End
      Begin VB.TextBox txtHuiZhenStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   44
         Text            =   "1"
         Top             =   3200
         Width           =   1095
      End
      Begin VB.TextBox txtShiJianStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   43
         Text            =   "1"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtSameStart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   42
         Text            =   "1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cbxShiJianNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":01CD
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":01EC
         TabIndex        =   41
         Text            =   "5"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox cbxHuiZhenNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":0214
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":0233
         TabIndex        =   40
         Text            =   "5"
         Top             =   3200
         Width           =   855
      End
      Begin VB.ComboBox cbxBingDongNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":025B
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":027A
         TabIndex        =   39
         Text            =   "5"
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cbxXiBaoNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":02A2
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":02C1
         TabIndex        =   38
         Text            =   "5"
         Top             =   2680
         Width           =   855
      End
      Begin VB.ComboBox cbxChangGuiNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":02E9
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":0308
         TabIndex        =   37
         Text            =   "5"
         Top             =   1120
         Width           =   855
      End
      Begin VB.ComboBox cbxSameNumLen 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholNumConfig.frx":0330
         Left            =   4680
         List            =   "frmPatholNumConfig.frx":034F
         TabIndex        =   36
         Text            =   "5"
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkBingDongDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   2175
         Width           =   255
      End
      Begin VB.CheckBox chkXiBaoDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   2700
         Width           =   255
      End
      Begin VB.CheckBox chkHuiZhenDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   3225
         Width           =   255
      End
      Begin VB.CheckBox chkShiJianDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   3755
         Width           =   255
      End
      Begin VB.CheckBox chkChangGuiDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   1125
         Width           =   255
      End
      Begin VB.CheckBox chkSameDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkChangGuiMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   1125
         Width           =   255
      End
      Begin VB.CheckBox chkBingDongMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   2175
         Width           =   255
      End
      Begin VB.CheckBox chkXiBaoMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   2700
         Width           =   255
      End
      Begin VB.CheckBox chkHuiZhenMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   3225
         Width           =   255
      End
      Begin VB.CheckBox chkShiJianMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   3755
         Width           =   255
      End
      Begin VB.CheckBox chkSameMonth 
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkChangGuiYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   1125
         Width           =   255
      End
      Begin VB.CheckBox chkBingDongYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   2175
         Width           =   255
      End
      Begin VB.CheckBox chkXiBaoYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   2700
         Width           =   255
      End
      Begin VB.CheckBox chkHuiZhenYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   3225
         Width           =   255
      End
      Begin VB.CheckBox chkShiJianYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   3755
         Width           =   255
      End
      Begin VB.CheckBox chkSameYear 
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtShiJianPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   17
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtChangGuiPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1120
         Width           =   615
      End
      Begin VB.TextBox txtBingDongPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtXiBaoPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   14
         Top             =   2680
         Width           =   615
      End
      Begin VB.TextBox txtHuiZhenPrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   13
         Top             =   3200
         Width           =   615
      End
      Begin VB.TextBox txtSamePrefix 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkUseSameRule 
         Caption         =   "使用相同规则"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "与[X]规则相同"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label labKuaiPianInf 
         Caption         =   "(快速石蜡)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1400
         TabIndex        =   61
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label labChangGuiInf 
         Caption         =   "(常规石蜡)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1400
         TabIndex        =   60
         Top             =   1360
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   840
         X2              =   1440
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label labKuaiPian 
         Caption         =   "快 片"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   1683
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "起始数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "序号位数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "年  月  日"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3330
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "前缀"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label labShiJian 
         Caption         =   "尸 检"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label labHuiZhen 
         Caption         =   "会 诊"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   3300
         Width           =   615
      End
      Begin VB.Label labXiBao 
         Caption         =   "细 胞"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   2761
         Width           =   615
      End
      Begin VB.Label labBingDong 
         Caption         =   "冰 冻"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   2222
         Width           =   615
      End
      Begin VB.Label labChangGui 
         Caption         =   "常 规"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1144
         Width           =   615
      End
      Begin VB.Line Line8 
         X1              =   840
         X2              =   1440
         Y1              =   3417
         Y2              =   3417
      End
      Begin VB.Line Line7 
         X1              =   840
         X2              =   1440
         Y1              =   2874
         Y2              =   2874
      End
      Begin VB.Line Line6 
         X1              =   840
         X2              =   1440
         Y1              =   2331
         Y2              =   2331
      End
      Begin VB.Line Line5 
         X1              =   840
         X2              =   1440
         Y1              =   1788
         Y2              =   1788
      End
      Begin VB.Line Line4 
         X1              =   840
         X2              =   1440
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line2 
         X1              =   840
         X2              =   840
         Y1              =   1245
         Y2              =   3960
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   840
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "病  理  号  码  生  成  规  则 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   3255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Label Label4 
      Caption         =   "规则中使用              年份"
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   4680
      Width           =   2535
   End
End
Attribute VB_Name = "frmPatholNumConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function GetSameRuleType(ByVal strRuleName As String) As Long
'获取与其他规则相同的规则类型
    GetSameRuleType = -2
    
    If chkUseSameRule.value Then
        GetSameRuleType = -1
        Exit Function
    End If


    Select Case strRuleName
        Case "常 规"
            GetSameRuleType = 0
        Case "冰 冻"
            GetSameRuleType = 1
        Case "细 胞"
            GetSameRuleType = 2
        Case "会 诊"
            GetSameRuleType = 3
        Case "尸 检"
            GetSameRuleType = 4
        Case "快 片"
            GetSameRuleType = 5
    End Select
    
    
End Function

Private Sub SavePatholNumRule()
'保存病理号规则
Dim strSql As String
Dim lngYearLen As Long

lngYearLen = 4
If opt2.value Then lngYearLen = 2

'通用规则
strSql = "zl_病理号码_规则(-1,'" & txtSamePrefix.Text & "'," & chkSameYear.value & "," & chkSameMonth.value & "," & chkSameDay.value & "," & Val(cbxSameNumLen.Text) & "," & lngYearLen & "," & Val(txtSameStart.Text) & "," & GetSameRuleType("") & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

'常规
strSql = "zl_病理号码_规则(0,'" & txtChangGuiPrefix.Text & "'," & chkChangGuiYear.value & "," & chkChangGuiMonth.value & "," & chkChangGuiDay.value & "," & Val(cbxChangGuiNumLen.Text) & "," & lngYearLen & "," & Val(txtChangGuiStart.Text) & "," & GetSameRuleType(cbxChangGuiGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)


'冰冻
strSql = "zl_病理号码_规则(1,'" & txtBingDongPrefix.Text & "'," & chkBingDongYear.value & "," & chkBingDongMonth.value & "," & chkBingDongDay.value & "," & Val(cbxBingDongNumLen.Text) & "," & lngYearLen & "," & Val(txtBingDongStart.Text) & "," & GetSameRuleType(cbxBingDongGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)


'细胞
strSql = "zl_病理号码_规则(2,'" & txtXiBaoPrefix.Text & "'," & chkXiBaoYear.value & "," & chkXiBaoMonth.value & "," & chkXiBaoDay.value & "," & Val(cbxXiBaoNumLen.Text) & "," & lngYearLen & "," & Val(txtXiBaoStart.Text) & "," & GetSameRuleType(cbxXiBaoGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)


'会诊
strSql = "zl_病理号码_规则(3,'" & txtHuiZhenPrefix.Text & "'," & chkHuiZhenYear.value & "," & chkHuiZhenMonth.value & "," & chkHuiZhenDay.value & "," & Val(cbxHuiZhenNumLen.Text) & "," & lngYearLen & "," & Val(txtHuiZhenStart.Text) & "," & GetSameRuleType(cbxHuiZhenGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)


'尸检
strSql = "zl_病理号码_规则(4,'" & txtShiJianPrefix.Text & "'," & chkShiJianYear.value & "," & chkShiJianMonth.value & "," & chkShiJianDay.value & "," & Val(cbxShiJianNumLen.Text) & "," & lngYearLen & "," & Val(txtShiJianStart.Text) & "," & GetSameRuleType(cbxShiJianGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)


'快速石蜡
strSql = "zl_病理号码_规则(5,'" & txtKuaiPianPrefix.Text & "'," & chkKuaiPianYear.value & "," & chkKuaiPianMonth.value & "," & chkKuaiPianDay.value & "," & Val(cbxKuaiPianNumLen.Text) & "," & lngYearLen & "," & Val(txtKuaiPianStart.Text) & "," & GetSameRuleType(cbxKuaiPianGuiZe.Text) & ")"
Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
End Sub


Private Function GetStartNum(ByVal lngRuleType As Long, ByVal blnUseYear As Boolean, _
    ByVal blnUseMonth As Boolean, ByVal blnUseDay As Boolean) As Long
'获取当前规则所使用到的序号
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    Dim curDate As Date
    
    GetStartNum = -1
    strSql = "select 当前序号 from 病理号码记录 where 类型=[1]"
    
    strWhere = ""
    If blnUseYear Then
        strWhere = strWhere & " and 年=[2]"
    End If
    
    If blnUseMonth Then
        strWhere = strWhere & " and 月=[3]"
    End If
    
    If blnUseDay Then
        strWhere = strWhere & " and 日=[4]"
    End If
    
    curDate = zlDatabase.Currentdate
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql & strWhere, Me.Caption, lngRuleType, Format(curDate, "yyyy"), Format(curDate, "mm"), Format(curDate, "dd"))
    
    If rsData.RecordCount <= 0 Then Exit Function

    GetStartNum = Val(Nvl(rsData!当前序号)) + 1
End Function


Private Function GetComboxIndex(cbxData As ComboBox, ByVal strCbxText As String) As Long
'查找combobox的索引
    Dim i As Long
    
    GetComboxIndex = -1
    For i = 0 To cbxData.ListCount - 1
        If cbxData.list(i) = strCbxText Then
            GetComboxIndex = i
            Exit Function
        End If
    Next i
End Function


Private Sub LoadPatholNumRule()
'载入病理号码规则
Dim strSql As String
Dim rsData As ADODB.Recordset
Dim lngStartNum As Long

strSql = "select * from 病理号码规则"
Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

If rsData.RecordCount <= 0 Then Exit Sub

'通用规则
rsData.Filter = "类型=-1"
If rsData.RecordCount > 0 Then
    txtSamePrefix.Text = Nvl(rsData!前缀)
    chkSameYear.value = Val(Nvl(rsData!年))
    chkSameMonth.value = Val(Nvl(rsData!月))
    chkSameDay.value = Val(Nvl(rsData!日))
    cbxSameNumLen.Text = Nvl(rsData!序号位数)
    
    lngStartNum = GetStartNum(-1, chkSameYear.value, chkSameMonth.value, chkSameDay.value)
    If lngStartNum > 0 Then
        txtSameStart.Text = lngStartNum
    Else
        txtSameStart.Text = Nvl(rsData!起始数)
    End If
    
    chkUseSameRule.value = IIf(Val(Nvl(rsData!相同规则)) = -1, 1, 0)
    If Val(Nvl(rsData!年份位数)) = 2 Then
        opt2.value = True
    Else
        opt1.value = True
    End If
End If


'常规
rsData.Filter = "类型=0"
If rsData.RecordCount > 0 Then
    txtChangGuiPrefix.Text = Nvl(rsData!前缀)
    chkChangGuiYear.value = Val(Nvl(rsData!年))
    chkChangGuiMonth.value = Val(Nvl(rsData!月))
    chkChangGuiDay.value = Val(Nvl(rsData!日))
    cbxChangGuiNumLen.Text = Nvl(rsData!序号位数)
    
    lngStartNum = GetStartNum(0, chkChangGuiYear.value, chkChangGuiMonth.value, chkChangGuiDay.value)
    If lngStartNum > 0 Then
        txtChangGuiStart.Text = lngStartNum
    Else
        txtChangGuiStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxChangGuiGuiZe.ListIndex = 0
        Case -1
            cbxChangGuiGuiZe.ListIndex = 0
        Case 0
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "常 规")
        Case 1
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "冰 冻")
        Case 2
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "细 胞")
        Case 3
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "会 诊")
        Case 4
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "尸 检")
        Case 5
            cbxChangGuiGuiZe.ListIndex = GetComboxIndex(cbxChangGuiGuiZe, "快 片")
    End Select
End If


'冰冻
rsData.Filter = "类型=1"
If rsData.RecordCount > 0 Then
    txtBingDongPrefix.Text = Nvl(rsData!前缀)
    chkBingDongYear.value = Val(Nvl(rsData!年))
    chkBingDongMonth.value = Val(Nvl(rsData!月))
    chkBingDongDay.value = Val(Nvl(rsData!日))
    cbxBingDongNumLen.Text = Nvl(rsData!序号位数)
    
    lngStartNum = GetStartNum(1, chkBingDongYear.value, chkBingDongMonth.value, chkBingDongDay.value)
    If lngStartNum > 0 Then
        txtBingDongStart.Text = lngStartNum
    Else
        txtBingDongStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxBingDongGuiZe.ListIndex = 0
        Case -1
            cbxBingDongGuiZe.ListIndex = 0
        Case 0
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "常 规")
        Case 1
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "冰 冻")
        Case 2
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "细 胞")
        Case 3
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "会 诊")
        Case 4
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "尸 检")
        Case 5
            cbxBingDongGuiZe.ListIndex = GetComboxIndex(cbxBingDongGuiZe, "快 片")
    End Select
End If


'细胞
rsData.Filter = "类型=2"
If rsData.RecordCount > 0 Then
    txtXiBaoPrefix.Text = Nvl(rsData!前缀)
    chkXiBaoYear.value = Val(Nvl(rsData!年))
    chkXiBaoMonth.value = Val(Nvl(rsData!月))
    chkXiBaoDay.value = Val(Nvl(rsData!日))
    cbxXiBaoNumLen.Text = Nvl(rsData!序号位数)

    lngStartNum = GetStartNum(2, chkXiBaoYear.value, chkXiBaoMonth.value, chkXiBaoDay.value)
    If lngStartNum > 0 Then
        txtXiBaoStart.Text = lngStartNum
    Else
        txtXiBaoStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxXiBaoGuiZe.ListIndex = 0
        Case -1
            cbxXiBaoGuiZe.ListIndex = 0
        Case 0
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "常 规")
        Case 1
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "冰 冻")
        Case 2
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "细 胞")
        Case 3
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "会 诊")
        Case 4
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "尸 检")
        Case 5
            cbxXiBaoGuiZe.ListIndex = GetComboxIndex(cbxXiBaoGuiZe, "快 片")
    End Select
End If


'会诊
rsData.Filter = "类型=3"
If rsData.RecordCount > 0 Then
    txtHuiZhenPrefix.Text = Nvl(rsData!前缀)
    chkHuiZhenYear.value = Val(Nvl(rsData!年))
    chkHuiZhenMonth.value = Val(Nvl(rsData!月))
    chkHuiZhenDay.value = Val(Nvl(rsData!日))
    cbxHuiZhenNumLen.Text = Nvl(rsData!序号位数)

    lngStartNum = GetStartNum(3, chkHuiZhenYear.value, chkHuiZhenMonth.value, chkHuiZhenDay.value)
    If lngStartNum > 0 Then
        txtHuiZhenStart.Text = lngStartNum
    Else
        txtHuiZhenStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxHuiZhenGuiZe.ListIndex = 0
        Case -1
            cbxHuiZhenGuiZe.ListIndex = 0
        Case 0
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "常 规")
        Case 1
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "冰 冻")
        Case 2
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "细 胞")
        Case 3
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "会 诊")
        Case 4
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "尸 检")
        Case 5
            cbxHuiZhenGuiZe.ListIndex = GetComboxIndex(cbxHuiZhenGuiZe, "快 片")
    End Select
End If


'尸检
rsData.Filter = "类型=4"
If rsData.RecordCount > 0 Then
    txtShiJianPrefix.Text = Nvl(rsData!前缀)
    chkShiJianYear.value = Val(Nvl(rsData!年))
    chkShiJianMonth.value = Val(Nvl(rsData!月))
    chkShiJianDay.value = Val(Nvl(rsData!日))
    cbxShiJianNumLen.Text = Nvl(rsData!序号位数)
    
    lngStartNum = GetStartNum(4, chkShiJianYear.value, chkShiJianMonth.value, chkShiJianDay.value)
    If lngStartNum > 0 Then
        txtShiJianStart.Text = lngStartNum
    Else
        txtShiJianStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxShiJianGuiZe.ListIndex = 0
        Case -1
            cbxShiJianGuiZe.ListIndex = 0
        Case 0
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "常 规")
        Case 1
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "冰 冻")
        Case 2
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "细 胞")
        Case 3
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "会 诊")
        Case 4
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "尸 检")
        Case 5
            cbxShiJianGuiZe.ListIndex = GetComboxIndex(cbxShiJianGuiZe, "快 片")
    End Select
End If


'快速石蜡
rsData.Filter = "类型=5"
If rsData.RecordCount > 0 Then
    txtKuaiPianPrefix.Text = Nvl(rsData!前缀)
    chkKuaiPianYear.value = Val(Nvl(rsData!年))
    chkKuaiPianMonth.value = Val(Nvl(rsData!月))
    chkKuaiPianDay.value = Val(Nvl(rsData!日))
    cbxKuaiPianNumLen.Text = Nvl(rsData!序号位数)
    txtKuaiPianStart.Text = Nvl(rsData!起始数)
    
    lngStartNum = GetStartNum(5, chkKuaiPianYear.value, chkKuaiPianMonth.value, chkKuaiPianDay.value)
    If lngStartNum > 0 Then
        txtKuaiPianStart.Text = lngStartNum
    Else
        txtKuaiPianStart.Text = Nvl(rsData!起始数)
    End If
    
    Select Case Val(Nvl(rsData!相同规则))
        Case -2
            cbxKuaiPianGuiZe.ListIndex = 0
        Case -1
            cbxKuaiPianGuiZe.ListIndex = 0
        Case 0
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "常 规")
        Case 1
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "冰 冻")
        Case 2
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "细 胞")
        Case 3
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "会 诊")
        Case 4
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "尸 检")
        Case 5
            cbxKuaiPianGuiZe.ListIndex = GetComboxIndex(cbxKuaiPianGuiZe, "快 片")
    End Select
End If
End Sub


Private Sub chkUseSameRule_Click()
On Error Resume Next
    Call ChangeConfigFace(chkUseSameRule.value)
End Sub


Private Sub ChangeConfigFace(ByVal blnIsSameRule As Boolean)
    txtChangGuiPrefix.Enabled = Not blnIsSameRule
    txtBingDongPrefix.Enabled = Not blnIsSameRule
    txtXiBaoPrefix.Enabled = Not blnIsSameRule
    txtHuiZhenPrefix.Enabled = Not blnIsSameRule
    txtShiJianPrefix.Enabled = Not blnIsSameRule
    txtKuaiPianPrefix.Enabled = Not blnIsSameRule
    
    chkChangGuiYear.Enabled = Not blnIsSameRule
    chkBingDongYear.Enabled = Not blnIsSameRule
    chkXiBaoYear.Enabled = Not blnIsSameRule
    chkHuiZhenYear.Enabled = Not blnIsSameRule
    chkShiJianYear.Enabled = Not blnIsSameRule
    chkKuaiPianYear.Enabled = Not blnIsSameRule
    
    chkChangGuiMonth.Enabled = Not blnIsSameRule
    chkBingDongMonth.Enabled = Not blnIsSameRule
    chkXiBaoMonth.Enabled = Not blnIsSameRule
    chkHuiZhenMonth.Enabled = Not blnIsSameRule
    chkShiJianMonth.Enabled = Not blnIsSameRule
    chkKuaiPianMonth.Enabled = Not blnIsSameRule
    
    chkChangGuiDay.Enabled = Not blnIsSameRule
    chkBingDongDay.Enabled = Not blnIsSameRule
    chkXiBaoDay.Enabled = Not blnIsSameRule
    chkHuiZhenDay.Enabled = Not blnIsSameRule
    chkShiJianDay.Enabled = Not blnIsSameRule
    chkKuaiPianDay.Enabled = Not blnIsSameRule
    
    cbxChangGuiNumLen.Enabled = Not blnIsSameRule
    cbxBingDongNumLen.Enabled = Not blnIsSameRule
    cbxXiBaoNumLen.Enabled = Not blnIsSameRule
    cbxHuiZhenNumLen.Enabled = Not blnIsSameRule
    cbxShiJianNumLen.Enabled = Not blnIsSameRule
    cbxKuaiPianNumLen.Enabled = Not blnIsSameRule
    
    
    cbxChangGuiGuiZe.Enabled = Not blnIsSameRule
    cbxBingDongGuiZe.Enabled = Not blnIsSameRule
    cbxXiBaoGuiZe.Enabled = Not blnIsSameRule
    cbxHuiZhenGuiZe.Enabled = Not blnIsSameRule
    cbxShiJianGuiZe.Enabled = Not blnIsSameRule
    cbxKuaiPianGuiZe.Enabled = Not blnIsSameRule
    
    
    txtChangGuiStart.Enabled = Not blnIsSameRule
    txtBingDongStart.Enabled = Not blnIsSameRule
    txtXiBaoStart.Enabled = Not blnIsSameRule
    txtHuiZhenStart.Enabled = Not blnIsSameRule
    txtShiJianStart.Enabled = Not blnIsSameRule
    txtKuaiPianStart.Enabled = Not blnIsSameRule
        
    labChangGui.Enabled = Not blnIsSameRule
    labBingDong.Enabled = Not blnIsSameRule
    labXiBao.Enabled = Not blnIsSameRule
    labHuiZhen.Enabled = Not blnIsSameRule
    labShiJian.Enabled = Not blnIsSameRule
    labKuaiPian.Enabled = Not blnIsSameRule
    
    labChangGuiInf.Enabled = Not blnIsSameRule
    labKuaiPianInf.Enabled = Not blnIsSameRule
End Sub


Private Sub cmdCancel_Click()
On Error Resume Next
    Me.Hide
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Call SavePatholNumRule
    
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadPatholNumRule
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
