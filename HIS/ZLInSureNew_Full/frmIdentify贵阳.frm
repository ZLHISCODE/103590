VERSION 5.00
Begin VB.Form frmIdentify贵阳 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmIdentify贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt工伤信息 
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1860
      TabIndex        =   75
      Top             =   7860
      Width           =   7515
   End
   Begin VB.TextBox txt备注 
      Enabled         =   0   'False
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
      IMEMode         =   3  'DISABLE
      Left            =   1860
      TabIndex        =   69
      Top             =   7440
      Width           =   7515
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "改密码(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   13
      Top             =   510
      Width           =   1335
   End
   Begin VB.TextBox txt封锁信息 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1860
      TabIndex        =   67
      Top             =   7050
      Width           =   7515
   End
   Begin VB.Frame Frame2 
      Caption         =   "累计信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   180
      TabIndex        =   41
      Top             =   4050
      Width           =   9195
      Begin VB.TextBox txt普通门诊医疗补助结转可使用 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   65
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   63
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助起付标准 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   61
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   59
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   57
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt大额支付累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   55
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txt大额统筹限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   53
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt统筹支付累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   51
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt基本统筹限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   49
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt已支付起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   47
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   45
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt住院次数 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   43
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label lbl普通门诊医疗补助结转可使用 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助结转可使用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4200
         TabIndex        =   64
         Top             =   2340
         Width           =   2730
      End
      Begin VB.Label lbl公务员门诊补助起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4620
         TabIndex        =   62
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Label lbl公务员门诊补助起付标准 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助起付标准"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4410
         TabIndex        =   60
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lbl公务员门诊补助累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   58
         Top             =   1170
         Width           =   2100
      End
      Begin VB.Label lbl公务员门诊补助限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   56
         Top             =   780
         Width           =   2100
      End
      Begin VB.Label lbl大额支付累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额支付累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5670
         TabIndex        =   54
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label lbl大额统筹限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   52
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label lbl统筹支付累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   50
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lbl基本统筹限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "基本统筹限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   48
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lbl已支付起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已支付起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   46
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label lbl起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1020
         TabIndex        =   44
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   42
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   180
      TabIndex        =   28
      Top             =   30
      Width           =   9195
      Begin VB.CommandButton cmd启动 
         Caption         =   "启动"
         Height          =   350
         Left            =   2940
         TabIndex        =   7
         ToolTipText     =   "启动密码键盘"
         Top             =   1875
         Width           =   675
      End
      Begin VB.CommandButton cmd读卡 
         Caption         =   "读卡"
         Height          =   350
         Left            =   3600
         TabIndex        =   76
         Top             =   1470
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.OptionButton opt卡类别 
         Caption         =   "身份证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   2
         Left            =   3060
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1365
      End
      Begin VB.OptionButton opt卡类别 
         Caption         =   "IC卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   1950
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   945
      End
      Begin VB.OptionButton opt卡类别 
         Caption         =   "磁卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   1170
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CheckBox chk转院治疗 
         Caption         =   "转院治疗"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3210
         TabIndex        =   12
         Top             =   2730
         Width           =   1185
      End
      Begin VB.TextBox txt人员类别 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   72
         Top             =   338
         Width           =   2595
      End
      Begin VB.TextBox txt性别 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7980
         TabIndex        =   70
         Top             =   1118
         Width           =   975
      End
      Begin VB.CheckBox chk工伤康复住院 
         Caption         =   "工伤康复住院"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   10
         Top             =   2730
         Width           =   1635
      End
      Begin VB.TextBox txt处方本编号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   9
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox cbo保险类别 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   2595
      End
      Begin VB.CheckBox chk计划生育 
         Caption         =   "计划生育"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1980
         TabIndex        =   11
         Top             =   2730
         Width           =   1185
      End
      Begin VB.TextBox txt缴费年度 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   40
         Top             =   3450
         Width           =   2595
      End
      Begin VB.TextBox txt帐户余额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   38
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt单位名称 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   36
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txt单位编码 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt出生日期 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txt身份证号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   24
         Top             =   1500
         Width           =   2595
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   20
         Top             =   1118
         Width           =   1065
      End
      Begin VB.TextBox txt医疗照顾人群 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   17
         Top             =   728
         Width           =   2595
      End
      Begin VB.TextBox txt分中心编号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   3450
         Width           =   2595
      End
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt密码 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1890
         Width           =   1245
      End
      Begin VB.ComboBox cbo支付类别 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2595
      End
      Begin VB.TextBox txt卡号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1500
         Width           =   2565
      End
      Begin VB.Label lbl人员类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   73
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7530
         TabIndex        =   71
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lbl处方本编号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方本编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   600
         TabIndex        =   25
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label lbl保险类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "保险类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   0
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl缴费年度 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "缴费年度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   39
         Top             =   3502
         Width           =   840
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   37
         Top             =   3112
         Width           =   840
      End
      Begin VB.Label lbl单位名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   35
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lbl单位编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位编码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   33
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label lbl出生日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   26
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl身份证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   23
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5880
         TabIndex        =   19
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lbl医疗照顾人群 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗照顾人群"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   16
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lbl分中心编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分中心编码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   600
         TabIndex        =   31
         Top             =   3502
         Width           =   1050
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   29
         Top             =   3112
         Width           =   840
      End
      Begin VB.Label lbl密码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   22
         Top             =   1950
         Width           =   420
      End
      Begin VB.Label lbl支付类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "支付类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   18
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   21
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Label Lab工伤信息 
      BackStyle       =   0  'Transparent
      Caption         =   "工伤信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   900
      TabIndex        =   74
      Top             =   7920
      Width           =   840
   End
   Begin VB.Label lbl备注 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1350
      TabIndex        =   68
      Top             =   7500
      Width           =   420
   End
   Begin VB.Label lbl封锁信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "封锁信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   930
      TabIndex        =   66
      Top             =   7110
      Width           =   840
   End
End
Attribute VB_Name = "frmIdentify贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr就诊时间 As String
Private mbytType As Byte
Private mstr分中心编号 As String
Private mstr保险类别 As String
Private mstr新密码 As String
Private mbln生育标志 As Boolean
Private mblnOK As Boolean
Private int门诊住院标志 As Integer   '门诊-0,住院-1
Private mstr工伤认定编号 As String
Private mlng病种ID As Long
Private mstr病种 As String
Private mstr支付类型 As String  '支付类别 31：住院；32:计划生育；37：转院

Private Sub cbo保险类别_Click()
    chk计划生育.Enabled = (cbo保险类别.Text Like "*生育*")
    chk计划生育.Value = 0
    chk转院治疗.Value = 0
    txt处方本编号.Enabled = False
    chk工伤康复住院.Enabled = (cbo保险类别.Text = "工伤保险" And mbytType = 0)
    If cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Then
        Me.cbo支付类别.ListIndex = 1
    End If
    'XieRong 2010.10.12 特殊门诊病人必须选择处方版本号
    txt处方本编号.Enabled = IIf(cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Or cbo支付类别.Text = "特殊门诊", True, False)
    lbl处方本编号.Enabled = IIf(cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Or cbo支付类别.Text = "特殊门诊", True, False)
   
End Sub

Private Sub cbo支付类别_Click()
    'XieRong 2010.10.12 特殊门诊病人必须选择处方版本号
    txt处方本编号.Enabled = IIf(cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Or cbo支付类别.Text = "特殊门诊", True, False)
    lbl处方本编号.Enabled = IIf(cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Or cbo支付类别.Text = "特殊门诊", True, False)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassWord_Click()
    Dim strNewPass As String
    strNewPass = frm修改密码.ChangePassword("", Me.txt密码.Text, 40)
    If strNewPass <> "" Then mstr新密码 = strNewPass
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    If cmdOK.Enabled = False Then Exit Sub
    If Trim(txt卡号.Text) = "" And opt卡类别(2).Value = False Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txt医保号.Text) = "" Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(mstr新密码) <> "" Then
        If 更改密码_贵阳市(txt卡号.Tag, txt密码.Text, mstr新密码) = False Then Exit Sub
        mstr密码 = mstr新密码
        mstr新密码 = ""
        txt密码.Text = mstr密码
    End If
    If chk计划生育.Value = 1 And chk转院治疗.Value = 1 Then MsgBox "“计划生育”和“转院治疗”只能选择其中之一，且不能选择错误！", vbInformation, gstrSysName: Exit Sub
    '2005.11.22,int门诊住院标志,住院强制选择保险类别
    If (int门诊住院标志 = 1 And cbo保险类别.ListIndex = 0) Then
       MsgBox "请选择保险类别！", vbInformation, gstrSysName
       cbo保险类别.SetFocus
       Exit Sub
    End If
        
    'XieRong 2010.10.12 特殊门诊病人必须选择处方版本号
    If cbo保险类别.Text = "居民保险" And (mbytType = 0 Or mbytType = 3) Or cbo支付类别.Text = "特殊门诊" Then
        If Trim(txt处方本编号.Text) = "" Then
            MsgBox "特殊门诊就诊必须录入处方本编号!", vbInformation, gstrSysName
            txt处方本编号.SetFocus
            Exit Sub
        End If
    End If
    If Me.cbo保险类别.Text = "居民医保" And Trim(txt处方本编号.Text) = "" Then
        MsgBox "居民医保就诊必须录入处方本编号!", vbInformation, gstrSysName
        txt处方本编号.SetFocus
        Exit Sub
    End If
    
    '20111116周玉强增加,是否为特殊门诊或工伤病人
     gstr门诊标志 = cbo支付类别.Text
     gstr工伤标志 = cbo保险类别.Text
    
    '有可能修改了密码，造成病人身份验证后返回的XML被破坏，再次调用读卡
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt卡号.Tag)            ' 磁卡数据
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt密码.Text)            ' 密码
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo支付类别.ItemData(Me.cbo支付类别.ListIndex))            ' 支付类别
    If Me.cbo支付类别.Text = "特殊门诊" Then Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", mstr病种)            ' 特殊病
    
    '2005.11.22,int门诊住院标志,医保返回
    If int门诊住院标志 = 0 Then
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex + 1)
    Else
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "GSRDBH", mstr工伤认定编号)
    Call InsertChild(mdomInput.documentElement, "STARTDATE", mstr就诊时间)           ' 开始时间
    '调用接口
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'mstr卡号 = Trim(txt卡号.Text)
    '医保接口未返回卡号，以前的卡号字段改为保存磁卡数据，后面虚拟结算要用
    mstr卡号 = Me.txt卡号.Tag
    mstr医保号 = Trim(txt医保号.Text)
    mstr分中心编号 = Trim(txt分中心编号.Text)
    mstr密码 = Trim(txt密码.Text)
    mstr保险类别 = cbo保险类别.ListIndex + 1
    mbln生育标志 = (chk计划生育.Value = 1)
    gint工伤康复住院 = Me.chk工伤康复住院.Value
    
    gstr处方本号 = txt处方本编号.Text
    'Add By 程池富 2010-01-16 贵阳医学院要求
    If chk计划生育.Value = 1 Then
        mstr支付类型 = "32" '计划生育
    ElseIf chk转院治疗.Value = 1 Then
        mstr支付类型 = "37" '转入住院
    Else
        mstr支付类型 = "31" '正常住院
    End If
    '保存此病人的医保档案
'    医保号_IN IN 医保病人档案.医保号%TYPE,
'    住院次数_IN IN 医保病人档案.住院次数%TYPE,
'    起付线_IN IN 医保病人档案.起付线%TYPE,
'    已支付起付线_IN IN 医保病人档案.已支付起付线%TYPE,
'    基本统筹限额_IN IN 医保病人档案.基本统筹限额%TYPE,
'    统筹支付累计_IN IN 医保病人档案.统筹支付累计%TYPE,
'    大额统筹限额_IN IN 医保病人档案.大额统筹限额%TYPE,
'    大额支付累计_IN IN 医保病人档案.大额支付累计%TYPE,
'    公务员补助限额_IN IN 医保病人档案.公务员补助限额%TYPE,
'    公务员补助累计_IN IN 医保病人档案.公务员补助累计%TYPE,
'    公务员起付标准_IN IN 医保病人档案.公务员起付标准%TYPE,
'    公务员补助起付线_IN IN 医保病人档案.公务员补助起付线%TYPE,
'    参加75公务员补助_IN IN 医保病人档案.参加75公务员补助%TYPE)
    On Error GoTo errHand
     '20110812周玉强将：txt备注.Text改为txt备注+Txt工伤信息,主要为了保存工伤信息
    gstrSQL = "zl_医保病人档案_INSERT(" & _
        "'" & mstr医保号 & "'," & Val(txt住院次数.Text) & "," & Val(txt起付线.Text) & "," & Val(txt已支付起付线.Text) & "," & _
        "" & Val(txt基本统筹限额.Text) & "," & Val(txt统筹支付累计.Text) & "," & Val(txt大额统筹限额.Text) & "," & Val(txt大额支付累计.Text) & "," & _
        "" & Val(txt普通门诊医疗补助限额.Text) & "," & Val(txt普通门诊医疗补助累计.Text) & "," & Val(txt普通门诊医疗补助起付标准.Text) & "," & _
        "" & Val(txt普通门诊医疗补助起付线.Text) & ",'" & txt普通门诊医疗补助结转可使用.Text & "','" & txt备注.Text & "|" & Txt工伤信息.Text & "','" & gstr处方本号 & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetIdentify(ByVal bytType As Byte, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, _
    Optional ByRef bln生育标志 As Boolean = False, Optional lng病种ID As Long, Optional str支付类型 As String = "31") As Boolean
    mblnOK = False
    mstr卡号 = str卡号
    mstr医保号 = str医保号
    mstr密码 = ""
    mstr新密码 = ""
    mstr工伤认定编号 = ""
    mbytType = bytType
    mlng病种ID = 0
    mstr支付类型 = ""
    
    gstrSNO = ""
    gstrIDNO = ""
    gstrPSAMNO = ""
    gintType = 1
    frmIdentify贵阳.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str卡号 = mstr卡号 & "^" & mstr保险类别
        str医保号 = mstr医保号
        str分中心编号 = mstr分中心编号
        str密码 = mstr密码
        bln生育标志 = mbln生育标志
        gstr工伤认定编号 = mstr工伤认定编号
        lng病种ID = mlng病种ID
        str支付类型 = mstr支付类型
    End If
End Function

Private Sub cmd读卡_Click()
    Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "公共模块\贵阳市医保", "端口", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If
    
    '打开读卡器
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '读取PSAM芯片号码
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    '读取社保卡号
    STRERR = Space(2000)
    gstrSNO = Space(2000)
    strPin = "000000"
    strAddr = "MF|EF05|07|$MF|EF06|01|$"
    If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, gstrSNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrSNO = TruncZero(gstrSNO)
    gstrIDNO = Split(gstrSNO, "|")(5)
    gstrSNO = Split(gstrSNO, "|")(2)
    txt卡号.Text = gstrSNO
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt密码.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    
    Call txt密码_KeyDown(vbKeyReturn, 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub

Private Sub cmd启动_Click()
        Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "公共模块\贵阳市医保", "端口", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If
    
    '打开读卡器
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '读取PSAM芯片号码
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt密码.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    
    Call txt密码_KeyDown(vbKeyReturn, 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub


Private Sub Form_Activate()
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mstr就诊时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If mbytType = 1 And mstr医保号 <> "" Then
        gstrSQL = " Select 病人ID,保险类别 From 保险帐户 Where 险类=[1] And 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险类别", TYPE_贵阳市, mstr医保号)
        If rsTemp.RecordCount <> 0 Then
            lng病人ID = rsTemp!病人ID
            Me.cbo保险类别.ListIndex = Nvl(rsTemp!保险类别, 0)
            
        End If
        '取入院日期
        gstrSQL = " Select A.入院日期 From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And A.病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期", lng病人ID)
        mstr就诊时间 = Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss")
        
        If txt卡号.Enabled Then Me.txt卡号.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '2005.11.22,int门诊住院标志,加载cbo保险类别itemdata
    With cbo支付类别
        .Clear
        If mbytType = 0 Or mbytType = 3 Then
            int门诊住院标志 = 0
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 11
            .AddItem "特殊门诊"
            .ItemData(.NewIndex) = 18
            With cbo保险类别
                .Clear
                .AddItem "企业职工基本医疗保险"
                .AddItem "企业离休医疗保险"
                .AddItem "机关事业单位医疗保险"
                .AddItem "生育保险"
                .AddItem "机关事业单位生育保险"
                .AddItem "居民保险"
                .AddItem "工伤保险"
                .ListIndex = 0
            End With
            .ListIndex = 0
         Else
            int门诊住院标志 = 1
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 31
            With cbo保险类别
                .Clear
                .AddItem ""
                .AddItem "企业职工基本医疗保险"
                .AddItem "企业离休医疗保险"
                .AddItem "机关事业单位医疗保险"
                .AddItem "生育保险"
                .AddItem "机关事业单位生育保险"
                .AddItem "居民保险"
                .AddItem "工伤保险"
                .ListIndex = 0
            End With
         End If
        .ListIndex = 0
    End With
    chk转院治疗.Visible = False '金保二期后未用，屏掉 2010-01-18
    
End Sub

Private Sub opt卡类别_Click(Index As Integer)
    gintType = Index + 1
    txt卡号.Enabled = (Index <> 1)
    cmdChangePassword.Enabled = (Index <> 2)
    cmd读卡.Visible = (Index = 1)
    cmd启动.Visible = (Index <> 1)
    Select Case Index
    Case 0
        lbl卡号.Caption = "卡号"
    Case 1
        lbl卡号.Caption = "IC卡号"
    Case 2
        lbl卡号.Caption = "身份证号"
    End Select
    If Index <> 1 Then txt卡号.SetFocus
End Sub

Private Sub txt卡号_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txt密码_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub txt密码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str封锁信息 As String
    Dim str病种 As String
    Dim rs病种 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txt卡号.Text) = "" Then
        MsgBox IIf(opt卡类别(2).Value = False, "请刷卡！", "请输入身份证号！"), vbInformation, gstrSysName
        txt卡号.SetFocus
        Exit Sub
    End If
    
    '2005.11.22,int门诊住院标志,住院强制选择保险类别
    If (int门诊住院标志 = 1 And cbo保险类别.ListIndex = 0) Then
       MsgBox "请选择保险类别！", vbInformation, gstrSysName
       cbo保险类别.SetFocus
       Exit Sub
    End If
    If opt卡类别(2).Value Then
        gstrIDNO = txt卡号.Text     '身份证号
        txt卡号.Text = ""
    End If
    If Me.cbo保险类别.Text = "工伤保险" Then
        '如果是工伤，则先获取工伤认定信息函数再读卡
        If InitXML = False Then Exit Sub
        Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
        Call InsertChild(mdomInput.documentElement, "CARDDATA", txt卡号.Text)            ' 磁卡数据
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txt密码.Text)            ' 密码
        Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                     ' 社会保障号
        Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
        Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
        Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo支付类别.ItemData(Me.cbo支付类别.ListIndex))            ' 支付类别
        If CommServer("GETGSINFO") = False Then Exit Sub
        mstr工伤认定编号 = frm工伤认定编号选择.ShowME
    End If
    
    '特殊门诊或住院
    'Modified By 朱玉宝 2003-12-03 地区： 原因：入院时取消病种选择，改为在虚拟结算时，如果没有病种，必需选择
    mstr病种 = ""
    mlng病种ID = 0
    If (Me.cbo支付类别.Text = "特殊门诊") Then
        gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                " From 保险病种 A where A.险类=" & TYPE_贵阳市
        
        Set rs病种 = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
        If Not rs病种 Is Nothing Then
            mlng病种ID = rs病种("ID")
            mstr病种 = rs病种!编码
        End If
        If mlng病种ID = 0 Then
            MsgBox "必须选择特殊病后才允许进行身份识别！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If InitXML = False Then Exit Sub
    '必须先修改密码
    If Trim(mstr新密码) <> "" Then
        If 更改密码_贵阳市(txt卡号.Text, mstr密码, mstr新密码) = False Then Exit Sub
        mstr密码 = mstr新密码
        mstr新密码 = ""
        txt密码.Text = mstr密码
    End If

    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt卡号.Text)            ' 磁卡数据
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt密码.Text)            ' 密码
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo支付类别.ItemData(Me.cbo支付类别.ListIndex))            ' 支付类别
    If Me.cbo支付类别.Text = "特殊门诊" Then Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", mstr病种)            ' 特殊病
    
    '2005.11.22,int门诊住院标志,医保返回
    If int门诊住院标志 = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex)
    End If
    
     '20110812周玉强增加判断是否有其他工伤信息
    gstr工伤信息 = mstr工伤认定编号
    If InStr(mstr工伤认定编号, "|") > 1 Then
     mstr工伤认定编号 = Mid(mstr工伤认定编号, 1, InStr(mstr工伤认定编号, "|") - 1)
    End If
    'end
    
    Call InsertChild(mdomInput.documentElement, "GSRDBH", mstr工伤认定编号)
    Call InsertChild(mdomInput.documentElement, "STARTDATE", mstr就诊时间)           ' 开始时间
    '调用接口
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    '取得返回值
    '基本信息
    txt卡号.Tag = txt卡号.Text                    '保存卡内数据，以便更新密码时使用
    'txt卡号.Text = GetElemnetValue("CARDID")
    txt医保号.Text = GetElemnetValue("PERSONCODE")
    txt分中心编号.Text = GetElemnetValue("CENTERCODE")
    txt医疗照顾人群.Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "否", "是")
    
    '2005.11.22,int门诊住院标志,住院必须自行选择保险类别,默认置空
    If int门诊住院标志 = 0 Then
        cbo保险类别.ListIndex = GetElemnetValue("INSURETYPE") - 1
    Else
'        cbo保险类别.ListIndex = 0
    End If
    
    '人员类别    11：在职；21：退休；32：省属离休；34：市属离休；41：普通居民；42：低保对象；43：三无人员；44：低收入家庭；45：重度残疾；
    txt人员类别.Text = GetElemnetValue("PERSONTYPE")
    txt人员类别.Text = Switch(txt人员类别.Text = "11", "在职", txt人员类别.Text = "21", "退休", _
                      txt人员类别.Text = "32", "省属离休", txt人员类别.Text = "34", "市属离休", _
                      txt人员类别.Text = "41", "普通居民", txt人员类别.Text = "42", "低保对象", _
                      txt人员类别.Text = "43", "三无人员", txt人员类别.Text = "44", "低收入家庭", _
                      txt人员类别.Text = "45", "重度残疾", True, "其他")
    txt姓名.Text = GetElemnetValue("PERSONNAME")
    txt性别.Text = GetElemnetValue("SEX")
    txt性别.Text = Switch(txt性别.Text = "1", "男", txt性别.Text = "2", "女", txt性别.Text = "9", "其它", True, txt性别.Text)
    txt身份证号.Text = GetElemnetValue("PID")
    txt出生日期.Text = GetElemnetValue("BIRTHDAY")
    txt单位编码.Text = GetElemnetValue("DEPTCODE")
    txt单位名称.Text = GetElemnetValue("DEPTNAME")
    txt帐户余额.Text = GetElemnetValue("ACCTBALANCE")
    '累计信息
    txt住院次数.Text = GetElemnetValue("HOSPTIMES")
    txt起付线.Text = GetElemnetValue("STARTFEE")
    txt已支付起付线.Text = GetElemnetValue("STARTFEEPAID")
    txt基本统筹限额.Text = GetElemnetValue("FUND1LMT")
    txt统筹支付累计.Text = GetElemnetValue("FUND1PAID")
    txt大额统筹限额.Text = GetElemnetValue("FUND2LMT")
    txt大额支付累计.Text = GetElemnetValue("FUND2PAID")
    txt普通门诊医疗补助限额.Text = GetElemnetValue("FUND3LMT")
    txt普通门诊医疗补助累计.Text = GetElemnetValue("FUND3PAID")
    txt普通门诊医疗补助起付标准.Text = GetElemnetValue("STARTFEE2STD")
    txt普通门诊医疗补助起付线.Text = GetElemnetValue("STARTFEE2")
    txt普通门诊医疗补助结转可使用.Text = GetElemnetValue("FUND75BALANCE")
    txt备注.Text = GetElemnetValue("NOTE")
    txt封锁信息.Text = GetElemnetValue("LOCKINFO")
    
    '20110812增加保存
    Txt工伤信息.Text = gstr工伤信息
     gstrSQL = " Select a.处方编号 From zlgyyb.医保病人档案 a,医保病人关联表 b" & _
   " Where a.医保号=b.医保号 and  b.标志=1 and b.险类=[1] And a.医保号=[2]"
       Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取处方编号", TYPE_贵阳市, CStr(Trim(txt医保号.Text)))
        If rsTemp.RecordCount = 1 Then
        If rsTemp!处方编号 <> "" And txt处方本编号.Enabled = True Then
        txt处方本编号.Text = rsTemp!处方编号
        End If
       End If
    'end
     cmdOK.Enabled = True
    If txt处方本编号.Enabled = True Then txt处方本编号.SetFocus
  
    If gblnLED Then
        zl9LedVoice.Speak "#26 " & Val(txt帐户余额.Text)
    End If
End Sub
