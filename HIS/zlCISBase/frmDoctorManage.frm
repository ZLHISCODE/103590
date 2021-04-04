VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDoctorManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医生授权管理"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14340
   Icon            =   "frmDoctorManage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14340
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1350
      Index           =   3
      Left            =   5040
      ScaleHeight     =   1320
      ScaleWidth      =   2085
      TabIndex        =   104
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1185
         Index           =   3
         Left            =   80
         ScaleHeight     =   1185
         ScaleWidth      =   1920
         TabIndex        =   105
         Top             =   80
         Width           =   1920
         Begin VB.CheckBox ChkFdKSSzy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "特殊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   1080
            TabIndex        =   115
            Top             =   720
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSzy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "限制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   1080
            TabIndex        =   114
            Top             =   480
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSzy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "非限制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   1080
            TabIndex        =   113
            Top             =   240
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSzy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   1080
            TabIndex        =   112
            Top             =   960
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "特殊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   111
            Top             =   720
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "限制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   108
            Top             =   480
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "非限制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   107
            Top             =   240
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdKSSmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   106
            Top             =   960
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.Line Line5 
            Index           =   1
            X1              =   960
            X2              =   960
            Y1              =   0
            Y2              =   1440
         End
         Begin VB.Label lblPic 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "门诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   110
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblPic 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   1320
            TabIndex        =   109
            Top             =   0
            Width           =   390
         End
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   2
      Left            =   7200
      ScaleHeight     =   1305
      ScaleWidth      =   1005
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1185
         Index           =   2
         Left            =   80
         ScaleHeight     =   1185
         ScaleWidth      =   855
         TabIndex        =   98
         Top             =   80
         Width           =   850
         Begin VB.CheckBox ChkFdSS 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   103
            Top             =   960
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.CheckBox ChkFdSS 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "四级"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   102
            Top             =   720
            Value           =   1  'Checked
            Width           =   720
         End
         Begin VB.CheckBox ChkFdSS 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "三级"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   101
            Top             =   480
            Value           =   1  'Checked
            Width           =   720
         End
         Begin VB.CheckBox ChkFdSS 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "二级"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   100
            Top             =   240
            Value           =   1  'Checked
            Width           =   720
         End
         Begin VB.CheckBox ChkFdSS 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "一级"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Value           =   1  'Checked
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   8280
      ScaleHeight     =   825
      ScaleWidth      =   2085
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Index           =   1
         Left            =   80
         ScaleHeight     =   705
         ScaleWidth      =   1920
         TabIndex        =   90
         Top             =   80
         Width           =   1920
         Begin VB.CheckBox ChkFdZy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   1080
            TabIndex        =   96
            Top             =   480
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "有权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   93
            Top             =   240
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdmz 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   92
            Top             =   480
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox ChkFdZy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "有权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   1080
            TabIndex        =   91
            Top             =   240
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.Label lblPic 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "住院"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   95
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblPic 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "门诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   94
            Top             =   0
            Width           =   390
         End
         Begin VB.Line Line5 
            Index           =   0
            X1              =   960
            X2              =   960
            Y1              =   0
            Y2              =   1440
         End
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   3840
      ScaleHeight     =   570
      ScaleWidth      =   1215
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   50
         ScaleHeight     =   450
         ScaleWidth      =   1035
         TabIndex        =   87
         Top             =   50
         Width           =   1030
         Begin VB.CheckBox ChkFd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "有处方权"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   89
            Top             =   0
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox ChkFd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无处方权"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   88
            Top             =   240
            Value           =   1  'Checked
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   14265
      TabIndex        =   5
      Top             =   1320
      Width           =   14295
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   0
         Left            =   3240
         Picture         =   "frmDoctorManage.frx":6852
         Top             =   430
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   6
         Left            =   13560
         Picture         =   "frmDoctorManage.frx":6BDC
         Top             =   430
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   5
         Left            =   12860
         Picture         =   "frmDoctorManage.frx":6F66
         Top             =   430
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   4
         Left            =   11880
         Picture         =   "frmDoctorManage.frx":72F0
         Top             =   435
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   3
         Left            =   11160
         Picture         =   "frmDoctorManage.frx":767A
         Top             =   435
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   2
         Left            =   9850
         Picture         =   "frmDoctorManage.frx":7A04
         Top             =   430
         Width           =   240
      End
      Begin VB.Image imgSentence 
         Height          =   240
         Index           =   1
         Left            =   6720
         Picture         =   "frmDoctorManage.frx":7D8E
         Top             =   430
         Width           =   240
      End
      Begin VB.Image imgDoctor 
         Height          =   240
         Left            =   650
         Picture         =   "frmDoctorManage.frx":8118
         Top             =   220
         Width           =   240
      End
      Begin VB.Image imgTs 
         Height          =   240
         Left            =   11520
         Picture         =   "frmDoctorManage.frx":E96A
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgSS 
         Height          =   240
         Left            =   7820
         Picture         =   "frmDoctorManage.frx":151BC
         Top             =   50
         Width           =   240
      End
      Begin VB.Image img处方权 
         Height          =   240
         Left            =   2780
         Picture         =   "frmDoctorManage.frx":1BA0E
         Top             =   225
         Width           =   240
      End
      Begin VB.Image imgKSS 
         Height          =   240
         Left            =   4680
         Picture         =   "frmDoctorManage.frx":22260
         Top             =   50
         Width           =   240
      End
      Begin VB.Line lineHead 
         Index           =   7
         X1              =   3840
         X2              =   14280
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line lineHead 
         Index           =   4
         X1              =   10200
         X2              =   10200
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line lineHead 
         Index           =   2
         X1              =   6960
         X2              =   6960
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line lineHead 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   3840
         X2              =   3840
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line lineHead 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   2760
         X2              =   2760
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "贵重"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   29
         Left            =   13160
         TabIndex        =   21
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "精神I类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   28
         Left            =   12180
         TabIndex        =   20
         Top             =   450
         Width           =   675
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "麻醉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   27
         Left            =   11475
         TabIndex        =   19
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "毒类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   26
         Left            =   10755
         TabIndex        =   18
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "四级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   25
         Left            =   9440
         TabIndex        =   17
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "三级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   24
         Left            =   8720
         TabIndex        =   16
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "二级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   23
         Left            =   8000
         TabIndex        =   15
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "一级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   22
         Left            =   7300
         TabIndex        =   14
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "特殊"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   21
         Left            =   6310
         TabIndex        =   13
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "限制"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   20
         Left            =   5520
         TabIndex        =   12
         Top             =   450
         Width           =   390
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "非限制"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   4560
         TabIndex        =   11
         Top             =   450
         Width           =   585
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "特殊医嘱权限"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   11880
         TabIndex        =   10
         Top             =   75
         Width           =   1170
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "处方权"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "医生信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "手术等级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   8160
         TabIndex        =   7
         Top             =   70
         Width           =   780
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "抗菌药物权限"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5040
         TabIndex        =   6
         Top             =   70
         Width           =   1170
      End
   End
   Begin MSComctlLib.ImageList imgUpDown 
      Left            =   10920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":28AB2
            Key             =   "Up"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":28E4C
            Key             =   "Down"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":291E6
            Key             =   "Used"
            Object.Tag             =   "Used"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6585
      ScaleWidth      =   14265
      TabIndex        =   22
      Top             =   2040
      Width           =   14295
      Begin VB.PictureBox picUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Left            =   0
         ScaleHeight     =   6495
         ScaleWidth      =   13995
         TabIndex        =   24
         Top             =   0
         Width           =   13995
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   5
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   75
            Top             =   5520
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   1170
               TabIndex        =   76
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   5
                  ItemData        =   "frmDoctorManage.frx":29580
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":29582
                  TabIndex        =   77
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Line LineBack 
               Index           =   5
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":29584
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":29F6E
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":2A958
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":2B342
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":2BD2C
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":2C716
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":2D100
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":2DAEA
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":2E4D4
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":2EEBE
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":2F8A8
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":30292
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":30C7C
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":31666
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":32050
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":32A3A
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":33424
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":33E0E
               Top             =   0
               Width           =   420
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   5
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":347F8
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   1200
               TabIndex        =   84
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   2400
               TabIndex        =   83
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   1200
               TabIndex        =   82
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   5
               Left            =   20
               Picture         =   "frmDoctorManage.frx":351E2
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   5
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   5
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   5
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   5
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   4080
               TabIndex        =   81
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   4080
               TabIndex        =   80
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   5
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   10320
               TabIndex        =   79
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   10320
               TabIndex        =   78
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   4
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   65
            Top             =   4440
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   1170
               TabIndex        =   66
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   4
                  ItemData        =   "frmDoctorManage.frx":35CAD
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":35CAF
                  TabIndex        =   67
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Line LineBack 
               Index           =   4
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":35CB1
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":3669B
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":37085
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":37A6F
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":38459
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":38E43
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":3982D
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":3A217
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":3AC01
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":3B5EB
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":3BFD5
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":3C9BF
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":3D3A9
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":3DD93
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":3E77D
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":3F167
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":3FB51
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":4053B
               Top             =   0
               Width           =   420
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   4
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":40F25
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   1200
               TabIndex        =   74
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   2400
               TabIndex        =   73
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   1200
               TabIndex        =   72
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   4
               Left            =   20
               Picture         =   "frmDoctorManage.frx":4190F
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   4
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   4
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   4
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   4
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   4080
               TabIndex        =   71
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   4080
               TabIndex        =   70
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   4
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   10320
               TabIndex        =   69
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   10320
               TabIndex        =   68
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   3
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   55
            Top             =   3360
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   1170
               TabIndex        =   56
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   3
                  ItemData        =   "frmDoctorManage.frx":423DA
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":423DC
                  TabIndex        =   57
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Line LineBack 
               Index           =   3
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":423DE
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":42DC8
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":437B2
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":4419C
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":44B86
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":45570
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":45F5A
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":46944
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":4732E
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":47D18
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":48702
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":490EC
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":49AD6
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":4A4C0
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":4AEAA
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":4B894
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":4C27E
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":4CC68
               Top             =   0
               Width           =   420
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   3
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":4D652
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   1200
               TabIndex        =   64
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   2400
               TabIndex        =   63
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   1200
               TabIndex        =   62
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   3
               Left            =   20
               Picture         =   "frmDoctorManage.frx":4E03C
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   4080
               TabIndex        =   61
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   4080
               TabIndex        =   60
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   10320
               TabIndex        =   59
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   10320
               TabIndex        =   58
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   2
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   45
            Top             =   2280
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   1170
               TabIndex        =   46
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   2
                  ItemData        =   "frmDoctorManage.frx":4EB07
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":4EB09
                  TabIndex        =   47
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   10320
               TabIndex        =   54
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   10320
               TabIndex        =   53
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   2
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   4080
               TabIndex        =   52
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   4080
               TabIndex        =   51
               Top             =   120
               Width           =   540
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   2
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   2
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   2
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   2
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   2
               Left            =   20
               Picture         =   "frmDoctorManage.frx":4EB0B
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   1200
               TabIndex        =   50
               Top             =   120
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   2400
               TabIndex        =   49
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   1200
               TabIndex        =   48
               Top             =   720
               Width           =   360
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":4F5D6
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":4FFC0
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":509AA
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":51394
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":51D7E
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":52768
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":53152
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":53B3C
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":54526
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":54F10
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":558FA
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":562E4
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":56CCE
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":576B8
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":580A2
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":58A8C
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":59476
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":59E60
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   2
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":5A84A
               Top             =   480
               Width           =   420
            End
            Begin VB.Line LineBack 
               Index           =   2
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   1
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   35
            Top             =   1200
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   1170
               TabIndex        =   36
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   1
                  ItemData        =   "frmDoctorManage.frx":5B234
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":5B236
                  TabIndex        =   37
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   10320
               TabIndex        =   44
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   10320
               TabIndex        =   43
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   4080
               TabIndex        =   42
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   4080
               TabIndex        =   41
               Top             =   120
               Width           =   540
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   1
               Left            =   20
               Picture         =   "frmDoctorManage.frx":5B238
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   40
               Top             =   120
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   2400
               TabIndex        =   39
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   38
               Top             =   720
               Width           =   360
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":5BD03
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":5C6ED
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":5D0D7
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":5DAC1
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":5E4AB
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":5EE95
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":5F87F
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":60269
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":60C53
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":6163D
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":62027
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":62A11
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":633FB
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":63DE5
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":647CF
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":651B9
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":65BA3
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":6658D
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   1
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":66F77
               Top             =   480
               Width           =   420
            End
            Begin VB.Line LineBack 
               Index           =   1
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   0
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   13935
            TabIndex        =   25
            Top             =   120
            Width           =   13935
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   1170
               TabIndex        =   33
               Top             =   380
               Width           =   1450
               Begin VB.ComboBox cboLevel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   300
                  Index           =   0
                  ItemData        =   "frmDoctorManage.frx":67961
                  Left            =   -23
                  List            =   "frmDoctorManage.frx":67963
                  TabIndex        =   34
                  Text            =   "cboLevel"
                  Top             =   -23
                  Width           =   1500
               End
            End
            Begin VB.Line LineBack 
               Index           =   0
               X1              =   1200
               X2              =   2590
               Y1              =   620
               Y2              =   620
            End
            Begin VB.Image PicZYGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":67965
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZGz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   13200
               Picture         =   "frmDoctorManage.frx":6834F
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":68D39
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZJs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   12480
               Picture         =   "frmDoctorManage.frx":69723
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":6A10D
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZMz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   11760
               Picture         =   "frmDoctorManage.frx":6AAF7
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":6B4E1
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZDl 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   11040
               Picture         =   "frmDoctorManage.frx":6BECB
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicSS4 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   9480
               Picture         =   "frmDoctorManage.frx":6C8B5
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS3 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   8760
               Picture         =   "frmDoctorManage.frx":6D29F
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS2 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   8040
               Picture         =   "frmDoctorManage.frx":6DC89
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicSS1 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   7320
               Picture         =   "frmDoctorManage.frx":6E673
               Top             =   240
               Width           =   420
            End
            Begin VB.Image PicZYTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":6F05D
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZTs 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   6360
               Picture         =   "frmDoctorManage.frx":6FA47
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":70431
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZXz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   5520
               Picture         =   "frmDoctorManage.frx":70E1B
               Top             =   0
               Width           =   420
            End
            Begin VB.Image PicZYFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":71805
               Top             =   480
               Width           =   420
            End
            Begin VB.Image PicMZFxz 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   4680
               Picture         =   "frmDoctorManage.frx":721EF
               Top             =   0
               Width           =   420
            End
            Begin VB.Image Pic处方权 
               Appearance      =   0  'Flat
               Height          =   420
               Index           =   0
               Left            =   3120
               Picture         =   "frmDoctorManage.frx":72BD9
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblZw 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "职务"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   1200
               TabIndex        =   32
               Top             =   720
               Width           =   360
            End
            Begin VB.Label lblSex 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   2400
               TabIndex        =   31
               Top             =   120
               Width           =   180
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   1200
               TabIndex        =   30
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgUser 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   945
               Index           =   0
               Left            =   20
               Picture         =   "frmDoctorManage.frx":735C3
               Stretch         =   -1  'True
               Top             =   0
               Width           =   945
            End
            Begin VB.Line LineFg 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   0
               X2              =   19800
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   10200
               X2              =   10200
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblKSSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   4080
               TabIndex        =   29
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lblKSSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   4080
               TabIndex        =   28
               Top             =   600
               Width           =   540
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   2760
               X2              =   2760
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Label lblTSZy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "住院："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   10320
               TabIndex        =   27
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblTSMz 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "门诊："
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   10320
               TabIndex        =   26
               Top             =   120
               Width           =   540
            End
         End
      End
      Begin VB.VScrollBar vscBar 
         Height          =   6615
         LargeChange     =   6
         Left            =   14040
         Max             =   1
         SmallChange     =   6
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4005
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1000
      Width           =   1905
   End
   Begin VB.ComboBox cboDept 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1000
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8685
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   635
      SimpleText      =   $"frmDoctorManage.frx":7408E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDoctorManage.frx":740D5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2

            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList img16 
      Left            =   11640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":74969
            Key             =   "Black1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":75363
            Key             =   "Black2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":75D5D
            Key             =   "Blue1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":76757
            Key             =   "Blue2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":77151
            Key             =   "Green1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":77B4B
            Key             =   "Green2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":78545
            Key             =   "None"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":78F3F
            Key             =   "Red1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorManage.frx":79939
            Key             =   "Red2"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgUserTmp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   13320
      Picture         =   "frmDoctorManage.frx":7A333
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   945
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   12480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&F)"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1050
      Width           =   705
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "科室(&D)"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1050
      Width           =   705
   End
End
Attribute VB_Name = "frmDoctorManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String               '当前用户拥有权限的字符串
Private mobjBar As CommandBar             'Comandbar对象
Private mstrUserDept As String            '当前医生所在科室
Private mrsTmp As ADODB.Recordset     '当前医生信息缓存数据集
Private mlngFindNum As Long               '是否启用查找功能
Private mlngFilter As Long               '是否启用过滤功能
Private mlngIndex As Long                 '光标跟随Pic的Index
Private mblnPic As Boolean                '是否加载照片
Private mblnMove As Boolean               '是否关闭过滤图片pic
Private mstrFdsql As String               '过滤条件的sql
Private Const conMenu_View_Photo = 999    '加载医生照片的按钮id
Private mcolCtl As New Collection         '窗体控件对象集合



Private Const conCtl = ",picInfo,lblName,lblSex,lblZw,cboLevel,Pic处方权,PicMZFxz,PicMZXz,PicMZTs,PicZYFxz,PicZYXz," & _
                        "PicZYTs,PicSS1,PicSS2,PicSS3,PicSS4,PicMZDl,PicMZMz,PicMZJs,PicMZGz,PicZYDl,PicZYMz,PicZYJs,PicZYGz"

Private Enum mEnumPrivsType        '医生权限类型：=1 处方权,=2 门诊抗菌药物权限,=3 住院抗菌药物权限,=4 手术等级权限,=5 门诊特殊医嘱权限,=6 住院特殊医嘱权限,=7 医生等级
    Priv处方权 = 1
    Priv门诊抗菌 = 2
    Priv住院抗菌 = 3
    Priv手术等级 = 4
    Priv门诊特殊 = 5
    Priv住院特殊 = 6
    Priv医生等级 = 7
End Enum

Private Enum imgType
    imgType处方权 = 0
    imgType抗菌药物 = 1
    imgType手术等级 = 2
    imgType毒类 = 3
    imgType麻醉 = 4
    imgType精神 = 5
    imgType贵重 = 6
End Enum

Private Enum FilterType
    Filter处方权 = 0
    Filter特殊医嘱 = 1
    Filter手术等级 = 2
    Filter抗菌药物 = 3
End Enum

Private Enum TsType
    Ts毒类 = 1
    Ts麻醉 = 2
    Ts精神 = 3
    Ts贵重 = 4
End Enum

Private Sub cboDept_Click()
    '加载当前选择的部门
    Dim i As Long
On Error GoTo errH
    mlngFindNum = 0
    Set mrsTmp = LoadPrss
    For i = 0 To 5
        If Not mrsTmp.EOF Then
            LoadData (i)
            mrsTmp.MoveNext
            picInfo.Item(i).Visible = True
        Else
            picInfo.Item(i).Visible = False
        End If
    Next
    vscBar.Value = 0
    picUser.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboLevel_Click(Index As Integer)
    '更新医生等级
    Dim strSql As String
    Dim intType As Integer
    Dim strCheck As String
On Error GoTo errH
    strCheck = cboLevel(Index).Text
    strSql = "Zl_医生权限_Update(" & Val(picInfo(Index).Tag) & "," & Priv医生等级 & ",'" & strCheck & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Call CacheData(Val(picInfo(Index).Tag), Priv医生等级, strCheck)
    picUser.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboLevel_KeyPress(Index As Integer, KeyAscii As Integer)
    '不允许手动输入医生等级
    KeyAscii = 0
End Sub

Private Sub LoadCbo()
    '加载医生等级下拉框数据
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim i As Integer
    
On Error GoTo errH
    strSql = "select 名称 from 专业技术职务 where decode( substr(编码,1,2),编码,null ,substr(编码,1,2))=23"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    For i = 0 To 5
        rsTmp.MoveFirst
        cboLevel.Item(i).AddItem ""
        Do While Not rsTmp.EOF
            cboLevel.Item(i).AddItem rsTmp!名称 & ""
            rsTmp.MoveNext
        Loop
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CloseFilter(picindex As Integer)
    Dim i As Integer
    '关闭过滤页面
    On Error GoTo errH
    If mblnMove Then
        picFilter(picindex).Visible = False
        picUser.SetFocus
        Call setUseIco(picFilter(picindex).Tag)
        mblnMove = False
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'组织过滤SQL
Private Function FilterSQL() As Boolean
    Dim i As Integer
    Dim strSql As String
    Dim strFilter As String
    Dim j As Integer
    Dim strTmp As String
    Dim strFind As String
    On Error GoTo errH
    
    '是否启用查找
    strFind = Replace(Replace(Replace(txtFind.Text, "'", ""), "*", ""), "%", "")
    If strFind = "" Then mlngFindNum = 0
    If mlngFindNum = 1 Then
        If zlCommFun.IsCharChinese(strFind) Then
            strSql = strSql & " And 姓名 like '*" & strFind & "*'"
        ElseIf IsNumeric(strFind) Then
            strSql = strSql & " And 编号 like '*" & strFind & "*'"
        Else
            strSql = strSql & " And 简码 like '*" & strFind & "*'"
        End If
    End If
    

    mlngFilter = 0
    For i = 0 To imgSentence.Count - 1
        If imgSentence(i).Tag <> "" Then
            '开启过滤
            mlngFilter = 1
            strTmp = ""
            Select Case i
                '处方权
                Case imgType处方权
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And 处方权标志 <> 1"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And 处方权标志 <> 0"
                '门诊住院抗菌药物
                Case imgType抗菌药物
                  '门诊
                  For j = 0 To 3
                    strTmp = decode(j, 0, "非限制使用", 1, "限制使用", 2, "特殊使用", 3, "null")
                    If Split(imgSentence(i).Tag, "|")(j) = "0" Then strSql = strSql & " And 门诊抗菌药物权限 <> '" & strTmp & "'"
                  Next
                  '住院
                  For j = 0 To 3
                    strTmp = decode(j, 0, "非限制使用", 1, "限制使用", 2, "特殊使用", 3, "null")
                    If Split(imgSentence(i).Tag, "|")(j + 4) = "0" Then strSql = strSql & " And 住院抗菌药物权限 <> '" & strTmp & "'"
                  Next
                '手术等级
                Case imgType手术等级
                  For j = 0 To 4
                    strTmp = decode(j, 0, "一级", 1, "二级", 2, "三级", 3, "四级", 4, "null")
                    If Split(imgSentence(i).Tag, "|")(j) = "0" Then strSql = strSql & " And 手术等级 <> '" & strTmp & "'"
                  Next
                '毒类
                Case imgType毒类
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And 门诊毒类 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And 门诊毒类 <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And 住院毒类 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And 住院毒类 <> '0'"
                '麻醉
                Case imgType麻醉
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And 门诊麻醉 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And 门诊麻醉 <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And 住院麻醉 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And 住院麻醉 <> '0'"
                 '精神
                Case imgType精神
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And 门诊精神 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And 门诊精神 <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And 住院精神 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And 住院精神 <> '0'"
                  '贵重
                Case imgType贵重
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And 门诊贵重 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And 门诊贵重 <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And 住院贵重 <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And 住院贵重 <> '0'"
            End Select
        End If
    Next
    If strSql <> "" Then strSql = Mid(strSql, 5)
    If mstrFdsql = strSql Then
       FilterSQL = False
    Else
       mstrFdsql = strSql
       FilterSQL = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CheckBtn(objCheck As Object, Index As Integer)
'过滤条件不允许全部取消
    Dim i As Integer
    Dim blnChk As Boolean

    For i = 0 To objCheck.Count - 1
        If objCheck(i).Value = 1 Then
            blnChk = True
        End If
    Next
    If Not blnChk Then objCheck(Index).Value = 1
End Sub

Private Sub Form_Activate()
    '用于支持滚轮
    picBack.SetFocus
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FmgFlexScroll
End Sub

Private Sub Form_Deactivate()
    '用于支持滚轮
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    On Error GoTo errH
    mlngIndex = 999
    mlngFindNum = 0
    mstrPrivs = gstrPrivs
    mstrFdsql = ""
    mstrUserDept = GetUser科室IDs(True)
    Call LoadDept
    Call LoadCbo
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    Call RestoreWinState(Me, App.ProductName)
    '初始化权限控件集合
    Set mcolCtl = GetControls
    '初始化界面信息
    Set mrsTmp = LoadPrss
    For i = 0 To 5
        If Not mrsTmp.EOF Then
            DoEvents
            LoadData (i)
            mrsTmp.MoveNext
            picInfo.Item(i).Visible = True
        Else
            picInfo.Item(i).Visible = False
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "手术授权管理(&N)")
            objControl.IconId = 9002
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Photo, "加载医生照片(&P)")
        objControl.Checked = False
        mblnPic = objControl.Checked
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "手术授权管理")
        objControl.IconId = 9002
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_Photo
            If Control.Checked = True Then
                Control.Checked = False
                mblnPic = Control.Checked
            Else
                Control.Checked = True
                mblnPic = Control.Checked
            End If
            Call vscBar_Change
        Case conMenu_Tool_OPSEmpower '手术授权管理
            '权限判断和处理
            gstrPrivs = GetPrivFunc(glngSys, 1080)
            If gstrPrivs = "" Then
                MsgBox "你没有手术授权管理权限，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Sub
            End If
            frmOPSEmpower.Show , Me
        Case conMenu_View_Refresh '刷新
            Call vscBar_Change
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Function GetControls() As Collection
    '获得ConCtl字符串的所有控件的对象
    Dim objTmp As Object
    Dim colCtl As New Collection
    For Each objTmp In Me.Controls
        If InStr(conCtl, objTmp.Name) > 0 Then
            colCtl.Add objTmp, objTmp.Name & objTmp.Index
        End If
    Next
    Set GetControls = colCtl
End Function

Private Sub LoadData(ByVal intIndex As Integer)
    '加载界面信息
    Dim str处方权 As String
    Dim strKSSMz As String
    Dim strKSSZy As String
    Dim strOPS As String
    Dim strTsMz As String
    Dim strTsZy As String
    Dim strTempFile As String
    Dim i As Integer
    Dim objTmp As Object
    
    On Error GoTo errH
    '清除数据
    clearInfo intIndex
    
    str处方权 = Val(mrsTmp!处方权标志 & "")
    strKSSMz = mrsTmp!门诊抗菌药物权限 & ""
    strKSSZy = mrsTmp!住院抗菌药物权限 & ""
    strOPS = mrsTmp!手术等级 & ""
    strTsMz = IIf(mrsTmp!门诊特殊医嘱权限 & "" <> "", mrsTmp!门诊特殊医嘱权限 & "", "0000")
    strTsZy = IIf(mrsTmp!住院特殊医嘱权限 & "" <> "", mrsTmp!住院特殊医嘱权限 & "", "0000")
    
    '加载医生信息
    cboLevel(intIndex).Text = mrsTmp!专业技术职务 & ""
    lblName.Item(intIndex).Caption = mrsTmp!姓名 & ""
    lblSex.Item(intIndex).Caption = mrsTmp!性别 & ""
    lblZw.Item(intIndex).Caption = mrsTmp!管理职务 & ""
    picInfo.Item(intIndex).Tag = mrsTmp!ID & ""
    
    '加载处方权
    If str处方权 = 1 Then
        Pic处方权(intIndex).Picture = img16.ListImages.Item("Black2").Picture: Pic处方权(intIndex).Tag = "True"
    End If
    
    For i = 1 To 4
        '加载住院特殊医嘱权限
        Set objTmp = mcolCtl(decode(i, Ts毒类, "PicZYDl", Ts麻醉, "PicZYMz", Ts精神, "PicZYJs", Ts贵重, "PicZYGz") & intIndex)
        If Val(Mid(strTsZy, i, 1)) = 1 Then objTmp.Picture = img16.ListImages.Item("Green2").Picture: objTmp.Tag = "True"
        '加载门诊特殊医嘱权限
        Set objTmp = mcolCtl(decode(i, Ts毒类, "PicMZDl", Ts麻醉, "PicMZMz", Ts精神, "PicMZJs", Ts贵重, "PicMZGz") & intIndex)
        If Val(Mid(strTsMz, i, 1)) = 1 Then objTmp.Picture = img16.ListImages.Item("Green2").Picture: objTmp.Tag = "True"
    Next
    
    '加载手术等级
    If strOPS <> "null" Then
        Set objTmp = mcolCtl(decode(strOPS, "一级", "PicSS1", "二级", "PicSS2", "三级", "PicSS3", "四级", "PicSS4") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Blue2").Picture: objTmp.Tag = "True"
    End If

    '加载门诊抗菌药物权限
    If strKSSMz <> "null" Then
        Set objTmp = mcolCtl(decode(strKSSMz, "非限制使用", "PicMZFxz", "限制使用", "PicMZXz", "特殊使用", "PicMZTs") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Red2").Picture: objTmp.Tag = "True"
    End If

    '加载住院抗菌药物权限
    If strKSSZy <> "null" Then
        Set objTmp = mcolCtl(decode(strKSSZy, "非限制使用", "PicZYFxz", "限制使用", "PicZYXz", "特殊使用", "PicZYTs") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Red2").Picture: objTmp.Tag = "True"
    End If
    
    '显示照片
    If mblnPic = True Then
        If mrsTmp!照片大小 & "" <> "" And Val(mrsTmp!照片大小 & "") < 10000 Then
            strTempFile = sys.Readlob(100, 16, Val(picInfo.Item(intIndex).Tag))
            If strTempFile <> "" Then
                imgUser(intIndex).Picture = LoadPicture(strTempFile)
                  '删除该临时文件
                Kill strTempFile
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
       
Private Sub clearInfo(ByVal intIndex As Long)
    '控件初始化,清空界面信息
    Dim objTmp As Object
    Dim i As Integer
On Error GoTo errH
    lblName.Item(intIndex).Caption = ""
    lblSex.Item(intIndex).Caption = ""
    lblZw.Item(intIndex).Caption = ""
    picInfo.Item(intIndex).Tag = ""
    imgUser(intIndex).Picture = imgUserTmp.Picture
    cboLevel(intIndex).Text = ""
    '清空图片勾选数据
    For i = 1 To mcolCtl.Count
        Set objTmp = mcolCtl(i)
        If TypeName(objTmp) = "Image" Then
            If objTmp.Index = intIndex Then
                objTmp.Picture = img16.ListImages.Item("None").Picture
                objTmp.Tag = ""
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDept()
'加载操作员所属科室
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select B.ID,B.编码,B.名称 " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", ",A.缺省") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", "部门人员 A, ") & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", " And a.部门id = B.Id And A.人员ID = [1] ") & vbNewLine & _
            "  And C.工作性质 = '临床' And C.服务对象 <> 0  And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"
    On Error GoTo errH
    cboDept.Clear
    '所有部门
    If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then
        cboDept.AddItem "所有部门"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '所属缺省
        If InStr(";" & mstrPrivs & ";", ";所有部门;") = 0 Then
            If rsTmp!缺省 = 1 Then
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    '没有强制限制临床,可能医技科室用
    If blnNew Then
        strSql = "Select 1 as 类别,部门ID From 部门人员 Where 人员ID=[1] Union" & _
                " Select Distinct 2 as 类别,B.科室ID From 部门人员 A,病区科室对应 B" & _
                " Where A.部门ID=B.病区ID And A.人员ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", UserInfo.ID)
    End If
    If bln病区 = False Then
        rsTmp.Filter = "类别 = 1"
    Else
        rsTmp.Filter = ""
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser科室IDs & ",", "," & rsTmp!部门ID & ",") = 0 Then
            GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadPrss() As Recordset
'功能：加载可授权的用户
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long, y As Long
    Dim clsRs As New clsRecordset
    
    If cboDept.ListIndex = -1 Then Exit Function
    strSql = "Select a.Id, c.部门id, a.姓名, a.编号, a.性别, a.简码, b.名称 As 所属部门, b.编码 As 部门编码, a.专业技术职务,nvl(a.处方权标志,0) as 处方权标志 ,nvl(a.门诊特殊医嘱权限,'0000') as 门诊特殊医嘱权限,nvl(a.住院特殊医嘱权限,'0000') as 住院特殊医嘱权限,a.管理职务," & vbNewLine & _
                "       Decode(d.级别, 1, '非限制使用', 2, '限制使用', 3, '特殊使用', 'null') As 住院抗菌药物权限," & vbNewLine & _
                "       Decode(e.级别, 1, '非限制使用', 2, '限制使用', 3, '特殊使用', 'null') As 门诊抗菌药物权限, nvl(a.手术等级,'null') as 手术等级,dbms_lob.getlength(G.照片) as 照片大小," & vbNewLine & _
                "nvl(Substr(A.门诊特殊医嘱权限, 1, 1),'0') as 门诊毒类,nvl(Substr(A.门诊特殊医嘱权限, 2, 1),'0') as 门诊麻醉,nvl(Substr(A.门诊特殊医嘱权限, 3, 1),'0') as 门诊精神,nvl(Substr(A.门诊特殊医嘱权限, 4, 1),'0') as 门诊贵重," & vbNewLine & _
                "nvl(Substr(A.住院特殊医嘱权限, 1, 1),'0') as 住院毒类,nvl(Substr(A.住院特殊医嘱权限, 2, 1),'0') as 住院麻醉,nvl(Substr(A.住院特殊医嘱权限, 3, 1),'0') as 住院精神,nvl(Substr(A.住院特殊医嘱权限, 4, 1),'0') as 住院贵重" & vbNewLine & _
                "From 人员表 A, 部门表 B, 部门人员 C, 人员抗菌药物权限 D, 人员抗菌药物权限 E, 人员性质说明 F,人员照片 G" & vbNewLine & _
                "Where a.Id = c.人员id And c.部门id = b.Id And d.人员id(+) = a.Id And a.Id = f.人员id  And A.id=G.人员id(+) And (d.记录状态 = 1 Or d.记录状态 Is Null) And d.场合(+) = 1  And" & vbNewLine & _
                "      e.人员id(+) = a.Id And (e.记录状态 = 1 Or e.记录状态 Is Null) And e.场合(+) = 2 And" & vbNewLine & _
                "      b.Id In (Select ID From 部门表 Start With 上级id Is Null Connect By Prior ID = 上级id) And" & vbNewLine & _
                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  And f.人员性质 = '医生'"
    On Error GoTo errH
    '判断是否有所有部门权限
    If cboDept.ItemData(cboDept.ListIndex) = -1 Then
        If InStr(mstrPrivs, ";所有部门;") = 0 Then
            strSql = strSql & " And Instr([1],','|| B.ID || ',')>0 Order By a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "," & mstrUserDept & ",")
        Else
            strSql = strSql & " And C.缺省 = 1 Order By a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        End If
    Else
        strSql = strSql & " And c.部门id=[1]  Order By a.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(cboDept.ItemData(cboDept.ListIndex)))
    End If
    
    Set LoadPrss = clsRs.CopyNew(rsTmp)
    
    Call FilterSQL
    LoadPrss.Filter = mstrFdsql
    
    '设置状态栏数据
    If mlngFindNum = 1 Then
        stbThis.Panels(2).Text = "当前查找一共有" & LoadPrss.RecordCount & "名医生"
    Else
        stbThis.Panels(2).Text = "当前部门一共有" & LoadPrss.RecordCount & "名医生"
    End If
    If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(已启用过滤)"
    
    If LoadPrss.RecordCount - 6 < 0 Then
        vscBar.Max = 0
    Else
        vscBar.Max = LoadPrss.RecordCount - 6
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mcolCtl Is Nothing Then Set mcolCtl = Nothing
    Unload Me
End Sub


'打开或关闭过滤页签
Private Sub imgSentence_Click(Index As Integer)
    Dim i As Integer
    On Error GoTo errH
    If imgSentence(Index).Picture = imgUpDown.ListImages.Item("Up").Picture Then
        For i = 0 To picFilter.Count - 1
            If picFilter(i).Visible Then
                picFilter(i).Visible = False
                picUser.SetFocus
                mblnMove = False
            End If
        Next
        Call setUseIco(Index)
    Else
        For i = 0 To picFilter.Count - 1
            If picFilter(i).Visible Then
                picFilter(i).Visible = False
                picUser.SetFocus
                mblnMove = False
                Call setUseIco(picFilter(i).Tag)
            End If
        Next
        picUser.SetFocus
        DoEvents
        '打开过滤页面
        Call setPicMove(Index, decode(Index, imgType处方权, Filter处方权, imgType抗菌药物, Filter抗菌药物, imgType手术等级, Filter手术等级, Filter特殊医嘱))
        mblnMove = True
        imgSentence(Index).Picture = imgUpDown.ListImages.Item("Up").Picture
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setPicMove(Index As Integer, picindex As Integer)
    Dim i As Integer
    On Error GoTo errH
    picFilter(picindex).Top = picHead.Top + imgSentence(Index).Top + imgSentence(Index).Height + 10
    If Index = 6 Then
        picFilter(picindex).Left = picHead.Left + imgSentence(Index).Left - picFilter(picindex).Width + imgSentence(Index).Width
    Else
        picFilter(picindex).Left = picHead.Left + imgSentence(Index).Left - (picFilter(picindex).Width / 2) + (imgSentence(Index).Width / 2)
    End If
    picFilter(picindex).Visible = True
    picFilter(picindex).Tag = Index
    
    Select Case picindex
        '处方权
        Case Filter处方权
            If imgSentence(Index).Tag = "" Then
                For i = 0 To ChkFd.Count - 1
                    ChkFd(i).Value = 1
                Next
            Else
                For i = 0 To ChkFd.Count - 1
                    ChkFd(i).Value = IIf(Split(imgSentence(Index).Tag, "|")(i) = 1, 1, 0)
                Next
            End If
            ChkFd(0).SetFocus
        '门诊与住院
        Case Filter特殊医嘱
            If imgSentence(Index).Tag = "" Then
                ChkFdmz(0).Value = 1: ChkFdmz(1).Value = 1: ChkFdZy(0).Value = 1: ChkFdZy(1).Value = 1
            Else
                ChkFdmz(0).Value = IIf(Split(imgSentence(Index).Tag, "|")(0) = 1, 1, 0)
                ChkFdmz(1).Value = IIf(Split(imgSentence(Index).Tag, "|")(1) = 1, 1, 0)
                ChkFdZy(0).Value = IIf(Split(imgSentence(Index).Tag, "|")(2) = 1, 1, 0)
                ChkFdZy(1).Value = IIf(Split(imgSentence(Index).Tag, "|")(3) = 1, 1, 0)
            End If
            ChkFdmz(0).SetFocus
        '手术等级
        Case Filter手术等级
            If imgSentence(Index).Tag = "" Then
                For i = 0 To ChkFdSS.Count - 1
                    ChkFdSS(i).Value = 1
                Next
            Else
                For i = 0 To ChkFdSS.Count - 1
                    ChkFdSS(i).Value = IIf(Split(imgSentence(Index).Tag, "|")(i) = 1, 1, 0)
                Next
            End If
            ChkFdSS(0).SetFocus
        '抗菌药物权限
        Case Filter抗菌药物
            If imgSentence(Index).Tag = "" Then
                For i = 0 To ChkFdKSSmz.Count - 1
                    ChkFdKSSmz(i).Value = 1
                Next
                For i = 0 To ChkFdKSSzy.Count - 1
                    ChkFdKSSzy(i).Value = 1
                Next
            Else
                For i = 0 To ChkFdKSSmz.Count - 1
                    ChkFdKSSmz(i).Value = IIf(Split(imgSentence(Index).Tag, "|")(i) = 1, 1, 0)
                Next
                For i = 0 To ChkFdKSSzy.Count - 1
                    ChkFdKSSzy(i).Value = IIf(Split(imgSentence(Index).Tag, "|")(i + 4) = 1, 1, 0)
                Next
            End If
            ChkFdKSSmz(0).SetFocus
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'设置过滤按钮图标,缓存过滤条件
Private Sub setUseIco(Index As Integer)
    Dim i As Integer
    Dim blnChk As Boolean
    Dim strTmp As String
    
    On Error GoTo errH
    
    blnChk = True
    Select Case Index
        '处方权
        Case imgType处方权
            blnChk = ChkFd(0).Value And ChkFd(1).Value
            strTmp = "|" & IIf(ChkFd(0).Value, "1", "0") & "|" & IIf(ChkFd(1).Value, "1", "0")
        '抗菌药物权限
        Case imgType抗菌药物
            For i = 0 To ChkFdKSSmz.Count - 1
                blnChk = blnChk And ChkFdKSSmz(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdKSSmz(i).Value, "1", "0")
            Next
            For i = 0 To ChkFdKSSzy.Count - 1
                blnChk = blnChk And ChkFdKSSzy(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdKSSzy(i).Value, "1", "0")
            Next
        '手术等级
        Case imgType手术等级
            For i = 0 To ChkFdSS.Count - 1
                blnChk = blnChk And ChkFdSS(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdSS(i).Value, "1", "0")
            Next
        '特殊医嘱权限
        Case imgType毒类, imgType麻醉, imgType精神, imgType贵重
            blnChk = ChkFdmz(0).Value And ChkFdmz(1).Value And ChkFdZy(0).Value And ChkFdZy(1).Value
            strTmp = "|" & IIf(ChkFdmz(0).Value, "1", "0") & "|" & IIf(ChkFdmz(1).Value, "1", "0") & "|" & IIf(ChkFdZy(0).Value, "1", "0") & "|" & IIf(ChkFdZy(1).Value, "1", "0")
    End Select
    
    If Not blnChk Then
        strTmp = Mid(strTmp, 2)
        imgSentence(Index).Tag = strTmp
    Else
        imgSentence(Index).Tag = ""
    End If
    
    imgSentence(Index).Picture = IIf(blnChk, imgUpDown.ListImages.Item("Down").Picture, imgUpDown.ListImages.Item("Used").Picture)
    
    If FilterSQL = True Then
        mrsTmp.Filter = mstrFdsql
        If Not mrsTmp.EOF Then
            mrsTmp.MoveFirst
        End If
        '设置状态栏数据
        If mlngFindNum = 1 Then
            stbThis.Panels(2).Text = "当前查找一共有" & mrsTmp.RecordCount & "名医生"
        Else
            stbThis.Panels(2).Text = "当前部门一共有" & mrsTmp.RecordCount & "名医生"
        End If
        If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(已启用过滤)"
        For i = 0 To 5
            If Not mrsTmp.EOF Then
                LoadData (i)
                mrsTmp.MoveNext
                picInfo.Item(i).Visible = True
            Else
                picInfo.Item(i).Visible = False
            End If
        Next
        vscBar.Value = 0
        vscBar.Max = IIf(mrsTmp.RecordCount - 6 > 0, mrsTmp.RecordCount - 6, 0)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'列表光标跟随效果
Private Sub picInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo errH
    If mlngIndex <> 999 And mlngIndex <> Index Then
        LineFg.Item(mlngIndex).BorderWidth = 1
        LineFg.Item(mlngIndex).BorderColor = &HE0E0E0
        picInfo.Item(mlngIndex).BackColor = &HFFFFFF
        cboLevel.Item(mlngIndex).BackColor = &HFFFFFF
    End If
    If mlngIndex <> Index Then
        LineFg.Item(Index).BorderWidth = 2
        LineFg.Item(Index).BorderColor = &HFF&
        picInfo.Item(Index).BackColor = 12648447
        cboLevel.Item(Index).BackColor = 12648447
    End If
    mlngIndex = Index
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



'勾选权限复选框,并执行
Private Sub PicCheck(obj As Image)
    Dim strSql As String
    Dim intIndex As Integer
    Dim intType As Integer  '权限类型
    Dim strCheck As String
    Dim int场合 As Integer  '=0 全部, =1 住院,=2 门诊
    Dim curDate As Date
    Dim i As Integer
    Dim objTmp As Object
    
    On Error GoTo errH
    intIndex = obj.Index
    DoEvents
    Select Case obj.Name
        '处方权勾选
        Case "Pic处方权"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                obj.Picture = img16.ListImages.Item("Black2").Picture
                obj.Tag = "True"
            End If
            intType = Priv处方权
            strCheck = Val(IIf(obj.Tag = "True", 1, 0))
        '门诊抗菌药物权限勾选
        Case "PicMZFxz", "PicMZXz", "PicMZTs"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                For i = 1 To 3
                    Set objTmp = mcolCtl(decode(i, 1, "PicMZFxz", 2, "PicMZXz", 3, "PicMZTs") & intIndex)
                    objTmp.Picture = img16.ListImages.Item("None").Picture: objTmp.Tag = ""
                Next
                obj.Picture = img16.ListImages.Item("Red2").Picture
                obj.Tag = "True"
            End If
            int场合 = 2
            intType = Priv门诊抗菌
            strCheck = IIf(obj.Tag = "True", IIf(obj.Name = "PicMZFxz", 1, IIf(obj.Name = "PicMZXz", 2, 3)), "")
        '住院抗菌药物权限勾选
        Case "PicZYFxz", "PicZYXz", "PicZYTs"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                For i = 1 To 3
                    Set objTmp = mcolCtl(decode(i, 1, "PicZYFxz", 2, "PicZYXz", 3, "PicZYTs") & intIndex)
                    objTmp.Picture = img16.ListImages.Item("None").Picture: objTmp.Tag = ""
                Next
                obj.Picture = img16.ListImages.Item("Red2").Picture
                obj.Tag = "True"
            End If
            int场合 = 1
            intType = Priv住院抗菌
            strCheck = IIf(obj.Tag = "True", IIf(obj.Name = "PicZYFxz", 1, IIf(obj.Name = "PicZYXz", 2, 3)), "")
        Case "PicSS1", "PicSS2", "PicSS3", "PicSS4"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                For i = 1 To 4
                    Set objTmp = mcolCtl(decode(i, 1, "PicSS1", 2, "PicSS2", 3, "PicSS3", 4, "PicSS4") & intIndex)
                    objTmp.Picture = img16.ListImages.Item("None").Picture: objTmp.Tag = ""
                Next
                obj.Picture = img16.ListImages.Item("Blue2").Picture
                obj.Tag = "True"
            End If
            intType = Priv手术等级
            
            If obj.Tag = "True" Then
                strCheck = decode(Val(Mid(obj.Name, Len(obj.Name), 1)), 1, "一级", 2, "二级", 3, "三级", 4, "四级")
            Else
                strCheck = ""
            End If


        Case "PicMZDl", "PicMZMz", "PicMZJs", "PicMZGz"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                obj.Picture = img16.ListImages.Item("Green2").Picture
                obj.Tag = "True"
            End If
            intType = Priv门诊特殊
            For i = 1 To 4
                Set objTmp = mcolCtl(decode(i, 1, "PicMZDl", 2, "PicMZMz", 3, "PicMZJs", 4, "PicMZGz") & intIndex)
                strCheck = strCheck & IIf(objTmp.Tag = "True", "1", "0")
            Next
        Case "PicZYDl", "PicZYGz", "PicZYJs", "PicZYMz"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                obj.Picture = img16.ListImages.Item("Green2").Picture
                obj.Tag = "True"
            End If
            intType = Priv住院特殊
            For i = 1 To 4
                Set objTmp = mcolCtl(decode(i, 1, "PicZYDl", 2, "PicZYMz", 3, "PicZYJs", 4, "PicZYGz") & intIndex)
                strCheck = strCheck & IIf(objTmp.Tag = "True", "1", "0")
            Next
    End Select
    
    If intType <> 0 Then
        If intType = Priv门诊抗菌 Or intType = Priv住院抗菌 Then
            curDate = sys.Currentdate
            strSql = "Zl_医生权限_Update(" & Val(picInfo(intIndex).Tag) & "," & intType & ",'" & strCheck & "'," & int场合 & ",'" & UserInfo.姓名 & "',to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Else
            strSql = "Zl_医生权限_Update(" & Val(picInfo(intIndex).Tag) & "," & intType & ",'" & strCheck & "')"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        
        Call CacheData(Val(picInfo(intIndex).Tag), intType, strCheck)

    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub CacheData(uid As Long, intType As Integer, strCheck As String)
    '缓存本地数据
    Dim lngDw As Long
    Dim strFilter As String
    Dim strTmp As String
On Error GoTo errH
    '保存记录数据
    strFilter = IIf(mrsTmp.Filter <> 0, mrsTmp.Filter, "")
    lngDw = mrsTmp.AbsolutePosition
    
    strTmp = decode(intType, Priv处方权, "处方权标志", Priv门诊抗菌, "门诊抗菌药物权限", Priv住院抗菌, "住院抗菌药物权限", Priv手术等级, "手术等级", Priv门诊特殊, "门诊特殊医嘱权限", Priv住院特殊, "住院特殊医嘱权限", Priv医生等级, "专业技术职务")
    If intType = Priv住院抗菌 Or intType = Priv门诊抗菌 Then
        strCheck = decode(Val(strCheck), 1, "非限制使用", 2, "限制使用", 3, "特殊使用", "null")
    End If
    If strCheck = "" Then
        If intType = Priv手术等级 Then
            strCheck = "null"
        End If
    End If
    mrsTmp.Filter = "Id=" & uid
    If Not mrsTmp.EOF Then
        mrsTmp.Update strTmp, strCheck
        If intType = Priv门诊特殊 Then
            mrsTmp.Update "门诊毒类", Val(Mid(strCheck, 1, 1))
            mrsTmp.Update "门诊麻醉", Val(Mid(strCheck, 2, 1))
            mrsTmp.Update "门诊精神", Val(Mid(strCheck, 3, 1))
            mrsTmp.Update "门诊贵重", Val(Mid(strCheck, 4, 1))
        ElseIf intType = Priv住院特殊 Then
            mrsTmp.Update "住院毒类", Val(Mid(strCheck, 1, 1))
            mrsTmp.Update "住院麻醉", Val(Mid(strCheck, 2, 1))
            mrsTmp.Update "住院精神", Val(Mid(strCheck, 3, 1))
            mrsTmp.Update "住院贵重", Val(Mid(strCheck, 4, 1))
        End If
    End If
    '恢复记录集
    mrsTmp.Filter = strFilter
    If (Not mrsTmp.EOF) And mrsTmp.RecordCount <> 0 Then
        lngDw = IIf(lngDw > 0, lngDw, 1)
        lngDw = IIf(lngDw > mrsTmp.RecordCount, mrsTmp.RecordCount, lngDw)
        mrsTmp.AbsolutePosition = lngDw
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii <> vbKeyReturn Then Exit Sub
    On Error GoTo errH
    
    If txtFind.Text <> "" Then
        mlngFindNum = 1
    Else
        mlngFindNum = 0
    End If
    
    If FilterSQL Then
        mrsTmp.Filter = mstrFdsql
    End If
    
    If Not mrsTmp.EOF Then
        mrsTmp.MoveFirst
    End If
    
    '设置状态栏数据
    If mlngFindNum = 1 Then
        stbThis.Panels(2).Text = "当前查找一共有" & mrsTmp.RecordCount & "名医生"
    Else
        stbThis.Panels(2).Text = "当前部门一共有" & mrsTmp.RecordCount & "名医生"
    End If
    If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(已启用过滤)"
    
    For i = 0 To 5
        If Not mrsTmp.EOF Then
            LoadData (i)
            mrsTmp.MoveNext
            picInfo.Item(i).Visible = True
        Else
            picInfo.Item(i).Visible = False
        End If
    Next
    vscBar.Value = 0
    vscBar.Max = IIf(mrsTmp.RecordCount - 6 > 0, mrsTmp.RecordCount - 6, 0)
    txtFind.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vscBar_Change()
    Dim i As Long
    On Error GoTo errH
    If vscBar.Value + 1 > 0 Then
        If mrsTmp.RecordCount <> 0 Then
            mrsTmp.AbsolutePosition = vscBar.Value + 1
        End If
        LockWindowUpdate Me.hwnd
        For i = 0 To 5
            If Not mrsTmp.EOF Then
                LoadData (i)
                mrsTmp.MoveNext
                picInfo.Item(i).Visible = True
            Else
                picInfo.Item(i).Visible = False
            End If
        Next
        LockWindowUpdate 0
    End If
    picBack.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub ChkFdKSSzy_LostFocus(Index As Integer)
    Call CloseFilter(Filter抗菌药物)
End Sub

Private Sub ChkFdKSSmz_LostFocus(Index As Integer)
    Call CloseFilter(Filter抗菌药物)
End Sub

Private Sub ChkFd_LostFocus(Index As Integer)
    Call CloseFilter(Filter处方权)
End Sub

Private Sub ChkFdmz_LostFocus(Index As Integer)
    Call CloseFilter(Filter特殊医嘱)
End Sub

Private Sub ChkFdZy_LostFocus(Index As Integer)
    Call CloseFilter(Filter特殊医嘱)
End Sub

Private Sub ChkFdSS_LostFocus(Index As Integer)
    Call CloseFilter(Filter手术等级)
End Sub

Private Sub picTmp_LostFocus(Index As Integer)
    Call CloseFilter(Index)
End Sub

Private Sub ChkFdKSSzy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub ChkFdKSSmz_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub ChkFdSS_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub ChkFdmz_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub ChkFdZy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub ChkFd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub lblPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub picTmp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub picFilter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = True
End Sub


Private Sub PicMZDl_Click(Index As Integer)
    PicCheck PicMZDl(Index)
End Sub

Private Sub PicMZFxz_Click(Index As Integer)
    PicCheck PicMZFxz(Index)
End Sub

Private Sub PicMZGz_Click(Index As Integer)
    PicCheck PicMZGz(Index)
End Sub

Private Sub PicMZJs_Click(Index As Integer)
    PicCheck PicMZJs(Index)
End Sub

Private Sub PicMZMz_Click(Index As Integer)
    PicCheck PicMZMz(Index)
End Sub

Private Sub PicMZTs_Click(Index As Integer)
    PicCheck PicMZTs(Index)
End Sub

Private Sub PicMZXz_Click(Index As Integer)
    PicCheck PicMZXz(Index)
End Sub

Private Sub PicSS1_Click(Index As Integer)
    PicCheck PicSS1(Index)
End Sub

Private Sub PicSS2_Click(Index As Integer)
    PicCheck PicSS2(Index)
End Sub

Private Sub PicSS3_Click(Index As Integer)
    PicCheck PicSS3(Index)
End Sub

Private Sub PicSS4_Click(Index As Integer)
    PicCheck PicSS4(Index)
End Sub

Private Sub PicZYDl_Click(Index As Integer)
    PicCheck PicZYDl(Index)
End Sub

Private Sub PicZYFxz_Click(Index As Integer)
    PicCheck PicZYFxz(Index)
End Sub

Private Sub PicZYGz_Click(Index As Integer)
    PicCheck PicZYGz(Index)
End Sub

Private Sub PicZYJs_Click(Index As Integer)
    PicCheck PicZYJs(Index)
End Sub

Private Sub PicZYMz_Click(Index As Integer)
    PicCheck PicZYMz(Index)
End Sub

Private Sub PicZYTs_Click(Index As Integer)
    PicCheck PicZYTs(Index)
End Sub

Private Sub PicZYXz_Click(Index As Integer)
    PicCheck PicZYXz(Index)
End Sub

Private Sub Pic处方权_Click(Index As Integer)
    PicCheck Pic处方权(Index)
End Sub

Private Sub ChkFdKSSmz_Click(Index As Integer)
    Call CheckBtn(ChkFdKSSmz, Index)
End Sub

Private Sub ChkFdKSSzy_Click(Index As Integer)
    Call CheckBtn(ChkFdKSSzy, Index)
End Sub

Private Sub ChkFdmz_Click(Index As Integer)
    Call CheckBtn(ChkFdmz, Index)
End Sub

Private Sub ChkFdSS_Click(Index As Integer)
    Call CheckBtn(ChkFdSS, Index)
End Sub

Private Sub ChkFdZy_Click(Index As Integer)
    Call CheckBtn(ChkFdZy, Index)
End Sub

Private Sub ChkFd_Click(Index As Integer)
    Call CheckBtn(ChkFd, Index)
End Sub

