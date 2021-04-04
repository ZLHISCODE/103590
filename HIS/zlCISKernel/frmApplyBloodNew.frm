VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyBloodNew 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输血申请单"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10650
   Icon            =   "frmApplyBloodNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   285
      Index           =   6
      Left            =   8430
      TabIndex        =   125
      Top             =   3690
      Width           =   2190
      Begin VB.OptionButton optAppraise 
         BackColor       =   &H8000000E&
         Caption         =   "未评估"
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
         Index           =   0
         Left            =   1155
         TabIndex        =   29
         Top             =   30
         Width           =   975
      End
      Begin VB.OptionButton optAppraise 
         BackColor       =   &H8000000E&
         Caption         =   "已评估"
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
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   30
         Width           =   945
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Index           =   5
      Left            =   2250
      TabIndex        =   123
      Top             =   3690
      Width           =   1815
      Begin VB.OptionButton optConsent 
         BackColor       =   &H8000000E&
         Caption         =   "未签"
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
         Index           =   0
         Left            =   960
         TabIndex        =   27
         Top             =   30
         Width           =   735
      End
      Begin VB.OptionButton optConsent 
         BackColor       =   &H8000000E&
         Caption         =   "已签"
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
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox picHisItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   6660
      TabIndex        =   121
      Top             =   10440
      Width           =   6660
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "本次历史申请项目："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   41
         Left            =   0
         TabIndex        =   122
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Index           =   5
      Left            =   1440
      ScaleHeight     =   270
      ScaleWidth      =   2115
      TabIndex        =   119
      Top             =   2430
      Width           =   2115
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   -30
         Width           =   2115
      End
   End
   Begin VB.PictureBox picBloodDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3750
      ScaleHeight     =   315
      ScaleWidth      =   6870
      TabIndex        =   106
      Top             =   4575
      Width           =   6870
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   305
         Index           =   3
         Left            =   5115
         ScaleHeight     =   300
         ScaleWidth      =   990
         TabIndex        =   109
         Top             =   0
         Width           =   990
         Begin VB.ComboBox cboInfo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   -45
            TabIndex        =   110
            Text            =   "cboInfo"
            Top             =   -30
            Width           =   1005
         End
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   305
         Index           =   8
         Left            =   1245
         ScaleHeight     =   300
         ScaleWidth      =   3090
         TabIndex        =   107
         Top             =   0
         Width           =   3090
         Begin VB.ComboBox cboInfo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   -30
            Width           =   3135
         End
      End
      Begin VB.Line Line1 
         Index           =   25
         X1              =   4965
         X2              =   6045
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "滴速"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   30
         Left            =   4485
         TabIndex        =   112
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "滴/分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   6120
         TabIndex        =   111
         Top             =   30
         Width           =   645
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "发血执行"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   18
         Left            =   105
         TabIndex        =   108
         Top             =   30
         Width           =   1020
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   1170
         X2              =   4350
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Index           =   3
      Left            =   5160
      TabIndex        =   101
      Top             =   3270
      Width           =   1515
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "无"
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
         Index           =   2
         Left            =   960
         TabIndex        =   23
         Top             =   30
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "有"
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
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   30
         Width           =   615
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Index           =   4
      Left            =   8835
      TabIndex        =   100
      Top             =   3270
      Width           =   1830
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "有"
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
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   30
         Width           =   615
      End
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "无"
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
         Index           =   4
         Left            =   960
         TabIndex        =   25
         Top             =   30
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd常用嘱托 
      Height          =   300
      Left            =   5355
      Picture         =   "frmApplyBloodNew.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "将当前嘱托设置为常用嘱托(Ctrl+D)"
      Top             =   9660
      Width           =   315
   End
   Begin VB.CommandButton cmd医生嘱托 
      Caption         =   "…"
      Height          =   265
      Left            =   5055
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   9675
      Width           =   285
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   345
      Index           =   1
      Left            =   1380
      TabIndex        =   93
      Top             =   2820
      Width           =   2565
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   13
         Left            =   75
         MaxLength       =   2
         TabIndex        =   16
         Top             =   45
         Width           =   700
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   14
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   17
         Top             =   45
         Width           =   700
      End
      Begin VB.Line Line1 
         Index           =   22
         X1              =   1155
         X2              =   1855
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "孕"
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
         Index           =   26
         Left            =   885
         TabIndex        =   95
         Top             =   90
         Width           =   210
      End
      Begin VB.Line Line1 
         Index           =   23
         X1              =   75
         X2              =   775
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "产"
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
         Index           =   27
         Left            =   1950
         TabIndex        =   94
         Top             =   90
         Width           =   210
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Index           =   9
      Left            =   795
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   9690
      Width           =   4425
   End
   Begin VB.CheckBox chkWait 
      BackColor       =   &H80000005&
      Caption         =   "待诊"
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
      Left            =   9795
      TabIndex        =   12
      Top             =   2025
      Width           =   735
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10740
      TabIndex        =   90
      Top             =   4590
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   251330561
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picNo 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8760
      ScaleHeight     =   495
      ScaleWidth      =   1935
      TabIndex        =   88
      Top             =   660
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   0
         TabIndex        =   89
         Top             =   165
         Width           =   255
      End
      Begin VB.Line Line1 
         Index           =   34
         X1              =   240
         X2              =   1680
         Y1              =   390
         Y2              =   390
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Index           =   11
      Left            =   4725
      ScaleHeight     =   270
      ScaleWidth      =   2895
      TabIndex        =   87
      Top             =   2430
      Width           =   2895
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   -30
         Width           =   2955
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Index           =   10
      Left            =   9240
      ScaleHeight     =   270
      ScaleWidth      =   1215
      TabIndex        =   86
      Top             =   1605
      Width           =   1215
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   -30
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Left            =   9375
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2025
      Width           =   270
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   1
      Left            =   10215
      Picture         =   "frmApplyBloodNew.frx":6DDC
      Style           =   1  'Graphical
      TabIndex        =   85
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   10770
      Width           =   285
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   0
      Left            =   4200
      Picture         =   "frmApplyBloodNew.frx":6ED2
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   4170
      Width           =   285
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   9
      Left            =   7380
      ScaleHeight     =   300
      ScaleWidth      =   3180
      TabIndex        =   83
      Top             =   7380
      Width           =   3180
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   -25
         Width           =   3135
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   20
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   9705
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Index           =   19
      Left            =   8535
      TabIndex        =   48
      Text            =   "2013-06-20 18:00"
      Top             =   10785
      Width           =   1935
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   18
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   10395
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   17
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   10065
      Width           =   1455
   End
   Begin VB.PictureBox picGet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   1
      Left            =   2085
      ScaleHeight     =   300
      ScaleWidth      =   3375
      TabIndex        =   75
      Top             =   7380
      Width           =   3375
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   1
         Left            =   2970
         Picture         =   "frmApplyBloodNew.frx":6FC8
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   0
         Width           =   285
      End
      Begin VB.TextBox txtGet 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   2940
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   2
      Left            =   9555
      ScaleHeight     =   300
      ScaleWidth      =   1095
      TabIndex        =   70
      Top             =   4140
      Width           =   1095
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   -30
         Width           =   975
      End
      Begin VB.PictureBox picRH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   660
         ScaleHeight     =   165
         ScaleWidth      =   270
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   75
         Visible         =   0   'False
         Width           =   270
         Begin VB.Label lblRh 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   128
            Top             =   105
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   1
      Left            =   7395
      ScaleHeight     =   300
      ScaleWidth      =   1380
      TabIndex        =   68
      Top             =   4140
      Width           =   1380
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "frmApplyBloodNew.frx":70BE
         Left            =   -25
         List            =   "frmApplyBloodNew.frx":70C0
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   -30
         Width           =   1335
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   30
      Text            =   "2013-06-20 18:00"
      Top             =   4170
      Width           =   2415
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Index           =   2
      Left            =   8820
      TabIndex        =   65
      Top             =   2865
      Width           =   1815
      Begin VB.OptionButton optPossession 
         BackColor       =   &H8000000E&
         Caption         =   "本市"
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   30
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPossession 
         BackColor       =   &H8000000E&
         Caption         =   "外埠"
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
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Index           =   0
      Left            =   1335
      TabIndex        =   62
      Top             =   3270
      Width           =   1575
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "无"
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
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   30
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optHistory 
         BackColor       =   &H8000000E&
         Caption         =   "有"
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
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   30
         Width           =   615
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Index           =   0
      Left            =   9000
      ScaleHeight     =   270
      ScaleWidth      =   1560
      TabIndex        =   60
      Top             =   2430
      Width           =   1560
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   -25
         Width           =   1530
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   7
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   5
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1605
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   4
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   3
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   1
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1605
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   0
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid vsLIS 
      Height          =   1560
      Left            =   240
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8025
      Width           =   10290
      _cx             =   18150
      _cy             =   2752
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   29
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmApplyBloodNew.frx":70C2
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   15
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   9705
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Index           =   8
      Left            =   1560
      MaxLength       =   500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2025
      Width           =   7800
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   750
      Top             =   570
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
            Picture         =   "frmApplyBloodNew.frx":7315
            Key             =   "c0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplyBloodNew.frx":78AF
            Key             =   "c1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplyBloodNew.frx":7E49
            Key             =   "o0"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplyBloodNew.frx":83E3
            Key             =   "o1"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPreBlood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   240
      ScaleHeight     =   2685
      ScaleWidth      =   10290
      TabIndex        =   104
      Top             =   4635
      Width           =   10290
      Begin VB.PictureBox picPreInfo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   10185
         TabIndex        =   113
         Top             =   2325
         Width           =   10185
         Begin VB.PictureBox picPreSum 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   7785
            ScaleHeight     =   450
            ScaleWidth      =   2370
            TabIndex        =   115
            Top             =   0
            Width           =   2370
            Begin VB.TextBox txt申请量 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   585
               Locked          =   -1  'True
               TabIndex        =   35
               TabStop         =   0   'False
               Text            =   "100000"
               Top             =   30
               Width           =   795
            End
            Begin VB.PictureBox picInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   305
               Index           =   4
               Left            =   1440
               ScaleHeight     =   300
               ScaleWidth      =   1005
               TabIndex        =   116
               Top             =   0
               Width           =   1005
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   360
                  Index           =   4
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   -30
                  Width           =   975
               End
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "总量"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   20
               Left            =   0
               TabIndex        =   117
               Top             =   45
               Width           =   510
            End
            Begin VB.Line Line1 
               Index           =   18
               X1              =   540
               X2              =   2475
               Y1              =   315
               Y2              =   315
            End
         End
         Begin VB.TextBox txt申请信息 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   114
            TabStop         =   0   'False
            Text            =   "品种："
            Top             =   45
            Width           =   7110
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBlood 
         Height          =   1935
         Left            =   0
         TabIndex        =   33
         Top             =   315
         Width           =   6375
         _cx             =   11245
         _cy             =   3413
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   16769985
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmApplyBloodNew.frx":897D
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1935
         Left            =   6420
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   315
         Width           =   3825
         _cx             =   6747
         _cy             =   3413
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   16769985
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmApplyBloodNew.frx":8AF2
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "血液信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   39
         Left            =   15
         TabIndex        =   105
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox picGet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2040
      ScaleHeight     =   300
      ScaleWidth      =   3375
      TabIndex        =   72
      Top             =   6855
      Width           =   3375
      Begin VB.TextBox txtGet 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   37
         Top             =   15
         Width           =   3015
      End
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7410
      MaxLength       =   10
      TabIndex        =   38
      Top             =   6855
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   9585
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6855
      Width           =   885
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血前评估"
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
      Index           =   43
      Left            =   7320
      TabIndex        =   126
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血治疗知情同意书"
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
      Index           =   42
      Left            =   270
      TabIndex        =   124
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   1920
      X2              =   5340
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "申请类型"
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
      Index           =   28
      Left            =   480
      TabIndex        =   120
      Top             =   2460
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   1440
      X2              =   3500
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "24小时内输血总量："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   40
      Left            =   240
      TabIndex        =   118
      Top             =   10800
      Width           =   2040
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血禁忌症及过敏史"
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
      Index           =   38
      Left            =   6900
      TabIndex        =   103
      Top             =   3300
      Width           =   1890
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "既往输血反应史"
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
      Index           =   32
      Left            =   3660
      TabIndex        =   102
      Top             =   3300
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "RH(D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   8820
      TabIndex        =   69
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "注意"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   25
      Left            =   240
      TabIndex        =   92
      Top             =   10050
      Width           =   450
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   660
      X2              =   5330
      Y1              =   9945
      Y2              =   9945
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
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
      Height          =   210
      Index           =   24
      Left            =   225
      TabIndex        =   91
      Top             =   9705
      Width           =   420
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   8535
      X2              =   10095
      Y1              =   9975
      Y2              =   9975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "上级医师签名"
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
      Index           =   36
      Left            =   7095
      TabIndex        =   82
      Top             =   10080
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   8535
      X2              =   10215
      Y1              =   11040
      Y2              =   11040
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "输血申请日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   35
      Left            =   7095
      TabIndex        =   81
      Top             =   10785
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   8535
      X2              =   10095
      Y1              =   10665
      Y2              =   10665
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "采 集 者签名"
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
      Index           =   34
      Left            =   7095
      TabIndex        =   80
      Top             =   10425
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   8550
      X2              =   10110
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "受 血 者"
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
      Index           =   23
      Left            =   240
      TabIndex        =   78
      Top             =   7770
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   7275
      X2              =   10470
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血执行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   22
      Left            =   6195
      TabIndex        =   77
      Top             =   7410
      Width           =   1020
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "输血途径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   21
      Left            =   780
      TabIndex        =   74
      Top             =   7410
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   7275
      X2              =   10500
      Y1              =   7125
      Y2              =   7125
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "预定输血量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   19
      Left            =   5940
      TabIndex        =   73
      Top             =   6915
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   1920
      X2              =   5340
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "预定输血成分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   17
      Left            =   270
      TabIndex        =   71
      Top             =   6915
      Width           =   1785
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   9435
      X2              =   10515
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   7275
      X2              =   8715
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   1920
      X2              =   4200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "预定输血日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   66
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "受血者属地"
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
      Index           =   13
      Left            =   7740
      TabIndex        =   64
      Top             =   2895
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "孕产情况"
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
      Index           =   12
      Left            =   480
      TabIndex        =   63
      Top             =   2895
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "既往输血史"
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
      Index           =   11
      Left            =   270
      TabIndex        =   61
      Top             =   3300
      Width           =   1050
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   8880
      X2              =   10470
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   4725
      X2              =   7625
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1440
      X2              =   9650
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   6600
      X2              =   7920
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   9120
      X2              =   10450
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1440
      X2              =   2880
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6600
      X2              =   7920
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4080
      X2              =   5400
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1440
      X2              =   2880
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4080
      X2              =   5400
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9120
      X2              =   10450
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3960
      X2              =   6720
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label lblHead 
      BackColor       =   &H8000000E&
      Caption         =   "临床输血申请单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4005
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血性质"
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
      Index           =   10
      Left            =   7950
      TabIndex        =   59
      Top             =   2460
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "输血目的"
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
      Index           =   9
      Left            =   3765
      TabIndex        =   58
      Top             =   2460
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "临床诊断"
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
      Index           =   8
      Left            =   480
      TabIndex        =   57
      Top             =   2055
      Width           =   840
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "用血安排"
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
      Index           =   6
      Left            =   8160
      TabIndex        =   55
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "床    号"
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
      Index           =   5
      Left            =   480
      TabIndex        =   54
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "科    别"
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
      Index           =   4
      Left            =   5625
      TabIndex        =   53
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "姓    名"
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
      Index           =   2
      Left            =   480
      TabIndex        =   51
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "住 院 号"
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
      Index           =   1
      Left            =   3120
      TabIndex        =   50
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "就诊类型"
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
      Index           =   0
      Left            =   8160
      TabIndex        =   49
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "年    龄"
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
      Index           =   7
      Left            =   5640
      TabIndex        =   56
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "性    别"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   52
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "血型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   6705
      TabIndex        =   67
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "申请医师签名"
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
      Index           =   33
      Left            =   7095
      TabIndex        =   79
      Top             =   9735
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "申请医师坐标"
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
      Index           =   29
      Left            =   7095
      TabIndex        =   97
      Top             =   9735
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   24
      Visible         =   0   'False
      X1              =   8535
      X2              =   10095
      Y1              =   9975
      Y2              =   9975
   End
End
Attribute VB_Name = "frmApplyBloodNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng病人性质 As Long   '0-住院，1-门诊
Private mblnChange As Boolean
Private mblnHaveAuditPriv As Boolean '执业医师资格
Private mintType As Integer   '0-新增，1-修改，2-查看,3-医嘱编辑调用，只能调整除输血成分，总量，申请时间，输血时间，执行科室，输血途径，输血执行科室，用血安排以外的内容；4-医生核对用血申请(输血科直接发血产生的医嘱)
Private mlngUpdateAdvice As Long  '修改的医嘱ID
Private mintPState As Integer
Private mdatTurn As Date
Private mlng病人科室id As Long
Private mlng病区ID As Long
Private mlng开单科室ID As Long
Private mlng输血途径 As Long
Private mlng输血项目ID As Long, mlngPre输血项目ID As Long
Private mstr输血项目 As String  '备血申请品种可选择多个。格式：项目ID,申请量,申请血型,申请RH
Private mstrLISAboRHCode As String
Private mstr入院时间 As String
Private mstr上次转科时间 As String
Private mrsDefine As Recordset
Private mobjVBA As Object
Private mobjScript As clsScript
Private mlng执行科室性质 As Long
Private mlng输血执行性质 As Long
Private mbln补录 As Boolean
Private mblnEditable As Boolean
Private mobjReport As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mstr诊断IDs As String  '诊断关联
Private mlng录入限量 As Long
Private mint申请单打印模式 As Integer  '1-发送时打印，0-新开时打印
Private mint险类 As Integer '当前病人险类
Private mbln提醒对码 As Boolean
Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象
Attribute mclsMipModule.VB_VarHelpID = -1
Private Const CON_LisResultCol = 3
Private Const CON_LisResultCount = 10
Private mobjPublicLis As Object
Private mint场合 As Integer '0-住院，1－门诊，默认为住院
Private mstr挂号单 As String '挂号单号
Private mlng挂号ID As Long
Private mlng前提ID As Long
Private mrsCard As ADODB.Recordset
Private mbytBaby As Integer  '婴儿序号
Private mstr检查入院诊断 As String
Private mint调用场合 As Integer '0-工作站调用，1－医嘱下达界面调用
Private mblnNewSpareBloood As Boolean  '是否采用新备血模式
Private mblnSpareBloood As Boolean  '是备血申请还是用血申请
Private mblnUseBloodSend As Boolean '用血医嘱是否已经发血
Private mstr摘要输血 As String '摘要，由 gclsInsure.GetItemInfo 获取
Private mstr摘要途径 As String
Private mstr费别 As String
Private mblnDataLoad As Boolean
Private mblnSelectBlood As Boolean '开用血申请是否采用医生选择血袋的模式(即存在可用的配血信息，通过医生指定血袋下达申请）

Private Enum Enum_Cbo
    cbo输血性质 = 0
    cbo输血血型 = 1
    cboRHD = 2
    cbo滴速 = 3
    cbo单位 = 4
    cbo输血类型 = 5
    cbo执行科室 = 8
    cbo输血执行 = 9
    cbo用血安排 = 10
    cbo输血目的 = 11
End Enum

Private Enum Enum_lbl
    lbl床号 = 5
    lbl门诊号 = 5
    lbl住院号 = 1
    lbl挂号单 = 1
    lbl既往输血史 = 11
    lbl既往输血反应史 = 32
    lbl输血禁忌及过敏史 = 38
    lbl孕产情况 = 12
    lbl受血者属地 = 13
    lbl预定输血日期 = 14
    lbl血型 = 15
    lblRHD = 16
    lbl预定输血成分 = 17
    lbl发血执行 = 18
    lbl预定输血量 = 19
    lbl输血途径 = 21
    lbl输血执行 = 22
    lbl检验结果 = 23
    lbl备注 = 24
    lbl申请医师签名 = 33
    lbl采集者签名 = 34
    lbl主治医师签名 = 36
    lbl注意 = 25
    '---
    lbl申请医师坐标 = 29
    lbl24H输血量 = 40
    lbl本次历史申请项目 = 41
    lbl知情同意书 = 42
    lbl输血评估 = 43
End Enum

Private Enum Enum_lin
    lin主治医师签名 = 30
    lin采血者签名 = 31
    lin申请医师签名 = 33
    '---
    lin申请医师坐标 = 24
End Enum

Private Enum Enum_txt
    txt就诊类型 = 0
    txt住院号 = 1
    txt挂号单 = 1
    txt姓名 = 2
    txt性别 = 3
    txt科室 = 4
    txt床号 = 5
    txt门诊号 = 5
    txtNO = 6
    txt年龄 = 7
    txt诊断信息 = 8
    txt备注 = 9 '医生嘱托
    txt预定输血时间 = 10
    txt预定输血量 = 11
    txt单位 = 12
    txt孕 = 13
    txt产 = 14
    txt主治医师签名 = 17
    txt采血者签名 = 18
    txt申请日期 = 19
    txt申请医师签名 = 20
    '----
    txt申请医师坐标 = 15
End Enum

Private Enum Enum_FraChk
    fra既往输血史 = 0
    fra孕产情况 = 1
    fra受血者属地 = 2
    fra既往输血反应史 = 3
    fra输血禁忌及过敏史 = 4
    fra知情同意书 = 5
    fra输血评估 = 6
End Enum

Private Enum Enum_Get
    txt预定输血成分 = 0
    txt输血途径 = 1
End Enum

Private Enum Enum_cmdDate
    cmd预定输血时间 = 0
    cmd申请日期 = 1
End Enum

Private Enum Enum_cmdGet
    cmd输血途径 = 1
End Enum

Private Enum Enum_Col
    COL_指标中文名 = 0
    COL_指标结果 = 1
    COL_结果单位 = 2
    COL_指标英文名 = 3
    COL_结果标志 = 4
    COL_结果参考 = 5
    COL_取值序列 = 6
    COL_指标代码 = 7
    COL_检验项目ID = 8
End Enum

Private Enum Enum_P_BloodCol
    COL_P_ID = 0
    COL_P_选择 = 1
    COL_P_编码 = 2
    COL_P_名称 = 3
    COL_P_申请量 = 4
    COL_P_单位 = 5
    COL_P_申请血型 = 6
    COL_P_申请RH = 7
    COL_P_执行分类ID = 8
    COL_P_执行科室ID = 9
    COL_P_录入限量ID = 10
    COL_P_计算系数 = 11
    COL_P_库存 = 12
End Enum

Private Enum Enum_S_BloodList
    COL_S_ID = 0
    COL_S_选择 = 1
    COL_S_编号 = 2
    COL_S_规格 = 3
    COL_S_效期 = 4
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病人性质 As Long, ByVal intType As Integer, Optional ByRef lngUpdateAdvice As Long, _
    Optional ByVal lng病人科室ID As Long, Optional ByVal lng病区ID As Long, Optional ByVal lng开单科室ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
    Optional ByRef rsDefine As Recordset, Optional ByRef objMip As Object, Optional ByVal int场合 As Integer, Optional ByVal str挂号单 As String, Optional ByVal lng项目id As Long, _
    Optional ByRef rsCard As ADODB.Recordset, Optional ByVal bytBaby As Byte, Optional ByVal int调用场合 As Integer, Optional ByVal lng前提ID As Long, Optional ByVal int申请单模式 As Integer = 0) As Boolean
      
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng病人性质 = lng病人性质
    mlng病人科室id = lng病人科室ID
    mlng病区ID = lng病区ID
    mlng开单科室ID = lng开单科室ID
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mint场合 = int场合
    mstr挂号单 = str挂号单
    mlng前提ID = lng前提ID
    Set mrsDefine = rsDefine
    
    mlngUpdateAdvice = lngUpdateAdvice
    
    mlng输血项目ID = lng项目id
    mstr输血项目 = lng项目id & ",,,"
    mlngPre输血项目ID = mlng输血项目ID
    Set mrsCard = rsCard
    mbytBaby = bytBaby
    mint调用场合 = int调用场合
    mblnSpareBloood = (int申请单模式 = 0)
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    If mblnOK = True Then lngUpdateAdvice = mlngUpdateAdvice
    Set rsCard = mrsCard
    ShowMe = mblnOK
End Function

Private Function SeekNextControl() As Boolean
'功能：定位到下一个焦点的控件上
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextControl = True
End Function

Private Sub cboInfo_Change(Index As Integer)
    If Visible And (Index = cbo输血目的 Or Index = cbo输血类型) And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSQL As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = cbo执行科室 Or Index = cbo输血执行 Then
        If cboInfo(Index).ItemData(cboInfo(Index).ListIndex) = -1 Then
            
            '他科执行，弹出选择执行科室
            strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                IIF(gstrNodeNo <> "", " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " Order by A.编码"
            vRect = zlControl.GetControlRect(cboInfo(Index).hwnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vRect.Left, vRect.Top, cboInfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = Cbo.FindIndex(cboInfo(Index), rsTmp!ID)
                If intIdx <> -1 Then
                    cboInfo(Index).ListIndex = intIdx
                Else
                    cboInfo(Index).AddItem rsTmp!编码 & "-" & rsTmp!名称, cboInfo(Index).ListCount - 1
                    cboInfo(Index).ItemData(cboInfo(Index).NewIndex) = rsTmp!ID
                    cboInfo(Index).ListIndex = cboInfo(Index).NewIndex
                End If
                If cboInfo(Index).ListIndex >= 0 Then
                    cboInfo(Index).Tag = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
                End If
                '恢复成现有的科室(不引发Click)
                If cboInfo(Index).Tag <> "" Then
                    intIdx = Cbo.FindIndex(cboInfo(Index), Val(cboInfo(Index).Tag))
                    Call zlControl.CboSetIndex(cboInfo(Index).hwnd, intIdx)
                End If
            End If
        End If
    ElseIf Index = cbo滴速 Then
        '快速和加压不显示单位
        If cboInfo(Index).ListIndex = 2 Or cboInfo(Index).ListIndex = 3 Then
            lblInfo(31).Visible = False
        Else
            lblInfo(31).Visible = True
        End If
    ElseIf Index = cbo单位 Then
        If Val(cboInfo(Index).Tag) = cboInfo(Index).ListIndex And cboInfo(Index).Tag <> "" Then Exit Sub
        cboInfo(Index).Tag = cboInfo(Index).ListIndex
        Call BloodSum
        Call RsetBreedUnit
    ElseIf Index = cbo输血血型 Then
        Call SetBloodLisAboRh(Index)
    ElseIf Index = cboRHD Then
        If cboInfo(Index).Text = "-" Then
            cboInfo(Index).ForeColor = vbRed
        Else
            cboInfo(Index).ForeColor = &H80000008
        End If
        Call SetLblRh
        Call SetBloodLisAboRh(Index)
    ElseIf Index = cbo用血安排 And mblnEditable = True Then
        intIdx = Val(GetBloodApplyCode(0))
        cboInfo(cbo输血血型).Enabled = intIdx = 0
        cboInfo(cboRHD).Enabled = intIdx = 0
        Call SetLblRh
    End If
    If Visible Then mblnChange = True
End Sub

Private Function FormatAdviceContext(ByVal strAdvicePro As String, ByVal strBloodWay As String) As String
'功能：根据系统基本参数，格式化医嘱内容
'参数：strBloodWay=输血途径,strAdvicePro=输血内容
    Dim strReturn As String, strText As String, strField As String
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    mrsDefine.Filter = "诊疗类别='K'"
    If mrsDefine.RecordCount > 0 Then
        strReturn = mrsDefine!医嘱内容 & ""
    End If
    If strReturn = "" Then
        If IsDate(txtInfo(txt预定输血时间).Text) Then
            strText = Format(txtInfo(txt预定输血时间).Text, "MM月dd日HH:mm")
        Else
            strText = Format(txtInfo(txt申请日期).Text, "MM月dd日HH:mm")
        End If
    
        strText = "于" & strText & "输" & strAdvicePro
        If strBloodWay <> "" Then
            strText = strText & "(" & strBloodWay & ")"
        End If
        strReturn = strText
    Else
        strText = strReturn
        If InStr(strText, "[输血时间]") > 0 Then
            If IsDate(txtInfo(txt预定输血时间).Text) Then
                strField = txtInfo(txt预定输血时间).Text
            Else
                strField = txtInfo(txt申请日期).Text
            End If
            strText = Replace(strText, "[输血时间]", """" & strField & """")
        End If
        If InStr(strText, "[诊疗项目]") > 0 Then
            strField = strAdvicePro
            strText = Replace(strText, "[诊疗项目]", """" & strField & """")
        End If
        If InStr(strText, "[输血项目]") > 0 Then
            strField = strAdvicePro
            strText = Replace(strText, "[输血项目]", """" & strField & """")
        End If
        If InStr(strText, "[输血途径]") > 0 Then
            strField = strBloodWay
            strText = Replace(strText, "[输血途径]", """" & strField & """")
        End If
        If InStr(strText, "[血型]") > 0 Then
            strField = Trim(cboInfo(cbo输血血型).Text)
            strText = Replace(strText, "[血型]", """" & strField & """")
        End If
        If InStr(strText, "[RH]") > 0 Then
            strField = Trim(cboInfo(cboRHD).Text)
            strText = Replace(strText, "[RH]", """" & strField & """")
        End If
        If InStr(strText, "[执行分类]") > 0 Then
            strField = IIF(mblnSpareBloood, 0, 1)
            strText = Replace(strText, "[执行分类]", """" & strField & """")
        End If
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

Private Function CheckUseBlood() As Boolean
'功能：启用了血库时，检查用血申请的量,是否超出了现有结果，只是进行提示，不强制禁止
    Dim lngRow As Long
    Dim dblTotal As Double   '待发总量
    Dim dblApplyTotal As Double  '申请总量
    Dim str申请单位 As String, lng换算系数 As Long, str替代项目 As String
    Dim strTmp As String, arrInfo, arrItem
    Dim j As Integer
    Dim strMsg As String
    Dim objCollection As New Collection
    
    If mblnSpareBloood = False And mblnSelectBlood = False Then
        With vsfBlood
            '获取每一个品种的待发血量
            For lngRow = .FixedRows To .Rows - 1
                strTmp = .TextMatrix(lngRow, COL_P_库存)
                dblTotal = 0: dblApplyTotal = 0
                If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                    '此处获取该品种的剩余量，统一转换为ML
                    lng换算系数 = Val(.TextMatrix(lngRow, COL_P_计算系数)): If lng换算系数 = 0 Then lng换算系数 = 1
                    str申请单位 = UCase(.TextMatrix(lngRow, COL_P_单位)): If str申请单位 = "" Then str申请单位 = "ML"
                    
                    dblApplyTotal = Val(.TextMatrix(lngRow, COL_P_申请量))
                    If str申请单位 <> "ML" Then
                        dblApplyTotal = dblApplyTotal * lng换算系数
                    End If
                End If
                If strTmp <> "" Then
                    arrInfo = Split(strTmp, "<Split2>")
                    If UBound(arrInfo) > 0 Then
                        arrItem = Split(arrInfo(1), "'")
                        dblTotal = Val(arrItem(2)) 'ML
                    End If
                End If
                '获取项目本身的剩余量
                dblTotal = dblTotal - dblApplyTotal
                objCollection.Add dblTotal, "A_" & .TextMatrix(lngRow, COL_P_ID)
            Next
            strMsg = ""
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                    '替代项目<Split2>品种ID'配发信息'待发量
                    '医嘱ID'医嘱内容'申请量'替代项目<Split2>品种ID'配发信息'待发量<Split3>品种ID'配发信息'待发量<Split3>品种ID'配发信息'待发量...<Split1>医嘱ID'医嘱内容'申请量<Split2>.....
                    strTmp = .TextMatrix(lngRow, COL_P_库存)
                    arrInfo = Split(strTmp, "<Split2>")
                    str替代项目 = arrInfo(0)
                    '不可能找不到,获取量
                    dblTotal = Val(objCollection("A_" & .TextMatrix(lngRow, COL_P_ID)))
                    If dblTotal < 0 Then
                        If str替代项目 <> "" Then
                            arrItem = Split(str替代项目, ",")
                            For j = 0 To UBound(arrItem)
                                If Val(arrItem(j)) > 0 Then
                                    dblApplyTotal = Val(ISExistCollection(objCollection, "A_" & Val(arrItem(j))))
                                    If dblApplyTotal > 0 And Val(arrItem(j)) <> Val(.TextMatrix(lngRow, COL_P_ID)) Then
                                        '剩余总量大于申请量，就退出，且从新更新替代项目的剩余量
                                        If dblApplyTotal >= dblTotal Then
                                            dblApplyTotal = dblApplyTotal - Abs(Val(dblTotal))
                                            dblTotal = 0
                                        Else
                                            dblApplyTotal = 0
                                            dblTotal = dblApplyTotal - Abs(Val(dblTotal))
                                        End If
                                        '更新集合
                                        objCollection.Remove "A_" & Val(arrItem(j))
                                        objCollection.Add dblApplyTotal, "A_" & Val(arrItem(j))
                                    End If
                                    If dblTotal >= 0 Then Exit For
                                End If
                            Next
                        End If
                        If dblTotal < 0 Then
                            strMsg = IIF(strMsg = "", "", strMsg & vbCrLf) & "[" & .TextMatrix(lngRow, COL_P_名称) & "]输入的申请量大于配血待发量。"
                        End If
                    End If
                End If
            Next
        End With
        
        If strMsg <> "" Then
            If MsgBox(strMsg & vbCrLf & "请问您是否要继续？。", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If vsfBlood.Enabled And vsfBlood.Visible Then vsfBlood.SetFocus
                Exit Function
            End If
        End If
    End If
    CheckUseBlood = True
End Function

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng执行性质 As Long
    Dim lng执行科室ID As Long
    Dim lng就诊ID As Long
    Dim bln中医 As Boolean
    Dim str类型 As String
    Dim blnSucceed As Boolean
    Dim strTmp As String
    Dim i As Long, intMax As Integer
    Dim strTabAdvice As String
    Dim strItems As String
    Dim blnCheck医保 As Boolean
    Dim rsPrice As ADODB.Recordset
    Dim lngRow As Long
    
    ' Call SeekNextControl  '用这种方式会出问题71290
    '这里采用两次设不同控件的焦点，确保validata事件的执行。
    txtGet(txt输血途径).SetFocus
    
    If txtInfo(txt诊断信息).Enabled = True And txtInfo(txt诊断信息).Locked = False Then
        If Trim(txtInfo(txt诊断信息).Text) = "" Then
            MsgBox "必须输入临床诊断！", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txtInfo(txt诊断信息))
            Exit Function
        End If
        intMax = txtInfo(txt诊断信息).MaxLength
        If LenB(StrConv(txtInfo(txt诊断信息).Text, vbFromUnicode)) > intMax Then
            MsgBox "临床诊断不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Call zlControl.ControlSetFocus(txtInfo(txt诊断信息))
            Exit Function
        End If
    End If
    '编辑附项不检查以下内容
    If mintType <> 3 Then
        '必须选择输血类型
        If cboInfo(cbo输血类型).Text = "" Then
            MsgBox "必须填写申请输血类型！", vbInformation, Me.Caption
            If cboInfo(cbo输血类型).Enabled Then cboInfo(cbo输血类型).SetFocus
            Exit Function
        End If
        
        '检查紧急医嘱必须填写输血目的
        If cboInfo(cbo用血安排).ListIndex = 1 And cboInfo(cbo输血目的).Text = "" Then
            MsgBox "紧急输血必须填写输血目的。", vbInformation, Me.Caption
            If cboInfo(cbo输血目的).Enabled Then cboInfo(cbo输血目的).SetFocus
            Exit Function
        End If
        
        '备血医嘱才检查
        If mblnSpareBloood = True Then
            '孕产情况检查
            If txtInfo(txt孕).Text <> "" And txtInfo(txt产).Text = "" Then
                MsgBox "输入了孕产情况中的孕次，则必须输入产次。", vbInformation, Me.Caption
                If txtInfo(txt产).Visible And txtInfo(txt产).Enabled Then txtInfo(txt产).SetFocus
                Exit Function
            End If
            If txtInfo(txt产).Text <> "" And txtInfo(txt孕).Text = "" Then
                MsgBox "输入了孕产情况中的产次，则必须输入孕次。", vbInformation, Me.Caption
                If txtInfo(txt孕).Visible And txtInfo(txt孕).Enabled Then txtInfo(txt孕).SetFocus
                Exit Function
            End If
            If Val(txtInfo(txt产).Text) > 0 Then
                If Val(txtInfo(txt孕).Text) = 0 Then
                    MsgBox "当孕产情况中的产次不为0时，则必须输入孕次，且次数必须大于0。", vbInformation, Me.Caption
                    If txtInfo(txt孕).Visible And txtInfo(txt孕).Enabled Then txtInfo(txt孕).SetFocus
                    Exit Function
                End If
            End If
            
            '输血治疗同意书和输血评估必须填写
            If optConsent(0).value = False And optConsent(1).value = False Then
                MsgBox "必须确认输血治疗同意书是否已签。", vbInformation, Me.Caption
                'option不能设置焦点，设置焦点会自动勾选
                Exit Function
            End If
            
            If optAppraise(0).value = False And optAppraise(1).value = False Then
                MsgBox "必须确定输血前评估是否已评估。", vbInformation, Me.Caption
                Exit Function
            End If
        End If
        '必须录入输血成分
        If mlng输血项目ID = 0 Then
            MsgBox "没有确定预定输血成分。", vbInformation, Me.Caption
            If vsfBlood.Visible And vsfBlood.Enabled Then
                If vsfBlood.Row > vsfBlood.FixedRows Then
                    vsfBlood.Row = vsfBlood.FixedRows
                    vsfBlood.Col = COL_P_选择
                End If
                vsfBlood.SetFocus
            End If
            Exit Function
        End If
        
        '检查执行科室
        If cboInfo(cbo执行科室).Text = "" Then
            MsgBox "没有确定执行科室。", vbInformation, Me.Caption
            If cboInfo(cbo执行科室).Enabled Then cboInfo(cbo执行科室).SetFocus
            Exit Function
        End If
        
        '检查滴速度
        If cboInfo(cbo滴速).Visible = True Then '可见说明肯定是用血申请
            If cboInfo(cbo滴速).ListIndex < 0 Then
                If LenB(StrConv(cboInfo(cbo滴速).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cboInfo(cbo滴速).Text) And cboInfo(cbo滴速).Text <> "") Then
                    MsgBox "自由录入的滴数只能是数字，且最多只允许录入3位数字！", vbInformation, gstrSysName
                    Call zlControl.ControlSetFocus(cboInfo(cbo滴速))
                    Exit Function
                End If
            End If
        End If
        
        '检查输血途径和输血执行
        If mlng输血途径 = 0 Then
            If mblnNewSpareBloood = False Then
                MsgBox "没有指定输血途径。", vbInformation, Me.Caption
            Else
                MsgBox "没有指定采集方式。", vbInformation, Me.Caption
            End If
            If txtGet(txt输血途径).Enabled Then txtGet(txt输血途径).SetFocus
            Exit Function
        End If
        
        If cboInfo(cbo输血执行).Text = "" Then
            If mblnNewSpareBloood = False Then
                MsgBox "没有确定输血执行科室。", vbInformation, Me.Caption
            Else
                MsgBox "没有确定采集执行科室。", vbInformation, Me.Caption
            End If
            If cboInfo(cbo输血执行).Enabled Then cboInfo(cbo输血执行).SetFocus
            Exit Function
        End If
        
        '必须录入总量
        If cboInfo(cbo单位).ListIndex = -1 Then
            MsgBox "请确定输血总量单位！", vbInformation, Me.Caption
            If cboInfo(cbo单位).Enabled And cboInfo(cbo单位).Visible Then cboInfo(cbo单位).SetFocus
            Exit Function
        End If
        
        '检查时间合法性
        If Not Check开始时间(txtInfo(txt申请日期).Text) Then
            If txtInfo(txt申请日期).Enabled Then txtInfo(txt申请日期).SetFocus
            Exit Function
        End If
        If Not Check安排时间(txtInfo(txt预定输血时间).Text, txtInfo(txt申请日期).Text) Then
            If txtInfo(txt预定输血时间).Enabled Then txtInfo(txt预定输血时间).SetFocus
            Exit Function
        End If
        
        With vsfBlood
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                    If Val(.TextMatrix(lngRow, COL_P_申请量)) <= 0 Then
                        If mblnSelectBlood = False Then
                            MsgBox "请录入大于0的输血申请量。", vbInformation, Me.Caption
                            .Row = lngRow: .Col = COL_P_申请量
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                        Else
                            MsgBox "[" & .TextMatrix(lngRow, COL_P_名称) & "]还未选择血液信息。", vbInformation, Me.Caption
                            .Row = lngRow: .Col = COL_P_名称
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                        End If
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngRow, COL_P_申请量)) > Val(.TextMatrix(lngRow, COL_P_录入限量ID)) And Val(.TextMatrix(lngRow, COL_P_录入限量ID)) > 0 Then
                        If MsgBox(.TextMatrix(lngRow, COL_P_名称) & " 的总量:" & Val(.TextMatrix(lngRow, COL_P_申请量)) & .TextMatrix(lngRow, COL_P_单位) & " 超过允许录入的最大限量:" & _
                            Val(.TextMatrix(lngRow, COL_P_录入限量ID)) & .TextMatrix(lngRow, COL_P_单位) & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            .Row = lngRow: .Col = COL_P_申请量
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
        
        If Trim(cboInfo(cbo输血血型).Text) = "" And Trim(cboInfo(cboRHD).Text) = "" Then
            If MsgBox("没有确定血型和RH(D)请问您是否要继续？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cbo输血血型).Enabled Then cboInfo(cbo输血血型).SetFocus
                Exit Function
            End If
        ElseIf Trim(cboInfo(cbo输血血型).Text) = "" Then
            If MsgBox("没有确定血型请问您是否要继续？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cbo输血血型).Enabled Then cboInfo(cbo输血血型).SetFocus
                Exit Function
            End If
        ElseIf Trim(cboInfo(cboRHD).Text) = "" Then
            If MsgBox("没有确定RH(D)请问您是否要继续？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cboRHD).Enabled Then cboInfo(cboRHD).SetFocus
                Exit Function
            End If
        End If
        If BloodApplyCheck = False Then Exit Function '调用自定义过程，由医院自行对整个申请单进行检查
        If CheckOrResetLisAboRH = False Then Exit Function '检查LIS结果中的血型是否和选择血型是否一致
        If CheckUseBlood = False Then Exit Function '用血申请检查申请量是否大于配发量（进行提示）
        
        If mint调用场合 = 0 Then
            lng就诊ID = IIF(mint场合 = 1, mlng挂号ID, mlng主页ID)
            strTmp = mlng输血项目ID & "||" & IIF(mint场合 = 1, 1, 2)
            mstr摘要输血 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", strTmp)
            strTmp = mlng输血途径 & "||" & IIF(mint场合 = 1, 1, 2)
            mstr摘要途径 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", strTmp)
            strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
            For i = 1 To 2
                If i = 1 Then
                    lng执行性质 = IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, 5, mlng执行科室性质)
                    lng执行科室ID = IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, 0, cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex))
                    
                    strTabAdvice = "select 1 as ID,1 as 序号,-null as 相关ID,'K' as 诊疗类别," & mlng输血项目ID & " as 管码项目ID," & _
                            mlng输血项目ID & " as 诊疗项目ID," & Val(txtInfo(txt预定输血量).Text) & " As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
                            "0 as 执行标记,0 as 计价特性, null As 附加手术," & lng执行性质 & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
                    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", IIF(1 = mint场合, 1, 2), mlng病人ID, lng就诊ID, mint险类, 1, _
                         "K", mlng输血项目ID, mlng开单科室ID, UserInfo.姓名, lng执行科室ID, lng执行性质, 0, 0, mstr摘要输血)
                Else
                    lng执行性质 = IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, 5, mlng输血执行性质)
                    lng执行科室ID = IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, 0, cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex))
                    
                    strTabAdvice = strTabAdvice & " Union All " & _
                         "select 2 as ID,2 as 序号,1 as 相关ID,'E' as 诊疗类别," & mlng输血途径 & " as 管码项目ID," & _
                            mlng输血途径 & " as 诊疗项目ID,1 As 总量, 0 As 单量,null as 标本部位,null As 检查方法," & _
                            "0 as 执行标记,0 as 计价特性, null As 附加手术," & lng执行性质 & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"

                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", IIF(1 = mint场合, 1, 2), mlng病人ID, lng就诊ID, mint险类, 1, _
                         "E", mlng输血途径, mlng开单科室ID, UserInfo.姓名, lng执行科室ID, lng执行性质, 0, 0, mstr摘要途径)
                End If
                
                If Not rsTmp.EOF Then
                    strMsg = NVL(rsTmp!结果)
                    If strMsg <> "" Then
                        Select Case Val(Split(strMsg, "|")(0))
                        Case 1 '提示
                            If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                strMsg = "": Exit Function
                            End If
                        Case 2 '禁止
                            MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                            strMsg = "": Exit Function
                        End Select
                        strMsg = ""
                    End If
                End If
            Next
            
            '诊断检查
             If InStr(mstr检查入院诊断, "K") > 0 And mint场合 = 0 Then
                bln中医 = Sys.DeptHaveProperty(mlng病人科室id, "中医科")
                str类型 = IIF(bln中医, "2,12", "2")
                If Not ExistsDiagNoses(mlng病人ID, mlng主页ID, str类型) Then
                    strMsg = "病人的入院诊断还没有输入，请先输入病人的入院诊断再下达输血申请。"
                End If
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '对码检查
            With vsfBlood
                For lngRow = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                        strIDs = IIF(strIDs = "", "", strIDs & ",") & Val(.TextMatrix(lngRow, COL_P_ID)) & ":"
                        If Val(cboInfo(cbo执行科室).Tag & "") <> 0 Then
                            strIDs = strIDs & Val(cboInfo(cbo执行科室).Tag & "")
                        End If
                    End If
                Next
            End With
            str医嘱内容 = FormatAdviceContext(Replace(txtGet(txt预定输血成分).Text, "'", ","), txtGet(txt输血途径).Text)
            
            strIDs = strIDs & "," & mlng输血途径 & ":"
            If Val(cboInfo(cbo输血执行).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo输血执行).Tag & "")
            End If
            If gint医保对码 = 2 Then mbln提醒对码 = True
            strItems = strIDs
            strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, IIF(mlng病人性质 = 0, 2, 1), "", strIDs, str医嘱内容)
            If strMsg <> "" Then
                If gint医保对码 = 1 Then
                    vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                    If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                    If vMsg = vbIgnore Then mbln提醒对码 = False
                ElseIf gint医保对码 = 2 Then
                    MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '医保管控实时监测
            If mint险类 <> 0 Then
                If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                    If MakePriceRecord申请单("3" & IIF(mint场合 = 1, "1", "2"), mlng病人ID, lng就诊ID, strTabAdvice, strItems, mstr费别, mlng开单科室ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint险类, IIF(mint场合 = 1, 0, 1), 0, rsPrice) Then
                            MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的输血申请单不能保存。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    CheckData = True
End Function

Private Function SaveData() As Boolean
'功能：保存数据
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lng医嘱ID As Long, lng医嘱序号 As Long, lng申请序号 As Long
    Dim strSQL As String, rsTmp As Recordset
    Dim rsData As New ADODB.Recordset, rsTemp As Recordset
    Dim str项目名称 As String, str输血途径 As String, strPrivs As String
    Dim curDate As Date, i As Long, lng相关ID As String, j As Long
    Dim lngCount As Long, int病人来源 As Integer
    Dim strTmp主页ID As String
    Dim strTmp挂号单 As String
    Dim str审核状态 As String
    Dim int紧急 As Integer
    Dim int分类 As Integer
    Dim int检查方法 As Integer
    Dim str滴速 As String
    Dim strErr As String
    Dim bln已审核 As Boolean
    Dim dbl24h量 As Double, dblTmp As Double
        
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If cboInfo(cbo滴速).Visible = True Then
        str滴速 = cboInfo(cbo滴速).Text
    End If
    If IsNumeric(str滴速) = True Then
        str滴速 = str滴速 & "滴/分钟"
    End If
    
    If mintType = 3 Then
        '申请附项编辑模式
        lng相关ID = mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Delete(" & lng相关ID & ")"
    ElseIf mintType = 4 Or (mintType = 1 And mblnUseBloodSend = True) Then '核对医嘱或修改用血医嘱
        '检查用血医嘱状态是已经审核状态(已经审核则说明本次就是审核再次修改)
        If mintType = 1 And mblnUseBloodSend = True Then
            gstrSQL = "Select 操作时间 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "检查用血医嘱是否审核", mlngUpdateAdvice, 11)
            bln已审核 = rsData.RecordCount > 0
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_医嘱审核管理_Cancel('" & mlngUpdateAdvice & "')"
        End If
        
        gstrSQL = "Select Id, 相关id, 序号, 医嘱状态, 医嘱期效, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性," & vbNewLine & _
            "       执行科室id, 执行性质, 紧急标志, 开始执行时间, 执行终止时间, 病人科室id, 开嘱科室id, 开嘱医生, 开嘱时间, 检查方法, 执行标记, 可否分零, 摘要, 零费记帐, 用药目的, 用药理由, 审核状态," & vbNewLine & _
            "       超量说明, 首次用量, 手术情况, 组合项目id, 皮试结果" & vbNewLine & _
            "From 病人医嘱记录" & vbNewLine & _
            "Where Id = [1] Or 相关id = [1]" & vbNewLine & _
            "Order By Nvl(相关id, 0)"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取待审核医嘱信息", mlngUpdateAdvice)
        
        str项目名称 = ""
        For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
            If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
                str项目名称 = IIF(str项目名称 = "", "", str项目名称 & ",") & vsfBlood.TextMatrix(i, COL_P_名称)
            End If
        Next
        If str项目名称 = "" Then str项目名称 = Sys.RowValue("诊疗项目目录", mlng输血项目ID, "名称")
         
        Set rsTmp = Get诊疗项目记录(mlng输血途径)
        str输血途径 = rsTmp!名称 & ""
        
        rsData.Filter = "ID=" & mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱记录_Update(" & rsData!ID & ",NULL," & rsData!序号 & ",1,1," & mlng输血项目ID & _
                                 "," & ZVal(NVL(rsData!收费细目ID, 0)) & "," & ZVal(NVL(rsData!天数, 0)) & "," & ZVal(NVL(rsData!单次用量, 0)) & "," & ZVal(txtInfo(txt预定输血量).Text) & ",'" & FormatAdviceContext(str项目名称, str输血途径) & _
                                 "'," & IIF(txtInfo(txt备注).Text = "", "NULL", "'" & txtInfo(txt备注).Text & "'") & ",'" & Format(txtInfo(txt预定输血时间).Text, "yyyy-MM-dd HH:mm:ss") & "','一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "Null", cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "5", mlng执行科室性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                   ZVal(NVL(rsData!检查方法, 0)) & ",0,NULL," & IIF(mstr摘要输血 = "", "null", "'" & mstr摘要输血 & "'") & ",'" & UserInfo.姓名 & "',Null,Null,'" & cboInfo(cbo输血目的).Text & "'," & IIF(bln已审核 = True, 1, ZVal(NVL(rsData!审核状态, 0))) & ")"
        
        lng相关ID = mlngUpdateAdvice
        rsData.Filter = "相关id=" & mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱记录_Update(" & rsData!ID & "," & lng相关ID & "," & rsData!序号 & ",1,1," & mlng输血途径 & ",NULL,NULL,NULL,Null,'" & str输血途径 & "'," & IIF(str滴速 = "", "NULL", "'" & str滴速 & "'") & ",NULL,'一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "Null", cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "5", mlng输血执行性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                ZVal(NVL(rsData!检查方法, 0)) & ",0,NULL," & IIF(mstr摘要途径 = "", "null", "'" & mstr摘要途径 & "'") & ",'" & UserInfo.姓名 & "',Null,NULL,''," & IIF(bln已审核 = True, 1, ZVal(NVL(rsData!审核状态, 0))) & ")"
        
        If mintType = 4 Or bln已审核 = True Then
            '完成核对
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_医嘱审核管理_Audit(" & lng相关ID & "," & 1 & "," & _
                            "'" & UserInfo.姓名 & "',to_date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
        End If
                        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Delete(" & lng相关ID & ")"
        str审核状态 = "NULL"
    Else
        
        lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
        
        '病人医嘱记录.序号，递增
        If mint场合 = 0 Then
            lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, mbytBaby) + 1
            strTmp主页ID = mlng主页ID
            strTmp挂号单 = "NULL"
            int病人来源 = 2
        Else
            lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, , mbytBaby, mstr挂号单) + 1
            strTmp主页ID = "NULL"
            strTmp挂号单 = "'" & mstr挂号单 & "'"
            int病人来源 = 1
        End If
        
        str项目名称 = ""
        For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
            If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
                str项目名称 = IIF(str项目名称 = "", "", str项目名称 & ",") & vsfBlood.TextMatrix(i, COL_P_名称)
            End If
        Next
        If str项目名称 = "" Then str项目名称 = Sys.RowValue("诊疗项目目录", mlng输血项目ID, "名称")
        
        Set rsTmp = Get诊疗项目记录(mlng输血途径)
        str输血途径 = rsTmp!名称 & "" ' Get项目名称(mlng输血途径)
        int分类 = Val(rsTmp!执行分类 & "")
        If mlngUpdateAdvice <> 0 Then
            '取申请序号
            strSQL = "Select 申请序号,检查方法 From 病人医嘱记录 where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            lng申请序号 = Val(rsTmp!申请序号 & "")
            int检查方法 = Val(rsTmp!检查方法 & "")
            
            '修改医嘱，删除后重新插入
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & mlngUpdateAdvice & ",1)"
        Else
            If mblnSpareBloood = True Then
                int检查方法 = 0
            Else
                int检查方法 = 1
            End If
        End If
        If lng申请序号 = 0 Then
            '取申请序号
            strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            lng申请序号 = Val(rsTmp!申请序号 & "")
        End If
        
        int紧急 = IIF(cboInfo(cbo用血安排).ListIndex <> 1, 0, 1)
        dblTmp = GetBloodTotalByML
        str审核状态 = GetBloodVerifyState(int病人来源, mlng病人ID, IIF(int病人来源 = 2, mlng主页ID, mlng挂号ID), txtInfo(txt预定输血时间).Text, dblTmp, int紧急, int分类, CInt(mbytBaby), mlngUpdateAdvice)
        If str审核状态 = "" Then str审核状态 = "NULL"
        '输血医嘱
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng医嘱ID & ",NULL," & lng医嘱序号 & "," & int病人来源 & "," & mlng病人ID & "," & strTmp主页ID & "," & mbytBaby & ",1,1,'K'," & mlng输血项目ID & _
                                 ",NULL,NULL,NULL," & ZVal(txtInfo(txt预定输血量).Text) & ",'" & FormatAdviceContext(str项目名称, str输血途径) & _
                                 "'," & IIF(txtInfo(txt备注).Text = "", "NULL", "'" & txtInfo(txt备注).Text & "'") & ",'" & Format(txtInfo(txt预定输血时间).Text, "yyyy-MM-dd HH:mm:ss") & "','一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "Null", cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "5", mlng执行科室性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp挂号单 & "," & ZVal(mlng前提ID) & "," & IIF(int检查方法 = 0, "NULL", "'" & int检查方法 & "'") & ",0,NULL," & IIF(mstr摘要输血 = "", "null", "'" & mstr摘要输血 & "'") & ",'" & UserInfo.姓名 & "',Null,Null,'" & cboInfo(cbo输血目的).Text & "'," & str审核状态 & "," & lng申请序号 & ")"
        
        '输血途径
        lng相关ID = lng医嘱ID
        lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
        lng医嘱序号 = lng医嘱序号 + 1
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng医嘱ID & "," & lng相关ID & "," & lng医嘱序号 & "," & int病人来源 & "," & mlng病人ID & "," & strTmp主页ID & _
                                 "," & mbytBaby & ",1,1,'E'," & mlng输血途径 & ",NULL,NULL,NULL,Null,'" & str输血途径 & "'," & IIF(str滴速 = "", "NULL", "'" & str滴速 & "'") & ",NULL,'一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "Null", cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "5", mlng输血执行性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp挂号单 & "," & ZVal(mlng前提ID) & ",NULL,0,NULL," & IIF(mstr摘要途径 = "", "null", "'" & mstr摘要途径 & "'") & ",'" & UserInfo.姓名 & "',Null,NULL,''," & str审核状态 & "," & lng申请序号 & ")"
    End If
    '输血申请其他项目
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_输血申请记录_Insert(" & lng相关ID & "," & chkWait.value & ",'" & cboInfo(cbo输血类型).Text & "','" & cboInfo(cbo输血目的).Text & "'," & cboInfo(cbo输血性质).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optHistory(2).value, 0, 1) & "," & IIF(optHistory(4).value, 0, 1) & ",'" & txtInfo(txt孕) & "/" & txtInfo(txt产) & "'," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo输血血型).ListIndex & "," & cboInfo(cboRHD).ListIndex & "," & IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, "Null")) & "," & IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, "Null")) & ",'" & GetBloodInfo & "')"
    '检验项目
    With vsLIS
        lngCount = 0
        For i = 0 To .Rows - 1
            For j = 0 To CON_LisResultCol - 1
                If Val(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) <> 0 Then
                    lngCount = lngCount + 1
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_输血检验结果_Insert(" & lng相关ID & "," & lngCount & "," & ZVal(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) & ",'" & .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_指标英文名 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_结果单位 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_结果标志 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_结果参考 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_取值序列 + (j * CON_LisResultCount)) & "'," & IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0) & ")"
                End If
            Next
        Next
    End With
    
    '诊断关联信息
    If mstr诊断IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng相关ID & ",'" & mstr诊断IDs & "')"
    End If
    If Trim(txtInfo(txt诊断信息).Text) <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng相关ID & ",'申请单诊断',null,1,null,'" & txtInfo(txt诊断信息).Text & "')"
    End If
    '申请内容插入医嘱申请附加项目
    str项目名称 = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
            str项目名称 = IIF(str项目名称 = "", "", str项目名称 & Space(2)) & vsfBlood.TextMatrix(i, COL_P_名称) & ":" & IIF(vsfBlood.TextMatrix(i, COL_P_申请血型) = "", "", vsfBlood.TextMatrix(i, COL_P_申请血型) & vsfBlood.TextMatrix(i, COL_P_申请RH)) & " " & vsfBlood.TextMatrix(i, COL_P_申请量) & vsfBlood.TextMatrix(i, COL_P_单位)
        End If
    Next
    If str项目名称 <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lng相关ID & ",'申请项目',null,2,null,'" & str项目名称 & "')"
    End If
        
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    '用血医嘱修改，先删除在新增
    If int检查方法 = 1 And InStr(1, ",3,4,", "," & mintType & ",") = 0 And mlngUpdateAdvice <> 0 And Not (mintType = 1 And mblnUseBloodSend = True) Then
        If InitObjBlood = True Then
            If gobjPublicBlood.AdviceOperation(IIF(mint场合 = 0, p住院医嘱下达, p门诊医嘱下达), mlngUpdateAdvice, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "血库公共部件调用失败，详细信息：" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "血库公共部件创建失败，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '医嘱相关过程执行
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '用血医嘱新增
    If int检查方法 = 1 And InStr(1, ",3,4,", "," & mintType & ",") = 0 And Not (mintType = 1 And mblnUseBloodSend = True) Then
        If InitObjBlood = True Then
            If gobjPublicBlood.AdviceOperation(IIF(mint场合 = 0, p住院医嘱下达, p门诊医嘱下达), lng相关ID, 0, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "血库公共部件调用失败，详细信息：" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "血库公共部件创建失败，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False

    mlngUpdateAdvice = lng相关ID
    
    Call SetCommandBarPara(conMenu_Tool_Archive, IIF(mblnSpareBloood = True, 1, 2), mlngUpdateAdvice)
    
    If mint场合 = 0 Then
        If str审核状态 = "NULL" Or str审核状态 = "4" Then
            Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, Trim(txtInfo(txt姓名).Text), Trim(txtInfo(txt住院号).Text), , IIF(mlng病人性质 = 1, 1, 2), _
                mlng主页ID, mlng病区ID, , mlng病人科室id, "", , Trim(txtInfo(txt床号).Text), lng相关ID, int紧急, 1, "K", "", UserInfo.姓名, _
                Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , , "")
        ElseIf str审核状态 = "1" Then
            Call ZLHIS_CIS_Audit("ZLHIS_CIS_030", mclsMipModule, mlng病人ID, Trim(txtInfo(txt姓名).Text), Trim(txtInfo(txt住院号).Text), , IIF(mlng病人性质 = 1, 1, 2), _
                mlng主页ID, mlng病区ID, , mlng病人科室id, "", , Trim(txtInfo(txt床号).Text), lng相关ID, UserInfo.姓名, _
                Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , , "")
        End If
    End If
    
    SaveData = True
    mblnChange = False
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboInfo_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If Index = cboRHD Then
        Call cboInfo_Click(Index)
    End If
End Sub

Private Sub cboInfo_DropDown(Index As Integer)
    If Index = cboRHD Then
        cboInfo(Index).ForeColor = &H80000008
    End If
End Sub

Private Sub cboInfo_GotFocus(Index As Integer)
    If Index = cbo滴速 Then Call zlControl.TxtSelAll(cboInfo(Index))
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
    If Index = cbo输血目的 Then
        If zlCommFun.ActualLen(cboInfo(Index).Text) > 50 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
    ElseIf Index = cbo滴速 And KeyAscii <> vbKeyReturn And KeyAscii <> 8 And KeyAscii <> vbKeyTab Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        If Chr(KeyAscii) = 0 Then
            If cboInfo(Index).Text = "" Or cboInfo(Index).SelLength = Len(cboInfo(Index).Text) Or cboInfo(Index).SelStart = 0 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
        If cboInfo(Index).SelLength = 0 And Len(cboInfo(Index).Text) > 2 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub PrintApply(ByVal intType As Integer)
'功能打印预览申请单
'参数：intType:1-预览，2-打印
    '判断如果还未保存则先保存再打印
    Dim strReportName As String
    If mintType <> 2 Then
        If mblnChange Then
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Else
            '如果不可用，则检查医嘱是否符合
            If CheckData = False Then Exit Sub
        End If
    End If
    If BloodApplyPrintCheck(mlngUpdateAdvice, IIF(1 = mint场合, 1, 2), IIF(mblnSpareBloood = True, 1, 2), intType - 1) = False Then Exit Sub
    strReportName = IIF(mblnSpareBloood = False, "ZL1_INSIDE_1254_17_2", "ZL1_INSIDE_1254_17_1")
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportName, Me, "医嘱ID=" & mlngUpdateAdvice, intType)
End Sub

Private Sub cboInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = cbo滴速 Then Call cboInfo_Click(Index)
End Sub

Private Sub cboInfo_LostFocus(Index As Integer)
    If Index = cboRHD Then Call cboInfo_Click(Index)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnChange As Boolean
    Dim lngPreUpdateAdivice As Long
    Select Case Control.ID
        Case conMenu_Tool_Archive * 10# + 1, conMenu_Tool_Archive * 10# + 2
            '只有mint调用场合=0和新增模式才允许切换，所以保存改变直接调用SaveData
            Me.Tag = ""
            blnChange = mblnChange
            lngPreUpdateAdivice = mlngUpdateAdvice
            If blnChange = True Then
                If MsgBox("当前申请单已经进行了调整尚未保存，请问是否需要保存？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes Then
                    '保存
                    If CheckData = False Then Exit Sub
                    mblnOK = SaveData
                    If mblnOK = False Then Exit Sub
                End If
                mblnChange = False
            End If
            
            mblnSpareBloood = Control.ID = (conMenu_Tool_Archive * 10# + 1)
            mblnNewSpareBloood = mblnSpareBloood And mintType = 0
            'Control.Category 切换后页面的申请ID,lngPreUpdateAdivice 切换前页面的申请ID(保存时存储Control.Category值)
            '总则：允许切换页面的情况下，只有两个页面都是起初的新增状态，且没有改变过，就只是调整页面控件位置,负责就重新加载控件和刷新数据
            If Not (lngPreUpdateAdivice = 0 And Val(Control.Category) = 0 And blnChange = False) Then
                '重新加载页面内容
                mlngUpdateAdvice = Val(Control.Category)
                Call SetControlContent
                If mlngUpdateAdvice <> 0 Then Me.Tag = "GOTO"
                Call Form_Load
                Me.Tag = ""
            Else
                Call SetControlContent(False)
                Call SetFormNature(False)
                Call LoadLastPrepareBlood
            End If
            mblnChange = False
        Case conMenu_File_PrintSet:
            If mblnSpareBloood = False Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_17_2", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_17_1", Me)
            End If
        Case conMenu_File_Preview, conMenu_File_Print
            Call PrintApply(IIF(Control.ID = conMenu_File_Preview, 1, 2))
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '保存
            If CheckData = False Then Exit Sub
            If mint调用场合 = 0 Then
                mblnOK = SaveData
            Else
                mblnOK = SaveCacheData
            End If
            If Control.ID = conMenu_Edit_SaveExit Then
                Unload Me
            Else
                lblInfo(lbl24H输血量) = "24小时内输血申请量：" & GetBloodCapacity(IIF(mint场合 = 0, 2, 1), mlng病人ID, IIF(mint场合 = 0, mlng主页ID, mlng挂号ID), zlDatabase.Currentdate, True, CInt(mbytBaby)) & "ML"
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Tool_Archive, conMenu_Tool_Archive * 10# + 1, conMenu_Tool_Archive * 10# + 2
            Control.Enabled = mintType = 0
            If Not mrsCard Is Nothing Then
                If Not mrsCard.EOF Then
                    Control.Enabled = False
                End If
            End If
            If Control.ID = conMenu_Tool_Archive * 10# + 1 Then
                Control.Checked = mblnSpareBloood
            ElseIf Control.ID = conMenu_Tool_Archive * 10# + 2 Then
                Control.Checked = Not mblnSpareBloood
            Else
                Control.Caption = IIF(mblnSpareBloood = True, "输血申请单▼", "取血通知单▼")
                Control.Checked = Control.Enabled
            End If
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
        Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview
            blnVisible = ((mint申请单打印模式 = 0 And InStr(GetInsidePrivs(p住院医嘱下达), ";输血申请单;") > 0) Or mint场合 = 1) And mint调用场合 = 0
            If mint申请单打印模式 = 0 And mint场合 = 0 Then
                If mintPState = ps出院 Then blnVisible = False
            End If
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub chkWait_Click()
    If chkWait.value = 1 Then
        txtInfo(txt诊断信息).Text = "待诊"
        txtInfo(txt诊断信息).Locked = True
        cmdInfo.Enabled = False
        mstr诊断IDs = ""
    Else
        txtInfo(txt诊断信息).Text = ""
        txtInfo(txt诊断信息).Locked = False
        cmdInfo.Enabled = True
    End If
    txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Dim lngIndex As Long
    
    If Index = 0 Then
        lngIndex = txt预定输血时间
    ElseIf Index = 1 Then
        lngIndex = txt申请日期
    End If
    If IsDate(txtInfo(lngIndex).Text) Then
        dtpDate.value = CDate(txtInfo(lngIndex).Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = lngIndex
    dtpDate.Left = txtInfo(lngIndex).Left + txtInfo(lngIndex).Width - dtpDate.Width
    dtpDate.Top = txtInfo(lngIndex).Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.ZOrder 0
    dtpDate.SetFocus
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Call TxtGetInfo(Index, 1)
End Sub

Private Sub cmdInfo_Click()
    Dim str诊断 As String
    Dim lng就诊ID As Long
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng病人性质 = 1, 1260, 1261), mclsMipModule)
    End If
    lng就诊ID = IIF(mint场合 = 0, mlng主页ID, mlng挂号ID)
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng病人ID, lng就诊ID, IIF(mlng病人性质 = 1, 1, 2), mlng病人科室id, mstr诊断IDs, str诊断, 0, mlngUpdateAdvice) Then
        txtInfo(txt诊断信息).Text = str诊断
        txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
        If mstr诊断IDs <> "" And chkWait.value = 1 Then
            chkWait.value = 0
        End If
    End If
    Call SeekNextControl
End Sub

Private Sub cmdInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdInfo_Click
End Sub

Private Function Check开始时间(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法
'说明：
'1.开始时间不能小于病人的入院时间
'2.开始时间不能小于病人的转科时间
'3.开始时间必须小于终止时间
'4.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
'5.补录的医嘱开始时间不能大于当前时间，转科补录不能大于转科开始时间
    Dim strInDate As String, blnOut As Boolean
    Dim rsBlood As New ADODB.Recordset
        
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
    strInDate = mstr入院时间
    '住院场合调用时才做以下检查
    If mint场合 = 0 Then
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "医嘱的开始执行时间不能小于病人的入院时间 " & strInDate & " 。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    
    
        strInDate = ""
        If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            If mdatTurn <> CDate(0) Then strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
        ElseIf IsDate(mstr上次转科时间) Then
            strInDate = mstr上次转科时间
        End If
    
        If strInDate <> "" Then
            If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
                If Format(strStart, "yyyy-MM-dd HH:mm") >= strInDate Then
                    strMsg = "医嘱的开始执行时间应小于病人" & IIF(mintPState = ps最近转出, "转出", IIF(mintPState = ps预出, "预出院", "出院")) & "的时间 " & strInDate & " 。"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                    strMsg = "医嘱的开始执行时间不能小于病人最近的转科时间 " & strInDate & " 。"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "医嘱的开始执行时间不能小于病人的就诊时间 " & strInDate & " 。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, ",1,4,", "," & mintType & ",") = 0 Then
        If InitObjBlood = True Then
            If gobjPublicBlood.GetPrepareBloodRs(mlngUpdateAdvice, rsBlood) = True Then
                '用血医嘱已经发血，则申请时间不能大于发血时间
                If Val(rsBlood!记录性质 & "") = 2 And Val(rsBlood!记录状态 & "") = 1 And IsDate(rsBlood!完成时间 & "") Then
                    If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(rsBlood!完成时间, "yyyy-MM-dd HH:mm") Then
                        strMsg = "医嘱的开始执行时间应小于血库发血的时间" & Format(rsBlood!完成时间, "yyyy-MM-dd HH:mm") & " 。"
                        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                '备血医嘱如果已经接收，则申请事件不能大于接收时间(主要是老数据)
                ElseIf Val(rsBlood!记录性质 & "") = 1 And Val(rsBlood!记录状态 & "") = 1 And IsDate(rsBlood!接收时间 & "") Then
                    If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(rsBlood!接收时间, "yyyy-MM-dd HH:mm") Then
                        strMsg = "医嘱的开始执行时间应小于血库接收的时间" & Format(rsBlood!接收时间, "yyyy-MM-dd HH:mm") & " 。"
                        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    Check开始时间 = True
End Function

Private Function Check安排时间(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的输血时间是否合法
'说明：
'1.输血时间不能小于医嘱的开始时间
    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "输入的输血时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "输血时间不能小于申请时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check安排时间 = True
End Function

Private Sub cmd常用嘱托_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    
    If Trim(txtInfo(txt备注).Text) = "" Then
        MsgBox "请输入嘱托内容。", vbInformation, gstrSysName
        If txtInfo(txt备注).Enabled Then txtInfo(txt备注).SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From 常用嘱托 Where 名称=[1] And (人员=[2] Or 人员 is null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtInfo(txt备注).Text), UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该嘱托内容已经在常用嘱托中。", vbInformation, gstrSysName
        If txtInfo(txt备注).Enabled Then txtInfo(txt备注).SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(txtInfo(txt备注).Text, CByte(Val(zlDatabase.GetPara("简码方式"))))
    strSQL = "zl_常用嘱托_Insert('" & Replace(txtInfo(txt备注).Text, "'", "''") & "','" & strSQL & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    MsgBox "已设置为常用嘱托。", vbInformation, gstrSysName
    If txtInfo(txt备注).Enabled Then txtInfo(txt备注).SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd医生嘱托_Click()
    Call ReasonSelect
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String, intIndex As Integer
    
    intIndex = Val(dtpDate.Tag)
    If intIndex = txt申请日期 Then
        '取值
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check开始时间(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '更新数据
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    ElseIf intIndex = txt预定输血时间 Then
        '取值
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check安排时间(strDate, txtInfo(txt申请日期).Text) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '更新数据
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    Dim intIndex As Integer
    
    If KeyAscii = vbKeyEscape Then
        intIndex = Val(dtpDate.Tag)
        If intIndex >= 0 Then txtInfo(intIndex).SetFocus
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub dtpDate_LostFocus()
    If Me.ActiveControl.Name <> "cmdDate" Then
        dtpDate.Visible = False
    Else
        If dtpDate.Visible = True Then
            Call cmdDate_Click(Me.ActiveControl.Index)
        End If
    End If
End Sub


Private Sub Form_Load()
    Dim strPar As String, i As Integer
    Dim rsBlood As New ADODB.Recordset
    
    mblnDataLoad = False
    mblnHaveAuditPriv = HaveAuditPriv
    mblnEditable = True
    mblnUseBloodSend = False
    mstr诊断IDs = ""
    mstrLISAboRHCode = ""
    If Me.Tag <> "GOTO" Then mblnOK = False
    mbln提醒对码 = True
    vsLIS.Rows = 0
    If mint场合 = 0 Then mint申请单打印模式 = Val(zlDatabase.GetPara("输血申请单打印模式", glngSys, p住院医嘱发送, "1"))
    
    strPar = zlDatabase.GetPara("输血申请注意事项", glngSys, IIF(mint场合 = 0, p住院医嘱下达, p门诊医嘱下达), "")
    lblInfo(lbl注意).Caption = Trim(strPar)
    lblInfo(lbl注意).Visible = Trim(strPar) <> ""
    
    '勾选了启用血库且安装了血库
    mblnNewSpareBloood = mblnSpareBloood And mintType = 0
    
    If mobjPublicLis Is Nothing Then
        On Error Resume Next
        Set mobjPublicLis = CreateObject("zlPublicLIS.clsSampleReprot")
        err.Clear: On Error GoTo 0
        If Not mobjPublicLis Is Nothing Then
            Call mobjPublicLis.InitSampleReprot(gcnOracle, glngSys, p住院医生站, "")
        End If
    End If
    If mintType = 2 Then
        picNo.Visible = True
        mblnEditable = False
    ElseIf mintType = 1 Then
        '修改时不允许调整开始执行时间，除非是补录医嘱
        SetControlEnabled txtInfo(txt申请日期), False
        SetControlEnabled cmdDate(cmd申请日期), False
        If mintType = 1 Then
            If InitObjBlood(True) = True Then
                If gobjPublicBlood.GetPrepareBloodRs(mlngUpdateAdvice, rsBlood) = True Then
                    '修改用血医嘱，如果输血科已经发血，则不允许:输血成分、执行科室、预定输血量
                    If Val(rsBlood!记录性质 & "") = 2 And Val(rsBlood!记录状态 & "") = 1 Then
                        mblnUseBloodSend = True
                        SetControlEnabled txtGet(txt预定输血成分), False
                        SetControlEnabled cboInfo(cbo执行科室), False
                        SetControlEnabled txtInfo(txt预定输血量), False
                        vsfList.Editable = flexEDNone
                        vsfBlood.Editable = flexEDNone
                        SetControlEnabled txtInfo(txt申请日期), True
                        SetControlEnabled cmdDate(cmd申请日期), True
                    End If
                End If
            End If
        End If
    ElseIf mintType = 3 Then
        '只能调整除输血成分，总量，申请时间，输血时间，执行科室，输血途径，输血执行科室，用血安排以外的内容
        SetControlEnabled txtInfo(txt申请日期), False
        SetControlEnabled cmdDate(cmd申请日期), False
        SetControlEnabled txtInfo(txt预定输血时间), False
        SetControlEnabled cmdDate(cmd预定输血时间), False
        SetControlEnabled txtGet(txt预定输血成分), False
        SetControlEnabled txtGet(txt输血途径), False
        SetControlEnabled cmdGet(txt输血途径), False
        SetControlEnabled txtInfo(txt预定输血量), False
        SetControlEnabled cboInfo(cbo执行科室), False
        SetControlEnabled cboInfo(cbo输血执行), False
        SetControlEnabled cboInfo(cbo用血安排), False
        SetControlEnabled cboInfo(cbo输血目的), False
        SetControlEnabled cboInfo(cbo输血类型), False
    End If
    mblnChange = mintType = 4
    If Me.Visible = False Then Call InitCommandBar
    If InitInfo = False Then Exit Sub
    Call LoadData
    Call SetFaceEnabledFalse
    Call SetFormNature
    If mbln补录 Then SetControlEnabled cboInfo(cbo用血安排), False
    '病人基本信息不可以编辑
    SetControlEnabled txtInfo(txt性别), False
    SetControlEnabled txtInfo(txt姓名), False
    SetControlEnabled txtInfo(txt年龄), False
    '初始化opt控件
    For i = 2 To 5
        optHistory(i).Enabled = IIF(optHistory(1).value = True, True, False) And optHistory(0).Enabled
    Next
    If optHistory(0).value = True Then
        optHistory(2).value = True
        optHistory(4).value = True
    End If
End Sub

Private Sub SetFaceEnabledFalse()
'功能：已审核不允许修改,已签名的不允许修改
    Dim objControl As Object
    If mblnEditable = False Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
    End If
End Sub

Private Sub SetControlEnabled(objControl As Object, ByVal blnEnabled As Boolean)
'功能：设置控件的可用性
    Select Case TypeName(objControl)
        Case "TextBox", "ComboBox"
            objControl.Locked = Not blnEnabled
            objControl.TabStop = blnEnabled
            objControl.BackColor = IIF(blnEnabled, vbWindowBackground, vbButtonFace)
        Case "CheckBox", "CommandButton", "OptionButton"
            objControl.Enabled = blnEnabled
            
    End Select
End Sub

Private Sub SetControlContent(Optional ByVal blnClearPatiInfo As Boolean = True)
'清空界面控件信息
    Dim objControl As Object
    For Each objControl In Me.Controls
        Select Case TypeName(objControl)
            Case "TextBox"
                If objControl.Name = "txtInfo" Then
                    If blnClearPatiInfo = True Or (blnClearPatiInfo = False And InStr(1, "," & txt诊断信息 & "," & txt孕 & "," & txt产 & "," & txt预定输血时间 & "," & txt备注 & ",", "," & objControl.Index & ",") <> 0) Then
                        objControl.Text = ""
                        objControl.Tag = ""
                    End If
                Else
                    objControl.Text = ""
                    objControl.Tag = ""
                End If
            Case "ComboBox"
                Call zlControl.CboSetIndex(objControl.hwnd, -1)
                objControl.Tag = ""
            Case "Label"
                objControl.Tag = ""
        End Select
    Next
    chkWait.value = 0
    optHistory(0).value = True
    optHistory(2).value = True
    optHistory(4).value = True
    optPossession(0).value = True
    optConsent(0).value = False
    optConsent(1).value = False
    optAppraise(0).value = False
    optAppraise(1).value = False
    
    lblInfo(lbl注意).Caption = ""
    lblInfo(lbl本次历史申请项目).Caption = ""
    lblInfo(lbl24H输血量).Caption = ""
    vsLIS.Rows = 0
    vsLIS.Tag = ""
    '设置两次行数，主要是用于清空实体行相关数据
    vsfBlood.Rows = 1
    vsfBlood.Rows = 2
    vsfList.Rows = 1
    mlng输血项目ID = mlngPre输血项目ID
    mlng输血途径 = 0
    mlng输血执行性质 = 0
    mlng执行科室性质 = 0
    mlng录入限量 = 0
End Sub

Private Function InitInfo(Optional blnFormLoad As Boolean = True) As Boolean
'功能：初始下拉菜单
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim curDate As Date
    Dim lng用法ID As Long
    Dim lng执行科室ID As Long
    Dim strMsg As String, strFilter As String
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    cboInfo(cbo执行科室).Clear
    If blnFormLoad = True Then
        '部分固定内容的下拉框
        Call Cbo.LoadFromList(cboInfo(cbo输血血型), Array(" ", "A", "B", "O", "AB", "不详", "未查"), 0)
        Call Cbo.LoadFromList(cboInfo(cbo用血安排), Array("普通", "急诊"))
        Call Cbo.SetIndex(cboInfo(cbo用血安排).hwnd, 0)
        Call Cbo.LoadFromList(cboInfo(cboRHD), Array(" ", "-", "+"), 0)
        Call Cbo.LoadFromList(cboInfo(cbo滴速), Array("15", "30", "快速", "加压"))
        
        '输血类型
        strSQL = "Select 名称,缺省标志 from 输血类型  order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            With cboInfo(cbo输血类型)
                .Clear
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!名称 & ""
                    If Val(rsTmp!缺省标志 & "") = 1 Then
                        .ListIndex = .ListCount - 1
                    End If
                    rsTmp.MoveNext
                Next
            End With
        End If
        Set rsTmp = Nothing
        
        strSQL = "select 名称,缺省标志 from 输血性质  order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            With cboInfo(cbo输血性质)
                .Clear
                .AddItem " "
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!名称 & ""
                    
                    If Val(rsTmp!缺省标志 & "") = 1 Then
                        .ListIndex = .ListCount - 1
                    End If
                    
                    rsTmp.MoveNext
                Next
                If .ListIndex = -1 Then .ListIndex = 1
            End With
        End If
        Set rsTmp = Nothing
        
        strSQL = "select 名称 from 输血目的 order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            With cboInfo(cbo输血目的)
                .Clear
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!名称 & ""
                    rsTmp.MoveNext
                Next
            End With
        End If
        Set rsTmp = Nothing
        
        '日期
        curDate = zlDatabase.Currentdate
        If mint场合 = 0 Then '只有住院场合才有补录
            If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
                If mdatTurn <> CDate(0) Then curDate = mdatTurn - 1 / 24 / 60
                mbln补录 = True
            End If
        Else
            mbln补录 = False
        End If
        
        'txtInfo(txt预定输血时间).Text = Format(curDate, "YYYY-MM-DD HH:mm")
        txtInfo(txt预定输血时间).Text = ""
        txtInfo(txt预定输血时间).Tag = txtInfo(txt预定输血时间).Text
        txtInfo(txt申请日期).Text = Format(curDate, "YYYY-MM-DD HH:mm")
        txtInfo(txt申请日期).Tag = txtInfo(txt申请日期).Text
    End If
    '缺省用法(新增时检查)
    If mintType = 0 Then
        strFilter = ""
        If mblnSpareBloood = True Then
            lng用法ID = Get缺省用法ID(9, IIF(mint场合 = 0, 2, 1))
            strMsg = "没有可用的输血采集方法,请先到诊疗项目管理中设置！"
        Else
            strFilter = " And nvl(执行分类,0)=" & IIF(mblnSpareBloood = False, 1, 0) '输血途径
            lng用法ID = Get缺省用法ID(8, IIF(mint场合 = 0, 2, 1), strFilter)
            strMsg = "没有可用的输血途径,请先到诊疗项目管理中设置！"
        End If
    
        If lng用法ID = 0 Then
            MsgBox strMsg, vbInformation, gstrSysName
            Screen.MousePointer = 0
            If blnFormLoad = True Then Unload Me
            Exit Function
        Else
            Set rsTmp = Get诊疗项目记录(lng用法ID)
            txtGet(txt输血途径).Text = rsTmp!名称 & ""
            mlng输血执行性质 = NVL(rsTmp!执行科室, 0)
            txtGet(txt输血途径).Tag = txtGet(txt输血途径).Text
            mlng输血途径 = lng用法ID
            cboInfo(cbo输血执行).Enabled = True
            Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(cbo输血执行), "E", mlng输血途径, 0, _
                Val(rsTmp!执行科室 & ""), mlng病人科室id, mlng开单科室ID, 0, 1, IIF(mlng病人性质 = 1, 1, 2))
            If cboInfo(cbo输血执行).ListIndex = -1 And cboInfo(cbo输血执行).ListCount > 1 Then
                Call Cbo.SetIndex(cboInfo(cbo输血执行).hwnd, 0)
            Else
                '如果有多项，则取默认的执行科室
                lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", mlng输血途径, 0, _
                        NVL(rsTmp!执行科室, 0), mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
                If lng执行科室ID <> 0 Then
                    Call Cbo.Locate(cboInfo(cbo输血执行), lng执行科室ID, True)
                End If
            End If
            If cboInfo(cbo输血执行).ListCount = 2 Then cboInfo(cbo输血执行).Enabled = False
            cboInfo(cbo输血执行).Tag = lng用法ID
        End If
    End If
    InitInfo = True
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadPatiInfo()
'功能：提取病人基本信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    '读取病人相关信息
    txtInfo(txt就诊类型).Text = IIF(mlng病人性质 = 1, "门诊", "住院")
    
    If mint场合 = 0 Then
        If mbytBaby = 0 Then
            strSQL = "Select A.住院号, Nvl(C.姓名, A.姓名) 姓名, Nvl(C.性别, A.性别) 性别, Nvl(C.年龄, A.年龄) 年龄, B.名称 As 科室, C.出院病床 As 当前床号, C.入院日期, C.险类,c.费别" & vbNewLine & _
                    "From 病人信息 A, 部门表 B, 病案主页 C" & vbNewLine & _
                    "Where C.出院科室id = B.Id And A.病人id = C.病人id And A.主页id = C.主页id And C.病人id = [1] And C.主页id = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        Else
            strSQL = "Select a.住院号, Nvl(q.婴儿姓名, a.姓名 || '之婴' || q.序号) 姓名, q.婴儿性别 性别, Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间) || '天' As 年龄, b.名称 As 科室," & vbNewLine & _
                    " c.出院病床 As 当前床号, c.入院日期, c.险类,c.费别" & vbNewLine & _
                    "From 病人信息 A, 部门表 B, 病案主页 C, 病人新生儿记录 Q" & vbNewLine & _
                    "Where a.病人id = c.病人id And a.主页id = c.主页id And A.病人id = q.病人id And A.主页id = q.主页id And c.出院科室id = b.Id And c.病人id = [1] " & vbNewLine & _
                    " And c.主页id = [2] And q.序号 = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mbytBaby)
        End If
        If rsTmp.RecordCount > 0 Then
            txtInfo(txt住院号).Text = rsTmp!住院号 & ""
            txtInfo(txt姓名).Text = rsTmp!姓名 & ""
            txtInfo(txt性别).Text = rsTmp!性别 & ""
            If txtInfo(txt性别).Text = "男" Or mbytBaby <> 0 Then
                SetControlEnabled txtInfo(txt孕), False
                SetControlEnabled txtInfo(txt产), False
            End If
            txtInfo(txt科室).Text = rsTmp!科室 & ""
            txtInfo(txt床号).Text = rsTmp!当前床号 & ""
            txtInfo(txt年龄).Text = rsTmp!年龄 & ""
            mstr入院时间 = Format(rsTmp!入院日期 & "", "YYYY-MM-DD HH:mm")
            mint险类 = Val(rsTmp!险类 & "")
            mstr费别 = rsTmp!费别 & ""
        End If
    Else
        strSQL = "Select a.ID, A.姓名,A.性别,A.年龄,a.no,a.门诊号,a.险类,b.名称 as 科室,a.执行时间,c.费别" & _
            " From 病人挂号记录 A,部门表 b,病人信息 c " & _
            " Where a.病人ID=c.病人ID and A.NO=[1] And a.记录性质=1 And a.记录状态=1 And A.病人ID+0=[2] and a.执行部门id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单, mlng病人ID)
        If rsTmp.RecordCount > 0 Then
            lblInfo(lbl挂号单).Caption = "挂 号 单"
            txtInfo(txt挂号单).Text = rsTmp!NO & ""
            txtInfo(txt姓名).Text = rsTmp!姓名 & ""
            txtInfo(txt性别).Text = rsTmp!性别 & ""
            If txtInfo(txt性别).Text = "男" Or mbytBaby <> 0 Then
                SetControlEnabled txtInfo(txt孕), False
                SetControlEnabled txtInfo(txt产), False
            End If
            txtInfo(txt科室).Text = rsTmp!科室 & ""
            lblInfo(lbl门诊号).Caption = "门 诊 号"
            txtInfo(txt门诊号).Text = rsTmp!门诊号 & ""
            txtInfo(txt年龄).Text = rsTmp!年龄 & ""
            mint险类 = Val(rsTmp!险类 & "")
            mstr费别 = rsTmp!费别 & ""
            mstr入院时间 = Format(rsTmp!执行时间 & "", "YYYY-MM-DD HH:mm")
            mlng挂号ID = Val(rsTmp!ID & "")
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadData() As Boolean
'功能：读取申请单信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim str诊断 As String
    Dim rsTmpOther As ADODB.Recordset
    Dim strTmp As String
    Dim strItemName As String, strIDs As String
    Dim arrItem
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    '读取病人相关信息
    Call LoadPatiInfo

    If mintType = 0 Then
        '如果是新增切换模式，则直接读取对应单据
        If Me.Tag = "GOTO" Then GoTo GoLoadData
        If mint场合 = 0 Then
            '读取上次转科时间
            strSQL = "Select 开始时间 From 病人变动记录" & _
                " Where 开始时间 is Not NULL And 开始原因=3" & _
                " And 病人ID=[1] And 主页ID=[2] Order by 开始时间 desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", mlng病人ID, mlng主页ID)
            If rsTmp.RecordCount > 0 Then
                mstr上次转科时间 = Format(rsTmp!开始时间 & "", "YYYY-MM-DD HH:mm")
            End If
        End If
        '下达用血申请时：诊断、输血目的、血型等默认去最后一次备血申请的信息
        Call LoadLastPrepareBlood
    ElseIf mintType = 1 Or mintType = 3 Or mintType = 2 Or mintType = 4 Then
GoLoadData:
        If mintType = 4 Then '医嘱审核状态还未有输血申请记录
            '直接发血获取对应备血医嘱的相关信息
            strSQL = "Select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='备血申请ID'"
            Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            If Not rsTmpOther.EOF Then
                Call LoadLastPrepareBlood(Val(rsTmpOther!内容 & ""))
            Else
                mstr诊断IDs = GetAdviceDiag(mlngUpdateAdvice, str诊断)
                txtInfo(txt诊断信息).Text = str诊断
                strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
                Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
                If Not rsTmpOther.EOF Then
                    txtInfo(txt诊断信息).Text = rsTmpOther!内容 & ""
                End If
                txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
                chkWait.value = IIF(txtInfo(txt诊断信息).Text = "待诊", 1, 0)
                
                '血型备血申请单可能没有，从病人信息从表中获取
                If cboInfo(cbo输血血型).ListIndex <= 0 Then
                    strSQL = "Select 信息值 from 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, "ABO")
                    If Not rsTmpOther.EOF Then
                        Select Case "" & rsTmpOther!信息值
                            Case "A", "A型"
                                cboInfo(cbo输血血型).ListIndex = 1
                            Case "B", "B型"
                                cboInfo(cbo输血血型).ListIndex = 2
                            Case "O", "O型"
                                cboInfo(cbo输血血型).ListIndex = 3
                            Case "AB", "AB型"
                                cboInfo(cbo输血血型).ListIndex = 4
                        End Select
                    End If
                End If
                If cboInfo(cboRHD).ListIndex <= 0 Then
                    strSQL = "Select 信息值 from 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, "RH")
                    If Not rsTmpOther.EOF Then
                        Select Case "" & rsTmpOther!信息值
                            Case "-", "阴"
                                cboInfo(cboRHD).ListIndex = 1
                            Case "+", "阳"
                                cboInfo(cboRHD).ListIndex = 2
                        End Select
                    End If
                End If
            End If
        Else
            '修改
            '读取输血相关信息
            strSQL = _
                " Select 是否待诊, 输血类型, 输血目的, 输血性质, 即往输血史, 既往输血反应史, 输血禁忌及过敏史, 孕产情况, 受血者属地, 是否签订同意书, 是否已评估, 输血血型, Rhd, 受血者血型, Hct, Alt, Hbsag," & vbNewLine & _
                "       梅毒, 血红蛋白, 血小板, Antihcv, Antihiv12" & vbNewLine & _
                " From 输血申请记录" & vbNewLine & _
                " Where 医嘱id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!是否待诊 & "") = 1 Then
                    txtInfo(txt诊断信息).Text = "待诊"
                    chkWait.value = 1
                Else
                    '读取诊断
                    mstr诊断IDs = GetAdviceDiag(mlngUpdateAdvice, str诊断)
                    txtInfo(txt诊断信息).Text = str诊断
                    '从附项中获取诊断如果附项中有以附项为准
                     strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
                     Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
                     If Not rsTmpOther.EOF Then
                         txtInfo(txt诊断信息).Text = rsTmpOther!内容 & ""
                     End If
                End If
                txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
                chkWait.value = Val(rsTmp!是否待诊 & "")
                Call zlControl.CboSetText(cboInfo(cbo输血类型), rsTmp!输血类型 & "", True, "'")
                Call zlControl.CboSetText(cboInfo(cbo输血目的), rsTmp!输血目的 & "", True, "'")
                cboInfo(cbo输血性质).ListIndex = Val(rsTmp!输血性质 & "")
                optHistory(Val(rsTmp!即往输血史 & "")).value = True
                optHistory(IIF(Val(rsTmp!既往输血反应史 & "") = 1, 3, 2)).value = True
                optHistory(IIF(Val(rsTmp!输血禁忌及过敏史 & "") = 1, 5, 4)).value = True
                If InStr(1, "" & rsTmp!孕产情况, "/") <= 0 Then
                    txtInfo(txt孕).Text = ""
                    txtInfo(txt产).Text = ""
                Else
                    txtInfo(txt孕).Text = Mid(rsTmp!孕产情况, 1, InStr(1, "" & rsTmp!孕产情况, "/") - 1)
                    If Not (txtInfo(txt孕).Text = "" Or IsNumeric(txtInfo(txt孕).Text)) Then
                        txtInfo(txt孕).Text = ""
                    End If
                    txtInfo(txt产).Text = Mid(rsTmp!孕产情况, InStr(1, "" & rsTmp!孕产情况, "/") + 1)
                    If Not (txtInfo(txt产).Text = "" Or IsNumeric(txtInfo(txt产).Text)) Then
                        txtInfo(txt产).Text = ""
                    End If
                End If
                optPossession(Val(rsTmp!受血者属地 & "")).value = True
                If InStr(1, ",0,1,", "," & rsTmp!是否签订同意书 & ",") <> 0 Then
                    optConsent(Val(rsTmp!是否签订同意书 & "")).value = True
                End If
                If InStr(1, ",0,1,", "," & rsTmp!是否已评估 & ",") <> 0 Then
                    optAppraise(Val(rsTmp!是否已评估 & "")).value = True
                End If
                
                cboInfo(cbo输血血型).ListIndex = Val(rsTmp!输血血型 & "")
                cboInfo(cboRHD).ListIndex = Val(rsTmp!RHD & "")
            End If
        End If
        '读取血液申请项目(如果申请项目记录数据<=1，则不用做任何处理)
        mstr输血项目 = ""
        strItemName = ""
        strIDs = ""
        strSQL = "Select A.名称,B.诊疗项目ID,B.申请量,B.申请血型,B.申请RH,b.血液信息 From 诊疗项目目录 A,输血申请项目 B where A.ID=B.诊疗项目ID And B.医嘱ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
        Do While Not rsTmp.EOF
            strIDs = IIF(strIDs = "", "", strIDs & ",") & rsTmp!诊疗项目ID
            mstr输血项目 = IIF(mstr输血项目 = "", "", mstr输血项目 & ";") & rsTmp!诊疗项目ID & "," & rsTmp!申请量 & "," & rsTmp!申请血型 & "," & rsTmp!申请rh & IIF(rsTmp!血液信息 & "" <> "", "," & rsTmp!血液信息, "")
            strItemName = IIF(strItemName = "", "", strItemName & "'") & rsTmp!名称
        rsTmp.MoveNext
        Loop
        '读取医嘱相关信息（主血液医嘱）
        strSQL = "Select A.ID,A.相关ID,a.紧急标志,a.用药理由,NVL(to_char(a.手术时间,'yyyy-MM-dd hh24:mi'),a.标本部位) as 预定输血时间,a.开始执行时间,a.诊疗项目ID," & _
                " a.执行科室ID,a.执行性质,a.总给予量,B.类别,B.操作类型,B.计算单位,B.名称 as 项目名称,b.执行分类,A.申请序号,A.审核状态,a.医生嘱托" & vbNewLine & _
                " From 病人医嘱记录 A,诊疗项目目录 B" & vbNewLine & _
                " Where a.诊疗项目ID=B.ID And (A.id = [1] or A.相关ID=[1])"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!紧急标志 & "") = 2 Then
                    mbln补录 = True
                    SetControlEnabled txtInfo(txt申请日期), True
                    SetControlEnabled cmdDate(cmd申请日期), True
                ElseIf Val(rsTmp!紧急标志 & "") = 1 Then
                    cboInfo(cbo用血安排).ListIndex = 1
                End If
                If cboInfo(cbo输血目的).Text = "" Then Call zlControl.CboSetText(cboInfo(cbo输血目的), rsTmp!用药理由 & "", True, "'")   '老的的输血目的存储在医嘱的用药理由里面
                txtInfo(txt预定输血时间).Text = Format(rsTmp!预定输血时间 & "", "YYYY-MM-DD HH:mm")
                txtInfo(txt预定输血时间).Tag = txtInfo(txt预定输血时间).Text
                txtInfo(txt申请日期).Text = Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD HH:mm")
                txtGet(txt预定输血成分).Text = IIF(strItemName = "", rsTmp!项目名称 & "", strItemName)
                txtInfo(txt单位).Text = rsTmp!计算单位 & ""
                txtGet(txt预定输血成分).Tag = txtGet(txt预定输血成分).Text
                mlng输血项目ID = Val(rsTmp!诊疗项目ID)
                
                Call Set执行科室(Val(rsTmp!执行性质 & ""), Val(rsTmp!执行科室ID & ""))
                Call LoadLisResult(mlngUpdateAdvice)
                
                txtInfo(txt预定输血量).Text = zl9ComLib.FormatEx(rsTmp!总给予量 & "", 5)
                txtInfo(txtNO).Text = rsTmp!申请序号 & ""
                txtInfo(txt备注).Text = rsTmp!医生嘱托 & ""
                '已经审核通过的不允许修改（备血完成或发血）
                If Val(rsTmp!审核状态 & "") = 2 And mblnUseBloodSend = False Then mblnEditable = False
                If InStr(1, "," & strIDs & ",", "," & mlng输血项目ID & ",") = 0 Then
                    mstr输血项目 = mlng输血项目ID & "," & txtInfo(txt预定输血量).Text & "," & cboInfo(cbo输血血型).Text & "," & cboInfo(cboRHD).Text & IIF(mstr输血项目 = "", "", ";" & mstr输血项目)
                End If
            End If
            rsTmp.Filter = "相关ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                txtGet(txt输血途径).Text = rsTmp!项目名称 & ""
                txtGet(txt输血途径).Tag = txtGet(txt输血途径).Text
                mlng输血途径 = Val(rsTmp!诊疗项目ID)
                If Not (rsTmp!类别 = "E" And rsTmp!操作类型 = "9") Then
                    mblnNewSpareBloood = False
                    If rsTmp!类别 = "E" And rsTmp!操作类型 = "8" Then
                        mblnSpareBloood = (Val(rsTmp!执行分类 & "") = 0) '不管新老流程 用血医嘱的执行分类=1
                    End If
                Else
                    mblnNewSpareBloood = True
                    mblnSpareBloood = True '诊疗类别为E,操作类型=9的就是备血医嘱
                End If
                Call Set输血执行(Val(rsTmp!执行性质 & ""), Val(rsTmp!执行科室ID & ""))
                '用血医嘱滴速
                strTmp = rsTmp!医生嘱托 & ""
                cboInfo(cbo滴速).Text = ""
                lblInfo(31).Visible = True
                If strTmp Like "*滴/分钟" Then
                    If IsNumeric(Split(strTmp, "滴/分钟")(0)) = True Then
                        cboInfo(cbo滴速).Text = Split(strTmp, "滴/分钟")(0)
                    End If
                ElseIf strTmp = "加压" Or strTmp = "快速" Then
                    cboInfo(cbo滴速).Text = strTmp
                    lblInfo(31).Visible = False
                End If
            End If
        End If
        '读取签名记录
        If gintCA <> 0 And Mid(gstrESign, 2, 1) = "1" Then
            strSQL = "Select b.签名人,A.操作类型 From 病人医嘱状态 A, 医嘱签名记录 B Where a.签名id = b.Id And a.医嘱id = [1] And 操作类型=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            If rsTmp.RecordCount > 0 Then
                mblnEditable = False
                '签名人
                rsTmp.Filter = "操作类型=1"
                If rsTmp.RecordCount > 0 Then
                    txtInfo(txt申请医师签名).Text = rsTmp!签名人 & ""
                End If
                '审核人（审核暂未启用签名功能，主要考虑到审核的人不一定有签名的U盾)
'                rsTmp.Filter = "操作类型=11"
'                If rsTmp.RecordCount > 0 Then
'                    txtInfo(txt主治医师签名).Text = rsTmp!签名人 & ""
'                End If
            End If
        End If
    End If
    
    Call LoadDataFromCache
    strIDs = ""
    arrItem = Split(mstr输血项目, ";")
    For i = 0 To UBound(arrItem)
        strIDs = strIDs & "," & Split(CStr(arrItem(i)), ",")(0)
    Next
    strIDs = Mid(strIDs, 2)

    If mlng输血项目ID <> 0 Then
        If InStr(1, "," & strIDs & ",", "," & mlng输血项目ID & ",") = 0 Then
            strIDs = mlng输血项目ID & "," & strIDs
            mstr输血项目 = mlng输血项目ID & "," & txtInfo(txt预定输血量).Text & "," & cboInfo(cbo输血血型).Text & "," & cboInfo(cboRHD).Text & IIF(mstr输血项目 = "", "", ";" & mstr输血项目)
        End If
        Set rsTmp = Get诊疗项目记录(mlng输血项目ID, strIDs)
        strTmp = ""
        Do While Not rsTmp.EOF
            strTmp = strTmp & IIF(strTmp = "", "", "'") & rsTmp!名称
            rsTmp.MoveNext
        Loop
        txtGet(txt预定输血成分).Text = strTmp
        rsTmp.Filter = "ID=" & mlng输血项目ID
        Call Set执行科室(Val(rsTmp!执行科室 & ""))
        txtInfo(txt单位).Text = rsTmp!计算单位 & ""
        txtGet(txt预定输血成分).Tag = txtGet(txt预定输血成分).Text
        mlng录入限量 = Val(rsTmp!录入限量 & "")
        If mrsCard Is Nothing And mint调用场合 = 1 Then Call SetLisResult(strIDs)
    End If
    
    Screen.MousePointer = 0
    LoadData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 保存(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " 保存退出(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagStretched
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    With objBar.Controls
        Set objMenuBar = .Add(xtpControlButtonPopup, conMenu_Tool_Archive, "申请单▼", "选择输血申请单的类型"): 'objMenuBar.IconId = 807
        objMenuBar.Style = xtpButtonCaption
        objMenuBar.Flags = xtpFlagRightAlign
        Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_Archive * 10# + 1, "输血申请单(&1)")
        Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_Archive * 10# + 2, "取血通知单(&2)")
    End With
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub Set执行科室(ByVal lng执行科室 As Long, Optional ByVal lng执行科室ID As Long)
'功能：设置执行科室
'参数：lng执行科室-执行性质，lng执行科室ID=如果传入，则表示设置此执行科室为当前执行科室
    Dim lngTmp As Long
 
    cboInfo(cbo执行科室).Enabled = True
    If lng执行科室 = 5 Then
        cboInfo(cbo执行科室).Clear: cboInfo(cbo执行科室).AddItem "-"
        cboInfo(cbo执行科室).ListIndex = 0
    Else
        If cboInfo(cbo执行科室).ListIndex >= 0 And lng执行科室ID = 0 Then
            lngTmp = cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex)
        ElseIf lng执行科室ID <> 0 Then
            lngTmp = lng执行科室ID
        End If
        
        Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(cbo执行科室), "K", mlng输血项目ID, 0, lng执行科室, mlng病人科室id, mlng开单科室ID, lngTmp, 1, IIF(mlng病人性质 = 1, 1, 2))
        If lng执行科室ID = 0 Then
            If cboInfo(cbo执行科室).ListIndex = -1 And cboInfo(cbo执行科室).ListCount = 1 Then
                cboInfo(cbo执行科室).ListIndex = 0
            Else
                 '如果有多项，则取默认的执行科室
                lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "K", mlng输血项目ID, 0, _
                        lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
            End If
        End If
        If lng执行科室ID <> 0 Then
            Call zlControl.CboLocate(cboInfo(cbo执行科室), lng执行科室ID, True)
        End If
    End If
    mlng执行科室性质 = lng执行科室
    If cboInfo(cbo执行科室).ListCount = 1 Then cboInfo(cbo执行科室).Enabled = False
    If cboInfo(cbo执行科室).ListIndex >= 0 Then
        cboInfo(cbo执行科室).Tag = cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex)
    End If
End Sub

Private Sub Set输血执行(ByVal lng执行科室 As Long, Optional ByVal lng执行科室ID As Long)
'功能：设置输血执行科室
'参数：lng执行科室-执行性质，lng执行科室ID=如果传入，则表示设置此执行科室为当前执行科室
    Dim lngTmp As Long
    
    cboInfo(cbo输血执行).Enabled = True
    If lng执行科室 = 5 Then
        cboInfo(cbo输血执行).Clear: cboInfo(cbo输血执行).AddItem "-"
        cboInfo(cbo输血执行).ListIndex = 0
    Else
        If cboInfo(cbo输血执行).ListIndex >= 0 And lng执行科室ID = 0 Then
            lngTmp = cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)
        ElseIf lng执行科室ID <> 0 Then
            lngTmp = lng执行科室ID
        End If
        
        Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(cbo输血执行), "E", mlng输血途径, 0, _
            lng执行科室, mlng病人科室id, mlng开单科室ID, lngTmp, 1, IIF(mlng病人性质 = 1, 1, 2), , , , , , , , mlng病人性质)
        If lng执行科室ID = 0 Then
            If cboInfo(cbo输血执行).ListIndex = -1 And cboInfo(cbo输血执行).ListCount = 1 Then
                cboInfo(cbo输血执行).ListIndex = 0
            Else
                 '如果有多项，则取默认的执行科室
                lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", mlng输血途径, 0, _
                        lng执行科室, mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
            End If
        End If
        If lng执行科室ID <> 0 Then
            Call zlControl.CboLocate(cboInfo(cbo输血执行), lng执行科室ID, True)
        End If
    End If
    mlng输血执行性质 = lng执行科室
    If cboInfo(cbo输血执行).ListCount = 1 Then cboInfo(cbo输血执行).Enabled = False
    If cboInfo(cbo输血执行).ListIndex >= 0 Then
    cboInfo(cbo输血执行).Tag = cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)
    End If
End Sub

Private Sub TxtGetInfo(Index As Integer, Optional ByVal intType As Integer)
'功能：设置文本框内容
'参数：intType =0 KeyPress调用，=1 下拉按钮调用
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    
    '用血申请不能敲回车
    If Index <> txt输血途径 Then Exit Sub
    
    If mblnNewSpareBloood = True And mblnSpareBloood = True Then
        strSQL = " And A.类别='E' And A.操作类型='9' "  '采集方式
    Else
        strSQL = " And A.类别='E' And A.操作类型='8' And nvl(A.执行分类,0)=" & IIF(mblnSpareBloood = False, 1, 0) '输血途径
    End If
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.执行分类 as 执行分类ID,A.计算单位,A.执行科室 as 执行科室ID,A.录入限量 as 录入限量ID" & _
    " From 诊疗项目目录 A,诊疗项目别名 B" & _
    " Where A.ID=B.诊疗项目ID" & _
    strSQL & "  And A.服务对象 IN(" & IIF(mlng病人性质 = 1, "1,2", 2) & ",3)" & _
    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
    IIF(intType = 0, " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])", "") & _
    IIF(mlng病人性质 = 1, "", " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4]) Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))") & _
    Decode(gbytCode, 0, " And B.码类 IN([3],3)", 1, " And B.码类 IN([3],3)", "") & _
    " Order by A.编码"
            
    vRect = zlControl.GetControlRect(txtGet(Index).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtGet(Index).Height, blnCancel, False, True, UCase(txtGet(Index).Text) & "%", _
        gstrLike & UCase(txtGet(Index).Text) & "%", gbytCode + 1, mlng病人科室id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到匹配的项目。", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtGet(Index))
        txtGet(Index).SetFocus: Exit Sub
    Else
        Call SetTxtBloodInfo(rsTmp, Index, True)
    End If
End Sub

Private Function SetTxtBloodInfo(ByVal rsTmp As ADODB.Recordset, Optional ByVal Index As Integer, Optional ByVal blnNextControl As Boolean = True) As Boolean
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    Dim strName As String, strID As String
    
    On Error GoTo ErrHand
    
    If Index = txt预定输血成分 Then
        If rsTmp.RecordCount > 0 Then
            mlng输血项目ID = Val(rsTmp!ID)
            mlng录入限量 = Val(rsTmp!录入限量ID & "")
            txtInfo(txt单位).Text = rsTmp!计算单位 & ""
            Call Set执行科室(Val(rsTmp!执行科室ID & "")) '备血选择多个品种以第一个为准确定执行科室
        Else
            mlng输血项目ID = 0
            mlng录入限量 = 0
            txtInfo(txt单位).Text = ""
        End If
        Do While Not rsTmp.EOF
            strName = IIF(strName = "", "", strName & "'") & rsTmp!名称
            strID = IIF(strID = "", "", strID & ",") & rsTmp!ID
            strIDs = IIF(strIDs = "", "", strIDs & ",") & rsTmp!ID & ":" & IIF(Val(cboInfo(cbo执行科室).Tag & "") <> 0, Val(cboInfo(cbo执行科室).Tag & ""), "")
            rsTmp.MoveNext
        Loop
        txtGet(Index).Text = strName
        txtGet(Index).Tag = txtGet(Index).Text
        Call SetLisResult(strID)
        '对码检查
        If strIDs <> "" Then
            str医嘱内容 = FormatAdviceContext(Replace(txtGet(txt预定输血成分).Text, "'", ","), txtGet(txt输血途径).Text)
        End If
    ElseIf Index = txt输血途径 Then
        txtGet(Index).Text = rsTmp!名称 & ""
        txtGet(Index).Tag = txtGet(Index).Text
        mlng输血途径 = Val(rsTmp!ID)
        Call Set输血执行(Val(rsTmp!执行科室ID & ""))
        '对码检查
        If mlng输血途径 <> 0 Then
            strIDs = strIDs & "," & mlng输血途径 & ":"
            If Val(cboInfo(cbo输血执行).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo输血执行).Tag & "")
            End If
        End If
    End If
    
    strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, IIF(mlng病人性质 = 0, 2, 1), "", strIDs, str医嘱内容)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    
    If blnNextControl = True Then
        If txtGet(Index).Enabled And txtGet(Index).Visible Then txtGet(Index).SetFocus
        Call SeekNextControl
    End If
    If Visible Then mblnChange = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetLisResult(ByVal str输血项目ID As String)
'功能：初始化输血项目对应的检验项目指标表格（备血申请可能有所个输血项目）
    Dim rsLIS As ADODB.Recordset '当前输血的检验项目
    Dim rs结果 As ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim strResult As String, arr项目名称() As String
    Dim str指标名 As String, str时间 As String
    Dim str检验编码 As String
    Dim strTmp As String, strTmp1 As String
    Dim arrTmp1 As Variant
    Dim arrTmp2 As Variant
    Dim arrBloodID As Variant, strBloodapplyrate As String, arrTmp4 As Variant, blnAdd As Boolean
    Dim i As Long, j As Long, k As Long
    Dim lngCol As Long
    Dim arrTmp3 As Variant
    Dim bln指定显示项 As Boolean, blnGet As Boolean
    Dim int历史结果天数 As Long, str历史检验编码 As String, str历史检验名称
    Dim arrDay, arrItem
    Dim strHisResult As String, strLisInfo As String
    '启用了血库时，用血申请不需要检查结果
    If mblnSpareBloood = False Then Exit Sub
    
    On Error GoTo errH
    arrBloodID = Split(str输血项目ID, ",")
    '130538:检验项目提取历次就诊结果支持指定天数
    If UBound(arrBloodID) > 0 Then
        strSQL = "Select /*+ CARDINALITY(C 10) */ A.检验项目ID,B.编码,B.名称,A.历史结果天数 from 输血检验对照 A,诊疗项目目录 B,Table(f_Num2list([1])) C " & _
            " Where A.检验项目ID=B.ID And A.项目ID=C.Column_Value Order by B.编码"
        Set rsLIS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str输血项目ID)
    Else
        strSQL = "Select A.检验项目ID,B.编码,B.名称,A.历史结果天数 from 输血检验对照 A,诊疗项目目录 B Where A.检验项目ID=B.ID And A.项目ID=[1] Order by B.编码"
        Set rsLIS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str输血项目ID))
    End If
    strTmp = ""
    Do While Not rsLIS.EOF
        If InStr(1, "," & str检验编码 & ",", "," & rsLIS!编码 & ",") = 0 Then
            str检验编码 = str检验编码 & "," & rsLIS!编码
        End If
        int历史结果天数 = IIF(Val("" & rsLIS!历史结果天数) <= 0, 7, Val("" & rsLIS!历史结果天数))
        If InStr(1, "," & strTmp & ",", "," & int历史结果天数 & ",") = 0 Then
            strTmp = strTmp & "," & int历史结果天数
        End If
        rsLIS.MoveNext
    Loop
    str检验编码 = Mid(str检验编码, 2)
    '历史检验指标按天数排列显示
    strTmp = Mid(strTmp, 2)
    arrDay = Split(strTmp, ",")
    arrItem = Array()
    strLisInfo = ""
    For i = 0 To UBound(arrDay)
        rsLIS.Filter = ""
        str历史检验编码 = ""
        str历史检验名称 = ""
        Do While Not rsLIS.EOF
            int历史结果天数 = IIF(Val("" & rsLIS!历史结果天数) <= 0, 7, Val("" & rsLIS!历史结果天数))
            If Val(arrDay(i)) = int历史结果天数 Then
                If InStr(1, "," & str历史检验编码 & ",", "," & rsLIS!编码 & ",") = 0 Then
                    str历史检验编码 = str历史检验编码 & "," & rsLIS!编码
                End If
                If InStr(1, "," & str历史检验名称 & ",", "," & rsLIS!名称 & ",") = 0 Then
                    str历史检验名称 = str历史检验名称 & ",[" & rsLIS!名称 & "]"
                End If
            End If
            rsLIS.MoveNext
        Loop
        str历史检验编码 = Mid(str历史检验编码, 2)
        str历史检验名称 = Mid(str历史检验名称, 2)
        strLisInfo = IIF(strLisInfo = "", "", strLisInfo & vbCrLf) & Val(arrDay(i)) & "天内：" & str历史检验名称
        ReDim Preserve arrItem(UBound(arrItem) + 1)
        arrItem(UBound(arrItem)) = str历史检验编码
    Next
    
    With vsLIS
        .Clear
        .Rows = 0
        If str检验编码 = "" Then Exit Sub

        strResult = mobjPublicLis.GetTransfusionApplyFor(str检验编码, mlng病人ID, IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mstr挂号单, CInt(mbytBaby), 0)
        strTmp = strResult
        strTmp = Replace(strTmp, "<split1>", "")
        strTmp = Replace(strTmp, "<split2>", "")
        strTmp = Replace(strTmp, "<split3>", "")
        strTmp = Trim(strTmp)
        
        If mint场合 = 0 Then
            strTmp1 = ""
            If strTmp <> "" Then
                arrTmp1 = Split(strResult, "<split3>")
                For i = 0 To UBound(arrTmp1)
                    If Replace(Replace(CStr(arrTmp1(i)), "<split1>", ""), "<split2>", "") <> "" Then 'strResult会存在<split1><split1><split3>这种
                        arrTmp2 = Split(arrTmp1(i), "<split1>")
                        If arrTmp2(8) <> "" Then
                            strTmp1 = "有结果"
                            Exit For
                        End If
                    End If
                Next
            End If
            If strTmp1 = "" Then
                If vsLIS.Tag = "YES" Then
                    blnGet = True
                Else
                    If UBound(arrDay) = 0 Then
                        blnGet = (MsgBox("本次住院未找到有效的检验指标，是否提取历次就诊" & Val(arrDay(0)) & "天内的检验指标？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes)
                    Else
                        blnGet = (MsgBox("本次住院未找到有效的检验指标，是否提取历次就诊指定天数内的检验指标？" & vbCrLf & strLisInfo, vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes)
                    End If
                End If
                vsLIS.Tag = IIF(blnGet = True, "YES", "")
                If blnGet = True Then
                    strResult = ""
                    For i = 0 To UBound(arrItem)
                        str历史检验编码 = CStr(arrItem(i))
                        int历史结果天数 = Val(arrDay(i))
                        strHisResult = ""
                        strHisResult = mobjPublicLis.GetTransfusionApplyFor(str历史检验编码, mlng病人ID, IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mstr挂号单, CInt(mbytBaby), 2, int历史结果天数)
                        strResult = IIF(strResult = "", "", strResult & "<split3>") & strHisResult
                    Next
                    strTmp = strResult
                    strTmp = Replace(strTmp, "<split1>", "")
                    strTmp = Replace(strTmp, "<split2>", "")
                    strTmp = Replace(strTmp, "<split3>", "")
                    strTmp = Trim(strTmp)
                End If
            End If
        End If
        
        If strTmp <> "" Then
'            指标1<split1>诊疗编码1<split1>单位1<split1>隐私项目1<split1>指标代码1<split1>中文名1<split1>英文名1<split1>取值序列1<split1>
                '检验结果1<split2>结果标志1<split2>结果参数1<split2>排列序号1<split2>标本类型1<split2>指标审核时间1<split3>
'            指标2<split1>诊疗编码2<split1>隐私项目2<split1>指标代码2<split1>中文名2<split1>英文名2<split1>取值序列2<split1>
              '  检验结果2<split2>结果标志2<split2>结果参数2<split2>排列序号2<split2>标本类型2<split2>指标审核时间1<split3>
            '重新赋值：strResult会存在<split1><split1><split3>这种
            arrTmp1 = Split(strResult, "<split3>")
            strTmp = "": str指标名 = "": strTmp1 = ""
            For i = 0 To UBound(arrTmp1)
                If Replace(Replace(CStr(arrTmp1(i)), "<split1>", ""), "<split2>", "") <> "" Then
                    str指标名 = Split(arrTmp1(i), "<split1>")(4)
                    If InStr(1, "'" & strTmp1 & "'", "'" & str指标名 & "'") = 0 Then '去掉重复的指标
                        strTmp = strTmp & IIF(strTmp = "", "", "<split3>") & CStr(arrTmp1(i))
                        strTmp1 = strTmp1 & IIF(strTmp1 = "", "", "'") & str指标名
                    End If
                End If
            Next i
            strResult = strTmp
            arrTmp1 = Split(strResult, "<split3>")
            
            strTmp = "": str指标名 = ""
            For i = 0 To UBound(arrTmp1)
                str指标名 = Split(arrTmp1(i), "<split1>")(5) '取诊治所见项目的 中文名
                str时间 = "无"
                arrTmp2 = Split(arrTmp1(i), "<split1>")
                If arrTmp2(8) <> "" Then
                    If UBound(Split(arrTmp2(8), "<split2>")) >= 5 Then
                        str时间 = Split(arrTmp2(8), "<split2>")(5)
                        
                        If IsDate(str时间) Then
                            str时间 = Format(str时间, "YYYY-MM-DD HH:MM:SS")
                        Else
                            str时间 = "无"
                        End If
                    End If
                End If
                strTmp = strTmp & "," & str指标名 & "," & str时间
            Next
            
            strPar = Mid(strTmp, 2)
            arr项目名称 = Split(txtGet(txt预定输血成分).Text, "'")
            strBloodapplyrate = ""
            For i = 0 To UBound(arrBloodID)
                strSQL = "select Zl_Fun_Bloodapplyrate([1],[2]) as 指标 from dual"
                Set rs结果 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(arr项目名称(i)), strPar)
                
                If Not rs结果.EOF Then
                    strTmp = rs结果!指标 & ""
                Else
                    strTmp = ""
                End If
                
                '如果只有一个血液品种则保持原有方式，否则统一按照Zl_Fun_Bloodapplyrate函数转换
                If UBound(arrBloodID) > 0 Then
                    If strTmp <> "" And strBloodapplyrate <> strTmp Then
                        strTmp1 = ""
                        arrTmp4 = Split(strTmp, ",")
                        For j = 0 To UBound(arrTmp4)
                            blnAdd = True
                            If CStr(arrTmp4(j)) <> "" And Not IsDate(CStr(arrTmp4(j))) And CStr(arrTmp4(j)) <> "无" Then
                                If strBloodapplyrate = "" Then
                                    strTmp1 = IIF(strTmp1 = "", "", strTmp1 & ",") & arrTmp4(j)
                                Else
                                    For k = 0 To UBound(Split(strBloodapplyrate, ","))
                                        If Split(arrTmp4(j), "|")(0) = Split(Split(strBloodapplyrate, ",")(k), "|")(0) Then
                                            blnAdd = False
                                            Exit For
                                        End If
                                    Next k
                                    If blnAdd = True Then
                                        strTmp1 = IIF(strTmp1 = "", "", strTmp1 & ",") & arrTmp4(j)
                                    End If
                                End If
                            End If
                        Next
                        If strTmp1 <> "" Then
                            strBloodapplyrate = IIF(strBloodapplyrate = "", "", strBloodapplyrate & ",") & strTmp1
                        End If
                    End If
                Else
                    strBloodapplyrate = strTmp
                End If
            Next
            strTmp = strBloodapplyrate
            '如果是原值反回说有过程没有做处理，程序也不用再做处理
            If strTmp <> strPar Then
                strResult = ""
                If strTmp <> "" Then
                    arrTmp3 = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp3)
                        If arrTmp3(i) <> "" And Not IsDate(CStr(arrTmp3(i))) And arrTmp3(i) <> "无" Then
                            For j = 0 To UBound(arrTmp1)
                                strTmp = Split(arrTmp1(j), "<split1>")(5)
                                If strTmp = Split(arrTmp3(i), "|")(0) Then
                                    strResult = IIF(strResult = "", "", strResult & "<split3>") & arrTmp1(j)
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If
                arrTmp1 = Split(strResult, "<split3>")
                bln指定显示项 = True
            End If
            
            .Rows = Int((UBound(arrTmp1) + 1) / CON_LisResultCol) + IIF((UBound(arrTmp1) + 1) Mod CON_LisResultCol = 0, 0, 1)
            For i = 0 To UBound(arrTmp1)
                '加载指标
                arrTmp2 = Split(arrTmp1(i), "<split1>")
                .TextMatrix(Int(i / CON_LisResultCol), COL_指标中文名 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(5)
                .TextMatrix(Int(i / CON_LisResultCol), COL_结果单位 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                .TextMatrix(Int(i / CON_LisResultCol), COL_指标英文名 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(6)
                .TextMatrix(Int(i / CON_LisResultCol), COL_取值序列 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(7)
                .TextMatrix(Int(i / CON_LisResultCol), COL_指标代码 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(4)
                rsLIS.Filter = "编码='" & arrTmp2(1) & "'"
                If rsLIS.RecordCount > 0 Then
                    .TextMatrix(Int(i / CON_LisResultCol), COL_检验项目ID + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsLIS!检验项目ID & ""
                End If
                
                If bln指定显示项 Then
                    strTmp = arrTmp3(i)
                    If InStr(strTmp, "|") <> 0 Then
                        strTmp = Split(strTmp, "|")(1)
                    Else
                        strTmp = "1"
                    End If
                Else
                    strTmp = "1"
                End If
                
                '加载指标结果
                If arrTmp2(8) <> "" And strTmp = "1" Then
                    arrTmp2 = Split(arrTmp2(8), "<split2>")
                    .TextMatrix(Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(0)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_结果标志 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(1)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_结果参考 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                Else
                    '未提取到结果表示可以医生录入
                    .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
                End If
                
                lngCol = COL_指标中文名 + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 1755
                lngCol = COL_结果单位 + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 500
                lngCol = 1 + COL_检验项目ID + (i Mod CON_LisResultCol) * CON_LisResultCount
                If lngCol <> 29 Then .ColWidth(lngCol) = 50
                lngCol = COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 1120
            Next
            '116848,根据检验结果设置ABO和 RH
            Call CheckOrResetLisAboRH(True)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadLisResult(ByVal lng医嘱ID As Long, Optional ByVal strResult As String)
'功能：修改\查看申请单时，传入医嘱ID，加载已填写的指标
    Dim rsTmp As Recordset, strSQL As String
    Dim i As Long, j As Long, lngCol As Long
    Dim varCol As Variant
    Dim varRow As Variant
    Dim varFields As Variant
    
    '启用了血库时，用血申请不需要检查结果
    If mblnSpareBloood = False Then Exit Sub
    
    strSQL = "select 序号,检验项目ID,指标代码,指标中文名,指标英文名,指标结果,结果单位,结果标志,结果参考,取值序列,是否人工填写 from 输血检验结果 Where 医嘱ID=[1] order by 序号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    If strResult <> "" Then
        varFields = Array("序号", "检验项目ID", "指标代码", "指标中文名", "指标英文名", "指标结果", "结果单位", "结果标志", "结果参考", "取值序列", "是否人工填写")
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp, True)
        varRow = Split(strResult, "<SplitRow>")
        For i = 0 To UBound(varRow)
            varCol = Split(varRow(i), "<SplitCol>")
            rsTmp.AddNew
'            varFields , varCol
            For j = 0 To UBound(varCol)
                rsTmp.Fields(j).value = varCol(j)
            Next
            rsTmp.Update
        Next
        rsTmp.MoveFirst
        rsTmp.Sort = "序号"
    End If
    
    With vsLIS
        .Clear
        .Rows = Int((rsTmp.RecordCount) / CON_LisResultCol) + IIF((rsTmp.RecordCount) Mod CON_LisResultCol = 0, 0, 1)
        For i = 0 To rsTmp.RecordCount - 1
            '加载指标
            .TextMatrix(Int(i / CON_LisResultCol), COL_指标中文名 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!指标中文名 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_结果单位 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!结果单位 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_指标英文名 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!指标英文名 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_取值序列 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!取值序列 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_指标代码 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!指标代码 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_检验项目ID + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!检验项目ID & ""
            '加载指标结果
            .TextMatrix(Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!指标结果 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_结果标志 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!结果标志 & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_结果参考 + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!结果参考 & ""

            '手工录入的可以修改
            If Val(rsTmp!是否人工填写 & "") = 1 Then
                .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
            End If
            
            lngCol = COL_指标中文名 + (i Mod CON_LisResultCol) * CON_LisResultCount
            .ColWidth(lngCol) = 1755
            lngCol = COL_结果单位 + (i Mod CON_LisResultCol) * CON_LisResultCount
            .ColWidth(lngCol) = 500
            lngCol = 1 + COL_检验项目ID + (i Mod CON_LisResultCol) * CON_LisResultCount
            If lngCol <> 29 Then .ColWidth(lngCol) = 50
            lngCol = COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount
            .ColWidth(lngCol) = 1120
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("当前申请单已经进行了调整尚未保存，是否要继续退出？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    mlng输血途径 = 0
    mlng输血项目ID = 0
    mlng输血执行性质 = 0
    mlng执行科室性质 = 0
    mbln补录 = False
    mstr入院时间 = ""
    mlng录入限量 = 0
    mstr上次转科时间 = ""
    mint险类 = 0
    mstr诊断IDs = ""
    mstrLISAboRHCode = ""
    Set mclsMipModule = Nothing
End Sub

Private Sub lblInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = lbl本次历史申请项目 Then
        Call picHisItem_MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub optAppraise_Click(Index As Integer)
    If Visible And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub optAppraise_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub optConsent_Click(Index As Integer)
    If Visible And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub optConsent_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub optHistory_Click(Index As Integer)
    Dim i As Integer
    
    If Visible And mblnDataLoad = False Then mblnChange = True
    If Visible And Index = 0 Or Index = 1 Then
        For i = 2 To 5
            optHistory(i).Enabled = IIF(optHistory(1).value = True, True, False)
        Next
    End If
    If optHistory(0).value = True Then
        optHistory(2).value = True
        optHistory(4).value = True
    End If
End Sub

Private Sub optHistory_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub optPossession_Click(Index As Integer)
    If Visible And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub optPossession_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub picHisItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    If lblInfo(lbl本次历史申请项目).Width > picHisItem.Width Then
        strInfo = lblInfo(lbl本次历史申请项目).Tag
    Else
        strInfo = ""
    End If
    Call zlCommFun.ShowTipInfo(picHisItem.hwnd, strInfo, True)
End Sub

Private Sub txtGet_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtGet(Index)
End Sub

Private Sub txtGet_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call TxtGetInfo(Index, 0)
    End If
End Sub

Private Sub txtGet_Validate(Index As Integer, Cancel As Boolean)
    '恢复人为的清除
    If txtGet(Index).Text <> txtGet(Index).Tag Then
        txtGet(Index).Text = txtGet(Index).Tag
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Dim intIme As Integer  '1-打开,2-关闭,0不调用关闭或打开输入发
    If Index = txt预定输血时间 Then
'        If txtInfo(Index).Text = "" Then txtInfo(Index).Text = txtInfo(txt申请日期).Text
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt申请日期 Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt孕 Or Index = txt产 Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt备注 Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 1
    ElseIf Index = txt诊断信息 Then
        intIme = 1
    End If
    
    If intIme <> 0 Then
        On Error Resume Next
        Call zlCommFun.OpenIme(intIme = 1)
        If err <> 0 Then err.Clear
    End If
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''''
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case txt预定输血时间
            Call cmdDate_Click(0)
        Case txt预定输血时间
            Call cmdDate_Click(1)
        End Select
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case txt预定输血量
            If InStr("1234567890.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
            If KeyAscii <> 0 Then
                If Chr(KeyAscii) = "." Then
                    If InStr(txtInfo(Index).Text, ".") > 0 Then KeyAscii = 0
                End If
            End If
        Case txt孕, txt产
            If InStr("1234567890", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            Else
                If InStr("1234567890", Chr(KeyAscii)) > 0 Then
                    If Val(txtInfo(Index).Text) = 0 Then txtInfo(Index).Text = ""
                End If
            End If
        Case txt诊断信息
            Call zlControl.TxtCheckKeyPress(txtInfo(Index), KeyAscii, m文本式)
        Case txt备注
            If KeyAscii = vbKeyReturn Then
                If txtInfo(txt备注).Text <> "" Then
                    Call ReasonSelect(txtInfo(txt备注).Text)
                End If
            End If
    End Select
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    If Index = txt备注 Or Index = txt诊断信息 Then
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    If Index = txt预定输血时间 Then
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Tag) Then
                    '恢复人为的清除，缺省为上次填写的时间
                    txtInfo(Index).Text = txtInfo(Index).Tag
                End If
            End If
        Else
            '检查时间合法性
            If Not Check安排时间(txtInfo(Index).Text, txtInfo(txt申请日期).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            txtInfo(Index).Tag = txtInfo(Index).Text
        End If
    ElseIf Index = txt申请日期 Then
            
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(txt申请日期).Tag) Then
                    '恢复人为的清除
                    txtInfo(Index).Text = txtInfo(txt申请日期).Tag
                End If
            End If
        Else
            '检查时间合法性
            If Not Check开始时间(txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            '判断是否是补录医嘱
            If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint补录间隔 _
                Or mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
                mbln补录 = True
                SetControlEnabled cboInfo(cbo用血安排), False
            Else
                mbln补录 = False
                SetControlEnabled cboInfo(cbo用血安排), True
            End If
        End If
    ElseIf Index = txt诊断信息 Then
        If txtInfo(Index).Tag <> txtInfo(Index).Text Then
            mstr诊断IDs = ""
        End If
    ElseIf Index = txt备注 Then
        If zlCommFun.ActualLen(txtInfo(Index).Text) > 100 Then
            MsgBox "输入内容不过超过 50 个汉字或 100 个字符。", vbInformation, gstrSysName
            Call txtInfo_GotFocus(Index)
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub txt申请量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cboInfo(cbo单位).Enabled And cboInfo(cbo单位).Visible Then
            cboInfo(cbo单位).SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub vsfBlood_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String, arrInfo, arrItem, arrCode
    Dim i As Integer, j As Integer
    Dim blnSelect As Boolean
    Dim strIDs As String
    
    If NewRow < vsfBlood.FixedRows Or NewCol < vsfBlood.FixedCols Then Exit Sub
    With vsfList
        .Rows = 1
        .Redraw = flexRDNone
        If mblnSelectBlood = False Then
            If mblnSpareBloood = True Then '备血申请
                '【血液室】- 800ml<Split2>A型+:400ml<Split3>B型+:400ml<Split1>【LW医技中心A】- 800ml<Split2>A型+:400ml<Split3>B型+:400ml
                strTmp = vsfBlood.TextMatrix(NewRow, COL_P_库存)
                arrInfo = Split(strTmp, "<Split1>") '单个库房信息
                For i = 0 To UBound(arrInfo)
                    arrItem = Split(arrInfo(i), "<Split2>") '分解库房和血型信息
                    .AddItem arrItem(0)   '库房信息
                    .IsSubtotal(.Rows - 1) = True
                    .RowOutlineLevel(.Rows - 1) = 0
                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                     arrItem = Split(arrItem(1), "<Split3>") '截取血型库存信息
                     For j = 0 To UBound(arrItem)
                        .AddItem arrItem(j)   '血型及总量
                        .IsSubtotal(.Rows - 1) = True
                        .RowOutlineLevel(.Rows - 1) = 1
                     Next j
                Next i
            Else '用血申请
                '替代项目<Split2>品种ID'配发信息'待发量
                '其中配发信息格式为：配血总量：400ml 已发量：0ml 未发量：400ml<Split4> 规格：200ml(未发)  效期:2016-09-17 16:13<Split3>0<Split4> 规格：200ml(未发)  效期:2016-08-14 11:17<Split3>0
                strTmp = vsfBlood.TextMatrix(NewRow, COL_P_库存)
                If strTmp <> "" Then
                    arrInfo = Split(strTmp, "<Split2>")
                    If UBound(arrInfo) > 0 Then
                        arrCode = Split(Split(arrInfo(1), "'")(1), "<Split4>") '分解配血血液和规格信息
                        If UBound(arrCode) >= 0 Then
                            .AddItem arrCode(0)  '血液
                            .IsSubtotal(.Rows - 1) = True
                            .RowOutlineLevel(.Rows - 1) = 0
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                            For j = 1 To UBound(arrCode)
                                arrItem = Split(arrCode(j), "<Split3>")
                                .AddItem arrItem(0)  '规格信息
                                .IsSubtotal(.Rows - 1) = True
                                .RowOutlineLevel(.Rows - 1) = 1
                                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = IIF(Val(arrItem(1)) = 0, &H80000008, &H8000000C)
                            Next
                        End If
                    End If
                End If
            End If
        Else '医生选择品种后，则可看到输血科配置的血液信息
            '格式：收发ID<Split1>血袋编号<Split1>规格<Split1>效期,多条配血记录之间用<Split4>相连。最后拼接<Split3>内容为已经选择的血液ID
            strIDs = ""
            strTmp = vsfBlood.TextMatrix(NewRow, COL_P_库存)
            If strTmp <> "" Then
                If InStr(1, strTmp, "<Split3>") <> 0 Then
                    arrInfo = Split(strTmp, "<Split3>") '首先将已配的血液和已选择的血液分开
                    strTmp = arrInfo(0)
                    strIDs = arrInfo(1)
                End If
                arrInfo = Split(strTmp, "<Split4>") '已配血液信息
                For i = 0 To UBound(arrInfo)
                    .Rows = .Rows + 1
                    arrItem = Split(arrInfo(i), "<Split1>") '分解血液信息
                    .TextMatrix(.Rows - 1, COL_S_ID) = Val(arrItem(0))
                    .TextMatrix(.Rows - 1, COL_S_选择) = ""
                    .TextMatrix(.Rows - 1, COL_S_编号) = arrItem(1)
                    .TextMatrix(.Rows - 1, COL_S_规格) = arrItem(2)
                    .TextMatrix(.Rows - 1, COL_S_效期) = Format(arrItem(3), "YYYY-MM-DD HH:mm")
                    blnSelect = InStr(1, "|" & strIDs & "|", "|" & .TextMatrix(.Rows - 1, COL_S_ID) & "|") <> 0
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_选择) = img16.ListImages(IIF(blnSelect = True, "c1", "c0")).Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_选择) = IIF(blnSelect = True, 1, 0)
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_选择, .Rows - 1, COL_S_效期) = blnSelect
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_选择, .Rows - 1, COL_S_效期) = IIF(blnSelect = True, &HC0E0FF, vbWhite)
                Next
                '将即将失效的血液放在前面
                .Cell(flexcpSort, .FixedRows, COL_S_效期, .Rows - 1, COL_S_效期) = 1
            End If
        End If
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    Dim intValue As Integer, blnNext As Boolean
    Dim i As Integer
    Dim dblSum As Double, strIDs As String
    Dim strTmp As String, arrInfo
    If mblnSelectBlood = False Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        blnNext = False
        With vsfList
            If .Editable = flexEDNone Then GoTo GONextControl
            If vsfBlood.Row < vsfBlood.FixedRows Then GoTo GONextControl
            If vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_选择) = 1 Then
                If .Row < .FixedRows And .Rows > .FixedRows Then
                    blnNext = True
                    .Row = .FixedRows
                    .Col = COL_S_选择
                    .ShowCell .FixedRows, COL_P_选择
                ElseIf .Row < .Rows - 1 Then
                    blnNext = True
                    .Row = .Row + 1
                    .Col = COL_S_选择
                    .ShowCell .Row, COL_S_选择
                End If
            End If
            If blnNext = False Then
                If vsfBlood.Row < vsfBlood.Rows - 1 Then
                    blnNext = True
                    vsfBlood.Row = vsfBlood.Row + 1
                    vsfBlood.Col = COL_P_选择
                    vsfBlood.ShowCell vsfBlood.Row, COL_P_选择
                    If vsfBlood.Enabled And vsfBlood.Visible Then vsfBlood.SetFocus
                End If
            End If
GONextControl:
            If blnNext = False Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        End With
    ElseIf KeyAscii = vbKeySpace Then
        With vsfList
            '必须选择血液品种后才能更改血液信息
            If vsfBlood.Row < vsfBlood.FixedRows Then Exit Sub
            If Val(vsfBlood.TextMatrix(vsfBlood.Row, COL_P_ID)) = 0 Then Exit Sub
            If vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_选择) = 0 Then Exit Sub
            If .Row >= .FixedRows And .Editable <> flexEDNone Then
                If Val(.TextMatrix(.Row, COL_S_ID)) = 0 Then Exit Sub
                intValue = Val(.Cell(flexcpData, .Row, COL_S_选择))
                Set .Cell(flexcpPicture, .Row, COL_S_选择, .Row, COL_S_选择) = img16.ListImages("c" & IIF(intValue = 1, "0", "1") & "").Picture
                .Cell(flexcpData, .Row, COL_S_选择, .Row, COL_S_选择) = IIF(intValue = 1, 0, 1)
                .Cell(flexcpFontBold, .Row, COL_S_选择, .Row, COL_S_效期) = IIF(intValue = 1, False, True)
                .Cell(flexcpBackColor, .Row, COL_S_选择, .Row, COL_S_效期) = IIF(intValue = 1, vbWhite, &HC0E0FF)
                strIDs = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_S_选择, i, COL_S_选择) = 1 Then
                        dblSum = dblSum + Val(.TextMatrix(i, COL_S_规格))
                        strIDs = strIDs & "|" & Val(.TextMatrix(i, COL_S_ID))
                    End If
                Next
                If Left(strIDs, 1) = "|" Then strIDs = Mid(strIDs, 2)
                strTmp = vsfBlood.TextMatrix(vsfBlood.Row, COL_P_库存)
                If InStr(1, strTmp, "<Split3>") <> 0 Then
                    arrInfo = Split(strTmp, "<Split3>")
                    strTmp = arrInfo(0)
                End If
                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_库存) = strTmp & IIF(strIDs <> "", "<Split3>" & strIDs, "")
                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请量) = IIF(dblSum = 0, "", dblSum)
                mblnChange = True
                Call BloodSum
            End If
        End With
    End If
End Sub

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandle
    If mblnSelectBlood = True And Button = 1 Then
        If Not (vsfList.Rows > vsfList.FixedRows And vsfList.MouseRow >= vsfList.FixedRows) Then Exit Sub
        If vsfList.Col = COL_S_选择 And vsfList.MouseCol = COL_S_选择 And Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择)) = 1 Then
            If X <= vsfList.CellLeft + 255 Then
                Call vsfList_KeyPress(vbKeySpace)
            End If
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '不允许编辑任何内容
    Cancel = True
End Sub

Private Sub vsfBlood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Visible Then mblnChange = True
    If Col = COL_P_申请量 Then Call BloodSum
End Sub

Private Sub vsfBlood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_P_申请量 Then
        vsfBlood.EditSelStart = 0
        vsfBlood.EditSelLength = Len(vsfBlood.TextMatrix(Row, Col))
        '关闭输入法
        On Error Resume Next
        Call zlCommFun.OpenIme
        If err <> 0 Then err.Clear
    End If
End Sub

Private Sub vsfBlood_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_P_选择 Then Cancel = True
End Sub

Private Sub vsfBlood_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer
    Dim blnFind As Boolean
    Dim blnNext As Boolean
    If KeyCode = vbKeyReturn Then
        Select Case Col
            Case COL_P_申请量
                vsfBlood.TextMatrix(Row, Col) = vsfBlood.EditText
            Case COL_P_申请血型, COL_P_申请RH
                If vsfBlood.ColHidden(Col) = False Then
                    vsfBlood.TextMatrix(Row, Col) = vsfBlood.ComboItem(vsfBlood.ComboIndex)
                End If
        End Select
        For i = Col + 1 To vsfBlood.Cols - 1
            If (i = COL_P_申请血型 Or i = COL_P_申请RH Or i = COL_P_申请量) And vsfBlood.ColHidden(i) = False Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = True Then
            vsfBlood.Col = i
            vsfBlood.ShowCell vsfBlood.Row, i
            blnNext = True
        Else
            If vsfBlood.Row < vsfBlood.Rows - 1 Then
                vsfBlood.Row = vsfBlood.Row + 1
                vsfBlood.Col = COL_P_选择
                vsfBlood.ShowCell vsfBlood.Row, COL_P_选择
                blnNext = True
            End If
        End If
        If blnNext = False Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End If
End Sub

Private Sub vsfBlood_KeyPress(KeyAscii As Integer)
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean
    Dim blnNext As Boolean, blnOne As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim intValue As Integer
    Dim blnSetUnit As Boolean
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        blnNext = False
        If vsfBlood.Row < vsfBlood.FixedRows And vsfBlood.Rows > vsfBlood.FixedRows Then
            blnNext = True
            vsfBlood.Row = vsfBlood.FixedRows
            vsfBlood.Col = COL_P_选择
            vsfBlood.ShowCell vsfBlood.FixedRows, COL_P_选择
        ElseIf vsfBlood.Row <= vsfBlood.Rows - 1 Then
            For j = vsfBlood.Col + 1 To vsfBlood.Cols - 1
                If (j = COL_P_申请量 Or j = COL_P_申请血型 Or j = COL_P_申请RH) And vsfBlood.ColHidden(j) = False Then
                    blnFind = True
                    Exit For
                End If
            Next
            If blnFind = True Then
                blnNext = True
                vsfBlood.Col = j
                vsfBlood.ShowCell vsfBlood.Row, j
            Else
                If vsfBlood.Row < vsfBlood.Rows - 1 Then
                    blnNext = True
                    vsfBlood.Row = vsfBlood.Row + 1
                    vsfBlood.Col = COL_P_选择
                    vsfBlood.ShowCell vsfBlood.Row, COL_P_选择
                End If
            End If
        End If
        If blnNext = False Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If vsfBlood.Editable <> flexEDNone Then
            txt申请信息.Text = "品种:"
            txt申请量.Text = ""
            cboInfo(cbo单位).ListIndex = -1
            If vsfBlood.Col = COL_P_选择 And vsfBlood.Row >= vsfBlood.FixedRows Then
                intValue = Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择))
                Set vsfBlood.Cell(flexcpPicture, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_选择) = img16.ListImages("c" & IIF(intValue = 1, "0", "1") & "").Picture
                vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_选择) = IIF(intValue = 1, 0, 1)
                vsfBlood.Cell(flexcpFontBold, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_库存) = IIF(intValue = 1, False, True)
                vsfBlood.Cell(flexcpBackColor, vsfBlood.Row, COL_P_选择, vsfBlood.Row, COL_P_库存) = IIF(intValue = 1, vbWhite, &HC0E0FF)
                
                If Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择)) = 1 Then
                    '申请血型和RH继承上一个
                    If vsfBlood.ColHidden(COL_P_申请血型) = False Then
                        For i = vsfBlood.Row - 1 To vsfBlood.FixedRows Step -1
                            If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请血型) = vsfBlood.TextMatrix(i, COL_P_申请血型)
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请RH) = vsfBlood.TextMatrix(i, COL_P_申请RH)
                                Exit For
                            End If
                            If vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请血型) = "" Then
                                If InStr(1, ",A,B,O,AB,", "," & cboInfo(cbo输血血型).Text & ",") <> 0 Then
                                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请血型) = cboInfo(cbo输血血型).Text
                                End If
                            End If
                            If vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请RH) = "" Then
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请RH) = cboInfo(cboRHD).Text
                            End If
                        Next
                    End If
                Else
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请量) = ""
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请血型) = ""
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_申请RH) = ""
                    '医生选择血液的模式，如果血液品种取消选择，则同步取消之前选择的血液信息
                    If mblnSelectBlood = True Then
                        For i = vsfList.FixedRows To vsfList.Rows - 1
                            If Val(vsfList.Cell(flexcpData, i, COL_S_选择, i, COL_S_选择)) = 1 Then
                                vsfList.Cell(flexcpPicture, i, COL_S_选择, i, COL_S_选择) = img16.ListImages("c0").Picture
                                vsfList.Cell(flexcpData, i, COL_S_选择, i, COL_S_选择) = 0
                                vsfList.Cell(flexcpFontBold, i, COL_S_选择, i, COL_S_效期) = False
                                vsfList.Cell(flexcpBackColor, i, COL_S_选择, i, COL_S_效期) = vbWhite
                            End If
                        Next
                        If InStr(1, vsfBlood.TextMatrix(vsfBlood.Row, COL_P_库存), "<Split3>") <> 0 Then
                            vsfBlood.TextMatrix(vsfBlood.Row, COL_P_库存) = Split(vsfBlood.TextMatrix(vsfBlood.Row, COL_P_库存), "<Split3>")(0)
                        End If
                    End If
                End If
                '设置血液和执行科室
                With rsTmp
                    If .State = 1 Then .Close
                    .Fields.Append "ID", adBigInt
                    .Fields.Append "编码", adVarChar, 20, adFldIsNullable
                    .Fields.Append "名称", adVarChar, 60, adFldIsNullable
                    .Fields.Append "计算单位", adVarChar, 20, adFldIsNullable
                    .Fields.Append "执行分类ID", adBigInt
                    .Fields.Append "执行科室ID", adBigInt
                    .Fields.Append "录入限量ID", adBigInt
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
                        If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
                            .AddNew
                            .Fields("ID") = Val(vsfBlood.TextMatrix(i, COL_P_ID))
                            .Fields("编码") = IIF(vsfBlood.TextMatrix(i, COL_P_编码) = "", Null, vsfBlood.TextMatrix(i, COL_P_编码))
                            .Fields("名称") = IIF(vsfBlood.TextMatrix(i, COL_P_名称) = "", Null, vsfBlood.TextMatrix(i, COL_P_名称))
                            .Fields("计算单位") = IIF(vsfBlood.TextMatrix(i, COL_P_单位) = "", Null, vsfBlood.TextMatrix(i, COL_P_单位))
                            .Fields("执行分类ID") = Val(vsfBlood.TextMatrix(i, COL_P_执行分类ID))
                            .Fields("执行科室ID") = Val(vsfBlood.TextMatrix(i, COL_P_执行科室ID))
                            .Fields("录入限量ID") = Val(vsfBlood.TextMatrix(i, COL_P_录入限量ID))
                            .Update
                            'iif(mid("" & 0.5,1,1)=".","0","") & 0.5，这种写法是为了保证小于的1的值能正常显示前缀0
                            txt申请信息.Text = txt申请信息.Text & "[" & vsfBlood.TextMatrix(i, COL_P_名称) & IIF(vsfBlood.TextMatrix(i, COL_P_申请量) <> "", "-" & IIF(Mid("" & vsfBlood.TextMatrix(i, COL_P_申请量), 1, 1) = ".", "0", "") & vsfBlood.TextMatrix(i, COL_P_申请量) & vsfBlood.TextMatrix(i, COL_P_单位), "") & "]"
                            '选择品种时缺省单位设置，规则如下：
                            '1、如果设置的品种单位中包含ML，则缺省设置单位为ML
                            '2、如果设置的品种单位中不包含ML，则缺省单位为到第一个品种的单位
                            blnSetUnit = False
                            If cboInfo(cbo单位).ListIndex = -1 Then
                                blnSetUnit = True
                            Else
                                If UCase(cboInfo(cbo单位).List(cboInfo(cbo单位).ListIndex)) <> "ML" And UCase(vsfBlood.TextMatrix(i, COL_P_单位)) = "ML" Then
                                    blnSetUnit = True
                                End If
                            End If
                            If blnSetUnit = True Then
                                For j = 0 To cboInfo(cbo单位).ListCount - 1
                                    If UCase(vsfBlood.TextMatrix(i, COL_P_单位)) = UCase(cboInfo(cbo单位).List(j)) Then
                                        Call zlControl.CboSetIndex(cboInfo(cbo单位).hwnd, j)
                                        cboInfo(cbo单位).Tag = j
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Call BloodSum
                    End If
                End With
                Call SetTxtBloodInfo(rsTmp, txt预定输血成分, False)
                Call RsetBreedUnit
            End If
        End If
    End If
End Sub

Private Sub vsfBlood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = COL_P_申请量 Then
        If InStr("1234567890.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
        If KeyAscii <> 0 Then
            If Chr(KeyAscii) = "." Then
                If InStr(vsfBlood.EditText, ".") > 0 Then KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub vsfBlood_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandle
    If Button <> 1 Then Exit Sub
    If vsfBlood.Col = COL_P_选择 And vsfBlood.MouseCol = COL_P_选择 Then
        If X <= vsfBlood.CellLeft + 255 Then
            Call vsfBlood_KeyPress(vbKeySpace)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfBlood_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnSpareBloood = True Then
        If Not ((Col = COL_P_申请量 Or Col = COL_P_申请血型 Or Col = COL_P_申请RH) And vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择) = 1 And vsfBlood.ColHidden(Col) = False) Then
            Cancel = True
            Exit Sub
        End If
    Else
        If Not (Col = COL_P_申请量 And vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_选择) = 1) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub vsLIS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Visible Then
        mblnChange = True
        '编辑检验结果ABO和RH时，重新设置ABO和RH选项值内容
        Call CheckOrResetLisAboRH(True, Row, Col)
    End If
End Sub

Private Sub vsLIS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    If mintType = 2 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    With vsLIS
        If .Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor Then
            .Editable = flexEDKbdMouse
            .FocusRect = flexFocusSolid
            If .TextMatrix(NewRow, NewCol + (COL_取值序列 - COL_指标结果)) <> "" Then
                '老版和新版用的分割符不同，新版是逗号，老版分号，做下兼容处理。
                strTmp = .TextMatrix(NewRow, NewCol + (COL_取值序列 - COL_指标结果))
                strTmp = Replace(strTmp, ";", "|")
                strTmp = Replace(strTmp, ",", "|")
                .ComboList = strTmp & "|已查未回报"
            Else
                .ComboList = ""
            End If
        Else
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsLIS_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long

    With vsLIS
        If Col <> CON_LisResultCount - 1 And Col <> CON_LisResultCount * 2 - 1 Then Exit Sub

        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
            '为了支持预览输出
            If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        
        Done = True
    End With
End Sub

Private Sub vsLIS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EnterNextCell
    End If
End Sub

Private Sub EnterNextCell()
'功能：输框定位到下一个
    With vsLIS
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '如果是隐藏行则递归再定位到下一个位置
        If .Cell(flexcpBackColor, .Row, .Col) <> COLEditBackColor Then Call EnterNextCell
        .ShowCell .Row, .Col
    End With
End Sub

Private Function SaveCacheData() As Boolean
'功能：缓存数据
    Dim strResult As String
    Dim rsCard As ADODB.Recordset
    Dim curDate As Date
    Dim str检验项目SQL As String
    Dim str诊断关联信息SQL As String
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long, j As Long
    Dim var1 As Variant
    Dim var2 As Variant
    Dim str滴速 As String
    Dim str项目名称 As String, str申请项目SQL As String
    
    If cboInfo(cbo滴速).Visible = True Then
        str滴速 = cboInfo(cbo滴速).Text
    End If
    If IsNumeric(str滴速) = True Then
        str滴速 = str滴速 & "滴/分钟"
    End If
    
    var1 = Array()
    var2 = Array()
    '检验项目
    With vsLIS
        lngCount = 0
        For i = 0 To .Rows - 1
            For j = 0 To CON_LisResultCol - 1
                If Val(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) <> 0 Then
                    lngCount = lngCount + 1
                    strTmp = "Zl_输血检验结果_Insert([相关ID]," & lngCount & "," & ZVal(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) & ",'" & .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_指标英文名 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_结果单位 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_结果标志 + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_结果参考 + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_取值序列 + (j * CON_LisResultCount)) & "'," & IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0) & ")"
                    str检验项目SQL = str检验项目SQL & "<splitSQL>" & strTmp
                    
                    var1 = Array(lngCount, .TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标英文名 + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_结果单位 + (j * CON_LisResultCount)), .TextMatrix(i, COL_结果标志 + (j * CON_LisResultCount)), .TextMatrix(i, COL_结果参考 + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_取值序列 + (j * CON_LisResultCount)), IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0))
                    strTmp = Join(var1, "<SplitCol>")
                    ReDim Preserve var2(UBound(var2) + 1)
                    var2(UBound(var2)) = strTmp
                End If
            Next
        Next
    End With
    strResult = Join(var2, "<SplitRow>")
    '诊断关联信息
    If mstr诊断IDs <> "" Then
        str诊断关联信息SQL = "Zl_病人诊断医嘱_Insert([相关ID],'" & mstr诊断IDs & "')"
        str诊断关联信息SQL = str诊断关联信息SQL & "<splitSQL>" & "Zl_病人医嘱附件_Insert([相关ID],'申请单诊断',null,null,null,'" & txtInfo(txt诊断信息).Text & "',1)"
    ElseIf Trim(txtInfo(txt诊断信息).Text) <> "" Then
        str诊断关联信息SQL = "Zl_病人医嘱附件_Insert([相关ID],'申请单诊断',null,null,null,'" & txtInfo(txt诊断信息).Text & "',1)"
    End If
    
    '申请项目SQL
    '申请内容插入医嘱申请附加项目
    str项目名称 = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
            str项目名称 = IIF(str项目名称 = "", "", str项目名称 & Space(2)) & vsfBlood.TextMatrix(i, COL_P_名称) & ":" & IIF(vsfBlood.TextMatrix(i, COL_P_申请血型) = "", "", vsfBlood.TextMatrix(i, COL_P_申请血型) & vsfBlood.TextMatrix(i, COL_P_申请RH)) & " " & vsfBlood.TextMatrix(i, COL_P_申请量) & vsfBlood.TextMatrix(i, COL_P_单位)
        End If
    Next
    If str项目名称 <> "" Then
        str申请项目SQL = "Zl_病人医嘱附件_Insert([相关ID],'申请项目',null,2,null,'" & str项目名称 & "')"
    End If
    
    If mrsCard Is Nothing Then
         Call InitCardRsBlood(mrsCard)
         mrsCard.AddNew
    End If
    
    With mrsCard
        !用血安排 = cboInfo(cbo用血安排).ListIndex
        !临床诊断IDs = mstr诊断IDs
        !待诊 = chkWait.value
        !输血类型 = cboInfo(cbo输血类型).Text
        !输血目的 = cboInfo(cbo输血目的).Text
        !输血性质 = cboInfo(cbo输血性质).ListIndex
        !即往输血史 = IIF(optHistory(0).value, 0, 1)
        !既往输血反应史 = IIF(optHistory(2).value, 0, 1)
        !输血禁忌及过敏史 = IIF(optHistory(4).value, 0, 1)
        !孕产情况 = txtInfo(txt孕) & "/" & txtInfo(txt产)
        !受血者属地 = IIF(optPossession(0).value, 0, 1)
        !是否签订同意书 = IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, Null))
        !是否已评估 = IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, Null))
        !预定输血日期 = txtInfo(txt预定输血时间).Text
        !血型 = cboInfo(cbo输血血型).ListIndex
        !RHD = cboInfo(cboRHD).ListIndex
        !输血项目ID = mlng输血项目ID
        !输血执行科室ID = IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, 0, cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex))
        !预定输血量 = Val(txtInfo(txt预定输血量).Text)
        !输血途径项目ID = mlng输血途径
        !输血途径执行科室ID = IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, 0, cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex))
        !备注 = txtInfo(txt备注).Text
        !滴速 = str滴速
        !输血申请日期 = txtInfo(txt申请日期).Text
        !申请科室id = mlng开单科室ID
        !临床诊断描述 = txtInfo(txt诊断信息).Text
        !检查结果 = strResult
        !申请项目 = GetBloodInfo
        !申请其他项目SQL = "Zl_输血申请记录_Insert([相关ID]," & chkWait.value & ",'" & cboInfo(cbo输血类型).Text & "','" & cboInfo(cbo输血目的).Text & "'," & cboInfo(cbo输血性质).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optHistory(2).value, 0, 1) & "," & IIF(optHistory(4).value, 0, 1) & ",'" & txtInfo(txt孕) & "/" & txtInfo(txt产) & "'," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo输血血型).ListIndex & "," & cboInfo(cboRHD).ListIndex & "," & IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, "Null")) & "," & IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, "Null")) & ",'" & !申请项目 & "')"
        !检验项目SQL = str检验项目SQL
        !诊断关联信息SQL = str诊断关联信息SQL
        !申请项目SQL = str申请项目SQL
        .Update
    End With
    SaveCacheData = True
    mblnChange = False
End Function

Private Sub LoadDataFromCache()
'功能：通过缓存数据加载界面
    Dim str诊断  As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    Dim strResult As String, strTmp As String
    Dim blnDo As Boolean
    Dim arrItem, strIDs As String
    Dim i As Integer
    On Error GoTo errH
    
    If Not mrsCard Is Nothing Then
        If Not mrsCard.EOF Then
            blnDo = True
        End If
    End If
    
    If blnDo Then
        With mrsCard
            cboInfo(cbo用血安排).ListIndex = IIF(1 = Val(!用血安排 & ""), 1, 0)
            If Val(!待诊 & "") = 1 Then
                txtInfo(txt诊断信息).Text = "待诊"
                chkWait.value = 1
            Else
               '读取诊断
                mstr诊断IDs = !临床诊断IDs & ""
                txtInfo(txt诊断信息).Text = !临床诊断描述 & ""
            End If
            txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
            chkWait.value = Val(!待诊 & "")
            If !输血类型 & "" <> "" Then
                Call zlControl.CboSetText(cboInfo(cbo输血类型), !输血类型 & "", True, "'")
            End If
            If !输血目的 & "" <> "" Then
                Call zlControl.CboSetText(cboInfo(cbo输血目的), !输血目的 & "", True, "'")
'                Call zlControl.CboSetIndex(cboInfo(cbo输血执行).hWnd, 0)
            End If
            txtInfo(txt预定输血时间).Text = !预定输血日期 & ""
            txtInfo(txt预定输血时间).Tag = txtInfo(txt预定输血时间).Text
            cboInfo(cbo输血性质).ListIndex = Val(!输血性质 & "")
            optHistory(Val(!即往输血史 & "")).value = True
            optHistory(IIF(Val(!既往输血反应史 & "") = 1, 3, 2)).value = True
            optHistory(IIF(Val(!输血禁忌及过敏史 & "") = 1, 5, 4)).value = True
            If InStr(1, "" & !孕产情况, "/") <= 0 Then
                txtInfo(txt孕).Text = ""
                txtInfo(txt产).Text = ""
            Else
                txtInfo(txt孕).Text = Mid(!孕产情况, 1, InStr(1, "" & !孕产情况, "/") - 1)
                If Not (txtInfo(txt孕).Text = "" Or IsNumeric(txtInfo(txt孕).Text)) Then
                    txtInfo(txt孕).Text = ""
                End If
                txtInfo(txt产).Text = Mid(!孕产情况, InStr(1, "" & !孕产情况, "/") + 1)
                If Not (txtInfo(txt产).Text = "" Or IsNumeric(txtInfo(txt产).Text)) Then
                    txtInfo(txt产).Text = ""
                End If
            End If
            optPossession(Val(!受血者属地 & "")).value = True
            If InStr(1, ",0,1,", "," & !是否签订同意书 & ",") <> 0 Then
                optConsent(Val(!是否签订同意书 & "")).value = True
            End If
            If InStr(1, ",0,1,", "," & !是否已评估 & ",") <> 0 Then
                optAppraise(Val(!是否已评估 & "")).value = True
            End If
        
            cboInfo(cbo输血血型).ListIndex = Val(!血型 & "")
            cboInfo(cboRHD).ListIndex = Val(!RHD & "")
            
            mstr输血项目 = !申请项目 & ""
            mlng输血项目ID = Val(!输血项目ID & "")
            
            strIDs = ""
            arrItem = Split(mstr输血项目, ";")
            For i = 0 To UBound(arrItem)
                strIDs = strIDs & "," & Split(CStr(arrItem(i)), ",")(0)
            Next
            strIDs = Mid(strIDs, 2)
            If InStr(1, "," & strIDs & ",", "," & mlng输血项目ID & ",") = 0 Then
                If strIDs <> "" Then
                    strIDs = mlng输血项目ID & "," & strIDs
                    mstr输血项目 = mlng输血项目ID & "," & Val(!预定输血量 & "") & ",,;" & mstr输血项目
                Else
                    strIDs = mlng输血项目ID
                    mstr输血项目 = mlng输血项目ID & "," & Val(!预定输血量 & "") & ",,"
                End If
            End If
            lngTmp = Val(!输血执行科室ID & "")
            Set rsTmp = Get诊疗项目记录(mlng输血项目ID, strIDs)
            strTmp = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & IIF(strTmp = "", "", "'") & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txtGet(txt预定输血成分).Text = strTmp
            rsTmp.Filter = "ID=" & mlng输血项目ID
            Call Set执行科室(Val(rsTmp!执行科室 & ""), lngTmp)
            txtInfo(txt单位).Text = rsTmp!计算单位 & ""
            txtGet(txt预定输血成分).Tag = txtGet(txt预定输血成分).Text
        
            mlng录入限量 = Val(rsTmp!录入限量 & "")
            mlng输血途径 = Val(!输血途径项目ID & "")
            lngTmp = Val(!输血途径执行科室ID & "")
            Set rsTmp = Get诊疗项目记录(mlng输血途径)
            If Not (rsTmp!类别 = "E" And rsTmp!操作类型 = "9") Then
                mblnNewSpareBloood = False
                If rsTmp!类别 = "E" And rsTmp!操作类型 = "8" Then
                    mblnSpareBloood = (Val(rsTmp!执行分类 & "") = 0)
                End If
            Else
                mblnNewSpareBloood = True
                mblnSpareBloood = True '诊疗类别为E,操作类型=9的就是备血医嘱
            End If
            txtGet(txt输血途径).Text = rsTmp!名称 & ""
            txtGet(txt输血途径).Tag = txtGet(txt输血途径).Text
            Call Set输血执行(Val(rsTmp!执行科室 & ""), lngTmp)
            
            txtInfo(txt预定输血量).Text = zl9ComLib.FormatEx((!预定输血量 & ""), 5)
            txtInfo(txt备注).Text = !备注 & ""
            txtInfo(txt申请日期).Text = !输血申请日期 & ""
            
            Call SetLisResult(strIDs)
            '用血医嘱滴速
            strTmp = !滴速 & ""
            cboInfo(cbo滴速).Text = ""
            lblInfo(31).Visible = True
            If strTmp Like "*滴/分钟" Then
                If IsNumeric(Split(strTmp, "滴/分钟")(0)) = True Then
                    cboInfo(cbo滴速).Text = Split(strTmp, "滴/分钟")(0)
                End If
            ElseIf strTmp = "加压" Or strTmp = "快速" Then
                cboInfo(cbo滴速).Text = strTmp
                lblInfo(31).Visible = False
            End If
            
            strResult = !检查结果 & ""
            If strResult <> "" Then
                Call LoadLisResult(0, strResult)
            End If
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetFormNature(Optional ByVal blnFormLoad As Boolean = True) As Boolean
'功能：根据申请类型设置窗体相关属性
    '检查类的放在最上面
    Dim bln用血申请 As Boolean '启用了血库，且是用血申请
    Dim arrItem, arrCode, i As Integer
    mblnSelectBlood = False
    
    bln用血申请 = mblnSpareBloood = False
    If bln用血申请 Then
        If mintType = 0 Then '如果是新增，则根据参数决定是否是医生选择品种
            mblnSelectBlood = gbln下达用血申请确定血液信息
        Else  '非新增模式，则根据输血项目内容决定是那种模式
            arrItem = Split(mstr输血项目, ";")
            For i = 0 To UBound(arrItem)
                arrCode = Split(arrItem(i), ",")
                If UBound(arrCode) > 3 Then
                    mblnSelectBlood = CStr(arrCode(4)) <> ""
                End If
            Next
        End If
    End If
    If blnFormLoad = False Then
        If InitInfo = False Then Exit Function
    End If
    If mblnSpareBloood = True Then
        lblHead.Caption = "临床输血申请单"
    Else
        lblHead.Caption = "临床取血通知单"
    End If
    If mblnSpareBloood = True And mblnNewSpareBloood = True Then
        lblInfo(lbl发血执行).Caption = "备血执行"
        lblInfo(lbl输血途径).Caption = "采集方法"
        lblInfo(lbl输血执行).Caption = "采集执行"
    Else
        lblInfo(lbl发血执行).Caption = "发血执行"
        lblInfo(lbl输血途径).Caption = "输血途径"
        lblInfo(lbl输血执行).Caption = "输血执行"
    End If
    '下方签名控件坐标调整
    If mint场合 = 0 And mblnSpareBloood Then
        lblInfo(lbl申请医师签名).Top = lblInfo(lbl申请医师坐标).Top
        Line1(lin申请医师签名).Y1 = Line1(lin申请医师坐标).Y1
        Line1(lin申请医师签名).Y2 = Line1(lin申请医师签名).Y1
        txtInfo(txt申请医师签名).Top = txtInfo(txt申请医师坐标).Top
        lblInfo(lbl主治医师签名).Visible = True
        Line1(lin主治医师签名).Visible = True
        txtInfo(txt主治医师签名).Visible = True
    Else
        lblInfo(lbl申请医师签名).Top = lblInfo(lbl主治医师签名).Top
        Line1(lin申请医师签名).Y1 = Line1(lin主治医师签名).Y1
        Line1(lin申请医师签名).Y2 = Line1(lin申请医师签名).Y1
        txtInfo(txt申请医师签名).Top = txtInfo(txt主治医师签名).Top
        lblInfo(lbl主治医师签名).Visible = False
        Line1(lin主治医师签名).Visible = False
        txtInfo(txt主治医师签名).Visible = False
    End If
    lblInfo(lbl采集者签名) = IIF(mblnSpareBloood = True, "采 集 者签名", "取 血 者签名")
    
    '控件位置调整(输血相关项目)
    On Error Resume Next
    lblInfo(23).Visible = bln用血申请 = False '受血者标签
    lblInfo(lbl既往输血史).Visible = bln用血申请 = False
    fraChk(fra既往输血史).Visible = bln用血申请 = False
    lblInfo(lbl既往输血反应史).Visible = bln用血申请 = False
    fraChk(fra既往输血反应史).Visible = bln用血申请 = False
    lblInfo(lbl输血禁忌及过敏史).Visible = bln用血申请 = False
    fraChk(fra输血禁忌及过敏史).Visible = bln用血申请 = False
    lblInfo(lbl孕产情况).Visible = bln用血申请 = False
    fraChk(fra孕产情况).Visible = bln用血申请 = False
    lblInfo(lbl受血者属地).Visible = bln用血申请 = False
    fraChk(fra受血者属地).Visible = bln用血申请 = False
    lblInfo(lbl知情同意书).Visible = bln用血申请 = False
    fraChk(fra知情同意书).Visible = bln用血申请 = False
    lblInfo(lbl输血评估).Visible = bln用血申请 = False
    fraChk(fra输血评估).Visible = bln用血申请 = False
    
    
    '预定输血日期
    lblInfo(lbl预定输血日期).Top = IIF(bln用血申请 = True, lblInfo(lbl孕产情况).Top, lblInfo(lbl知情同意书).Top + lblInfo(lbl知情同意书).Height + 210)
    txtInfo(txt预定输血时间).Top = lblInfo(lbl预定输血日期).Top - 30
    Line1(12).Y1 = txtInfo(txt预定输血时间).Top + txtInfo(txt预定输血时间).Height + 15
    Line1(12).Y2 = Line1(12).Y1
    cmdDate(cmd预定输血时间).Top = txtInfo(txt预定输血时间).Top
    '血型
    lblInfo(lbl血型).Top = lblInfo(lbl预定输血日期).Top
    picInfo(1).Top = txtInfo(txt预定输血时间).Top - 30
    Line1(13).Y1 = Line1(12).Y1
    Line1(13).Y2 = Line1(13).Y1
    'RH
    lblInfo(lblRHD).Top = lblInfo(lbl预定输血日期).Top
    picInfo(2).Top = picInfo(1).Top
    Line1(14).Y1 = Line1(13).Y1
    Line1(14).Y2 = Line1(14).Y1
    
    With picPreBlood
        .Left = lblInfo(lbl预定输血日期).Left
        .Top = lblInfo(lbl预定输血日期).Top + lblInfo(lbl预定输血日期).Height + 180
        .Visible = True
    End With
    
    With picBloodDept
        .Top = picPreBlood.Top - 30
        .Left = picPreBlood.Width + picPreBlood.Left - .Width
        .Visible = True
        .ZOrder 0
    End With
    '滴速用血申请才显示(目前暂时不用)
    lblInfo(30).Visible = bln用血申请
    picInfo(3).Visible = bln用血申请
    Line1(25).Visible = bln用血申请
    lblInfo(31).Visible = bln用血申请
    If bln用血申请 Then
        If cboInfo(cbo滴速).Text <> "" And IsNumeric(cboInfo(cbo滴速).Text) = False Then lblInfo(31).Visible = False
    End If
    '发血执行位置变动调整
    If bln用血申请 = False Then
        picInfo(8).Left = picBloodDept.Width - picInfo(8).Width - 30
    Else
        picInfo(8).Left = lblInfo(30).Left - picInfo(8).Width - 120
    End If
    Line1(16).X1 = picInfo(8).Left - 75
    Line1(16).X2 = picInfo(8).Left + picInfo(8).Width + 15
    lblInfo(lbl发血执行).Left = picInfo(8).Left - lblInfo(lbl发血执行).Width - 120
    
    '输血成分和输血量控制
    vsLIS.Visible = bln用血申请 = False
    lblInfo(lbl预定输血成分).Visible = False
    txtGet(txt预定输血成分).Visible = False
    txtGet(txt预定输血成分).Locked = True
    txtGet(txt预定输血成分).BackColor = &H8000000F
    picGet(0).Visible = False
    lblInfo(lbl预定输血量).Visible = False
    txtInfo(txt预定输血量).Visible = False
    txtInfo(txt单位).Visible = False
    Line1(15).Visible = False
    Line1(17).Visible = False
    If bln用血申请 = False Then
        '预定输血成分
        lblInfo(lbl预定输血成分).Top = lblInfo(lbl检验结果).Top - 765
    Else
        '预定输血成分
        lblInfo(lbl预定输血成分).Top = lblInfo(lbl备注).Top - 1000
    End If
    picGet(0).Top = lblInfo(lbl预定输血成分).Top - 30
    Line1(15).Y1 = picGet(0).Top + picGet(0).Height - 5
    Line1(15).Y2 = Line1(15).Y1
    picPreBlood.Height = Line1(15).Y1 - picPreBlood.Top
        
    '备血列表
    picPreInfo.Left = 15
    picPreInfo.Top = picPreBlood.Height - picPreInfo.Height
    picPreInfo.Width = picPreBlood.Width - 15
    picPreSum.Left = picPreInfo.Width - picPreSum.Width - 30
    txt申请信息.Width = picPreSum.Left - 240
    
    '血液列表
    vsfBlood.Left = 15
    vsfBlood.Top = 330
    vsfBlood.Width = IIF(bln用血申请 = False, picPreBlood.Width - IIF(gbln显示血液库存 = True, 3000, 45), 5000)
    vsfBlood.Height = picPreInfo.Top - vsfBlood.Top - 30
    
    '库存信息
    vsfList.Left = vsfBlood.Left + vsfBlood.Width + 15
    vsfList.Top = vsfBlood.Top
    vsfList.Width = picPreBlood.Width - vsfList.Left - 45
    vsfList.Height = vsfBlood.Height
    vsfList.Visible = IIF(bln用血申请 = False, gbln显示血液库存, True)
    
    '预定输血量
    lblInfo(lbl预定输血量).Top = lblInfo(lbl预定输血成分).Top
    txtInfo(txt预定输血量).Top = picGet(0).Top
    Line1(17).Y1 = txtInfo(txt预定输血量).Top + txtInfo(txt预定输血量).Height + 15
    Line1(17).Y2 = Line1(17).Y1
    '单位
    txtInfo(txt单位).Top = txtInfo(txt预定输血量).Top
    txtInfo(txt单位).Left = txtInfo(txt预定输血量).Left + txtInfo(txt预定输血量).Width + 120

    '输血途径
    lblInfo(lbl输血途径).Top = lblInfo(lbl预定输血量).Top + lblInfo(lbl预定输血量).Height + 225
    picGet(1).Top = lblInfo(lbl输血途径).Top - 30
    Line1(19).Y1 = picGet(1).Top + picGet(1).Height - 5
    Line1(19).Y2 = Line1(19).Y1
    '输血执行
    lblInfo(lbl输血执行).Top = lblInfo(lbl输血途径).Top
    picInfo(9).Top = picGet(1).Top
    Line1(20).Y1 = Line1(19).Y1
    Line1(20).Y2 = Line1(20).Y1
    
    '计算24小时输血量
    lblInfo(lbl24H输血量).Visible = Not bln用血申请
    If Not bln用血申请 Then lblInfo(lbl24H输血量).Caption = "24小时内输血申请量：" & GetBloodCapacity(IIF(mint场合 = 0, 2, 1), mlng病人ID, IIF(mint场合 = 0, mlng主页ID, mlng挂号ID), zlDatabase.Currentdate, True, CInt(mbytBaby)) & "ML"
    
    If bln用血申请 Then
        picHisItem.Visible = False
    Else
        picHisItem.Visible = True
        lblInfo(lbl本次历史申请项目).Tag = GetPatiHisBloodItem
        If lblInfo(lbl本次历史申请项目).Tag <> "" Then
            lblInfo(lbl本次历史申请项目).Caption = "本次历史申请项目:[" & Replace(lblInfo(lbl本次历史申请项目).Tag, "'", "][") & "]"
            lblInfo(lbl本次历史申请项目).Tag = Replace(lblInfo(lbl本次历史申请项目).Tag, "'", vbCrLf)
        Else
            lblInfo(lbl本次历史申请项目).Caption = ""
        End If
    End If
    
    Call cboInfo_Click(cbo用血安排)
    On Error GoTo 0
    If mblnSelectBlood = True Then
        Call LoadBloodListBySelect
    Else
        Call LoadBloodList(bln用血申请)
    End If
    SetFormNature = True
End Function

Private Sub LoadBloodListBySelect()
'用血申请如果是通过医生选择血液的模式则调用次函数
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim strWhere As String, str单位 As String, strTmp As String
    Dim strSQLChild1 As String, strSQLChild2 As String
    Dim arrItem, arrCode() As String  '这两个数组在开始存放输血申请项目，后面如果使用注意影响
    Dim strID As String
    On Error GoTo ErrHand
    
    txt申请信息.Text = "品种:"
    txt申请量.Text = ""
    
    arrItem = Split(mstr输血项目, ";")
    '获取血液收发ID，血液如果已发，查看时需要提取原始记录
    For i = 0 To UBound(arrItem)
        arrCode = Split(CStr(arrItem(i)), ",")
        If UBound(arrCode) > 3 Then
            strID = strID & IIF(arrCode(4) <> "", "|" & arrCode(4), "")
        End If
        If Left(strID, 1) = "|" Then strID = Mid(strID, 2)
    Next
    With vsfList
        .Clear
        .WordWrap = True
        .ExtendLastCol = True
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 1
        .Cols = 5
        .Editable = flexEDNone
        .TextMatrix(0, COL_S_ID) = "ID": .ColWidth(COL_S_ID) = 0
        .TextMatrix(0, COL_S_选择) = "": .ColWidth(COL_S_选择) = 255
        .TextMatrix(0, COL_S_编号) = "血袋编号": .ColWidth(COL_S_编号) = 1200:
        .TextMatrix(0, COL_S_规格) = "规格": .ColWidth(COL_S_规格) = 1000
        .TextMatrix(0, COL_S_效期) = "效期": .ColWidth(COL_S_效期) = 1400
        .ColHidden(0) = True
        .ColDataType(COL_S_选择) = flexDTString
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_S_规格) = flexAlignCenterCenter
        .GridLines = flexGridFlat
        .Editable = flexEDNone
    End With
    
    With vsfBlood
        .Clear
        .WordWrap = True
        .ExtendLastCol = False
        .FixedRows = 1
        .FixedCols = 0
        .Rows = .FixedRows
        .Cols = 13
        For i = 0 To .Cols - 1
            .ColHidden(i) = False
        Next
        
        .TextMatrix(0, COL_P_ID) = "ID": .ColWidth(COL_P_ID) = 0
        .TextMatrix(0, COL_P_选择) = "": .ColWidth(COL_P_选择) = 255
        .TextMatrix(0, COL_P_编码) = "编码": .ColWidth(COL_P_编码) = 1000:
        .TextMatrix(0, COL_P_名称) = "名称": .ColWidth(COL_P_名称) = 3000
        .TextMatrix(0, COL_P_申请量) = "申请量": .ColWidth(COL_P_申请量) = 800
        .TextMatrix(0, COL_P_单位) = "单位": .ColWidth(COL_P_单位) = 600
        .TextMatrix(0, COL_P_申请血型) = "申请血型": .ColWidth(COL_P_申请血型) = 1000
        .TextMatrix(0, COL_P_申请RH) = "申请RH": .ColWidth(COL_P_申请RH) = 800
        .TextMatrix(0, COL_P_执行分类ID) = "执行分类ID": .ColWidth(COL_P_执行分类ID) = 0
        .TextMatrix(0, COL_P_执行科室ID) = "执行科室ID": .ColWidth(COL_P_执行科室ID) = 0
        .TextMatrix(0, COL_P_录入限量ID) = "录入限量ID": .ColWidth(COL_P_录入限量ID) = 0
        .TextMatrix(0, COL_P_计算系数) = "计算系数": .ColWidth(COL_P_计算系数) = 0
        .TextMatrix(0, COL_P_库存) = "库存": .ColWidth(COL_P_库存) = 0
        
        .ColHidden(COL_P_ID) = True
        .ColHidden(COL_P_执行分类ID) = True
        .ColHidden(COL_P_执行科室ID) = True
        .ColHidden(COL_P_录入限量ID) = True
        .ColHidden(COL_P_计算系数) = True
        .ColHidden(COL_P_库存) = True
        .ColHidden(COL_P_申请量) = True
        .ColHidden(COL_P_申请血型) = True
        .ColHidden(COL_P_申请RH) = True
        .ColDataType(COL_P_选择) = flexDTString
        
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_P_申请量) = flexAlignCenterCenter
        .ColAlignment(COL_P_申请血型) = flexAlignCenterCenter
        .ColAlignment(COL_P_申请RH) = flexAlignCenterCenter
        .Editable = flexEDNone
    End With
    
    If mint场合 = 0 Then
        strWhere = " And b.病人id = [1] And b.主页id = [2] And Nvl(b.婴儿, 0) = [3] "
    Else
        strWhere = " And b.病人id = [1] And b.挂号单 = [2] And Nvl(b.婴儿, 0) = [3]"
    End If
    
    gstrSQL = _
        " Select a.Id, a.编码, a.名称, a.计算单位, a.执行分类 执行分类id, a.执行科室 执行科室id, a.录入限量 录入限量id, a.计算系数, b.血液信息" & vbNewLine & _
        " From 诊疗项目目录 a," & vbNewLine & _
        "     ("
    strSQLChild1 = "Select h.Id," & vbNewLine & _
        "              f_List2str(Cast(Collect(f.Id || '<Split1>' || f.血袋编号 || '<Split1>' || decode(substr('' || Nvl(f.填写数量, 0) * Nvl(换算系数, 1),1,1),'.',0,'') || Nvl(f.填写数量, 0) * Nvl(换算系数, 1) || h.计算单位 || '<Split1>' ||" & vbNewLine & _
        "                                       To_Char(f.效期, 'yyyy-mm-dd hh24:mi')) As t_Strlist)," & vbNewLine & _
        "                          '<Split4>') 血液信息" & vbNewLine & _
        "       From 诊疗项目目录 h, 血液规格 g, 血液收发记录 f, 血液配血记录 e, 病人医嘱记录 b" & vbNewLine & _
        "       Where h.Id = g.品种id And g.规格id = f.血液id  And f.审核人 is Null  And f.配发id = e.Id And Mod(f.记录状态, 3) = 1 And f.发血状态 = 1 And " & vbNewLine & _
        "             e.申请id = b.Id And b.诊疗类别 = 'K' And b.医嘱状态 In (1, 3, 8) " & strWhere & " And" & vbNewLine & _
        "             Exists" & vbNewLine & _
        "        (Select 1" & vbNewLine & _
        "              From 诊疗项目目录 p, 病人医嘱记录 q" & vbNewLine & _
        "              Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = b.Id And q.诊疗类别 = 'E')" & vbNewLine & _
        "       And  Not Exists (Select 1" & vbNewLine & _
        "              From 输血申请项目 a, 病人医嘱记录 b" & vbNewLine & _
        "              Where a.医嘱id = b.Id And Instr('|' || a.血液信息 || '|' ,'|' || f.Id || '|') <> 0  And b.医嘱状态 In (1, 3, 8) And b.诊疗类别 = 'K'  " & strWhere & ")" & vbNewLine & _
        "       Group By h.Id, h.名称, h.计算单位"
        
    strSQLChild2 = "Select h.Id," & vbNewLine & _
        "              f_List2str(Cast(Collect(f.Id || '<Split1>' || f.血袋编号 || '<Split1>' || decode(substr('' || Nvl(f.填写数量, 0) * Nvl(换算系数, 1),1,1),'.',0,'') || Nvl(f.填写数量, 0) * Nvl(换算系数, 1) || h.计算单位 || '<Split1>' ||" & vbNewLine & _
        "                                       To_Char(f.效期, 'yyyy-mm-dd hh24:mi')) As t_Strlist)," & vbNewLine & _
        "                          '<Split4>') 血液信息" & vbNewLine & _
        "       From 诊疗项目目录 h, 血液规格 g, 血液收发记录 f, 血液配血记录 e, 病人医嘱记录 b" & vbNewLine & _
        "       Where h.Id = g.品种id And g.规格id = f.血液id  And instr([4],'|' || f.id || '|',1)<>0  And f.配发id = e.Id And Mod(f.记录状态, 3) = 1 And" & vbNewLine & _
        "             e.申请id = b.Id And b.诊疗类别 = 'K' And b.医嘱状态 In (1, 3, 8) " & strWhere & " And" & vbNewLine & _
        "             Exists" & vbNewLine & _
        "        (Select 1" & vbNewLine & _
        "              From 诊疗项目目录 p, 病人医嘱记录 q" & vbNewLine & _
        "              Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = b.Id And q.诊疗类别 = 'E')" & vbNewLine & _
        "       Group By h.Id, h.名称, h.计算单位"
    If mintType = 0 Then '新增
        gstrSQL = gstrSQL & strSQLChild1
    ElseIf mintType = 2 Then '查看
        gstrSQL = gstrSQL & strSQLChild2
    Else '修改
        gstrSQL = gstrSQL & "Select id, f_List2str(Cast(Collect(血液信息) As t_Strlist),'<Split4>') 血液信息 From (" & strSQLChild2 & vbNewLine & " Union ALL" & vbNewLine & strSQLChild1 & ") Group By Id"
    End If
    
    gstrSQL = gstrSQL & ") b" & vbNewLine & _
            " Where a.Id = b.Id" & vbNewLine & _
            " Order By a.编码"
            
    If mint场合 = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取备血信息", mlng病人ID, mlng主页ID, mbytBaby, "|" & strID & "|")
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取备血信息", mlng病人ID, mstr挂号单, mbytBaby, "|" & strID & "|")
    End If
    
    If rsTmp.RecordCount > 0 Then
        If mintType <> 0 And mintType <> 1 And mintType <> 4 Then '只有新增或修改时才允许选择
            vsfBlood.Editable = flexEDNone
        Else
            vsfBlood.Editable = IIF(mblnUseBloodSend = True, flexEDNone, flexEDKbdMouse)
        End If
    Else
        vsfBlood.Editable = flexEDNone
    End If
    
    vsfList.Editable = vsfBlood.Editable  '血液列表编辑和用血医嘱相同
    
    With vsfBlood
        .Redraw = flexRDNone
        Do While Not rsTmp.EOF
            If .Rows <= .FixedRows Then
                .Rows = .FixedRows + 1
            Else
                If .TextMatrix(.Rows - 1, COL_P_ID) <> "" Then .Rows = .Rows + 1
            End If
            
            .TextMatrix(.Rows - 1, COL_P_ID) = Val(rsTmp!ID & "")
            .TextMatrix(.Rows - 1, COL_P_选择) = ""
            .TextMatrix(.Rows - 1, COL_P_编码) = rsTmp!编码 & ""
            .TextMatrix(.Rows - 1, COL_P_名称) = rsTmp!名称 & ""
            .TextMatrix(.Rows - 1, COL_P_申请量) = ""
            .TextMatrix(.Rows - 1, COL_P_单位) = rsTmp!计算单位 & ""
            If InStr(1, "'" & UCase(str单位) & "'", "'" & UCase(rsTmp!计算单位 & "") & "'") = 0 Then
                str单位 = IIF(str单位 = "", "", str单位 & "'") & rsTmp!计算单位 & ""
            End If
            .TextMatrix(.Rows - 1, COL_P_申请血型) = ""
            .TextMatrix(.Rows - 1, COL_P_申请RH) = ""
            .TextMatrix(.Rows - 1, COL_P_执行分类ID) = Val(rsTmp!执行分类ID & "")
            .TextMatrix(.Rows - 1, COL_P_执行科室ID) = Val(rsTmp!执行科室ID & "")
            .TextMatrix(.Rows - 1, COL_P_录入限量ID) = Val(rsTmp!录入限量ID & "")
            .TextMatrix(.Rows - 1, COL_P_计算系数) = Val(rsTmp!计算系数 & "")
            .TextMatrix(.Rows - 1, COL_P_库存) = rsTmp!血液信息 & ""
            
            Set .Cell(flexcpPicture, .Rows - 1, COL_P_选择) = img16.ListImages("c0").Picture
            .Cell(flexcpData, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_选择) = 0
            .Cell(flexcpFontBold, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = False
            .Cell(flexcpBackColor, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = vbWhite
            For j = 0 To UBound(arrItem)
                arrCode = Split(CStr(arrItem(j)), ",")
                If Val(arrCode(0)) = Val(rsTmp!ID & "") Then
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_选择) = img16.ListImages("c1").Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_选择) = 1
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = True
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = &HC0E0FF
                    .TextMatrix(.Rows - 1, COL_P_申请量) = arrCode(1)
                    .TextMatrix(.Rows - 1, COL_P_申请血型) = arrCode(2)
                    .TextMatrix(.Rows - 1, COL_P_申请RH) = arrCode(3)
                    If UBound(arrCode) > 3 Then
                        .TextMatrix(.Rows - 1, COL_P_库存) = .TextMatrix(.Rows - 1, COL_P_库存) & IIF(arrCode(4) <> "", "<Split3>" & arrCode(4), "")
                    End If
                    'iif(mid("" & 0.5,1,1)=".","0","") & 0.5，这种写法是为了保证小于的1的值能正常显示前缀0
                    txt申请信息.Text = txt申请信息.Text & "[" & .TextMatrix(.Rows - 1, COL_P_名称) & IIF(.TextMatrix(.Rows - 1, COL_P_申请量) <> "", "-" & IIF(Mid("" & .TextMatrix(.Rows - 1, COL_P_申请量), 1, 1) = ".", "0", "") & .TextMatrix(.Rows - 1, COL_P_申请量) & .TextMatrix(.Rows - 1, COL_P_单位), "") & "]"
                End If
            Next
            rsTmp.MoveNext
        Loop
        If .Rows > .FixedRows Then
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
            '确定表格尺寸
            .AutoSize 0, .Cols - 1
            .ColWidth(COL_P_选择) = 255
            .Redraw = flexRDDirect
            Call vsfBlood_AfterRowColChange(0, 0, 1, 1)
        Else
            .Redraw = flexRDDirect
        End If
    End With
    
    '加载单位
    arrItem = Split(str单位, "'")
    cboInfo(cbo单位).Clear
    cboInfo(cbo单位).Tag = ""
    For i = 0 To UBound(arrItem)
        cboInfo(cbo单位).AddItem CStr(arrItem(i))
        If UCase(txtInfo(txt单位).Text) = UCase(CStr(arrItem(i))) Then
            Call zlControl.CboSetIndex(cboInfo(cbo单位).hwnd, i)
            cboInfo(cbo单位).Tag = i
        End If
    Next
    Call BloodSum
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBloodList(ByVal bln用血申请 As Boolean)
'功能：备血申请加载血液信息
    '备血申请相关变量
    Dim str血型 As String, strRH As String
    '用血申请相关变量
    Dim strWhere As String

    '公共变量
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lng诊疗项目ID As Long
    Dim arrItem, arrCode() As String
    Dim str单位 As String
    Dim arrRecord, arrInfo, str备血信息 As String
    Dim str未发 As String, str已发 As String, str总量 As String, str待发 As String
    Dim str配发信息 As String, str替代项目 As String, str计算单位 As String, blnLast As Boolean
    Dim objCollection As New Collection
    Dim bln显示申请量 As Boolean
    
    On Error GoTo ErrHand
    
    If bln用血申请 = False Then
        For i = 0 To cboInfo(cbo输血血型).ListCount - 1
            If InStr(1, ",A,B,O,AB,", "," & cboInfo(cbo输血血型).List(i) & ",") <> 0 Then
                str血型 = str血型 & "|" & cboInfo(cbo输血血型).List(i)
            End If
        Next i
        str血型 = Mid(str血型, 2)
        For i = 0 To cboInfo(cboRHD).ListCount - 1
            strRH = strRH & "|" & cboInfo(cboRHD).List(i)
        Next i
        strRH = Mid(strRH, 2)
    End If
    txt申请信息.Text = "品种:"
    txt申请量.Text = ""
    
    With vsfList
        .Clear
        .WordWrap = True
        .ExtendLastCol = True
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 1
        .Cols = 1
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarSimpleLeaf
        .GridLines = flexGridFlatVert
        .Editable = flexEDNone
        .TextMatrix(0, 0) = IIF(bln用血申请 = True, "配发信息", "库存信息")
        .ColHidden(0) = False
    End With
    
    With vsfBlood
        .Clear
        .WordWrap = True
        .ExtendLastCol = False
        .FixedRows = 1
        .FixedCols = 0
        .Rows = .FixedRows
        .Cols = 13
        For i = 0 To .Cols - 1
            .ColHidden(i) = False
        Next

        .TextMatrix(0, COL_P_ID) = "ID": .ColWidth(COL_P_ID) = 0
        .TextMatrix(0, COL_P_选择) = "": .ColWidth(COL_P_选择) = 255
        .TextMatrix(0, COL_P_编码) = "编码": .ColWidth(COL_P_编码) = 1000:
        .TextMatrix(0, COL_P_名称) = "名称": .ColWidth(COL_P_名称) = 2000
        .TextMatrix(0, COL_P_申请量) = "申请量": .ColWidth(COL_P_申请量) = 800
        .TextMatrix(0, COL_P_单位) = "单位": .ColWidth(COL_P_单位) = 600
        .TextMatrix(0, COL_P_申请血型) = "申请血型": .ColWidth(COL_P_申请血型) = 1000
        .TextMatrix(0, COL_P_申请RH) = "申请RH": .ColWidth(COL_P_申请RH) = 800
        .TextMatrix(0, COL_P_执行分类ID) = "执行分类ID": .ColWidth(COL_P_执行分类ID) = 0
        .TextMatrix(0, COL_P_执行科室ID) = "执行科室ID": .ColWidth(COL_P_执行科室ID) = 0
        .TextMatrix(0, COL_P_录入限量ID) = "录入限量ID": .ColWidth(COL_P_录入限量ID) = 0
        .TextMatrix(0, COL_P_计算系数) = "计算系数": .ColWidth(COL_P_计算系数) = 0
        .TextMatrix(0, COL_P_库存) = "库存": .ColWidth(COL_P_库存) = 0
        
        .ColHidden(COL_P_ID) = True
        .ColHidden(COL_P_执行分类ID) = True
        .ColHidden(COL_P_执行科室ID) = True
        .ColHidden(COL_P_录入限量ID) = True
        .ColHidden(COL_P_计算系数) = True
        .ColHidden(COL_P_库存) = True
        .ColHidden(COL_P_申请血型) = Not (bln用血申请 = False And bln显示申请量 = True)
        .ColHidden(COL_P_申请RH) = Not (bln用血申请 = False And bln显示申请量 = True)
        .ColDataType(COL_P_选择) = flexDTString
        If bln用血申请 = False Then
            .ColComboList(COL_P_申请血型) = str血型
            .ColComboList(COL_P_申请RH) = strRH
            If gbln显示血液库存 = False Then
                .ColWidth(COL_P_名称) = 4000
            ElseIf bln显示申请量 = False Then
                .ColWidth(COL_P_名称) = 4000
            End If
        End If
        
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_P_申请量) = flexAlignCenterCenter
        .ColAlignment(COL_P_申请血型) = flexAlignCenterCenter
        .ColAlignment(COL_P_申请RH) = flexAlignCenterCenter
    End With
    

    If bln用血申请 = False Then
        If gbln显示血液库存 = True Then
            gstrSQL = _
                    " Select Id, 编码,名称, Sum(总量) 总量, f_List2str(Cast(Collect(分类总量) As t_Strlist), '<Split1>') 库存信息, 计算单位,执行科室ID,录入限量ID,执行分类id,计算系数" & vbNewLine & _
                    " From (Select Id, 编码,名称, 计算单位,执行科室ID,录入限量ID,计算系数,执行分类id, Sum(总量) 总量," & vbNewLine & _
                    "              Decode(库房名称, '', '', '【' || 库房名称 || '】- ' || Sum(总量) || 计算单位 || '<Split2>' || f_List2str(Cast(Collect(分类总量 || 计算单位) As t_Strlist),'<Split3>')) 分类总量 " & vbNewLine & _
                    "       From (Select a.Id,A.编码, a.名称, e.库房id, Nvl(Max(f.名称), '') 库房名称," & vbNewLine & _
                    "                     e.Abo || e.Rh || ':' || Nvl(Sum(e.可用数量 * d.换算系数), 0) 分类总量, Nvl(Sum(e.可用数量 * d.换算系数), 0) 总量, a.计算单位,a.计算系数,A.执行科室 as 执行科室ID,A.录入限量 as 录入限量ID,a.执行分类 as 执行分类id" & vbNewLine & _
                    "              From 部门表 f, 血液库存记录 e, 血液规格 d, 诊疗分类目录 c, 诊疗项目目录 a,诊疗项目别名 B" & vbNewLine & _
                    "              Where e.库房id = f.Id(+) And e.血液id(+) = d.规格id And  e.效期(+)>Sysdate And d.品种id = a.Id And c.Id = a.分类id And c.类型 = 8 And A.ID=B.诊疗项目ID" & vbNewLine & _
                    "                   And A.类别='K'  And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & vbNewLine & _
                    "                   And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                    "                           And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[1])  Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & vbNewLine & _
                    "                   " & Decode(gbytCode, 0, " And B.码类 IN([2],3)", 1, " And B.码类 IN([2],3)", "") & vbNewLine & _
                    "              Group By a.Id, A.编码,a.名称, a.计算单位,A.执行科室,A.录入限量, e.库房id, e.Abo, e.Rh,执行分类,计算系数)" & vbNewLine & _
                    "       Group By Id, 编码,名称, 计算单位,执行科室ID,录入限量ID, 库房名称,执行分类id,计算系数)" & vbNewLine & _
                    " Group By Id, 编码,名称, 计算单位,执行科室ID,录入限量ID,执行分类id,计算系数" & vbNewLine & _
                    " Order by 编码"
        Else
            gstrSQL = "Select Distinct a.Id, a.编码, a.名称, a.执行分类 As 执行分类id, a.计算单位, a.执行科室 As 执行科室id, a.录入限量 As 录入限量id, a.计算系数,' ' as 库存信息,' ' as 总量" & vbNewLine & _
                " From 诊疗分类目录 c, 诊疗项目目录 a, 诊疗项目别名 b, 血液品种 d" & vbNewLine & _
                " Where c.Id = a.分类id And c.类型 = 8 And a.Id = b.诊疗项目id And a.Id = d.品种id And a.类别 = 'K' And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And" & vbNewLine & _
                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null) And" & vbNewLine & _
                "      (Exists (Select 1 From 诊疗适用科室 Where 项目id = a.Id And 科室id = [1]) Or Not Exists" & vbNewLine & _
                "       (Select 1 From 诊疗适用科室 Where 项目id = a.Id)) " & Decode(gbytCode, 0, " And B.码类 IN([2],3)", 1, " And B.码类 IN([2],3)", "") & vbNewLine & _
                " Order By a.编码"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取血液信息", mlng病人科室id, gbytCode + 1)
         If rsTmp.RecordCount > 0 Then
            If mintType <> 0 And mintType <> 1 Then  '只有新增或修改时才允许选择
                vsfBlood.Editable = flexEDNone
            Else
                vsfBlood.Editable = flexEDKbdMouse
            End If
        Else
            vsfBlood.Editable = flexEDNone
        End If
        arrRecord = Array()
        Do While Not rsTmp.EOF
            '一个品种一条信息
            ReDim Preserve arrRecord(UBound(arrRecord) + 1)
            arrRecord(UBound(arrRecord)) = rsTmp!ID & "'" & rsTmp!编码 & "'" & rsTmp!名称 & "'" & rsTmp!计算单位 & "'" & rsTmp!执行分类ID & "'" & rsTmp!执行科室ID & "'" & rsTmp!录入限量ID & "'" & rsTmp!计算系数
            '库存格式SQL已经处理好。格式：【血液室】- 800ml<Split2>A型+:400ml<Split3>B型+:400ml<Split1>【LW医技中心A】- 800ml<Split2>A型+:400ml<Split3>B型+:400ml
            If gbln显示血液库存 = True Then
                str备血信息 = "" & rsTmp!库存信息
            Else
                str备血信息 = ""
            End If
            
            objCollection.Add str备血信息, "A_" & rsTmp!ID
            rsTmp.MoveNext
        Loop
    Else
        If mint场合 = 0 Then
            strWhere = " And b.病人id = [1] And b.主页id = [2] And Nvl(b.婴儿, 0) = [3] "
        Else
            strWhere = " And b.病人id = [1] And b.挂号单 = [2] And Nvl(b.婴儿, 0) = [3]"
        End If
        '123316,调整血液信息，改为逐条提取在拼接的方式（以前是sql中直接处理好的，但是由于用户病人血液太多拼接字符超出4000报错）
        gstrSQL = "Select a.Id, a.编码, a.名称, a.计算单位, a.执行分类id, a.执行科室id, a.录入限量id, a.计算系数, a.总量, a.已发, a.未发, a.血液信息, a.是否配血, a.待发, a.待发单位," & vbNewLine & _
                        "       a.替代项目" & vbNewLine & _
                        "From (With 配血记录 As (Select h.Id," & vbNewLine & _
                        "                           Decode(Nvl(f.审核人, ''), '', Nvl(f.填写数量, 0) * Nvl(换算系数, 1), 0) *" & vbNewLine & _
                        "                            Decode(Upper(h.计算单位), 'ML', 1, Nvl(h.计算系数, 1)) 待发, Nvl(f.填写数量, 0) * Nvl(换算系数, 1) 总量," & vbNewLine & _
                        "                           Decode(Nvl(f.审核人, ''), '', 0, Nvl(f.填写数量, 0) * Nvl(换算系数, 1)) 已发," & vbNewLine & _
                        "                           Decode(Nvl(f.审核人, ''), '', Nvl(f.填写数量, 0) * Nvl(换算系数, 1), 0) 未发," & vbNewLine & _
                        "                           ' 规格：' || Decode(Substr('' || Nvl(f.填写数量, 0) * Nvl(换算系数, 1), 1, 1), '.', 0, '') ||" & vbNewLine & _
                        "                            Nvl(f.填写数量, 0) * Nvl(换算系数, 1) || h.计算单位 || Decode(Nvl(f.审核人, ''), '', '(未发)', '(已发)') ||" & vbNewLine & _
                        "                            '  效期:' || To_Char(f.效期, 'yyyy-mm-dd hh24:mi') || '<Split3>' ||" & vbNewLine & _
                        "                            Decode(Nvl(f.审核人, ''), '', 0, 1) 血液信息, Decode(Nvl(f.审核人, ''), '', 0, 1) 是否已发, f.效期" & vbNewLine & _
                        "                    From 诊疗项目目录 h, 血液规格 g, 血液收发记录 f, 血液配血记录 e, 病人医嘱记录 b" & vbNewLine & _
                        "                    Where h.Id = g.品种id And g.规格id = f.血液id And f.配发id = e.Id And Mod(f.记录状态, 3) = 1 And" & vbNewLine & _
                        "                          Instr(',0,3,', ',' || f.发血状态 || ',') = 0 And e.申请id = b.Id And b.诊疗类别 = 'K' And" & vbNewLine & _
                        "                          b.医嘱状态 In (1, 3, 8) " & strWhere & " And Exists" & vbNewLine & _
                        "                     (Select 1" & vbNewLine & _
                        "                           From 诊疗项目目录 p, 病人医嘱记录 q" & vbNewLine & _
                        "                           Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = b.Id And" & vbNewLine & _
                        "                                 q.诊疗类别 = 'E'))"

        gstrSQL = gstrSQL & vbNewLine & _
                        "       Select a.Id, a.编码, a.名称, a.计算单位, a.执行分类 执行分类id, a.执行科室 执行科室id, a.录入限量 录入限量id, a.计算系数," & vbNewLine & _
                        "               (Select f_List2str(Cast(Collect('' || 替代项目id) As t_Strlist)) From 诊疗项目替代 Where 项目id = a.Id) 替代项目, b.总量," & vbNewLine & _
                        "               b.已发, b.未发, 是否已发, b.血液信息, b.效期, Decode(b.Id, Null, 0, 1) 是否配血, b.待发, 'ml' 待发单位" & vbNewLine & _
                        "       From 诊疗项目目录 a, 配血记录 b," & vbNewLine & _
                        "           (Select Id 诊疗项目id" & vbNewLine & _
                        "               From 配血记录" & vbNewLine & _
                        "               Union" & vbNewLine & _
                        "               Select Decode(Nvl(c.医嘱id, 0), 0, b.诊疗项目id, c.诊疗项目id) 诊疗项目id" & vbNewLine & _
                        "               From 输血申请项目 c, 病人医嘱记录 b" & vbNewLine & _
                        "               Where c.医嘱id(+) = b.Id " & strWhere & " And b.诊疗类别 = 'K' And" & vbNewLine & _
                        "                    b.医嘱状态 In (1, 3, 8) And Exists" & vbNewLine & _
                        "               (Select 1" & vbNewLine & _
                        "                    From 诊疗项目目录 p, 病人医嘱记录 q" & vbNewLine & _
                        "                   Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = b.Id And q.诊疗类别 = 'E')) c" & vbNewLine & _
                        "       Where a.Id = c.诊疗项目id And c.诊疗项目id = b.Id(+)) a" & vbNewLine & _
                        "       Order By a.编码, a.是否已发, a.效期"
            
        If mint场合 = 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取备血信息", mlng病人ID, mlng主页ID, mbytBaby)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取备血信息", mlng病人ID, mstr挂号单, mbytBaby)
        End If
        
        If rsTmp.RecordCount > 0 Then
            If mintType <> 0 And mintType <> 1 And mintType <> 4 Then '只有新增或修改时才允许选择
                vsfBlood.Editable = flexEDNone
            Else
                vsfBlood.Editable = IIF(mblnUseBloodSend = True, flexEDNone, flexEDKbdMouse)
            End If
        Else
            vsfBlood.Editable = flexEDNone
        End If
        '开始进行内容组装
        blnLast = False
        arrRecord = Array()
        str配发信息 = "": str未发 = "": str已发 = "": str总量 = "": str待发 = ""
        lng诊疗项目ID = -999: str替代项目 = "": str计算单位 = ""
        Do While Not rsTmp.EOF
            If lng诊疗项目ID <> Val("" & rsTmp!ID) Then
                If lng诊疗项目ID <> -999 Then
GOWORK:
                    If str配发信息 = "" Then
                        str配发信息 = "还未配血"
                    Else
                        str配发信息 = "配血总量：" & IIF(Left(str总量, 1) = ".", "0", "") & str总量 & str计算单位 & " 已发：" & IIF(Left(str已发, 1) = ".", "0", "") & str已发 & str计算单位 & " 未发：" & IIF(Left(str未发, 1) = ".", "0", "") & str未发 & str计算单位 & str配发信息
                    End If
                    '配发信息格式：配血总量：400ml 已发量：0ml 未发量：400ml<Split4> 规格：200ml(未发)  效期:2016-09-17 16:13<Split3>0<Split4> 规格：200ml(未发)  效期:2016-08-14 11:17<Split3>0
                    str备血信息 = str替代项目 & "<Split2>" & lng诊疗项目ID & "'" & str配发信息 & "'" & str待发
                    objCollection.Add str备血信息, "A_" & lng诊疗项目ID
                    If blnLast = True Then GoTo GONEXT
                End If
                str配发信息 = "": str未发 = "": str已发 = "": str总量 = "": str待发 = ""
                lng诊疗项目ID = Val("" & rsTmp!ID)
                str替代项目 = "" & rsTmp!替代项目
                str计算单位 = "" & rsTmp!计算单位
                ReDim Preserve arrRecord(UBound(arrRecord) + 1)
                arrRecord(UBound(arrRecord)) = rsTmp!ID & "'" & rsTmp!编码 & "'" & rsTmp!名称 & "'" & rsTmp!计算单位 & "'" & rsTmp!执行分类ID & "'" & rsTmp!执行科室ID & "'" & rsTmp!录入限量ID & "'" & rsTmp!计算系数
            End If
            
            If Val(rsTmp!是否配血 & "") = 1 Then
                str总量 = Val(str总量) + Val(rsTmp!总量 & "")
                str已发 = Val(str已发) + Val(rsTmp!已发 & "")
                str未发 = Val(str未发) + Val(rsTmp!未发 & "")
                str待发 = Val(str待发) + Val(rsTmp!待发 & "")
                If rsTmp!血液信息 & "" <> "" Then str配发信息 = str配发信息 & "<Split4>" & rsTmp!血液信息
            End If
            rsTmp.MoveNext
        Loop
        If lng诊疗项目ID <> -999 Then
            blnLast = True
            GoTo GOWORK
        End If
GONEXT:
    End If
    
    arrItem = Split(mstr输血项目, ";")
    With vsfBlood
        .Redraw = flexRDNone
        For i = 0 To UBound(arrRecord)
            arrInfo = Split(CStr(arrRecord(i)), "'")
            If .Rows <= .FixedRows Then
                .Rows = .FixedRows + 1
            Else
                If .TextMatrix(.Rows - 1, COL_P_ID) <> "" Then
                    .Rows = .Rows + 1
                End If
            End If
            
            .TextMatrix(.Rows - 1, COL_P_ID) = Val(arrInfo(0))
            .TextMatrix(.Rows - 1, COL_P_选择) = ""
            .TextMatrix(.Rows - 1, COL_P_编码) = CStr(arrInfo(1))
            .TextMatrix(.Rows - 1, COL_P_名称) = CStr(arrInfo(2))
            .TextMatrix(.Rows - 1, COL_P_申请量) = ""
            .TextMatrix(.Rows - 1, COL_P_单位) = CStr(arrInfo(3))
            If InStr(1, "'" & UCase(str单位) & "'", "'" & UCase(CStr(arrInfo(3))) & "'") = 0 Then
                str单位 = IIF(str单位 = "", "", str单位 & "'") & CStr(arrInfo(3))
            End If
            .TextMatrix(.Rows - 1, COL_P_申请血型) = ""
            .TextMatrix(.Rows - 1, COL_P_申请RH) = ""
            .TextMatrix(.Rows - 1, COL_P_执行分类ID) = Val(arrInfo(4))
            .TextMatrix(.Rows - 1, COL_P_执行科室ID) = Val(arrInfo(5))
            .TextMatrix(.Rows - 1, COL_P_录入限量ID) = Val(arrInfo(6))
            .TextMatrix(.Rows - 1, COL_P_计算系数) = Val(arrInfo(7))
            .TextMatrix(.Rows - 1, COL_P_库存) = CStr(ISExistCollection(objCollection, "A_" & Val(arrInfo(0))))
            
            Set .Cell(flexcpPicture, .Rows - 1, COL_P_选择) = img16.ListImages("c0").Picture
            .Cell(flexcpData, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_选择) = 0
            .Cell(flexcpFontBold, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = False
            .Cell(flexcpBackColor, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = vbWhite
            For j = 0 To UBound(arrItem)
                arrCode = Split(CStr(arrItem(j)), ",")
                If Val(arrCode(0)) = Val(arrInfo(0)) Then
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_选择) = img16.ListImages("c1").Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_选择) = 1
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = True
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_选择, .Rows - 1, COL_P_库存) = &HC0E0FF
                    .TextMatrix(.Rows - 1, COL_P_申请量) = arrCode(1)
                    .TextMatrix(.Rows - 1, COL_P_申请血型) = arrCode(2)
                    .TextMatrix(.Rows - 1, COL_P_申请RH) = arrCode(3)
                    
                    txt申请信息.Text = txt申请信息.Text & "[" & .TextMatrix(.Rows - 1, COL_P_名称) & IIF(.TextMatrix(.Rows - 1, COL_P_申请量) <> "", "-" & .TextMatrix(.Rows - 1, COL_P_申请量) & .TextMatrix(.Rows - 1, COL_P_单位), "") & "]"
                End If
            Next
        Next
        If .Rows > .FixedRows Then
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
            .CellBorderRange .FixedRows, COL_P_申请量, .Rows - 1, COL_P_申请量, vbGreen, 1, 1, 1, 1, 1, 1
            If bln用血申请 = False Then
                .CellBorderRange .FixedRows, COL_P_申请血型, .Rows - 1, COL_P_申请血型, vbGreen, 1, 1, 1, 1, 1, 1
                .CellBorderRange .FixedRows, COL_P_申请RH, .Rows - 1, COL_P_申请RH, vbGreen, 1, 1, 1, 1, 1, 1
            End If
            '确定表格尺寸
            .AutoSize 0, .Cols - 1
            .ColWidth(COL_P_选择) = 255
            .Redraw = flexRDDirect
            Call vsfBlood_AfterRowColChange(0, 0, 1, 1)
        Else
            .Redraw = flexRDDirect
        End If
    End With
    
    '加载单位
    arrItem = Split(str单位, "'")
    cboInfo(cbo单位).Clear
    cboInfo(cbo单位).Tag = ""
    For i = 0 To UBound(arrItem)
        cboInfo(cbo单位).AddItem CStr(arrItem(i))
        If UCase(txtInfo(txt单位).Text) = UCase(CStr(arrItem(i))) Then
            Call zlControl.CboSetIndex(cboInfo(cbo单位).hwnd, i)
            cboInfo(cbo单位).Tag = i
        End If
    Next
    Call BloodSum
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ISExistCollection(ByVal objColl As Collection, ByVal strKey As String) As Variant
'功能:判断key值是否存在集合中，如果存在则返回对应内容,否则返回空
    Dim strReturn
    On Error Resume Next
    err.Clear
    strReturn = objColl(strKey)
    If err <> 0 Then
        strReturn = ""
        err.Clear
    End If
    On Error GoTo 0
    ISExistCollection = strReturn
End Function

Private Sub LoadLastPrepareBlood(Optional ByVal lngActiveID As Long = 0)
'功能：'下达用血申请时：诊断、输血目的、血型等默认去最后一次备血申请的信息
    Dim strWhere As String, strSQL As String, str诊断 As String
    Dim rsTmp As New ADODB.Recordset, rsTmpOther As New ADODB.Recordset
    Dim lng医嘱ID As Long, int紧急标志 As Integer, str用药理由 As String
    
    If mblnSpareBloood = True Then Exit Sub
    
    If mint场合 = 0 Then
        strWhere = " And A.病人ID=[1] And A.主页ID=[2]"
    Else
        strWhere = " And A.病人id =[1] And A.挂号单=[2] "
    End If
    
    On Error GoTo ErrHand
    If lngActiveID = 0 Then
        '获取最后一次备血申请
        strSQL = _
            " Select a.id,a.紧急标志,a.用药理由" & vbNewLine & _
            " From 诊疗项目目录 p, 病人医嘱记录 q, 病人医嘱记录 a" & vbNewLine & _
            " Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = a.Id And q.诊疗类别 = 'E' And a.诊疗类别 = 'K' And" & vbNewLine & _
            "      a.医嘱状态 In (1, 3, 8) " & strWhere & vbNewLine & _
            " Order By a.开始执行时间 Desc"
        If mint场合 = 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取最后一次备血申请", mlng病人ID, mlng主页ID)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取最后一次备血申请", mlng病人ID, mstr挂号单)
        End If
        If rsTmp.EOF Then Exit Sub
        
        mblnDataLoad = True
        lng医嘱ID = Val("" & rsTmp!ID)
        int紧急标志 = Val("" & rsTmp!紧急标志)
        str用药理由 = "" & rsTmp!用药理由
        If int紧急标志 = 1 Then
            cboInfo(cbo用血安排).ListIndex = 1
        End If
        Call zlControl.CboSetText(cboInfo(cbo输血目的), str用药理由, True, "'")
    Else
        lng医嘱ID = lngActiveID
    End If
    
    strSQL = "Select 是否待诊,输血类型, 输血目的, 输血性质, 输血血型, Rhd" & vbNewLine & _
                " From 输血申请记录" & vbNewLine & _
                " Where 医嘱id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        If Val(rsTmp!是否待诊 & "") = 1 Then
            txtInfo(txt诊断信息).Text = "待诊"
            chkWait.value = 1
        Else
            '读取诊断
            mstr诊断IDs = GetAdviceDiag(lng医嘱ID, str诊断)
            txtInfo(txt诊断信息).Text = str诊断
            '从附项中获取诊断如果附项中有以附项为准
             strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
             Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
             If Not rsTmpOther.EOF Then
                 txtInfo(txt诊断信息).Text = rsTmpOther!内容 & ""
             End If
        End If
        txtInfo(txt诊断信息).Tag = txtInfo(txt诊断信息).Text
        chkWait.value = Val(rsTmp!是否待诊 & "")
        Call zlControl.CboSetText(cboInfo(cbo输血类型), rsTmp!输血类型 & "", True, "'")
        If "" & rsTmp!输血目的 <> "" Then Call zlControl.CboSetText(cboInfo(cbo输血目的), rsTmp!输血目的 & "", True, "'") '老的的输血目的存储在医嘱的用药理由里面
        cboInfo(cbo输血性质).ListIndex = Val(rsTmp!输血性质 & "")
        cboInfo(cbo输血血型).ListIndex = Val(rsTmp!输血血型 & "")
        cboInfo(cboRHD).ListIndex = Val(rsTmp!RHD & "")
    End If
    '血型备血申请单可能没有，从病人信息从表中获取
    If cboInfo(cbo输血血型).ListIndex <= 0 Then
        strSQL = "Select 信息值 from 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
        Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, "ABO") '这里信息名是ABO,而不是'血型',是无法读取血型的原因
        If Not rsTmpOther.EOF Then
            Select Case "" & rsTmpOther!信息值
                Case "A", "A型"
                    cboInfo(cbo输血血型).ListIndex = 1
                Case "B", "B型"
                    cboInfo(cbo输血血型).ListIndex = 2
                Case "O", "O型"
                    cboInfo(cbo输血血型).ListIndex = 3
                Case "AB", "AB型"
                    cboInfo(cbo输血血型).ListIndex = 4
                Case "不详"
                    cboInfo(cbo输血血型).ListIndex = 5
                Case "未查"
                    cboInfo(cbo输血血型).ListIndex = 6
            End Select
        End If
    End If
    If cboInfo(cboRHD).ListIndex <= 0 Then
        strSQL = "Select 信息值 from 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
        Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, "RH")
        If Not rsTmpOther.EOF Then
            Select Case "" & rsTmpOther!信息值
                Case "-", "阴"
                    cboInfo(cboRHD).ListIndex = 1
                Case "+", "阳"
                    cboInfo(cboRHD).ListIndex = 2
            End Select
        End If
    End If
    
    mblnDataLoad = False
    Exit Sub
ErrHand:
    mblnDataLoad = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetCommandBarPara(ByVal lngControlID As Long, ByVal lngIndex As Long, ByVal strValue As String)
'功能：设置申请下拉菜单控件相应的属性值
    Dim cbsControl As CommandBarControl
    Dim cbsControlPopup As CommandBarControl
    On Error Resume Next
    Set cbsControlPopup = cbsMain(2).Controls.Find(, lngControlID)
    If Not cbsControlPopup Is Nothing Then
        Set cbsControl = cbsControlPopup.CommandBar.Controls(lngIndex)
        If Not cbsControl Is Nothing Then
            cbsControl.Category = strValue
        End If
    End If
    If err < 0 Then err.Clear
End Sub

Private Function BloodApplyCheck() As Boolean
'功能：新开和修改输血申请时，保存数据之前对申请的相关内容进行检查，并返回提示及处理结果。
'问题号：116846:刘鹏飞,2017-11-23
    Dim strResult As String
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim var1 As Variant
    Dim var2 As Variant
    Dim strSQL As String, strMsg As String
    Dim rsTmp As New ADODB.Recordset
    Dim str申请项目 As String
    Dim lng输血执行科室ID As Long, lng输血途径执行科室ID As Long
    
    On Error GoTo ErrHand
    var1 = Array()
    var2 = Array()
    '输血申请才有检验项目
    If mblnSpareBloood = True Then
        '检验项目
        With vsLIS
            For i = 0 To .Rows - 1
                For j = 0 To CON_LisResultCol - 1
                    If Val(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) <> 0 Then
                        var1 = Array(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标英文名 + (j * CON_LisResultCount)), .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_结果单位 + (j * CON_LisResultCount)), .TextMatrix(i, COL_结果标志 + (j * CON_LisResultCount)), .TextMatrix(i, COL_结果参考 + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_取值序列 + (j * CON_LisResultCount)), IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0))
                        strTmp = Join(var1, "<SplitCol>")
                        ReDim Preserve var2(UBound(var2) + 1)
                        var2(UBound(var2)) = strTmp
                    End If
                Next
            Next
        End With
        strResult = Join(var2, "<SplitRow>")
    End If
    str申请项目 = GetBloodInfo(False)
    lng输血执行科室ID = IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, 0, cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex))
    lng输血途径执行科室ID = IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, 0, cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex))
    
    If mblnSpareBloood = True Then '备血申请单
        strSQL = "Select Zl1_EX_BloodApplyCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]) as 结果 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_EX_BloodApplyCheck", IIF(1 = mint场合, 1, 2), mlng病人ID, IIF(mint场合 = 1, mlng挂号ID, mlng主页ID), IIF(mblnSpareBloood = True, 1, 2), _
            cboInfo(cbo用血安排).ListIndex, chkWait.value, IIF(chkWait.value = 0, txtInfo(txt诊断信息).Text, ""), mstr诊断IDs, cboInfo(cbo输血类型).Text, cboInfo(cbo输血目的).Text, cboInfo(cbo输血性质).Text, txtInfo(txt预定输血时间).Text, _
            cboInfo(cbo输血血型).Text, cboInfo(cboRHD).Text, str申请项目, lng输血执行科室ID, mlng输血途径, lng输血途径执行科室ID, txtInfo(txt备注).Text, mbytBaby, IIF(optHistory(0).value, 0, 1), _
            IIF(optHistory(2).value, 0, 1), IIF(optHistory(4).value, 0, 1), txtInfo(txt孕) & "/" & txtInfo(txt产), IIF(optPossession(0).value, 0, 1), strResult)
    Else '用血申请单
        strSQL = "Select Zl1_EX_BloodApplyCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20]) as 结果 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_EX_BloodApplyCheck", IIF(1 = mint场合, 1, 2), mlng病人ID, IIF(mint场合 = 1, mlng挂号ID, mlng主页ID), IIF(mblnSpareBloood = True, 1, 2), _
            cboInfo(cbo用血安排).ListIndex, chkWait.value, IIF(chkWait.value = 0, txtInfo(txt诊断信息).Text, ""), mstr诊断IDs, cboInfo(cbo输血类型).Text, cboInfo(cbo输血目的).Text, cboInfo(cbo输血性质).Text, txtInfo(txt预定输血时间).Text, _
            cboInfo(cbo输血血型).Text, cboInfo(cboRHD).Text, str申请项目, lng输血执行科室ID, mlng输血途径, lng输血途径执行科室ID, txtInfo(txt备注).Text, mbytBaby)
    End If
    
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!结果)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '提示
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '禁止
                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                strMsg = "": Exit Function
            End Select
            strMsg = ""
        End If
    End If
                
    BloodApplyCheck = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckOrResetLisAboRH(Optional ByVal blnReset As Boolean = False, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
'功能：备血申请时：1、如果有LIS检验结果，自动设置LIS检验结果信息。2：保存时检查ABO和RH是否填写的一致，不一致则禁止
'参数：blnReset：TRUE 根据检验结果设置ABO和RH，False，保存是检查是否和检验结果一致
'         lngRow和lngCol 在编辑检验结果时传入调用，便于同步更新ABO和RH（blnReset=true时有效）,lngCOl当前编辑的指标结果列

    Dim lngCount As Long
    Dim i As Integer, j As Integer
    Dim strAboCode As String, strRHCode As String
    Dim intAboType As Integer, intRHType As Integer
    Dim ArrResult(0 To 1) As String, arrCode() As String
    Dim blnIsAbo As Boolean, blnIsRH As Boolean  '检查结果是否返回了ABO和RH
    Dim strTemp As String, blnMsg As Boolean
    Dim strResult As String

    
    If mblnSpareBloood = False Then CheckOrResetLisAboRH = True: Exit Function '备血申请才有检验结果
    
    '列合法性检查
    If lngRow <> -1 Then
        If Not (lngRow >= vsLIS.FixedRows And lngRow < vsLIS.Rows) Then Exit Function
    End If
    If lngCol <> -1 Then
        '列只能是指标结果列
        If Not (lngCol Mod 10 = COL_指标结果) Then Exit Function
    End If
    
    On Error GoTo ErrHand
    If mstrLISAboRHCode = "" Then
        strTemp = GetBloodApplyCode(1)
        mstrLISAboRHCode = strTemp
    Else
        strTemp = mstrLISAboRHCode
    End If
    If InStr(1, strTemp, ",") > 0 Then
        strAboCode = Mid(strTemp, 1, InStr(1, strTemp, ",") - 1)
        strRHCode = Mid(strTemp, InStr(1, strTemp, ",") + 1)
    Else
        strAboCode = strTemp
        strRHCode = ""
    End If
    If InStr(1, strAboCode, ":") > 0 Then
        intAboType = Val(Mid(strAboCode, InStr(1, strAboCode, ":") + 1))
        strAboCode = Mid(strAboCode, 1, InStr(1, strAboCode, ":") - 1)
    End If
    If InStr(1, strRHCode, ":") > 0 Then
        intRHType = Val(Mid(strRHCode, InStr(1, strRHCode, ":") + 1))
        strRHCode = Mid(strRHCode, 1, InStr(1, strRHCode, ":") - 1)
    End If
    
    If blnReset = True And lngRow <> -1 And lngCol <> -1 Then
        With vsLIS
            If Val(.TextMatrix(lngRow, lngCol + (COL_检验项目ID - COL_指标结果))) <> 0 Then
                If strAboCode = .TextMatrix(lngRow, lngCol + (COL_指标代码 - COL_指标结果)) And strAboCode <> "" Then
                    ArrResult(0) = .TextMatrix(lngRow, lngCol + (COL_指标中文名 - COL_指标结果)) & "'" & .TextMatrix(lngRow, lngCol)
                    blnIsAbo = IIF(.Cell(flexcpBackColor, lngRow, lngCol) = COLEditBackColor, False, True)
                End If
                
                If strRHCode = .TextMatrix(lngRow, lngCol + (COL_指标代码 - COL_指标结果)) And strRHCode <> "" Then
                    ArrResult(1) = .TextMatrix(lngRow, lngCol + (COL_指标中文名 - COL_指标结果)) & "'" & .TextMatrix(lngRow, lngCol)
                    blnIsRH = IIF(.Cell(flexcpBackColor, lngRow, lngCol) = COLEditBackColor, False, True)
                    .Cell(flexcpForeColor, lngRow, lngCol) = IIF(.TextMatrix(lngRow, lngCol) = "-", vbRed, &H80000012)
                End If
            End If
        End With
    Else
        '获取ABO、RH在检验结果中对应的的指标名称和指标结果
        With vsLIS
            lngCount = 0
            For i = 0 To .Rows - 1
                For j = 0 To CON_LisResultCol - 1
                    If Val(.TextMatrix(i, COL_检验项目ID + (j * CON_LisResultCount))) <> 0 Then
                        If strAboCode = .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)) And strAboCode <> "" Then
                            ArrResult(0) = .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)) & "'" & .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount))
                            blnIsAbo = IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, False, True)
                        End If
                        If strRHCode = .TextMatrix(i, COL_指标代码 + (j * CON_LisResultCount)) And strRHCode <> "" Then
                            ArrResult(1) = .TextMatrix(i, COL_指标中文名 + (j * CON_LisResultCount)) & "'" & .TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount))
                            blnIsRH = IIF(.Cell(flexcpBackColor, i, COL_指标结果 + (j * CON_LisResultCount)) = COLEditBackColor, False, True)
                            .Cell(flexcpForeColor, i, COL_指标结果 + (j * CON_LisResultCount)) = IIF(.TextMatrix(i, COL_指标结果 + (j * CON_LisResultCount)) = "-", vbRed, &H80000012)
                        End If
                    End If
                Next
            Next
        End With
    End If
    If blnReset = True Then
        '116848
        '根据检验结果设置ABO和RH
        If ArrResult(0) <> "" Then
            arrCode = Split(ArrResult(0), "'")
            strTemp = UCase(arrCode(1))
            If strTemp = "A" Or strTemp = "A型" Then
                cboInfo(cbo输血血型).ListIndex = 1
            ElseIf strTemp = "B" Or strTemp = "B型" Then
                cboInfo(cbo输血血型).ListIndex = 2
            ElseIf strTemp = "AB" Or strTemp = "AB型" Then
                cboInfo(cbo输血血型).ListIndex = 4
            ElseIf strTemp = "O" Or strTemp = "O型" Then
                cboInfo(cbo输血血型).ListIndex = 3
            ElseIf strTemp = "不详" Then
                cboInfo(cbo输血血型).ListIndex = 5
            ElseIf strTemp = "未查" Then
                cboInfo(cbo输血血型).ListIndex = 6
            Else
                If strTemp = "" Then
                    If blnIsAbo = True Then '不详(表明LIS未出结果)
                        cboInfo(cbo输血血型).ListIndex = 5
                    Else '未查 (表明未做血型检查)
                        cboInfo(cbo输血血型).ListIndex = 6
                    End If
                Else
                    cboInfo(cbo输血血型).ListIndex = 0
                End If
            End If
        End If
        
        If ArrResult(1) <> "" Then
            arrCode = Split(ArrResult(1), "'")
            strTemp = arrCode(1)
            If strTemp = "-" Or strTemp Like "阴性*" Then
                cboInfo(cboRHD).ListIndex = 1
            ElseIf strTemp = "+" Or strTemp Like "阳性*" Then
                cboInfo(cboRHD).ListIndex = 2
            Else
                cboInfo(cboRHD).ListIndex = 0
            End If
        End If
    Else
        '结果一致性检查
        For i = 0 To 1
            If ArrResult(i) <> "" Then
                arrCode = Split(ArrResult(i), "'")
                If i = 0 Then
                    strTemp = UCase(cboInfo(cbo输血血型).Text)
                    strResult = UCase(arrCode(1))
                Else
                    strTemp = cboInfo(cboRHD).Text
                    If arrCode(1) Like "阳性*" Then
                        strResult = "+"
                    ElseIf arrCode(1) Like "阴性*" Then
                        strResult = "-"
                    Else
                        strResult = arrCode(1)
                    End If
                End If
                If strResult <> "" And strTemp <> strResult Then
                    blnMsg = True
                    If i = 0 Then
                        If strResult Like "*型" And Trim(strTemp) <> "" And strTemp & "型" = strResult Then
                            blnMsg = False
                        End If
                    End If
                    If blnMsg = True Then
                        If i = 0 Then
                            If intAboType = 0 Then
                                If MsgBox("申请单中的血型和检验结果中指标[" & arrCode(0) & "]的结果不符，请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Function
                                End If
                            Else
                                MsgBox "申请单中的血型和检验结果中指标[" & arrCode(0) & "]的结果不符，请检查！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        Else
                            If intRHType = 0 Then
                                If MsgBox("申请单中的RHD和检验结果中指标[" & arrCode(0) & "]的结果不符，请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Function
                                End If
                            Else
                                MsgBox "申请单中的RHD和检验结果中指标[" & arrCode(0) & "]的结果不符，请检查！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    CheckOrResetLisAboRH = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetBloodInfo(Optional ByVal blnABO As Boolean = True) As String
'申请返回选择的血液信息：格式：诊疗项目ID,申请量,申请血型,申请RH;诊疗项目ID,申请量,申请血型,申请RH
    Dim lngRow As Long
    Dim strRow As String, strTmp As String
    Dim strIDs As String

    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                '医生选择血袋信息，则需要记录选择的收发ID
                strIDs = ""
                If mblnSelectBlood Then
                    strTmp = .TextMatrix(lngRow, COL_P_库存)
                    If InStr(1, strTmp, "<Split3>") <> 0 Then
                        strIDs = Split(strTmp, "<Split3>")(1) '获取到选中的血液信息
                    End If
                End If
                If blnABO = True Then
                    strRow = strRow & ";" & .TextMatrix(lngRow, COL_P_ID) & "," & .TextMatrix(lngRow, COL_P_申请量) & "," & .TextMatrix(lngRow, COL_P_申请血型) & "," & .TextMatrix(lngRow, COL_P_申请RH) & IIF(strIDs <> "", "," & strIDs, "")
                Else
                    strRow = strRow & ";" & .TextMatrix(lngRow, COL_P_ID) & "," & .TextMatrix(lngRow, COL_P_申请量) & IIF(strIDs <> "", "," & strIDs, "")
                End If
            End If
        Next
    End With
    If Left(strRow, 1) = ";" Then strRow = Mid(strRow, 2)
    GetBloodInfo = strRow
End Function

Private Sub ReasonSelect(Optional ByVal strFind As String = "")
    Dim blnCancle As Boolean
    Dim strRetrun As String
    Dim lngLeft As Long, lngTop As Long
    Dim strName As String
    
    lngLeft = txtInfo(txt备注).Left
    lngTop = txtInfo(txt备注).Top
    strName = "常用嘱托。"
    
    lngLeft = lngLeft + Me.Left
    lngTop = lngTop + Me.Top - 2700
    
    strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, 2)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "没有找到可用的" & strName, vbInformation, Me.Caption
            End If
        Else
            txtInfo(txt备注).Text = strRetrun
        End If
    End If
End Sub

Private Sub Get输血执行科室()
'血液项目的执行科室，肯定是血库科室
    Dim bln上班安排 As Boolean
    Dim strSQL As String, bytDay As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    bln上班安排 = Check上班安排(False)
    If Not bln上班安排 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
            " From 部门表 A,部门性质说明 C" & _
            " Where  A.ID=C.部门ID" & _
            " And C.服务对象 IN([1],3) And C.工作性质='血库'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by 编码"
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
        strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
            " From 部门安排 B,部门表 C,部门性质说明 D" & _
            " Where  B.部门ID=C.ID And B.星期=[2]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
            " And C.ID=D.部门ID And D.服务对象 IN([1],3) And C.工作性质='血库'" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " Order by 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "", IIF(mlng病人性质 = 1, 1, 2), bytDay)
    cboInfo(cbo执行科室).Clear
    For i = 1 To rsTmp.RecordCount
        cboInfo(cbo执行科室).AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).NewIndex) = CLng(rsTmp!ID)
'        If lngDeptID = rsTmp!ID Then
'            Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
'        End If
        rsTmp.MoveNext
    Next
    If cboInfo(cbo执行科室).ListIndex = -1 And cboInfo(cbo执行科室).ListCount > 0 Then
        Call zlControl.CboSetIndex(cboInfo(cbo执行科室).hwnd, 0)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub BloodSum()
'功能：备血申请计算总量（可能多个品种的单位不一样，需要进行转换）
    Dim lngRow As Integer, i As Integer
    Dim str单位 As String, dblNum As Double, dbl计算系数 As Double
    Dim dblSum As Double '总量
    Dim strCur单位 As String, dblCur计算系数 As Double
    
    txt申请信息.Text = "品种:"
    txt申请量.Text = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_选择)) = 1 Then
            'iif(mid("" & 0.5,1,1)=".","0","") & 0.5，这种写法是为了保证小于的1的值能正常显示前缀0
            txt申请信息.Text = txt申请信息.Text & "[" & vsfBlood.TextMatrix(i, COL_P_名称) & IIF(vsfBlood.TextMatrix(i, COL_P_申请量) <> "", "-" & IIF(Mid("" & vsfBlood.TextMatrix(i, COL_P_申请量), 1, 1) = ".", "0", "") & vsfBlood.TextMatrix(i, COL_P_申请量) & vsfBlood.TextMatrix(i, COL_P_单位), "") & "]"
        End If
    Next
    
    If cboInfo(cbo单位).ListIndex >= 0 Then
        strCur单位 = UCase(cboInfo(cbo单位).List(cboInfo(cbo单位).ListIndex))
    End If
    
    With vsfBlood
        If strCur单位 <> "" Then
            For lngRow = .FixedRows To .Rows - 1
                If strCur单位 = UCase(.TextMatrix(lngRow, COL_P_单位)) Then
                    dblCur计算系数 = Val(.TextMatrix(lngRow, COL_P_计算系数))
                    Exit For
                End If
            Next
        End If
        If dblCur计算系数 <= 0 Then dblCur计算系数 = 1
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                str单位 = UCase(.TextMatrix(lngRow, COL_P_单位))
                dbl计算系数 = Val(.TextMatrix(lngRow, COL_P_计算系数))
                If dbl计算系数 <= 0 Then dbl计算系数 = 1
                dblNum = Val(.TextMatrix(lngRow, COL_P_申请量))
                If strCur单位 = "" Then
                    For i = 0 To cboInfo(cbo单位).ListCount - 1
                        If str单位 = UCase(cboInfo(cbo单位).List(i)) Then
                            Call zlControl.CboSetIndex(cboInfo(cbo单位).hwnd, i)
                            strCur单位 = str单位
                            dblCur计算系数 = dbl计算系数
                            cboInfo(cbo单位).Tag = i
                            Exit For
                        End If
                    Next
                End If
                If UCase(strCur单位) = UCase(str单位) Then
                    dblSum = dblSum + dblNum
                Else
                    If str单位 <> "ML" Then
                        dblNum = dblNum * dbl计算系数
                    End If
                    dblSum = dblSum + Format(dblNum / dblCur计算系数, "#0.00;-#0.00")
                End If
            End If
        Next
    End With
    If Val(dblSum) = 0 Then
        txtInfo(txt预定输血量) = ""
        txt申请量.Text = ""
    Else
        txtInfo(txt预定输血量).Text = zl9ComLib.FormatEx(dblSum, 5)
        txt申请量.Text = zl9ComLib.FormatEx(dblSum, 5)
    End If
    txtInfo(txt单位).Text = strCur单位
End Sub

Private Sub RsetBreedUnit()
'功能：单位切换如果医嘱默认品种的单位和选择单位不符合则进行品种切换(备血可选择多个品种，但医嘱记录只能记录一个品种)
    Dim lngRow As Long
    Dim strCur单位, str单位 As String
    If mlng输血项目ID <= 0 Then Exit Sub
    If cboInfo(cbo单位).ListIndex >= 0 Then
        strCur单位 = UCase(cboInfo(cbo单位).List(cboInfo(cbo单位).ListIndex))
    Else
        Exit Sub
    End If
    
    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If mlng输血项目ID = Val(.TextMatrix(lngRow, COL_P_ID)) Then
                str单位 = UCase(.TextMatrix(lngRow, COL_P_单位))
                Exit For
            End If
        Next
        If strCur单位 <> str单位 Then
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                    If strCur单位 = UCase(.TextMatrix(lngRow, COL_P_单位)) Then
                        mlng输血项目ID = Val(.TextMatrix(lngRow, COL_P_ID))
                        txtInfo(txt单位).Text = .TextMatrix(lngRow, COL_P_单位)
                        mlng录入限量 = Val(.TextMatrix(lngRow, COL_P_录入限量ID))
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Function GetBloodTotalByML() As Double
    '功能：本次申请输血总量(转换后为ML)
    Dim lngRow As Long
    Dim str单位 As String, dbl计算系数 As Double, dblNum As Double
    Dim dblTotal As Double
    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                str单位 = UCase(.TextMatrix(lngRow, COL_P_单位))
                dbl计算系数 = Val(.TextMatrix(lngRow, COL_P_计算系数))
                If dbl计算系数 <= 0 Then dbl计算系数 = 1
                dblNum = Val(.TextMatrix(lngRow, COL_P_申请量))
                If str单位 <> "ML" Then
                    dblNum = dblNum * dbl计算系数
                End If
                dblTotal = dblTotal + dblNum
            End If
        Next
    End With
    GetBloodTotalByML = dblTotal
End Function

Private Function GetPatiHisBloodItem() As String
'功能：获取病人本次就诊的历史申请血液品种信息
    Dim strSQL As String, strWhere As String
    Dim rsTmp As New ADODB.Recordset
    Dim strRetrun As String
    On Error GoTo ErrHand
     
    If mint场合 = 0 Then
        strWhere = " And b.病人id = [1] And b.主页id = [2] And Nvl(b.婴儿, 0) = [3] And B.id<>[4]"
    Else
        strWhere = " And b.病人id = [1] And b.挂号单 = [2] And Nvl(b.婴儿, 0) = [3] And B.id<>[4]"
    End If
    strSQL = _
        " Select 名称" & vbNewLine & _
        " From 诊疗项目目录 a," & vbNewLine & _
        "     (Select 诊疗项目id, Min(开始执行时间) 开始执行时间" & vbNewLine & _
        "       From (Select Decode(Nvl(c.医嘱id, 0), 0, b.诊疗项目id, c.诊疗项目id) 诊疗项目id, b.开始执行时间" & vbNewLine & _
        "              From 输血申请项目 c, 病人医嘱记录 b" & vbNewLine & _
        "              Where c.医嘱id(+) = b.Id And b.诊疗类别 = 'K' And b.医嘱状态 In (1, 3, 8) " & strWhere & " And Exists" & vbNewLine & _
        "               (Select 1" & vbNewLine & _
        "                     From 诊疗项目目录 p, 病人医嘱记录 q" & vbNewLine & _
        "                     Where p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And q.相关id = b.Id And q.诊疗类别 = 'E'))" & vbNewLine & _
        "       Group By 诊疗项目id) b" & vbNewLine & _
        " Where a.Id = b.诊疗项目id" & vbNewLine & _
        " Order By 开始执行时间"
    If mint场合 = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取备血信息", mlng病人ID, mlng主页ID, mbytBaby, mlngUpdateAdvice)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取备血信息", mlng病人ID, mstr挂号单, mbytBaby, mlngUpdateAdvice)
    End If
    Do While Not rsTmp.EOF
        strRetrun = strRetrun & "'" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If Left(strRetrun, 1) = "'" Then strRetrun = Mid(strRetrun, 2)
    GetPatiHisBloodItem = strRetrun
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetBloodApplyCode(ByVal int模式 As Integer) As String
'int模式:0=返回是否允许修改ABO和RH；1=返回ABO和RH指标打吗，便于根据检验结果更新ABO和RH，以及保存是检查ABO、RH是否和检验结果一致(输血申请单时有效)
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    gstrSQL = "Select Zl_Fun_BloodApplyCode([1],[2],[3]) as 指标 from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Zl_Fun_BloodApplyCode", IIF(mblnSpareBloood = True, 1, 2), cboInfo(cbo用血安排).ListIndex, int模式)
    If Not rsTemp.EOF Then
        strValue = "" & rsTemp!指标
    End If
    GetBloodApplyCode = strValue
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetLblRh()
    With cboInfo(cboRHD)
        lblRh.Caption = .Text
        Set lblRh.Font = .Font
        lblRh.Left = IIF(.Left < 0, Abs(.Left), 0)
        If .Top < 0 Then
            lblRh.Top = (.Height + .Top - lblRh.Height) \ 2
        Else
            lblRh.Top = (.Height - lblRh.Height) \ 2
        End If
        lblRh.ForeColor = IIF(.Text = "-", vbRed, vbBlack)
    
        picRH.Left = .Left
        picRH.Top = .Top
        picRH.Height = .Height
        picRH.Width = .Width - 300
        picRH.Visible = (.Enabled = False And .Text = "-")
        picRH.ZOrder 0
    End With
End Sub

Private Sub SetBloodLisAboRh(ByVal intIndex As Integer)
    Dim lngRow As Long
    If intIndex = cbo输血血型 Or intIndex = cboRHD Then
        With vsfBlood
            If .ColHidden(IIF(intIndex = cbo输血血型, COL_P_申请血型, COL_P_申请RH)) = True Then Exit Sub
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_选择)) = 1 Then
                    If intIndex = cbo输血血型 Then
                        If InStr(1, ",A,AB,B,O,", "," & cboInfo(intIndex).Text & ",") <> 0 Then
                            .TextMatrix(lngRow, COL_P_申请血型) = cboInfo(intIndex).Text
                        Else
                            .TextMatrix(lngRow, COL_P_申请血型) = ""
                        End If
                    Else
                        .TextMatrix(lngRow, COL_P_申请RH) = cboInfo(intIndex).Text
                    End If
                End If
            Next lngRow
        End With
    End If
End Sub
