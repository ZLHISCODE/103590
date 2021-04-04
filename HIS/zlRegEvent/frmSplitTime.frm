VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSplitTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间段设置"
   ClientHeight    =   4080
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "frmSplitTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdChange 
      Caption         =   "转为自定义(C)"
      Height          =   350
      Left            =   2430
      TabIndex        =   50
      Top             =   3675
      Width           =   1380
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   1260
      TabIndex        =   51
      Top             =   3675
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   49
      Top             =   3675
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5415
      TabIndex        =   48
      Top             =   3675
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   90
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   9105
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   15
      Width           =   9105
      Begin VB.TextBox txt后夜预留 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   8100
         TabIndex        =   19
         Top             =   2820
         Width           =   720
      End
      Begin VB.TextBox txt前夜预留 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   8100
         TabIndex        =   14
         Top             =   2370
         Width           =   720
      End
      Begin VB.TextBox txt下午预留 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   8100
         TabIndex        =   9
         Top             =   1935
         Width           =   720
      End
      Begin VB.TextBox txt上午预留 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   8100
         TabIndex        =   4
         Top             =   1470
         Width           =   720
      End
      Begin VB.PictureBox pic提前颜色 
         BackColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   3405
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   20
         Top             =   3240
         Width           =   270
      End
      Begin MSMask.MaskEdBox txt后夜S 
         Height          =   270
         Left            =   2505
         TabIndex        =   15
         Top             =   2820
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt后夜E 
         Height          =   270
         Left            =   3840
         TabIndex        =   16
         Top             =   2820
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt前夜S 
         Height          =   270
         Left            =   2505
         TabIndex        =   10
         Top             =   2370
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt前夜E 
         Height          =   270
         Left            =   3840
         TabIndex        =   11
         Top             =   2370
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt上午S 
         Height          =   270
         Left            =   2505
         TabIndex        =   0
         Top             =   1470
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt上午E 
         Height          =   270
         Left            =   3855
         TabIndex        =   1
         Top             =   1470
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt下午S 
         Height          =   270
         Left            =   2505
         TabIndex        =   5
         Top             =   1935
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt下午E 
         Height          =   270
         Left            =   3840
         TabIndex        =   6
         Top             =   1935
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt后夜缺省 
         Height          =   270
         Left            =   5265
         TabIndex        =   17
         Top             =   2820
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt前夜缺省 
         Height          =   270
         Left            =   5265
         TabIndex        =   12
         Top             =   2370
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt下午缺省 
         Height          =   270
         Left            =   5265
         TabIndex        =   7
         Top             =   1935
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt上午提前 
         Height          =   270
         Left            =   6705
         TabIndex        =   3
         Top             =   1470
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt下午提前 
         Height          =   270
         Left            =   6705
         TabIndex        =   8
         Top             =   1935
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt前夜提前 
         Height          =   270
         Left            =   6705
         TabIndex        =   13
         Top             =   2370
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt后夜提前 
         Height          =   270
         Left            =   6705
         TabIndex        =   18
         Top             =   2820
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt上午缺省 
         Height          =   270
         Left            =   5265
         TabIndex        =   2
         Top             =   1470
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "预留时间"
         Height          =   180
         Left            =   8100
         TabIndex        =   62
         Top             =   1110
         Width           =   720
      End
      Begin VB.Line Line7 
         X1              =   7935
         X2              =   7935
         Y1              =   1050
         Y2              =   3165
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "缺省时间"
         Height          =   180
         Left            =   5445
         TabIndex        =   61
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "放号时间"
         Height          =   180
         Left            =   6885
         TabIndex        =   60
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   4
         Left            =   3675
         TabIndex        =   59
         Top             =   1110
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "结束时间"
         Height          =   180
         Left            =   4020
         TabIndex        =   58
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "开始时间"
         Height          =   180
         Left            =   2685
         TabIndex        =   57
         Top             =   1110
         Width           =   720
      End
      Begin VB.Line Line6 
         X1              =   6540
         X2              =   6540
         Y1              =   1050
         Y2              =   3165
      End
      Begin VB.Line Line5 
         X1              =   5055
         X2              =   5055
         Y1              =   1050
         Y2              =   3165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "时间段"
         Height          =   180
         Left            =   1275
         TabIndex        =   56
         Top             =   1110
         Width           =   540
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   795
         X2              =   8970
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Label9 
         Caption         =   $"frmSplitTime.frx":000C
         Height          =   825
         Left            =   780
         TabIndex        =   55
         Top             =   165
         Width           =   5670
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "提前放号时间挂号安排显示颜色"
         Height          =   180
         Left            =   780
         TabIndex        =   52
         ToolTipText     =   "设置提前放号时间在挂号时号码列表中显示的颜色，若当前时间处于提前放号时间与开始时间之间时，用此颜色显示该号码，方便区分"
         Top             =   3285
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmSplitTime.frx":0108
         Top             =   337
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   1275
         X2              =   1275
         Y1              =   1380
         Y2              =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "后夜"
         Height          =   180
         Left            =   1920
         TabIndex        =   39
         Top             =   2865
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "前夜"
         Height          =   180
         Left            =   1920
         TabIndex        =   38
         Top             =   2430
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "上午"
         Height          =   180
         Left            =   1920
         TabIndex        =   37
         Top             =   1515
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "下午"
         Height          =   180
         Left            =   1920
         TabIndex        =   36
         Top             =   1980
         Width           =   360
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   1800
         Y1              =   1380
         Y2              =   3165
      End
      Begin VB.Line Line3 
         X1              =   2370
         X2              =   2370
         Y1              =   1050
         Y2              =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "全        日"
         Height          =   930
         Left            =   960
         TabIndex        =   35
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "白天"
         Height          =   180
         Left            =   1365
         TabIndex        =   34
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "夜间"
         Height          =   180
         Left            =   1365
         TabIndex        =   33
         Top             =   2625
         Width           =   360
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   1800
         X2              =   8970
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   1275
         X2              =   8970
         Y1              =   2265
         Y2              =   2265
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   1800
         X2              =   8970
         Y1              =   2730
         Y2              =   2730
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   0
         Left            =   3675
         TabIndex        =   32
         Top             =   1515
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   1
         Left            =   3675
         TabIndex        =   31
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   2
         Left            =   3675
         TabIndex        =   30
         Top             =   1980
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   3
         Left            =   3675
         TabIndex        =   29
         Top             =   2865
         Width           =   90
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         Height          =   2145
         Left            =   780
         Top             =   1035
         Width           =   8190
      End
   End
   Begin VB.PictureBox picCus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   9105
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.TextBox txt预留 
         Height          =   300
         Left            =   5025
         TabIndex        =   46
         Top             =   3255
         Width           =   870
      End
      Begin VB.PictureBox pic提前颜色 
         BackColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2850
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   45
         Top             =   3255
         Width           =   270
      End
      Begin MSMask.MaskEdBox txt结束 
         Height          =   300
         Left            =   2940
         TabIndex        =   24
         Top             =   2895
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt开始 
         Height          =   300
         Left            =   2010
         TabIndex        =   23
         Top             =   2895
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lvwSeg 
         Height          =   2595
         Left            =   210
         TabIndex        =   21
         Top             =   165
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "时间段"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "开始时间"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "终止时间"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "缺省预约时间"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "提前放号时间"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "预留时间"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.TextBox txt时间段 
         Height          =   300
         Left            =   825
         MaxLength       =   4
         TabIndex        =   22
         Top             =   2895
         Width           =   720
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   495
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSplitTime.frx":09D2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   7815
         TabIndex        =   27
         Top             =   1830
         Width           =   1100
      End
      Begin VB.CommandButton cmdModi 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   7815
         TabIndex        =   26
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   7815
         TabIndex        =   25
         Top             =   825
         Width           =   1100
      End
      Begin MSMask.MaskEdBox mkTxt缺省 
         Height          =   300
         Left            =   5025
         TabIndex        =   43
         Top             =   2895
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt提前 
         Height          =   300
         Left            =   7080
         TabIndex        =   44
         Top             =   2895
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:MM:SS"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束预留时间"
         Height          =   180
         Left            =   3900
         TabIndex        =   63
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "提前放号时间挂号安排显示颜色"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   3300
         Width           =   2520
      End
      Begin VB.Label lbl提前 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提前放号时间"
         Height          =   180
         Left            =   5955
         TabIndex        =   53
         Top             =   2955
         Width           =   1080
      End
      Begin VB.Label lbl缺省 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省预约时间"
         Height          =   180
         Left            =   3900
         TabIndex        =   47
         Top             =   2955
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "范围"
         Height          =   180
         Left            =   1590
         TabIndex        =   42
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间段"
         Height          =   180
         Left            =   240
         TabIndex        =   41
         Top             =   2955
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSplitTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrColor As String

Private Sub cmdAdd_Click()
    Dim ObjItem As ListItem, i As Integer
    
    If Not OneValid Then Exit Sub
    
    For i = 1 To lvwSeg.ListItems.Count
        If lvwSeg.ListItems(i).Text = txt时间段.Text Then
            MsgBox "该时间段名称已经存在！", vbInformation, gstrSysName
            txt时间段.SetFocus: Exit Sub
        End If
    Next
    If Check缺省时间 = False Then Exit Sub
    Set ObjItem = lvwSeg.ListItems.Add(, , txt时间段.Text, , 1)
    ObjItem.SubItems(1) = txt开始.Text
    ObjItem.SubItems(2) = txt结束.Text
    ObjItem.Selected = True
    ObjItem.EnsureVisible
    lvwSeg.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Dim ObjItem As ListItem
    
    If Not CheckValid Then Exit Sub
    
    If picStd.Visible Then
        lvwSeg.ListItems.Clear
        Set ObjItem = lvwSeg.ListItems.Add(, , "上午", , 1)
        ObjItem.SubItems(1) = txt上午S.Text
        ObjItem.SubItems(2) = txt上午E.Text
        ObjItem.SubItems(3) = txt上午缺省.Text
        ObjItem.SubItems(4) = txt上午提前.Text
        ObjItem.SubItems(5) = txt上午预留.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "下午", , 1)
        ObjItem.SubItems(1) = txt下午S.Text
        ObjItem.SubItems(2) = txt下午E.Text
        ObjItem.SubItems(3) = txt下午缺省.Text
        ObjItem.SubItems(4) = txt下午提前.Text
        ObjItem.SubItems(5) = txt下午预留.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "白天", , 1)
        ObjItem.SubItems(1) = txt上午S.Text
        ObjItem.SubItems(2) = txt下午E.Text
        ObjItem.SubItems(3) = txt上午缺省.Text
        ObjItem.SubItems(4) = txt上午提前.Text
        ObjItem.SubItems(5) = txt上午预留.Text
        ObjItem.Selected = True
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "前夜", , 1)
        ObjItem.SubItems(1) = txt前夜S.Text
        ObjItem.SubItems(2) = txt前夜E.Text
        ObjItem.SubItems(3) = txt前夜缺省.Text
        ObjItem.SubItems(4) = txt前夜提前.Text
        ObjItem.SubItems(5) = txt前夜预留.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "后夜", , 1)
        ObjItem.SubItems(1) = txt后夜S.Text
        ObjItem.SubItems(2) = txt后夜E.Text
        ObjItem.SubItems(3) = txt后夜缺省.Text
        ObjItem.SubItems(4) = txt后夜提前.Text
        ObjItem.SubItems(5) = txt后夜预留.Text
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "夜间", , 1)
        ObjItem.SubItems(1) = txt前夜S.Text
        ObjItem.SubItems(2) = txt后夜E.Text
        ObjItem.SubItems(3) = txt前夜缺省.Text
        ObjItem.SubItems(4) = txt前夜提前.Text
        ObjItem.SubItems(5) = txt前夜预留.Text
        ObjItem.Selected = True
        
        Set ObjItem = lvwSeg.ListItems.Add(, , "全日", , 1)
        ObjItem.SubItems(1) = txt上午S.Text
        ObjItem.SubItems(2) = txt后夜E.Text
        ObjItem.SubItems(3) = txt后夜缺省.Text
        ObjItem.SubItems(4) = txt后夜提前.Text
        ObjItem.SubItems(5) = txt后夜预留.Text
        ObjItem.Selected = True
                    
        txt上午S.Text = "__:__:__"
        txt下午S.Text = "__:__:__"
        txt前夜S.Text = "__:__:__"
        txt后夜S.Text = "__:__:__"
        
        lvwSeg.ListItems(1).Selected = True
        Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
        
        cmdChange.Caption = "转为标准(&C)"
        picStd.Visible = False
        picCus.Visible = True
        lvwSeg.SetFocus
    Else
        Call SetStandard
        
        cmdChange.Caption = "转为自定义(&C)"
        lvwSeg.ListItems.Clear
        picCus.Visible = False
        picStd.Visible = True
        txt上午S.SetFocus
    End If
End Sub

Private Sub cmdDel_Click()
    Dim intIdx As Integer
    
    If lvwSeg.SelectedItem Is Nothing Then
        MsgBox "没有可以删除的时间段！", vbInformation, gstrSysName
        lvwSeg.SetFocus: Exit Sub
    End If
    
    intIdx = lvwSeg.SelectedItem.index
    
    lvwSeg.ListItems.Remove intIdx
    
    If lvwSeg.ListItems.Count > 0 Then
        If intIdx <= lvwSeg.ListItems.Count Then
            lvwSeg.ListItems(intIdx).Selected = True
        Else
            lvwSeg.ListItems(lvwSeg.ListItems.Count).Selected = True
        End If
        lvwSeg.SelectedItem.EnsureVisible
        Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
    Else
        txt时间段.Text = ""
        txt开始.Text = "__:__:__"
        txt结束.Text = "__:__:__"
        mkTxt缺省.Text = "__:__:__"
        txt提前.Text = "__:__:__"
        txt预留.Text = ""
    End If
    
    lvwSeg.SetFocus
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function CheckValid() As Boolean
    If picStd.Visible Then
        If Not IsDate(txt上午S.Text) Then
            MsgBox "上午的开始时间设置不正确！", vbInformation, gstrSysName
            txt上午S.SetFocus: Exit Function
        End If
        If Not IsDate(txt上午E.Text) Then
            MsgBox "上午的终止时间设置不正确！", vbInformation, gstrSysName
            txt上午S.SetFocus: Exit Function
        End If
        If Not IsDate(txt上午缺省.Text) Then
            MsgBox "上午的缺省预约时间设置不正确！", vbInformation, gstrSysName
            txt上午缺省.SetFocus: Exit Function
        End If
        If Not IsDate(txt下午S.Text) Then
            MsgBox "下午的开始时间设置不正确！", vbInformation, gstrSysName
            txt下午S.SetFocus: Exit Function
        End If
        If Not IsDate(txt下午E.Text) Then
            MsgBox "下午的终止时间设置不正确！", vbInformation, gstrSysName
            txt下午S.SetFocus: Exit Function
        End If
        If Not IsDate(txt下午缺省.Text) Then
            MsgBox "下午的缺省预约时间设置不正确！", vbInformation, gstrSysName
            txt下午缺省.SetFocus: Exit Function
        End If
        
        If Not (IIf(txt上午S.Text = "00:00:00", "24:00:00", txt上午S.Text) < IIf(txt下午S.Text = "00:00:00", "24:00:00", txt下午S.Text)) Then
            MsgBox "上午开始时间应该小于下午开始时间！", vbInformation, gstrSysName
            txt上午S.SetFocus: Exit Function
        End If
        
        If DateDiff("n", txt上午S.Text, txt上午E.Text) < 0 Then
            If Val(txt上午预留.Text) >= Abs(DateDiff("n", "2000-01-01 " & txt上午S.Text, "2000-01-02 " & txt上午E.Text)) Then
                MsgBox "上午预留时间过长！", vbInformation, gstrSysName
                txt上午预留.SetFocus: Exit Function
            End If
        Else
            If Val(txt上午预留.Text) >= Abs(DateDiff("n", txt上午S.Text, txt上午E.Text)) Then
                MsgBox "上午预留时间过长！", vbInformation, gstrSysName
                txt上午预留.SetFocus: Exit Function
            End If
        End If
        
        If DateDiff("n", txt下午S.Text, txt下午E.Text) < 0 Then
            If Val(txt下午预留.Text) >= Abs(DateDiff("n", "2000-01-01 " & txt下午S.Text, "2000-01-02 " & txt下午E.Text)) Then
                MsgBox "下午预留时间过长！", vbInformation, gstrSysName
                txt下午预留.SetFocus: Exit Function
            End If
        Else
            If Val(txt下午预留.Text) >= Abs(DateDiff("n", txt下午S.Text, txt下午E.Text)) Then
                MsgBox "下午预留时间过长！", vbInformation, gstrSysName
                txt下午预留.SetFocus: Exit Function
            End If
        End If
        
        If DateDiff("n", txt前夜S.Text, txt前夜E.Text) < 0 Then
            If Val(txt前夜预留.Text) >= Abs(DateDiff("n", "2000-01-01 " & txt前夜S.Text, "2000-01-02 " & txt前夜E.Text)) Then
                MsgBox "前夜预留时间过长！", vbInformation, gstrSysName
                txt前夜预留.SetFocus: Exit Function
            End If
        Else
            If Val(txt前夜预留.Text) >= Abs(DateDiff("n", txt前夜S.Text, txt前夜E.Text)) Then
                MsgBox "前夜预留时间过长！", vbInformation, gstrSysName
                txt前夜预留.SetFocus: Exit Function
            End If
        End If
        
        If DateDiff("n", txt后夜S.Text, txt后夜E.Text) < 0 Then
            If Val(txt后夜预留.Text) >= Abs(DateDiff("n", "2000-01-01 " & txt后夜S.Text, "2000-01-02 " & txt后夜E.Text)) Then
                MsgBox "后夜预留时间过长！", vbInformation, gstrSysName
                txt后夜预留.SetFocus: Exit Function
            End If
        Else
            If Val(txt后夜预留.Text) >= Abs(DateDiff("n", txt后夜S.Text, txt后夜E.Text)) Then
                MsgBox "后夜预留时间过长！", vbInformation, gstrSysName
                txt后夜预留.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt上午提前.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt上午提前.Text) Then
                MsgBox "上午的放号时间设置不正确！", vbInformation, gstrSysName
                txt上午提前.SetFocus: Exit Function
            End If
            If Format(txt上午提前.Text, "HH:MM:SS") > Format(txt上午S.Text, "HH:MM:SS") Then
                MsgBox "上午的放号时间不能大于开始时间！", vbInformation, gstrSysName
                txt上午提前.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt下午提前.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt下午提前.Text) Then
                MsgBox "下午的放号时间设置不正确！", vbInformation, gstrSysName
                txt下午提前.SetFocus: Exit Function
            End If
            If Format(txt下午提前.Text, "HH:MM:SS") > Format(txt下午S.Text, "HH:MM:SS") Then
                MsgBox "下午的放号时间不能大于开始时间！", vbInformation, gstrSysName
                txt下午提前.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt前夜提前.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt前夜提前.Text) Then
                MsgBox "前夜的放号时间设置不正确！", vbInformation, gstrSysName
                txt前夜提前.SetFocus: Exit Function
            End If
            If Format(txt前夜提前.Text, "HH:MM:SS") > Format(txt前夜S.Text, "HH:MM:SS") Then
                MsgBox "前夜的放号时间不能大于开始时间！", vbInformation, gstrSysName
                txt前夜提前.SetFocus: Exit Function
            End If
        End If
        
        If Replace(Replace(txt后夜提前.Text, "_", ""), ":", "") <> "" Then
            If Not IsDate(txt后夜提前.Text) Then
                MsgBox "后夜的放号时间设置不正确！", vbInformation, gstrSysName
                txt后夜提前.SetFocus: Exit Function
            End If
            If Format(txt后夜提前.Text, "HH:MM:SS") > Format(txt后夜S.Text, "HH:MM:SS") Then
                MsgBox "后夜的放号时间不能大于开始时间！", vbInformation, gstrSysName
                txt后夜提前.SetFocus: Exit Function
            End If
        End If
        
        If Not IsDate(txt前夜S.Text) Then
            MsgBox "前夜的开始时间设置不正确！", vbInformation, gstrSysName
            txt前夜S.SetFocus: Exit Function
        End If
        If Not IsDate(txt前夜E.Text) Then
            MsgBox "前夜的终止时间设置不正确！", vbInformation, gstrSysName
            txt前夜S.SetFocus: Exit Function
        End If
        If Not IsDate(txt前夜缺省.Text) Then
            MsgBox "前夜的缺省预约时间设置不正确！", vbInformation, gstrSysName
            txt前夜缺省.SetFocus: Exit Function
        End If
        
        If Not IsDate(txt后夜S.Text) Then
            MsgBox "后夜的开始时间设置不正确！", vbInformation, gstrSysName
            txt后夜S.SetFocus: Exit Function
        End If
        If Not IsDate(txt后夜E.Text) Then
            MsgBox "后夜的终止时间设置不正确！", vbInformation, gstrSysName
            txt后夜S.SetFocus: Exit Function
        End If
        If Not IsDate(txt后夜缺省.Text) Then
            MsgBox "后夜的缺省预约时间设置不正确！", vbInformation, gstrSysName
            txt后夜缺省.SetFocus: Exit Function
        End If
        If Not (IIf(txt前夜S.Text = "00:00:00", "24:00:00", txt前夜S.Text) < IIf(txt后夜S.Text = "00:00:00", "24:00:00", txt后夜S.Text)) Then
            MsgBox "前夜开始时间应该小于后夜开始时间！", vbInformation, gstrSysName
            txt前夜S.SetFocus: Exit Function
        End If
    Else
        If lvwSeg.ListItems.Count = 0 Then
            MsgBox "必须至少设置一个时间段！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckValid = True
End Function

Private Function OneValid() As Boolean
    If Trim(txt时间段.Text) = "" Then
        MsgBox "必须输入时间段名称！", vbInformation, gstrSysName
        txt时间段.SetFocus: Exit Function
    End If
    If zlCommFun.ActualLen(txt时间段.Text) > 4 Then
        MsgBox "时间段名称只能为两个汉字或4个字母！", vbInformation, gstrSysName
        txt时间段.SetFocus: Exit Function
    End If
    If Not IsDate(txt开始.Text) Then
        MsgBox "开始时间设置不正确！", vbInformation, gstrSysName
        txt开始.SetFocus: Exit Function
    End If
    If Not IsDate(txt结束.Text) Then
        MsgBox "结束时间设置不正确！", vbInformation, gstrSysName
        txt结束.SetFocus: Exit Function
    End If
    If txt开始.Text = txt结束.Text Then
        MsgBox "开始和结束时间不应该相同！", vbInformation, gstrSysName
        txt结束.SetFocus: Exit Function
    End If
    
    If DateDiff("n", txt开始.Text, txt结束.Text) < 0 Then
        If Val(txt预留.Text) >= Abs(DateDiff("n", "2000-01-01 " & txt开始.Text, "2000-01-02 " & txt结束.Text)) Then
            MsgBox "预留时间过长！", vbInformation, gstrSysName
            txt预留.SetFocus: Exit Function
        End If
    Else
        If Val(txt预留.Text) >= Abs(DateDiff("n", txt开始.Text, txt结束.Text)) Then
            MsgBox "预留时间过长！", vbInformation, gstrSysName
            txt预留.SetFocus: Exit Function
        End If
    End If
    
    If Check缺省时间 = False Then Exit Function
    
    OneValid = True
End Function

Private Sub cmdModi_Click()
    Dim i As Integer
    
    If lvwSeg.SelectedItem Is Nothing Then
        MsgBox "没有时间段可以修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Not OneValid Then Exit Sub
    
    For i = 1 To lvwSeg.ListItems.Count
        If i <> lvwSeg.SelectedItem.index And lvwSeg.ListItems(i).Text = txt时间段.Text Then
            MsgBox "该时间段名称已经存在！", vbInformation, gstrSysName
            txt时间段.SetFocus: Exit Sub
        End If
    Next
    If Check缺省时间 = False Then Exit Sub
    
    lvwSeg.SelectedItem.Text = txt时间段.Text
    lvwSeg.SelectedItem.SubItems(1) = txt开始.Text
    lvwSeg.SelectedItem.SubItems(2) = txt结束.Text
    lvwSeg.SelectedItem.SubItems(3) = mkTxt缺省.Text
    lvwSeg.SelectedItem.SubItems(4) = txt提前.Text
    lvwSeg.SelectedItem.SubItems(5) = txt预留.Text
    
    lvwSeg.SetFocus
End Sub
Private Function Check缺省时间() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查缺省时间是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-12 14:46:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strDate As String
    
    If mkTxt缺省.Text Like "*_*" Then
        MsgBox "缺省预约时间有误,请重新输入!", vbInformation + vbOKOnly, gstrSysName
        If mkTxt缺省.Enabled Then mkTxt缺省.SetFocus
        Exit Function
    End If
    
    If IsDate(mkTxt缺省.Text) = False Then
        MsgBox "缺省预约时间格式不对,请重新输入!", vbInformation + vbOKOnly, gstrSysName
        If mkTxt缺省.Enabled Then mkTxt缺省.SetFocus
        Exit Function
    End If
    
    If Replace(Replace(txt提前.Text, "_", ""), ":", "") <> "" Then
        If IsDate(txt提前.Text) = False Then
            MsgBox "放号时间格式不对,请重新输入!", vbInformation + vbOKOnly, gstrSysName
            If txt提前.Enabled Then txt提前.SetFocus
            Exit Function
        End If
    End If
    
    strDate = "2010-01-01 "
    If CDate("2010-01-01 " & txt开始.Text) > CDate("2010-01-01 " & txt结束.Text) Then
        strDate = "2010-01-02 "
    End If
    
    If CDate(strDate & mkTxt缺省.Text) < CDate("2010-01-01 " & txt开始.Text) _
        Or CDate(strDate & mkTxt缺省.Text) > CDate(strDate & txt结束.Text) Then
        MsgBox "缺省预约时间必须在时间范围内,请重新输入!", vbInformation + vbOKOnly, gstrSysName
        If mkTxt缺省.Enabled Then mkTxt缺省.SetFocus
        Exit Function
    End If
    
    If Replace(Replace(txt提前.Text, "_", ""), ":", "") <> "" Then
        If Format(txt开始.Text, "HH:MM:SS") < Format(txt提前.Text, "HH:MM:SS") Then
            MsgBox "放号时间必须小于开始时间,请重新输入!", vbInformation + vbOKOnly, gstrSysName
            If txt提前.Enabled Then txt提前.SetFocus
            Exit Function
        End If
    End If
    
    Check缺省时间 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    Dim arrSQL() As String, i As Integer, blnTrans As Boolean
    
    '完整性检查
    If Not CheckValid Then Exit Sub
    
    ReDim arrSQL(0)
    arrSQL(0) = "zl_时间段_Clear"
    
    If picStd.Visible Then
        ReDim Preserve arrSQL(7)
        arrSQL(1) = "zl_时间段_INSERT('白天',To_Date('" & Format(txt上午S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt下午E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt下午缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt上午提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt上午提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt上午预留.Text) & ")"
        arrSQL(2) = "zl_时间段_INSERT('后夜',To_Date('" & Format(txt后夜S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt后夜提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt后夜提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt后夜预留.Text) & ")"
        arrSQL(3) = "zl_时间段_INSERT('前夜',To_Date('" & Format(txt前夜S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt前夜E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt前夜缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt前夜提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt前夜提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt前夜预留.Text) & ")"
        arrSQL(4) = "zl_时间段_INSERT('全日',To_Date('" & Format(txt上午S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & "Null,'" & mstrColor & "'," & Val(txt后夜预留.Text) & ")"
        arrSQL(5) = "zl_时间段_INSERT('上午',To_Date('" & Format(txt上午S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt上午E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt上午缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt上午提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt上午提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt上午预留.Text) & ")"
        arrSQL(6) = "zl_时间段_INSERT('下午',To_Date('" & Format(txt下午S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt下午E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt下午缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt下午提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt下午提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt下午预留.Text) & ")"
        arrSQL(7) = "zl_时间段_INSERT('夜间',To_Date('" & Format(txt前夜S.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜E.Text, "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(txt后夜缺省.Text, "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(txt前夜提前.Text, "_", ""), ":", "") = "", "Null", "To_Date('" & Format(txt前夜提前.Text, "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(txt前夜预留.Text) & ")"
    Else
        ReDim Preserve arrSQL(lvwSeg.ListItems.Count)
        For i = 1 To lvwSeg.ListItems.Count
            With lvwSeg.ListItems(i)
                arrSQL(i) = "zl_时间段_INSERT('" & .Text & "',To_Date('" & Format(.SubItems(1), "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(.SubItems(2), "HH:MM:SS") & "','HH24:MI:SS'),To_Date('" & Format(.SubItems(3), "HH:MM:SS") & "','HH24:MI:SS')," & IIf(Replace(Replace(.SubItems(4), "_", ""), ":", "") = "", "Null", "To_Date('" & Format(.SubItems(4), "HH:MM:SS") & "','HH24:MI:SS')") & ",'" & mstrColor & "'," & Val(.SubItems(5)) & ")"
            End With
        Next
    End If
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call FillData
End Sub

Public Function FillData() As Boolean
'功能:到所有时间段装入到msfTime
    Dim rsTime As New ADODB.Recordset
    Dim strSQL As String, ObjItem As ListItem
    
    Dim vBegin As Date, vEnd As Date, blnStd As Boolean

    On Error GoTo errH
    

    strSQL = "Select 时间段,To_Char(开始时间,'HH24:MI:SS') As 开始时间,to_char(终止时间,'HH24:MI:SS') As 终止时间 ,to_char(缺省时间,'HH24:MI:SS') As 缺省时间,to_char(提前时间,'HH24:MI:SS') As 提前时间,提前颜色,出诊预留时间 as 预留时间  From 时间段 Where 号类 Is Null And 站点 Is Null"
    Set rsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTime
        If Not .EOF Then
            '判断是否符合标准时间段的条件
            mstrColor = Nvl(rsTime!提前颜色)
            If mstrColor = "" Then mstrColor = &H0&
            pic提前颜色(0).BackColor = mstrColor
            pic提前颜色(1).BackColor = mstrColor
            
            blnStd = rsTime.RecordCount = 7
            If blnStd Then
                rsTime.Filter = "时间段='上午' or 时间段='下午' or 时间段='白天' or 时间段='前夜' or 时间段='后夜' or 时间段='夜间' or 时间段='全日'"
                blnStd = blnStd And rsTime.RecordCount = 7
            End If
            
            If blnStd Then
                '上午的结束=下午的开始
                rsTime.Filter = "时间段='上午'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='下午'": vBegin = rsTime!开始时间
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                '下午的结束=前夜的开始
                rsTime.Filter = "时间段='下午'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='前夜'": vBegin = rsTime!开始时间
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                '前夜的结束=后夜的开始
                rsTime.Filter = "时间段='前夜'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='后夜'": vBegin = rsTime!开始时间
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                
                '后夜的结束=上午的开始
                rsTime.Filter = "时间段='后夜'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='上午'": vBegin = rsTime!开始时间
                blnStd = blnStd And Format(vEnd + 1 / 24 / 60 / 60, "HH:mm:ss") = Format(vBegin, "HH:mm:ss")
                '--------------------------------------------------------------------------
                '白天的开始=上午的开始
                rsTime.Filter = "时间段='白天'": vEnd = rsTime!开始时间
                rsTime.Filter = "时间段='上午'": vBegin = rsTime!开始时间
                blnStd = blnStd And vEnd = vBegin
                
                '白天的结束=下午的结束
                rsTime.Filter = "时间段='白天'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='下午'": vBegin = rsTime!终止时间
                blnStd = blnStd And vEnd = vBegin
                
                '夜间的开始=前夜的开始
                rsTime.Filter = "时间段='夜间'": vEnd = rsTime!开始时间
                rsTime.Filter = "时间段='前夜'": vBegin = rsTime!开始时间
                blnStd = blnStd And vEnd = vBegin
                
                '夜间的结束=后夜的结束
                rsTime.Filter = "时间段='夜间'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='后夜'": vBegin = rsTime!终止时间
                blnStd = blnStd And vEnd = vBegin
                
                '全日的开始=上午的开始
                rsTime.Filter = "时间段='全日'": vEnd = rsTime!开始时间
                rsTime.Filter = "时间段='上午'": vBegin = rsTime!开始时间
                blnStd = blnStd And vEnd = vBegin
                
                '全日的结束=后夜的结束
                rsTime.Filter = "时间段='全日'": vEnd = rsTime!终止时间
                rsTime.Filter = "时间段='后夜'": vBegin = rsTime!终止时间
                blnStd = blnStd And vEnd = vBegin
            End If
            
            .Filter = 0
            .MoveFirst
            If blnStd Then
                Do Until .EOF
                    Select Case .Fields("时间段").Value
                    Case "后夜"
                        txt后夜S.Text = IIf(IsNull(.Fields("开始时间").Value), "__:__:__", .Fields("开始时间").Value)
                        txt后夜E.Text = IIf(IsNull(.Fields("终止时间").Value), "__:__:__", .Fields("终止时间").Value)
                        txt后夜缺省.Text = IIf(IsNull(.Fields("缺省时间").Value), "__:__:__", .Fields("缺省时间").Value)
                        txt后夜提前.Text = IIf(IsNull(.Fields("提前时间").Value), "__:__:__", .Fields("提前时间").Value)
                        txt后夜预留.Text = IIf(IsNull(.Fields("预留时间").Value), "", .Fields("预留时间").Value)
                    Case "前夜"
                        txt前夜S.Text = IIf(IsNull(.Fields("开始时间").Value), "__:__:__", .Fields("开始时间").Value)
                        txt前夜E.Text = IIf(IsNull(.Fields("终止时间").Value), "__:__:__", .Fields("终止时间").Value)
                        txt前夜缺省.Text = IIf(IsNull(.Fields("缺省时间").Value), "__:__:__", .Fields("缺省时间").Value)
                        txt前夜提前.Text = IIf(IsNull(.Fields("提前时间").Value), "__:__:__", .Fields("提前时间").Value)
                        txt前夜预留.Text = IIf(IsNull(.Fields("预留时间").Value), "", .Fields("预留时间").Value)
                    Case "上午"
                        txt上午S.Text = IIf(IsNull(.Fields("开始时间").Value), "__:__:__", .Fields("开始时间").Value)
                        txt上午E.Text = IIf(IsNull(.Fields("终止时间").Value), "__:__:__", .Fields("终止时间").Value)
                        txt上午缺省.Text = IIf(IsNull(.Fields("缺省时间").Value), "__:__:__", .Fields("缺省时间").Value)
                        txt上午提前.Text = IIf(IsNull(.Fields("提前时间").Value), "__:__:__", .Fields("提前时间").Value)
                        txt上午预留.Text = IIf(IsNull(.Fields("预留时间").Value), "", .Fields("预留时间").Value)
                    Case "下午"
                        txt下午S.Text = IIf(IsNull(.Fields("开始时间").Value), "__:__:__", .Fields("开始时间").Value)
                        txt下午E.Text = IIf(IsNull(.Fields("终止时间").Value), "__:__:__", .Fields("终止时间").Value)
                        txt下午缺省.Text = IIf(IsNull(.Fields("缺省时间").Value), "__:__:__", .Fields("缺省时间").Value)
                        txt下午提前.Text = IIf(IsNull(.Fields("提前时间").Value), "__:__:__", .Fields("提前时间").Value)
                        txt下午预留.Text = IIf(IsNull(.Fields("预留时间").Value), "", .Fields("预留时间").Value)
                    End Select
                    rsTime.MoveNext
                Loop
            Else
                Do Until .EOF
                    Set ObjItem = lvwSeg.ListItems.Add(, , !时间段, , 1)
                    ObjItem.SubItems(1) = IIf(IsNull(!开始时间), "__:__:__", !开始时间)
                    ObjItem.SubItems(2) = IIf(IsNull(!终止时间), "__:__:__", !终止时间)
                    ObjItem.SubItems(3) = IIf(IsNull(!缺省时间), "__:__:__", !缺省时间)
                    ObjItem.SubItems(4) = IIf(IsNull(!提前时间), "__:__:__", !提前时间)
                    ObjItem.SubItems(5) = IIf(IsNull(!预留时间), "", !预留时间)
                    rsTime.MoveNext
                Loop
                lvwSeg.ListItems(1).Selected = True
                Call lvwSeg_ItemClick(lvwSeg.SelectedItem)
                
                cmdChange.Caption = "转为标准(&C)"
                picStd.Visible = False
                picCus.Visible = True
            End If
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ActiveControl.Name = "cmdOK" Or ActiveControl.Name = "cmdCancel" Then Exit Sub
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub lvwSeg_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt时间段.Text = Item.Text
    txt开始.Text = Item.SubItems(1)
    txt结束.Text = Item.SubItems(2)
    mkTxt缺省.Text = Item.SubItems(3)
    txt提前.Text = Item.SubItems(4)
    txt预留.Text = Item.SubItems(5)
End Sub

Private Sub pic提前颜色_Click(index As Integer)
    dlgColor.ShowColor
    mstrColor = dlgColor.Color
    pic提前颜色(0).BackColor = mstrColor
    pic提前颜色(1).BackColor = mstrColor
End Sub

Private Sub txt后夜S_GotFocus()
    zlControl.TxtSelAll txt后夜S
End Sub

Private Sub txt后夜S_LostFocus()
    If IsDate(txt后夜S.Text) Then
        Me.txt前夜E.Text = Format(DateAdd("s", -1, CDate(Me.txt后夜S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txt后夜预留_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt结束_Change()
    If mkTxt缺省.Text Like "*_*" Then mkTxt缺省.Text = txt结束
End Sub

Private Sub txt结束_GotFocus()
   zlControl.TxtSelAll txt结束
End Sub

Private Sub txt开始_GotFocus()
   zlControl.TxtSelAll txt开始
End Sub

Private Sub txt前夜S_GotFocus()
   zlControl.TxtSelAll txt前夜S
End Sub

Private Sub txt前夜S_LostFocus()
    If IsDate(txt前夜S.Text) Then
        Me.txt下午E.Text = Format(DateAdd("s", -1, CDate(Me.txt前夜S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txt前夜预留_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt上午S_GotFocus()
   zlControl.TxtSelAll txt上午S
End Sub

Private Sub txt上午S_LostFocus()
    If IsDate(txt上午S.Text) Then
        Me.txt后夜E.Text = Format(DateAdd("s", -1, CDate(Me.txt上午S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub txt上午预留_GotFocus()
   zlControl.TxtSelAll txt上午预留
End Sub

Private Sub txt上午预留_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt下午预留_GotFocus()
   zlControl.TxtSelAll txt下午预留
End Sub

Private Sub txt前夜预留_GotFocus()
   zlControl.TxtSelAll txt前夜预留
End Sub

Private Sub txt后夜预留_GotFocus()
   zlControl.TxtSelAll txt后夜预留
End Sub

Private Sub txt时间段_GotFocus()
   zlControl.TxtSelAll txt时间段
End Sub

Private Sub txt提前_GotFocus()
   zlControl.TxtSelAll txt提前
End Sub

Private Sub txt下午S_GotFocus()
   zlControl.TxtSelAll txt下午S
End Sub

Private Sub txt下午S_LostFocus()
    If IsDate(Me.txt下午S.Text) Then
        Me.txt上午E.Text = Format(DateAdd("s", -1, CDate(Me.txt下午S.Text)), "HH:mm:ss")
    End If
End Sub

Private Sub SetStandard()
'功能：将自定义时间段转换为标准时间段
    Dim i As Integer
    
    For i = 1 To lvwSeg.ListItems.Count
        With lvwSeg.ListItems(i)
            Select Case .Text
                Case "上午"
                    txt上午S.Text = .SubItems(1)
                    txt上午E.Text = .SubItems(2)
                Case "下午"
                    txt下午S.Text = .SubItems(1)
                    txt下午E.Text = .SubItems(2)
                Case "白天", "白日"
                    If Not IsDate(txt上午S.Text) Then txt上午S.Text = .SubItems(1)
                    If Not IsDate(txt下午E.Text) Then txt下午E.Text = .SubItems(2)
                Case "前夜", "上夜"
                    txt前夜S.Text = .SubItems(1)
                    txt前夜E.Text = .SubItems(2)
                Case "后夜", "下夜"
                    txt后夜S.Text = .SubItems(1)
                    txt后夜E.Text = .SubItems(2)
                Case "夜间", "晚上"
                    If Not IsDate(txt前夜S.Text) Then txt前夜S.Text = .SubItems(1)
                    If Not IsDate(txt后夜E.Text) Then txt后夜E.Text = .SubItems(2)
                Case "全日", "全天"
                    If Not IsDate(txt上午S.Text) Then txt上午S.Text = .SubItems(1)
                    If Not IsDate(txt后夜E.Text) Then txt后夜E.Text = .SubItems(2)
            End Select
        End With
    Next
End Sub

Private Sub txt下午预留_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt预留_GotFocus()
   zlControl.TxtSelAll txt预留
End Sub

Private Sub txt预留_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
