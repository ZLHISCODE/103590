VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseRegist 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "传染病阳性结果反馈单"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   Icon            =   "frmDiseaseRegist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13455.55
   ScaleMode       =   0  'User
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11175
      TabIndex        =   54
      Top             =   0
      Width           =   11175
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.VScrollBar vsbReport 
      Height          =   7335
      LargeChange     =   50
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar hsbReport 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Frame frmMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   240
      TabIndex        =   52
      Top             =   360
      Width           =   10455
      Begin MSComCtl2.MonthView dtpDate 
         Bindings        =   "frmDiseaseRegist.frx":6852
         Height          =   2220
         Left            =   4440
         TabIndex        =   51
         Top             =   4200
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         StartOfWeek     =   41091073
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   0
         Left            =   3910
         Picture         =   "frmDiseaseRegist.frx":6866
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   5
         Left            =   3090
         Picture         =   "frmDiseaseRegist.frx":6B80
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   4
         Left            =   9810
         Picture         =   "frmDiseaseRegist.frx":6E9A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2045
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   3
         Left            =   6450
         Picture         =   "frmDiseaseRegist.frx":71B4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2045
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   2
         Left            =   6450
         Picture         =   "frmDiseaseRegist.frx":74CE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   1
         Left            =   3090
         Picture         =   "frmDiseaseRegist.frx":77E8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2035
         Width           =   250
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
         Height          =   210
         Index           =   6
         Left            =   4660
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2045
         Width           =   1725
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
         Height          =   210
         Index           =   11
         Left            =   8020
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2045
         Width           =   1725
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
         Height          =   210
         Index           =   16
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   6180
         Width           =   1725
      End
      Begin VB.Frame fraIdea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   0
         TabIndex        =   53
         Top             =   6570
         Width           =   10700
         Begin VB.ComboBox cboReport 
            BackColor       =   &H8000000E&
            Height          =   300
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1800
            Width           =   3255
         End
         Begin VB.TextBox txtInfo 
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
            Height          =   1440
            Index           =   18
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   180
            Width           =   8685
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
            Height          =   210
            Index           =   9
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   2280
            Width           =   1725
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
            Height          =   210
            Index           =   10
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1800
            Width           =   1725
         End
         Begin VB.Label lblReport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "反馈单关联的报告卡:"
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
            Height          =   255
            Left            =   300
            TabIndex        =   43
            Top             =   1850
            Width           =   2175
         End
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   -30
            X2              =   11130
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "处理情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   41
            Top             =   195
            Width           =   900
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "确认医师"
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
            Index           =   10
            Left            =   7320
            TabIndex        =   45
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "处理时间"
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
            Index           =   11
            Left            =   7320
            TabIndex        =   47
            Top             =   2325
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   8160
            X2              =   10035
            Y1              =   2505
            Y2              =   2505
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   8160
            X2              =   10035
            Y1              =   2025
            Y2              =   2025
         End
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   1410
         Index           =   17
         Left            =   1350
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2845
         Width           =   8685
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
         Height          =   210
         Index           =   7
         Left            =   6340
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
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
         Height          =   210
         Index           =   5
         Left            =   6340
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
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
         Height          =   210
         Index           =   4
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
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
         Height          =   210
         Index           =   3
         Left            =   3820
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
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
         Height          =   210
         Index           =   2
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
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
         Height          =   210
         Index           =   1
         Left            =   3820
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
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
         Height          =   210
         Index           =   19
         Left            =   1300
         TabIndex        =   16
         Top             =   2045
         Width           =   1725
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
         Height          =   210
         Index           =   12
         Left            =   4660
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2445
         Width           =   1725
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
         Height          =   210
         Index           =   8
         Left            =   1320
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4445
         Width           =   2525
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
         Height          =   210
         Index           =   13
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   5655
         Width           =   1725
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
         Height          =   210
         Index           =   14
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5130
         Width           =   1725
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
         Height          =   210
         Index           =   15
         Left            =   1300
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2445
         Width           =   1725
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
         Height          =   210
         Index           =   0
         Left            =   8860
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   4605
         X2              =   6510
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检科室"
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
         Index           =   16
         Left            =   3720
         TabIndex        =   18
         Top             =   2045
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检医生"
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
         Index           =   6
         Left            =   7080
         TabIndex        =   21
         Top             =   2045
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   7965
         X2              =   9870
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
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
         Index           =   36
         Left            =   7530
         TabIndex        =   35
         Top             =   5160
         Width           =   630
      End
      Begin VB.Line Line1 
         Index           =   31
         X1              =   8175
         X2              =   10030
         Y1              =   6405
         Y2              =   6405
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
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
         Index           =   34
         Left            =   7320
         TabIndex        =   39
         Top             =   6180
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登记科室"
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
         Index           =   33
         Left            =   7320
         TabIndex        =   37
         Top             =   5655
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   1245
         X2              =   3960
         Y1              =   4675
         Y2              =   4675
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "反馈结果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   13
         Left            =   300
         TabIndex        =   30
         Top             =   2845
         Width           =   900
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "疑似疾病"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   19
         Left            =   300
         TabIndex        =   32
         Top             =   4445
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6285
         X2              =   7515
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8805
         X2              =   10035
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6285
         X2              =   7515
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3765
         X2              =   4995
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   1245
         X2              =   2475
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1245
         X2              =   2475
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   3765
         X2              =   4995
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3240
         X2              =   7080
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "传染病阳性结果反馈单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3270
         TabIndex        =   0
         Top             =   360
         Width           =   3750
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   5400
         TabIndex        =   13
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科    室"
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
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   5400
         TabIndex        =   5
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检时间"
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
         Index           =   35
         Left            =   360
         TabIndex        =   15
         Top             =   2045
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   1245
         X2              =   3150
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病    情"
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
         Index           =   8
         Left            =   7920
         TabIndex        =   7
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标本名称"
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
         Index           =   9
         Left            =   3750
         TabIndex        =   27
         Top             =   2445
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4605
         X2              =   6510
         Y1              =   2675
         Y2              =   2675
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   8175
         X2              =   10030
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   8175
         X2              =   10030
         Y1              =   5355
         Y2              =   5355
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   1245
         X2              =   3150
         Y1              =   2675
         Y2              =   2675
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检查时间"
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
         Index           =   12
         Left            =   360
         TabIndex        =   24
         Top             =   2445
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmDiseaseRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'本窗体之所以把CommandBars放在PictureBox里面，是因为该窗体在其他地方会被镶嵌到TabControl里面，
'不把CommandBars放在PictureBox里面的会出现问题
'***************************************************************************************************
Public Event PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String) '转科
Public Event Closed(ByVal lngFunID As Long, ByVal strTag As String) '如果填写了诊断，关闭窗体时触发lngFunID固定为0；strTag 扩展参数未使用 ""。

Private Enum mCtlID '界面上的控件索引值
    txt病情 = 0
    txt住院号 = 1
    txt姓名 = 2
    txt性别 = 3
    txt科室 = 4
    txt床号 = 5
    txt送检科室 = 6
    txt年龄 = 7
    txt疑似疾病 = 8
    txt处理时间 = 9
    txt确认医师 = 10
    txt送检医生 = 11
    txt标本名称 = 12
    txt登记科室 = 13
    txt登记人 = 14
    txt检查时间 = 15
    txt登记时间 = 16
    txt反馈结果 = 17
    txt处理情况 = 18
    txt送检时间 = 19
    
    cmd疑似疾病 = 0
    cmd送检时间 = 1
    cmd标本名称 = 2
    cmd送检科室 = 3
    cmd送检医生 = 4
    cmd检查时间 = 5
End Enum

Private mlngID As Long   '疾病阳性记录 表的ID
Private mint场合 As Integer    '0-住院，1-门诊
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr挂号NO As String
Private mlng挂号ID As Long
Private mlng病区ID As Long '病区ID－－ 病案主页.当前病区ID
Private mstr诊室 As String '门诊病人诊室 病人挂号记录.诊室
Private mlng登记科室ID As Long
Private mdat送检时间 As Date
Private mdat检查时间 As Date
Private mlng送检科室ID As Long
Private mstr送检医生 As String
Private mstr标本名称 As String
Private mstr反馈结果 As String
Private mstr疑似传染病 As String
Private mintType As Integer  '0表示填写（只显示上半部分），1-表示医生处理（只有下半部分可编辑），2-查看（所有不可编辑，可查看下半部分）,3-修改（可编辑，可查看上半部分）
Private mlng科室ID As Long
Private mIntState As Integer '反馈单的当前状态：0-正在填写；1-待医生确认，2-医生已处理,3-非传染病，4-转科待处理
Private mblnOk As Boolean

Private mdat登记时间 As Date
Private mstr登记人 As String
Private mstr处理情况 As String

Private mstr就诊时间 As Date
Private mintResult As Integer   '1-发送，2-完成，3-转诊
Private WithEvents mclsDiagEdit As zlMedRecPage.clsDiagEdit
Attribute mclsDiagEdit.VB_VarHelpID = -1
Private mclsMipModule As zl9ComLib.clsMipModule
Private mblnDiagnose As Boolean     '是否填写了诊断

Private mblnDialog As Boolean        '是否显示为窗体
Private mlngTop As Long              '不显示为窗体时，上边距
Private mblnSbSisible As Boolean     '不显示为窗体时，滚动条是否可见
Private mlng医嘱ID As Long           '填写该反馈单的医嘱ID
Private mstr类别 As String           '阳性结果对应的医嘱的类别 D/E   检查/检验(采集方式)
Private mblnNoID As Boolean

Public Function ShowDiseaseRegist(ByRef frmParent As Object, ByVal intType As Integer, Optional ByVal lngID As Long, _
                Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号No As String, _
                Optional ByVal lng医嘱id As Long, Optional ByVal var登记科室 As Variant, Optional ByVal dat送检时间 As Date, Optional ByVal var送检科室 As Variant, _
                Optional ByVal str送检医生 As String, Optional ByVal str标本名称 As String, Optional ByVal str反馈结果 As String, _
                Optional ByVal dat检查时间 As Date, Optional ByVal str疑似传染病 As String, Optional ByRef objMip As Object, Optional ByVal str登记人 As String) As Integer
'功能：调用传染病阳性结果反馈单
'参数：intType 0表示填写（只显示上半部分），1-表示医生处理（只有下半部分可编辑），2-查看（所有不可编辑，可查看下半部分）,3-修改（可编辑，可查看上半部分）
'      lngID  = 疾病阳性记录 ID
'      lng病人ID = 病人ID
'      lng主页ID=住院:主页ID
'      str挂号No =门诊：挂号单NO
'      lng登记科室ID 和 lng送检科室ID 之所以是可变数据类型 是为了兼容LIS独立安装
    mintType = intType
    mlngID = lngID
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr挂号NO = str挂号No
    mlng医嘱ID = lng医嘱id
    mdat送检时间 = dat送检时间
    mstr送检医生 = str送检医生
    mstr标本名称 = str标本名称
    mstr反馈结果 = str反馈结果
    mdat检查时间 = dat检查时间
    mstr疑似传染病 = str疑似传染病
    mstr登记人 = str登记人
    
    If TypeName(var送检科室) = "String" Then         '传的编码
        mlng送检科室ID = GetDeptID(var送检科室)
    ElseIf IsNumeric(var送检科室) Then
        mlng送检科室ID = Val(var送检科室)
    Else
        mlng送检科室ID = 0
    End If
    
    If TypeName(var登记科室) = "String" Then
        mlng登记科室ID = GetDeptID(var登记科室)
    ElseIf IsNumeric(var登记科室) Then
        mlng登记科室ID = Val(var登记科室)
    Else
        mlng登记科室ID = 0
    End If
    
    mintResult = 0
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip

     '判断是住院病人还是门诊病人
    If intType = 0 Then
        If mlng主页ID = 0 And str挂号No <> "" Then
            mint场合 = 1
        ElseIf mlng主页ID <> 0 And str挂号No = "" Then
            mint场合 = 0
        Else
            Call MsgBox("不处理门诊和出院以外的病人!", vbInformation, gstrSysName)
            Exit Function
        End If
    End If
    mblnDialog = True
    On Error Resume Next
    Me.Show 1, frmParent
        
    ShowDiseaseRegist = mintResult
End Function


Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
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
    Set cbsMain.Icons = gobjComlib.zlCommFun.GetPubIcons

    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Send, "发送")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OK, "确认为传染病")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_NO, "非传染病")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ViewReport, "报告查看"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Transfer, "转诊"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Function SetDiagnose() As Boolean
    Dim lng就诊ID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancle As Boolean
    Dim mstr诊断IDs As String, str诊断 As String
On Error GoTo errH

    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIf(mint场合 = 1, 1260, 1261), gclsMipModule)
    End If
    
    lng就诊ID = IIf(mint场合 = 0, mlng主页ID, mlng挂号ID)

    strSQL = "Select rownum as ID, a.疾病id, a.诊断id, b.编码 || '-' || b.名称 As 疾病, c.编码 || '-' || c.名称 As 诊断, a.报告病种" & vbNewLine & _
            "From 疾病报告前提 A, 疾病编码目录 B, 疾病诊断目录 C" & vbNewLine & _
            "Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.报告病种 = [1]"
    Set rsTemp = gobjComlib.FS.ShowSQLSelectEx(Me, txtInfo(2), strSQL, 0, "诊断", False, "", "疑似疾病的诊断选择", False, False, False, blnCancle, True, False, True, "MultiCheckReturn=1", mstr疑似传染病)
    
    If Not blnCancle Then
        If rsTemp Is Nothing Then
            SetDiagnose = mclsDiagEdit.ShowDiagEdit(Me, mlngID, mlng病人ID, lng就诊ID, IIf(mint场合 = 1, 1, 2), mlng科室ID, mstr诊断IDs, str诊断, 0)
        Else
            SetDiagnose = mclsDiagEdit.ConfirmInfectiousDiseases(Me, mlngID, mlng病人ID, lng就诊ID, IIf(mint场合 = 1, 1, 2), mlng科室ID, rsTemp)
        End If
    End If
    SetDiagnose = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Print
            If mintType = 0 Or mintType = 3 Then
                If CheckData Then
                    If SendDiseaseRecord Then
                        Call PrintDiseaseRegist(2, mlngID, Me)
                        mintType = 3
                    End If
                End If
            ElseIf mintType = 1 Then
                If SaveDisProcessData(6) Then
                    Call PrintDiseaseRegist(2, mlngID, Me)
                End If
            Else
                Call PrintDiseaseRegist(2, mlngID, Me)
            End If
        Case conMenu_Tool_Send
            If CheckData Then
                If SendDiseaseRecord Then
                    If mintType = 0 Then
                        Call SendMsg
                    End If
                    mintResult = 1
                    Unload Me
                End If
            End If
        Case conMenu_Tool_OK
            If SaveDisProcessData(2) Then
                 mintResult = 2
                 Call SetDiagnose
                 mblnOk = True
                 Unload Me
            End If
        Case conMenu_Tool_NO
            If cboReport.Text <> "" Then
                If MsgBox("该反馈单已经关联了报告卡，确认为非传染病将取消关联，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If SaveDisProcessData(3) Then
                mintResult = 2
                Unload Me
            End If
        Case conMenu_Tool_Transfer
            If SaveDisProcessData(4) Then
                RaiseEvent PatiTransfer(mlng病人ID, mstr挂号NO)
                mintResult = 3
                Unload Me
            End If
        Case conMenu_Tool_ViewReport
            Call ViewEPRReport
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not mblnDialog Then Exit Sub
    Select Case Control.ID
        Case conMenu_Tool_Send
            Control.Visible = (mintType = 0 Or mintType = 3)
        Case conMenu_Tool_OK
            Control.Visible = (mintType = 1)
        Case conMenu_Tool_NO
            Control.Visible = (mintType = 1)
        Case conMenu_Tool_Transfer
            Control.Visible = (mintType = 1)
            If Control.Visible Then Control.Visible = (mint场合 = 1)
        Case conMenu_Tool_ViewReport
            Control.Visible = (mintType = 1)
    End Select
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Select Case Index
        Case cmd疑似疾病
            Call GetDiseaseList(1)
        Case cmd送检时间
            If IsDate(txtInfo(txt送检时间).Text) Then
                dtpDate.Value = CDate(txtInfo(txt送检时间).Text)
            Else
                dtpDate.Value = gobjComlib.zlDatabase.Currentdate
            End If
            dtpDate.Tag = "送检时间"
            dtpDate.Left = txtInfo(txt送检时间).Left
            dtpDate.Top = txtInfo(txt送检时间).Top + txtInfo(txt送检时间).Height
            dtpDate.Visible = True
            dtpDate.SetFocus
        Case cmd标本名称
            Call GetSampleList(1)
        Case cmd送检科室
            Call GetInspectDept(1)
        Case cmd送检医生
            Call GetInspectDoctor(1)
        Case cmd检查时间
            If IsDate(txtInfo(txt检查时间).Text) Then
                dtpDate.Value = CDate(txtInfo(txt检查时间).Text)
            Else
                dtpDate.Value = gobjComlib.zlDatabase.Currentdate
            End If
            dtpDate.Tag = "检查时间"
            dtpDate.Left = txtInfo(txt检查时间).Left
            dtpDate.Top = txtInfo(txt检查时间).Top + txtInfo(txt检查时间).Height
            dtpDate.Visible = True
            dtpDate.SetFocus
    End Select
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "送检时间" Then
        '取值
        If IsDate(txtInfo(txt送检时间).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt送检时间).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt送检时间).Text = strDate
        txtInfo(txt送检时间).Tag = strDate
        mdat送检时间 = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt送检时间).SetFocus
    ElseIf dtpDate.Tag = "检查时间" Then
        '取值
        If IsDate(txtInfo(txt检查时间).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt检查时间).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt检查时间).Text = strDate
        txtInfo(txt检查时间).Tag = strDate
        mdat检查时间 = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt检查时间).SetFocus
    End If
End Sub

Private Sub Form_Load()
    picMenu.Visible = mblnDialog
    mblnNoID = False
    If mblnDialog Then
        Me.BorderStyle = 3
        lblReport.Visible = True
        cboReport.Visible = True
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIf(mint场合 = 1, 1260, 1261), gclsMipModule)
        Call InitCommandBar
        Call SetFormState(mintType)
        Call LoadPatiInfo
        If mblnNoID Then
            Unload Me
        Else
            Call SaveDisProcessData(1)
        End If
    Else
        lblReport.Visible = False
        cboReport.Visible = False
    End If
End Sub

'获取登记人，登记时间，登记科室
Private Sub GetRegistInfo()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
On Error GoTo errH
    
    '读取登记科室
    strSQL = "Select a.Id, a.名称 From 部门表 A Where ID = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng登记科室ID)
    If rsTmp.RecordCount > 0 Then
        mlng登记科室ID = Val(rsTmp!ID)
        txtInfo(txt登记科室).Text = rsTmp!名称 & ""
    Else
        mlng登记科室ID = UserInfo.部门ID
        txtInfo(txt登记科室).Text = UserInfo.部门名
    End If

    If mstr登记人 = "" Then mstr登记人 = UserInfo.姓名
    mdat登记时间 = gobjComlib.zlDatabase.Currentdate
    txtInfo(txt登记人).Text = mstr登记人
    txtInfo(txt登记时间).Text = Format(mdat登记时间, "yyyy-MM-dd HH:mm")

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDiseaseList(ByVal intType As Integer)
'功能：获取疑似疾病目录
'参数：0 文本框按回车，1 点按钮
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt疑似疾病).Tag = txtInfo(txt疑似疾病).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt疑似疾病).Text = "" Then  '相当于是清除该项目
            txtInfo(txt疑似疾病).Tag = ""
            mstr疑似传染病 = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "select A.编码 as ID,A.简码,A.名称 from 传染病目录 A" & IIf(intType = 0, " where A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2] ", "") & " order by A.编码"
        
    strInput = Trim(UCase(txtInfo(txt疑似疾病).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt疑似疾病).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "传染病", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt疑似疾病).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的疾病!", vbInformation, gstrSysName)
            txtInfo(txt疑似疾病).SetFocus
            mstr疑似传染病 = ""
            gobjComlib.zlControl.TxtSelAll txtInfo(txt疑似疾病)
        End If
        Exit Sub
    Else
        mstr疑似传染病 = rsTmp!名称 & ""
        txtInfo(txt疑似疾病).Text = rsTmp!名称 & ""
        txtInfo(txt疑似疾病).Tag = rsTmp!名称 & ""
        txtInfo(txt疑似疾病).SetFocus
        Call SeekNextCtl
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatiInfo()
'功能：提取病人基本信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '读取病人相关信息
    If mintType <> 0 Then
        strSQL = "Select a.Id, a.病人id, a.主页id, a.挂号单, a.医嘱ID,a.送检时间, a.送检科室id, a.送检医生, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, " & vbNewLine & _
                " a.检查时间, a.登记时间, a.登记人, a.登记科室id,a.记录状态, a.处理人, a.处理时间, a.处理情况说明,b.诊疗类别" & vbNewLine & _
                "From 疾病阳性记录 A,病人医嘱记录 b Where a.医嘱ID=b.id(+) and a.Id = [1]"

        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        
        If rsTmp.RecordCount > 0 Then
            mlng病人ID = Val(rsTmp!病人ID & "")
            mlng主页ID = Val(rsTmp!主页ID & "")
            mstr挂号NO = rsTmp!挂号单 & ""
            mlng医嘱ID = Val(rsTmp!医嘱id & "")
            mIntState = Val(rsTmp!记录状态 & "")
            mstr类别 = rsTmp!诊疗类别 & ""
            If mlng主页ID = 0 And mstr挂号NO <> "" Then
                mint场合 = 1
            Else
                mint场合 = 0
            End If

            If IsDate(rsTmp!登记时间 & "") Then
                txtInfo(txt登记时间).Text = Format(rsTmp!登记时间 & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!处理时间 & "") Then
                txtInfo(txt处理时间).Text = Format(rsTmp!处理时间 & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!送检时间 & "") Then
                mdat送检时间 = Format(rsTmp!送检时间 & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!检查时间 & "") Then
                mdat检查时间 = Format(rsTmp!检查时间 & "", "YYYY-MM-DD HH:mm")
            End If
            
            txtInfo(txt登记人).Text = rsTmp!登记人 & ""
            txtInfo(txt确认医师).Text = rsTmp!处理人 & ""
            txtInfo(txt处理情况).Text = rsTmp!处理情况说明 & ""
            mstr疑似传染病 = rsTmp!疑似疾病 & ""
            mlng送检科室ID = Val(rsTmp!送检科室ID & "")
            mlng登记科室ID = Val(rsTmp!登记科室ID & "")
            mstr送检医生 = rsTmp!送检医生 & ""
            mstr标本名称 = rsTmp!标本名称 & ""
            mstr反馈结果 = rsTmp!反馈结果 & ""
        Else
            mblnNoID = True
        End If
    End If

    If mint场合 = 0 Then
        strSQL = "Select A.住院号, Nvl(C.姓名, A.姓名) 姓名, Nvl(C.性别, A.性别) 性别, Nvl(C.年龄, A.年龄) 年龄,B.ID as 科室ID, B.名称 As 科室, C.出院病床 As 当前床号, c.当前病况 as 病情," & _
                "C.入院日期 as 就诊时间, C.险类,c.当前病区ID,c.出院科室ID" & vbNewLine & _
                "From 病人信息 A, 部门表 B, 病案主页 C" & vbNewLine & _
                "Where C.出院科室id = B.Id And A.病人id = C.病人id And A.主页id = C.主页id And C.病人id = [1] And C.主页id = [2]"
        
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Else
        strSQL = "Select a.id,A.姓名,A.性别,A.年龄,a.no,a.门诊号 as 住院号,B.ID as 科室ID, b.名称 as 科室, null as 病情,a.执行时间 as 就诊时间,a.诊室" & _
                " From 病人挂号记录 A,部门表 b " & _
                " Where A.NO=[1] And a.记录性质=1 And a.记录状态=1 And A.病人ID+0=[2] and a.执行部门id=b.id"
        
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号NO, mlng病人ID)
    End If
    
    If rsTmp.RecordCount > 0 Then
        txtInfo(txt姓名).Text = rsTmp!姓名 & ""
        txtInfo(txt性别).Text = rsTmp!性别 & ""
        txtInfo(txt科室).Text = rsTmp!科室 & ""
        txtInfo(txt年龄).Text = rsTmp!年龄 & ""
        txtInfo(txt病情).Text = rsTmp!病情 & ""
        If mint场合 = 0 Then
            txtInfo(txt床号).Text = rsTmp!当前床号 & ""
            mlng病区ID = Val(rsTmp!当前病区ID & "")
        Else
            txtInfo(txt床号).Text = ""
            mlng挂号ID = Val(rsTmp!ID & "")
            mstr诊室 = rsTmp!诊室 & ""
        End If
        If IsDate(rsTmp!就诊时间 & "") Then
             mstr就诊时间 = Format(rsTmp!就诊时间 & "", "YYYY-MM-DD HH:mm")
        End If
        txtInfo(txt住院号).Text = rsTmp!住院号 & ""
        mlng科室ID = Val(rsTmp!科室ID & "")
    End If
        
    If mdat送检时间 <> CDate(0) Then
        txtInfo(txt送检时间).Text = Format(mdat送检时间, "yyyy-MM-dd HH:mm")
        txtInfo(txt送检时间).Tag = txtInfo(txt送检时间).Text
    End If
    
    '读取送检科室
    If mlng送检科室ID <> 0 Then
        strSQL = "Select a.Id, a.名称 From 部门表 A Where ID = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng送检科室ID)
        If rsTmp.RecordCount > 0 Then
            txtInfo(txt送检科室).Text = rsTmp!名称 & ""
            txtInfo(txt送检科室).Tag = txtInfo(txt送检科室).Text
        End If
    End If
    
    If mdat检查时间 <> CDate(0) Then
        txtInfo(txt检查时间).Text = Format(mdat检查时间, "yyyy-MM-dd HH:mm")
        txtInfo(txt检查时间).Tag = txtInfo(txt检查时间).Text
    End If
    
    txtInfo(txt送检医生).Text = mstr送检医生
    txtInfo(txt标本名称).Text = mstr标本名称
    txtInfo(txt反馈结果).Text = mstr反馈结果
    txtInfo(txt疑似疾病).Text = mstr疑似传染病
    txtInfo(txt送检医生).Tag = txtInfo(txt送检医生).Text
    txtInfo(txt标本名称).Tag = txtInfo(txt标本名称).Text
    txtInfo(txt疑似疾病).Tag = txtInfo(txt疑似疾病).Text
    
    If mintType = 1 Or mintType = 2 Then
        '读取登记科室
        If mlng登记科室ID <> 0 Then
            strSQL = "Select a.Id, a.名称 From 部门表 A Where ID = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng登记科室ID)
            If rsTmp.RecordCount > 0 Then
                txtInfo(txt登记科室).Text = rsTmp!名称 & ""
            End If
        End If
    ElseIf mintType = 0 Or mintType = 3 Then
         Call GetRegistInfo
    End If
    
    If mintType = 1 Then
        glngOpenedID = mlngID
        txtInfo(txt确认医师).Text = UserInfo.姓名
        txtInfo(txt处理时间).Text = Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    End If

    If mblnDialog Then
        Call SetCboReportData(mlng病人ID, mstr疑似传染病, mlngID)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSampleList(ByVal intType As Integer)
'功能：获取标本名称
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt标本名称).Tag = txtInfo(txt标本名称).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt标本名称).Text = "" Then '相当于是清除该项目
            txtInfo(txt标本名称).Tag = ""
            mstr标本名称 = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "select A.编码 as ID,A.名称 from 诊疗检验标本 A" & IIf(intType = 0, " where A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2]", "") & " order by A.编码"
        
    strInput = Trim(UCase(txtInfo(txt标本名称).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt标本名称).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "传染病", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt标本名称).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的标本!", vbInformation, gstrSysName)
            txtInfo(txt标本名称).SetFocus
            mstr标本名称 = ""
            gobjComlib.zlControl.TxtSelAll txtInfo(txt标本名称)
        End If
        Exit Sub
    Else
        txtInfo(txt标本名称).Text = rsTmp!名称 & ""
        txtInfo(txt标本名称).Tag = rsTmp!名称 & ""
        txtInfo(txt标本名称).SetFocus
        mstr标本名称 = rsTmp!名称 & ""
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetInspectDept(ByVal intType As Integer)
'功能：获取送检科室
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt送检科室).Tag = txtInfo(txt送检科室).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt送检科室).Text = "" Then '相当于是清除该项目
            txtInfo(txt送检科室).Tag = ""
            mlng送检科室ID = 0
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    If mint场合 = 0 Then
        strTemp = " and B.服务对象 in (2,3) "
    ElseIf mint场合 = 1 Then
        strTemp = " and B.服务对象 in (1,3) "
    End If
    
    strInput = Trim(UCase(txtInfo(txt送检科室).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt送检科室).hwnd)
        
    If mstr送检医生 = "" Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 as 科室,A.简码 From 部门表 A,部门性质说明 B " & _
                " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And a.Id = b.部门id" & _
                IIf(intType = 0, " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])", "") & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.工作性质= '临床' " & strTemp & "  Order by A.编码"
    Else
        strSQL = "Select Distinct d.Id, d.编码, d.名称 As 科室, d.简码 " & vbNewLine & _
                "From 人员表 A, 部门性质说明 B,部门人员 C, 部门表 D " & vbNewLine & _
                "Where (D.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL) " & vbNewLine & _
                 IIf(intType = 0, " And (D.编码 Like [1] Or D.名称 Like [2] Or D.简码 Like [2])", "") & vbNewLine & _
                "and a.Id = c.人员id And d.Id = B.部门id And c.部门id = d.Id  And a.姓名 = [3]" & vbNewLine & _
                " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) And B.工作性质= '临床' " & strTemp & "  Order by D.编码"
    End If
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "传染病", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtInfo(txt送检科室).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", mstr送检医生)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的科室!", vbInformation, gstrSysName)
            mlng送检科室ID = 0
            txtInfo(txt送检科室).SetFocus
            gobjComlib.zlControl.TxtSelAll txtInfo(txt送检科室)
        End If
        Exit Sub
    Else
        mlng送检科室ID = Val(rsTmp!ID)
        txtInfo(txt送检科室).Text = rsTmp!科室 & ""
        txtInfo(txt送检科室).Tag = rsTmp!科室 & ""
        txtInfo(txt送检科室).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetInspectDoctor(ByVal intType As Integer)
'功能：获取送检医生,人员性质为 医生 ，部门性质为 临床，服务于住院或者门诊 下面的人员
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim lngDeptId As Long
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt送检医生).Tag = txtInfo(txt送检医生).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt送检医生).Text = "" Then '相当于是清除该项目
            txtInfo(txt送检医生).Tag = ""
             mstr送检医生 = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "Select Distinct a.Id, a.编号, a.姓名, a.简码, d.名称 As 部门 ,d.ID as 部门ID" & vbNewLine & _
            "From 人员表 A, 人员性质说明 B, 部门人员 C, 部门表 D, 部门性质说明 E " & vbNewLine & _
            "Where a.Id = b.人员id And b.人员性质 = '医生' And a.Id = c.人员id  " & vbNewLine & _
             IIf(mlng送检科室ID = 0, "And c.缺省 = 1 ", "And c.部门id = [1] ") & vbNewLine & _
             IIf(intType = 0, " And (A.编号 Like [2] Or A.姓名 Like [3] Or A.简码 Like [3]) ", "") & vbNewLine & _
            "And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And " & vbNewLine & _
            "(d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And c.部门id = d.Id And d.Id = e.部门id  " & vbNewLine & _
            IIf(mint场合 = 0, "and e.服务对象 In (2, 3) ", "and e.服务对象 In (1, 3) ") & vbNewLine & _
            "And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)  Order By a.编号"
    
    strInput = Trim(UCase(txtInfo(txt送检医生).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt送检医生).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "传染病", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtInfo(txt送检医生).Height, blnCancel, False, True, mlng送检科室ID, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("没有找到匹配的医生!", vbInformation, gstrSysName)
             mstr送检医生 = ""
            txtInfo(txt送检医生).SetFocus
            gobjComlib.zlControl.TxtSelAll txtInfo(txt送检医生)
        End If
        Exit Sub
    Else
        txtInfo(txt送检医生).Text = rsTmp!姓名 & ""
        txtInfo(txt送检医生).Tag = rsTmp!姓名 & ""
        mstr送检医生 = rsTmp!姓名 & ""
        If (mlng送检科室ID = 0) Then
            txtInfo(txt送检科室).Text = rsTmp!部门 & ""
            txtInfo(txt送检科室).Tag = rsTmp!部门 & ""
            mlng送检科室ID = Val(rsTmp!部门ID & "")
        End If
        
        txtInfo(txt送检医生).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetFormState(ByVal intType As Integer)
    Dim objControl As Object
    
    If intType = 0 Or intType = 3 Then
        SetControlEnabled txtInfo(txt姓名), False
        SetControlEnabled txtInfo(txt性别), False
        SetControlEnabled txtInfo(txt年龄), False
        SetControlEnabled txtInfo(txt病情), False
        SetControlEnabled txtInfo(txt床号), False
        SetControlEnabled txtInfo(txt住院号), False
        SetControlEnabled txtInfo(txt科室), False

        SetControlEnabled txtInfo(txt登记人), False, False
        SetControlEnabled txtInfo(txt登记科室), False, False
        SetControlEnabled txtInfo(txt登记时间), False, False
        If intType = 0 Then
            SetControlEnabled txtInfo(txt送检科室), mlng送检科室ID = 0
            SetControlEnabled txtInfo(txt送检医生), mstr送检医生 = ""
            SetControlEnabled txtInfo(txt标本名称), mstr标本名称 = ""
            SetControlEnabled txtInfo(txt送检时间), mdat送检时间 = 0
            SetControlEnabled txtInfo(txt检查时间), mdat检查时间 = 0
            SetControlEnabled txtInfo(txt疑似疾病), mstr疑似传染病 = ""
            
            SetControlEnabled cmdInfo(cmd送检科室), mlng送检科室ID = 0
            SetControlEnabled cmdInfo(cmd送检医生), mstr送检医生 = ""
            SetControlEnabled cmdInfo(cmd标本名称), mstr标本名称 = ""
            SetControlEnabled cmdInfo(cmd送检时间), mdat送检时间 = 0
            SetControlEnabled cmdInfo(cmd检查时间), mdat检查时间 = 0
            SetControlEnabled cmdInfo(cmd疑似疾病), mstr疑似传染病 = ""
        End If
        fraIdea.Visible = False
        Me.Height = 7800
    ElseIf intType = 1 Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
        SetControlEnabled txtInfo(txt处理情况), True
        SetControlEnabled cboReport, True
    ElseIf intType = 2 Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
    End If
    
    lblInfo(txt住院号).Caption = IIf(mint场合 = 0, "住 院 号", "门 诊 号")
    lblInfo(txt床号).Visible = (mint场合 = 0)
    txtInfo(txt床号).Visible = (mint场合 = 0)
    Line1(txt床号).Visible = (mint场合 = 0)
End Sub

Private Sub SetControlEnabled(objControl As Object, ByVal blnEnabled As Boolean, Optional blnColor As Boolean = True)
'功能：设置控件的可用性
    Select Case TypeName(objControl)
        Case "TextBox"
            objControl.Locked = Not blnEnabled
            objControl.TabStop = blnEnabled
            If blnColor Then objControl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
        Case "CommandButton", "ComboBox"
            objControl.Enabled = blnEnabled
    End Select
End Sub

Private Function SeekNextCtl() As Boolean
'功能：定位到下一个焦点的控件上
    Call gobjComlib.zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Function SendDiseaseRecord() As Boolean
    '发送
    Dim strSQL As String
    Dim str主页ID As String
    Dim str挂号No As String
    Dim str送检时间 As String
    Dim str送检医生 As String
    Dim str标本名称 As String
    Dim str反馈结果 As String
    Dim str疑似传染病 As String
    Dim str检查时间 As String
    Dim str登记时间 As String
    Dim str登记人 As String
 On Error GoTo errH
    If mint场合 = 0 Then
        str挂号No = "NULL"
        str主页ID = CStr(mlng主页ID)
    ElseIf mint场合 = 1 Then
        str主页ID = "NULL"
        str挂号No = "'" & mstr挂号NO & "'"
    End If
    
    If mdat送检时间 = CDate(0) Then
        str送检时间 = "NULL"
    Else
        str送检时间 = "to_date('" & Format(mdat送检时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    If mdat检查时间 = CDate(0) Then
        str检查时间 = "NULL"
    Else
        str检查时间 = "to_date('" & Format(mdat检查时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    If mdat登记时间 = CDate(0) Then
        str登记时间 = "NULL"
    Else
        str登记时间 = "to_date('" & Format(mdat登记时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    str送检医生 = "'" & mstr送检医生 & "'"
    str标本名称 = "'" & mstr标本名称 & "'"
    str反馈结果 = "'" & mstr反馈结果 & "'"
    str疑似传染病 = "'" & mstr疑似传染病 & "'"
    str登记人 = "'" & mstr登记人 & "'"
    If mintType = 0 Then
        mlngID = gobjComlib.zlDatabase.GetNextId("疾病阳性记录")       '获取ID
        strSQL = "Zl_疾病阳性检测记录_Insert(" & mlngID & "," & mlng病人ID & "," & str主页ID & "," & str挂号No & "," & IIf(mlng医嘱ID = 0, "NULL", mlng医嘱ID) & "," _
                & str送检时间 & "," & IIf(mlng送检科室ID = 0, "NULL", mlng送检科室ID) & "," & str送检医生 & "," & str标本名称 & "," & str反馈结果 & "," _
                & str疑似传染病 & "," & str检查时间 & "," & str登记时间 & "," & str登记人 & "," & IIf(mlng登记科室ID = 0, "NULL", mlng登记科室ID) & "," & 1 & ")"
    ElseIf mintType = 3 Then
        strSQL = "Zl_疾病阳性检测记录_Update(" & 4 & "," & mlngID & ",NULL,NULL,NULL,NULL,NULL," & str送检时间 & "," & IIf(mlng送检科室ID = 0, "NULL", mlng送检科室ID) & "," & str送检医生 & "," & str标本名称 & "," & str反馈结果 & "," _
                & str疑似传染病 & "," & str检查时间 & "," & str登记时间 & "," & str登记人 & "," & IIf(mlng登记科室ID = 0, "NULL", mlng登记科室ID) & ")"
    End If
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    SendDiseaseRecord = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveDisProcessData(ByVal intType As Integer) As Boolean
'功能：保存数据
'参数：intType   1-最开始进来时改变状态，时间，医生；2-确定；3-非传染病；4-转科；5-关闭；6-打印前保存
    Dim strSQL As String
    Dim str处理时间 As String
    Dim str处理医生 As String
    Dim str处理情况 As String, strTmp As String
    Dim lngReportID As Long
    Dim intDisState As Integer      '疾病阳性记录.记录状态,反馈单的当前状态：1-待医生确认，2-医生已处理(点击确认),3-非传染病，4-转科待处理
    
    On Error GoTo errH
    
    If mintType <> 1 Then Exit Function
    
    strTmp = Trim(txtInfo(txt处理情况).Text)
    str处理时间 = "to_date('" & Format(txtInfo(txt处理时间).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str处理医生 = "'" & txtInfo(txt确认医师).Text & "'"
    str处理情况 = "'" & strTmp & "'"


    If cboReport.ListCount > 1 Then
        lngReportID = cboReport.ItemData(cboReport.ListIndex)
    End If

    If (mIntState = 1 Or mIntState = 4) And (intType = 1 Or intType = 2) Or (mIntState = 3 And intType = 2) Then
        strSQL = "Zl_疾病阳性检测记录_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & "2" & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mIntState = 2
    ElseIf intType = 3 Or intType = 4 Then
        If strTmp = "" Then
            MsgBox "处理情况说明不能为空，请先填写处理情况说明！", vbInformation, gstrSysName
            Exit Function
        End If
        If intType = 3 Then lngReportID = 0
        strSQL = "Zl_疾病阳性检测记录_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(intType) & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mIntState = intType
    ElseIf intType = 5 Then
        If strTmp = "" Then
            MsgBox "处理情况说明不能为空，请先填写处理情况说明！", vbInformation, gstrSysName
            Exit Function
        End If
        strSQL = "Zl_疾病阳性检测记录_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(mIntState) & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    ElseIf intType = 6 Then
        strSQL = "Zl_疾病阳性检测记录_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(mIntState) & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    If (intType = 2 Or intType = 3) And mstr类别 = "E" Then
        Call InitObjLis(1278)
        If Not gobjLIS Is Nothing Then
            strTmp = ""
            Call gobjLIS.WriteInLisNotify(2, CStr(mlng医嘱ID), , strTmp)
            If strTmp <> "" Then MsgBox "zl9LisInsideComm部件(WriteInLisNotify)方法错误：" & vbCrLf & strTmp, vbInformation, gstrSysName
        End If
    End If
    SaveDisProcessData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strMsg As String
    If mintType = 0 Then
        '必须录入疑似疾病
        If txtInfo(txt疑似疾病).Text = "" Then
            MsgBox "没有确定疑似疾病。", vbInformation, gstrSysName
            If txtInfo(txt疑似疾病).Enabled Then txtInfo(txt疑似疾病).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt反馈结果).Text = "" Then
            MsgBox "没有填写反馈结果。", vbInformation, gstrSysName
            If txtInfo(txt反馈结果).Enabled Then txtInfo(txt反馈结果).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt检查时间).Text = "" Then
            MsgBox "没有填写检查时间。", vbInformation, gstrSysName
            Exit Function
        Else
            If Not Check时间("检查时间", txtInfo(txt检查时间).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt检查时间).Enabled Then txtInfo(txt检查时间).SetFocus
                Exit Function
            End If
        End If
        
        If txtInfo(txt送检时间).Text <> "" Then
            If Not Check时间("送检时间", txtInfo(txt送检时间).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt送检时间).Enabled Then txtInfo(txt送检时间).SetFocus
                Exit Function
            End If
        End If
    End If
    CheckData = True
End Function

Private Sub Form_Resize()
On Error Resume Next
    If mblnDialog Then
        frmMain.Top = 800
        frmMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (frmMain.Width / 2)
    Else
        frmMain.Top = mlngTop
        frmMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (frmMain.Width / 2)
    
        If mblnSbSisible Then
            If Me.ScaleWidth < frmMain.Width Then
                hsbReport.Visible = True
            Else
                hsbReport.Visible = False
            End If
        
            If Me.ScaleHeight < frmMain.Height Then
                vsbReport.Visible = True
            Else
                vsbReport.Visible = False
            End If
            vsbReport.Top = Me.ScaleTop
            vsbReport.Left = Me.ScaleLeft + Me.ScaleWidth - vsbReport.Width
            vsbReport.Height = Me.ScaleHeight - IIf(hsbReport.Visible = True, hsbReport.Height, 0)
            vsbReport.LargeChange = 100 / ((frmMain.Height + 800) / Me.ScaleHeight)
            vsbReport.SmallChange = vsbReport.LargeChange
            
            hsbReport.Top = vsbReport.Top + vsbReport.Height
            hsbReport.Left = Me.ScaleLeft
            hsbReport.Width = Me.ScaleLeft + Me.ScaleWidth
            hsbReport.LargeChange = 100 / (frmMain.Width / Me.ScaleWidth)
            hsbReport.SmallChange = hsbReport.LargeChange
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mintType = 1 And Not mblnOk Then
        If Not SaveDisProcessData(5) Then
            Cancel = True
            Exit Sub
        End If
    End If
    If mblnDiagnose Then
        RaiseEvent Closed(0, "")
    End If
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    If Not mclsMipModule Is Nothing Then Set mclsMipModule = Nothing
    If mintType = 1 Then
        glngOpenedID = 0
    End If
End Sub

Private Sub mclsDiagEdit_Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
    Dim clsDisease As New cDockDisease
    Dim strName As String, strReason As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim frmDisStation As frmDiseaseStation
    Dim blnNotView As Boolean
    
    On Error GoTo errH
     If Not blnEditCancel Then
        strName = txtInfo(txt姓名).Text
        mblnDiagnose = True
        If str诊断ID = "" And str疾病ID = "" Then Exit Sub
        
        If InStr(";" & gobjComlib.GetPrivFunc(glngSys, 1249) & ";", ";病历书写;") <= 0 Then
            Exit Sub
        End If
        
        Set rsTemp = clsDisease.SatisfyEditDiseaseDoc(mlng病人ID, mlng主页ID, IIf(mint场合 = 0, 2, 1), mlng科室ID, str疾病ID, str诊断ID)
        
        If rsTemp Is Nothing Then
            Exit Sub
        ElseIf rsTemp.RecordCount = 0 Then
            Exit Sub          '不符合疾病报告前提，退出
        End If
        If cboReport.ListCount > 1 Then
            If cboReport.ListIndex > 0 Then
                If MsgBox("该反馈单已经关联了一份疾病报告单，是否填写新的报告单？！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                 If MsgBox("该病人已经填写过" & "“" & mstr疑似传染病 & "”疾病的报告单，是否关联？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    cboReport.ListIndex = 1
                    Call SaveDisProcessData(2)
                    Exit Sub
                End If
            End If
        End If
        Set frmDisStation = New frmDiseaseStation

        '没有查找到一年之内的重复的报告卡
        If Not frmDisStation.ShowDiseaseStation(Me, mlng病人ID, IIf(mint场合 = 0, mlng主页ID, mlng挂号ID), IIf(mint场合 = 0, 2, 1), _
                                    mlng科室ID, str疾病ID, str诊断ID, blnNotView) Then
            Call clsDisease.EditDiseaseReport(Me, rsTemp, mlng病人ID, IIf(mint场合 = 0, mlng主页ID, mlng挂号ID), IIf(mint场合 = 0, 2, 1), mlng科室ID, str疾病ID, str诊断ID, strReason)
            If strReason <> "" Then txtInfo(txt处理情况).Text = strReason
        ElseIf blnNotView Then
            Call clsDisease.EditDiseaseReport(Me, rsTemp, mlng病人ID, IIf(mint场合 = 0, mlng主页ID, mlng挂号ID), IIf(mint场合 = 0, 2, 1), mlng科室ID, str疾病ID, str诊断ID, strReason)
            If strReason <> "" Then txtInfo(txt处理情况).Text = strReason
        End If
        Call SetCboReportData(mlng病人ID, mstr疑似传染病, mlngID)
        
        If Not frmDisStation Is Nothing Then
            Unload frmDisStation
            Set frmDisStation = Nothing
        End If
    End If
    Set clsDisease = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCboReportData(ByVal lng病人ID As Long, ByVal str疑似传染病 As String, ByVal lngID As Long) As Boolean
'功能：查询阳性反馈单关联的疾病报告
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
On Error GoTo errH
    cboReport.Clear
    strSQL = "Select rowNum as NO,A.ID,B.ID as 反馈单ID,A.创建时间,A.病历名称,B.传染病名称  from 电子病历记录 A, 疾病阳性记录 B where A.ID = B.文件ID and A.病人ID = B.病人ID and B.病人ID = [1] and B.传染病名称 = [2]"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性反馈单关联的疾病报告", lng病人ID, str疑似传染病)
    
    If rsTemp.RecordCount > 0 Then
        cboReport.AddItem ""
        cboReport.ItemData(cboReport.NewIndex) = 0
        cboReport.ListIndex = 0
        For i = 1 To rsTemp.RecordCount
            cboReport.AddItem rsTemp!NO & "-" & rsTemp!病历名称 & "(" & rsTemp!传染病名称 & ")"
            cboReport.ItemData(cboReport.NewIndex) = rsTemp!ID
            If lngID = rsTemp!反馈单ID Then
                cboReport.ListIndex = i
            End If
            rsTemp.MoveNext
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtInfo_Change(Index As Integer)
    Select Case Index
        Case txt反馈结果
            mstr反馈结果 = txtInfo(txt反馈结果).Text
    End Select
End Sub

Private Function Check时间(ByVal strTimeType As String, ByVal str时间 As String, Optional ByRef strMsg As String) As Boolean
'功能：检查输入的时间是否合法
    Dim strInDate As String
    Dim datCurrent As Date
    
    datCurrent = gobjComlib.zlDatabase.Currentdate
    strInDate = mstr就诊时间
    If Not IsDate(str时间) Then
        strMsg = "输入的" & strTimeType & "无效。"
        Exit Function
    End If

    If mint场合 = 0 Then
        If Format(str时间, "yyyy-MM-dd HH:mm") < Format(mstr就诊时间, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "不能小于病人的入院时间 " & strInDate & " 。"
            Exit Function
        End If
    Else
        If Format(str时间, "yyyy-MM-dd HH:mm") < Format(mstr就诊时间, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "不能小于病人的就诊时间 " & strInDate & " 。"
            Exit Function
        End If
    End If
    
    If Format(str时间, "yyyy-MM-dd HH:mm") > Format(datCurrent, "yyyy-MM-dd HH:mm") Then
         strMsg = strTimeType & "不能大于当前时间 " & Format(datCurrent, "yyyy-MM-dd HH:mm") & " 。"
         Exit Function
     End If
        
    Check时间 = True
End Function

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'按键事件，糊模查找
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Select Case Index
            Case txt送检科室
                Call GetInspectDept(0)
            Case txt送检医生
                Call GetInspectDoctor(0)
            Case txt疑似疾病
                Call GetDiseaseList(0)
            Case txt标本名称
                Call GetSampleList(0)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case txt送检科室
            Call GetInspectDept(1)
        Case txt送检医生
            Call GetInspectDoctor(1)
        Case txt疑似疾病
            Call GetDiseaseList(1)
        Case txt标本名称
            Call GetSampleList(1)
        End Select
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strMsg As String
    
    If mintType <> 0 Then
        Exit Sub
    End If
    Select Case Index
        Case txt送检科室, txt送检医生, txt标本名称, txt疑似疾病
            If txtInfo(Index).Text <> txtInfo(Index).Tag Then
                If txtInfo(Index).Text = "" Then
                    txtInfo(Index).Tag = ""
                    If Index = txt送检科室 Then
                        mlng送检科室ID = 0
                    ElseIf Index = txt送检医生 Then
                        mstr送检医生 = ""
                    ElseIf Index = txt标本名称 Then
                        mstr标本名称 = ""
                    ElseIf Index = txt疑似疾病 Then
                        mstr疑似传染病 = ""
                    End If
                Else
                    txtInfo(Index).Text = txtInfo(Index).Tag
                    If txtInfo(Index).Enabled Then
                        txtInfo(Index).SetFocus
                        gobjComlib.zlControl.TxtSelAll txtInfo(Index)
                    End If
                End If
            End If
            
        Case txt送检时间, txt检查时间
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = txtInfo(Index).Tag
            Else
                txtInfo(Index).Tag = txtInfo(Index).Text
                If Index = txt送检时间 Then
                    mdat送检时间 = Format(txtInfo(txt送检时间).Text, "yyyy-MM-dd HH:mm")
                ElseIf Index = txt检查时间 Then
                    mdat检查时间 = Format(txtInfo(txt检查时间).Text, "yyyy-MM-dd HH:mm")
                End If
            End If
    End Select
End Sub

Private Sub SendMsg()
'功能：发送 传染病阳性结果 消息
    Dim strXML As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String

    On Error GoTo errH
    strXML = "<patient_info><patient_id>" & mlng病人ID & "</patient_id><patient_name>" & txtInfo(txt姓名).Text & "</patient_name>"
    If mint场合 = 0 Then
        strXML = strXML & "<in_number>" & txtInfo(txt住院号).Text & "</in_number>"
    Else
        strXML = strXML & "<out_number>" & txtInfo(txt住院号).Text & "</out_number>"
    End If
    strXML = strXML & "</patient_info><patient_clinic><patient_source>" & IIf(mint场合 = 0, 2, 1) & "</patient_source>"
    strXML = strXML & "<clinic_id>" & IIf(mint场合 = 0, mlng主页ID, mlng挂号ID) & "</clinic_id>"
    If mlng病区ID <> 0 Then
        strXML = strXML & "<clinic_area_id>" & mlng病区ID & "</clinic_area_id>"
        strTmp = ""
        strTmp = gobjComlib.Sys.RowValue("部门表", mlng病区ID, "名称")
        If strTmp <> "" Then
            strXML = strXML & "<clinic_area_title>" & strTmp & "</clinic_area_title>"
        End If
    End If
    strXML = strXML & "<clinic_dept_id>" & mlng科室ID & "</clinic_dept_id>"
    strTmp = "" & gobjComlib.Sys.RowValue("部门表", mlng科室ID, "名称")
    strXML = strXML & "<clinic_dept_title>" & strTmp & "</clinic_dept_title>"
    strXML = strXML & "<clinic_room>" & mstr诊室 & "</clinic_room>"
    If "" <> txtInfo(txt床号).Text Then
        strXML = strXML & "<clinic_bed>" & strTmp & "</clinic_bed>"
    End If
    strXML = strXML & "</patient_clinic><positive_info><info_id>" & mlngID & "</info_id>"
    strXML = strXML & "<sample_name>" & mstr标本名称 & "</sample_name>"
    strXML = strXML & "<disease_name>" & mstr疑似传染病 & "</disease_name>"
    strXML = strXML & "<create_time>" & Format(mdat登记时间, "yyyy-MM-dd HH:mm:ss") & "</create_time>"
    strXML = strXML & "<create_doctor>" & mstr登记人 & "</create_doctor>"
    strXML = strXML & "<create_dept_id>" & IIf(mlng登记科室ID = 0, "NULL", mlng登记科室ID) & "</create_dept_id>"

    strTmp = ""
    strTmp = "" & gobjComlib.Sys.RowValue("部门表", IIf(mlng登记科室ID = 0, "NULL", mlng登记科室ID), "名称")
    If strTmp <> "" Then
        strXML = strXML & "<create_dept>" & strTmp & "</create_dept>"
    End If

    strXML = strXML & "</positive_info>"

    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then Call mclsMipModule.CommitMessage("ZLHIS_CIS_032", strXML)
    End If

    Call gobjComlib.zlDatabase.SendMsg("ZLHIS_CIS_032", strXML)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFrmInset(ByVal blnSbSisible As Boolean)
    mblnSbSisible = blnSbSisible
    Me.Appearance = 1
    Me.BackColor = &HC0C0C0
    Me.BorderStyle = 0
    Me.Caption = Me.Caption
    mblnDialog = False
End Sub

Public Sub zlRefresh(ByVal lngID As Long)
    mintType = 2
    mlngID = lngID
    Call LoadPatiInfo
    Call SetFormState(2)
End Sub

Private Sub hsbReport_Change()
    frmMain.Left = -((frmMain.Width - Me.Width) * (hsbReport.Value / 100))
End Sub

Public Sub SetReportTop(ByVal lngTop As Long)
     frmMain.Top = lngTop
     mlngTop = lngTop
End Sub

Private Sub vsbReport_Change()
    frmMain.Top = 200 - ((frmMain.Height + 800 - Me.Height) * (vsbReport.Value / 100))
End Sub

Private Sub ViewEPRReport()
'功能：查阅报告
    Dim lng报告ID As Long
    Dim str检查报告ID As String
    Dim objPublicPACS As Object

    '先判断是否可以继续操作
    If mlng医嘱ID = 0 Then
        MsgBox "该反馈单对应的医嘱为空，无法查看检验检查报告！", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case CheckEPRReport(mlng医嘱ID, lng报告ID, str检查报告ID)
    Case 0
        MsgBox "该医嘱的报告没有书写！", vbInformation, gstrSysName
        Exit Sub
    Case 2
        If InStr(gobjComlib.GetPrivFunc(glngSys, 1253), "查阅未完成报告") > 0 Then
            MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
        Else
            MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select

    '执行操作
    '新版PACS报告，直接强制使用新版PACS报告编辑器
    If str检查报告ID <> "" Then
        Call CreateObjectPacs(objPublicPACS)
        Call objPublicPACS.zlDocShowReport(mlng医嘱ID, , False, Me, True)
    Else
        Call gObjRichEPR.ViewDocument(Me, lng报告ID)
        '查阅报告
    End If
    If objPublicPACS Is Nothing Then Set objPublicPACS = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.用户名)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS公共部件未创建成功！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function


Private Function CheckEPRReport(ByVal lng医嘱id As Long, ByRef lng报告ID As Long, ByRef str检查报告ID As String) As Integer
'功能：检查对应项目的报告填写情况
'参数：lng医嘱ID=可见行的医嘱ID
'      lng报告ID=可以传入，主要用于返回报告病历ID
'      int执行状态=用于检验完成时，传入综合的执行状态
'返回：0-报告还没有填写
'      1-报告已填写完成(已签名,包括修订后签名,或已执行完成)
'      2-报告未填写完成(未签名,或修订后未签名,且未执行完成)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '检查报告是否已书写
    strSQL = "Select 病历ID,检查报告ID || ''  as 检查报告ID From 病人医嘱报告 Where 医嘱ID=[1]"
    
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱id)
    If Not rsTmp.EOF Then lng报告ID = Val(rsTmp!病历id & ""): str检查报告ID = rsTmp!检查报告ID & ""
    If lng报告ID = 0 And str检查报告ID = "" Then
        CheckEPRReport = 0: Exit Function
    End If
    
    '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
    '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
    strSQL = _
        " Select 2 as 排序,医嘱ID,执行过程,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
        " Union ALL" & _
        " Select 排序,医嘱ID,执行过程, 执行状态,发送时间" & _
        " From (" & _
            " Select 1 as 排序,B.医嘱ID,B.执行过程,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
            " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
            " Order by A.序号" & _
        " ) Where Rownum=1" & _
        " Order by 排序,发送时间 Desc"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱id, 0)
    If NVL(rsTmp!执行过程, 0) >= 5 Or NVL(rsTmp!执行状态, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub
