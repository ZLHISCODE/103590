VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCriticalEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "危急值登记单"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   Icon            =   "frmCriticalEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   9570
      TabIndex        =   40
      Top             =   8880
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   117702657
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7665
      Left            =   450
      ScaleHeight     =   7665
      ScaleWidth      =   10500
      TabIndex        =   0
      Top             =   660
      Width           =   10500
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "235"
         Top             =   1785
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
         Index           =   0
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "张三"
         Top             =   1230
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "男"
         Top             =   1230
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "内科"
         Top             =   1785
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "6"
         Top             =   1785
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "28岁"
         Top             =   1230
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
         Index           =   6
         Left            =   1260
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "2017-10-11 12:00"
         Top             =   4080
         Width           =   1725
      End
      Begin VB.PictureBox picCL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   -30
         ScaleHeight     =   2895
         ScaleWidth      =   10455
         TabIndex        =   3
         Top             =   4530
         Width           =   10455
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   270
            Index           =   13
            Left            =   9705
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2250
            Width           =   270
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   270
            Index           =   12
            Left            =   6570
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2265
            Width           =   270
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
            Left            =   1260
            TabIndex        =   8
            Text            =   "2013-06-20 18:00"
            Top             =   2280
            Width           =   1695
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
            Index           =   10
            Left            =   1275
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   135
            Width           =   8685
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Index           =   0
            Left            =   1200
            ScaleHeight     =   390
            ScaleWidth      =   1965
            TabIndex        =   4
            Top             =   1665
            Width           =   1965
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "是"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   75
               TabIndex        =   6
               Top             =   60
               Value           =   -1  'True
               Width           =   600
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "否"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   825
               TabIndex        =   5
               Top             =   75
               Width           =   600
            End
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            Left            =   8775
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "张三"
            Top             =   2280
            Width           =   1125
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            Left            =   5505
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "内科"
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   -105
            X2              =   11055
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "确认时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   11
            Left            =   315
            TabIndex        =   17
            Top             =   2295
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   1260
            X2              =   2910
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   5490
            X2              =   6795
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "确认科室"
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
            Left            =   4560
            TabIndex        =   16
            Top             =   2295
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   8700
            X2              =   9960
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "确 认 人"
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
            Left            =   7830
            TabIndex        =   15
            Top             =   2295
            Width           =   840
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   10
            Left            =   315
            TabIndex        =   14
            Top             =   150
            Width           =   900
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "是否是危值"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   14
            Left            =   105
            TabIndex        =   13
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Image imgDate 
            Height          =   240
            Index           =   11
            Left            =   3015
            Picture         =   "frmCriticalEdit.frx":6852
            Top             =   2280
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   270
         Index           =   7
         Left            =   6480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4050
         Width           =   270
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   270
         Index           =   8
         Left            =   9660
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4050
         Width           =   270
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "B超室"
         Top             =   4080
         Width           =   1185
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "张永康"
         Top             =   4080
         Width           =   1230
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
         Index           =   9
         Left            =   1290
         MaxLength       =   4000
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2385
         Width           =   8685
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "危急值登记单"
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
         Left            =   4065
         TabIndex        =   39
         Top             =   120
         Width           =   2250
      End
      Begin VB.Line linHead 
         X1              =   4065
         X2              =   6315
         Y1              =   495
         Y2              =   495
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
         Index           =   1
         Left            =   4560
         TabIndex        =   38
         Top             =   1230
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
         Index           =   2
         Left            =   7785
         TabIndex        =   37
         Top             =   1230
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
         Index           =   4
         Left            =   4560
         TabIndex        =   36
         Top             =   1785
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
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   1230
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
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   1785
         Width           =   840
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
         Left            =   7785
         TabIndex        =   33
         Top             =   1785
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   4
         X1              =   5505
         X2              =   6735
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line linL 
         Index           =   0
         X1              =   1275
         X2              =   2505
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   3
         X1              =   1275
         X2              =   2505
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line linL 
         Index           =   1
         X1              =   5505
         X2              =   6735
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   2
         X1              =   8700
         X2              =   9930
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   5
         X1              =   8700
         X2              =   9930
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报告时间"
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
         Left            =   360
         TabIndex        =   32
         Top             =   4065
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   6
         X1              =   1275
         X2              =   2940
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "危急值描述"
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
         Index           =   9
         Left            =   105
         TabIndex        =   31
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报 告 人"
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
         Left            =   7785
         TabIndex        =   30
         Top             =   4065
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报告科室"
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
         Left            =   4560
         TabIndex        =   29
         Top             =   4065
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   8
         X1              =   8700
         X2              =   9930
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line linL 
         Index           =   7
         X1              =   5505
         X2              =   6735
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "检查类"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7785
         TabIndex        =   28
         Top             =   510
         Width           =   900
      End
      Begin VB.Image imgDate 
         Height          =   240
         Index           =   6
         Left            =   2985
         Picture         =   "frmCriticalEdit.frx":D0A4
         Top             =   4080
         Width           =   240
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCriticalEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CtlID
    txt姓名 = 0
    txt性别 = 1
    txt年龄 = 2
    txt科室 = 3
    txt住院号 = 4
    txt床号 = 5
    txt报告时间 = 6
    txt报告科室 = 7
    txt报告人 = 8
     
    txt危急值描述 = 9
    txt处理情况 = 10
    txt确认时间 = 11
    txt确认科室 = 12
    txt确认人 = 13
End Enum

Private mclsMipModule As zl9ComLib.clsMipModule '消息对象
Private mblnModal As Boolean '显示方式，模态，非模态
Private mfrmParent As Object '父窗口对象
Private mint调用类型 As Integer  '1-门诊,2-住院,3-其它来源病人
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr挂号单 As String
Private mlng医嘱ID As Long
Private mintType As Integer '0-新增，1-修改，2-查看，3-处理
Private mstr就诊时间 As String
Private mlng危急值ID As Long
Private mrs危急值 As ADODB.Recordset
Private mlng标本ID As Long
Private mstr危急值描述 As String
Private mdat报告时间 As Date
Private mlng报告科室ID As Long
Private mstr报告人 As String
Private mstr处理情况 As String
Private mdat确认时间 As Date
Private mstr确认人 As String
Private mlng确认科室ID As Long
Private mobjReport As Object
Private mlng婴儿 As Long
Private mblnOK As Boolean
 
Private mblnChange As Boolean

Public Function ShowMe(frmParent As Object, ByVal blnModal As Boolean, ByVal intType As Integer, ByVal int调用类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal lng婴儿 As Long, ByRef lng危急值ID As Long, ByVal lng医嘱ID As Long, _
    Optional ByVal lng标本id As Long, Optional ByVal str危急值描述 As String, Optional ByVal dat报告时间 As Date, Optional ByVal lng报告科室ID As Long, Optional ByVal str报告人 As String, Optional ByRef objMip As Object) As Boolean
'功能：显示窗体，新增一条危急值
'返回：true 提交了数据， false 未保存
'参数：frmParent 父窗体
'      blnModal 显示方式，1-模态，0-非模态
'      intType  0-新增，1-修改，2-查看，3-处理
'      int调用类型 1-门诊病人,2-住院病人,3-外来病人
'      lng病人ID,lng主页ID,str挂号单,病人相关
'      lng危急值ID 当前记录ID,可返回
'      lng医嘱ID 对应的医嘱项目
'      lng标本ID LIS调用是传入

'      str危急值描述 新增时界面缺省值
'      dat报告时间   新增时界面缺省值
'      lng报告科室ID 新增时界面缺省值
'      str报告人     新增时界面缺省值

'      objMip 用于发送消息的对象 zl9ComLib.clsMipModule

    Set mfrmParent = frmParent
    mblnModal = blnModal
    mintType = intType
    mint调用类型 = int调用类型
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr挂号单 = str挂号单
    
    If mlng主页ID = 0 And mstr挂号单 = "" Then
        mint调用类型 = 3
    End If
    
    mlng婴儿 = lng婴儿
    mlng危急值ID = lng危急值ID
    mlng医嘱ID = lng医嘱ID
    mlng标本ID = lng标本id
    mstr危急值描述 = str危急值描述
    mdat报告时间 = dat报告时间
    mlng报告科室ID = lng报告科室ID
    mstr报告人 = str报告人
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    Me.Show IIF(blnModal, 1, 0), frmParent
    
    lng危急值ID = mlng危急值ID
    
    ShowMe = mblnOK
    
End Function

Public Function ShowApp(frmParent As Object, ByVal lng危急值ID As Long)
'功能：查看危急值记录
    Set mfrmParent = frmParent
    mlng危急值ID = lng危急值ID
    mintType = 2
    Call InitBaseBy记录ID(lng危急值ID)
    Me.Show 1, frmParent
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet '打印设置
             Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_20", Me)
        Case conMenu_File_Preview '预览
            Call PrintApply(1)
        Case conMenu_File_Print '打印
            Call PrintApply(2)
        Case conMenu_Edit_Save
            If CheckData() Then
                Call SaveData
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save
            Control.Enabled = mblnChange
            If mintType = 2 Then
                blnVisible = False
            End If
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Select Case Index
    Case txt报告科室
    Case txt报告人
        Call GetItem报告人(1)
    Case txt确认科室
    Case txt确认人
    End Select
End Sub

 
Private Sub Form_Load()
    mblnOK = False
    Call InitCommandBar
    Call LoadBaseInfo
    If mintType <> 0 Then
        Call LoadInputData
    End If
    Call SetFaceCtrl
    mblnChange = False
End Sub

Private Sub SetFaceCtrl()
    Select Case mintType
    Case 0
        txtInfo(txt报告时间).Locked = False
        txtInfo(txt报告科室).Locked = True
        txtInfo(txt报告人).Locked = False
        txtInfo(txt危急值描述).Locked = False
        
        cmdSel(txt报告科室).Visible = False
        
        picCL.Visible = False
    Case 1
        txtInfo(txt报告时间).Locked = False
        txtInfo(txt报告科室).Locked = True
        txtInfo(txt报告人).Locked = False
        txtInfo(txt危急值描述).Locked = False
        
        cmdSel(txt报告科室).Visible = False
        
        picCL.Visible = False
    Case 2
        txtInfo(txt报告时间).Locked = True
        txtInfo(txt报告科室).Locked = True
        txtInfo(txt报告人).Locked = True
        txtInfo(txt危急值描述).Locked = True
        imgDate(txt报告时间).Visible = False
        cmdSel(txt报告科室).Visible = False
        cmdSel(txt报告人).Visible = False
        
        
        txtInfo(txt处理情况).Locked = True
        txtInfo(txt确认时间).Locked = True
        txtInfo(txt确认科室).Locked = True
        txtInfo(txt确认人).Locked = True
        imgDate(txt确认时间).Visible = False
        cmdSel(txt确认科室).Visible = False
        cmdSel(txt确认人).Visible = False
        
        picInfo(0).Enabled = False
        
        '如果医生未处理则不显示下面部分
        If txtInfo(txt处理情况).Text = "" Then
            picCL.Visible = False
        Else
            picCL.Visible = True
        End If
    Case 3
        txtInfo(txt报告时间).Locked = True
        txtInfo(txt报告科室).Locked = True
        txtInfo(txt报告人).Locked = True
        txtInfo(txt危急值描述).Locked = True
        imgDate(txt报告时间).Visible = False
        cmdSel(txt报告科室).Visible = False
        cmdSel(txt报告人).Visible = False
        cmdSel(txt确认科室).Visible = False
        cmdSel(txt确认人).Visible = False
        
        txtInfo(txt处理情况).Locked = False
        txtInfo(txt确认时间).Locked = False
        txtInfo(txt确认科室).Locked = True
        txtInfo(txt确认人).Locked = True
        picCL.Visible = True
    End Select
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    picMain.Top = 525
    picMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (picMain.Width / 2)
    picCL.Left = 0
    picCL.Width = picMain.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
End Sub

Private Sub optInfo_Click(Index As Integer)
    mblnChange = True
End Sub

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
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
            objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 保存")
            objControl.BeginGroup = True
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

Private Sub LoadBaseInfo()
'功能：加载初始数据
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim i As Long
    
    '清空
    For i = 0 To txt报告时间
        txtInfo(i).Text = ""
    Next
    
    txtInfo(txt报告人).Text = ""
    txtInfo(txt报告科室).Text = ""
    txtInfo(txt报告时间).Text = ""
    txtInfo(txt危急值描述).Text = ""
    txtInfo(txt处理情况).Text = ""
    txtInfo(txt确认时间).Text = ""
    txtInfo(txt确认科室).Text = ""
    txtInfo(txt确认人).Text = ""
    
    On Error GoTo errH
    
    If mint调用类型 = 1 Then
        strSql = "select a.姓名, a.性别, a.年龄,b.名称 as 科室,a.门诊号 as 住院号,null as 床号,a.发生时间 as 就诊时间,b.id as 科室ID,a.复诊 from 病人挂号记录 a,部门表 b where a.执行部门id=b.id and a.no=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mstr挂号单)
        lblInfo(txt住院号).Caption = "门 诊 号"
        lblInfo(txt床号).Caption = "复    诊"
        txtInfo(txt床号).Text = IIF(Val(rsTmp!复诊 & "") = 1, "是", "否")
    ElseIf mint调用类型 = 2 Then
        strSql = "Select a.姓名, a.性别, a.年龄, b.名称 As 科室, a.住院号, a.出院病床 As 床号,a.入院日期 as 就诊时间,b.id as 科室ID" & vbNewLine & _
            "From 病案主页 A, 部门表 B Where a.出院科室id = b.Id And a.病人id = [1] And a.主页id = [2]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
        txtInfo(txt床号).Text = rsTmp!床号 & ""
    ElseIf mint调用类型 = 3 Then
        strSql = "select  a.姓名, a.性别, a.年龄 ,b.名称 as 科室,null as 住院号,null as 床号,a.开始执行时间 as 就诊时间,b.id as 科室ID,null as 复诊  from 病人医嘱记录 a,部门表 b where a.病人科室id=b.id and a.id=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID)
        lblInfo(txt住院号).Caption = "门 诊 号"
        lblInfo(txt床号).Caption = "复    诊"
        txtInfo(txt床号).Text = IIF(Val(rsTmp!复诊 & "") = 1, "是", "否")
    End If
    
    txtInfo(txt姓名).Text = rsTmp!姓名 & ""
    txtInfo(txt性别).Text = rsTmp!性别 & ""
    txtInfo(txt年龄).Text = rsTmp!年龄 & ""
    txtInfo(txt科室).Text = rsTmp!科室 & ""
        txtInfo(txt科室).Tag = Val(rsTmp!科室ID & "")
    txtInfo(txt住院号).Text = rsTmp!住院号 & ""
    
    mstr就诊时间 = Format(rsTmp!就诊时间 & "", "YYYY-MM-DD HH:mm")
     
    
    '新增
    strSql = "select b.名称,b.id as 登记科室ID,a.诊疗类别 from 病人医嘱记录 a,部门表 b where a.执行科室id=b.id and a.id=[1]"
    Set rsAdvice = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID)
    If Not rsAdvice.EOF Then
        mlng报告科室ID = Val(rsAdvice!登记科室id & "")
        txtInfo(txt报告科室).Text = rsAdvice!名称 & ""
        If rsAdvice!诊疗类别 & "" = "D" Then
            lblType.Caption = "检查类"
        Else
            lblType.Caption = "检验类"
        End If
    End If
         
    If mintType = 0 Then
        txtInfo(txt报告人).Text = UserInfo.姓名
        txtInfo(txt报告人).Tag = UserInfo.姓名
        
        txtInfo(txt报告时间).Text = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 39 Then KeyAscii = 0 '单引号蔽屏
    If KeyAscii = 13 Then
        Select Case Index
        Case txt报告人
            Call GetItem报告人(0)
        End Select
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case txt报告时间
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = txtInfo(Index).Tag
            Else
                txtInfo(Index).Tag = txtInfo(Index).Text
            End If
    End Select
End Sub

Private Sub imgDate_Click(Index As Integer)
    Select Case Index
        Case txt报告时间
            If IsDate(txtInfo(txt报告时间).Text) Then
                dtpDate.value = CDate(txtInfo(txt报告时间).Text)
            Else
                dtpDate.value = zldatabase.Currentdate
            End If
            dtpDate.Tag = "报告时间"
            dtpDate.Left = txtInfo(txt报告时间).Left + picMain.Left
            dtpDate.Top = txtInfo(txt报告时间).Top + txtInfo(txt报告时间).Height + picMain.Top + 20
            dtpDate.Visible = True
            dtpDate.SetFocus
        Case txt确认时间
            If IsDate(txtInfo(Index).Text) Then
                dtpDate.value = CDate(txtInfo(Index).Text)
            Else
                dtpDate.value = zldatabase.Currentdate
            End If
            dtpDate.Tag = "确认时间"
            dtpDate.Left = txtInfo(Index).Left + picMain.Left
            dtpDate.Top = txtInfo(Index).Top + txtInfo(Index).Height + picCL.Top + picMain.Top + 20
            dtpDate.Visible = True
            dtpDate.SetFocus
    End Select
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "报告时间" Then
        '取值
        If IsDate(txtInfo(txt报告时间).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt报告时间).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt报告时间).Text = strDate
        txtInfo(txt报告时间).Tag = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt报告时间).SetFocus
    ElseIf dtpDate.Tag = "确认时间" Then
        '取值
        If IsDate(txtInfo(txt确认时间).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt确认时间).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt确认时间).Text = strDate
        txtInfo(txt确认时间).Tag = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt确认时间).SetFocus
    End If
End Sub

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strMsg As String
    Dim strTmp As String
    
    If mintType = 0 Or mintType = 1 Then
        If txtInfo(txt危急值描述).Text = "" Then
            MsgBox "没有填写 危急值描述 。", vbInformation, gstrSysName
            If txtInfo(txt危急值描述).Enabled Then txtInfo(txt危急值描述).SetFocus
            Exit Function
        End If
        
        strTmp = txtInfo(txt危急值描述).Text
        If zlCommFun.ActualLen(strTmp) > txtInfo(txt危急值描述).MaxLength Then
            strMsg = "危急值描述-内容太长(允许录入" & txtInfo(txt危急值描述).MaxLength & "个字符或" & txtInfo(txt危急值描述).MaxLength \ 2 & "个汉字)。"
            MsgBox strMsg, vbInformation, gstrSysName
            If txtInfo(txt危急值描述).Enabled Then txtInfo(txt危急值描述).SetFocus
            Exit Function
        End If
                
        If txtInfo(txt报告时间).Text = "" Then
            MsgBox "没有填写 报告时间 。", vbInformation, gstrSysName
            If txtInfo(txt报告时间).Enabled Then txtInfo(txt报告时间).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt报告时间).Text <> "" Then
            If Not Check时间("报告时间", txtInfo(txt报告时间).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt报告时间).Enabled Then txtInfo(txt报告时间).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If mintType = 3 Then
        If txtInfo(txt处理情况).Text = "" Then
            MsgBox "没有填写 处理情况 。", vbInformation, gstrSysName
            If txtInfo(txt处理情况).Enabled Then txtInfo(txt处理情况).SetFocus
            Exit Function
        End If
        
        strTmp = txtInfo(txt处理情况).Text
        If zlCommFun.ActualLen(strTmp) > txtInfo(txt处理情况).MaxLength Then
            strMsg = "处理情况-内容太长(允许录入" & txtInfo(txt处理情况).MaxLength & "个字符或" & txtInfo(txt处理情况).MaxLength \ 2 & "个汉字)。"
            MsgBox strMsg, vbInformation, gstrSysName
            If txtInfo(txt处理情况).Enabled Then txtInfo(txt处理情况).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt确认时间).Text = "" Then
            MsgBox "没有填写 确认时间 。", vbInformation, gstrSysName
            If txtInfo(txt确认时间).Enabled Then txtInfo(txt确认时间).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt确认时间).Text <> "" Then
            If Not Check时间("确认时间", txtInfo(txt确认时间).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt确认时间).Enabled Then txtInfo(txt确认时间).SetFocus
                Exit Function
            End If
        End If
        
    End If
    
    CheckData = True
End Function

Private Function Check时间(ByVal strTimeType As String, ByVal str时间 As String, Optional ByRef strMsg As String) As Boolean
'功能：检查输入的时间是否合法
    Dim strInDate As String
    Dim datCurrent As Date
    
    If Not IsDate(str时间) Then
        strMsg = "输入的" & strTimeType & "无效。"
        Exit Function
    End If
    
    If "报告时间" = strTimeType Then
        datCurrent = zldatabase.Currentdate
        strInDate = mstr就诊时间
        
    
        If Format(str时间, "yyyy-MM-dd HH:mm") < Format(mstr就诊时间, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "不能小于病人的入院时间 " & strInDate & " 。"
            Exit Function
        End If
       
        If Format(str时间, "yyyy-MM-dd HH:mm") > Format(datCurrent, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "不能大于当前时间 " & Format(datCurrent, "yyyy-MM-dd HH:mm") & " 。"
            Exit Function
        End If
    ElseIf "确认时间" = strTimeType Then
        '确认时间应该大于报告时间
        If Format(str时间, "yyyy-MM-dd HH:mm") < txtInfo(txt报告时间).Text Then
            strMsg = strTimeType & "不能小于报告时间 " & strInDate & " 。"
            Exit Function
        End If
    End If
    Check时间 = True
End Function

Private Function SaveData() As Boolean
'功能：保存数据
    Dim strSql As String
    Dim strPars As String
    Dim int是否危急值 As Integer
    
    On Error GoTo errH
    
    mstr危急值描述 = txtInfo(txt危急值描述).Text
    mdat报告时间 = CDate(txtInfo(txt报告时间).Text)
    mstr报告人 = txtInfo(txt报告人).Text
    
    If mintType = 0 Then
        mlng危急值ID = zldatabase.GetNextID("病人危急值记录")        '获取危急值记录ID
    End If
    
    strPars = "(" & mlng危急值ID & ",null," & mlng病人ID & "," & ZVal(mlng主页ID) & "," & IIF(mstr挂号单 = "", "null", "'" & mstr挂号单 & "'") & "," & mlng婴儿 & ","
    strPars = strPars & "'" & txtInfo(txt姓名).Text & "','" & txtInfo(txt性别).Text & "','" & txtInfo(txt年龄).Text & "'," & mlng医嘱ID & "," & mlng标本ID & ",'" & mstr危急值描述 & "',"
    strPars = strPars & "to_date('" & Format(mdat报告时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),"
    strPars = strPars & mlng报告科室ID & ",'" & mstr报告人 & "')"
    
    If mintType = 0 Then
        strSql = "Zl_病人危急值记录_Insert" & strPars
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
        If lblType.Caption = "检查类" Then
            Call Send危急值消息(0)
        Else
            Call Send危急值消息(1)
        End If
        '新增状态保存后变为修改状态
        mintType = 1
    ElseIf mintType = 1 Then
        strSql = "Zl_病人危急值记录_Update" & strPars
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    ElseIf mintType = 3 Then
        mdat确认时间 = CDate(txtInfo(txt确认时间).Text)
        mstr处理情况 = txtInfo(txt处理情况).Text
        mstr确认人 = txtInfo(txt确认人).Text
        If optInfo(0).value Then
            int是否危急值 = 1
        Else
            int是否危急值 = 0
        End If
        strSql = "Zl_病人危急值记录_处理(" & mlng危急值ID & ",'" & mstr处理情况 & "',to_date('" & Format(mdat确认时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & mstr确认人 & "'," & mlng确认科室ID & "," & int是否危急值 & ")"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    mblnOK = True
    mblnChange = False
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitBaseBy记录ID(ByVal lngIdin As Long)
'功能：通过危急值初始基础信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
     
    On Error GoTo errH
    mlng病人ID = 4
    strSql = "Select a.Id, a.数据来源, a.病人id, a.主页id, a.挂号单, a.婴儿, a.姓名, a.性别, a.年龄, a.医嘱id, a.标本id, a.危急值描述, a.报告时间, a.报告科室id, a.报告人,a.处理情况, a.确认时间, a.确认人, a.确认科室id, a.状态, a.是否危急值 " & _
        " From 病人危急值记录 A where a.id=[1]"
    Set mrs危急值 = zldatabase.OpenSQLRecord(strSql, Me.Caption, lngIdin)
    
    mlng病人ID = Val(mrs危急值!病人ID & "")
    mlng主页ID = Val(mrs危急值!主页ID & "")
    mstr挂号单 = mrs危急值!挂号单 & ""
    mint调用类型 = IIF(mstr挂号单 = "", 2, 1)
    mlng医嘱ID = Val(mrs危急值!医嘱ID & "")
    mlng婴儿 = Val(mrs危急值!婴儿 & "")
    
    mlng标本ID = Val(mrs危急值!标本ID & "")
    mstr危急值描述 = Val(mrs危急值!危急值描述 & "")
    
    If Not IsNull(mrs危急值!报告时间) Then
        mdat报告时间 = Format(mrs危急值!报告时间 & "", "yyyy-MM-dd HH:mm:ss")
    End If
    
    mlng报告科室ID = Val(mrs危急值!报告科室id & "")
    mstr报告人 = Val(mrs危急值!报告人 & "")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If dtpDate.Tag = "报告时间" Then
            txtInfo(txt报告时间).SetFocus
        ElseIf dtpDate.Tag = "确认时间" Then
            txtInfo(txt确认时间).SetFocus
        End If
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub LoadInputData()
'功能：查看时加载界面新增时填写的信息
    
    Call InitBaseBy记录ID(mlng危急值ID)
    
    txtInfo(txt报告时间).Text = Format(mrs危急值!报告时间, "yyyy-MM-dd HH:mm")
    txtInfo(txt危急值描述).Text = mrs危急值!危急值描述 & ""
    txtInfo(txt处理情况).Text = mrs危急值!处理情况 & ""
    txtInfo(txt报告人).Text = mrs危急值!报告人 & ""
    txtInfo(txt报告科室).Text = Sys.RowValue("部门表", Val(mrs危急值!报告科室id & ""), "名称")
    txtInfo(txt报告时间).Text = Format(mrs危急值!报告时间, "yyyy-MM-dd HH:mm")
    
    
    Select Case mintType
    Case 0
    Case 1
    Case 2
        If txtInfo(txt处理情况).Text <> "" Then
            txtInfo(txt确认时间).Text = Format(mrs危急值!确认时间, "yyyy-MM-dd HH:mm")
            txtInfo(txt确认人).Text = mrs危急值!确认人 & ""
            mlng确认科室ID = Val(mrs危急值!确认科室ID & "")
            txtInfo(txt确认科室).Text = Sys.RowValue("部门表", mlng确认科室ID, "名称")
            If Val(mrs危急值!是否危急值 & "") = 1 Then
                optInfo(0).value = True
                optInfo(1).value = False
            Else
                optInfo(0).value = False
                optInfo(1).value = True
            End If
        End If
    Case 3
        If txtInfo(txt处理情况).Text = "" Then
            '如果没有填写处理情况，加载缺省值，确认时间默认为报告时间，确认科室为病人科室，确认人为当前操作员，默认为是危急值
            txtInfo(txt确认时间).Text = Format(mrs危急值!报告时间, "yyyy-MM-dd HH:mm")
            txtInfo(txt确认人).Text = UserInfo.姓名
            mlng确认科室ID = Val(txtInfo(txt科室).Tag)
            txtInfo(txt确认科室).Text = txtInfo(txt科室).Text
            optInfo(0).value = True
            optInfo(1).value = False
        Else
            txtInfo(txt确认时间).Text = Format(mrs危急值!确认时间, "yyyy-MM-dd HH:mm")
            txtInfo(txt确认人).Text = mrs危急值!确认人 & ""
            mlng确认科室ID = Val(mrs危急值!确认科室ID & "")
            txtInfo(txt确认科室).Text = Sys.RowValue("部门表", mlng确认科室ID, "名称")
            If Val(mrs危急值!是否危急值 & "") = 1 Then
                optInfo(0).value = True
                optInfo(1).value = False
            Else
                optInfo(0).value = False
                optInfo(1).value = True
            End If
        End If
    End Select
    
End Sub

Private Sub Send危急值消息(ByVal intType As Integer)
'功能：发送PACS危值消息,ZLHIS_PACS_005,LIS危急值消息 ZLHIS_LIS_003
'参数：intType，0-ZLHIS_PACS_005，1-ZLHIS_LIS_003
    Dim strXML As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim str姓名 As String
    Dim str门诊号 As String
    Dim str住院号 As String
    Dim int病人来源 As Integer
    Dim lng就诊ID As Long
    Dim lng病区ID As Long '病案主页.当前病区ID
    Dim lng科室id As Long '开嘱科室ID
    Dim lng项目id As Long
    Dim str医嘱内容 As String
    Dim lng执行科室ID As Long
    Dim str部门IDs As String
    
    On Error GoTo errH
    
    str姓名 = txtInfo(txt姓名).Text
    
    If mint调用类型 = 1 Then
        str门诊号 = txtInfo(txt住院号).Text
        strSql = "select b.id as 就诊ID,a.病人来源,a.开嘱科室id,a.诊疗项目id,a.医嘱内容,a.执行科室id from 病人医嘱记录 a,病人挂号记录 b" & _
            "  where a.挂号单=b.no and a.挂号单=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mstr挂号单)
        If Not rsTmp.EOF Then
            int病人来源 = Val(rsTmp!病人来源 & "")
            lng科室id = Val(rsTmp!开嘱科室id & "")
            lng项目id = Val(rsTmp!诊疗项目ID & "")
            str医嘱内容 = rsTmp!医嘱内容 & ""
            lng执行科室ID = Val(rsTmp!执行科室ID & "")
            lng就诊ID = Val(rsTmp!就诊ID & "")
        End If
    Else
        str住院号 = txtInfo(txt住院号).Text
        lng就诊ID = mlng主页ID
        strSql = "select a.病人来源,b.当前病区id,a.开嘱科室id,a.诊疗项目id,a.医嘱内容,a.执行科室id from 病人医嘱记录 a,病案主页 b" & _
            "  where a.病人id=b.病人id and a.主页id=b.主页id and a.id=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID)
        If Not rsTmp.EOF Then
            int病人来源 = Val(rsTmp!病人来源 & "")
            lng病区ID = Val(rsTmp!当前病区ID & "")
            lng科室id = Val(rsTmp!开嘱科室id & "")
            lng项目id = Val(rsTmp!诊疗项目ID & "")
            str医嘱内容 = rsTmp!医嘱内容 & ""
            lng执行科室ID = Val(rsTmp!执行科室ID & "")
        End If
    End If
    
    If intType = 0 Then
        strXML = "<patient_info>" & vbNewLine & _
            "              <patient_id>" & mlng病人ID & "</patient_id>" & vbNewLine & _
            "              <patient_name>" & str姓名 & "</patient_name>" & vbNewLine & _
            "              <in_number>" & str住院号 & "</in_number>" & vbNewLine & _
            "              <out_number>" & str门诊号 & "</out_number>" & vbNewLine & _
            "          </patient_info>" & vbNewLine & _
            "          <patient_clinic>" & vbNewLine & _
            "              <patient_source>" & int病人来源 & "</patient_source>" & vbNewLine & _
            "              <clinic_id>" & lng就诊ID & "</clinic_id>" & vbNewLine & _
            "              <clinic_area_id>" & lng病区ID & "</clinic_area_id>" & vbNewLine & _
            "              <clinic_dept_id>" & lng科室id & "</clinic_dept_id>" & vbNewLine & _
            "          </patient_clinic>" & vbNewLine & _
            "          <check_order>" & vbNewLine & _
            "              <order_id>" & mlng医嘱ID & "</order_id>" & vbNewLine & _
            "              <check_item_id>" & lng项目id & "</check_item_id>" & vbNewLine & _
            "              <check_item_title>" & str医嘱内容 & "</check_item_title>" & vbNewLine & _
            "              <study_execute_id>" & lng执行科室ID & "</study_execute_id>" & vbNewLine & _
            "          </check_order>"
        If Not (mclsMipModule Is Nothing) Then
            If mclsMipModule.IsConnect Then
                Call mclsMipModule.CommitMessage("ZLHIS_PACS_005", strXML)
            End If
        End If
        Call zldatabase.SendMsg("ZLHIS_PACS_005", strXML)
    Else
        '产生新开消息，LIS消息先不按消息平台处理
        str部门IDs = lng科室id
        If lng病区ID <> 0 Then
            If lng病区ID <> lng科室id Then
                str部门IDs = str部门IDs & "," & lng病区ID
            End If
        End If
        strSql = "Zl_业务消息清单_Insert(" & mlng病人ID & "," & lng就诊ID & "," & lng科室id & ","
        strSql = strSql & IIF(lng病区ID = 0, "NULL", lng病区ID) & "," & int病人来源 & ","
        strSql = strSql & "'" & mstr危急值描述 & "','1110','ZLHIS_LIS_003','" & mlng医嘱ID & "',3,0,sysdate,'" & str部门IDs & "',null)"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    
    If mblnChange Then
        
        strMsg = "当前内容编辑后尚未保存，确实要退出吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        
    End If
End Sub

Private Sub PrintApply(ByVal intType As Integer)
'功能打印预览申请单
'参数：intType:1-预览，2-打印
    '判断如果还未保存则先保存再打印
    If mblnChange Then
        If CheckData = False Then Exit Sub
        If SaveData() Then
            mblnOK = True
        End If
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_20", Me, "记录ID=" & mlng危急值ID, intType)
End Sub

Private Sub GetItem报告人(ByVal intType As Integer)
'功能：获取主刀医生项目
'参数：0 文本框按回车，1 点按钮
    Dim strSql As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean, strTmp As String
    Dim blnDo As Boolean, str部门 As String
    Dim lng部门ID As Long, lng人员id As Long
    Dim i As Integer
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt报告人).Tag = txtInfo(txt报告人).Text Then
'            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt报告人).Text = "" Then '相当于是清除该项目
            txtInfo(txt报告人).Tag = ""
'            Call SeekNextCtl
            Exit Sub
        End If
    End If
            
    strInput = Trim(UCase(txtInfo(txt报告人).Text))   '传入的值存在前缀空格
    
    strSql = "Select A.ID,A.编号,A.姓名,A.简码,A.手术等级" & _
        " From 人员表 A,人员性质说明 B" & _
        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        IIF(intType = 0, " And (A.编号 Like [1] Or A.姓名 Like [2] Or A.简码 Like [2])", "") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编号"
    vRect = zlControl.GetControlRect(txtInfo(txt报告人).Hwnd)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "报告人", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt报告人).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            If MsgBox("没有找到匹配的医生，你确定要输入没有建立人员档案的医生吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnDo = True
                strTmp = strInput
            Else
                blnDo = False
            End If
        Else
            Call MsgBox("没有找到匹配的医生!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        txtInfo(txt报告人).Text = rsTmp!姓名 & ""
        txtInfo(txt报告人).Tag = rsTmp!姓名 & ""
        lng人员id = rsTmp!ID
        txtInfo(txt报告人).SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
