VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBloodApply 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "输血申请单"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10650
   Icon            =   "frmBloodApply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   10650
   StartUpPosition =   2  '屏幕中心
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
      Left            =   9480
      TabIndex        =   82
      Top             =   2520
      Width           =   735
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10080
      TabIndex        =   81
      Top             =   5880
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   113115137
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picNo 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8400
      ScaleHeight     =   495
      ScaleWidth      =   1935
      TabIndex        =   78
      Top             =   1080
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
         TabIndex        =   79
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
         TabIndex        =   80
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
      Left            =   1440
      ScaleHeight     =   270
      ScaleWidth      =   4095
      TabIndex        =   77
      Top             =   3000
      Width           =   4095
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
         Left            =   -25
         TabIndex        =   10
         Top             =   -25
         Width           =   3960
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
      ScaleWidth      =   915
      TabIndex        =   76
      Top             =   2070
      Width           =   915
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
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   -25
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "…"
      Height          =   270
      Left            =   9060
      TabIndex        =   9
      Top             =   2520
      Width           =   270
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   1
      Left            =   10080
      Picture         =   "frmBloodApply.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   9105
      Width           =   285
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   0
      Left            =   4200
      Picture         =   "frmBloodApply.frx":6948
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "编辑(F4)"
      Top             =   3960
      Width           =   285
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   9
      Left            =   6720
      ScaleHeight     =   300
      ScaleWidth      =   3375
      TabIndex        =   73
      Top             =   5400
      Width           =   3375
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
         TabIndex        =   28
         Top             =   -25
         Width           =   3135
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   8
      Left            =   6720
      ScaleHeight     =   300
      ScaleWidth      =   3375
      TabIndex        =   72
      Top             =   4440
      Width           =   3375
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
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   24
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7920
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
      Left            =   8400
      TabIndex        =   33
      Text            =   "2013-06-20 18:00"
      Top             =   9120
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8640
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8280
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
      TabIndex        =   63
      Top             =   5400
      Width           =   3375
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   1
         Left            =   2970
         Picture         =   "frmBloodApply.frx":6A3E
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   10
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
         TabIndex        =   27
         Top             =   0
         Width           =   2940
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   25
      Top             =   4920
      Width           =   3375
   End
   Begin VB.PictureBox picGet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   0
      Left            =   2040
      ScaleHeight     =   300
      ScaleWidth      =   3375
      TabIndex        =   58
      Top             =   4440
      Width           =   3375
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   3020
         Picture         =   "frmBloodApply.frx":6B34
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   10
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
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   2
      Left            =   8880
      ScaleHeight     =   300
      ScaleWidth      =   1095
      TabIndex        =   56
      Top             =   3930
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
         TabIndex        =   22
         Top             =   -25
         Width           =   975
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Index           =   1
      Left            =   6720
      ScaleHeight     =   300
      ScaleWidth      =   1455
      TabIndex        =   54
      Top             =   3930
      Width           =   1455
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
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   -25
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
      TabIndex        =   20
      Text            =   "2013-06-20 18:00"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   2
      Left            =   8520
      TabIndex        =   51
      Top             =   3360
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
         Top             =   120
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
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   49
      Top             =   3360
      Width           =   3015
      Begin VB.OptionButton optPregnancy 
         BackColor       =   &H8000000E&
         Caption         =   "正常"
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
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPregnancy 
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
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton optPregnancy 
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
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame fraChk 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   47
      Top             =   3360
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
         Left            =   720
         TabIndex        =   14
         Top             =   120
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
         TabIndex        =   13
         Top             =   120
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
      Left            =   6720
      ScaleHeight     =   270
      ScaleWidth      =   3495
      TabIndex        =   45
      Top             =   3000
      Width           =   3495
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
         TabIndex        =   12
         Top             =   -25
         Width           =   3375
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
      Index           =   8
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7335
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
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
      Top             =   2070
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1650
      Width           =   855
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
      Height          =   255
      Index           =   12
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4920
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid vsLIS 
      Height          =   1605
      Left            =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6240
      Width           =   10125
      _cx             =   17859
      _cy             =   2831
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
      FormatString    =   $"frmBloodApply.frx":6C2A
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
      X1              =   8400
      X2              =   9960
      Y1              =   8190
      Y2              =   8190
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
      Index           =   36
      Left            =   6960
      TabIndex        =   71
      Top             =   7950
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   8400
      X2              =   10080
      Y1              =   9270
      Y2              =   9270
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
      Left            =   6960
      TabIndex        =   70
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   8400
      X2              =   9960
      Y1              =   8910
      Y2              =   8910
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
      Left            =   6960
      TabIndex        =   69
      Top             =   8670
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   8400
      X2              =   9960
      Y1              =   8550
      Y2              =   8550
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   23
      Left            =   240
      TabIndex        =   67
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   6600
      X2              =   9840
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   22
      Left            =   5475
      TabIndex        =   66
      Top             =   5430
      Width           =   1140
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   1920
      X2              =   5340
      Y1              =   5700
      Y2              =   5700
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
      Left            =   720
      TabIndex        =   62
      Top             =   5430
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   6600
      X2              =   7680
      Y1              =   5190
      Y2              =   5190
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   1920
      X2              =   5320
      Y1              =   5190
      Y2              =   5190
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   19
      Left            =   480
      TabIndex        =   60
      Top             =   4950
      Width           =   1305
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   6600
      X2              =   9840
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   18
      Left            =   5475
      TabIndex        =   59
      Top             =   4470
      Width           =   1140
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   1920
      X2              =   5340
      Y1              =   4740
      Y2              =   4740
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
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   57
      Top             =   4470
      Width           =   1785
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   8760
      X2              =   9840
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "RHD"
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
      Left            =   8280
      TabIndex        =   55
      Top             =   3960
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   6600
      X2              =   8040
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   1920
      X2              =   4200
      Y1              =   4230
      Y2              =   4230
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
      TabIndex        =   52
      Top             =   3990
      Width           =   1815
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   13
      Left            =   7320
      TabIndex        =   50
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   48
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   11
      Left            =   300
      TabIndex        =   46
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   6600
      X2              =   10080
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1440
      X2              =   5340
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1440
      X2              =   9360
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   6600
      X2              =   7920
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   9120
      X2              =   10080
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1440
      X2              =   2880
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6600
      X2              =   7920
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4080
      X2              =   5400
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1440
      X2              =   2880
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4080
      X2              =   5400
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9120
      X2              =   10080
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3960
      X2              =   6720
      Y1              =   1155
      Y2              =   1155
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
      Top             =   750
      Width           =   3615
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   44
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   43
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   42
      Top             =   2550
      Width           =   975
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
      TabIndex        =   40
      Top             =   2100
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
      TabIndex        =   39
      Top             =   2100
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
      TabIndex        =   38
      Top             =   2070
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
      TabIndex        =   36
      Top             =   1710
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
      TabIndex        =   35
      Top             =   2070
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
      TabIndex        =   34
      Top             =   1680
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
      TabIndex        =   41
      Top             =   1710
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
      TabIndex        =   37
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label lblInfo 
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
      Height          =   255
      Index           =   15
      Left            =   6000
      TabIndex        =   53
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "单位"
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
      Index           =   20
      Left            =   5985
      TabIndex        =   61
      Top             =   4950
      Width           =   540
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "主治医师签名"
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
      Left            =   6960
      TabIndex        =   68
      Top             =   8310
      Width           =   1335
   End
End
Attribute VB_Name = "frmBloodApply"
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
Private mintType As Integer   '0-新增，1-修改，2-查看,3-医嘱编辑调用，只能调整除输血成分，总量，申请时间，输血时间，执行科室，输血途径，输血执行科室，用血安排以外的内容
Private mlngUpdateAdvice As Long  '修改的医嘱ID
Private mintPState As Integer
Private mdatTurn As Date
Private mlng病人科室id As Long
Private mlng病区ID As Long
Private mlng开单科室ID As Long
Private mlng输血途径 As Long
Private mlng输血项目ID As Long
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
Private mint申请单打印模式 As Integer  '1-发送时打印，2-新开时打印
Private mint险类 As Integer '当前病人险类
Private mbln提醒对码 As Boolean
Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象
Attribute mclsMipModule.VB_VarHelpID = -1
Private Const CON_LisResultCol = 3
Private Const CON_LisResultCount = 10
Private mobjPublicLis As Object
Private mint场合 As Integer '0-住院，1－门诊，默认为住院
Private mstr挂号单 As String '挂号单号

Private Enum Enum_Cbo
    cbo输血性质 = 0
    cbo输血血型 = 1
    cboRhd = 2
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
    txt预定输血时间 = 10
    txt预定输血量 = 11
    txt单位 = 12
    txt主治医师签名 = 17
    txt申请日期 = 19
    txt申请医师签名 = 20
End Enum

Private Enum Enum_Get
    txt预定输血成分 = 0
    txt输血途径 = 1
End Enum

Private Enum Enum_Date
    cmd预定输血时间 = 0
    cmd申请日期 = 1
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

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病人性质 As Long, ByVal intType As Integer, Optional ByVal lngUpdateAdvice As Long, _
            Optional ByVal lng病人科室ID As Long, Optional ByVal lng病区ID As Long, Optional ByVal lng开单科室ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
            Optional ByRef rsDefine As Recordset, Optional ByRef objMip As Object, Optional ByVal int场合 As Integer, Optional ByVal str挂号单 As String) As Boolean
    
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
    Set mrsDefine = rsDefine
    mlngUpdateAdvice = lngUpdateAdvice
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    ShowMe = mblnOK
End Function

Private Function SeekNextControl() As Boolean
'功能：定位到下一个焦点的控件上
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextControl = True
End Function

Private Sub cboInfo_Change(Index As Integer)
    If Visible And Index = cbo输血目的 Then mblnChange = True
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSql As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = cbo执行科室 Or Index = cbo输血执行 Then
        If cboInfo(Index).ItemData(cboInfo(Index).ListIndex) = -1 Then
            
            '他科执行，弹出选择执行科室
            strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                IIF(gstrNodeNo <> "", " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " Order by A.编码"
            vRect = GetControlRect(cboInfo(Index).hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "执行科室", , , , , , True, vRect.Left, vRect.Top, cboInfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = SeekCboIndex(cboInfo(Index), rsTmp!ID)
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
                    intIdx = SeekCboIndex(cboInfo(Index), Val(cboInfo(Index).Tag))
                    Call zlControl.CboSetIndex(cboInfo(Index).hWnd, intIdx)
                End If
            End If
        End If
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
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

Private Function CheckData() As Boolean
'功能：检查数据正确性
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    
'    Call SeekNextControl  '用这种方式会出问题71290
'这里采用两次设不同控件的焦点，确保validata事件的执行。
    txtGet(txt预定输血成分).SetFocus
    txtGet(txt输血途径).SetFocus
    
    '编辑附项不检查以下内容
    If mintType <> 3 Then
        '必须录入输血成分
        If mlng输血项目ID = 0 Then
            MsgBox "没有确定预定输血成分。", vbInformation, Me.Caption
            If txtGet(txt预定输血成分).Enabled Then txtGet(txt预定输血成分).SetFocus
            Exit Function
        End If
        
        '检查执行科室
        If cboInfo(cbo执行科室).Text = "" Then
            MsgBox "没有确定执行科室。", vbInformation, Me.Caption
            If cboInfo(cbo执行科室).Enabled Then cboInfo(cbo执行科室).SetFocus
            Exit Function
        End If
        
        '检查输血途径和输血执行
        If mlng输血途径 = 0 Then
            MsgBox "没有指定输血途径。", vbInformation, Me.Caption
            If txtGet(txt输血途径).Enabled Then txtGet(txt输血途径).SetFocus
            Exit Function
        End If
        If cboInfo(cbo输血执行).Text = "" Then
            MsgBox "没有确定输血执行科室。", vbInformation, Me.Caption
            If cboInfo(cbo输血执行).Enabled Then cboInfo(cbo输血执行).SetFocus
            Exit Function
        End If
        
        '必须录入总量
        If Val(txtInfo(txt预定输血量).Text) <= 0 Then
            MsgBox "请录入大于0的预定输血量。", vbInformation, Me.Caption
            If txtInfo(txt预定输血量).Enabled Then txtInfo(txt预定输血量).SetFocus
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
        '检查紧急医嘱必须填写输血目的
        If cboInfo(cbo用血安排).ListIndex = 1 And cboInfo(cbo输血目的).Text = "" Then
            MsgBox "紧急输血必须填写输血目的。", vbInformation, Me.Caption
            If cboInfo(cbo输血目的).Enabled Then cboInfo(cbo输血目的).SetFocus
            Exit Function
        End If
        '对码检查
        strIDs = mlng输血项目ID & ":"
        If Val(cboInfo(cbo执行科室).Tag & "") <> 0 Then
            strIDs = strIDs & Val(cboInfo(cbo执行科室).Tag & "")
        End If
        str医嘱内容 = FormatAdviceContext(txtGet(txt预定输血成分).Text, txtGet(txt输血途径).Text)
        strIDs = strIDs & "," & mlng输血途径 & ":"
        If Val(cboInfo(cbo输血执行).Tag & "") <> 0 Then
            strIDs = strIDs & Val(cboInfo(cbo输血执行).Tag & "")
        End If
        If gint医保对码 = 2 Then mbln提醒对码 = True
        
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
    End If
    
    CheckData = True
End Function

Private Function SaveData() As Boolean
'功能：保存数据
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lng医嘱ID As Long, lng医嘱序号 As Long, lng申请序号 As Long
    Dim strSql As String, rsTmp As Recordset
    Dim str项目名称 As String, str输血途径 As String
    Dim curDate As Date, i As Long, lng相关ID As String, j As Long
    Dim lngCount As Long, int病人来源 As Integer
    Dim strTmp主页ID As String
    Dim strTmp挂号单 As String

    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    If mintType = 3 Then
        '申请附项编辑模式
        lng相关ID = mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Delete(" & lng相关ID & ")"
    Else
        
        lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
        
        '病人医嘱记录.序号，递增
        If mint场合 = 0 Then
            lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, 0) + 1
            strTmp主页ID = mlng主页ID
            strTmp挂号单 = "NULL"
            int病人来源 = 2
        Else
            lng医嘱序号 = GetMaxAdviceNO(mlng病人ID, , 0, mstr挂号单) + 1
            strTmp主页ID = "NULL"
            strTmp挂号单 = "'" & mstr挂号单 & "'"
            int病人来源 = 1
        End If
        
        str项目名称 = Get项目名称(mlng输血项目ID)
        str输血途径 = Get项目名称(mlng输血途径)
        If mlngUpdateAdvice <> 0 Then
            '修改医嘱，删除后重新插入
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & mlngUpdateAdvice & ",1)"
            
            '取申请序号
            strSql = "Select 申请序号 From 病人医嘱记录 where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
            lng申请序号 = Val(rsTmp!申请序号 & "")
        End If
        If lng申请序号 = 0 Then
            '取申请序号
            strSql = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
            lng申请序号 = Val(rsTmp!申请序号 & "")
        End If
        '输血医嘱
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng医嘱ID & ",NULL," & lng医嘱序号 & "," & int病人来源 & "," & mlng病人ID & "," & strTmp主页ID & ",0,1,1,'K'," & mlng输血项目ID & _
                                 ",NULL,NULL,NULL," & ZVal(txtInfo(txt预定输血量).Text) & ",'" & FormatAdviceContext(str项目名称, str输血途径) & _
                                 "',Null,'" & Format(txtInfo(txt预定输血时间).Text, "yyyy-MM-dd HH:mm:ss") & "','一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "Null", cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo执行科室).ItemData(cboInfo(cbo执行科室).ListIndex) <= 0, "5", mlng执行科室性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp挂号单 & ",NULL,Null,0,NULL,NULL,'" & UserInfo.姓名 & "',Null,NULL,'" & cboInfo(cbo输血目的).Text & "'," & _
                                 IIF(gbln输血分级管理 And cboInfo(cbo用血安排).ListIndex <> 1, 1, IIF(gbln血库系统, 4, "NULL")) & "," & lng申请序号 & ")"
        
        '输血途径
        lng相关ID = lng医嘱ID
        lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
        lng医嘱序号 = lng医嘱序号 + 1
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng医嘱ID & "," & lng相关ID & "," & lng医嘱序号 & "," & int病人来源 & "," & mlng病人ID & "," & strTmp主页ID & _
                                 ",0,1,1,'E'," & mlng输血途径 & ",NULL,NULL,NULL,Null,'" & str输血途径 & "',Null,NULL,'一次性',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "Null", cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex) <= 0, "5", mlng输血执行性质) & "," & IIF(mbln补录, 2, cboInfo(cbo用血安排).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng病人科室id & "," & mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt申请日期).Text), txtInfo(txt申请日期).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp挂号单 & ",NULL,Null,0,NULL,NULL,'" & UserInfo.姓名 & "',Null,NULL,''," & _
                                 IIF(gbln输血分级管理 And cboInfo(cbo用血安排).ListIndex <> 1, 1, IIF(gbln血库系统, 4, "NULL")) & "," & lng申请序号 & ")"
    End If
    '输血申请其他项目
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_输血申请记录_Insert(" & lng相关ID & "," & chkWait.value & "," & cboInfo(cbo输血性质).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optPregnancy(0).value, 0, IIF(optPregnancy(1).value, 1, 2)) & "," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo输血血型).ListIndex & "," & cboInfo(cboRhd).ListIndex & ")"
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
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mlngUpdateAdvice = lng相关ID
    
    If Not (mclsMipModule Is Nothing) And mint场合 = 0 Then
        If mclsMipModule.IsConnect Then
            Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, Trim(txtInfo(txt姓名).Text), Trim(txtInfo(txt住院号).Text), , IIF(mlng病人性质 = 1, 1, 2), _
                mlng主页ID, mlng病区ID, , mlng病人科室id, "", , Trim(txtInfo(txt床号).Text), _
                lng医嘱ID, 0, 1, "K", "", UserInfo.姓名, Format(txtInfo(txt申请日期).Text, "yyyy-MM-dd HH:mm:ss"), mlng开单科室ID, "", , , "")
        End If
    End If
    
    SaveData = True
    mblnChange = False
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
    If Index = cbo输血目的 Then
        If zlCommFun.ActualLen(cboInfo(Index).Text) > 50 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
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
    Else
        '如果不可用，则检查医嘱是否符合
        If CheckData = False Then Exit Sub
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_17", Me, "医嘱ID=" & mlngUpdateAdvice, intType)
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_17", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '保存
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
            If Control.ID = conMenu_Edit_SaveExit Then
                Unload Me
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
        Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview
            blnVisible = (mint申请单打印模式 = 2 And InStr(GetInsidePrivs(p住院医嘱下达), ";输血申请单;") > 0) Or mint场合 = 1
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub chkWait_Click()
    If chkWait.value = 1 Then
        txtInfo(txt诊断信息).Text = "待诊"
        cmdInfo.Enabled = False
        mstr诊断IDs = ""
    Else
        txtInfo(txt诊断信息).Text = ""
        cmdInfo.Enabled = True
    End If
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
    dtpDate.SetFocus
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Call TxtGetInfo(Index, 1)
End Sub

Private Sub cmdInfo_Click()
    Dim str诊断 As String

    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng病人性质 = 1, 1260, 1261), mclsMipModule)
    End If
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 2), mlng病人科室id, mstr诊断IDs, str诊断, 0, mlngUpdateAdvice) Then
        txtInfo(txt诊断信息).Text = str诊断
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
        
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '住院场合调用时才做以下检查
    If mint场合 = 0 Then
        strInDate = mstr入院时间
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "医嘱的开始执行时间不能小于病人的入院时间 " & strInDate & " 。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    
    
        strInDate = ""
        If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
            strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
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
            strMsg = "输血时间不能小于医嘱开始时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check安排时间 = True
End Function

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

Private Sub Form_Load()
    mblnHaveAuditPriv = HaveAuditPriv
    mblnEditable = True
    mstr诊断IDs = ""
    mblnOK = False
    mbln提醒对码 = True
    vsLIS.Rows = 0
    If mint场合 = 0 Then mint申请单打印模式 = Val(zlDatabase.GetPara("输血申请单打印模式", glngSys, p住院医嘱发送, "1"))
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
    ElseIf mintType = 3 Then
        '只能调整除输血成分，总量，申请时间，输血时间，执行科室，输血途径，输血执行科室，用血安排以外的内容
        SetControlEnabled txtInfo(txt申请日期), False
        SetControlEnabled cmdDate(cmd申请日期), False
        SetControlEnabled txtInfo(txt预定输血时间), False
        SetControlEnabled cmdDate(cmd预定输血时间), False
        SetControlEnabled txtInfo(txt预定输血成分), False
        SetControlEnabled txtGet(txt预定输血成分), False
        SetControlEnabled cmdGet(txt预定输血成分), False
        SetControlEnabled txtGet(txt输血途径), False
        SetControlEnabled cmdGet(txt输血途径), False
        SetControlEnabled txtInfo(txt预定输血量), False
        SetControlEnabled cboInfo(cbo执行科室), False
        SetControlEnabled cboInfo(cbo输血执行), False
        SetControlEnabled cboInfo(cbo用血安排), False
        SetControlEnabled cboInfo(cbo输血目的), False
    End If
    mblnChange = False
    Call InitCommandBar
    If InitInfo = False Then Exit Sub
    Call LoadData
    Call SetFaceEnabledFalse
    If mbln补录 Then SetControlEnabled cboInfo(cbo用血安排), False
    '病人基本信息不可以编辑
    SetControlEnabled txtInfo(txt性别), False
    SetControlEnabled txtInfo(txt姓名), False
    SetControlEnabled txtInfo(txt年龄), False
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

Private Function InitInfo() As Boolean
'功能：初始下拉菜单
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim curDate As Date
    Dim lng用法ID As Long
    Dim lng执行科室ID As Long
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '部分固定内容的下拉框
    Call SetCboFromList(cboInfo, Array(" ", "A", "B", "O", "AB"), Array(cbo输血血型), 0)
    Call SetCboFromList(cboInfo, Array(" ", "常规", "紧急", "大量", "特殊"), Array(cbo输血性质), 1)
    Call SetCboFromList(cboInfo, Array("普通", "急诊"), Array(cbo用血安排), 0)
    Call SetCboFromList(cboInfo, Array(" ", "-", "+"), Array(cboRhd), 0)
    
    strSql = "select 名称 from 输血目的 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
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
    
    txtInfo(txt预定输血时间).Text = Format(curDate, "YYYY-MM-DD HH:mm")
    txtInfo(txt申请日期).Text = Format(curDate, "YYYY-MM-DD HH:mm")
    txtInfo(txt申请日期).Tag = txtInfo(txt申请日期).Text
    
    '缺省用法
    lng用法ID = Get缺省用法ID(8, IIF(mint场合 = 0, 2, 1))
    
    If lng用法ID = 0 Then
        MsgBox "没有可用的输血途径,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
        Screen.MousePointer = 0
        Unload Me
        Exit Function
    Else
        Set rsTmp = Get诊疗项目记录(lng用法ID)
        txtGet(txt输血途径).Text = rsTmp!名称 & ""
        mlng输血执行性质 = Nvl(rsTmp!执行科室, 0)
        txtGet(txt输血途径).Tag = txtGet(txt输血途径).Text
        mlng输血途径 = lng用法ID
        cboInfo(cbo输血执行).Enabled = True
        Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(cbo输血执行), "E", mlng输血途径, 0, _
            Val(rsTmp!执行科室 & ""), mlng病人科室id, mlng开单科室ID, 0, 1, IIF(mlng病人性质 = 1, 1, 2))
        If cboInfo(cbo输血执行).ListIndex = -1 And cboInfo(cbo输血执行).ListCount > 1 Then
'            cboInfo(cbo输血执行).ListIndex = 0
            Call zlControl.CboSetIndex(cboInfo(cbo输血执行).hWnd, 0)
        Else
            '如果有多项，则取默认的执行科室
            lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", mlng输血途径, 0, _
                    Nvl(rsTmp!执行科室, 0), mlng病人科室id, mlng开单科室ID, 1, IIF(mlng病人性质 = 1, 1, 2))
            If lng执行科室ID <> 0 Then
                Call zlControl.CboLocate(cboInfo(cbo输血执行), lng执行科室ID, True)
            End If
        End If
        If cboInfo(cbo输血执行).ListCount = 2 Then cboInfo(cbo输血执行).Enabled = False
        cboInfo(cbo输血执行).Tag = lng用法ID
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
    Dim strSql As String
    On Error GoTo errH
    
    '读取病人相关信息
    txtInfo(txt就诊类型).Text = IIF(mlng病人性质 = 1, "门诊", "住院")
    
    If mint场合 = 0 Then
        strSql = "Select A.住院号, Nvl(C.姓名, A.姓名) 姓名, Nvl(C.性别, A.性别) 性别, Nvl(C.年龄, A.年龄) 年龄, B.名称 As 科室, C.出院病床 As 当前床号, C.入院日期, C.险类" & vbNewLine & _
                "From 病人信息 A, 部门表 B, 病案主页 C" & vbNewLine & _
                "Where C.出院科室id = B.Id And A.病人id = C.病人id And A.主页id = C.主页id And C.病人id = [1] And C.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.RecordCount > 0 Then
            txtInfo(txt住院号).Text = rsTmp!住院号 & ""
            txtInfo(txt姓名).Text = rsTmp!姓名 & ""
            txtInfo(txt性别).Text = rsTmp!性别 & ""
            If txtInfo(txt性别).Text = "男" Then
                SetControlEnabled optPregnancy(0), False
                SetControlEnabled optPregnancy(1), False
                SetControlEnabled optPregnancy(2), False
            End If
            txtInfo(txt科室).Text = rsTmp!科室 & ""
            txtInfo(txt床号).Text = rsTmp!当前床号 & ""
            txtInfo(txt年龄).Text = rsTmp!年龄 & ""
            mstr入院时间 = Format(rsTmp!入院日期 & "", "YYYY-MM-DD HH:mm")
            mint险类 = Val(rsTmp!险类 & "")
        End If
    Else
        strSql = "Select A.姓名,A.性别,A.年龄,a.no,a.门诊号,a.险类,b.名称 as 科室" & _
            " From 病人挂号记录 A,部门表 b " & _
            " Where A.NO=[1] And a.记录性质=1 And a.记录状态=1 And A.病人ID+0=[2] and a.执行部门id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr挂号单, mlng病人ID)
        If rsTmp.RecordCount > 0 Then
            lblInfo(lbl挂号单).Caption = "挂 号 单"
            txtInfo(txt挂号单).Text = rsTmp!NO & ""
            txtInfo(txt姓名).Text = rsTmp!姓名 & ""
            txtInfo(txt性别).Text = rsTmp!性别 & ""
            If txtInfo(txt性别).Text = "男" Then
                SetControlEnabled optPregnancy(0), False
                SetControlEnabled optPregnancy(1), False
                SetControlEnabled optPregnancy(2), False
            End If
            txtInfo(txt科室).Text = rsTmp!科室 & ""
            lblInfo(lbl门诊号).Caption = "门 诊 号"
            txtInfo(txt门诊号).Text = rsTmp!门诊号 & ""
            txtInfo(txt年龄).Text = rsTmp!年龄 & ""
            mint险类 = Val(rsTmp!险类 & "")
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
    Dim strSql As String, i As Long, j As Long
    Dim str诊断 As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    '读取病人相关信息
    Call LoadPatiInfo

    If mintType = 0 Then
        If mint场合 = 0 Then
            '读取上次转科时间
            strSql = "Select 开始时间 From 病人变动记录" & _
                " Where 开始时间 is Not NULL And 开始原因=3" & _
                " And 病人ID=[1] And 主页ID=[2] Order by 开始时间 desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", mlng病人ID, mlng主页ID)
            If rsTmp.RecordCount > 0 Then
                mstr上次转科时间 = Format(rsTmp!开始时间 & "", "YYYY-MM-DD HH:mm")
            End If
        End If
    ElseIf mintType = 1 Or mintType = 3 Or mintType = 2 Then
        '修改
        '读取输血相关信息
        strSql = "Select 是否待诊, 输血性质, 即往输血史, 孕产情况, 受血者属地, 输血血型, Rhd, 受血者血型, Hct, Alt, Hbsag, 梅毒, 血红蛋白, 血小板, Antihcv, Antihiv12" & vbNewLine & _
                " From 输血申请记录" & vbNewLine & _
                " Where 医嘱id = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp!是否待诊 & "") = 1 Then
                txtInfo(txt诊断信息).Text = "待诊"
                chkWait.value = 1
            Else
               '读取诊断
               mstr诊断IDs = GetAdviceDiag(mlngUpdateAdvice, str诊断)
               txtInfo(txt诊断信息).Text = str诊断
            End If
            chkWait.value = Val(rsTmp!是否待诊 & "")
            cboInfo(cbo输血性质).ListIndex = Val(rsTmp!输血性质 & "")
            optHistory(Val(rsTmp!即往输血史 & "")).value = True
            optPregnancy(Val(rsTmp!孕产情况 & "")).value = True
            optPossession(Val(rsTmp!受血者属地 & "")).value = True
            cboInfo(cbo输血血型).ListIndex = Val(rsTmp!输血血型 & "")
            cboInfo(cboRhd).ListIndex = Val(rsTmp!Rhd & "")
        End If
        
        '读取医嘱相关信息
        strSql = "Select A.ID,A.相关ID,a.紧急标志,a.用药理由,NVL(to_char(a.手术时间,'yyyy-MM-dd hh24:mi'),a.标本部位) as 预定输血时间,a.开始执行时间,a.诊疗项目ID,a.执行科室ID,a.执行性质,a.总给予量,B.计算单位,B.名称 as 项目名称,A.申请序号,A.审核状态" & vbNewLine & _
                " From 病人医嘱记录 A,诊疗项目目录 B" & vbNewLine & _
                " Where a.诊疗项目ID=B.ID And (A.id = [1] or A.相关ID=[1])"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
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
                
                cboInfo(cbo输血目的).Text = rsTmp!用药理由 & ""
                txtInfo(txt预定输血时间).Text = Format(rsTmp!预定输血时间 & "", "YYYY-MM-DD HH:mm")
                txtInfo(txt申请日期).Text = Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD HH:mm")
                txtGet(txt预定输血成分).Text = rsTmp!项目名称 & ""
                txtInfo(txt单位).Text = rsTmp!计算单位 & ""
                txtGet(txt预定输血成分).Tag = txtGet(txt预定输血成分).Text
                mlng输血项目ID = Val(rsTmp!诊疗项目ID)
                
                Call Set执行科室(Val(rsTmp!执行性质 & ""), Val(rsTmp!执行科室ID & ""))
                Call LoadLisResult(mlngUpdateAdvice)
                
                txtInfo(txt预定输血量).Text = rsTmp!总给予量 & ""
                txtInfo(txtNO).Text = rsTmp!申请序号 & ""
                '已经审核通过的不允许修改
                If Val(rsTmp!审核状态 & "") = 2 Then mblnEditable = False
            End If
            rsTmp.Filter = "相关ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                txtGet(txt输血途径).Text = rsTmp!项目名称 & ""
                txtGet(txt输血途径).Tag = txtGet(txt输血途径).Text
                mlng输血途径 = Val(rsTmp!诊疗项目ID)
                Call Set输血执行(Val(rsTmp!执行性质 & ""), Val(rsTmp!执行科室ID & ""))
            End If
        End If
        '读取签名记录
        If gintCA <> 0 And Mid(gstrESign, 2, 1) = "1" Then
            strSql = "Select b.签名人,A.操作类型 From 病人医嘱状态 A, 医嘱签名记录 B Where a.签名id = b.Id And a.医嘱id = [1] And 操作类型=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
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
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
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
        cboInfo(cbo执行科室).Clear: cboInfo(cbo执行科室).AddItem "<院外执行>"
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
        cboInfo(cbo输血执行).Clear: cboInfo(cbo输血执行).AddItem "<院外执行>"
        cboInfo(cbo输血执行).ListIndex = 0
    Else
        If cboInfo(cbo输血执行).ListIndex >= 0 And lng执行科室ID = 0 Then
            lngTmp = cboInfo(cbo输血执行).ItemData(cboInfo(cbo输血执行).ListIndex)
        ElseIf lng执行科室ID <> 0 Then
            lngTmp = lng执行科室ID
        End If
        
        Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cboInfo(cbo输血执行), "E", mlng输血途径, 0, _
            lng执行科室, mlng病人科室id, mlng开单科室ID, lngTmp, 1, IIF(mlng病人性质 = 1, 1, 2))
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
    Dim strSql As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim strIDs As String, str医嘱内容 As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    

    If Index = txt预定输血成分 Then
        strSql = " And A.类别='K' "
    ElseIf Index = txt输血途径 Then
        strSql = " And A.类别='E' And A.操作类型='8' "
    End If
    strSql = "Select Distinct A.ID,A.编码,A.名称,A.执行分类 as 执行分类ID,A.计算单位,A.执行科室 as 执行科室ID,A.录入限量 as 录入限量ID" & _
        " From 诊疗项目目录 A,诊疗项目别名 B" & _
        " Where A.ID=B.诊疗项目ID" & _
        strSql & "  And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        IIF(intType = 0, " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])", "") & _
        " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[4])" & _
                " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
        Decode(gbytCode, 0, " And B.码类 IN([3],3)", 1, " And B.码类 IN([3],3)", "") & _
        " Order by A.编码"
    vRect = GetControlRect(txtGet(Index).hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtGet(Index).Height, blnCancel, False, True, UCase(txtGet(Index).Text) & "%", _
        gstrLike & UCase(txtGet(Index).Text) & "%", gbytCode + 1, mlng病人科室id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到匹配的项目。", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtGet(Index))
        txtGet(Index).SetFocus: Exit Sub
    Else
        txtGet(Index).Text = rsTmp!名称 & ""
        txtGet(Index).Tag = txtGet(Index).Text
        If Index = txt预定输血成分 Then
            mlng输血项目ID = Val(rsTmp!ID)
            mlng录入限量 = Val(rsTmp!录入限量ID & "")
            txtInfo(txt单位).Text = rsTmp!计算单位 & ""
            Call Set执行科室(Val(rsTmp!执行科室ID & ""))
            Call SetLisResult(mlng输血项目ID)
        ElseIf Index = txt输血途径 Then
            mlng输血途径 = Val(rsTmp!ID)
            Call Set输血执行(Val(rsTmp!执行科室ID & ""))
        End If
        '对码检查
        If mlng输血项目ID <> 0 Then
            strIDs = mlng输血项目ID & ":"
            If Val(cboInfo(cbo执行科室).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo执行科室).Tag & "")
            End If
            str医嘱内容 = FormatAdviceContext(txtGet(txt预定输血成分).Text, txtGet(txt输血途径).Text)
        End If
        If mlng输血途径 <> 0 Then
            strIDs = strIDs & "," & mlng输血途径 & ":"
            If Val(cboInfo(cbo输血执行).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo输血执行).Tag & "")
            End If
        End If
        
        strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, IIF(mlng病人性质 = 0, 2, 1), "", strIDs, str医嘱内容)
        If strMsg <> "" Then
            If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
            vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
            If vMsg = vbIgnore Then mbln提醒对码 = False
        End If
        
        txtGet(Index).SetFocus
        Call SeekNextControl
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub SetLisResult(ByVal lng输血项目ID As Long)
'功能：初始化输血项目对应的检验项目指标表格
    Dim rsTmp As Recordset, strSql As String
    Dim i As Long, j As Long
    Dim str检验编码 As String
    Dim strResult As String
    Dim rsLIS As Recordset '当前输血的检验项目
    Dim arrTmp1 As Variant
    Dim arrTmp2 As Variant
    Dim strTmp As String
    
    strSql = "select A.检验项目ID,B.编码 from 输血检验对照 A,诊疗项目目录 B Where A.检验项目ID=B.ID And A.项目ID=[1]"
    On Error GoTo errH
    Set rsLIS = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng输血项目ID)
    Do While Not rsLIS.EOF
        str检验编码 = str检验编码 & "," & rsLIS!编码
        rsLIS.MoveNext
    Loop
    str检验编码 = Mid(str检验编码, 2)
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
        
        If mint场合 = 0 And strTmp = "" Then
            If MsgBox("本次住院未找到有效的检验指标，是否提取历次就诊七天内的检验指标？", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes Then
                strResult = ""
                strResult = mobjPublicLis.GetTransfusionApplyFor(str检验编码, mlng病人ID, IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mstr挂号单, CInt(mbytBaby), 1)
                strTmp = strResult
                strTmp = Replace(strTmp, "<split1>", "")
                strTmp = Replace(strTmp, "<split2>", "")
                strTmp = Replace(strTmp, "<split3>", "")
                strTmp = Trim(strTmp)
            End If
        End If
        
        If strTmp <> "" Then
'            指标1<split1>诊疗编码1<split1>单位1<split1>隐私项目1<split1>指标代码1<split1>中文名1<split1>英文名1<split1>取值序列1<split1>
                '检验结果1<split2>结果标志1<split2>结果参数1<split2>排列序号1<split2>标本类型1<split3>
'            指标2<split1>诊疗编码2<split1>隐私项目2<split1>指标代码2<split1>中文名2<split1>英文名2<split1>取值序列2<split1>
              '  检验结果2<split2>结果标志2<split2>结果参数2<split2>排列序号2<split2>标本类型2<split3>
            arrTmp1 = Split(strResult, "<split3>")
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
                '加载指标结果
                If arrTmp2(8) <> "" Then
                    arrTmp2 = Split(arrTmp2(8), "<split2>")
                    .TextMatrix(Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(0)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_结果标志 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(1)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_结果参考 + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                Else
                    '未提取到结果表示可以医生录入
                    .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_指标结果 + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadLisResult(ByVal lng医嘱ID As Long)
'功能：修改\查看申请单时，传入医嘱ID，加载已填写的指标
    Dim rsTmp As Recordset, strSql As String
    Dim i As Long, j As Long

    strSql = "select 检验项目ID,指标代码,指标中文名,指标英文名,指标结果,结果单位,结果标志,结果参考,取值序列,是否人工填写 from 输血检验结果 Where 医嘱ID=[1] order by 序号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng医嘱ID)

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
    Set mclsMipModule = Nothing
End Sub

Private Sub optHistory_Click(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub optHistory_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub optPossession_Click(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub optPossession_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub optPregnancy_Click(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub optPregnancy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
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
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    If Index = txt预定输血时间 Then
        If txtInfo(Index).Text = "" Then txtInfo(Index).Text = txtInfo(txt申请日期).Text
        zlControl.TxtSelAll txtInfo(Index)
    ElseIf Index = txt申请日期 Then
        zlControl.TxtSelAll txtInfo(Index)
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
    End Select
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    If Index = txt预定输血时间 Then
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(txt申请日期).Text) Then
                    '恢复人为的清除缺省为开始时间
                    txtInfo(Index).Text = txtInfo(txt申请日期).Text
                End If
            End If
        Else
            '检查时间合法性
            If Not Check安排时间(txtInfo(Index).Text, txtInfo(txt申请日期).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
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
    ElseIf Index = txt预定输血量 Then
        If Val(txtInfo(Index).Text) > mlng录入限量 And mlng录入限量 > 0 Then
            If MsgBox(txtGet(txt预定输血成分).Text & " 的总量:" & Val(txtInfo(Index).Text) & txtInfo(txt单位).Text & " 超过允许录入的最大限量:" & _
                mlng录入限量 & txtInfo(txt单位).Text & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Cancel = True: txtInfo(Index).SetFocus: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub vsLIS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
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
                .ComboList = strTmp
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
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
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
