VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBloodApply 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ѫ���뵥"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkWait 
      BackColor       =   &H80000005&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
      Caption         =   "��"
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
      ToolTipText     =   "�༭(F4)"
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
      ToolTipText     =   "�༭(F4)"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   10
         Width           =   285
      End
      Begin VB.TextBox txtGet 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
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
         Name            =   "����"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   10
         Width           =   285
      End
      Begin VB.TextBox txtGet 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�Ⲻ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "����ҽʦǩ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�� �� ��ǩ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�� Ѫ ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫִ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ;��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ԥ����Ѫ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫִ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ԥ����Ѫ�ɷ�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "Ԥ����Ѫ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�в����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������Ѫʷ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ٴ���Ѫ���뵥"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ѪĿ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ٴ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ס Ժ ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ѫ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��λ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����ҽʦǩ��"
      BeginProperty Font 
         Name            =   "����"
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
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng�������� As Long   '0-סԺ��1-����
Private mblnChange As Boolean
Private mblnHaveAuditPriv As Boolean 'ִҵҽʦ�ʸ�
Private mintType As Integer   '0-������1-�޸ģ�2-�鿴,3-ҽ���༭���ã�ֻ�ܵ�������Ѫ�ɷ֣�����������ʱ�䣬��Ѫʱ�䣬ִ�п��ң���Ѫ;������Ѫִ�п��ң���Ѫ�������������
Private mlngUpdateAdvice As Long  '�޸ĵ�ҽ��ID
Private mintPState As Integer
Private mdatTurn As Date
Private mlng���˿���id As Long
Private mlng����ID As Long
Private mlng��������ID As Long
Private mlng��Ѫ;�� As Long
Private mlng��Ѫ��ĿID As Long
Private mstr��Ժʱ�� As String
Private mstr�ϴ�ת��ʱ�� As String
Private mrsDefine As Recordset
Private mobjVBA As Object
Private mobjScript As clsScript
Private mlngִ�п������� As Long
Private mlng��Ѫִ������ As Long
Private mbln��¼ As Boolean
Private mblnEditable As Boolean
Private mobjReport As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mstr���IDs As String  '��Ϲ���
Private mlng¼������ As Long
Private mint���뵥��ӡģʽ As Integer  '1-����ʱ��ӡ��2-�¿�ʱ��ӡ
Private mint���� As Integer '��ǰ��������
Private mbln���Ѷ��� As Boolean
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����
Attribute mclsMipModule.VB_VarHelpID = -1
Private Const CON_LisResultCol = 3
Private Const CON_LisResultCount = 10
Private mobjPublicLis As Object
Private mint���� As Integer '0-סԺ��1�����Ĭ��ΪסԺ
Private mstr�Һŵ� As String '�Һŵ���

Private Enum Enum_Cbo
    cbo��Ѫ���� = 0
    cbo��ѪѪ�� = 1
    cboRhd = 2
    cboִ�п��� = 8
    cbo��Ѫִ�� = 9
    cbo��Ѫ���� = 10
    cbo��ѪĿ�� = 11
End Enum

Private Enum Enum_lbl
    lbl���� = 5
    lbl����� = 5
    
    lblסԺ�� = 1
    lbl�Һŵ� = 1

End Enum

Private Enum Enum_txt
    txt�������� = 0
    
    txtסԺ�� = 1
    txt�Һŵ� = 1
    
    txt���� = 2
    txt�Ա� = 3
    
    txt���� = 4
    
    txt���� = 5
    txt����� = 5
    
    txtNO = 6
    txt���� = 7
    txt�����Ϣ = 8
    txtԤ����Ѫʱ�� = 10
    txtԤ����Ѫ�� = 11
    txt��λ = 12
    txt����ҽʦǩ�� = 17
    txt�������� = 19
    txt����ҽʦǩ�� = 20
End Enum

Private Enum Enum_Get
    txtԤ����Ѫ�ɷ� = 0
    txt��Ѫ;�� = 1
End Enum

Private Enum Enum_Date
    cmdԤ����Ѫʱ�� = 0
    cmd�������� = 1
End Enum

Private Enum Enum_Col
    COL_ָ�������� = 0
    COL_ָ���� = 1
    COL_�����λ = 2
    COL_ָ��Ӣ���� = 3
    COL_�����־ = 4
    COL_����ο� = 5
    COL_ȡֵ���� = 6
    COL_ָ����� = 7
    COL_������ĿID = 8
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�������� As Long, ByVal intType As Integer, Optional ByVal lngUpdateAdvice As Long, _
            Optional ByVal lng���˿���ID As Long, Optional ByVal lng����ID As Long, Optional ByVal lng��������ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
            Optional ByRef rsDefine As Recordset, Optional ByRef objMip As Object, Optional ByVal int���� As Integer, Optional ByVal str�Һŵ� As String) As Boolean
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng�������� = lng��������
    mlng���˿���id = lng���˿���ID
    mlng����ID = lng����ID
    mlng��������ID = lng��������ID
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mint���� = int����
    mstr�Һŵ� = str�Һŵ�
    Set mrsDefine = rsDefine
    mlngUpdateAdvice = lngUpdateAdvice
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    ShowMe = mblnOK
End Function

Private Function SeekNextControl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextControl = True
End Function

Private Sub cboInfo_Change(Index As Integer)
    If Visible And Index = cbo��ѪĿ�� Then mblnChange = True
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSql As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = cboִ�п��� Or Index = cbo��Ѫִ�� Then
        If cboInfo(Index).ItemData(cboInfo(Index).ListIndex) = -1 Then
            
            '����ִ�У�����ѡ��ִ�п���
            strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                IIF(gstrNodeNo <> "", " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " Order by A.����"
            vRect = GetControlRect(cboInfo(Index).hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "ִ�п���", , , , , , True, vRect.Left, vRect.Top, cboInfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = SeekCboIndex(cboInfo(Index), rsTmp!ID)
                If intIdx <> -1 Then
                    cboInfo(Index).ListIndex = intIdx
                Else
                    cboInfo(Index).AddItem rsTmp!���� & "-" & rsTmp!����, cboInfo(Index).ListCount - 1
                    cboInfo(Index).ItemData(cboInfo(Index).NewIndex) = rsTmp!ID
                    cboInfo(Index).ListIndex = cboInfo(Index).NewIndex
                End If
                If cboInfo(Index).ListIndex >= 0 Then
                    cboInfo(Index).Tag = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
                End If
                '�ָ������еĿ���(������Click)
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
'���ܣ�����ϵͳ������������ʽ��ҽ������
'������strBloodWay=��Ѫ;��,strAdvicePro=��Ѫ����
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
    mrsDefine.Filter = "�������='K'"
    If mrsDefine.RecordCount > 0 Then
        strReturn = mrsDefine!ҽ������ & ""
    End If
    If strReturn = "" Then
        If IsDate(txtInfo(txtԤ����Ѫʱ��).Text) Then
            strText = Format(txtInfo(txtԤ����Ѫʱ��).Text, "MM��dd��HH:mm")
        Else
            strText = Format(txtInfo(txt��������).Text, "MM��dd��HH:mm")
        End If
    
        strText = "��" & strText & "��" & strAdvicePro
        If strBloodWay <> "" Then
            strText = strText & "(" & strBloodWay & ")"
        End If
    Else
        strText = strReturn
        If InStr(strText, "[��Ѫʱ��]") > 0 Then
            If IsDate(txtInfo(txtԤ����Ѫʱ��).Text) Then
                strField = txtInfo(txtԤ����Ѫʱ��).Text
            Else
                strField = txtInfo(txt��������).Text
            End If
            strText = Replace(strText, "[��Ѫʱ��]", """" & strField & """")
        End If
        If InStr(strText, "[������Ŀ]") > 0 Then
            strField = strAdvicePro
            strText = Replace(strText, "[������Ŀ]", """" & strField & """")
        End If
        If InStr(strText, "[��Ѫ��Ŀ]") > 0 Then
            strField = strAdvicePro
            strText = Replace(strText, "[��Ѫ��Ŀ]", """" & strField & """")
        End If
        If InStr(strText, "[��Ѫ;��]") > 0 Then
            strField = strBloodWay
            strText = Replace(strText, "[��Ѫ;��]", """" & strField & """")
        End If
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    
'    Call SeekNextControl  '�����ַ�ʽ�������71290
'������������費ͬ�ؼ��Ľ��㣬ȷ��validata�¼���ִ�С�
    txtGet(txtԤ����Ѫ�ɷ�).SetFocus
    txtGet(txt��Ѫ;��).SetFocus
    
    '�༭��������������
    If mintType <> 3 Then
        '����¼����Ѫ�ɷ�
        If mlng��Ѫ��ĿID = 0 Then
            MsgBox "û��ȷ��Ԥ����Ѫ�ɷ֡�", vbInformation, Me.Caption
            If txtGet(txtԤ����Ѫ�ɷ�).Enabled Then txtGet(txtԤ����Ѫ�ɷ�).SetFocus
            Exit Function
        End If
        
        '���ִ�п���
        If cboInfo(cboִ�п���).Text = "" Then
            MsgBox "û��ȷ��ִ�п��ҡ�", vbInformation, Me.Caption
            If cboInfo(cboִ�п���).Enabled Then cboInfo(cboִ�п���).SetFocus
            Exit Function
        End If
        
        '�����Ѫ;������Ѫִ��
        If mlng��Ѫ;�� = 0 Then
            MsgBox "û��ָ����Ѫ;����", vbInformation, Me.Caption
            If txtGet(txt��Ѫ;��).Enabled Then txtGet(txt��Ѫ;��).SetFocus
            Exit Function
        End If
        If cboInfo(cbo��Ѫִ��).Text = "" Then
            MsgBox "û��ȷ����Ѫִ�п��ҡ�", vbInformation, Me.Caption
            If cboInfo(cbo��Ѫִ��).Enabled Then cboInfo(cbo��Ѫִ��).SetFocus
            Exit Function
        End If
        
        '����¼������
        If Val(txtInfo(txtԤ����Ѫ��).Text) <= 0 Then
            MsgBox "��¼�����0��Ԥ����Ѫ����", vbInformation, Me.Caption
            If txtInfo(txtԤ����Ѫ��).Enabled Then txtInfo(txtԤ����Ѫ��).SetFocus
            Exit Function
        End If
        
        '���ʱ��Ϸ���
        If Not Check��ʼʱ��(txtInfo(txt��������).Text) Then
            If txtInfo(txt��������).Enabled Then txtInfo(txt��������).SetFocus
            Exit Function
        End If
        If Not Check����ʱ��(txtInfo(txtԤ����Ѫʱ��).Text, txtInfo(txt��������).Text) Then
            If txtInfo(txtԤ����Ѫʱ��).Enabled Then txtInfo(txtԤ����Ѫʱ��).SetFocus
            Exit Function
        End If
        '������ҽ��������д��ѪĿ��
        If cboInfo(cbo��Ѫ����).ListIndex = 1 And cboInfo(cbo��ѪĿ��).Text = "" Then
            MsgBox "������Ѫ������д��ѪĿ�ġ�", vbInformation, Me.Caption
            If cboInfo(cbo��ѪĿ��).Enabled Then cboInfo(cbo��ѪĿ��).SetFocus
            Exit Function
        End If
        '������
        strIDs = mlng��Ѫ��ĿID & ":"
        If Val(cboInfo(cboִ�п���).Tag & "") <> 0 Then
            strIDs = strIDs & Val(cboInfo(cboִ�п���).Tag & "")
        End If
        strҽ������ = FormatAdviceContext(txtGet(txtԤ����Ѫ�ɷ�).Text, txtGet(txt��Ѫ;��).Text)
        strIDs = strIDs & "," & mlng��Ѫ;�� & ":"
        If Val(cboInfo(cbo��Ѫִ��).Tag & "") <> 0 Then
            strIDs = strIDs & Val(cboInfo(cbo��Ѫִ��).Tag & "")
        End If
        If gintҽ������ = 2 Then mbln���Ѷ��� = True
        
        strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, IIF(mlng�������� = 0, 2, 1), "", strIDs, strҽ������)
        If strMsg <> "" Then
            If gintҽ������ = 1 Then
                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", Me)
                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                If vMsg = vbIgnore Then mbln���Ѷ��� = False
            ElseIf gintҽ������ = 2 Then
                MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    CheckData = True
End Function

Private Function SaveData() As Boolean
'���ܣ���������
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lngҽ��ID As Long, lngҽ����� As Long, lng������� As Long
    Dim strSql As String, rsTmp As Recordset
    Dim str��Ŀ���� As String, str��Ѫ;�� As String
    Dim curDate As Date, i As Long, lng���ID As String, j As Long
    Dim lngCount As Long, int������Դ As Integer
    Dim strTmp��ҳID As String
    Dim strTmp�Һŵ� As String

    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    If mintType = 3 Then
        '���븽��༭ģʽ
        lng���ID = mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Delete(" & lng���ID & ")"
    Else
        
        lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")        '��ȡҽ��ID
        
        '����ҽ����¼.��ţ�����
        If mint���� = 0 Then
            lngҽ����� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, 0) + 1
            strTmp��ҳID = mlng��ҳID
            strTmp�Һŵ� = "NULL"
            int������Դ = 2
        Else
            lngҽ����� = GetMaxAdviceNO(mlng����ID, , 0, mstr�Һŵ�) + 1
            strTmp��ҳID = "NULL"
            strTmp�Һŵ� = "'" & mstr�Һŵ� & "'"
            int������Դ = 1
        End If
        
        str��Ŀ���� = Get��Ŀ����(mlng��Ѫ��ĿID)
        str��Ѫ;�� = Get��Ŀ����(mlng��Ѫ;��)
        If mlngUpdateAdvice <> 0 Then
            '�޸�ҽ����ɾ�������²���
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & mlngUpdateAdvice & ",1)"
            
            'ȡ�������
            strSql = "Select ������� From ����ҽ����¼ where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
            lng������� = Val(rsTmp!������� & "")
        End If
        If lng������� = 0 Then
            'ȡ�������
            strSql = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
            lng������� = Val(rsTmp!������� & "")
        End If
        '��Ѫҽ��
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngҽ��ID & ",NULL," & lngҽ����� & "," & int������Դ & "," & mlng����ID & "," & strTmp��ҳID & ",0,1,1,'K'," & mlng��Ѫ��ĿID & _
                                 ",NULL,NULL,NULL," & ZVal(txtInfo(txtԤ����Ѫ��).Text) & ",'" & FormatAdviceContext(str��Ŀ����, str��Ѫ;��) & _
                                 "',Null,'" & Format(txtInfo(txtԤ����Ѫʱ��).Text, "yyyy-MM-dd HH:mm:ss") & "','һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "Null", cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex)) & _
                                 "," & IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "5", mlngִ�п�������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp�Һŵ� & ",NULL,Null,0,NULL,NULL,'" & UserInfo.���� & "',Null,NULL,'" & cboInfo(cbo��ѪĿ��).Text & "'," & _
                                 IIF(gbln��Ѫ�ּ����� And cboInfo(cbo��Ѫ����).ListIndex <> 1, 1, IIF(gblnѪ��ϵͳ, 4, "NULL")) & "," & lng������� & ")"
        
        '��Ѫ;��
        lng���ID = lngҽ��ID
        lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")        '��ȡҽ��ID
        lngҽ����� = lngҽ����� + 1
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngҽ��ID & "," & lng���ID & "," & lngҽ����� & "," & int������Դ & "," & mlng����ID & "," & strTmp��ҳID & _
                                 ",0,1,1,'E'," & mlng��Ѫ;�� & ",NULL,NULL,NULL,Null,'" & str��Ѫ;�� & "',Null,NULL,'һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "Null", cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "5", mlng��Ѫִ������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp�Һŵ� & ",NULL,Null,0,NULL,NULL,'" & UserInfo.���� & "',Null,NULL,''," & _
                                 IIF(gbln��Ѫ�ּ����� And cboInfo(cbo��Ѫ����).ListIndex <> 1, 1, IIF(gblnѪ��ϵͳ, 4, "NULL")) & "," & lng������� & ")"
    End If
    '��Ѫ����������Ŀ
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_��Ѫ�����¼_Insert(" & lng���ID & "," & chkWait.value & "," & cboInfo(cbo��Ѫ����).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optPregnancy(0).value, 0, IIF(optPregnancy(1).value, 1, 2)) & "," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo��ѪѪ��).ListIndex & "," & cboInfo(cboRhd).ListIndex & ")"
    '������Ŀ
    With vsLIS
        lngCount = 0
        For i = 0 To .Rows - 1
            For j = 0 To CON_LisResultCol - 1
                If Val(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) <> 0 Then
                    lngCount = lngCount + 1
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_��Ѫ������_Insert(" & lng���ID & "," & lngCount & "," & ZVal(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) & ",'" & .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_ָ��Ӣ���� + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_�����λ + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_�����־ + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_����ο� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_ȡֵ���� + (j * CON_LisResultCount)) & "'," & IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0) & ")"
                End If
            Next
        Next
    End With
    
    '��Ϲ�����Ϣ
    If mstr���IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(" & lng���ID & ",'" & mstr���IDs & "')"
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mlngUpdateAdvice = lng���ID
    
    If Not (mclsMipModule Is Nothing) And mint���� = 0 Then
        If mclsMipModule.IsConnect Then
            Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, Trim(txtInfo(txt����).Text), Trim(txtInfo(txtסԺ��).Text), , IIF(mlng�������� = 1, 1, 2), _
                mlng��ҳID, mlng����ID, , mlng���˿���id, "", , Trim(txtInfo(txt����).Text), _
                lngҽ��ID, 0, 1, "K", "", UserInfo.����, Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , , "")
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
    If Index = cbo��ѪĿ�� Then
        If zlCommFun.ActualLen(cboInfo(Index).Text) > 50 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
    End If
End Sub

Private Sub PrintApply(ByVal intType As Integer)
'���ܴ�ӡԤ�����뵥
'������intType:1-Ԥ����2-��ӡ
    '�ж������δ�������ȱ����ٴ�ӡ
    
    If mblnChange Then
        If CheckData = False Then Exit Sub
        If SaveData() Then
            mblnOK = True
        End If
    Else
        '��������ã�����ҽ���Ƿ����
        If CheckData = False Then Exit Sub
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_17", Me, "ҽ��ID=" & mlngUpdateAdvice, intType)
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_17", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '����
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
            blnVisible = (mint���뵥��ӡģʽ = 2 And InStr(GetInsidePrivs(pסԺҽ���´�), ";��Ѫ���뵥;") > 0) Or mint���� = 1
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub chkWait_Click()
    If chkWait.value = 1 Then
        txtInfo(txt�����Ϣ).Text = "����"
        cmdInfo.Enabled = False
        mstr���IDs = ""
    Else
        txtInfo(txt�����Ϣ).Text = ""
        cmdInfo.Enabled = True
    End If
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Dim lngIndex As Long
    
    If Index = 0 Then
        lngIndex = txtԤ����Ѫʱ��
    ElseIf Index = 1 Then
        lngIndex = txt��������
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
    Dim str��� As String

    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng�������� = 1, 1260, 1261), mclsMipModule)
    End If
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 2), mlng���˿���id, mstr���IDs, str���, 0, mlngUpdateAdvice) Then
        txtInfo(txt�����Ϣ).Text = str���
        If mstr���IDs <> "" And chkWait.value = 1 Then
            chkWait.value = 0
        End If
    End If
    Call SeekNextControl
End Sub

Private Sub cmdInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdInfo_Click
End Sub

Private Function Check��ʼʱ��(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵���Ժʱ��
'2.��ʼʱ�䲻��С�ڲ��˵�ת��ʱ��
'3.��ʼʱ�����С����ֹʱ��
'4.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
'5.��¼��ҽ����ʼʱ�䲻�ܴ��ڵ�ǰʱ�䣬ת�Ʋ�¼���ܴ���ת�ƿ�ʼʱ��
    Dim strInDate As String, blnOut As Boolean
        
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'סԺ���ϵ���ʱ�������¼��
    If mint���� = 0 Then
        strInDate = mstr��Ժʱ��
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    
    
        strInDate = ""
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
        ElseIf IsDate(mstr�ϴ�ת��ʱ��) Then
            strInDate = mstr�ϴ�ת��ʱ��
        End If
    
        If strInDate <> "" Then
            If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
                If Format(strStart, "yyyy-MM-dd HH:mm") >= strInDate Then
                    strMsg = "ҽ���Ŀ�ʼִ��ʱ��ӦС�ڲ���" & IIF(mintPState = ps���ת��, "ת��", IIF(mintPState = psԤ��, "Ԥ��Ժ", "��Ժ")) & "��ʱ�� " & strInDate & " ��"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                    strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��������ת��ʱ�� " & strInDate & " ��"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    Check��ʼʱ�� = True
End Function

Private Function Check����ʱ��(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ�����������Ѫʱ���Ƿ�Ϸ�
'˵����
'1.��Ѫʱ�䲻��С��ҽ���Ŀ�ʼʱ��
    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "�������Ѫʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "��Ѫʱ�䲻��С��ҽ����ʼʱ�䡣"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check����ʱ�� = True
End Function

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String, intIndex As Integer
    
    intIndex = Val(dtpDate.Tag)
    If intIndex = txt�������� Then
        'ȡֵ
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check��ʼʱ��(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '��������
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    ElseIf intIndex = txtԤ����Ѫʱ�� Then
        'ȡֵ
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check����ʱ��(strDate, txtInfo(txt��������).Text) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '��������
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
    mstr���IDs = ""
    mblnOK = False
    mbln���Ѷ��� = True
    vsLIS.Rows = 0
    If mint���� = 0 Then mint���뵥��ӡģʽ = Val(zlDatabase.GetPara("��Ѫ���뵥��ӡģʽ", glngSys, pסԺҽ������, "1"))
    If mobjPublicLis Is Nothing Then
        On Error Resume Next
        Set mobjPublicLis = CreateObject("zlPublicLIS.clsSampleReprot")
        err.Clear: On Error GoTo 0
        If Not mobjPublicLis Is Nothing Then
            Call mobjPublicLis.InitSampleReprot(gcnOracle, glngSys, pסԺҽ��վ, "")
        End If
    End If
    If mintType = 2 Then
        picNo.Visible = True
        mblnEditable = False
    ElseIf mintType = 1 Then
        '�޸�ʱ�����������ʼִ��ʱ�䣬�����ǲ�¼ҽ��
        SetControlEnabled txtInfo(txt��������), False
        SetControlEnabled cmdDate(cmd��������), False
    ElseIf mintType = 3 Then
        'ֻ�ܵ�������Ѫ�ɷ֣�����������ʱ�䣬��Ѫʱ�䣬ִ�п��ң���Ѫ;������Ѫִ�п��ң���Ѫ�������������
        SetControlEnabled txtInfo(txt��������), False
        SetControlEnabled cmdDate(cmd��������), False
        SetControlEnabled txtInfo(txtԤ����Ѫʱ��), False
        SetControlEnabled cmdDate(cmdԤ����Ѫʱ��), False
        SetControlEnabled txtInfo(txtԤ����Ѫ�ɷ�), False
        SetControlEnabled txtGet(txtԤ����Ѫ�ɷ�), False
        SetControlEnabled cmdGet(txtԤ����Ѫ�ɷ�), False
        SetControlEnabled txtGet(txt��Ѫ;��), False
        SetControlEnabled cmdGet(txt��Ѫ;��), False
        SetControlEnabled txtInfo(txtԤ����Ѫ��), False
        SetControlEnabled cboInfo(cboִ�п���), False
        SetControlEnabled cboInfo(cbo��Ѫִ��), False
        SetControlEnabled cboInfo(cbo��Ѫ����), False
        SetControlEnabled cboInfo(cbo��ѪĿ��), False
    End If
    mblnChange = False
    Call InitCommandBar
    If InitInfo = False Then Exit Sub
    Call LoadData
    Call SetFaceEnabledFalse
    If mbln��¼ Then SetControlEnabled cboInfo(cbo��Ѫ����), False
    '���˻�����Ϣ�����Ա༭
    SetControlEnabled txtInfo(txt�Ա�), False
    SetControlEnabled txtInfo(txt����), False
    SetControlEnabled txtInfo(txt����), False
End Sub

Private Sub SetFaceEnabledFalse()
'���ܣ�����˲������޸�,��ǩ���Ĳ������޸�
    Dim objControl As Object
    If mblnEditable = False Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
    End If
End Sub

Private Sub SetControlEnabled(objControl As Object, ByVal blnEnabled As Boolean)
'���ܣ����ÿؼ��Ŀ�����
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
'���ܣ���ʼ�����˵�
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim curDate As Date
    Dim lng�÷�ID As Long
    Dim lngִ�п���ID As Long
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '���̶ֹ����ݵ�������
    Call SetCboFromList(cboInfo, Array(" ", "A", "B", "O", "AB"), Array(cbo��ѪѪ��), 0)
    Call SetCboFromList(cboInfo, Array(" ", "����", "����", "����", "����"), Array(cbo��Ѫ����), 1)
    Call SetCboFromList(cboInfo, Array("��ͨ", "����"), Array(cbo��Ѫ����), 0)
    Call SetCboFromList(cboInfo, Array(" ", "-", "+"), Array(cboRhd), 0)
    
    strSql = "select ���� from ��ѪĿ�� order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        With cboInfo(cbo��ѪĿ��)
            .Clear
            For i = 1 To rsTmp.RecordCount
                .AddItem rsTmp!���� & ""
                rsTmp.MoveNext
            Next
        End With
    End If
    Set rsTmp = Nothing
    
    '����
    curDate = zlDatabase.Currentdate
    If mint���� = 0 Then 'ֻ��סԺ���ϲ��в�¼
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            If mdatTurn <> CDate(0) Then curDate = mdatTurn - 1 / 24 / 60
            mbln��¼ = True
        End If
    Else
        mbln��¼ = False
    End If
    
    txtInfo(txtԤ����Ѫʱ��).Text = Format(curDate, "YYYY-MM-DD HH:mm")
    txtInfo(txt��������).Text = Format(curDate, "YYYY-MM-DD HH:mm")
    txtInfo(txt��������).Tag = txtInfo(txt��������).Text
    
    'ȱʡ�÷�
    lng�÷�ID = Getȱʡ�÷�ID(8, IIF(mint���� = 0, 2, 1))
    
    If lng�÷�ID = 0 Then
        MsgBox "û�п��õ���Ѫ;��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
        Screen.MousePointer = 0
        Unload Me
        Exit Function
    Else
        Set rsTmp = Get������Ŀ��¼(lng�÷�ID)
        txtGet(txt��Ѫ;��).Text = rsTmp!���� & ""
        mlng��Ѫִ������ = Nvl(rsTmp!ִ�п���, 0)
        txtGet(txt��Ѫ;��).Tag = txtGet(txt��Ѫ;��).Text
        mlng��Ѫ;�� = lng�÷�ID
        cboInfo(cbo��Ѫִ��).Enabled = True
        Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(cbo��Ѫִ��), "E", mlng��Ѫ;��, 0, _
            Val(rsTmp!ִ�п��� & ""), mlng���˿���id, mlng��������ID, 0, 1, IIF(mlng�������� = 1, 1, 2))
        If cboInfo(cbo��Ѫִ��).ListIndex = -1 And cboInfo(cbo��Ѫִ��).ListCount > 1 Then
'            cboInfo(cbo��Ѫִ��).ListIndex = 0
            Call zlControl.CboSetIndex(cboInfo(cbo��Ѫִ��).hWnd, 0)
        Else
            '����ж����ȡĬ�ϵ�ִ�п���
            lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "E", mlng��Ѫ;��, 0, _
                    Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
            If lngִ�п���ID <> 0 Then
                Call zlControl.CboLocate(cboInfo(cbo��Ѫִ��), lngִ�п���ID, True)
            End If
        End If
        If cboInfo(cbo��Ѫִ��).ListCount = 2 Then cboInfo(cbo��Ѫִ��).Enabled = False
        cboInfo(cbo��Ѫִ��).Tag = lng�÷�ID
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
'���ܣ���ȡ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    
    '��ȡ���������Ϣ
    txtInfo(txt��������).Text = IIF(mlng�������� = 1, "����", "סԺ")
    
    If mint���� = 0 Then
        strSql = "Select A.סԺ��, Nvl(C.����, A.����) ����, Nvl(C.�Ա�, A.�Ա�) �Ա�, Nvl(C.����, A.����) ����, B.���� As ����, C.��Ժ���� As ��ǰ����, C.��Ժ����, C.����" & vbNewLine & _
                "From ������Ϣ A, ���ű� B, ������ҳ C" & vbNewLine & _
                "Where C.��Ժ����id = B.Id And A.����id = C.����id And A.��ҳid = C.��ҳid And C.����id = [1] And C.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.RecordCount > 0 Then
            txtInfo(txtסԺ��).Text = rsTmp!סԺ�� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
            If txtInfo(txt�Ա�).Text = "��" Then
                SetControlEnabled optPregnancy(0), False
                SetControlEnabled optPregnancy(1), False
                SetControlEnabled optPregnancy(2), False
            End If
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt����).Text = rsTmp!��ǰ���� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            mstr��Ժʱ�� = Format(rsTmp!��Ժ���� & "", "YYYY-MM-DD HH:mm")
            mint���� = Val(rsTmp!���� & "")
        End If
    Else
        strSql = "Select A.����,A.�Ա�,A.����,a.no,a.�����,a.����,b.���� as ����" & _
            " From ���˹Һż�¼ A,���ű� b " & _
            " Where A.NO=[1] And a.��¼����=1 And a.��¼״̬=1 And A.����ID+0=[2] and a.ִ�в���id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr�Һŵ�, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            lblInfo(lbl�Һŵ�).Caption = "�� �� ��"
            txtInfo(txt�Һŵ�).Text = rsTmp!NO & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
            If txtInfo(txt�Ա�).Text = "��" Then
                SetControlEnabled optPregnancy(0), False
                SetControlEnabled optPregnancy(1), False
                SetControlEnabled optPregnancy(2), False
            End If
            txtInfo(txt����).Text = rsTmp!���� & ""
            lblInfo(lbl�����).Caption = "�� �� ��"
            txtInfo(txt�����).Text = rsTmp!����� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            mint���� = Val(rsTmp!���� & "")
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
'���ܣ���ȡ���뵥��Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim str��� As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    '��ȡ���������Ϣ
    Call LoadPatiInfo

    If mintType = 0 Then
        If mint���� = 0 Then
            '��ȡ�ϴ�ת��ʱ��
            strSql = "Select ��ʼʱ�� From ���˱䶯��¼" & _
                " Where ��ʼʱ�� is Not NULL And ��ʼԭ��=3" & _
                " And ����ID=[1] And ��ҳID=[2] Order by ��ʼʱ�� desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", mlng����ID, mlng��ҳID)
            If rsTmp.RecordCount > 0 Then
                mstr�ϴ�ת��ʱ�� = Format(rsTmp!��ʼʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
        End If
    ElseIf mintType = 1 Or mintType = 3 Or mintType = 2 Then
        '�޸�
        '��ȡ��Ѫ�����Ϣ
        strSql = "Select �Ƿ����, ��Ѫ����, ������Ѫʷ, �в����, ��Ѫ������, ��ѪѪ��, Rhd, ��Ѫ��Ѫ��, Hct, Alt, Hbsag, ÷��, Ѫ�쵰��, ѪС��, Antihcv, Antihiv12" & vbNewLine & _
                " From ��Ѫ�����¼" & vbNewLine & _
                " Where ҽ��id = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp!�Ƿ���� & "") = 1 Then
                txtInfo(txt�����Ϣ).Text = "����"
                chkWait.value = 1
            Else
               '��ȡ���
               mstr���IDs = GetAdviceDiag(mlngUpdateAdvice, str���)
               txtInfo(txt�����Ϣ).Text = str���
            End If
            chkWait.value = Val(rsTmp!�Ƿ���� & "")
            cboInfo(cbo��Ѫ����).ListIndex = Val(rsTmp!��Ѫ���� & "")
            optHistory(Val(rsTmp!������Ѫʷ & "")).value = True
            optPregnancy(Val(rsTmp!�в���� & "")).value = True
            optPossession(Val(rsTmp!��Ѫ������ & "")).value = True
            cboInfo(cbo��ѪѪ��).ListIndex = Val(rsTmp!��ѪѪ�� & "")
            cboInfo(cboRhd).ListIndex = Val(rsTmp!Rhd & "")
        End If
        
        '��ȡҽ�������Ϣ
        strSql = "Select A.ID,A.���ID,a.������־,a.��ҩ����,NVL(to_char(a.����ʱ��,'yyyy-MM-dd hh24:mi'),a.�걾��λ) as Ԥ����Ѫʱ��,a.��ʼִ��ʱ��,a.������ĿID,a.ִ�п���ID,a.ִ������,a.�ܸ�����,B.���㵥λ,B.���� as ��Ŀ����,A.�������,A.���״̬" & vbNewLine & _
                " From ����ҽ����¼ A,������ĿĿ¼ B" & vbNewLine & _
                " Where a.������ĿID=B.ID And (A.id = [1] or A.���ID=[1])"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!������־ & "") = 2 Then
                    mbln��¼ = True
                    SetControlEnabled txtInfo(txt��������), True
                    SetControlEnabled cmdDate(cmd��������), True
                ElseIf Val(rsTmp!������־ & "") = 1 Then
                    cboInfo(cbo��Ѫ����).ListIndex = 1
                End If
                
                cboInfo(cbo��ѪĿ��).Text = rsTmp!��ҩ���� & ""
                txtInfo(txtԤ����Ѫʱ��).Text = Format(rsTmp!Ԥ����Ѫʱ�� & "", "YYYY-MM-DD HH:mm")
                txtInfo(txt��������).Text = Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:mm")
                txtGet(txtԤ����Ѫ�ɷ�).Text = rsTmp!��Ŀ���� & ""
                txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
                txtGet(txtԤ����Ѫ�ɷ�).Tag = txtGet(txtԤ����Ѫ�ɷ�).Text
                mlng��Ѫ��ĿID = Val(rsTmp!������ĿID)
                
                Call Setִ�п���(Val(rsTmp!ִ������ & ""), Val(rsTmp!ִ�п���ID & ""))
                Call LoadLisResult(mlngUpdateAdvice)
                
                txtInfo(txtԤ����Ѫ��).Text = rsTmp!�ܸ����� & ""
                txtInfo(txtNO).Text = rsTmp!������� & ""
                '�Ѿ����ͨ���Ĳ������޸�
                If Val(rsTmp!���״̬ & "") = 2 Then mblnEditable = False
            End If
            rsTmp.Filter = "���ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                txtGet(txt��Ѫ;��).Text = rsTmp!��Ŀ���� & ""
                txtGet(txt��Ѫ;��).Tag = txtGet(txt��Ѫ;��).Text
                mlng��Ѫ;�� = Val(rsTmp!������ĿID)
                Call Set��Ѫִ��(Val(rsTmp!ִ������ & ""), Val(rsTmp!ִ�п���ID & ""))
            End If
        End If
        '��ȡǩ����¼
        If gintCA <> 0 And Mid(gstrESign, 2, 1) = "1" Then
            strSql = "Select b.ǩ����,A.�������� From ����ҽ��״̬ A, ҽ��ǩ����¼ B Where a.ǩ��id = b.Id And a.ҽ��id = [1] And ��������=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUpdateAdvice)
            If rsTmp.RecordCount > 0 Then
                mblnEditable = False
                'ǩ����
                rsTmp.Filter = "��������=1"
                If rsTmp.RecordCount > 0 Then
                    txtInfo(txt����ҽʦǩ��).Text = rsTmp!ǩ���� & ""
                End If
                '����ˣ������δ����ǩ�����ܣ���Ҫ���ǵ���˵��˲�һ����ǩ����U��)
'                rsTmp.Filter = "��������=11"
'                If rsTmp.RecordCount > 0 Then
'                    txtInfo(txt����ҽʦǩ��).Text = rsTmp!ǩ���� & ""
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
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ����(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " �����˳�(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&X)"): objControl.BeginGroup = True
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

Private Sub Setִ�п���(ByVal lngִ�п��� As Long, Optional ByVal lngִ�п���ID As Long)
'���ܣ�����ִ�п���
'������lngִ�п���-ִ�����ʣ�lngִ�п���ID=������룬���ʾ���ô�ִ�п���Ϊ��ǰִ�п���
    Dim lngTmp As Long
 
    cboInfo(cboִ�п���).Enabled = True
    If lngִ�п��� = 5 Then
        cboInfo(cboִ�п���).Clear: cboInfo(cboִ�п���).AddItem "<Ժ��ִ��>"
        cboInfo(cboִ�п���).ListIndex = 0
    Else
        If cboInfo(cboִ�п���).ListIndex >= 0 And lngִ�п���ID = 0 Then
            lngTmp = cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex)
        ElseIf lngִ�п���ID <> 0 Then
            lngTmp = lngִ�п���ID
        End If
        
        Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(cboִ�п���), "K", mlng��Ѫ��ĿID, 0, lngִ�п���, mlng���˿���id, mlng��������ID, lngTmp, 1, IIF(mlng�������� = 1, 1, 2))
        If lngִ�п���ID = 0 Then
            If cboInfo(cboִ�п���).ListIndex = -1 And cboInfo(cboִ�п���).ListCount = 1 Then
                cboInfo(cboִ�п���).ListIndex = 0
            Else
                 '����ж����ȡĬ�ϵ�ִ�п���
                lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "K", mlng��Ѫ��ĿID, 0, _
                        lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
            End If
        End If
        If lngִ�п���ID <> 0 Then
            Call zlControl.CboLocate(cboInfo(cboִ�п���), lngִ�п���ID, True)
        End If
    End If
    mlngִ�п������� = lngִ�п���
    If cboInfo(cboִ�п���).ListCount = 1 Then cboInfo(cboִ�п���).Enabled = False
    If cboInfo(cboִ�п���).ListIndex >= 0 Then
        cboInfo(cboִ�п���).Tag = cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex)
    End If
End Sub

Private Sub Set��Ѫִ��(ByVal lngִ�п��� As Long, Optional ByVal lngִ�п���ID As Long)
'���ܣ�������Ѫִ�п���
'������lngִ�п���-ִ�����ʣ�lngִ�п���ID=������룬���ʾ���ô�ִ�п���Ϊ��ǰִ�п���
    Dim lngTmp As Long
    
    cboInfo(cbo��Ѫִ��).Enabled = True
    If lngִ�п��� = 5 Then
        cboInfo(cbo��Ѫִ��).Clear: cboInfo(cbo��Ѫִ��).AddItem "<Ժ��ִ��>"
        cboInfo(cbo��Ѫִ��).ListIndex = 0
    Else
        If cboInfo(cbo��Ѫִ��).ListIndex >= 0 And lngִ�п���ID = 0 Then
            lngTmp = cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)
        ElseIf lngִ�п���ID <> 0 Then
            lngTmp = lngִ�п���ID
        End If
        
        Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(cbo��Ѫִ��), "E", mlng��Ѫ;��, 0, _
            lngִ�п���, mlng���˿���id, mlng��������ID, lngTmp, 1, IIF(mlng�������� = 1, 1, 2))
        If lngִ�п���ID = 0 Then
            If cboInfo(cbo��Ѫִ��).ListIndex = -1 And cboInfo(cbo��Ѫִ��).ListCount = 1 Then
                cboInfo(cbo��Ѫִ��).ListIndex = 0
            Else
                 '����ж����ȡĬ�ϵ�ִ�п���
                lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "E", mlng��Ѫ;��, 0, _
                        lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
            End If
        End If
        If lngִ�п���ID <> 0 Then
            Call zlControl.CboLocate(cboInfo(cbo��Ѫִ��), lngִ�п���ID, True)
        End If
    End If
    mlng��Ѫִ������ = lngִ�п���
    If cboInfo(cbo��Ѫִ��).ListCount = 1 Then cboInfo(cbo��Ѫִ��).Enabled = False
    If cboInfo(cbo��Ѫִ��).ListIndex >= 0 Then
        cboInfo(cbo��Ѫִ��).Tag = cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)
    End If
End Sub

Private Sub TxtGetInfo(Index As Integer, Optional ByVal intType As Integer)
'���ܣ������ı�������
'������intType =0 KeyPress���ã�=1 ������ť����
    Dim strSql As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    

    If Index = txtԤ����Ѫ�ɷ� Then
        strSql = " And A.���='K' "
    ElseIf Index = txt��Ѫ;�� Then
        strSql = " And A.���='E' And A.��������='8' "
    End If
    strSql = "Select Distinct A.ID,A.����,A.����,A.ִ�з��� as ִ�з���ID,A.���㵥λ,A.ִ�п��� as ִ�п���ID,A.¼������ as ¼������ID" & _
        " From ������ĿĿ¼ A,������Ŀ���� B" & _
        " Where A.ID=B.������ĿID" & _
        strSql & "  And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        IIF(intType = 0, " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])", "") & _
        " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4])" & _
                " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
        Decode(gbytCode, 0, " And B.���� IN([3],3)", 1, " And B.���� IN([3],3)", "") & _
        " Order by A.����"
    vRect = GetControlRect(txtGet(Index).hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtGet(Index).Height, blnCancel, False, True, UCase(txtGet(Index).Text) & "%", _
        gstrLike & UCase(txtGet(Index).Text) & "%", gbytCode + 1, mlng���˿���id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ�ƥ�����Ŀ��", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtGet(Index))
        txtGet(Index).SetFocus: Exit Sub
    Else
        txtGet(Index).Text = rsTmp!���� & ""
        txtGet(Index).Tag = txtGet(Index).Text
        If Index = txtԤ����Ѫ�ɷ� Then
            mlng��Ѫ��ĿID = Val(rsTmp!ID)
            mlng¼������ = Val(rsTmp!¼������ID & "")
            txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
            Call Setִ�п���(Val(rsTmp!ִ�п���ID & ""))
            Call SetLisResult(mlng��Ѫ��ĿID)
        ElseIf Index = txt��Ѫ;�� Then
            mlng��Ѫ;�� = Val(rsTmp!ID)
            Call Set��Ѫִ��(Val(rsTmp!ִ�п���ID & ""))
        End If
        '������
        If mlng��Ѫ��ĿID <> 0 Then
            strIDs = mlng��Ѫ��ĿID & ":"
            If Val(cboInfo(cboִ�п���).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cboִ�п���).Tag & "")
            End If
            strҽ������ = FormatAdviceContext(txtGet(txtԤ����Ѫ�ɷ�).Text, txtGet(txt��Ѫ;��).Text)
        End If
        If mlng��Ѫ;�� <> 0 Then
            strIDs = strIDs & "," & mlng��Ѫ;�� & ":"
            If Val(cboInfo(cbo��Ѫִ��).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo��Ѫִ��).Tag & "")
            End If
        End If
        
        strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, IIF(mlng�������� = 0, 2, 1), "", strIDs, strҽ������)
        If strMsg <> "" Then
            If gintҽ������ = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档"
            vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
            If vMsg = vbIgnore Then mbln���Ѷ��� = False
        End If
        
        txtGet(Index).SetFocus
        Call SeekNextControl
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub SetLisResult(ByVal lng��Ѫ��ĿID As Long)
'���ܣ���ʼ����Ѫ��Ŀ��Ӧ�ļ�����Ŀָ����
    Dim rsTmp As Recordset, strSql As String
    Dim i As Long, j As Long
    Dim str������� As String
    Dim strResult As String
    Dim rsLIS As Recordset '��ǰ��Ѫ�ļ�����Ŀ
    Dim arrTmp1 As Variant
    Dim arrTmp2 As Variant
    Dim strTmp As String
    
    strSql = "select A.������ĿID,B.���� from ��Ѫ������� A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.��ĿID=[1]"
    On Error GoTo errH
    Set rsLIS = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��Ѫ��ĿID)
    Do While Not rsLIS.EOF
        str������� = str������� & "," & rsLIS!����
        rsLIS.MoveNext
    Loop
    str������� = Mid(str�������, 2)
    With vsLIS
        .Clear
        .Rows = 0
        If str������� = "" Then Exit Sub

        strResult = mobjPublicLis.GetTransfusionApplyFor(str�������, mlng����ID, IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mstr�Һŵ�, CInt(mbytBaby), 0)
        strTmp = strResult
        strTmp = Replace(strTmp, "<split1>", "")
        strTmp = Replace(strTmp, "<split2>", "")
        strTmp = Replace(strTmp, "<split3>", "")
        strTmp = Trim(strTmp)
        
        If mint���� = 0 And strTmp = "" Then
            If MsgBox("����סԺδ�ҵ���Ч�ļ���ָ�꣬�Ƿ���ȡ���ξ��������ڵļ���ָ�ꣿ", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes Then
                strResult = ""
                strResult = mobjPublicLis.GetTransfusionApplyFor(str�������, mlng����ID, IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mstr�Һŵ�, CInt(mbytBaby), 1)
                strTmp = strResult
                strTmp = Replace(strTmp, "<split1>", "")
                strTmp = Replace(strTmp, "<split2>", "")
                strTmp = Replace(strTmp, "<split3>", "")
                strTmp = Trim(strTmp)
            End If
        End If
        
        If strTmp <> "" Then
'            ָ��1<split1>���Ʊ���1<split1>��λ1<split1>��˽��Ŀ1<split1>ָ�����1<split1>������1<split1>Ӣ����1<split1>ȡֵ����1<split1>
                '������1<split2>�����־1<split2>�������1<split2>�������1<split2>�걾����1<split3>
'            ָ��2<split1>���Ʊ���2<split1>��˽��Ŀ2<split1>ָ�����2<split1>������2<split1>Ӣ����2<split1>ȡֵ����2<split1>
              '  ������2<split2>�����־2<split2>�������2<split2>�������2<split2>�걾����2<split3>
            arrTmp1 = Split(strResult, "<split3>")
            .Rows = Int((UBound(arrTmp1) + 1) / CON_LisResultCol) + IIF((UBound(arrTmp1) + 1) Mod CON_LisResultCol = 0, 0, 1)
            For i = 0 To UBound(arrTmp1)
                '����ָ��
                arrTmp2 = Split(arrTmp1(i), "<split1>")
                .TextMatrix(Int(i / CON_LisResultCol), COL_ָ�������� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(5)
                .TextMatrix(Int(i / CON_LisResultCol), COL_�����λ + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                .TextMatrix(Int(i / CON_LisResultCol), COL_ָ��Ӣ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(6)
                .TextMatrix(Int(i / CON_LisResultCol), COL_ȡֵ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(7)
                .TextMatrix(Int(i / CON_LisResultCol), COL_ָ����� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(4)
                rsLIS.Filter = "����='" & arrTmp2(1) & "'"
                If rsLIS.RecordCount > 0 Then
                    .TextMatrix(Int(i / CON_LisResultCol), COL_������ĿID + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsLIS!������ĿID & ""
                End If
                '����ָ����
                If arrTmp2(8) <> "" Then
                    arrTmp2 = Split(arrTmp2(8), "<split2>")
                    .TextMatrix(Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(0)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_�����־ + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(1)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_����ο� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                Else
                    'δ��ȡ�������ʾ����ҽ��¼��
                    .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadLisResult(ByVal lngҽ��ID As Long)
'���ܣ��޸�\�鿴���뵥ʱ������ҽ��ID����������д��ָ��
    Dim rsTmp As Recordset, strSql As String
    Dim i As Long, j As Long

    strSql = "select ������ĿID,ָ�����,ָ��������,ָ��Ӣ����,ָ����,�����λ,�����־,����ο�,ȡֵ����,�Ƿ��˹���д from ��Ѫ������ Where ҽ��ID=[1] order by ���"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID)

    With vsLIS
        .Clear
        .Rows = Int((rsTmp.RecordCount) / CON_LisResultCol) + IIF((rsTmp.RecordCount) Mod CON_LisResultCol = 0, 0, 1)
        For i = 0 To rsTmp.RecordCount - 1
            '����ָ��
            .TextMatrix(Int(i / CON_LisResultCol), COL_ָ�������� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!ָ�������� & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_�����λ + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!�����λ & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_ָ��Ӣ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!ָ��Ӣ���� & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_ȡֵ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!ȡֵ���� & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_ָ����� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!ָ����� & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_������ĿID + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!������ĿID & ""
            '����ָ����
            .TextMatrix(Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!ָ���� & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_�����־ + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!�����־ & ""
            .TextMatrix(Int(i / CON_LisResultCol), COL_����ο� + (i Mod CON_LisResultCol) * CON_LisResultCount) = rsTmp!����ο� & ""

            '�ֹ�¼��Ŀ����޸�
            If Val(rsTmp!�Ƿ��˹���д & "") = 1 Then
                .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
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
        If MsgBox("��ǰ���뵥�Ѿ������˵�����δ���棬�Ƿ�Ҫ�����˳���", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    mlng��Ѫ;�� = 0
    mlng��Ѫ��ĿID = 0
    mlng��Ѫִ������ = 0
    mlngִ�п������� = 0
    mbln��¼ = False
    mstr��Ժʱ�� = ""
    mlng¼������ = 0
    mstr�ϴ�ת��ʱ�� = ""
    mint���� = 0
    mstr���IDs = ""
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
    '�ָ���Ϊ�����
    If txtGet(Index).Text <> txtGet(Index).Tag Then
        txtGet(Index).Text = txtGet(Index).Tag
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    If Index = txtԤ����Ѫʱ�� Then
        If txtInfo(Index).Text = "" Then txtInfo(Index).Text = txtInfo(txt��������).Text
        zlControl.TxtSelAll txtInfo(Index)
    ElseIf Index = txt�������� Then
        zlControl.TxtSelAll txtInfo(Index)
    End If
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''''
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case txtԤ����Ѫʱ��
            Call cmdDate_Click(0)
        Case txtԤ����Ѫʱ��
            Call cmdDate_Click(1)
        End Select
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case txtԤ����Ѫ��
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
    If Index = txtԤ����Ѫʱ�� Then
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(txt��������).Text) Then
                    '�ָ���Ϊ�����ȱʡΪ��ʼʱ��
                    txtInfo(Index).Text = txtInfo(txt��������).Text
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check����ʱ��(txtInfo(Index).Text, txtInfo(txt��������).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
        End If
    ElseIf Index = txt�������� Then
            
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(txt��������).Tag) Then
                    '�ָ���Ϊ�����
                    txtInfo(Index).Text = txtInfo(txt��������).Tag
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check��ʼʱ��(txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            '�ж��Ƿ��ǲ�¼ҽ��
            If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint��¼��� _
                Or mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
                mbln��¼ = True
                SetControlEnabled cboInfo(cbo��Ѫ����), False
            Else
                mbln��¼ = False
                SetControlEnabled cboInfo(cbo��Ѫ����), True
            End If
        End If
    ElseIf Index = txtԤ����Ѫ�� Then
        If Val(txtInfo(Index).Text) > mlng¼������ And mlng¼������ > 0 Then
            If MsgBox(txtGet(txtԤ����Ѫ�ɷ�).Text & " ������:" & Val(txtInfo(Index).Text) & txtInfo(txt��λ).Text & " ��������¼����������:" & _
                mlng¼������ & txtInfo(txt��λ).Text & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
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
            If .TextMatrix(NewRow, NewCol + (COL_ȡֵ���� - COL_ָ����)) <> "" Then
                '�ϰ���°��õķָ����ͬ���°��Ƕ��ţ��ϰ�ֺţ����¼��ݴ���
                strTmp = .TextMatrix(NewRow, NewCol + (COL_ȡֵ���� - COL_ָ����))
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

        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
            'Ϊ��֧��Ԥ�����
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
'���ܣ����λ����һ��
    With vsLIS
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '�������������ݹ��ٶ�λ����һ��λ��
        If .Cell(flexcpBackColor, .Row, .Col) <> COLEditBackColor Then Call EnterNextCell
        .ShowCell .Row, .Col
    End With
End Sub
