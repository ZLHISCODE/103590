VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyBloodNew 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ѫ���뵥"
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
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "δ����"
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
         Left            =   1155
         TabIndex        =   29
         Top             =   30
         Width           =   975
      End
      Begin VB.OptionButton optAppraise 
         BackColor       =   &H8000000E&
         Caption         =   "������"
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
         Caption         =   "δǩ"
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
         Left            =   960
         TabIndex        =   27
         Top             =   30
         Width           =   735
      End
      Begin VB.OptionButton optConsent 
         BackColor       =   &H8000000E&
         Caption         =   "��ǩ"
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
         Caption         =   "������ʷ������Ŀ��"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
               Name            =   "����"
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
         Caption         =   "����"
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
         Index           =   30
         Left            =   4485
         TabIndex        =   112
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "��/��"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   960
         TabIndex        =   23
         Top             =   30
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
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   30
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
         Index           =   4
         Left            =   960
         TabIndex        =   25
         Top             =   30
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd�������� 
      Height          =   300
      Left            =   5355
      Picture         =   "frmApplyBloodNew.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "����ǰ��������Ϊ��������(Ctrl+D)"
      Top             =   9660
      Width           =   315
   End
   Begin VB.CommandButton cmdҽ������ 
      Caption         =   "��"
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
            Name            =   "����"
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
      Left            =   795
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   9690
      Width           =   4425
   End
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
         TabIndex        =   1
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
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   -30
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
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
      ToolTipText     =   "�༭(F4)"
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
      ToolTipText     =   "�༭(F4)"
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
         ToolTipText     =   "�༭(F4)"
         Top             =   0
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
         Top             =   30
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
         Left            =   960
         TabIndex        =   21
         Top             =   30
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
      Top             =   1605
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
         Name            =   "����"
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
            Begin VB.TextBox txt������ 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
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
                     Name            =   "����"
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
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "����"
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
         Begin VB.TextBox txt������Ϣ 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
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
            Text            =   "Ʒ�֣�"
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
         Caption         =   "ѪҺ��Ϣ"
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
         Name            =   "����"
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
      Caption         =   "��Ѫǰ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ����֪��ͬ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "24Сʱ����Ѫ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ѫ����֢������ʷ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������Ѫ��Ӧʷ"
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
      Left            =   8820
      TabIndex        =   69
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ע��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ע"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ϼ�ҽʦǩ��"
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
      Height          =   240
      Index           =   22
      Left            =   6195
      TabIndex        =   77
      Top             =   7410
      Width           =   1020
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
      TabIndex        =   66
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   57
      Top             =   2055
      Width           =   840
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
      TabIndex        =   55
      Top             =   1665
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
      TabIndex        =   54
      Top             =   1665
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
      TabIndex        =   53
      Top             =   1665
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
      TabIndex        =   51
      Top             =   1290
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
      TabIndex        =   50
      Top             =   1665
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
      TabIndex        =   49
      Top             =   1260
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
      TabIndex        =   56
      Top             =   1290
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
      TabIndex        =   52
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   15
      Left            =   6705
      TabIndex        =   67
      Top             =   4200
      Width           =   510
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
      Left            =   7095
      TabIndex        =   79
      Top             =   9735
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "����ҽʦ����"
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
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng�������� As Long   '0-סԺ��1-����
Private mblnChange As Boolean
Private mblnHaveAuditPriv As Boolean 'ִҵҽʦ�ʸ�
Private mintType As Integer   '0-������1-�޸ģ�2-�鿴,3-ҽ���༭���ã�ֻ�ܵ�������Ѫ�ɷ֣�����������ʱ�䣬��Ѫʱ�䣬ִ�п��ң���Ѫ;������Ѫִ�п��ң���Ѫ������������ݣ�4-ҽ���˶���Ѫ����(��Ѫ��ֱ�ӷ�Ѫ������ҽ��)
Private mlngUpdateAdvice As Long  '�޸ĵ�ҽ��ID
Private mintPState As Integer
Private mdatTurn As Date
Private mlng���˿���id As Long
Private mlng����ID As Long
Private mlng��������ID As Long
Private mlng��Ѫ;�� As Long
Private mlng��Ѫ��ĿID As Long, mlngPre��Ѫ��ĿID As Long
Private mstr��Ѫ��Ŀ As String  '��Ѫ����Ʒ�ֿ�ѡ��������ʽ����ĿID,������,����Ѫ��,����RH
Private mstrLISAboRHCode As String
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
Private mint���뵥��ӡģʽ As Integer  '1-����ʱ��ӡ��0-�¿�ʱ��ӡ
Private mint���� As Integer '��ǰ��������
Private mbln���Ѷ��� As Boolean
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����
Attribute mclsMipModule.VB_VarHelpID = -1
Private Const CON_LisResultCol = 3
Private Const CON_LisResultCount = 10
Private mobjPublicLis As Object
Private mint���� As Integer '0-סԺ��1�����Ĭ��ΪסԺ
Private mstr�Һŵ� As String '�Һŵ���
Private mlng�Һ�ID As Long
Private mlngǰ��ID As Long
Private mrsCard As ADODB.Recordset
Private mbytBaby As Integer  'Ӥ�����
Private mstr�����Ժ��� As String
Private mint���ó��� As Integer '0-����վ���ã�1��ҽ���´�������
Private mblnNewSpareBloood As Boolean  '�Ƿ�����±�Ѫģʽ
Private mblnSpareBloood As Boolean  '�Ǳ�Ѫ���뻹����Ѫ����
Private mblnUseBloodSend As Boolean '��Ѫҽ���Ƿ��Ѿ���Ѫ
Private mstrժҪ��Ѫ As String 'ժҪ���� gclsInsure.GetItemInfo ��ȡ
Private mstrժҪ;�� As String
Private mstr�ѱ� As String
Private mblnDataLoad As Boolean
Private mblnSelectBlood As Boolean '����Ѫ�����Ƿ����ҽ��ѡ��Ѫ����ģʽ(�����ڿ��õ���Ѫ��Ϣ��ͨ��ҽ��ָ��Ѫ���´����룩

Private Enum Enum_Cbo
    cbo��Ѫ���� = 0
    cbo��ѪѪ�� = 1
    cboRHD = 2
    cbo���� = 3
    cbo��λ = 4
    cbo��Ѫ���� = 5
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
    lbl������Ѫʷ = 11
    lbl������Ѫ��Ӧʷ = 32
    lbl��Ѫ���ɼ�����ʷ = 38
    lbl�в���� = 12
    lbl��Ѫ������ = 13
    lblԤ����Ѫ���� = 14
    lblѪ�� = 15
    lblRHD = 16
    lblԤ����Ѫ�ɷ� = 17
    lbl��Ѫִ�� = 18
    lblԤ����Ѫ�� = 19
    lbl��Ѫ;�� = 21
    lbl��Ѫִ�� = 22
    lbl������ = 23
    lbl��ע = 24
    lbl����ҽʦǩ�� = 33
    lbl�ɼ���ǩ�� = 34
    lbl����ҽʦǩ�� = 36
    lblע�� = 25
    '---
    lbl����ҽʦ���� = 29
    lbl24H��Ѫ�� = 40
    lbl������ʷ������Ŀ = 41
    lbl֪��ͬ���� = 42
    lbl��Ѫ���� = 43
End Enum

Private Enum Enum_lin
    lin����ҽʦǩ�� = 30
    lin��Ѫ��ǩ�� = 31
    lin����ҽʦǩ�� = 33
    '---
    lin����ҽʦ���� = 24
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
    txt��ע = 9 'ҽ������
    txtԤ����Ѫʱ�� = 10
    txtԤ����Ѫ�� = 11
    txt��λ = 12
    txt�� = 13
    txt�� = 14
    txt����ҽʦǩ�� = 17
    txt��Ѫ��ǩ�� = 18
    txt�������� = 19
    txt����ҽʦǩ�� = 20
    '----
    txt����ҽʦ���� = 15
End Enum

Private Enum Enum_FraChk
    fra������Ѫʷ = 0
    fra�в���� = 1
    fra��Ѫ������ = 2
    fra������Ѫ��Ӧʷ = 3
    fra��Ѫ���ɼ�����ʷ = 4
    fra֪��ͬ���� = 5
    fra��Ѫ���� = 6
End Enum

Private Enum Enum_Get
    txtԤ����Ѫ�ɷ� = 0
    txt��Ѫ;�� = 1
End Enum

Private Enum Enum_cmdDate
    cmdԤ����Ѫʱ�� = 0
    cmd�������� = 1
End Enum

Private Enum Enum_cmdGet
    cmd��Ѫ;�� = 1
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

Private Enum Enum_P_BloodCol
    COL_P_ID = 0
    COL_P_ѡ�� = 1
    COL_P_���� = 2
    COL_P_���� = 3
    COL_P_������ = 4
    COL_P_��λ = 5
    COL_P_����Ѫ�� = 6
    COL_P_����RH = 7
    COL_P_ִ�з���ID = 8
    COL_P_ִ�п���ID = 9
    COL_P_¼������ID = 10
    COL_P_����ϵ�� = 11
    COL_P_��� = 12
End Enum

Private Enum Enum_S_BloodList
    COL_S_ID = 0
    COL_S_ѡ�� = 1
    COL_S_��� = 2
    COL_S_��� = 3
    COL_S_Ч�� = 4
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�������� As Long, ByVal intType As Integer, Optional ByRef lngUpdateAdvice As Long, _
    Optional ByVal lng���˿���ID As Long, Optional ByVal lng����ID As Long, Optional ByVal lng��������ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
    Optional ByRef rsDefine As Recordset, Optional ByRef objMip As Object, Optional ByVal int���� As Integer, Optional ByVal str�Һŵ� As String, Optional ByVal lng��Ŀid As Long, _
    Optional ByRef rsCard As ADODB.Recordset, Optional ByVal bytBaby As Byte, Optional ByVal int���ó��� As Integer, Optional ByVal lngǰ��ID As Long, Optional ByVal int���뵥ģʽ As Integer = 0) As Boolean
      
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
    mlngǰ��ID = lngǰ��ID
    Set mrsDefine = rsDefine
    
    mlngUpdateAdvice = lngUpdateAdvice
    
    mlng��Ѫ��ĿID = lng��Ŀid
    mstr��Ѫ��Ŀ = lng��Ŀid & ",,,"
    mlngPre��Ѫ��ĿID = mlng��Ѫ��ĿID
    Set mrsCard = rsCard
    mbytBaby = bytBaby
    mint���ó��� = int���ó���
    mblnSpareBloood = (int���뵥ģʽ = 0)
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    If mblnOK = True Then lngUpdateAdvice = mlngUpdateAdvice
    Set rsCard = mrsCard
    ShowMe = mblnOK
End Function

Private Function SeekNextControl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextControl = True
End Function

Private Sub cboInfo_Change(Index As Integer)
    If Visible And (Index = cbo��ѪĿ�� Or Index = cbo��Ѫ����) And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSQL As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = cboִ�п��� Or Index = cbo��Ѫִ�� Then
        If cboInfo(Index).ItemData(cboInfo(Index).ListIndex) = -1 Then
            
            '����ִ�У�����ѡ��ִ�п���
            strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                IIF(gstrNodeNo <> "", " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " Order by A.����"
            vRect = zlControl.GetControlRect(cboInfo(Index).hwnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vRect.Left, vRect.Top, cboInfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = Cbo.FindIndex(cboInfo(Index), rsTmp!ID)
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
                    intIdx = Cbo.FindIndex(cboInfo(Index), Val(cboInfo(Index).Tag))
                    Call zlControl.CboSetIndex(cboInfo(Index).hwnd, intIdx)
                End If
            End If
        End If
    ElseIf Index = cbo���� Then
        '���ٺͼ�ѹ����ʾ��λ
        If cboInfo(Index).ListIndex = 2 Or cboInfo(Index).ListIndex = 3 Then
            lblInfo(31).Visible = False
        Else
            lblInfo(31).Visible = True
        End If
    ElseIf Index = cbo��λ Then
        If Val(cboInfo(Index).Tag) = cboInfo(Index).ListIndex And cboInfo(Index).Tag <> "" Then Exit Sub
        cboInfo(Index).Tag = cboInfo(Index).ListIndex
        Call BloodSum
        Call RsetBreedUnit
    ElseIf Index = cbo��ѪѪ�� Then
        Call SetBloodLisAboRh(Index)
    ElseIf Index = cboRHD Then
        If cboInfo(Index).Text = "-" Then
            cboInfo(Index).ForeColor = vbRed
        Else
            cboInfo(Index).ForeColor = &H80000008
        End If
        Call SetLblRh
        Call SetBloodLisAboRh(Index)
    ElseIf Index = cbo��Ѫ���� And mblnEditable = True Then
        intIdx = Val(GetBloodApplyCode(0))
        cboInfo(cbo��ѪѪ��).Enabled = intIdx = 0
        cboInfo(cboRHD).Enabled = intIdx = 0
        Call SetLblRh
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
        strReturn = strText
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
        If InStr(strText, "[Ѫ��]") > 0 Then
            strField = Trim(cboInfo(cbo��ѪѪ��).Text)
            strText = Replace(strText, "[Ѫ��]", """" & strField & """")
        End If
        If InStr(strText, "[RH]") > 0 Then
            strField = Trim(cboInfo(cboRHD).Text)
            strText = Replace(strText, "[RH]", """" & strField & """")
        End If
        If InStr(strText, "[ִ�з���]") > 0 Then
            strField = IIF(mblnSpareBloood, 0, 1)
            strText = Replace(strText, "[ִ�з���]", """" & strField & """")
        End If
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

Private Function CheckUseBlood() As Boolean
'���ܣ�������Ѫ��ʱ�������Ѫ�������,�Ƿ񳬳������н����ֻ�ǽ�����ʾ����ǿ�ƽ�ֹ
    Dim lngRow As Long
    Dim dblTotal As Double   '��������
    Dim dblApplyTotal As Double  '��������
    Dim str���뵥λ As String, lng����ϵ�� As Long, str�����Ŀ As String
    Dim strTmp As String, arrInfo, arrItem
    Dim j As Integer
    Dim strMsg As String
    Dim objCollection As New Collection
    
    If mblnSpareBloood = False And mblnSelectBlood = False Then
        With vsfBlood
            '��ȡÿһ��Ʒ�ֵĴ���Ѫ��
            For lngRow = .FixedRows To .Rows - 1
                strTmp = .TextMatrix(lngRow, COL_P_���)
                dblTotal = 0: dblApplyTotal = 0
                If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                    '�˴���ȡ��Ʒ�ֵ�ʣ������ͳһת��ΪML
                    lng����ϵ�� = Val(.TextMatrix(lngRow, COL_P_����ϵ��)): If lng����ϵ�� = 0 Then lng����ϵ�� = 1
                    str���뵥λ = UCase(.TextMatrix(lngRow, COL_P_��λ)): If str���뵥λ = "" Then str���뵥λ = "ML"
                    
                    dblApplyTotal = Val(.TextMatrix(lngRow, COL_P_������))
                    If str���뵥λ <> "ML" Then
                        dblApplyTotal = dblApplyTotal * lng����ϵ��
                    End If
                End If
                If strTmp <> "" Then
                    arrInfo = Split(strTmp, "<Split2>")
                    If UBound(arrInfo) > 0 Then
                        arrItem = Split(arrInfo(1), "'")
                        dblTotal = Val(arrItem(2)) 'ML
                    End If
                End If
                '��ȡ��Ŀ�����ʣ����
                dblTotal = dblTotal - dblApplyTotal
                objCollection.Add dblTotal, "A_" & .TextMatrix(lngRow, COL_P_ID)
            Next
            strMsg = ""
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                    '�����Ŀ<Split2>Ʒ��ID'�䷢��Ϣ'������
                    'ҽ��ID'ҽ������'������'�����Ŀ<Split2>Ʒ��ID'�䷢��Ϣ'������<Split3>Ʒ��ID'�䷢��Ϣ'������<Split3>Ʒ��ID'�䷢��Ϣ'������...<Split1>ҽ��ID'ҽ������'������<Split2>.....
                    strTmp = .TextMatrix(lngRow, COL_P_���)
                    arrInfo = Split(strTmp, "<Split2>")
                    str�����Ŀ = arrInfo(0)
                    '�������Ҳ���,��ȡ��
                    dblTotal = Val(objCollection("A_" & .TextMatrix(lngRow, COL_P_ID)))
                    If dblTotal < 0 Then
                        If str�����Ŀ <> "" Then
                            arrItem = Split(str�����Ŀ, ",")
                            For j = 0 To UBound(arrItem)
                                If Val(arrItem(j)) > 0 Then
                                    dblApplyTotal = Val(ISExistCollection(objCollection, "A_" & Val(arrItem(j))))
                                    If dblApplyTotal > 0 And Val(arrItem(j)) <> Val(.TextMatrix(lngRow, COL_P_ID)) Then
                                        'ʣ���������������������˳����Ҵ��¸��������Ŀ��ʣ����
                                        If dblApplyTotal >= dblTotal Then
                                            dblApplyTotal = dblApplyTotal - Abs(Val(dblTotal))
                                            dblTotal = 0
                                        Else
                                            dblApplyTotal = 0
                                            dblTotal = dblApplyTotal - Abs(Val(dblTotal))
                                        End If
                                        '���¼���
                                        objCollection.Remove "A_" & Val(arrItem(j))
                                        objCollection.Add dblApplyTotal, "A_" & Val(arrItem(j))
                                    End If
                                    If dblTotal >= 0 Then Exit For
                                End If
                            Next
                        End If
                        If dblTotal < 0 Then
                            strMsg = IIF(strMsg = "", "", strMsg & vbCrLf) & "[" & .TextMatrix(lngRow, COL_P_����) & "]�����������������Ѫ��������"
                        End If
                    End If
                End If
            Next
        End With
        
        If strMsg <> "" Then
            If MsgBox(strMsg & vbCrLf & "�������Ƿ�Ҫ��������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If vsfBlood.Enabled And vsfBlood.Visible Then vsfBlood.SetFocus
                Exit Function
            End If
        End If
    End If
    CheckUseBlood = True
End Function

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngִ������ As Long
    Dim lngִ�п���ID As Long
    Dim lng����ID As Long
    Dim bln��ҽ As Boolean
    Dim str���� As String
    Dim blnSucceed As Boolean
    Dim strTmp As String
    Dim i As Long, intMax As Integer
    Dim strTabAdvice As String
    Dim strItems As String
    Dim blnCheckҽ�� As Boolean
    Dim rsPrice As ADODB.Recordset
    Dim lngRow As Long
    
    ' Call SeekNextControl  '�����ַ�ʽ�������71290
    '������������費ͬ�ؼ��Ľ��㣬ȷ��validata�¼���ִ�С�
    txtGet(txt��Ѫ;��).SetFocus
    
    If txtInfo(txt�����Ϣ).Enabled = True And txtInfo(txt�����Ϣ).Locked = False Then
        If Trim(txtInfo(txt�����Ϣ).Text) = "" Then
            MsgBox "���������ٴ���ϣ�", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txtInfo(txt�����Ϣ))
            Exit Function
        End If
        intMax = txtInfo(txt�����Ϣ).MaxLength
        If LenB(StrConv(txtInfo(txt�����Ϣ).Text, vbFromUnicode)) > intMax Then
            MsgBox "�ٴ���ϲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Call zlControl.ControlSetFocus(txtInfo(txt�����Ϣ))
            Exit Function
        End If
    End If
    '�༭��������������
    If mintType <> 3 Then
        '����ѡ����Ѫ����
        If cboInfo(cbo��Ѫ����).Text = "" Then
            MsgBox "������д������Ѫ���ͣ�", vbInformation, Me.Caption
            If cboInfo(cbo��Ѫ����).Enabled Then cboInfo(cbo��Ѫ����).SetFocus
            Exit Function
        End If
        
        '������ҽ��������д��ѪĿ��
        If cboInfo(cbo��Ѫ����).ListIndex = 1 And cboInfo(cbo��ѪĿ��).Text = "" Then
            MsgBox "������Ѫ������д��ѪĿ�ġ�", vbInformation, Me.Caption
            If cboInfo(cbo��ѪĿ��).Enabled Then cboInfo(cbo��ѪĿ��).SetFocus
            Exit Function
        End If
        
        '��Ѫҽ���ż��
        If mblnSpareBloood = True Then
            '�в�������
            If txtInfo(txt��).Text <> "" And txtInfo(txt��).Text = "" Then
                MsgBox "�������в�����е��дΣ������������Ρ�", vbInformation, Me.Caption
                If txtInfo(txt��).Visible And txtInfo(txt��).Enabled Then txtInfo(txt��).SetFocus
                Exit Function
            End If
            If txtInfo(txt��).Text <> "" And txtInfo(txt��).Text = "" Then
                MsgBox "�������в�����еĲ��Σ�����������дΡ�", vbInformation, Me.Caption
                If txtInfo(txt��).Visible And txtInfo(txt��).Enabled Then txtInfo(txt��).SetFocus
                Exit Function
            End If
            If Val(txtInfo(txt��).Text) > 0 Then
                If Val(txtInfo(txt��).Text) = 0 Then
                    MsgBox "���в�����еĲ��β�Ϊ0ʱ������������дΣ��Ҵ����������0��", vbInformation, Me.Caption
                    If txtInfo(txt��).Visible And txtInfo(txt��).Enabled Then txtInfo(txt��).SetFocus
                    Exit Function
                End If
            End If
            
            '��Ѫ����ͬ�������Ѫ����������д
            If optConsent(0).value = False And optConsent(1).value = False Then
                MsgBox "����ȷ����Ѫ����ͬ�����Ƿ���ǩ��", vbInformation, Me.Caption
                'option�������ý��㣬���ý�����Զ���ѡ
                Exit Function
            End If
            
            If optAppraise(0).value = False And optAppraise(1).value = False Then
                MsgBox "����ȷ����Ѫǰ�����Ƿ���������", vbInformation, Me.Caption
                Exit Function
            End If
        End If
        '����¼����Ѫ�ɷ�
        If mlng��Ѫ��ĿID = 0 Then
            MsgBox "û��ȷ��Ԥ����Ѫ�ɷ֡�", vbInformation, Me.Caption
            If vsfBlood.Visible And vsfBlood.Enabled Then
                If vsfBlood.Row > vsfBlood.FixedRows Then
                    vsfBlood.Row = vsfBlood.FixedRows
                    vsfBlood.Col = COL_P_ѡ��
                End If
                vsfBlood.SetFocus
            End If
            Exit Function
        End If
        
        '���ִ�п���
        If cboInfo(cboִ�п���).Text = "" Then
            MsgBox "û��ȷ��ִ�п��ҡ�", vbInformation, Me.Caption
            If cboInfo(cboִ�п���).Enabled Then cboInfo(cboִ�п���).SetFocus
            Exit Function
        End If
        
        '�����ٶ�
        If cboInfo(cbo����).Visible = True Then '�ɼ�˵���϶�����Ѫ����
            If cboInfo(cbo����).ListIndex < 0 Then
                If LenB(StrConv(cboInfo(cbo����).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cboInfo(cbo����).Text) And cboInfo(cbo����).Text <> "") Then
                    MsgBox "����¼��ĵ���ֻ�������֣������ֻ����¼��3λ���֣�", vbInformation, gstrSysName
                    Call zlControl.ControlSetFocus(cboInfo(cbo����))
                    Exit Function
                End If
            End If
        End If
        
        '�����Ѫ;������Ѫִ��
        If mlng��Ѫ;�� = 0 Then
            If mblnNewSpareBloood = False Then
                MsgBox "û��ָ����Ѫ;����", vbInformation, Me.Caption
            Else
                MsgBox "û��ָ���ɼ���ʽ��", vbInformation, Me.Caption
            End If
            If txtGet(txt��Ѫ;��).Enabled Then txtGet(txt��Ѫ;��).SetFocus
            Exit Function
        End If
        
        If cboInfo(cbo��Ѫִ��).Text = "" Then
            If mblnNewSpareBloood = False Then
                MsgBox "û��ȷ����Ѫִ�п��ҡ�", vbInformation, Me.Caption
            Else
                MsgBox "û��ȷ���ɼ�ִ�п��ҡ�", vbInformation, Me.Caption
            End If
            If cboInfo(cbo��Ѫִ��).Enabled Then cboInfo(cbo��Ѫִ��).SetFocus
            Exit Function
        End If
        
        '����¼������
        If cboInfo(cbo��λ).ListIndex = -1 Then
            MsgBox "��ȷ����Ѫ������λ��", vbInformation, Me.Caption
            If cboInfo(cbo��λ).Enabled And cboInfo(cbo��λ).Visible Then cboInfo(cbo��λ).SetFocus
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
        
        With vsfBlood
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                    If Val(.TextMatrix(lngRow, COL_P_������)) <= 0 Then
                        If mblnSelectBlood = False Then
                            MsgBox "��¼�����0����Ѫ��������", vbInformation, Me.Caption
                            .Row = lngRow: .Col = COL_P_������
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                        Else
                            MsgBox "[" & .TextMatrix(lngRow, COL_P_����) & "]��δѡ��ѪҺ��Ϣ��", vbInformation, Me.Caption
                            .Row = lngRow: .Col = COL_P_����
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                        End If
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngRow, COL_P_������)) > Val(.TextMatrix(lngRow, COL_P_¼������ID)) And Val(.TextMatrix(lngRow, COL_P_¼������ID)) > 0 Then
                        If MsgBox(.TextMatrix(lngRow, COL_P_����) & " ������:" & Val(.TextMatrix(lngRow, COL_P_������)) & .TextMatrix(lngRow, COL_P_��λ) & " ��������¼����������:" & _
                            Val(.TextMatrix(lngRow, COL_P_¼������ID)) & .TextMatrix(lngRow, COL_P_��λ) & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            .Row = lngRow: .Col = COL_P_������
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
        
        If Trim(cboInfo(cbo��ѪѪ��).Text) = "" And Trim(cboInfo(cboRHD).Text) = "" Then
            If MsgBox("û��ȷ��Ѫ�ͺ�RH(D)�������Ƿ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cbo��ѪѪ��).Enabled Then cboInfo(cbo��ѪѪ��).SetFocus
                Exit Function
            End If
        ElseIf Trim(cboInfo(cbo��ѪѪ��).Text) = "" Then
            If MsgBox("û��ȷ��Ѫ���������Ƿ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cbo��ѪѪ��).Enabled Then cboInfo(cbo��ѪѪ��).SetFocus
                Exit Function
            End If
        ElseIf Trim(cboInfo(cboRHD).Text) = "" Then
            If MsgBox("û��ȷ��RH(D)�������Ƿ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If cboInfo(cboRHD).Enabled Then cboInfo(cboRHD).SetFocus
                Exit Function
            End If
        End If
        If BloodApplyCheck = False Then Exit Function '�����Զ�����̣���ҽԺ���ж��������뵥���м��
        If CheckOrResetLisAboRH = False Then Exit Function '���LIS����е�Ѫ���Ƿ��ѡ��Ѫ���Ƿ�һ��
        If CheckUseBlood = False Then Exit Function '��Ѫ�������������Ƿ�����䷢����������ʾ��
        
        If mint���ó��� = 0 Then
            lng����ID = IIF(mint���� = 1, mlng�Һ�ID, mlng��ҳID)
            strTmp = mlng��Ѫ��ĿID & "||" & IIF(mint���� = 1, 1, 2)
            mstrժҪ��Ѫ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", strTmp)
            strTmp = mlng��Ѫ;�� & "||" & IIF(mint���� = 1, 1, 2)
            mstrժҪ;�� = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", strTmp)
            strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
            For i = 1 To 2
                If i = 1 Then
                    lngִ������ = IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, 5, mlngִ�п�������)
                    lngִ�п���ID = IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, 0, cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex))
                    
                    strTabAdvice = "select 1 as ID,1 as ���,-null as ���ID,'K' as �������," & mlng��Ѫ��ĿID & " as ������ĿID," & _
                            mlng��Ѫ��ĿID & " as ������ĿID," & Val(txtInfo(txtԤ����Ѫ��).Text) & " As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
                            "0 as ִ�б��,0 as �Ƽ�����, null As ��������," & lngִ������ & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
                    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", IIF(1 = mint����, 1, 2), mlng����ID, lng����ID, mint����, 1, _
                         "K", mlng��Ѫ��ĿID, mlng��������ID, UserInfo.����, lngִ�п���ID, lngִ������, 0, 0, mstrժҪ��Ѫ)
                Else
                    lngִ������ = IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, 5, mlng��Ѫִ������)
                    lngִ�п���ID = IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, 0, cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex))
                    
                    strTabAdvice = strTabAdvice & " Union All " & _
                         "select 2 as ID,2 as ���,1 as ���ID,'E' as �������," & mlng��Ѫ;�� & " as ������ĿID," & _
                            mlng��Ѫ;�� & " as ������ĿID,1 As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
                            "0 as ִ�б��,0 as �Ƽ�����, null As ��������," & lngִ������ & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"

                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", IIF(1 = mint����, 1, 2), mlng����ID, lng����ID, mint����, 1, _
                         "E", mlng��Ѫ;��, mlng��������ID, UserInfo.����, lngִ�п���ID, lngִ������, 0, 0, mstrժҪ;��)
                End If
                
                If Not rsTmp.EOF Then
                    strMsg = NVL(rsTmp!���)
                    If strMsg <> "" Then
                        Select Case Val(Split(strMsg, "|")(0))
                        Case 1 '��ʾ
                            If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                strMsg = "": Exit Function
                            End If
                        Case 2 '��ֹ
                            MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                            strMsg = "": Exit Function
                        End Select
                        strMsg = ""
                    End If
                End If
            Next
            
            '��ϼ��
             If InStr(mstr�����Ժ���, "K") > 0 And mint���� = 0 Then
                bln��ҽ = Sys.DeptHaveProperty(mlng���˿���id, "��ҽ��")
                str���� = IIF(bln��ҽ, "2,12", "2")
                If Not ExistsDiagNoses(mlng����ID, mlng��ҳID, str����) Then
                    strMsg = "���˵���Ժ��ϻ�û�����룬�������벡�˵���Ժ������´���Ѫ���롣"
                End If
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '������
            With vsfBlood
                For lngRow = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                        strIDs = IIF(strIDs = "", "", strIDs & ",") & Val(.TextMatrix(lngRow, COL_P_ID)) & ":"
                        If Val(cboInfo(cboִ�п���).Tag & "") <> 0 Then
                            strIDs = strIDs & Val(cboInfo(cboִ�п���).Tag & "")
                        End If
                    End If
                Next
            End With
            strҽ������ = FormatAdviceContext(Replace(txtGet(txtԤ����Ѫ�ɷ�).Text, "'", ","), txtGet(txt��Ѫ;��).Text)
            
            strIDs = strIDs & "," & mlng��Ѫ;�� & ":"
            If Val(cboInfo(cbo��Ѫִ��).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo��Ѫִ��).Tag & "")
            End If
            If gintҽ������ = 2 Then mbln���Ѷ��� = True
            strItems = strIDs
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
            'ҽ���ܿ�ʵʱ���
            If mint���� <> 0 Then
                If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                    If MakePriceRecord���뵥("3" & IIF(mint���� = 1, "1", "2"), mlng����ID, lng����ID, strTabAdvice, strItems, mstr�ѱ�, mlng��������ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint����, IIF(mint���� = 1, 0, 1), 0, rsPrice) Then
                            MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´����Ѫ���뵥���ܱ��档", vbInformation, gstrSysName
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
'���ܣ���������
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lngҽ��ID As Long, lngҽ����� As Long, lng������� As Long
    Dim strSQL As String, rsTmp As Recordset
    Dim rsData As New ADODB.Recordset, rsTemp As Recordset
    Dim str��Ŀ���� As String, str��Ѫ;�� As String, strPrivs As String
    Dim curDate As Date, i As Long, lng���ID As String, j As Long
    Dim lngCount As Long, int������Դ As Integer
    Dim strTmp��ҳID As String
    Dim strTmp�Һŵ� As String
    Dim str���״̬ As String
    Dim int���� As Integer
    Dim int���� As Integer
    Dim int��鷽�� As Integer
    Dim str���� As String
    Dim strErr As String
    Dim bln����� As Boolean
    Dim dbl24h�� As Double, dblTmp As Double
        
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If cboInfo(cbo����).Visible = True Then
        str���� = cboInfo(cbo����).Text
    End If
    If IsNumeric(str����) = True Then
        str���� = str���� & "��/����"
    End If
    
    If mintType = 3 Then
        '���븽��༭ģʽ
        lng���ID = mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Delete(" & lng���ID & ")"
    ElseIf mintType = 4 Or (mintType = 1 And mblnUseBloodSend = True) Then '�˶�ҽ�����޸���Ѫҽ��
        '�����Ѫҽ��״̬���Ѿ����״̬(�Ѿ������˵�����ξ�������ٴ��޸�)
        If mintType = 1 And mblnUseBloodSend = True Then
            gstrSQL = "Select ����ʱ�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ѫҽ���Ƿ����", mlngUpdateAdvice, 11)
            bln����� = rsData.RecordCount > 0
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_ҽ����˹���_Cancel('" & mlngUpdateAdvice & "')"
        End If
        
        gstrSQL = "Select Id, ���id, ���, ҽ��״̬, ҽ����Ч, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����," & vbNewLine & _
            "       ִ�п���id, ִ������, ������־, ��ʼִ��ʱ��, ִ����ֹʱ��, ���˿���id, ��������id, ����ҽ��, ����ʱ��, ��鷽��, ִ�б��, �ɷ����, ժҪ, ��Ѽ���, ��ҩĿ��, ��ҩ����, ���״̬," & vbNewLine & _
            "       ����˵��, �״�����, �������, �����Ŀid, Ƥ�Խ��" & vbNewLine & _
            "From ����ҽ����¼" & vbNewLine & _
            "Where Id = [1] Or ���id = [1]" & vbNewLine & _
            "Order By Nvl(���id, 0)"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ҽ����Ϣ", mlngUpdateAdvice)
        
        str��Ŀ���� = ""
        For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
            If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
                str��Ŀ���� = IIF(str��Ŀ���� = "", "", str��Ŀ���� & ",") & vsfBlood.TextMatrix(i, COL_P_����)
            End If
        Next
        If str��Ŀ���� = "" Then str��Ŀ���� = Sys.RowValue("������ĿĿ¼", mlng��Ѫ��ĿID, "����")
         
        Set rsTmp = Get������Ŀ��¼(mlng��Ѫ;��)
        str��Ѫ;�� = rsTmp!���� & ""
        
        rsData.Filter = "ID=" & mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ����¼_Update(" & rsData!ID & ",NULL," & rsData!��� & ",1,1," & mlng��Ѫ��ĿID & _
                                 "," & ZVal(NVL(rsData!�շ�ϸĿID, 0)) & "," & ZVal(NVL(rsData!����, 0)) & "," & ZVal(NVL(rsData!��������, 0)) & "," & ZVal(txtInfo(txtԤ����Ѫ��).Text) & ",'" & FormatAdviceContext(str��Ŀ����, str��Ѫ;��) & _
                                 "'," & IIF(txtInfo(txt��ע).Text = "", "NULL", "'" & txtInfo(txt��ע).Text & "'") & ",'" & Format(txtInfo(txtԤ����Ѫʱ��).Text, "yyyy-MM-dd HH:mm:ss") & "','һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "Null", cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex)) & _
                                 "," & IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "5", mlngִ�п�������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                   ZVal(NVL(rsData!��鷽��, 0)) & ",0,NULL," & IIF(mstrժҪ��Ѫ = "", "null", "'" & mstrժҪ��Ѫ & "'") & ",'" & UserInfo.���� & "',Null,Null,'" & cboInfo(cbo��ѪĿ��).Text & "'," & IIF(bln����� = True, 1, ZVal(NVL(rsData!���״̬, 0))) & ")"
        
        lng���ID = mlngUpdateAdvice
        rsData.Filter = "���id=" & mlngUpdateAdvice
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ����¼_Update(" & rsData!ID & "," & lng���ID & "," & rsData!��� & ",1,1," & mlng��Ѫ;�� & ",NULL,NULL,NULL,Null,'" & str��Ѫ;�� & "'," & IIF(str���� = "", "NULL", "'" & str���� & "'") & ",NULL,'һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "Null", cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "5", mlng��Ѫִ������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                ZVal(NVL(rsData!��鷽��, 0)) & ",0,NULL," & IIF(mstrժҪ;�� = "", "null", "'" & mstrժҪ;�� & "'") & ",'" & UserInfo.���� & "',Null,NULL,''," & IIF(bln����� = True, 1, ZVal(NVL(rsData!���״̬, 0))) & ")"
        
        If mintType = 4 Or bln����� = True Then
            '��ɺ˶�
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_ҽ����˹���_Audit(" & lng���ID & "," & 1 & "," & _
                            "'" & UserInfo.���� & "',to_date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'))"
        End If
                        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Delete(" & lng���ID & ")"
        str���״̬ = "NULL"
    Else
        
        lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")        '��ȡҽ��ID
        
        '����ҽ����¼.��ţ�����
        If mint���� = 0 Then
            lngҽ����� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, mbytBaby) + 1
            strTmp��ҳID = mlng��ҳID
            strTmp�Һŵ� = "NULL"
            int������Դ = 2
        Else
            lngҽ����� = GetMaxAdviceNO(mlng����ID, , mbytBaby, mstr�Һŵ�) + 1
            strTmp��ҳID = "NULL"
            strTmp�Һŵ� = "'" & mstr�Һŵ� & "'"
            int������Դ = 1
        End If
        
        str��Ŀ���� = ""
        For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
            If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
                str��Ŀ���� = IIF(str��Ŀ���� = "", "", str��Ŀ���� & ",") & vsfBlood.TextMatrix(i, COL_P_����)
            End If
        Next
        If str��Ŀ���� = "" Then str��Ŀ���� = Sys.RowValue("������ĿĿ¼", mlng��Ѫ��ĿID, "����")
        
        Set rsTmp = Get������Ŀ��¼(mlng��Ѫ;��)
        str��Ѫ;�� = rsTmp!���� & "" ' Get��Ŀ����(mlng��Ѫ;��)
        int���� = Val(rsTmp!ִ�з��� & "")
        If mlngUpdateAdvice <> 0 Then
            'ȡ�������
            strSQL = "Select �������,��鷽�� From ����ҽ����¼ where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            lng������� = Val(rsTmp!������� & "")
            int��鷽�� = Val(rsTmp!��鷽�� & "")
            
            '�޸�ҽ����ɾ�������²���
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & mlngUpdateAdvice & ",1)"
        Else
            If mblnSpareBloood = True Then
                int��鷽�� = 0
            Else
                int��鷽�� = 1
            End If
        End If
        If lng������� = 0 Then
            'ȡ�������
            strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            lng������� = Val(rsTmp!������� & "")
        End If
        
        int���� = IIF(cboInfo(cbo��Ѫ����).ListIndex <> 1, 0, 1)
        dblTmp = GetBloodTotalByML
        str���״̬ = GetBloodVerifyState(int������Դ, mlng����ID, IIF(int������Դ = 2, mlng��ҳID, mlng�Һ�ID), txtInfo(txtԤ����Ѫʱ��).Text, dblTmp, int����, int����, CInt(mbytBaby), mlngUpdateAdvice)
        If str���״̬ = "" Then str���״̬ = "NULL"
        '��Ѫҽ��
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngҽ��ID & ",NULL," & lngҽ����� & "," & int������Դ & "," & mlng����ID & "," & strTmp��ҳID & "," & mbytBaby & ",1,1,'K'," & mlng��Ѫ��ĿID & _
                                 ",NULL,NULL,NULL," & ZVal(txtInfo(txtԤ����Ѫ��).Text) & ",'" & FormatAdviceContext(str��Ŀ����, str��Ѫ;��) & _
                                 "'," & IIF(txtInfo(txt��ע).Text = "", "NULL", "'" & txtInfo(txt��ע).Text & "'") & ",'" & Format(txtInfo(txtԤ����Ѫʱ��).Text, "yyyy-MM-dd HH:mm:ss") & "','һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "Null", cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex)) & _
                                 "," & IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, "5", mlngִ�п�������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp�Һŵ� & "," & ZVal(mlngǰ��ID) & "," & IIF(int��鷽�� = 0, "NULL", "'" & int��鷽�� & "'") & ",0,NULL," & IIF(mstrժҪ��Ѫ = "", "null", "'" & mstrժҪ��Ѫ & "'") & ",'" & UserInfo.���� & "',Null,Null,'" & cboInfo(cbo��ѪĿ��).Text & "'," & str���״̬ & "," & lng������� & ")"
        
        '��Ѫ;��
        lng���ID = lngҽ��ID
        lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")        '��ȡҽ��ID
        lngҽ����� = lngҽ����� + 1
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngҽ��ID & "," & lng���ID & "," & lngҽ����� & "," & int������Դ & "," & mlng����ID & "," & strTmp��ҳID & _
                                 "," & mbytBaby & ",1,1,'E'," & mlng��Ѫ;�� & ",NULL,NULL,NULL,Null,'" & str��Ѫ;�� & "'," & IIF(str���� = "", "NULL", "'" & str���� & "'") & ",NULL,'һ����',NULL,NULL,NULL,NULL,0," & _
                                 IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "Null", cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)) & _
                                 "," & IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, "5", mlng��Ѫִ������) & "," & IIF(mbln��¼, 2, cboInfo(cbo��Ѫ����).ListIndex) & _
                                 ",to_date('" & Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                                 ",NULL," & mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
                                 "to_date('" & Format(IIF(curDate > CDate(txtInfo(txt��������).Text), txtInfo(txt��������).Text, curDate), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                                 strTmp�Һŵ� & "," & ZVal(mlngǰ��ID) & ",NULL,0,NULL," & IIF(mstrժҪ;�� = "", "null", "'" & mstrժҪ;�� & "'") & ",'" & UserInfo.���� & "',Null,NULL,''," & str���״̬ & "," & lng������� & ")"
    End If
    '��Ѫ����������Ŀ
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_��Ѫ�����¼_Insert(" & lng���ID & "," & chkWait.value & ",'" & cboInfo(cbo��Ѫ����).Text & "','" & cboInfo(cbo��ѪĿ��).Text & "'," & cboInfo(cbo��Ѫ����).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optHistory(2).value, 0, 1) & "," & IIF(optHistory(4).value, 0, 1) & ",'" & txtInfo(txt��) & "/" & txtInfo(txt��) & "'," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo��ѪѪ��).ListIndex & "," & cboInfo(cboRHD).ListIndex & "," & IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, "Null")) & "," & IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, "Null")) & ",'" & GetBloodInfo & "')"
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
    If Trim(txtInfo(txt�����Ϣ).Text) <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lng���ID & ",'���뵥���',null,1,null,'" & txtInfo(txt�����Ϣ).Text & "')"
    End If
    '�������ݲ���ҽ�����븽����Ŀ
    str��Ŀ���� = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
            str��Ŀ���� = IIF(str��Ŀ���� = "", "", str��Ŀ���� & Space(2)) & vsfBlood.TextMatrix(i, COL_P_����) & ":" & IIF(vsfBlood.TextMatrix(i, COL_P_����Ѫ��) = "", "", vsfBlood.TextMatrix(i, COL_P_����Ѫ��) & vsfBlood.TextMatrix(i, COL_P_����RH)) & " " & vsfBlood.TextMatrix(i, COL_P_������) & vsfBlood.TextMatrix(i, COL_P_��λ)
        End If
    Next
    If str��Ŀ���� <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lng���ID & ",'������Ŀ',null,2,null,'" & str��Ŀ���� & "')"
    End If
        
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    '��Ѫҽ���޸ģ���ɾ��������
    If int��鷽�� = 1 And InStr(1, ",3,4,", "," & mintType & ",") = 0 And mlngUpdateAdvice <> 0 And Not (mintType = 1 And mblnUseBloodSend = True) Then
        If InitObjBlood = True Then
            If gobjPublicBlood.AdviceOperation(IIF(mint���� = 0, pסԺҽ���´�, p����ҽ���´�), mlngUpdateAdvice, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "Ѫ�⹫����������ʧ�ܣ���ϸ��Ϣ��" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "Ѫ�⹫����������ʧ�ܣ����飡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    'ҽ����ع���ִ��
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '��Ѫҽ������
    If int��鷽�� = 1 And InStr(1, ",3,4,", "," & mintType & ",") = 0 And Not (mintType = 1 And mblnUseBloodSend = True) Then
        If InitObjBlood = True Then
            If gobjPublicBlood.AdviceOperation(IIF(mint���� = 0, pסԺҽ���´�, p����ҽ���´�), lng���ID, 0, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "Ѫ�⹫����������ʧ�ܣ���ϸ��Ϣ��" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "Ѫ�⹫����������ʧ�ܣ����飡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False

    mlngUpdateAdvice = lng���ID
    
    Call SetCommandBarPara(conMenu_Tool_Archive, IIF(mblnSpareBloood = True, 1, 2), mlngUpdateAdvice)
    
    If mint���� = 0 Then
        If str���״̬ = "NULL" Or str���״̬ = "4" Then
            Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, Trim(txtInfo(txt����).Text), Trim(txtInfo(txtסԺ��).Text), , IIF(mlng�������� = 1, 1, 2), _
                mlng��ҳID, mlng����ID, , mlng���˿���id, "", , Trim(txtInfo(txt����).Text), lng���ID, int����, 1, "K", "", UserInfo.����, _
                Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , , "")
        ElseIf str���״̬ = "1" Then
            Call ZLHIS_CIS_Audit("ZLHIS_CIS_030", mclsMipModule, mlng����ID, Trim(txtInfo(txt����).Text), Trim(txtInfo(txtסԺ��).Text), , IIF(mlng�������� = 1, 1, 2), _
                mlng��ҳID, mlng����ID, , mlng���˿���id, "", , Trim(txtInfo(txt����).Text), lng���ID, UserInfo.����, _
                Format(txtInfo(txt��������).Text, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , , "")
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
    If Index = cbo���� Then Call zlControl.TxtSelAll(cboInfo(Index))
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
    If Index = cbo��ѪĿ�� Then
        If zlCommFun.ActualLen(cboInfo(Index).Text) > 50 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then KeyAscii = 0
    ElseIf Index = cbo���� And KeyAscii <> vbKeyReturn And KeyAscii <> 8 And KeyAscii <> vbKeyTab Then
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
'���ܴ�ӡԤ�����뵥
'������intType:1-Ԥ����2-��ӡ
    '�ж������δ�������ȱ����ٴ�ӡ
    Dim strReportName As String
    If mintType <> 2 Then
        If mblnChange Then
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Else
            '��������ã�����ҽ���Ƿ����
            If CheckData = False Then Exit Sub
        End If
    End If
    If BloodApplyPrintCheck(mlngUpdateAdvice, IIF(1 = mint����, 1, 2), IIF(mblnSpareBloood = True, 1, 2), intType - 1) = False Then Exit Sub
    strReportName = IIF(mblnSpareBloood = False, "ZL1_INSIDE_1254_17_2", "ZL1_INSIDE_1254_17_1")
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportName, Me, "ҽ��ID=" & mlngUpdateAdvice, intType)
End Sub

Private Sub cboInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = cbo���� Then Call cboInfo_Click(Index)
End Sub

Private Sub cboInfo_LostFocus(Index As Integer)
    If Index = cboRHD Then Call cboInfo_Click(Index)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnChange As Boolean
    Dim lngPreUpdateAdivice As Long
    Select Case Control.ID
        Case conMenu_Tool_Archive * 10# + 1, conMenu_Tool_Archive * 10# + 2
            'ֻ��mint���ó���=0������ģʽ�������л������Ա���ı�ֱ�ӵ���SaveData
            Me.Tag = ""
            blnChange = mblnChange
            lngPreUpdateAdivice = mlngUpdateAdvice
            If blnChange = True Then
                If MsgBox("��ǰ���뵥�Ѿ������˵�����δ���棬�����Ƿ���Ҫ���棿", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes Then
                    '����
                    If CheckData = False Then Exit Sub
                    mblnOK = SaveData
                    If mblnOK = False Then Exit Sub
                End If
                mblnChange = False
            End If
            
            mblnSpareBloood = Control.ID = (conMenu_Tool_Archive * 10# + 1)
            mblnNewSpareBloood = mblnSpareBloood And mintType = 0
            'Control.Category �л���ҳ�������ID,lngPreUpdateAdivice �л�ǰҳ�������ID(����ʱ�洢Control.Categoryֵ)
            '���������л�ҳ�������£�ֻ������ҳ�涼�����������״̬����û�иı������ֻ�ǵ���ҳ��ؼ�λ��,��������¼��ؿؼ���ˢ������
            If Not (lngPreUpdateAdivice = 0 And Val(Control.Category) = 0 And blnChange = False) Then
                '���¼���ҳ������
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
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '����
            If CheckData = False Then Exit Sub
            If mint���ó��� = 0 Then
                mblnOK = SaveData
            Else
                mblnOK = SaveCacheData
            End If
            If Control.ID = conMenu_Edit_SaveExit Then
                Unload Me
            Else
                lblInfo(lbl24H��Ѫ��) = "24Сʱ����Ѫ��������" & GetBloodCapacity(IIF(mint���� = 0, 2, 1), mlng����ID, IIF(mint���� = 0, mlng��ҳID, mlng�Һ�ID), zlDatabase.Currentdate, True, CInt(mbytBaby)) & "ML"
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
                Control.Caption = IIF(mblnSpareBloood = True, "��Ѫ���뵥��", "ȡѪ֪ͨ����")
                Control.Checked = Control.Enabled
            End If
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
        Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview
            blnVisible = ((mint���뵥��ӡģʽ = 0 And InStr(GetInsidePrivs(pסԺҽ���´�), ";��Ѫ���뵥;") > 0) Or mint���� = 1) And mint���ó��� = 0
            If mint���뵥��ӡģʽ = 0 And mint���� = 0 Then
                If mintPState = ps��Ժ Then blnVisible = False
            End If
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub chkWait_Click()
    If chkWait.value = 1 Then
        txtInfo(txt�����Ϣ).Text = "����"
        txtInfo(txt�����Ϣ).Locked = True
        cmdInfo.Enabled = False
        mstr���IDs = ""
    Else
        txtInfo(txt�����Ϣ).Text = ""
        txtInfo(txt�����Ϣ).Locked = False
        cmdInfo.Enabled = True
    End If
    txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
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
    dtpDate.ZOrder 0
    dtpDate.SetFocus
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Call TxtGetInfo(Index, 1)
End Sub

Private Sub cmdInfo_Click()
    Dim str��� As String
    Dim lng����ID As Long
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng�������� = 1, 1260, 1261), mclsMipModule)
    End If
    lng����ID = IIF(mint���� = 0, mlng��ҳID, mlng�Һ�ID)
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng����ID, lng����ID, IIF(mlng�������� = 1, 1, 2), mlng���˿���id, mstr���IDs, str���, 0, mlngUpdateAdvice) Then
        txtInfo(txt�����Ϣ).Text = str���
        txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
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
    Dim rsBlood As New ADODB.Recordset
        
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    strInDate = mstr��Ժʱ��
    'סԺ���ϵ���ʱ�������¼��
    If mint���� = 0 Then
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    
    
        strInDate = ""
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            If mdatTurn <> CDate(0) Then strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
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
    Else
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵ľ���ʱ�� " & strInDate & " ��"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, ",1,4,", "," & mintType & ",") = 0 Then
        If InitObjBlood = True Then
            If gobjPublicBlood.GetPrepareBloodRs(mlngUpdateAdvice, rsBlood) = True Then
                '��Ѫҽ���Ѿ���Ѫ��������ʱ�䲻�ܴ��ڷ�Ѫʱ��
                If Val(rsBlood!��¼���� & "") = 2 And Val(rsBlood!��¼״̬ & "") = 1 And IsDate(rsBlood!���ʱ�� & "") Then
                    If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(rsBlood!���ʱ��, "yyyy-MM-dd HH:mm") Then
                        strMsg = "ҽ���Ŀ�ʼִ��ʱ��ӦС��Ѫ�ⷢѪ��ʱ��" & Format(rsBlood!���ʱ��, "yyyy-MM-dd HH:mm") & " ��"
                        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                '��Ѫҽ������Ѿ����գ��������¼����ܴ��ڽ���ʱ��(��Ҫ��������)
                ElseIf Val(rsBlood!��¼���� & "") = 1 And Val(rsBlood!��¼״̬ & "") = 1 And IsDate(rsBlood!����ʱ�� & "") Then
                    If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(rsBlood!����ʱ��, "yyyy-MM-dd HH:mm") Then
                        strMsg = "ҽ���Ŀ�ʼִ��ʱ��ӦС��Ѫ����յ�ʱ��" & Format(rsBlood!����ʱ��, "yyyy-MM-dd HH:mm") & " ��"
                        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
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
            strMsg = "��Ѫʱ�䲻��С������ʱ�䡣"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check����ʱ�� = True
End Function

Private Sub cmd��������_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    
    If Trim(txtInfo(txt��ע).Text) = "" Then
        MsgBox "�������������ݡ�", vbInformation, gstrSysName
        If txtInfo(txt��ע).Enabled Then txtInfo(txt��ע).SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From �������� Where ����=[1] And (��Ա=[2] Or ��Ա is null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtInfo(txt��ע).Text), UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        MsgBox "�����������Ѿ��ڳ��������С�", vbInformation, gstrSysName
        If txtInfo(txt��ע).Enabled Then txtInfo(txt��ע).SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(txtInfo(txt��ע).Text, CByte(Val(zlDatabase.GetPara("���뷽ʽ"))))
    strSQL = "zl_��������_Insert('" & Replace(txtInfo(txt��ע).Text, "'", "''") & "','" & strSQL & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    MsgBox "������Ϊ�������С�", vbInformation, gstrSysName
    If txtInfo(txt��ע).Enabled Then txtInfo(txt��ע).SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdҽ������_Click()
    Call ReasonSelect
End Sub

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
    mstr���IDs = ""
    mstrLISAboRHCode = ""
    If Me.Tag <> "GOTO" Then mblnOK = False
    mbln���Ѷ��� = True
    vsLIS.Rows = 0
    If mint���� = 0 Then mint���뵥��ӡģʽ = Val(zlDatabase.GetPara("��Ѫ���뵥��ӡģʽ", glngSys, pסԺҽ������, "1"))
    
    strPar = zlDatabase.GetPara("��Ѫ����ע������", glngSys, IIF(mint���� = 0, pסԺҽ���´�, p����ҽ���´�), "")
    lblInfo(lblע��).Caption = Trim(strPar)
    lblInfo(lblע��).Visible = Trim(strPar) <> ""
    
    '��ѡ������Ѫ���Ұ�װ��Ѫ��
    mblnNewSpareBloood = mblnSpareBloood And mintType = 0
    
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
        If mintType = 1 Then
            If InitObjBlood(True) = True Then
                If gobjPublicBlood.GetPrepareBloodRs(mlngUpdateAdvice, rsBlood) = True Then
                    '�޸���Ѫҽ���������Ѫ���Ѿ���Ѫ��������:��Ѫ�ɷ֡�ִ�п��ҡ�Ԥ����Ѫ��
                    If Val(rsBlood!��¼���� & "") = 2 And Val(rsBlood!��¼״̬ & "") = 1 Then
                        mblnUseBloodSend = True
                        SetControlEnabled txtGet(txtԤ����Ѫ�ɷ�), False
                        SetControlEnabled cboInfo(cboִ�п���), False
                        SetControlEnabled txtInfo(txtԤ����Ѫ��), False
                        vsfList.Editable = flexEDNone
                        vsfBlood.Editable = flexEDNone
                        SetControlEnabled txtInfo(txt��������), True
                        SetControlEnabled cmdDate(cmd��������), True
                    End If
                End If
            End If
        End If
    ElseIf mintType = 3 Then
        'ֻ�ܵ�������Ѫ�ɷ֣�����������ʱ�䣬��Ѫʱ�䣬ִ�п��ң���Ѫ;������Ѫִ�п��ң���Ѫ�������������
        SetControlEnabled txtInfo(txt��������), False
        SetControlEnabled cmdDate(cmd��������), False
        SetControlEnabled txtInfo(txtԤ����Ѫʱ��), False
        SetControlEnabled cmdDate(cmdԤ����Ѫʱ��), False
        SetControlEnabled txtGet(txtԤ����Ѫ�ɷ�), False
        SetControlEnabled txtGet(txt��Ѫ;��), False
        SetControlEnabled cmdGet(txt��Ѫ;��), False
        SetControlEnabled txtInfo(txtԤ����Ѫ��), False
        SetControlEnabled cboInfo(cboִ�п���), False
        SetControlEnabled cboInfo(cbo��Ѫִ��), False
        SetControlEnabled cboInfo(cbo��Ѫ����), False
        SetControlEnabled cboInfo(cbo��ѪĿ��), False
        SetControlEnabled cboInfo(cbo��Ѫ����), False
    End If
    mblnChange = mintType = 4
    If Me.Visible = False Then Call InitCommandBar
    If InitInfo = False Then Exit Sub
    Call LoadData
    Call SetFaceEnabledFalse
    Call SetFormNature
    If mbln��¼ Then SetControlEnabled cboInfo(cbo��Ѫ����), False
    '���˻�����Ϣ�����Ա༭
    SetControlEnabled txtInfo(txt�Ա�), False
    SetControlEnabled txtInfo(txt����), False
    SetControlEnabled txtInfo(txt����), False
    '��ʼ��opt�ؼ�
    For i = 2 To 5
        optHistory(i).Enabled = IIF(optHistory(1).value = True, True, False) And optHistory(0).Enabled
    Next
    If optHistory(0).value = True Then
        optHistory(2).value = True
        optHistory(4).value = True
    End If
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

Private Sub SetControlContent(Optional ByVal blnClearPatiInfo As Boolean = True)
'��ս���ؼ���Ϣ
    Dim objControl As Object
    For Each objControl In Me.Controls
        Select Case TypeName(objControl)
            Case "TextBox"
                If objControl.Name = "txtInfo" Then
                    If blnClearPatiInfo = True Or (blnClearPatiInfo = False And InStr(1, "," & txt�����Ϣ & "," & txt�� & "," & txt�� & "," & txtԤ����Ѫʱ�� & "," & txt��ע & ",", "," & objControl.Index & ",") <> 0) Then
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
    
    lblInfo(lblע��).Caption = ""
    lblInfo(lbl������ʷ������Ŀ).Caption = ""
    lblInfo(lbl24H��Ѫ��).Caption = ""
    vsLIS.Rows = 0
    vsLIS.Tag = ""
    '����������������Ҫ���������ʵ�����������
    vsfBlood.Rows = 1
    vsfBlood.Rows = 2
    vsfList.Rows = 1
    mlng��Ѫ��ĿID = mlngPre��Ѫ��ĿID
    mlng��Ѫ;�� = 0
    mlng��Ѫִ������ = 0
    mlngִ�п������� = 0
    mlng¼������ = 0
End Sub

Private Function InitInfo(Optional blnFormLoad As Boolean = True) As Boolean
'���ܣ���ʼ�����˵�
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim curDate As Date
    Dim lng�÷�ID As Long
    Dim lngִ�п���ID As Long
    Dim strMsg As String, strFilter As String
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    cboInfo(cboִ�п���).Clear
    If blnFormLoad = True Then
        '���̶ֹ����ݵ�������
        Call Cbo.LoadFromList(cboInfo(cbo��ѪѪ��), Array(" ", "A", "B", "O", "AB", "����", "δ��"), 0)
        Call Cbo.LoadFromList(cboInfo(cbo��Ѫ����), Array("��ͨ", "����"))
        Call Cbo.SetIndex(cboInfo(cbo��Ѫ����).hwnd, 0)
        Call Cbo.LoadFromList(cboInfo(cboRHD), Array(" ", "-", "+"), 0)
        Call Cbo.LoadFromList(cboInfo(cbo����), Array("15", "30", "����", "��ѹ"))
        
        '��Ѫ����
        strSQL = "Select ����,ȱʡ��־ from ��Ѫ����  order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            With cboInfo(cbo��Ѫ����)
                .Clear
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!���� & ""
                    If Val(rsTmp!ȱʡ��־ & "") = 1 Then
                        .ListIndex = .ListCount - 1
                    End If
                    rsTmp.MoveNext
                Next
            End With
        End If
        Set rsTmp = Nothing
        
        strSQL = "select ����,ȱʡ��־ from ��Ѫ����  order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            With cboInfo(cbo��Ѫ����)
                .Clear
                .AddItem " "
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!���� & ""
                    
                    If Val(rsTmp!ȱʡ��־ & "") = 1 Then
                        .ListIndex = .ListCount - 1
                    End If
                    
                    rsTmp.MoveNext
                Next
                If .ListIndex = -1 Then .ListIndex = 1
            End With
        End If
        Set rsTmp = Nothing
        
        strSQL = "select ���� from ��ѪĿ�� order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
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
        
        'txtInfo(txtԤ����Ѫʱ��).Text = Format(curDate, "YYYY-MM-DD HH:mm")
        txtInfo(txtԤ����Ѫʱ��).Text = ""
        txtInfo(txtԤ����Ѫʱ��).Tag = txtInfo(txtԤ����Ѫʱ��).Text
        txtInfo(txt��������).Text = Format(curDate, "YYYY-MM-DD HH:mm")
        txtInfo(txt��������).Tag = txtInfo(txt��������).Text
    End If
    'ȱʡ�÷�(����ʱ���)
    If mintType = 0 Then
        strFilter = ""
        If mblnSpareBloood = True Then
            lng�÷�ID = Getȱʡ�÷�ID(9, IIF(mint���� = 0, 2, 1))
            strMsg = "û�п��õ���Ѫ�ɼ�����,���ȵ�������Ŀ���������ã�"
        Else
            strFilter = " And nvl(ִ�з���,0)=" & IIF(mblnSpareBloood = False, 1, 0) '��Ѫ;��
            lng�÷�ID = Getȱʡ�÷�ID(8, IIF(mint���� = 0, 2, 1), strFilter)
            strMsg = "û�п��õ���Ѫ;��,���ȵ�������Ŀ���������ã�"
        End If
    
        If lng�÷�ID = 0 Then
            MsgBox strMsg, vbInformation, gstrSysName
            Screen.MousePointer = 0
            If blnFormLoad = True Then Unload Me
            Exit Function
        Else
            Set rsTmp = Get������Ŀ��¼(lng�÷�ID)
            txtGet(txt��Ѫ;��).Text = rsTmp!���� & ""
            mlng��Ѫִ������ = NVL(rsTmp!ִ�п���, 0)
            txtGet(txt��Ѫ;��).Tag = txtGet(txt��Ѫ;��).Text
            mlng��Ѫ;�� = lng�÷�ID
            cboInfo(cbo��Ѫִ��).Enabled = True
            Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(cbo��Ѫִ��), "E", mlng��Ѫ;��, 0, _
                Val(rsTmp!ִ�п��� & ""), mlng���˿���id, mlng��������ID, 0, 1, IIF(mlng�������� = 1, 1, 2))
            If cboInfo(cbo��Ѫִ��).ListIndex = -1 And cboInfo(cbo��Ѫִ��).ListCount > 1 Then
                Call Cbo.SetIndex(cboInfo(cbo��Ѫִ��).hwnd, 0)
            Else
                '����ж����ȡĬ�ϵ�ִ�п���
                lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "E", mlng��Ѫ;��, 0, _
                        NVL(rsTmp!ִ�п���, 0), mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
                If lngִ�п���ID <> 0 Then
                    Call Cbo.Locate(cboInfo(cbo��Ѫִ��), lngִ�п���ID, True)
                End If
            End If
            If cboInfo(cbo��Ѫִ��).ListCount = 2 Then cboInfo(cbo��Ѫִ��).Enabled = False
            cboInfo(cbo��Ѫִ��).Tag = lng�÷�ID
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
'���ܣ���ȡ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    '��ȡ���������Ϣ
    txtInfo(txt��������).Text = IIF(mlng�������� = 1, "����", "סԺ")
    
    If mint���� = 0 Then
        If mbytBaby = 0 Then
            strSQL = "Select A.סԺ��, Nvl(C.����, A.����) ����, Nvl(C.�Ա�, A.�Ա�) �Ա�, Nvl(C.����, A.����) ����, B.���� As ����, C.��Ժ���� As ��ǰ����, C.��Ժ����, C.����,c.�ѱ�" & vbNewLine & _
                    "From ������Ϣ A, ���ű� B, ������ҳ C" & vbNewLine & _
                    "Where C.��Ժ����id = B.Id And A.����id = C.����id And A.��ҳid = C.��ҳid And C.����id = [1] And C.��ҳid = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        Else
            strSQL = "Select a.סԺ��, Nvl(q.Ӥ������, a.���� || '֮Ӥ' || q.���) ����, q.Ӥ���Ա� �Ա�, Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��) || '��' As ����, b.���� As ����," & vbNewLine & _
                    " c.��Ժ���� As ��ǰ����, c.��Ժ����, c.����,c.�ѱ�" & vbNewLine & _
                    "From ������Ϣ A, ���ű� B, ������ҳ C, ������������¼ Q" & vbNewLine & _
                    "Where a.����id = c.����id And a.��ҳid = c.��ҳid And A.����id = q.����id And A.��ҳid = q.��ҳid And c.��Ժ����id = b.Id And c.����id = [1] " & vbNewLine & _
                    " And c.��ҳid = [2] And q.��� = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mbytBaby)
        End If
        If rsTmp.RecordCount > 0 Then
            txtInfo(txtסԺ��).Text = rsTmp!סԺ�� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
            If txtInfo(txt�Ա�).Text = "��" Or mbytBaby <> 0 Then
                SetControlEnabled txtInfo(txt��), False
                SetControlEnabled txtInfo(txt��), False
            End If
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt����).Text = rsTmp!��ǰ���� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            mstr��Ժʱ�� = Format(rsTmp!��Ժ���� & "", "YYYY-MM-DD HH:mm")
            mint���� = Val(rsTmp!���� & "")
            mstr�ѱ� = rsTmp!�ѱ� & ""
        End If
    Else
        strSQL = "Select a.ID, A.����,A.�Ա�,A.����,a.no,a.�����,a.����,b.���� as ����,a.ִ��ʱ��,c.�ѱ�" & _
            " From ���˹Һż�¼ A,���ű� b,������Ϣ c " & _
            " Where a.����ID=c.����ID and A.NO=[1] And a.��¼����=1 And a.��¼״̬=1 And A.����ID+0=[2] and a.ִ�в���id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            lblInfo(lbl�Һŵ�).Caption = "�� �� ��"
            txtInfo(txt�Һŵ�).Text = rsTmp!NO & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
            If txtInfo(txt�Ա�).Text = "��" Or mbytBaby <> 0 Then
                SetControlEnabled txtInfo(txt��), False
                SetControlEnabled txtInfo(txt��), False
            End If
            txtInfo(txt����).Text = rsTmp!���� & ""
            lblInfo(lbl�����).Caption = "�� �� ��"
            txtInfo(txt�����).Text = rsTmp!����� & ""
            txtInfo(txt����).Text = rsTmp!���� & ""
            mint���� = Val(rsTmp!���� & "")
            mstr�ѱ� = rsTmp!�ѱ� & ""
            mstr��Ժʱ�� = Format(rsTmp!ִ��ʱ�� & "", "YYYY-MM-DD HH:mm")
            mlng�Һ�ID = Val(rsTmp!ID & "")
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
    Dim strSQL As String, i As Long, j As Long
    Dim str��� As String
    Dim rsTmpOther As ADODB.Recordset
    Dim strTmp As String
    Dim strItemName As String, strIDs As String
    Dim arrItem
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    '��ȡ���������Ϣ
    Call LoadPatiInfo

    If mintType = 0 Then
        '����������л�ģʽ����ֱ�Ӷ�ȡ��Ӧ����
        If Me.Tag = "GOTO" Then GoTo GoLoadData
        If mint���� = 0 Then
            '��ȡ�ϴ�ת��ʱ��
            strSQL = "Select ��ʼʱ�� From ���˱䶯��¼" & _
                " Where ��ʼʱ�� is Not NULL And ��ʼԭ��=3" & _
                " And ����ID=[1] And ��ҳID=[2] Order by ��ʼʱ�� desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", mlng����ID, mlng��ҳID)
            If rsTmp.RecordCount > 0 Then
                mstr�ϴ�ת��ʱ�� = Format(rsTmp!��ʼʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
        End If
        '�´���Ѫ����ʱ����ϡ���ѪĿ�ġ�Ѫ�͵�Ĭ��ȥ���һ�α�Ѫ�������Ϣ
        Call LoadLastPrepareBlood
    ElseIf mintType = 1 Or mintType = 3 Or mintType = 2 Or mintType = 4 Then
GoLoadData:
        If mintType = 4 Then 'ҽ�����״̬��δ����Ѫ�����¼
            'ֱ�ӷ�Ѫ��ȡ��Ӧ��Ѫҽ���������Ϣ
            strSQL = "Select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='��Ѫ����ID'"
            Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            If Not rsTmpOther.EOF Then
                Call LoadLastPrepareBlood(Val(rsTmpOther!���� & ""))
            Else
                mstr���IDs = GetAdviceDiag(mlngUpdateAdvice, str���)
                txtInfo(txt�����Ϣ).Text = str���
                strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
                Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
                If Not rsTmpOther.EOF Then
                    txtInfo(txt�����Ϣ).Text = rsTmpOther!���� & ""
                End If
                txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
                chkWait.value = IIF(txtInfo(txt�����Ϣ).Text = "����", 1, 0)
                
                'Ѫ�ͱ�Ѫ���뵥����û�У��Ӳ�����Ϣ�ӱ��л�ȡ
                If cboInfo(cbo��ѪѪ��).ListIndex <= 0 Then
                    strSQL = "Select ��Ϣֵ from ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, "ABO")
                    If Not rsTmpOther.EOF Then
                        Select Case "" & rsTmpOther!��Ϣֵ
                            Case "A", "A��"
                                cboInfo(cbo��ѪѪ��).ListIndex = 1
                            Case "B", "B��"
                                cboInfo(cbo��ѪѪ��).ListIndex = 2
                            Case "O", "O��"
                                cboInfo(cbo��ѪѪ��).ListIndex = 3
                            Case "AB", "AB��"
                                cboInfo(cbo��ѪѪ��).ListIndex = 4
                        End Select
                    End If
                End If
                If cboInfo(cboRHD).ListIndex <= 0 Then
                    strSQL = "Select ��Ϣֵ from ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, "RH")
                    If Not rsTmpOther.EOF Then
                        Select Case "" & rsTmpOther!��Ϣֵ
                            Case "-", "��"
                                cboInfo(cboRHD).ListIndex = 1
                            Case "+", "��"
                                cboInfo(cboRHD).ListIndex = 2
                        End Select
                    End If
                End If
            End If
        Else
            '�޸�
            '��ȡ��Ѫ�����Ϣ
            strSQL = _
                " Select �Ƿ����, ��Ѫ����, ��ѪĿ��, ��Ѫ����, ������Ѫʷ, ������Ѫ��Ӧʷ, ��Ѫ���ɼ�����ʷ, �в����, ��Ѫ������, �Ƿ�ǩ��ͬ����, �Ƿ�������, ��ѪѪ��, Rhd, ��Ѫ��Ѫ��, Hct, Alt, Hbsag," & vbNewLine & _
                "       ÷��, Ѫ�쵰��, ѪС��, Antihcv, Antihiv12" & vbNewLine & _
                " From ��Ѫ�����¼" & vbNewLine & _
                " Where ҽ��id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!�Ƿ���� & "") = 1 Then
                    txtInfo(txt�����Ϣ).Text = "����"
                    chkWait.value = 1
                Else
                    '��ȡ���
                    mstr���IDs = GetAdviceDiag(mlngUpdateAdvice, str���)
                    txtInfo(txt�����Ϣ).Text = str���
                    '�Ӹ����л�ȡ���������������Ը���Ϊ׼
                     strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
                     Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
                     If Not rsTmpOther.EOF Then
                         txtInfo(txt�����Ϣ).Text = rsTmpOther!���� & ""
                     End If
                End If
                txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
                chkWait.value = Val(rsTmp!�Ƿ���� & "")
                Call zlControl.CboSetText(cboInfo(cbo��Ѫ����), rsTmp!��Ѫ���� & "", True, "'")
                Call zlControl.CboSetText(cboInfo(cbo��ѪĿ��), rsTmp!��ѪĿ�� & "", True, "'")
                cboInfo(cbo��Ѫ����).ListIndex = Val(rsTmp!��Ѫ���� & "")
                optHistory(Val(rsTmp!������Ѫʷ & "")).value = True
                optHistory(IIF(Val(rsTmp!������Ѫ��Ӧʷ & "") = 1, 3, 2)).value = True
                optHistory(IIF(Val(rsTmp!��Ѫ���ɼ�����ʷ & "") = 1, 5, 4)).value = True
                If InStr(1, "" & rsTmp!�в����, "/") <= 0 Then
                    txtInfo(txt��).Text = ""
                    txtInfo(txt��).Text = ""
                Else
                    txtInfo(txt��).Text = Mid(rsTmp!�в����, 1, InStr(1, "" & rsTmp!�в����, "/") - 1)
                    If Not (txtInfo(txt��).Text = "" Or IsNumeric(txtInfo(txt��).Text)) Then
                        txtInfo(txt��).Text = ""
                    End If
                    txtInfo(txt��).Text = Mid(rsTmp!�в����, InStr(1, "" & rsTmp!�в����, "/") + 1)
                    If Not (txtInfo(txt��).Text = "" Or IsNumeric(txtInfo(txt��).Text)) Then
                        txtInfo(txt��).Text = ""
                    End If
                End If
                optPossession(Val(rsTmp!��Ѫ������ & "")).value = True
                If InStr(1, ",0,1,", "," & rsTmp!�Ƿ�ǩ��ͬ���� & ",") <> 0 Then
                    optConsent(Val(rsTmp!�Ƿ�ǩ��ͬ���� & "")).value = True
                End If
                If InStr(1, ",0,1,", "," & rsTmp!�Ƿ������� & ",") <> 0 Then
                    optAppraise(Val(rsTmp!�Ƿ������� & "")).value = True
                End If
                
                cboInfo(cbo��ѪѪ��).ListIndex = Val(rsTmp!��ѪѪ�� & "")
                cboInfo(cboRHD).ListIndex = Val(rsTmp!RHD & "")
            End If
        End If
        '��ȡѪҺ������Ŀ(���������Ŀ��¼����<=1���������κδ���)
        mstr��Ѫ��Ŀ = ""
        strItemName = ""
        strIDs = ""
        strSQL = "Select A.����,B.������ĿID,B.������,B.����Ѫ��,B.����RH,b.ѪҺ��Ϣ From ������ĿĿ¼ A,��Ѫ������Ŀ B where A.ID=B.������ĿID And B.ҽ��ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
        Do While Not rsTmp.EOF
            strIDs = IIF(strIDs = "", "", strIDs & ",") & rsTmp!������ĿID
            mstr��Ѫ��Ŀ = IIF(mstr��Ѫ��Ŀ = "", "", mstr��Ѫ��Ŀ & ";") & rsTmp!������ĿID & "," & rsTmp!������ & "," & rsTmp!����Ѫ�� & "," & rsTmp!����rh & IIF(rsTmp!ѪҺ��Ϣ & "" <> "", "," & rsTmp!ѪҺ��Ϣ, "")
            strItemName = IIF(strItemName = "", "", strItemName & "'") & rsTmp!����
        rsTmp.MoveNext
        Loop
        '��ȡҽ�������Ϣ����ѪҺҽ����
        strSQL = "Select A.ID,A.���ID,a.������־,a.��ҩ����,NVL(to_char(a.����ʱ��,'yyyy-MM-dd hh24:mi'),a.�걾��λ) as Ԥ����Ѫʱ��,a.��ʼִ��ʱ��,a.������ĿID," & _
                " a.ִ�п���ID,a.ִ������,a.�ܸ�����,B.���,B.��������,B.���㵥λ,B.���� as ��Ŀ����,b.ִ�з���,A.�������,A.���״̬,a.ҽ������" & vbNewLine & _
                " From ����ҽ����¼ A,������ĿĿ¼ B" & vbNewLine & _
                " Where a.������ĿID=B.ID And (A.id = [1] or A.���ID=[1])"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
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
                If cboInfo(cbo��ѪĿ��).Text = "" Then Call zlControl.CboSetText(cboInfo(cbo��ѪĿ��), rsTmp!��ҩ���� & "", True, "'")   '�ϵĵ���ѪĿ�Ĵ洢��ҽ������ҩ��������
                txtInfo(txtԤ����Ѫʱ��).Text = Format(rsTmp!Ԥ����Ѫʱ�� & "", "YYYY-MM-DD HH:mm")
                txtInfo(txtԤ����Ѫʱ��).Tag = txtInfo(txtԤ����Ѫʱ��).Text
                txtInfo(txt��������).Text = Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:mm")
                txtGet(txtԤ����Ѫ�ɷ�).Text = IIF(strItemName = "", rsTmp!��Ŀ���� & "", strItemName)
                txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
                txtGet(txtԤ����Ѫ�ɷ�).Tag = txtGet(txtԤ����Ѫ�ɷ�).Text
                mlng��Ѫ��ĿID = Val(rsTmp!������ĿID)
                
                Call Setִ�п���(Val(rsTmp!ִ������ & ""), Val(rsTmp!ִ�п���ID & ""))
                Call LoadLisResult(mlngUpdateAdvice)
                
                txtInfo(txtԤ����Ѫ��).Text = zl9ComLib.FormatEx(rsTmp!�ܸ����� & "", 5)
                txtInfo(txtNO).Text = rsTmp!������� & ""
                txtInfo(txt��ע).Text = rsTmp!ҽ������ & ""
                '�Ѿ����ͨ���Ĳ������޸ģ���Ѫ��ɻ�Ѫ��
                If Val(rsTmp!���״̬ & "") = 2 And mblnUseBloodSend = False Then mblnEditable = False
                If InStr(1, "," & strIDs & ",", "," & mlng��Ѫ��ĿID & ",") = 0 Then
                    mstr��Ѫ��Ŀ = mlng��Ѫ��ĿID & "," & txtInfo(txtԤ����Ѫ��).Text & "," & cboInfo(cbo��ѪѪ��).Text & "," & cboInfo(cboRHD).Text & IIF(mstr��Ѫ��Ŀ = "", "", ";" & mstr��Ѫ��Ŀ)
                End If
            End If
            rsTmp.Filter = "���ID=" & mlngUpdateAdvice
            If rsTmp.RecordCount > 0 Then
                txtGet(txt��Ѫ;��).Text = rsTmp!��Ŀ���� & ""
                txtGet(txt��Ѫ;��).Tag = txtGet(txt��Ѫ;��).Text
                mlng��Ѫ;�� = Val(rsTmp!������ĿID)
                If Not (rsTmp!��� = "E" And rsTmp!�������� = "9") Then
                    mblnNewSpareBloood = False
                    If rsTmp!��� = "E" And rsTmp!�������� = "8" Then
                        mblnSpareBloood = (Val(rsTmp!ִ�з��� & "") = 0) '������������ ��Ѫҽ����ִ�з���=1
                    End If
                Else
                    mblnNewSpareBloood = True
                    mblnSpareBloood = True '�������ΪE,��������=9�ľ��Ǳ�Ѫҽ��
                End If
                Call Set��Ѫִ��(Val(rsTmp!ִ������ & ""), Val(rsTmp!ִ�п���ID & ""))
                '��Ѫҽ������
                strTmp = rsTmp!ҽ������ & ""
                cboInfo(cbo����).Text = ""
                lblInfo(31).Visible = True
                If strTmp Like "*��/����" Then
                    If IsNumeric(Split(strTmp, "��/����")(0)) = True Then
                        cboInfo(cbo����).Text = Split(strTmp, "��/����")(0)
                    End If
                ElseIf strTmp = "��ѹ" Or strTmp = "����" Then
                    cboInfo(cbo����).Text = strTmp
                    lblInfo(31).Visible = False
                End If
            End If
        End If
        '��ȡǩ����¼
        If gintCA <> 0 And Mid(gstrESign, 2, 1) = "1" Then
            strSQL = "Select b.ǩ����,A.�������� From ����ҽ��״̬ A, ҽ��ǩ����¼ B Where a.ǩ��id = b.Id And a.ҽ��id = [1] And ��������=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
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
    
    Call LoadDataFromCache
    strIDs = ""
    arrItem = Split(mstr��Ѫ��Ŀ, ";")
    For i = 0 To UBound(arrItem)
        strIDs = strIDs & "," & Split(CStr(arrItem(i)), ",")(0)
    Next
    strIDs = Mid(strIDs, 2)

    If mlng��Ѫ��ĿID <> 0 Then
        If InStr(1, "," & strIDs & ",", "," & mlng��Ѫ��ĿID & ",") = 0 Then
            strIDs = mlng��Ѫ��ĿID & "," & strIDs
            mstr��Ѫ��Ŀ = mlng��Ѫ��ĿID & "," & txtInfo(txtԤ����Ѫ��).Text & "," & cboInfo(cbo��ѪѪ��).Text & "," & cboInfo(cboRHD).Text & IIF(mstr��Ѫ��Ŀ = "", "", ";" & mstr��Ѫ��Ŀ)
        End If
        Set rsTmp = Get������Ŀ��¼(mlng��Ѫ��ĿID, strIDs)
        strTmp = ""
        Do While Not rsTmp.EOF
            strTmp = strTmp & IIF(strTmp = "", "", "'") & rsTmp!����
            rsTmp.MoveNext
        Loop
        txtGet(txtԤ����Ѫ�ɷ�).Text = strTmp
        rsTmp.Filter = "ID=" & mlng��Ѫ��ĿID
        Call Setִ�п���(Val(rsTmp!ִ�п��� & ""))
        txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
        txtGet(txtԤ����Ѫ�ɷ�).Tag = txtGet(txtԤ����Ѫ�ɷ�).Text
        mlng¼������ = Val(rsTmp!¼������ & "")
        If mrsCard Is Nothing And mint���ó��� = 1 Then Call SetLisResult(strIDs)
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
    objBar.EnableDocking xtpFlagStretched
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    With objBar.Controls
        Set objMenuBar = .Add(xtpControlButtonPopup, conMenu_Tool_Archive, "���뵥��", "ѡ����Ѫ���뵥������"): 'objMenuBar.IconId = 807
        objMenuBar.Style = xtpButtonCaption
        objMenuBar.Flags = xtpFlagRightAlign
        Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_Archive * 10# + 1, "��Ѫ���뵥(&1)")
        Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_Archive * 10# + 2, "ȡѪ֪ͨ��(&2)")
    End With
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
        cboInfo(cboִ�п���).Clear: cboInfo(cboִ�п���).AddItem "-"
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
        cboInfo(cbo��Ѫִ��).Clear: cboInfo(cbo��Ѫִ��).AddItem "-"
        cboInfo(cbo��Ѫִ��).ListIndex = 0
    Else
        If cboInfo(cbo��Ѫִ��).ListIndex >= 0 And lngִ�п���ID = 0 Then
            lngTmp = cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex)
        ElseIf lngִ�п���ID <> 0 Then
            lngTmp = lngִ�п���ID
        End If
        
        Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(cbo��Ѫִ��), "E", mlng��Ѫ;��, 0, _
            lngִ�п���, mlng���˿���id, mlng��������ID, lngTmp, 1, IIF(mlng�������� = 1, 1, 2), , , , , , , , mlng��������)
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
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    
    '��Ѫ���벻���ûس�
    If Index <> txt��Ѫ;�� Then Exit Sub
    
    If mblnNewSpareBloood = True And mblnSpareBloood = True Then
        strSQL = " And A.���='E' And A.��������='9' "  '�ɼ���ʽ
    Else
        strSQL = " And A.���='E' And A.��������='8' And nvl(A.ִ�з���,0)=" & IIF(mblnSpareBloood = False, 1, 0) '��Ѫ;��
    End If
    
    strSQL = "Select Distinct A.ID,A.����,A.����,A.ִ�з��� as ִ�з���ID,A.���㵥λ,A.ִ�п��� as ִ�п���ID,A.¼������ as ¼������ID" & _
    " From ������ĿĿ¼ A,������Ŀ���� B" & _
    " Where A.ID=B.������ĿID" & _
    strSQL & "  And A.������� IN(" & IIF(mlng�������� = 1, "1,2", 2) & ",3)" & _
    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
    IIF(intType = 0, " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])", "") & _
    IIF(mlng�������� = 1, "", " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4]) Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))") & _
    Decode(gbytCode, 0, " And B.���� IN([3],3)", 1, " And B.���� IN([3],3)", "") & _
    " Order by A.����"
            
    vRect = zlControl.GetControlRect(txtGet(Index).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtGet(Index).Height, blnCancel, False, True, UCase(txtGet(Index).Text) & "%", _
        gstrLike & UCase(txtGet(Index).Text) & "%", gbytCode + 1, mlng���˿���id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ�ƥ�����Ŀ��", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtGet(Index))
        txtGet(Index).SetFocus: Exit Sub
    Else
        Call SetTxtBloodInfo(rsTmp, Index, True)
    End If
End Sub

Private Function SetTxtBloodInfo(ByVal rsTmp As ADODB.Recordset, Optional ByVal Index As Integer, Optional ByVal blnNextControl As Boolean = True) As Boolean
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim vMsg As VbMsgBoxResult
    Dim strName As String, strID As String
    
    On Error GoTo ErrHand
    
    If Index = txtԤ����Ѫ�ɷ� Then
        If rsTmp.RecordCount > 0 Then
            mlng��Ѫ��ĿID = Val(rsTmp!ID)
            mlng¼������ = Val(rsTmp!¼������ID & "")
            txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
            Call Setִ�п���(Val(rsTmp!ִ�п���ID & "")) '��Ѫѡ����Ʒ���Ե�һ��Ϊ׼ȷ��ִ�п���
        Else
            mlng��Ѫ��ĿID = 0
            mlng¼������ = 0
            txtInfo(txt��λ).Text = ""
        End If
        Do While Not rsTmp.EOF
            strName = IIF(strName = "", "", strName & "'") & rsTmp!����
            strID = IIF(strID = "", "", strID & ",") & rsTmp!ID
            strIDs = IIF(strIDs = "", "", strIDs & ",") & rsTmp!ID & ":" & IIF(Val(cboInfo(cboִ�п���).Tag & "") <> 0, Val(cboInfo(cboִ�п���).Tag & ""), "")
            rsTmp.MoveNext
        Loop
        txtGet(Index).Text = strName
        txtGet(Index).Tag = txtGet(Index).Text
        Call SetLisResult(strID)
        '������
        If strIDs <> "" Then
            strҽ������ = FormatAdviceContext(Replace(txtGet(txtԤ����Ѫ�ɷ�).Text, "'", ","), txtGet(txt��Ѫ;��).Text)
        End If
    ElseIf Index = txt��Ѫ;�� Then
        txtGet(Index).Text = rsTmp!���� & ""
        txtGet(Index).Tag = txtGet(Index).Text
        mlng��Ѫ;�� = Val(rsTmp!ID)
        Call Set��Ѫִ��(Val(rsTmp!ִ�п���ID & ""))
        '������
        If mlng��Ѫ;�� <> 0 Then
            strIDs = strIDs & "," & mlng��Ѫ;�� & ":"
            If Val(cboInfo(cbo��Ѫִ��).Tag & "") <> 0 Then
                strIDs = strIDs & Val(cboInfo(cbo��Ѫִ��).Tag & "")
            End If
        End If
    End If
    
    strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, IIF(mlng�������� = 0, 2, 1), "", strIDs, strҽ������)
    If strMsg <> "" Then
        If gintҽ������ = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln���Ѷ��� = False
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

Private Sub SetLisResult(ByVal str��Ѫ��ĿID As String)
'���ܣ���ʼ����Ѫ��Ŀ��Ӧ�ļ�����Ŀָ���񣨱�Ѫ���������������Ѫ��Ŀ��
    Dim rsLIS As ADODB.Recordset '��ǰ��Ѫ�ļ�����Ŀ
    Dim rs��� As ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim strResult As String, arr��Ŀ����() As String
    Dim strָ���� As String, strʱ�� As String
    Dim str������� As String
    Dim strTmp As String, strTmp1 As String
    Dim arrTmp1 As Variant
    Dim arrTmp2 As Variant
    Dim arrBloodID As Variant, strBloodapplyrate As String, arrTmp4 As Variant, blnAdd As Boolean
    Dim i As Long, j As Long, k As Long
    Dim lngCol As Long
    Dim arrTmp3 As Variant
    Dim blnָ����ʾ�� As Boolean, blnGet As Boolean
    Dim int��ʷ������� As Long, str��ʷ������� As String, str��ʷ��������
    Dim arrDay, arrItem
    Dim strHisResult As String, strLisInfo As String
    '������Ѫ��ʱ����Ѫ���벻��Ҫ�����
    If mblnSpareBloood = False Then Exit Sub
    
    On Error GoTo errH
    arrBloodID = Split(str��Ѫ��ĿID, ",")
    '130538:������Ŀ��ȡ���ξ�����֧��ָ������
    If UBound(arrBloodID) > 0 Then
        strSQL = "Select /*+ CARDINALITY(C 10) */ A.������ĿID,B.����,B.����,A.��ʷ������� from ��Ѫ������� A,������ĿĿ¼ B,Table(f_Num2list([1])) C " & _
            " Where A.������ĿID=B.ID And A.��ĿID=C.Column_Value Order by B.����"
        Set rsLIS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��Ѫ��ĿID)
    Else
        strSQL = "Select A.������ĿID,B.����,B.����,A.��ʷ������� from ��Ѫ������� A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.��ĿID=[1] Order by B.����"
        Set rsLIS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str��Ѫ��ĿID))
    End If
    strTmp = ""
    Do While Not rsLIS.EOF
        If InStr(1, "," & str������� & ",", "," & rsLIS!���� & ",") = 0 Then
            str������� = str������� & "," & rsLIS!����
        End If
        int��ʷ������� = IIF(Val("" & rsLIS!��ʷ�������) <= 0, 7, Val("" & rsLIS!��ʷ�������))
        If InStr(1, "," & strTmp & ",", "," & int��ʷ������� & ",") = 0 Then
            strTmp = strTmp & "," & int��ʷ�������
        End If
        rsLIS.MoveNext
    Loop
    str������� = Mid(str�������, 2)
    '��ʷ����ָ�갴����������ʾ
    strTmp = Mid(strTmp, 2)
    arrDay = Split(strTmp, ",")
    arrItem = Array()
    strLisInfo = ""
    For i = 0 To UBound(arrDay)
        rsLIS.Filter = ""
        str��ʷ������� = ""
        str��ʷ�������� = ""
        Do While Not rsLIS.EOF
            int��ʷ������� = IIF(Val("" & rsLIS!��ʷ�������) <= 0, 7, Val("" & rsLIS!��ʷ�������))
            If Val(arrDay(i)) = int��ʷ������� Then
                If InStr(1, "," & str��ʷ������� & ",", "," & rsLIS!���� & ",") = 0 Then
                    str��ʷ������� = str��ʷ������� & "," & rsLIS!����
                End If
                If InStr(1, "," & str��ʷ�������� & ",", "," & rsLIS!���� & ",") = 0 Then
                    str��ʷ�������� = str��ʷ�������� & ",[" & rsLIS!���� & "]"
                End If
            End If
            rsLIS.MoveNext
        Loop
        str��ʷ������� = Mid(str��ʷ�������, 2)
        str��ʷ�������� = Mid(str��ʷ��������, 2)
        strLisInfo = IIF(strLisInfo = "", "", strLisInfo & vbCrLf) & Val(arrDay(i)) & "���ڣ�" & str��ʷ��������
        ReDim Preserve arrItem(UBound(arrItem) + 1)
        arrItem(UBound(arrItem)) = str��ʷ�������
    Next
    
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
        
        If mint���� = 0 Then
            strTmp1 = ""
            If strTmp <> "" Then
                arrTmp1 = Split(strResult, "<split3>")
                For i = 0 To UBound(arrTmp1)
                    If Replace(Replace(CStr(arrTmp1(i)), "<split1>", ""), "<split2>", "") <> "" Then 'strResult�����<split1><split1><split3>����
                        arrTmp2 = Split(arrTmp1(i), "<split1>")
                        If arrTmp2(8) <> "" Then
                            strTmp1 = "�н��"
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
                        blnGet = (MsgBox("����סԺδ�ҵ���Ч�ļ���ָ�꣬�Ƿ���ȡ���ξ���" & Val(arrDay(0)) & "���ڵļ���ָ�ꣿ", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes)
                    Else
                        blnGet = (MsgBox("����סԺδ�ҵ���Ч�ļ���ָ�꣬�Ƿ���ȡ���ξ���ָ�������ڵļ���ָ�ꣿ" & vbCrLf & strLisInfo, vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbYes)
                    End If
                End If
                vsLIS.Tag = IIF(blnGet = True, "YES", "")
                If blnGet = True Then
                    strResult = ""
                    For i = 0 To UBound(arrItem)
                        str��ʷ������� = CStr(arrItem(i))
                        int��ʷ������� = Val(arrDay(i))
                        strHisResult = ""
                        strHisResult = mobjPublicLis.GetTransfusionApplyFor(str��ʷ�������, mlng����ID, IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mstr�Һŵ�, CInt(mbytBaby), 2, int��ʷ�������)
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
'            ָ��1<split1>���Ʊ���1<split1>��λ1<split1>��˽��Ŀ1<split1>ָ�����1<split1>������1<split1>Ӣ����1<split1>ȡֵ����1<split1>
                '������1<split2>�����־1<split2>�������1<split2>�������1<split2>�걾����1<split2>ָ�����ʱ��1<split3>
'            ָ��2<split1>���Ʊ���2<split1>��˽��Ŀ2<split1>ָ�����2<split1>������2<split1>Ӣ����2<split1>ȡֵ����2<split1>
              '  ������2<split2>�����־2<split2>�������2<split2>�������2<split2>�걾����2<split2>ָ�����ʱ��1<split3>
            '���¸�ֵ��strResult�����<split1><split1><split3>����
            arrTmp1 = Split(strResult, "<split3>")
            strTmp = "": strָ���� = "": strTmp1 = ""
            For i = 0 To UBound(arrTmp1)
                If Replace(Replace(CStr(arrTmp1(i)), "<split1>", ""), "<split2>", "") <> "" Then
                    strָ���� = Split(arrTmp1(i), "<split1>")(4)
                    If InStr(1, "'" & strTmp1 & "'", "'" & strָ���� & "'") = 0 Then 'ȥ���ظ���ָ��
                        strTmp = strTmp & IIF(strTmp = "", "", "<split3>") & CStr(arrTmp1(i))
                        strTmp1 = strTmp1 & IIF(strTmp1 = "", "", "'") & strָ����
                    End If
                End If
            Next i
            strResult = strTmp
            arrTmp1 = Split(strResult, "<split3>")
            
            strTmp = "": strָ���� = ""
            For i = 0 To UBound(arrTmp1)
                strָ���� = Split(arrTmp1(i), "<split1>")(5) 'ȡ����������Ŀ�� ������
                strʱ�� = "��"
                arrTmp2 = Split(arrTmp1(i), "<split1>")
                If arrTmp2(8) <> "" Then
                    If UBound(Split(arrTmp2(8), "<split2>")) >= 5 Then
                        strʱ�� = Split(arrTmp2(8), "<split2>")(5)
                        
                        If IsDate(strʱ��) Then
                            strʱ�� = Format(strʱ��, "YYYY-MM-DD HH:MM:SS")
                        Else
                            strʱ�� = "��"
                        End If
                    End If
                End If
                strTmp = strTmp & "," & strָ���� & "," & strʱ��
            Next
            
            strPar = Mid(strTmp, 2)
            arr��Ŀ���� = Split(txtGet(txtԤ����Ѫ�ɷ�).Text, "'")
            strBloodapplyrate = ""
            For i = 0 To UBound(arrBloodID)
                strSQL = "select Zl_Fun_Bloodapplyrate([1],[2]) as ָ�� from dual"
                Set rs��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(arr��Ŀ����(i)), strPar)
                
                If Not rs���.EOF Then
                    strTmp = rs���!ָ�� & ""
                Else
                    strTmp = ""
                End If
                
                '���ֻ��һ��ѪҺƷ���򱣳�ԭ�з�ʽ������ͳһ����Zl_Fun_Bloodapplyrate����ת��
                If UBound(arrBloodID) > 0 Then
                    If strTmp <> "" And strBloodapplyrate <> strTmp Then
                        strTmp1 = ""
                        arrTmp4 = Split(strTmp, ",")
                        For j = 0 To UBound(arrTmp4)
                            blnAdd = True
                            If CStr(arrTmp4(j)) <> "" And Not IsDate(CStr(arrTmp4(j))) And CStr(arrTmp4(j)) <> "��" Then
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
            '�����ԭֵ����˵�й���û������������Ҳ������������
            If strTmp <> strPar Then
                strResult = ""
                If strTmp <> "" Then
                    arrTmp3 = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp3)
                        If arrTmp3(i) <> "" And Not IsDate(CStr(arrTmp3(i))) And arrTmp3(i) <> "��" Then
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
                blnָ����ʾ�� = True
            End If
            
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
                
                If blnָ����ʾ�� Then
                    strTmp = arrTmp3(i)
                    If InStr(strTmp, "|") <> 0 Then
                        strTmp = Split(strTmp, "|")(1)
                    Else
                        strTmp = "1"
                    End If
                Else
                    strTmp = "1"
                End If
                
                '����ָ����
                If arrTmp2(8) <> "" And strTmp = "1" Then
                    arrTmp2 = Split(arrTmp2(8), "<split2>")
                    .TextMatrix(Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(0)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_�����־ + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(1)
                    .TextMatrix(Int(i / CON_LisResultCol), COL_����ο� + (i Mod CON_LisResultCol) * CON_LisResultCount) = arrTmp2(2)
                Else
                    'δ��ȡ�������ʾ����ҽ��¼��
                    .Cell(flexcpBackColor, Int(i / CON_LisResultCol), COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount) = COLEditBackColor
                End If
                
                lngCol = COL_ָ�������� + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 1755
                lngCol = COL_�����λ + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 500
                lngCol = 1 + COL_������ĿID + (i Mod CON_LisResultCol) * CON_LisResultCount
                If lngCol <> 29 Then .ColWidth(lngCol) = 50
                lngCol = COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount
                .ColWidth(lngCol) = 1120
            Next
            '116848,���ݼ���������ABO�� RH
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

Private Sub LoadLisResult(ByVal lngҽ��ID As Long, Optional ByVal strResult As String)
'���ܣ��޸�\�鿴���뵥ʱ������ҽ��ID����������д��ָ��
    Dim rsTmp As Recordset, strSQL As String
    Dim i As Long, j As Long, lngCol As Long
    Dim varCol As Variant
    Dim varRow As Variant
    Dim varFields As Variant
    
    '������Ѫ��ʱ����Ѫ���벻��Ҫ�����
    If mblnSpareBloood = False Then Exit Sub
    
    strSQL = "select ���,������ĿID,ָ�����,ָ��������,ָ��Ӣ����,ָ����,�����λ,�����־,����ο�,ȡֵ����,�Ƿ��˹���д from ��Ѫ������ Where ҽ��ID=[1] order by ���"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    If strResult <> "" Then
        varFields = Array("���", "������ĿID", "ָ�����", "ָ��������", "ָ��Ӣ����", "ָ����", "�����λ", "�����־", "����ο�", "ȡֵ����", "�Ƿ��˹���д")
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
        rsTmp.Sort = "���"
    End If
    
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
            
            lngCol = COL_ָ�������� + (i Mod CON_LisResultCol) * CON_LisResultCount
            .ColWidth(lngCol) = 1755
            lngCol = COL_�����λ + (i Mod CON_LisResultCol) * CON_LisResultCount
            .ColWidth(lngCol) = 500
            lngCol = 1 + COL_������ĿID + (i Mod CON_LisResultCol) * CON_LisResultCount
            If lngCol <> 29 Then .ColWidth(lngCol) = 50
            lngCol = COL_ָ���� + (i Mod CON_LisResultCol) * CON_LisResultCount
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
    mstrLISAboRHCode = ""
    Set mclsMipModule = Nothing
End Sub

Private Sub lblInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = lbl������ʷ������Ŀ Then
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
    If lblInfo(lbl������ʷ������Ŀ).Width > picHisItem.Width Then
        strInfo = lblInfo(lbl������ʷ������Ŀ).Tag
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
    '�ָ���Ϊ�����
    If txtGet(Index).Text <> txtGet(Index).Tag Then
        txtGet(Index).Text = txtGet(Index).Tag
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible And mblnDataLoad = False Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Dim intIme As Integer  '1-��,2-�ر�,0�����ùرջ�����뷢
    If Index = txtԤ����Ѫʱ�� Then
'        If txtInfo(Index).Text = "" Then txtInfo(Index).Text = txtInfo(txt��������).Text
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt�������� Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt�� Or Index = txt�� Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 2
    ElseIf Index = txt��ע Then
        zlControl.TxtSelAll txtInfo(Index)
        intIme = 1
    ElseIf Index = txt�����Ϣ Then
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
        Case txt��, txt��
            If InStr("1234567890", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            Else
                If InStr("1234567890", Chr(KeyAscii)) > 0 Then
                    If Val(txtInfo(Index).Text) = 0 Then txtInfo(Index).Text = ""
                End If
            End If
        Case txt�����Ϣ
            Call zlControl.TxtCheckKeyPress(txtInfo(Index), KeyAscii, m�ı�ʽ)
        Case txt��ע
            If KeyAscii = vbKeyReturn Then
                If txtInfo(txt��ע).Text <> "" Then
                    Call ReasonSelect(txtInfo(txt��ע).Text)
                End If
            End If
    End Select
    If KeyAscii = vbKeyReturn Then Call SeekNextControl
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    If Index = txt��ע Or Index = txt�����Ϣ Then
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    If Index = txtԤ����Ѫʱ�� Then
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Tag) Then
                    '�ָ���Ϊ�������ȱʡΪ�ϴ���д��ʱ��
                    txtInfo(Index).Text = txtInfo(Index).Tag
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check����ʱ��(txtInfo(Index).Text, txtInfo(txt��������).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            txtInfo(Index).Tag = txtInfo(Index).Text
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
    ElseIf Index = txt�����Ϣ Then
        If txtInfo(Index).Tag <> txtInfo(Index).Text Then
            mstr���IDs = ""
        End If
    ElseIf Index = txt��ע Then
        If zlCommFun.ActualLen(txtInfo(Index).Text) > 100 Then
            MsgBox "�������ݲ������� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
            Call txtInfo_GotFocus(Index)
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cboInfo(cbo��λ).Enabled And cboInfo(cbo��λ).Visible Then
            cboInfo(cbo��λ).SetFocus
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
            If mblnSpareBloood = True Then '��Ѫ����
                '��ѪҺ�ҡ�- 800ml<Split2>A��+:400ml<Split3>B��+:400ml<Split1>��LWҽ������A��- 800ml<Split2>A��+:400ml<Split3>B��+:400ml
                strTmp = vsfBlood.TextMatrix(NewRow, COL_P_���)
                arrInfo = Split(strTmp, "<Split1>") '�����ⷿ��Ϣ
                For i = 0 To UBound(arrInfo)
                    arrItem = Split(arrInfo(i), "<Split2>") '�ֽ�ⷿ��Ѫ����Ϣ
                    .AddItem arrItem(0)   '�ⷿ��Ϣ
                    .IsSubtotal(.Rows - 1) = True
                    .RowOutlineLevel(.Rows - 1) = 0
                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                     arrItem = Split(arrItem(1), "<Split3>") '��ȡѪ�Ϳ����Ϣ
                     For j = 0 To UBound(arrItem)
                        .AddItem arrItem(j)   'Ѫ�ͼ�����
                        .IsSubtotal(.Rows - 1) = True
                        .RowOutlineLevel(.Rows - 1) = 1
                     Next j
                Next i
            Else '��Ѫ����
                '�����Ŀ<Split2>Ʒ��ID'�䷢��Ϣ'������
                '�����䷢��Ϣ��ʽΪ����Ѫ������400ml �ѷ�����0ml δ������400ml<Split4> ���200ml(δ��)  Ч��:2016-09-17 16:13<Split3>0<Split4> ���200ml(δ��)  Ч��:2016-08-14 11:17<Split3>0
                strTmp = vsfBlood.TextMatrix(NewRow, COL_P_���)
                If strTmp <> "" Then
                    arrInfo = Split(strTmp, "<Split2>")
                    If UBound(arrInfo) > 0 Then
                        arrCode = Split(Split(arrInfo(1), "'")(1), "<Split4>") '�ֽ���ѪѪҺ�͹����Ϣ
                        If UBound(arrCode) >= 0 Then
                            .AddItem arrCode(0)  'ѪҺ
                            .IsSubtotal(.Rows - 1) = True
                            .RowOutlineLevel(.Rows - 1) = 0
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                            For j = 1 To UBound(arrCode)
                                arrItem = Split(arrCode(j), "<Split3>")
                                .AddItem arrItem(0)  '�����Ϣ
                                .IsSubtotal(.Rows - 1) = True
                                .RowOutlineLevel(.Rows - 1) = 1
                                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = IIF(Val(arrItem(1)) = 0, &H80000008, &H8000000C)
                            Next
                        End If
                    End If
                End If
            End If
        Else 'ҽ��ѡ��Ʒ�ֺ���ɿ�����Ѫ�����õ�ѪҺ��Ϣ
            '��ʽ���շ�ID<Split1>Ѫ�����<Split1>���<Split1>Ч��,������Ѫ��¼֮����<Split4>���������ƴ��<Split3>����Ϊ�Ѿ�ѡ���ѪҺID
            strIDs = ""
            strTmp = vsfBlood.TextMatrix(NewRow, COL_P_���)
            If strTmp <> "" Then
                If InStr(1, strTmp, "<Split3>") <> 0 Then
                    arrInfo = Split(strTmp, "<Split3>") '���Ƚ������ѪҺ����ѡ���ѪҺ�ֿ�
                    strTmp = arrInfo(0)
                    strIDs = arrInfo(1)
                End If
                arrInfo = Split(strTmp, "<Split4>") '����ѪҺ��Ϣ
                For i = 0 To UBound(arrInfo)
                    .Rows = .Rows + 1
                    arrItem = Split(arrInfo(i), "<Split1>") '�ֽ�ѪҺ��Ϣ
                    .TextMatrix(.Rows - 1, COL_S_ID) = Val(arrItem(0))
                    .TextMatrix(.Rows - 1, COL_S_ѡ��) = ""
                    .TextMatrix(.Rows - 1, COL_S_���) = arrItem(1)
                    .TextMatrix(.Rows - 1, COL_S_���) = arrItem(2)
                    .TextMatrix(.Rows - 1, COL_S_Ч��) = Format(arrItem(3), "YYYY-MM-DD HH:mm")
                    blnSelect = InStr(1, "|" & strIDs & "|", "|" & .TextMatrix(.Rows - 1, COL_S_ID) & "|") <> 0
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_ѡ��) = img16.ListImages(IIF(blnSelect = True, "c1", "c0")).Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_ѡ��) = IIF(blnSelect = True, 1, 0)
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_S_Ч��) = blnSelect
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_S_Ч��) = IIF(blnSelect = True, &HC0E0FF, vbWhite)
                Next
                '������ʧЧ��ѪҺ����ǰ��
                .Cell(flexcpSort, .FixedRows, COL_S_Ч��, .Rows - 1, COL_S_Ч��) = 1
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
            If vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_ѡ��) = 1 Then
                If .Row < .FixedRows And .Rows > .FixedRows Then
                    blnNext = True
                    .Row = .FixedRows
                    .Col = COL_S_ѡ��
                    .ShowCell .FixedRows, COL_P_ѡ��
                ElseIf .Row < .Rows - 1 Then
                    blnNext = True
                    .Row = .Row + 1
                    .Col = COL_S_ѡ��
                    .ShowCell .Row, COL_S_ѡ��
                End If
            End If
            If blnNext = False Then
                If vsfBlood.Row < vsfBlood.Rows - 1 Then
                    blnNext = True
                    vsfBlood.Row = vsfBlood.Row + 1
                    vsfBlood.Col = COL_P_ѡ��
                    vsfBlood.ShowCell vsfBlood.Row, COL_P_ѡ��
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
            '����ѡ��ѪҺƷ�ֺ���ܸ���ѪҺ��Ϣ
            If vsfBlood.Row < vsfBlood.FixedRows Then Exit Sub
            If Val(vsfBlood.TextMatrix(vsfBlood.Row, COL_P_ID)) = 0 Then Exit Sub
            If vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_ѡ��) = 0 Then Exit Sub
            If .Row >= .FixedRows And .Editable <> flexEDNone Then
                If Val(.TextMatrix(.Row, COL_S_ID)) = 0 Then Exit Sub
                intValue = Val(.Cell(flexcpData, .Row, COL_S_ѡ��))
                Set .Cell(flexcpPicture, .Row, COL_S_ѡ��, .Row, COL_S_ѡ��) = img16.ListImages("c" & IIF(intValue = 1, "0", "1") & "").Picture
                .Cell(flexcpData, .Row, COL_S_ѡ��, .Row, COL_S_ѡ��) = IIF(intValue = 1, 0, 1)
                .Cell(flexcpFontBold, .Row, COL_S_ѡ��, .Row, COL_S_Ч��) = IIF(intValue = 1, False, True)
                .Cell(flexcpBackColor, .Row, COL_S_ѡ��, .Row, COL_S_Ч��) = IIF(intValue = 1, vbWhite, &HC0E0FF)
                strIDs = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_S_ѡ��, i, COL_S_ѡ��) = 1 Then
                        dblSum = dblSum + Val(.TextMatrix(i, COL_S_���))
                        strIDs = strIDs & "|" & Val(.TextMatrix(i, COL_S_ID))
                    End If
                Next
                If Left(strIDs, 1) = "|" Then strIDs = Mid(strIDs, 2)
                strTmp = vsfBlood.TextMatrix(vsfBlood.Row, COL_P_���)
                If InStr(1, strTmp, "<Split3>") <> 0 Then
                    arrInfo = Split(strTmp, "<Split3>")
                    strTmp = arrInfo(0)
                End If
                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_���) = strTmp & IIF(strIDs <> "", "<Split3>" & strIDs, "")
                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_������) = IIF(dblSum = 0, "", dblSum)
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
        If vsfList.Col = COL_S_ѡ�� And vsfList.MouseCol = COL_S_ѡ�� And Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��)) = 1 Then
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
    '������༭�κ�����
    Cancel = True
End Sub

Private Sub vsfBlood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Visible Then mblnChange = True
    If Col = COL_P_������ Then Call BloodSum
End Sub

Private Sub vsfBlood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_P_������ Then
        vsfBlood.EditSelStart = 0
        vsfBlood.EditSelLength = Len(vsfBlood.TextMatrix(Row, Col))
        '�ر����뷨
        On Error Resume Next
        Call zlCommFun.OpenIme
        If err <> 0 Then err.Clear
    End If
End Sub

Private Sub vsfBlood_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_P_ѡ�� Then Cancel = True
End Sub

Private Sub vsfBlood_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer
    Dim blnFind As Boolean
    Dim blnNext As Boolean
    If KeyCode = vbKeyReturn Then
        Select Case Col
            Case COL_P_������
                vsfBlood.TextMatrix(Row, Col) = vsfBlood.EditText
            Case COL_P_����Ѫ��, COL_P_����RH
                If vsfBlood.ColHidden(Col) = False Then
                    vsfBlood.TextMatrix(Row, Col) = vsfBlood.ComboItem(vsfBlood.ComboIndex)
                End If
        End Select
        For i = Col + 1 To vsfBlood.Cols - 1
            If (i = COL_P_����Ѫ�� Or i = COL_P_����RH Or i = COL_P_������) And vsfBlood.ColHidden(i) = False Then
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
                vsfBlood.Col = COL_P_ѡ��
                vsfBlood.ShowCell vsfBlood.Row, COL_P_ѡ��
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
            vsfBlood.Col = COL_P_ѡ��
            vsfBlood.ShowCell vsfBlood.FixedRows, COL_P_ѡ��
        ElseIf vsfBlood.Row <= vsfBlood.Rows - 1 Then
            For j = vsfBlood.Col + 1 To vsfBlood.Cols - 1
                If (j = COL_P_������ Or j = COL_P_����Ѫ�� Or j = COL_P_����RH) And vsfBlood.ColHidden(j) = False Then
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
                    vsfBlood.Col = COL_P_ѡ��
                    vsfBlood.ShowCell vsfBlood.Row, COL_P_ѡ��
                End If
            End If
        End If
        If blnNext = False Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If vsfBlood.Editable <> flexEDNone Then
            txt������Ϣ.Text = "Ʒ��:"
            txt������.Text = ""
            cboInfo(cbo��λ).ListIndex = -1
            If vsfBlood.Col = COL_P_ѡ�� And vsfBlood.Row >= vsfBlood.FixedRows Then
                intValue = Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��))
                Set vsfBlood.Cell(flexcpPicture, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_ѡ��) = img16.ListImages("c" & IIF(intValue = 1, "0", "1") & "").Picture
                vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_ѡ��) = IIF(intValue = 1, 0, 1)
                vsfBlood.Cell(flexcpFontBold, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_���) = IIF(intValue = 1, False, True)
                vsfBlood.Cell(flexcpBackColor, vsfBlood.Row, COL_P_ѡ��, vsfBlood.Row, COL_P_���) = IIF(intValue = 1, vbWhite, &HC0E0FF)
                
                If Val(vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��)) = 1 Then
                    '����Ѫ�ͺ�RH�̳���һ��
                    If vsfBlood.ColHidden(COL_P_����Ѫ��) = False Then
                        For i = vsfBlood.Row - 1 To vsfBlood.FixedRows Step -1
                            If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����Ѫ��) = vsfBlood.TextMatrix(i, COL_P_����Ѫ��)
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����RH) = vsfBlood.TextMatrix(i, COL_P_����RH)
                                Exit For
                            End If
                            If vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����Ѫ��) = "" Then
                                If InStr(1, ",A,B,O,AB,", "," & cboInfo(cbo��ѪѪ��).Text & ",") <> 0 Then
                                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����Ѫ��) = cboInfo(cbo��ѪѪ��).Text
                                End If
                            End If
                            If vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����RH) = "" Then
                                vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����RH) = cboInfo(cboRHD).Text
                            End If
                        Next
                    End If
                Else
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_������) = ""
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����Ѫ��) = ""
                    vsfBlood.TextMatrix(vsfBlood.Row, COL_P_����RH) = ""
                    'ҽ��ѡ��ѪҺ��ģʽ�����ѪҺƷ��ȡ��ѡ����ͬ��ȡ��֮ǰѡ���ѪҺ��Ϣ
                    If mblnSelectBlood = True Then
                        For i = vsfList.FixedRows To vsfList.Rows - 1
                            If Val(vsfList.Cell(flexcpData, i, COL_S_ѡ��, i, COL_S_ѡ��)) = 1 Then
                                vsfList.Cell(flexcpPicture, i, COL_S_ѡ��, i, COL_S_ѡ��) = img16.ListImages("c0").Picture
                                vsfList.Cell(flexcpData, i, COL_S_ѡ��, i, COL_S_ѡ��) = 0
                                vsfList.Cell(flexcpFontBold, i, COL_S_ѡ��, i, COL_S_Ч��) = False
                                vsfList.Cell(flexcpBackColor, i, COL_S_ѡ��, i, COL_S_Ч��) = vbWhite
                            End If
                        Next
                        If InStr(1, vsfBlood.TextMatrix(vsfBlood.Row, COL_P_���), "<Split3>") <> 0 Then
                            vsfBlood.TextMatrix(vsfBlood.Row, COL_P_���) = Split(vsfBlood.TextMatrix(vsfBlood.Row, COL_P_���), "<Split3>")(0)
                        End If
                    End If
                End If
                '����ѪҺ��ִ�п���
                With rsTmp
                    If .State = 1 Then .Close
                    .Fields.Append "ID", adBigInt
                    .Fields.Append "����", adVarChar, 20, adFldIsNullable
                    .Fields.Append "����", adVarChar, 60, adFldIsNullable
                    .Fields.Append "���㵥λ", adVarChar, 20, adFldIsNullable
                    .Fields.Append "ִ�з���ID", adBigInt
                    .Fields.Append "ִ�п���ID", adBigInt
                    .Fields.Append "¼������ID", adBigInt
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
                        If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
                            .AddNew
                            .Fields("ID") = Val(vsfBlood.TextMatrix(i, COL_P_ID))
                            .Fields("����") = IIF(vsfBlood.TextMatrix(i, COL_P_����) = "", Null, vsfBlood.TextMatrix(i, COL_P_����))
                            .Fields("����") = IIF(vsfBlood.TextMatrix(i, COL_P_����) = "", Null, vsfBlood.TextMatrix(i, COL_P_����))
                            .Fields("���㵥λ") = IIF(vsfBlood.TextMatrix(i, COL_P_��λ) = "", Null, vsfBlood.TextMatrix(i, COL_P_��λ))
                            .Fields("ִ�з���ID") = Val(vsfBlood.TextMatrix(i, COL_P_ִ�з���ID))
                            .Fields("ִ�п���ID") = Val(vsfBlood.TextMatrix(i, COL_P_ִ�п���ID))
                            .Fields("¼������ID") = Val(vsfBlood.TextMatrix(i, COL_P_¼������ID))
                            .Update
                            'iif(mid("" & 0.5,1,1)=".","0","") & 0.5������д����Ϊ�˱�֤С�ڵ�1��ֵ��������ʾǰ׺0
                            txt������Ϣ.Text = txt������Ϣ.Text & "[" & vsfBlood.TextMatrix(i, COL_P_����) & IIF(vsfBlood.TextMatrix(i, COL_P_������) <> "", "-" & IIF(Mid("" & vsfBlood.TextMatrix(i, COL_P_������), 1, 1) = ".", "0", "") & vsfBlood.TextMatrix(i, COL_P_������) & vsfBlood.TextMatrix(i, COL_P_��λ), "") & "]"
                            'ѡ��Ʒ��ʱȱʡ��λ���ã��������£�
                            '1��������õ�Ʒ�ֵ�λ�а���ML����ȱʡ���õ�λΪML
                            '2��������õ�Ʒ�ֵ�λ�в�����ML����ȱʡ��λΪ����һ��Ʒ�ֵĵ�λ
                            blnSetUnit = False
                            If cboInfo(cbo��λ).ListIndex = -1 Then
                                blnSetUnit = True
                            Else
                                If UCase(cboInfo(cbo��λ).List(cboInfo(cbo��λ).ListIndex)) <> "ML" And UCase(vsfBlood.TextMatrix(i, COL_P_��λ)) = "ML" Then
                                    blnSetUnit = True
                                End If
                            End If
                            If blnSetUnit = True Then
                                For j = 0 To cboInfo(cbo��λ).ListCount - 1
                                    If UCase(vsfBlood.TextMatrix(i, COL_P_��λ)) = UCase(cboInfo(cbo��λ).List(j)) Then
                                        Call zlControl.CboSetIndex(cboInfo(cbo��λ).hwnd, j)
                                        cboInfo(cbo��λ).Tag = j
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
                Call SetTxtBloodInfo(rsTmp, txtԤ����Ѫ�ɷ�, False)
                Call RsetBreedUnit
            End If
        End If
    End If
End Sub

Private Sub vsfBlood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = COL_P_������ Then
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
    If vsfBlood.Col = COL_P_ѡ�� And vsfBlood.MouseCol = COL_P_ѡ�� Then
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
        If Not ((Col = COL_P_������ Or Col = COL_P_����Ѫ�� Or Col = COL_P_����RH) And vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��) = 1 And vsfBlood.ColHidden(Col) = False) Then
            Cancel = True
            Exit Sub
        End If
    Else
        If Not (Col = COL_P_������ And vsfBlood.Cell(flexcpData, vsfBlood.Row, COL_P_ѡ��) = 1) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub vsLIS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Visible Then
        mblnChange = True
        '�༭������ABO��RHʱ����������ABO��RHѡ��ֵ����
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
            If .TextMatrix(NewRow, NewCol + (COL_ȡֵ���� - COL_ָ����)) <> "" Then
                '�ϰ���°��õķָ����ͬ���°��Ƕ��ţ��ϰ�ֺţ����¼��ݴ���
                strTmp = .TextMatrix(NewRow, NewCol + (COL_ȡֵ���� - COL_ָ����))
                strTmp = Replace(strTmp, ";", "|")
                strTmp = Replace(strTmp, ",", "|")
                .ComboList = strTmp & "|�Ѳ�δ�ر�"
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

Private Function SaveCacheData() As Boolean
'���ܣ���������
    Dim strResult As String
    Dim rsCard As ADODB.Recordset
    Dim curDate As Date
    Dim str������ĿSQL As String
    Dim str��Ϲ�����ϢSQL As String
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long, j As Long
    Dim var1 As Variant
    Dim var2 As Variant
    Dim str���� As String
    Dim str��Ŀ���� As String, str������ĿSQL As String
    
    If cboInfo(cbo����).Visible = True Then
        str���� = cboInfo(cbo����).Text
    End If
    If IsNumeric(str����) = True Then
        str���� = str���� & "��/����"
    End If
    
    var1 = Array()
    var2 = Array()
    '������Ŀ
    With vsLIS
        lngCount = 0
        For i = 0 To .Rows - 1
            For j = 0 To CON_LisResultCol - 1
                If Val(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) <> 0 Then
                    lngCount = lngCount + 1
                    strTmp = "Zl_��Ѫ������_Insert([���ID]," & lngCount & "," & ZVal(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) & ",'" & .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_ָ��Ӣ���� + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_�����λ + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_�����־ + (j * CON_LisResultCount)) & "','" & .TextMatrix(i, COL_����ο� + (j * CON_LisResultCount)) & "','" & _
                                 .TextMatrix(i, COL_ȡֵ���� + (j * CON_LisResultCount)) & "'," & IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0) & ")"
                    str������ĿSQL = str������ĿSQL & "<splitSQL>" & strTmp
                    
                    var1 = Array(lngCount, .TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ��Ӣ���� + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_�����λ + (j * CON_LisResultCount)), .TextMatrix(i, COL_�����־ + (j * CON_LisResultCount)), .TextMatrix(i, COL_����ο� + (j * CON_LisResultCount)), _
                        .TextMatrix(i, COL_ȡֵ���� + (j * CON_LisResultCount)), IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0))
                    strTmp = Join(var1, "<SplitCol>")
                    ReDim Preserve var2(UBound(var2) + 1)
                    var2(UBound(var2)) = strTmp
                End If
            Next
        Next
    End With
    strResult = Join(var2, "<SplitRow>")
    '��Ϲ�����Ϣ
    If mstr���IDs <> "" Then
        str��Ϲ�����ϢSQL = "Zl_�������ҽ��_Insert([���ID],'" & mstr���IDs & "')"
        str��Ϲ�����ϢSQL = str��Ϲ�����ϢSQL & "<splitSQL>" & "Zl_����ҽ������_Insert([���ID],'���뵥���',null,null,null,'" & txtInfo(txt�����Ϣ).Text & "',1)"
    ElseIf Trim(txtInfo(txt�����Ϣ).Text) <> "" Then
        str��Ϲ�����ϢSQL = "Zl_����ҽ������_Insert([���ID],'���뵥���',null,null,null,'" & txtInfo(txt�����Ϣ).Text & "',1)"
    End If
    
    '������ĿSQL
    '�������ݲ���ҽ�����븽����Ŀ
    str��Ŀ���� = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
            str��Ŀ���� = IIF(str��Ŀ���� = "", "", str��Ŀ���� & Space(2)) & vsfBlood.TextMatrix(i, COL_P_����) & ":" & IIF(vsfBlood.TextMatrix(i, COL_P_����Ѫ��) = "", "", vsfBlood.TextMatrix(i, COL_P_����Ѫ��) & vsfBlood.TextMatrix(i, COL_P_����RH)) & " " & vsfBlood.TextMatrix(i, COL_P_������) & vsfBlood.TextMatrix(i, COL_P_��λ)
        End If
    Next
    If str��Ŀ���� <> "" Then
        str������ĿSQL = "Zl_����ҽ������_Insert([���ID],'������Ŀ',null,2,null,'" & str��Ŀ���� & "')"
    End If
    
    If mrsCard Is Nothing Then
         Call InitCardRsBlood(mrsCard)
         mrsCard.AddNew
    End If
    
    With mrsCard
        !��Ѫ���� = cboInfo(cbo��Ѫ����).ListIndex
        !�ٴ����IDs = mstr���IDs
        !���� = chkWait.value
        !��Ѫ���� = cboInfo(cbo��Ѫ����).Text
        !��ѪĿ�� = cboInfo(cbo��ѪĿ��).Text
        !��Ѫ���� = cboInfo(cbo��Ѫ����).ListIndex
        !������Ѫʷ = IIF(optHistory(0).value, 0, 1)
        !������Ѫ��Ӧʷ = IIF(optHistory(2).value, 0, 1)
        !��Ѫ���ɼ�����ʷ = IIF(optHistory(4).value, 0, 1)
        !�в���� = txtInfo(txt��) & "/" & txtInfo(txt��)
        !��Ѫ������ = IIF(optPossession(0).value, 0, 1)
        !�Ƿ�ǩ��ͬ���� = IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, Null))
        !�Ƿ������� = IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, Null))
        !Ԥ����Ѫ���� = txtInfo(txtԤ����Ѫʱ��).Text
        !Ѫ�� = cboInfo(cbo��ѪѪ��).ListIndex
        !RHD = cboInfo(cboRHD).ListIndex
        !��Ѫ��ĿID = mlng��Ѫ��ĿID
        !��Ѫִ�п���ID = IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, 0, cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex))
        !Ԥ����Ѫ�� = Val(txtInfo(txtԤ����Ѫ��).Text)
        !��Ѫ;����ĿID = mlng��Ѫ;��
        !��Ѫ;��ִ�п���ID = IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, 0, cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex))
        !��ע = txtInfo(txt��ע).Text
        !���� = str����
        !��Ѫ�������� = txtInfo(txt��������).Text
        !�������id = mlng��������ID
        !�ٴ�������� = txtInfo(txt�����Ϣ).Text
        !����� = strResult
        !������Ŀ = GetBloodInfo
        !����������ĿSQL = "Zl_��Ѫ�����¼_Insert([���ID]," & chkWait.value & ",'" & cboInfo(cbo��Ѫ����).Text & "','" & cboInfo(cbo��ѪĿ��).Text & "'," & cboInfo(cbo��Ѫ����).ListIndex & "," & IIF(optHistory(0).value, 0, 1) & _
                             "," & IIF(optHistory(2).value, 0, 1) & "," & IIF(optHistory(4).value, 0, 1) & ",'" & txtInfo(txt��) & "/" & txtInfo(txt��) & "'," & IIF(optPossession(0).value, 0, 1) & _
                             "," & cboInfo(cbo��ѪѪ��).ListIndex & "," & cboInfo(cboRHD).ListIndex & "," & IIF(optConsent(0).value, 0, IIF(optConsent(1).value, 1, "Null")) & "," & IIF(optAppraise(0).value, 0, IIF(optAppraise(1).value, 1, "Null")) & ",'" & !������Ŀ & "')"
        !������ĿSQL = str������ĿSQL
        !��Ϲ�����ϢSQL = str��Ϲ�����ϢSQL
        !������ĿSQL = str������ĿSQL
        .Update
    End With
    SaveCacheData = True
    mblnChange = False
End Function

Private Sub LoadDataFromCache()
'���ܣ�ͨ���������ݼ��ؽ���
    Dim str���  As String
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
            cboInfo(cbo��Ѫ����).ListIndex = IIF(1 = Val(!��Ѫ���� & ""), 1, 0)
            If Val(!���� & "") = 1 Then
                txtInfo(txt�����Ϣ).Text = "����"
                chkWait.value = 1
            Else
               '��ȡ���
                mstr���IDs = !�ٴ����IDs & ""
                txtInfo(txt�����Ϣ).Text = !�ٴ�������� & ""
            End If
            txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
            chkWait.value = Val(!���� & "")
            If !��Ѫ���� & "" <> "" Then
                Call zlControl.CboSetText(cboInfo(cbo��Ѫ����), !��Ѫ���� & "", True, "'")
            End If
            If !��ѪĿ�� & "" <> "" Then
                Call zlControl.CboSetText(cboInfo(cbo��ѪĿ��), !��ѪĿ�� & "", True, "'")
'                Call zlControl.CboSetIndex(cboInfo(cbo��Ѫִ��).hWnd, 0)
            End If
            txtInfo(txtԤ����Ѫʱ��).Text = !Ԥ����Ѫ���� & ""
            txtInfo(txtԤ����Ѫʱ��).Tag = txtInfo(txtԤ����Ѫʱ��).Text
            cboInfo(cbo��Ѫ����).ListIndex = Val(!��Ѫ���� & "")
            optHistory(Val(!������Ѫʷ & "")).value = True
            optHistory(IIF(Val(!������Ѫ��Ӧʷ & "") = 1, 3, 2)).value = True
            optHistory(IIF(Val(!��Ѫ���ɼ�����ʷ & "") = 1, 5, 4)).value = True
            If InStr(1, "" & !�в����, "/") <= 0 Then
                txtInfo(txt��).Text = ""
                txtInfo(txt��).Text = ""
            Else
                txtInfo(txt��).Text = Mid(!�в����, 1, InStr(1, "" & !�в����, "/") - 1)
                If Not (txtInfo(txt��).Text = "" Or IsNumeric(txtInfo(txt��).Text)) Then
                    txtInfo(txt��).Text = ""
                End If
                txtInfo(txt��).Text = Mid(!�в����, InStr(1, "" & !�в����, "/") + 1)
                If Not (txtInfo(txt��).Text = "" Or IsNumeric(txtInfo(txt��).Text)) Then
                    txtInfo(txt��).Text = ""
                End If
            End If
            optPossession(Val(!��Ѫ������ & "")).value = True
            If InStr(1, ",0,1,", "," & !�Ƿ�ǩ��ͬ���� & ",") <> 0 Then
                optConsent(Val(!�Ƿ�ǩ��ͬ���� & "")).value = True
            End If
            If InStr(1, ",0,1,", "," & !�Ƿ������� & ",") <> 0 Then
                optAppraise(Val(!�Ƿ������� & "")).value = True
            End If
        
            cboInfo(cbo��ѪѪ��).ListIndex = Val(!Ѫ�� & "")
            cboInfo(cboRHD).ListIndex = Val(!RHD & "")
            
            mstr��Ѫ��Ŀ = !������Ŀ & ""
            mlng��Ѫ��ĿID = Val(!��Ѫ��ĿID & "")
            
            strIDs = ""
            arrItem = Split(mstr��Ѫ��Ŀ, ";")
            For i = 0 To UBound(arrItem)
                strIDs = strIDs & "," & Split(CStr(arrItem(i)), ",")(0)
            Next
            strIDs = Mid(strIDs, 2)
            If InStr(1, "," & strIDs & ",", "," & mlng��Ѫ��ĿID & ",") = 0 Then
                If strIDs <> "" Then
                    strIDs = mlng��Ѫ��ĿID & "," & strIDs
                    mstr��Ѫ��Ŀ = mlng��Ѫ��ĿID & "," & Val(!Ԥ����Ѫ�� & "") & ",,;" & mstr��Ѫ��Ŀ
                Else
                    strIDs = mlng��Ѫ��ĿID
                    mstr��Ѫ��Ŀ = mlng��Ѫ��ĿID & "," & Val(!Ԥ����Ѫ�� & "") & ",,"
                End If
            End If
            lngTmp = Val(!��Ѫִ�п���ID & "")
            Set rsTmp = Get������Ŀ��¼(mlng��Ѫ��ĿID, strIDs)
            strTmp = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & IIF(strTmp = "", "", "'") & rsTmp!����
                rsTmp.MoveNext
            Loop
            txtGet(txtԤ����Ѫ�ɷ�).Text = strTmp
            rsTmp.Filter = "ID=" & mlng��Ѫ��ĿID
            Call Setִ�п���(Val(rsTmp!ִ�п��� & ""), lngTmp)
            txtInfo(txt��λ).Text = rsTmp!���㵥λ & ""
            txtGet(txtԤ����Ѫ�ɷ�).Tag = txtGet(txtԤ����Ѫ�ɷ�).Text
        
            mlng¼������ = Val(rsTmp!¼������ & "")
            mlng��Ѫ;�� = Val(!��Ѫ;����ĿID & "")
            lngTmp = Val(!��Ѫ;��ִ�п���ID & "")
            Set rsTmp = Get������Ŀ��¼(mlng��Ѫ;��)
            If Not (rsTmp!��� = "E" And rsTmp!�������� = "9") Then
                mblnNewSpareBloood = False
                If rsTmp!��� = "E" And rsTmp!�������� = "8" Then
                    mblnSpareBloood = (Val(rsTmp!ִ�з��� & "") = 0)
                End If
            Else
                mblnNewSpareBloood = True
                mblnSpareBloood = True '�������ΪE,��������=9�ľ��Ǳ�Ѫҽ��
            End If
            txtGet(txt��Ѫ;��).Text = rsTmp!���� & ""
            txtGet(txt��Ѫ;��).Tag = txtGet(txt��Ѫ;��).Text
            Call Set��Ѫִ��(Val(rsTmp!ִ�п��� & ""), lngTmp)
            
            txtInfo(txtԤ����Ѫ��).Text = zl9ComLib.FormatEx((!Ԥ����Ѫ�� & ""), 5)
            txtInfo(txt��ע).Text = !��ע & ""
            txtInfo(txt��������).Text = !��Ѫ�������� & ""
            
            Call SetLisResult(strIDs)
            '��Ѫҽ������
            strTmp = !���� & ""
            cboInfo(cbo����).Text = ""
            lblInfo(31).Visible = True
            If strTmp Like "*��/����" Then
                If IsNumeric(Split(strTmp, "��/����")(0)) = True Then
                    cboInfo(cbo����).Text = Split(strTmp, "��/����")(0)
                End If
            ElseIf strTmp = "��ѹ" Or strTmp = "����" Then
                cboInfo(cbo����).Text = strTmp
                lblInfo(31).Visible = False
            End If
            
            strResult = !����� & ""
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
'���ܣ����������������ô����������
    '�����ķ���������
    Dim bln��Ѫ���� As Boolean '������Ѫ�⣬������Ѫ����
    Dim arrItem, arrCode, i As Integer
    mblnSelectBlood = False
    
    bln��Ѫ���� = mblnSpareBloood = False
    If bln��Ѫ���� Then
        If mintType = 0 Then '���������������ݲ��������Ƿ���ҽ��ѡ��Ʒ��
            mblnSelectBlood = gbln�´���Ѫ����ȷ��ѪҺ��Ϣ
        Else  '������ģʽ���������Ѫ��Ŀ���ݾ���������ģʽ
            arrItem = Split(mstr��Ѫ��Ŀ, ";")
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
        lblHead.Caption = "�ٴ���Ѫ���뵥"
    Else
        lblHead.Caption = "�ٴ�ȡѪ֪ͨ��"
    End If
    If mblnSpareBloood = True And mblnNewSpareBloood = True Then
        lblInfo(lbl��Ѫִ��).Caption = "��Ѫִ��"
        lblInfo(lbl��Ѫ;��).Caption = "�ɼ�����"
        lblInfo(lbl��Ѫִ��).Caption = "�ɼ�ִ��"
    Else
        lblInfo(lbl��Ѫִ��).Caption = "��Ѫִ��"
        lblInfo(lbl��Ѫ;��).Caption = "��Ѫ;��"
        lblInfo(lbl��Ѫִ��).Caption = "��Ѫִ��"
    End If
    '�·�ǩ���ؼ��������
    If mint���� = 0 And mblnSpareBloood Then
        lblInfo(lbl����ҽʦǩ��).Top = lblInfo(lbl����ҽʦ����).Top
        Line1(lin����ҽʦǩ��).Y1 = Line1(lin����ҽʦ����).Y1
        Line1(lin����ҽʦǩ��).Y2 = Line1(lin����ҽʦǩ��).Y1
        txtInfo(txt����ҽʦǩ��).Top = txtInfo(txt����ҽʦ����).Top
        lblInfo(lbl����ҽʦǩ��).Visible = True
        Line1(lin����ҽʦǩ��).Visible = True
        txtInfo(txt����ҽʦǩ��).Visible = True
    Else
        lblInfo(lbl����ҽʦǩ��).Top = lblInfo(lbl����ҽʦǩ��).Top
        Line1(lin����ҽʦǩ��).Y1 = Line1(lin����ҽʦǩ��).Y1
        Line1(lin����ҽʦǩ��).Y2 = Line1(lin����ҽʦǩ��).Y1
        txtInfo(txt����ҽʦǩ��).Top = txtInfo(txt����ҽʦǩ��).Top
        lblInfo(lbl����ҽʦǩ��).Visible = False
        Line1(lin����ҽʦǩ��).Visible = False
        txtInfo(txt����ҽʦǩ��).Visible = False
    End If
    lblInfo(lbl�ɼ���ǩ��) = IIF(mblnSpareBloood = True, "�� �� ��ǩ��", "ȡ Ѫ ��ǩ��")
    
    '�ؼ�λ�õ���(��Ѫ�����Ŀ)
    On Error Resume Next
    lblInfo(23).Visible = bln��Ѫ���� = False '��Ѫ�߱�ǩ
    lblInfo(lbl������Ѫʷ).Visible = bln��Ѫ���� = False
    fraChk(fra������Ѫʷ).Visible = bln��Ѫ���� = False
    lblInfo(lbl������Ѫ��Ӧʷ).Visible = bln��Ѫ���� = False
    fraChk(fra������Ѫ��Ӧʷ).Visible = bln��Ѫ���� = False
    lblInfo(lbl��Ѫ���ɼ�����ʷ).Visible = bln��Ѫ���� = False
    fraChk(fra��Ѫ���ɼ�����ʷ).Visible = bln��Ѫ���� = False
    lblInfo(lbl�в����).Visible = bln��Ѫ���� = False
    fraChk(fra�в����).Visible = bln��Ѫ���� = False
    lblInfo(lbl��Ѫ������).Visible = bln��Ѫ���� = False
    fraChk(fra��Ѫ������).Visible = bln��Ѫ���� = False
    lblInfo(lbl֪��ͬ����).Visible = bln��Ѫ���� = False
    fraChk(fra֪��ͬ����).Visible = bln��Ѫ���� = False
    lblInfo(lbl��Ѫ����).Visible = bln��Ѫ���� = False
    fraChk(fra��Ѫ����).Visible = bln��Ѫ���� = False
    
    
    'Ԥ����Ѫ����
    lblInfo(lblԤ����Ѫ����).Top = IIF(bln��Ѫ���� = True, lblInfo(lbl�в����).Top, lblInfo(lbl֪��ͬ����).Top + lblInfo(lbl֪��ͬ����).Height + 210)
    txtInfo(txtԤ����Ѫʱ��).Top = lblInfo(lblԤ����Ѫ����).Top - 30
    Line1(12).Y1 = txtInfo(txtԤ����Ѫʱ��).Top + txtInfo(txtԤ����Ѫʱ��).Height + 15
    Line1(12).Y2 = Line1(12).Y1
    cmdDate(cmdԤ����Ѫʱ��).Top = txtInfo(txtԤ����Ѫʱ��).Top
    'Ѫ��
    lblInfo(lblѪ��).Top = lblInfo(lblԤ����Ѫ����).Top
    picInfo(1).Top = txtInfo(txtԤ����Ѫʱ��).Top - 30
    Line1(13).Y1 = Line1(12).Y1
    Line1(13).Y2 = Line1(13).Y1
    'RH
    lblInfo(lblRHD).Top = lblInfo(lblԤ����Ѫ����).Top
    picInfo(2).Top = picInfo(1).Top
    Line1(14).Y1 = Line1(13).Y1
    Line1(14).Y2 = Line1(14).Y1
    
    With picPreBlood
        .Left = lblInfo(lblԤ����Ѫ����).Left
        .Top = lblInfo(lblԤ����Ѫ����).Top + lblInfo(lblԤ����Ѫ����).Height + 180
        .Visible = True
    End With
    
    With picBloodDept
        .Top = picPreBlood.Top - 30
        .Left = picPreBlood.Width + picPreBlood.Left - .Width
        .Visible = True
        .ZOrder 0
    End With
    '������Ѫ�������ʾ(Ŀǰ��ʱ����)
    lblInfo(30).Visible = bln��Ѫ����
    picInfo(3).Visible = bln��Ѫ����
    Line1(25).Visible = bln��Ѫ����
    lblInfo(31).Visible = bln��Ѫ����
    If bln��Ѫ���� Then
        If cboInfo(cbo����).Text <> "" And IsNumeric(cboInfo(cbo����).Text) = False Then lblInfo(31).Visible = False
    End If
    '��Ѫִ��λ�ñ䶯����
    If bln��Ѫ���� = False Then
        picInfo(8).Left = picBloodDept.Width - picInfo(8).Width - 30
    Else
        picInfo(8).Left = lblInfo(30).Left - picInfo(8).Width - 120
    End If
    Line1(16).X1 = picInfo(8).Left - 75
    Line1(16).X2 = picInfo(8).Left + picInfo(8).Width + 15
    lblInfo(lbl��Ѫִ��).Left = picInfo(8).Left - lblInfo(lbl��Ѫִ��).Width - 120
    
    '��Ѫ�ɷֺ���Ѫ������
    vsLIS.Visible = bln��Ѫ���� = False
    lblInfo(lblԤ����Ѫ�ɷ�).Visible = False
    txtGet(txtԤ����Ѫ�ɷ�).Visible = False
    txtGet(txtԤ����Ѫ�ɷ�).Locked = True
    txtGet(txtԤ����Ѫ�ɷ�).BackColor = &H8000000F
    picGet(0).Visible = False
    lblInfo(lblԤ����Ѫ��).Visible = False
    txtInfo(txtԤ����Ѫ��).Visible = False
    txtInfo(txt��λ).Visible = False
    Line1(15).Visible = False
    Line1(17).Visible = False
    If bln��Ѫ���� = False Then
        'Ԥ����Ѫ�ɷ�
        lblInfo(lblԤ����Ѫ�ɷ�).Top = lblInfo(lbl������).Top - 765
    Else
        'Ԥ����Ѫ�ɷ�
        lblInfo(lblԤ����Ѫ�ɷ�).Top = lblInfo(lbl��ע).Top - 1000
    End If
    picGet(0).Top = lblInfo(lblԤ����Ѫ�ɷ�).Top - 30
    Line1(15).Y1 = picGet(0).Top + picGet(0).Height - 5
    Line1(15).Y2 = Line1(15).Y1
    picPreBlood.Height = Line1(15).Y1 - picPreBlood.Top
        
    '��Ѫ�б�
    picPreInfo.Left = 15
    picPreInfo.Top = picPreBlood.Height - picPreInfo.Height
    picPreInfo.Width = picPreBlood.Width - 15
    picPreSum.Left = picPreInfo.Width - picPreSum.Width - 30
    txt������Ϣ.Width = picPreSum.Left - 240
    
    'ѪҺ�б�
    vsfBlood.Left = 15
    vsfBlood.Top = 330
    vsfBlood.Width = IIF(bln��Ѫ���� = False, picPreBlood.Width - IIF(gbln��ʾѪҺ��� = True, 3000, 45), 5000)
    vsfBlood.Height = picPreInfo.Top - vsfBlood.Top - 30
    
    '�����Ϣ
    vsfList.Left = vsfBlood.Left + vsfBlood.Width + 15
    vsfList.Top = vsfBlood.Top
    vsfList.Width = picPreBlood.Width - vsfList.Left - 45
    vsfList.Height = vsfBlood.Height
    vsfList.Visible = IIF(bln��Ѫ���� = False, gbln��ʾѪҺ���, True)
    
    'Ԥ����Ѫ��
    lblInfo(lblԤ����Ѫ��).Top = lblInfo(lblԤ����Ѫ�ɷ�).Top
    txtInfo(txtԤ����Ѫ��).Top = picGet(0).Top
    Line1(17).Y1 = txtInfo(txtԤ����Ѫ��).Top + txtInfo(txtԤ����Ѫ��).Height + 15
    Line1(17).Y2 = Line1(17).Y1
    '��λ
    txtInfo(txt��λ).Top = txtInfo(txtԤ����Ѫ��).Top
    txtInfo(txt��λ).Left = txtInfo(txtԤ����Ѫ��).Left + txtInfo(txtԤ����Ѫ��).Width + 120

    '��Ѫ;��
    lblInfo(lbl��Ѫ;��).Top = lblInfo(lblԤ����Ѫ��).Top + lblInfo(lblԤ����Ѫ��).Height + 225
    picGet(1).Top = lblInfo(lbl��Ѫ;��).Top - 30
    Line1(19).Y1 = picGet(1).Top + picGet(1).Height - 5
    Line1(19).Y2 = Line1(19).Y1
    '��Ѫִ��
    lblInfo(lbl��Ѫִ��).Top = lblInfo(lbl��Ѫ;��).Top
    picInfo(9).Top = picGet(1).Top
    Line1(20).Y1 = Line1(19).Y1
    Line1(20).Y2 = Line1(20).Y1
    
    '����24Сʱ��Ѫ��
    lblInfo(lbl24H��Ѫ��).Visible = Not bln��Ѫ����
    If Not bln��Ѫ���� Then lblInfo(lbl24H��Ѫ��).Caption = "24Сʱ����Ѫ��������" & GetBloodCapacity(IIF(mint���� = 0, 2, 1), mlng����ID, IIF(mint���� = 0, mlng��ҳID, mlng�Һ�ID), zlDatabase.Currentdate, True, CInt(mbytBaby)) & "ML"
    
    If bln��Ѫ���� Then
        picHisItem.Visible = False
    Else
        picHisItem.Visible = True
        lblInfo(lbl������ʷ������Ŀ).Tag = GetPatiHisBloodItem
        If lblInfo(lbl������ʷ������Ŀ).Tag <> "" Then
            lblInfo(lbl������ʷ������Ŀ).Caption = "������ʷ������Ŀ:[" & Replace(lblInfo(lbl������ʷ������Ŀ).Tag, "'", "][") & "]"
            lblInfo(lbl������ʷ������Ŀ).Tag = Replace(lblInfo(lbl������ʷ������Ŀ).Tag, "'", vbCrLf)
        Else
            lblInfo(lbl������ʷ������Ŀ).Caption = ""
        End If
    End If
    
    Call cboInfo_Click(cbo��Ѫ����)
    On Error GoTo 0
    If mblnSelectBlood = True Then
        Call LoadBloodListBySelect
    Else
        Call LoadBloodList(bln��Ѫ����)
    End If
    SetFormNature = True
End Function

Private Sub LoadBloodListBySelect()
'��Ѫ���������ͨ��ҽ��ѡ��ѪҺ��ģʽ����ôκ���
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim strWhere As String, str��λ As String, strTmp As String
    Dim strSQLChild1 As String, strSQLChild2 As String
    Dim arrItem, arrCode() As String  '�����������ڿ�ʼ�����Ѫ������Ŀ���������ʹ��ע��Ӱ��
    Dim strID As String
    On Error GoTo ErrHand
    
    txt������Ϣ.Text = "Ʒ��:"
    txt������.Text = ""
    
    arrItem = Split(mstr��Ѫ��Ŀ, ";")
    '��ȡѪҺ�շ�ID��ѪҺ����ѷ����鿴ʱ��Ҫ��ȡԭʼ��¼
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
        .TextMatrix(0, COL_S_ѡ��) = "": .ColWidth(COL_S_ѡ��) = 255
        .TextMatrix(0, COL_S_���) = "Ѫ�����": .ColWidth(COL_S_���) = 1200:
        .TextMatrix(0, COL_S_���) = "���": .ColWidth(COL_S_���) = 1000
        .TextMatrix(0, COL_S_Ч��) = "Ч��": .ColWidth(COL_S_Ч��) = 1400
        .ColHidden(0) = True
        .ColDataType(COL_S_ѡ��) = flexDTString
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_S_���) = flexAlignCenterCenter
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
        .TextMatrix(0, COL_P_ѡ��) = "": .ColWidth(COL_P_ѡ��) = 255
        .TextMatrix(0, COL_P_����) = "����": .ColWidth(COL_P_����) = 1000:
        .TextMatrix(0, COL_P_����) = "����": .ColWidth(COL_P_����) = 3000
        .TextMatrix(0, COL_P_������) = "������": .ColWidth(COL_P_������) = 800
        .TextMatrix(0, COL_P_��λ) = "��λ": .ColWidth(COL_P_��λ) = 600
        .TextMatrix(0, COL_P_����Ѫ��) = "����Ѫ��": .ColWidth(COL_P_����Ѫ��) = 1000
        .TextMatrix(0, COL_P_����RH) = "����RH": .ColWidth(COL_P_����RH) = 800
        .TextMatrix(0, COL_P_ִ�з���ID) = "ִ�з���ID": .ColWidth(COL_P_ִ�з���ID) = 0
        .TextMatrix(0, COL_P_ִ�п���ID) = "ִ�п���ID": .ColWidth(COL_P_ִ�п���ID) = 0
        .TextMatrix(0, COL_P_¼������ID) = "¼������ID": .ColWidth(COL_P_¼������ID) = 0
        .TextMatrix(0, COL_P_����ϵ��) = "����ϵ��": .ColWidth(COL_P_����ϵ��) = 0
        .TextMatrix(0, COL_P_���) = "���": .ColWidth(COL_P_���) = 0
        
        .ColHidden(COL_P_ID) = True
        .ColHidden(COL_P_ִ�з���ID) = True
        .ColHidden(COL_P_ִ�п���ID) = True
        .ColHidden(COL_P_¼������ID) = True
        .ColHidden(COL_P_����ϵ��) = True
        .ColHidden(COL_P_���) = True
        .ColHidden(COL_P_������) = True
        .ColHidden(COL_P_����Ѫ��) = True
        .ColHidden(COL_P_����RH) = True
        .ColDataType(COL_P_ѡ��) = flexDTString
        
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_P_������) = flexAlignCenterCenter
        .ColAlignment(COL_P_����Ѫ��) = flexAlignCenterCenter
        .ColAlignment(COL_P_����RH) = flexAlignCenterCenter
        .Editable = flexEDNone
    End With
    
    If mint���� = 0 Then
        strWhere = " And b.����id = [1] And b.��ҳid = [2] And Nvl(b.Ӥ��, 0) = [3] "
    Else
        strWhere = " And b.����id = [1] And b.�Һŵ� = [2] And Nvl(b.Ӥ��, 0) = [3]"
    End If
    
    gstrSQL = _
        " Select a.Id, a.����, a.����, a.���㵥λ, a.ִ�з��� ִ�з���id, a.ִ�п��� ִ�п���id, a.¼������ ¼������id, a.����ϵ��, b.ѪҺ��Ϣ" & vbNewLine & _
        " From ������ĿĿ¼ a," & vbNewLine & _
        "     ("
    strSQLChild1 = "Select h.Id," & vbNewLine & _
        "              f_List2str(Cast(Collect(f.Id || '<Split1>' || f.Ѫ����� || '<Split1>' || decode(substr('' || Nvl(f.��д����, 0) * Nvl(����ϵ��, 1),1,1),'.',0,'') || Nvl(f.��д����, 0) * Nvl(����ϵ��, 1) || h.���㵥λ || '<Split1>' ||" & vbNewLine & _
        "                                       To_Char(f.Ч��, 'yyyy-mm-dd hh24:mi')) As t_Strlist)," & vbNewLine & _
        "                          '<Split4>') ѪҺ��Ϣ" & vbNewLine & _
        "       From ������ĿĿ¼ h, ѪҺ��� g, ѪҺ�շ���¼ f, ѪҺ��Ѫ��¼ e, ����ҽ����¼ b" & vbNewLine & _
        "       Where h.Id = g.Ʒ��id And g.���id = f.ѪҺid  And f.����� is Null  And f.�䷢id = e.Id And Mod(f.��¼״̬, 3) = 1 And f.��Ѫ״̬ = 1 And " & vbNewLine & _
        "             e.����id = b.Id And b.������� = 'K' And b.ҽ��״̬ In (1, 3, 8) " & strWhere & " And" & vbNewLine & _
        "             Exists" & vbNewLine & _
        "        (Select 1" & vbNewLine & _
        "              From ������ĿĿ¼ p, ����ҽ����¼ q" & vbNewLine & _
        "              Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = b.Id And q.������� = 'E')" & vbNewLine & _
        "       And  Not Exists (Select 1" & vbNewLine & _
        "              From ��Ѫ������Ŀ a, ����ҽ����¼ b" & vbNewLine & _
        "              Where a.ҽ��id = b.Id And Instr('|' || a.ѪҺ��Ϣ || '|' ,'|' || f.Id || '|') <> 0  And b.ҽ��״̬ In (1, 3, 8) And b.������� = 'K'  " & strWhere & ")" & vbNewLine & _
        "       Group By h.Id, h.����, h.���㵥λ"
        
    strSQLChild2 = "Select h.Id," & vbNewLine & _
        "              f_List2str(Cast(Collect(f.Id || '<Split1>' || f.Ѫ����� || '<Split1>' || decode(substr('' || Nvl(f.��д����, 0) * Nvl(����ϵ��, 1),1,1),'.',0,'') || Nvl(f.��д����, 0) * Nvl(����ϵ��, 1) || h.���㵥λ || '<Split1>' ||" & vbNewLine & _
        "                                       To_Char(f.Ч��, 'yyyy-mm-dd hh24:mi')) As t_Strlist)," & vbNewLine & _
        "                          '<Split4>') ѪҺ��Ϣ" & vbNewLine & _
        "       From ������ĿĿ¼ h, ѪҺ��� g, ѪҺ�շ���¼ f, ѪҺ��Ѫ��¼ e, ����ҽ����¼ b" & vbNewLine & _
        "       Where h.Id = g.Ʒ��id And g.���id = f.ѪҺid  And instr([4],'|' || f.id || '|',1)<>0  And f.�䷢id = e.Id And Mod(f.��¼״̬, 3) = 1 And" & vbNewLine & _
        "             e.����id = b.Id And b.������� = 'K' And b.ҽ��״̬ In (1, 3, 8) " & strWhere & " And" & vbNewLine & _
        "             Exists" & vbNewLine & _
        "        (Select 1" & vbNewLine & _
        "              From ������ĿĿ¼ p, ����ҽ����¼ q" & vbNewLine & _
        "              Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = b.Id And q.������� = 'E')" & vbNewLine & _
        "       Group By h.Id, h.����, h.���㵥λ"
    If mintType = 0 Then '����
        gstrSQL = gstrSQL & strSQLChild1
    ElseIf mintType = 2 Then '�鿴
        gstrSQL = gstrSQL & strSQLChild2
    Else '�޸�
        gstrSQL = gstrSQL & "Select id, f_List2str(Cast(Collect(ѪҺ��Ϣ) As t_Strlist),'<Split4>') ѪҺ��Ϣ From (" & strSQLChild2 & vbNewLine & " Union ALL" & vbNewLine & strSQLChild1 & ") Group By Id"
    End If
    
    gstrSQL = gstrSQL & ") b" & vbNewLine & _
            " Where a.Id = b.Id" & vbNewLine & _
            " Order By a.����"
            
    If mint���� = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mlng��ҳID, mbytBaby, "|" & strID & "|")
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mstr�Һŵ�, mbytBaby, "|" & strID & "|")
    End If
    
    If rsTmp.RecordCount > 0 Then
        If mintType <> 0 And mintType <> 1 And mintType <> 4 Then 'ֻ���������޸�ʱ������ѡ��
            vsfBlood.Editable = flexEDNone
        Else
            vsfBlood.Editable = IIF(mblnUseBloodSend = True, flexEDNone, flexEDKbdMouse)
        End If
    Else
        vsfBlood.Editable = flexEDNone
    End If
    
    vsfList.Editable = vsfBlood.Editable  'ѪҺ�б�༭����Ѫҽ����ͬ
    
    With vsfBlood
        .Redraw = flexRDNone
        Do While Not rsTmp.EOF
            If .Rows <= .FixedRows Then
                .Rows = .FixedRows + 1
            Else
                If .TextMatrix(.Rows - 1, COL_P_ID) <> "" Then .Rows = .Rows + 1
            End If
            
            .TextMatrix(.Rows - 1, COL_P_ID) = Val(rsTmp!ID & "")
            .TextMatrix(.Rows - 1, COL_P_ѡ��) = ""
            .TextMatrix(.Rows - 1, COL_P_����) = rsTmp!���� & ""
            .TextMatrix(.Rows - 1, COL_P_����) = rsTmp!���� & ""
            .TextMatrix(.Rows - 1, COL_P_������) = ""
            .TextMatrix(.Rows - 1, COL_P_��λ) = rsTmp!���㵥λ & ""
            If InStr(1, "'" & UCase(str��λ) & "'", "'" & UCase(rsTmp!���㵥λ & "") & "'") = 0 Then
                str��λ = IIF(str��λ = "", "", str��λ & "'") & rsTmp!���㵥λ & ""
            End If
            .TextMatrix(.Rows - 1, COL_P_����Ѫ��) = ""
            .TextMatrix(.Rows - 1, COL_P_����RH) = ""
            .TextMatrix(.Rows - 1, COL_P_ִ�з���ID) = Val(rsTmp!ִ�з���ID & "")
            .TextMatrix(.Rows - 1, COL_P_ִ�п���ID) = Val(rsTmp!ִ�п���ID & "")
            .TextMatrix(.Rows - 1, COL_P_¼������ID) = Val(rsTmp!¼������ID & "")
            .TextMatrix(.Rows - 1, COL_P_����ϵ��) = Val(rsTmp!����ϵ�� & "")
            .TextMatrix(.Rows - 1, COL_P_���) = rsTmp!ѪҺ��Ϣ & ""
            
            Set .Cell(flexcpPicture, .Rows - 1, COL_P_ѡ��) = img16.ListImages("c0").Picture
            .Cell(flexcpData, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_ѡ��) = 0
            .Cell(flexcpFontBold, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = False
            .Cell(flexcpBackColor, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = vbWhite
            For j = 0 To UBound(arrItem)
                arrCode = Split(CStr(arrItem(j)), ",")
                If Val(arrCode(0)) = Val(rsTmp!ID & "") Then
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_ѡ��) = img16.ListImages("c1").Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_ѡ��) = 1
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = True
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = &HC0E0FF
                    .TextMatrix(.Rows - 1, COL_P_������) = arrCode(1)
                    .TextMatrix(.Rows - 1, COL_P_����Ѫ��) = arrCode(2)
                    .TextMatrix(.Rows - 1, COL_P_����RH) = arrCode(3)
                    If UBound(arrCode) > 3 Then
                        .TextMatrix(.Rows - 1, COL_P_���) = .TextMatrix(.Rows - 1, COL_P_���) & IIF(arrCode(4) <> "", "<Split3>" & arrCode(4), "")
                    End If
                    'iif(mid("" & 0.5,1,1)=".","0","") & 0.5������д����Ϊ�˱�֤С�ڵ�1��ֵ��������ʾǰ׺0
                    txt������Ϣ.Text = txt������Ϣ.Text & "[" & .TextMatrix(.Rows - 1, COL_P_����) & IIF(.TextMatrix(.Rows - 1, COL_P_������) <> "", "-" & IIF(Mid("" & .TextMatrix(.Rows - 1, COL_P_������), 1, 1) = ".", "0", "") & .TextMatrix(.Rows - 1, COL_P_������) & .TextMatrix(.Rows - 1, COL_P_��λ), "") & "]"
                End If
            Next
            rsTmp.MoveNext
        Loop
        If .Rows > .FixedRows Then
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
            'ȷ�����ߴ�
            .AutoSize 0, .Cols - 1
            .ColWidth(COL_P_ѡ��) = 255
            .Redraw = flexRDDirect
            Call vsfBlood_AfterRowColChange(0, 0, 1, 1)
        Else
            .Redraw = flexRDDirect
        End If
    End With
    
    '���ص�λ
    arrItem = Split(str��λ, "'")
    cboInfo(cbo��λ).Clear
    cboInfo(cbo��λ).Tag = ""
    For i = 0 To UBound(arrItem)
        cboInfo(cbo��λ).AddItem CStr(arrItem(i))
        If UCase(txtInfo(txt��λ).Text) = UCase(CStr(arrItem(i))) Then
            Call zlControl.CboSetIndex(cboInfo(cbo��λ).hwnd, i)
            cboInfo(cbo��λ).Tag = i
        End If
    Next
    Call BloodSum
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBloodList(ByVal bln��Ѫ���� As Boolean)
'���ܣ���Ѫ�������ѪҺ��Ϣ
    '��Ѫ������ر���
    Dim strѪ�� As String, strRH As String
    '��Ѫ������ر���
    Dim strWhere As String

    '��������
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lng������ĿID As Long
    Dim arrItem, arrCode() As String
    Dim str��λ As String
    Dim arrRecord, arrInfo, str��Ѫ��Ϣ As String
    Dim strδ�� As String, str�ѷ� As String, str���� As String, str���� As String
    Dim str�䷢��Ϣ As String, str�����Ŀ As String, str���㵥λ As String, blnLast As Boolean
    Dim objCollection As New Collection
    Dim bln��ʾ������ As Boolean
    
    On Error GoTo ErrHand
    
    If bln��Ѫ���� = False Then
        For i = 0 To cboInfo(cbo��ѪѪ��).ListCount - 1
            If InStr(1, ",A,B,O,AB,", "," & cboInfo(cbo��ѪѪ��).List(i) & ",") <> 0 Then
                strѪ�� = strѪ�� & "|" & cboInfo(cbo��ѪѪ��).List(i)
            End If
        Next i
        strѪ�� = Mid(strѪ��, 2)
        For i = 0 To cboInfo(cboRHD).ListCount - 1
            strRH = strRH & "|" & cboInfo(cboRHD).List(i)
        Next i
        strRH = Mid(strRH, 2)
    End If
    txt������Ϣ.Text = "Ʒ��:"
    txt������.Text = ""
    
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
        .TextMatrix(0, 0) = IIF(bln��Ѫ���� = True, "�䷢��Ϣ", "�����Ϣ")
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
        .TextMatrix(0, COL_P_ѡ��) = "": .ColWidth(COL_P_ѡ��) = 255
        .TextMatrix(0, COL_P_����) = "����": .ColWidth(COL_P_����) = 1000:
        .TextMatrix(0, COL_P_����) = "����": .ColWidth(COL_P_����) = 2000
        .TextMatrix(0, COL_P_������) = "������": .ColWidth(COL_P_������) = 800
        .TextMatrix(0, COL_P_��λ) = "��λ": .ColWidth(COL_P_��λ) = 600
        .TextMatrix(0, COL_P_����Ѫ��) = "����Ѫ��": .ColWidth(COL_P_����Ѫ��) = 1000
        .TextMatrix(0, COL_P_����RH) = "����RH": .ColWidth(COL_P_����RH) = 800
        .TextMatrix(0, COL_P_ִ�з���ID) = "ִ�з���ID": .ColWidth(COL_P_ִ�з���ID) = 0
        .TextMatrix(0, COL_P_ִ�п���ID) = "ִ�п���ID": .ColWidth(COL_P_ִ�п���ID) = 0
        .TextMatrix(0, COL_P_¼������ID) = "¼������ID": .ColWidth(COL_P_¼������ID) = 0
        .TextMatrix(0, COL_P_����ϵ��) = "����ϵ��": .ColWidth(COL_P_����ϵ��) = 0
        .TextMatrix(0, COL_P_���) = "���": .ColWidth(COL_P_���) = 0
        
        .ColHidden(COL_P_ID) = True
        .ColHidden(COL_P_ִ�з���ID) = True
        .ColHidden(COL_P_ִ�п���ID) = True
        .ColHidden(COL_P_¼������ID) = True
        .ColHidden(COL_P_����ϵ��) = True
        .ColHidden(COL_P_���) = True
        .ColHidden(COL_P_����Ѫ��) = Not (bln��Ѫ���� = False And bln��ʾ������ = True)
        .ColHidden(COL_P_����RH) = Not (bln��Ѫ���� = False And bln��ʾ������ = True)
        .ColDataType(COL_P_ѡ��) = flexDTString
        If bln��Ѫ���� = False Then
            .ColComboList(COL_P_����Ѫ��) = strѪ��
            .ColComboList(COL_P_����RH) = strRH
            If gbln��ʾѪҺ��� = False Then
                .ColWidth(COL_P_����) = 4000
            ElseIf bln��ʾ������ = False Then
                .ColWidth(COL_P_����) = 4000
            End If
        End If
        
        .RowHeight(0) = 300
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        .ColAlignment(COL_P_������) = flexAlignCenterCenter
        .ColAlignment(COL_P_����Ѫ��) = flexAlignCenterCenter
        .ColAlignment(COL_P_����RH) = flexAlignCenterCenter
    End With
    

    If bln��Ѫ���� = False Then
        If gbln��ʾѪҺ��� = True Then
            gstrSQL = _
                    " Select Id, ����,����, Sum(����) ����, f_List2str(Cast(Collect(��������) As t_Strlist), '<Split1>') �����Ϣ, ���㵥λ,ִ�п���ID,¼������ID,ִ�з���id,����ϵ��" & vbNewLine & _
                    " From (Select Id, ����,����, ���㵥λ,ִ�п���ID,¼������ID,����ϵ��,ִ�з���id, Sum(����) ����," & vbNewLine & _
                    "              Decode(�ⷿ����, '', '', '��' || �ⷿ���� || '��- ' || Sum(����) || ���㵥λ || '<Split2>' || f_List2str(Cast(Collect(�������� || ���㵥λ) As t_Strlist),'<Split3>')) �������� " & vbNewLine & _
                    "       From (Select a.Id,A.����, a.����, e.�ⷿid, Nvl(Max(f.����), '') �ⷿ����," & vbNewLine & _
                    "                     e.Abo || e.Rh || ':' || Nvl(Sum(e.�������� * d.����ϵ��), 0) ��������, Nvl(Sum(e.�������� * d.����ϵ��), 0) ����, a.���㵥λ,a.����ϵ��,A.ִ�п��� as ִ�п���ID,A.¼������ as ¼������ID,a.ִ�з��� as ִ�з���id" & vbNewLine & _
                    "              From ���ű� f, ѪҺ����¼ e, ѪҺ��� d, ���Ʒ���Ŀ¼ c, ������ĿĿ¼ a,������Ŀ���� B" & vbNewLine & _
                    "              Where e.�ⷿid = f.Id(+) And e.ѪҺid(+) = d.���id And  e.Ч��(+)>Sysdate And d.Ʒ��id = a.Id And c.Id = a.����id And c.���� = 8 And A.ID=B.������ĿID" & vbNewLine & _
                    "                   And A.���='K'  And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & vbNewLine & _
                    "                   And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
                    "                           And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[1])  Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & vbNewLine & _
                    "                   " & Decode(gbytCode, 0, " And B.���� IN([2],3)", 1, " And B.���� IN([2],3)", "") & vbNewLine & _
                    "              Group By a.Id, A.����,a.����, a.���㵥λ,A.ִ�п���,A.¼������, e.�ⷿid, e.Abo, e.Rh,ִ�з���,����ϵ��)" & vbNewLine & _
                    "       Group By Id, ����,����, ���㵥λ,ִ�п���ID,¼������ID, �ⷿ����,ִ�з���id,����ϵ��)" & vbNewLine & _
                    " Group By Id, ����,����, ���㵥λ,ִ�п���ID,¼������ID,ִ�з���id,����ϵ��" & vbNewLine & _
                    " Order by ����"
        Else
            gstrSQL = "Select Distinct a.Id, a.����, a.����, a.ִ�з��� As ִ�з���id, a.���㵥λ, a.ִ�п��� As ִ�п���id, a.¼������ As ¼������id, a.����ϵ��,' ' as �����Ϣ,' ' as ����" & vbNewLine & _
                " From ���Ʒ���Ŀ¼ c, ������ĿĿ¼ a, ������Ŀ���� b, ѪҺƷ�� d" & vbNewLine & _
                " Where c.Id = a.����id And c.���� = 8 And a.Id = b.������Ŀid And a.Id = d.Ʒ��id And a.��� = 'K' And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And" & vbNewLine & _
                "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null) And" & vbNewLine & _
                "      (Exists (Select 1 From �������ÿ��� Where ��Ŀid = a.Id And ����id = [1]) Or Not Exists" & vbNewLine & _
                "       (Select 1 From �������ÿ��� Where ��Ŀid = a.Id)) " & Decode(gbytCode, 0, " And B.���� IN([2],3)", 1, " And B.���� IN([2],3)", "") & vbNewLine & _
                " Order By a.����"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡѪҺ��Ϣ", mlng���˿���id, gbytCode + 1)
         If rsTmp.RecordCount > 0 Then
            If mintType <> 0 And mintType <> 1 Then  'ֻ���������޸�ʱ������ѡ��
                vsfBlood.Editable = flexEDNone
            Else
                vsfBlood.Editable = flexEDKbdMouse
            End If
        Else
            vsfBlood.Editable = flexEDNone
        End If
        arrRecord = Array()
        Do While Not rsTmp.EOF
            'һ��Ʒ��һ����Ϣ
            ReDim Preserve arrRecord(UBound(arrRecord) + 1)
            arrRecord(UBound(arrRecord)) = rsTmp!ID & "'" & rsTmp!���� & "'" & rsTmp!���� & "'" & rsTmp!���㵥λ & "'" & rsTmp!ִ�з���ID & "'" & rsTmp!ִ�п���ID & "'" & rsTmp!¼������ID & "'" & rsTmp!����ϵ��
            '����ʽSQL�Ѿ�����á���ʽ����ѪҺ�ҡ�- 800ml<Split2>A��+:400ml<Split3>B��+:400ml<Split1>��LWҽ������A��- 800ml<Split2>A��+:400ml<Split3>B��+:400ml
            If gbln��ʾѪҺ��� = True Then
                str��Ѫ��Ϣ = "" & rsTmp!�����Ϣ
            Else
                str��Ѫ��Ϣ = ""
            End If
            
            objCollection.Add str��Ѫ��Ϣ, "A_" & rsTmp!ID
            rsTmp.MoveNext
        Loop
    Else
        If mint���� = 0 Then
            strWhere = " And b.����id = [1] And b.��ҳid = [2] And Nvl(b.Ӥ��, 0) = [3] "
        Else
            strWhere = " And b.����id = [1] And b.�Һŵ� = [2] And Nvl(b.Ӥ��, 0) = [3]"
        End If
        '123316,����ѪҺ��Ϣ����Ϊ������ȡ��ƴ�ӵķ�ʽ����ǰ��sql��ֱ�Ӵ���õģ����������û�����ѪҺ̫��ƴ���ַ�����4000����
        gstrSQL = "Select a.Id, a.����, a.����, a.���㵥λ, a.ִ�з���id, a.ִ�п���id, a.¼������id, a.����ϵ��, a.����, a.�ѷ�, a.δ��, a.ѪҺ��Ϣ, a.�Ƿ���Ѫ, a.����, a.������λ," & vbNewLine & _
                        "       a.�����Ŀ" & vbNewLine & _
                        "From (With ��Ѫ��¼ As (Select h.Id," & vbNewLine & _
                        "                           Decode(Nvl(f.�����, ''), '', Nvl(f.��д����, 0) * Nvl(����ϵ��, 1), 0) *" & vbNewLine & _
                        "                            Decode(Upper(h.���㵥λ), 'ML', 1, Nvl(h.����ϵ��, 1)) ����, Nvl(f.��д����, 0) * Nvl(����ϵ��, 1) ����," & vbNewLine & _
                        "                           Decode(Nvl(f.�����, ''), '', 0, Nvl(f.��д����, 0) * Nvl(����ϵ��, 1)) �ѷ�," & vbNewLine & _
                        "                           Decode(Nvl(f.�����, ''), '', Nvl(f.��д����, 0) * Nvl(����ϵ��, 1), 0) δ��," & vbNewLine & _
                        "                           ' ���' || Decode(Substr('' || Nvl(f.��д����, 0) * Nvl(����ϵ��, 1), 1, 1), '.', 0, '') ||" & vbNewLine & _
                        "                            Nvl(f.��д����, 0) * Nvl(����ϵ��, 1) || h.���㵥λ || Decode(Nvl(f.�����, ''), '', '(δ��)', '(�ѷ�)') ||" & vbNewLine & _
                        "                            '  Ч��:' || To_Char(f.Ч��, 'yyyy-mm-dd hh24:mi') || '<Split3>' ||" & vbNewLine & _
                        "                            Decode(Nvl(f.�����, ''), '', 0, 1) ѪҺ��Ϣ, Decode(Nvl(f.�����, ''), '', 0, 1) �Ƿ��ѷ�, f.Ч��" & vbNewLine & _
                        "                    From ������ĿĿ¼ h, ѪҺ��� g, ѪҺ�շ���¼ f, ѪҺ��Ѫ��¼ e, ����ҽ����¼ b" & vbNewLine & _
                        "                    Where h.Id = g.Ʒ��id And g.���id = f.ѪҺid And f.�䷢id = e.Id And Mod(f.��¼״̬, 3) = 1 And" & vbNewLine & _
                        "                          Instr(',0,3,', ',' || f.��Ѫ״̬ || ',') = 0 And e.����id = b.Id And b.������� = 'K' And" & vbNewLine & _
                        "                          b.ҽ��״̬ In (1, 3, 8) " & strWhere & " And Exists" & vbNewLine & _
                        "                     (Select 1" & vbNewLine & _
                        "                           From ������ĿĿ¼ p, ����ҽ����¼ q" & vbNewLine & _
                        "                           Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = b.Id And" & vbNewLine & _
                        "                                 q.������� = 'E'))"

        gstrSQL = gstrSQL & vbNewLine & _
                        "       Select a.Id, a.����, a.����, a.���㵥λ, a.ִ�з��� ִ�з���id, a.ִ�п��� ִ�п���id, a.¼������ ¼������id, a.����ϵ��," & vbNewLine & _
                        "               (Select f_List2str(Cast(Collect('' || �����Ŀid) As t_Strlist)) From ������Ŀ��� Where ��Ŀid = a.Id) �����Ŀ, b.����," & vbNewLine & _
                        "               b.�ѷ�, b.δ��, �Ƿ��ѷ�, b.ѪҺ��Ϣ, b.Ч��, Decode(b.Id, Null, 0, 1) �Ƿ���Ѫ, b.����, 'ml' ������λ" & vbNewLine & _
                        "       From ������ĿĿ¼ a, ��Ѫ��¼ b," & vbNewLine & _
                        "           (Select Id ������Ŀid" & vbNewLine & _
                        "               From ��Ѫ��¼" & vbNewLine & _
                        "               Union" & vbNewLine & _
                        "               Select Decode(Nvl(c.ҽ��id, 0), 0, b.������Ŀid, c.������Ŀid) ������Ŀid" & vbNewLine & _
                        "               From ��Ѫ������Ŀ c, ����ҽ����¼ b" & vbNewLine & _
                        "               Where c.ҽ��id(+) = b.Id " & strWhere & " And b.������� = 'K' And" & vbNewLine & _
                        "                    b.ҽ��״̬ In (1, 3, 8) And Exists" & vbNewLine & _
                        "               (Select 1" & vbNewLine & _
                        "                    From ������ĿĿ¼ p, ����ҽ����¼ q" & vbNewLine & _
                        "                   Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = b.Id And q.������� = 'E')) c" & vbNewLine & _
                        "       Where a.Id = c.������Ŀid And c.������Ŀid = b.Id(+)) a" & vbNewLine & _
                        "       Order By a.����, a.�Ƿ��ѷ�, a.Ч��"
            
        If mint���� = 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mlng��ҳID, mbytBaby)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mstr�Һŵ�, mbytBaby)
        End If
        
        If rsTmp.RecordCount > 0 Then
            If mintType <> 0 And mintType <> 1 And mintType <> 4 Then 'ֻ���������޸�ʱ������ѡ��
                vsfBlood.Editable = flexEDNone
            Else
                vsfBlood.Editable = IIF(mblnUseBloodSend = True, flexEDNone, flexEDKbdMouse)
            End If
        Else
            vsfBlood.Editable = flexEDNone
        End If
        '��ʼ����������װ
        blnLast = False
        arrRecord = Array()
        str�䷢��Ϣ = "": strδ�� = "": str�ѷ� = "": str���� = "": str���� = ""
        lng������ĿID = -999: str�����Ŀ = "": str���㵥λ = ""
        Do While Not rsTmp.EOF
            If lng������ĿID <> Val("" & rsTmp!ID) Then
                If lng������ĿID <> -999 Then
GOWORK:
                    If str�䷢��Ϣ = "" Then
                        str�䷢��Ϣ = "��δ��Ѫ"
                    Else
                        str�䷢��Ϣ = "��Ѫ������" & IIF(Left(str����, 1) = ".", "0", "") & str���� & str���㵥λ & " �ѷ���" & IIF(Left(str�ѷ�, 1) = ".", "0", "") & str�ѷ� & str���㵥λ & " δ����" & IIF(Left(strδ��, 1) = ".", "0", "") & strδ�� & str���㵥λ & str�䷢��Ϣ
                    End If
                    '�䷢��Ϣ��ʽ����Ѫ������400ml �ѷ�����0ml δ������400ml<Split4> ���200ml(δ��)  Ч��:2016-09-17 16:13<Split3>0<Split4> ���200ml(δ��)  Ч��:2016-08-14 11:17<Split3>0
                    str��Ѫ��Ϣ = str�����Ŀ & "<Split2>" & lng������ĿID & "'" & str�䷢��Ϣ & "'" & str����
                    objCollection.Add str��Ѫ��Ϣ, "A_" & lng������ĿID
                    If blnLast = True Then GoTo GONEXT
                End If
                str�䷢��Ϣ = "": strδ�� = "": str�ѷ� = "": str���� = "": str���� = ""
                lng������ĿID = Val("" & rsTmp!ID)
                str�����Ŀ = "" & rsTmp!�����Ŀ
                str���㵥λ = "" & rsTmp!���㵥λ
                ReDim Preserve arrRecord(UBound(arrRecord) + 1)
                arrRecord(UBound(arrRecord)) = rsTmp!ID & "'" & rsTmp!���� & "'" & rsTmp!���� & "'" & rsTmp!���㵥λ & "'" & rsTmp!ִ�з���ID & "'" & rsTmp!ִ�п���ID & "'" & rsTmp!¼������ID & "'" & rsTmp!����ϵ��
            End If
            
            If Val(rsTmp!�Ƿ���Ѫ & "") = 1 Then
                str���� = Val(str����) + Val(rsTmp!���� & "")
                str�ѷ� = Val(str�ѷ�) + Val(rsTmp!�ѷ� & "")
                strδ�� = Val(strδ��) + Val(rsTmp!δ�� & "")
                str���� = Val(str����) + Val(rsTmp!���� & "")
                If rsTmp!ѪҺ��Ϣ & "" <> "" Then str�䷢��Ϣ = str�䷢��Ϣ & "<Split4>" & rsTmp!ѪҺ��Ϣ
            End If
            rsTmp.MoveNext
        Loop
        If lng������ĿID <> -999 Then
            blnLast = True
            GoTo GOWORK
        End If
GONEXT:
    End If
    
    arrItem = Split(mstr��Ѫ��Ŀ, ";")
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
            .TextMatrix(.Rows - 1, COL_P_ѡ��) = ""
            .TextMatrix(.Rows - 1, COL_P_����) = CStr(arrInfo(1))
            .TextMatrix(.Rows - 1, COL_P_����) = CStr(arrInfo(2))
            .TextMatrix(.Rows - 1, COL_P_������) = ""
            .TextMatrix(.Rows - 1, COL_P_��λ) = CStr(arrInfo(3))
            If InStr(1, "'" & UCase(str��λ) & "'", "'" & UCase(CStr(arrInfo(3))) & "'") = 0 Then
                str��λ = IIF(str��λ = "", "", str��λ & "'") & CStr(arrInfo(3))
            End If
            .TextMatrix(.Rows - 1, COL_P_����Ѫ��) = ""
            .TextMatrix(.Rows - 1, COL_P_����RH) = ""
            .TextMatrix(.Rows - 1, COL_P_ִ�з���ID) = Val(arrInfo(4))
            .TextMatrix(.Rows - 1, COL_P_ִ�п���ID) = Val(arrInfo(5))
            .TextMatrix(.Rows - 1, COL_P_¼������ID) = Val(arrInfo(6))
            .TextMatrix(.Rows - 1, COL_P_����ϵ��) = Val(arrInfo(7))
            .TextMatrix(.Rows - 1, COL_P_���) = CStr(ISExistCollection(objCollection, "A_" & Val(arrInfo(0))))
            
            Set .Cell(flexcpPicture, .Rows - 1, COL_P_ѡ��) = img16.ListImages("c0").Picture
            .Cell(flexcpData, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_ѡ��) = 0
            .Cell(flexcpFontBold, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = False
            .Cell(flexcpBackColor, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = vbWhite
            For j = 0 To UBound(arrItem)
                arrCode = Split(CStr(arrItem(j)), ",")
                If Val(arrCode(0)) = Val(arrInfo(0)) Then
                    Set .Cell(flexcpPicture, .Rows - 1, COL_P_ѡ��) = img16.ListImages("c1").Picture
                    .Cell(flexcpData, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_ѡ��) = 1
                    .Cell(flexcpFontBold, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = True
                    .Cell(flexcpBackColor, .Rows - 1, COL_P_ѡ��, .Rows - 1, COL_P_���) = &HC0E0FF
                    .TextMatrix(.Rows - 1, COL_P_������) = arrCode(1)
                    .TextMatrix(.Rows - 1, COL_P_����Ѫ��) = arrCode(2)
                    .TextMatrix(.Rows - 1, COL_P_����RH) = arrCode(3)
                    
                    txt������Ϣ.Text = txt������Ϣ.Text & "[" & .TextMatrix(.Rows - 1, COL_P_����) & IIF(.TextMatrix(.Rows - 1, COL_P_������) <> "", "-" & .TextMatrix(.Rows - 1, COL_P_������) & .TextMatrix(.Rows - 1, COL_P_��λ), "") & "]"
                End If
            Next
        Next
        If .Rows > .FixedRows Then
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
            .CellBorderRange .FixedRows, COL_P_������, .Rows - 1, COL_P_������, vbGreen, 1, 1, 1, 1, 1, 1
            If bln��Ѫ���� = False Then
                .CellBorderRange .FixedRows, COL_P_����Ѫ��, .Rows - 1, COL_P_����Ѫ��, vbGreen, 1, 1, 1, 1, 1, 1
                .CellBorderRange .FixedRows, COL_P_����RH, .Rows - 1, COL_P_����RH, vbGreen, 1, 1, 1, 1, 1, 1
            End If
            'ȷ�����ߴ�
            .AutoSize 0, .Cols - 1
            .ColWidth(COL_P_ѡ��) = 255
            .Redraw = flexRDDirect
            Call vsfBlood_AfterRowColChange(0, 0, 1, 1)
        Else
            .Redraw = flexRDDirect
        End If
    End With
    
    '���ص�λ
    arrItem = Split(str��λ, "'")
    cboInfo(cbo��λ).Clear
    cboInfo(cbo��λ).Tag = ""
    For i = 0 To UBound(arrItem)
        cboInfo(cbo��λ).AddItem CStr(arrItem(i))
        If UCase(txtInfo(txt��λ).Text) = UCase(CStr(arrItem(i))) Then
            Call zlControl.CboSetIndex(cboInfo(cbo��λ).hwnd, i)
            cboInfo(cbo��λ).Tag = i
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
'����:�ж�keyֵ�Ƿ���ڼ����У���������򷵻ض�Ӧ����,���򷵻ؿ�
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
'���ܣ�'�´���Ѫ����ʱ����ϡ���ѪĿ�ġ�Ѫ�͵�Ĭ��ȥ���һ�α�Ѫ�������Ϣ
    Dim strWhere As String, strSQL As String, str��� As String
    Dim rsTmp As New ADODB.Recordset, rsTmpOther As New ADODB.Recordset
    Dim lngҽ��ID As Long, int������־ As Integer, str��ҩ���� As String
    
    If mblnSpareBloood = True Then Exit Sub
    
    If mint���� = 0 Then
        strWhere = " And A.����ID=[1] And A.��ҳID=[2]"
    Else
        strWhere = " And A.����id =[1] And A.�Һŵ�=[2] "
    End If
    
    On Error GoTo ErrHand
    If lngActiveID = 0 Then
        '��ȡ���һ�α�Ѫ����
        strSQL = _
            " Select a.id,a.������־,a.��ҩ����" & vbNewLine & _
            " From ������ĿĿ¼ p, ����ҽ����¼ q, ����ҽ����¼ a" & vbNewLine & _
            " Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = a.Id And q.������� = 'E' And a.������� = 'K' And" & vbNewLine & _
            "      a.ҽ��״̬ In (1, 3, 8) " & strWhere & vbNewLine & _
            " Order By a.��ʼִ��ʱ�� Desc"
        If mint���� = 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���һ�α�Ѫ����", mlng����ID, mlng��ҳID)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���һ�α�Ѫ����", mlng����ID, mstr�Һŵ�)
        End If
        If rsTmp.EOF Then Exit Sub
        
        mblnDataLoad = True
        lngҽ��ID = Val("" & rsTmp!ID)
        int������־ = Val("" & rsTmp!������־)
        str��ҩ���� = "" & rsTmp!��ҩ����
        If int������־ = 1 Then
            cboInfo(cbo��Ѫ����).ListIndex = 1
        End If
        Call zlControl.CboSetText(cboInfo(cbo��ѪĿ��), str��ҩ����, True, "'")
    Else
        lngҽ��ID = lngActiveID
    End If
    
    strSQL = "Select �Ƿ����,��Ѫ����, ��ѪĿ��, ��Ѫ����, ��ѪѪ��, Rhd" & vbNewLine & _
                " From ��Ѫ�����¼" & vbNewLine & _
                " Where ҽ��id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        If Val(rsTmp!�Ƿ���� & "") = 1 Then
            txtInfo(txt�����Ϣ).Text = "����"
            chkWait.value = 1
        Else
            '��ȡ���
            mstr���IDs = GetAdviceDiag(lngҽ��ID, str���)
            txtInfo(txt�����Ϣ).Text = str���
            '�Ӹ����л�ȡ���������������Ը���Ϊ׼
             strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
             Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
             If Not rsTmpOther.EOF Then
                 txtInfo(txt�����Ϣ).Text = rsTmpOther!���� & ""
             End If
        End If
        txtInfo(txt�����Ϣ).Tag = txtInfo(txt�����Ϣ).Text
        chkWait.value = Val(rsTmp!�Ƿ���� & "")
        Call zlControl.CboSetText(cboInfo(cbo��Ѫ����), rsTmp!��Ѫ���� & "", True, "'")
        If "" & rsTmp!��ѪĿ�� <> "" Then Call zlControl.CboSetText(cboInfo(cbo��ѪĿ��), rsTmp!��ѪĿ�� & "", True, "'") '�ϵĵ���ѪĿ�Ĵ洢��ҽ������ҩ��������
        cboInfo(cbo��Ѫ����).ListIndex = Val(rsTmp!��Ѫ���� & "")
        cboInfo(cbo��ѪѪ��).ListIndex = Val(rsTmp!��ѪѪ�� & "")
        cboInfo(cboRHD).ListIndex = Val(rsTmp!RHD & "")
    End If
    'Ѫ�ͱ�Ѫ���뵥����û�У��Ӳ�����Ϣ�ӱ��л�ȡ
    If cboInfo(cbo��ѪѪ��).ListIndex <= 0 Then
        strSQL = "Select ��Ϣֵ from ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
        Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, "ABO") '������Ϣ����ABO,������'Ѫ��',���޷���ȡѪ�͵�ԭ��
        If Not rsTmpOther.EOF Then
            Select Case "" & rsTmpOther!��Ϣֵ
                Case "A", "A��"
                    cboInfo(cbo��ѪѪ��).ListIndex = 1
                Case "B", "B��"
                    cboInfo(cbo��ѪѪ��).ListIndex = 2
                Case "O", "O��"
                    cboInfo(cbo��ѪѪ��).ListIndex = 3
                Case "AB", "AB��"
                    cboInfo(cbo��ѪѪ��).ListIndex = 4
                Case "����"
                    cboInfo(cbo��ѪѪ��).ListIndex = 5
                Case "δ��"
                    cboInfo(cbo��ѪѪ��).ListIndex = 6
            End Select
        End If
    End If
    If cboInfo(cboRHD).ListIndex <= 0 Then
        strSQL = "Select ��Ϣֵ from ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��=[2]"
        Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, "RH")
        If Not rsTmpOther.EOF Then
            Select Case "" & rsTmpOther!��Ϣֵ
                Case "-", "��"
                    cboInfo(cboRHD).ListIndex = 1
                Case "+", "��"
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
'���ܣ��������������˵��ؼ���Ӧ������ֵ
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
'���ܣ��¿����޸���Ѫ����ʱ����������֮ǰ�������������ݽ��м�飬��������ʾ����������
'����ţ�116846:������,2017-11-23
    Dim strResult As String
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim var1 As Variant
    Dim var2 As Variant
    Dim strSQL As String, strMsg As String
    Dim rsTmp As New ADODB.Recordset
    Dim str������Ŀ As String
    Dim lng��Ѫִ�п���ID As Long, lng��Ѫ;��ִ�п���ID As Long
    
    On Error GoTo ErrHand
    var1 = Array()
    var2 = Array()
    '��Ѫ������м�����Ŀ
    If mblnSpareBloood = True Then
        '������Ŀ
        With vsLIS
            For i = 0 To .Rows - 1
                For j = 0 To CON_LisResultCol - 1
                    If Val(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) <> 0 Then
                        var1 = Array(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ��Ӣ���� + (j * CON_LisResultCount)), .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_�����λ + (j * CON_LisResultCount)), .TextMatrix(i, COL_�����־ + (j * CON_LisResultCount)), .TextMatrix(i, COL_����ο� + (j * CON_LisResultCount)), _
                            .TextMatrix(i, COL_ȡֵ���� + (j * CON_LisResultCount)), IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, 1, 0))
                        strTmp = Join(var1, "<SplitCol>")
                        ReDim Preserve var2(UBound(var2) + 1)
                        var2(UBound(var2)) = strTmp
                    End If
                Next
            Next
        End With
        strResult = Join(var2, "<SplitRow>")
    End If
    str������Ŀ = GetBloodInfo(False)
    lng��Ѫִ�п���ID = IIF(cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex) <= 0, 0, cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).ListIndex))
    lng��Ѫ;��ִ�п���ID = IIF(cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex) <= 0, 0, cboInfo(cbo��Ѫִ��).ItemData(cboInfo(cbo��Ѫִ��).ListIndex))
    
    If mblnSpareBloood = True Then '��Ѫ���뵥
        strSQL = "Select Zl1_EX_BloodApplyCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]) as ��� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_EX_BloodApplyCheck", IIF(1 = mint����, 1, 2), mlng����ID, IIF(mint���� = 1, mlng�Һ�ID, mlng��ҳID), IIF(mblnSpareBloood = True, 1, 2), _
            cboInfo(cbo��Ѫ����).ListIndex, chkWait.value, IIF(chkWait.value = 0, txtInfo(txt�����Ϣ).Text, ""), mstr���IDs, cboInfo(cbo��Ѫ����).Text, cboInfo(cbo��ѪĿ��).Text, cboInfo(cbo��Ѫ����).Text, txtInfo(txtԤ����Ѫʱ��).Text, _
            cboInfo(cbo��ѪѪ��).Text, cboInfo(cboRHD).Text, str������Ŀ, lng��Ѫִ�п���ID, mlng��Ѫ;��, lng��Ѫ;��ִ�п���ID, txtInfo(txt��ע).Text, mbytBaby, IIF(optHistory(0).value, 0, 1), _
            IIF(optHistory(2).value, 0, 1), IIF(optHistory(4).value, 0, 1), txtInfo(txt��) & "/" & txtInfo(txt��), IIF(optPossession(0).value, 0, 1), strResult)
    Else '��Ѫ���뵥
        strSQL = "Select Zl1_EX_BloodApplyCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20]) as ��� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_EX_BloodApplyCheck", IIF(1 = mint����, 1, 2), mlng����ID, IIF(mint���� = 1, mlng�Һ�ID, mlng��ҳID), IIF(mblnSpareBloood = True, 1, 2), _
            cboInfo(cbo��Ѫ����).ListIndex, chkWait.value, IIF(chkWait.value = 0, txtInfo(txt�����Ϣ).Text, ""), mstr���IDs, cboInfo(cbo��Ѫ����).Text, cboInfo(cbo��ѪĿ��).Text, cboInfo(cbo��Ѫ����).Text, txtInfo(txtԤ����Ѫʱ��).Text, _
            cboInfo(cbo��ѪѪ��).Text, cboInfo(cboRHD).Text, str������Ŀ, lng��Ѫִ�п���ID, mlng��Ѫ;��, lng��Ѫ;��ִ�п���ID, txtInfo(txt��ע).Text, mbytBaby)
    End If
    
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!���)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '��ʾ
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '��ֹ
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
'���ܣ���Ѫ����ʱ��1�������LIS���������Զ�����LIS��������Ϣ��2������ʱ���ABO��RH�Ƿ���д��һ�£���һ�����ֹ
'������blnReset��TRUE ���ݼ���������ABO��RH��False�������Ǽ���Ƿ�ͼ�����һ��
'         lngRow��lngCol �ڱ༭������ʱ������ã�����ͬ������ABO��RH��blnReset=trueʱ��Ч��,lngCOl��ǰ�༭��ָ������

    Dim lngCount As Long
    Dim i As Integer, j As Integer
    Dim strAboCode As String, strRHCode As String
    Dim intAboType As Integer, intRHType As Integer
    Dim ArrResult(0 To 1) As String, arrCode() As String
    Dim blnIsAbo As Boolean, blnIsRH As Boolean  '������Ƿ񷵻���ABO��RH
    Dim strTemp As String, blnMsg As Boolean
    Dim strResult As String

    
    If mblnSpareBloood = False Then CheckOrResetLisAboRH = True: Exit Function '��Ѫ������м�����
    
    '�кϷ��Լ��
    If lngRow <> -1 Then
        If Not (lngRow >= vsLIS.FixedRows And lngRow < vsLIS.Rows) Then Exit Function
    End If
    If lngCol <> -1 Then
        '��ֻ����ָ������
        If Not (lngCol Mod 10 = COL_ָ����) Then Exit Function
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
            If Val(.TextMatrix(lngRow, lngCol + (COL_������ĿID - COL_ָ����))) <> 0 Then
                If strAboCode = .TextMatrix(lngRow, lngCol + (COL_ָ����� - COL_ָ����)) And strAboCode <> "" Then
                    ArrResult(0) = .TextMatrix(lngRow, lngCol + (COL_ָ�������� - COL_ָ����)) & "'" & .TextMatrix(lngRow, lngCol)
                    blnIsAbo = IIF(.Cell(flexcpBackColor, lngRow, lngCol) = COLEditBackColor, False, True)
                End If
                
                If strRHCode = .TextMatrix(lngRow, lngCol + (COL_ָ����� - COL_ָ����)) And strRHCode <> "" Then
                    ArrResult(1) = .TextMatrix(lngRow, lngCol + (COL_ָ�������� - COL_ָ����)) & "'" & .TextMatrix(lngRow, lngCol)
                    blnIsRH = IIF(.Cell(flexcpBackColor, lngRow, lngCol) = COLEditBackColor, False, True)
                    .Cell(flexcpForeColor, lngRow, lngCol) = IIF(.TextMatrix(lngRow, lngCol) = "-", vbRed, &H80000012)
                End If
            End If
        End With
    Else
        '��ȡABO��RH�ڼ������ж�Ӧ�ĵ�ָ�����ƺ�ָ����
        With vsLIS
            lngCount = 0
            For i = 0 To .Rows - 1
                For j = 0 To CON_LisResultCol - 1
                    If Val(.TextMatrix(i, COL_������ĿID + (j * CON_LisResultCount))) <> 0 Then
                        If strAboCode = .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)) And strAboCode <> "" Then
                            ArrResult(0) = .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)) & "'" & .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount))
                            blnIsAbo = IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, False, True)
                        End If
                        If strRHCode = .TextMatrix(i, COL_ָ����� + (j * CON_LisResultCount)) And strRHCode <> "" Then
                            ArrResult(1) = .TextMatrix(i, COL_ָ�������� + (j * CON_LisResultCount)) & "'" & .TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount))
                            blnIsRH = IIF(.Cell(flexcpBackColor, i, COL_ָ���� + (j * CON_LisResultCount)) = COLEditBackColor, False, True)
                            .Cell(flexcpForeColor, i, COL_ָ���� + (j * CON_LisResultCount)) = IIF(.TextMatrix(i, COL_ָ���� + (j * CON_LisResultCount)) = "-", vbRed, &H80000012)
                        End If
                    End If
                Next
            Next
        End With
    End If
    If blnReset = True Then
        '116848
        '���ݼ���������ABO��RH
        If ArrResult(0) <> "" Then
            arrCode = Split(ArrResult(0), "'")
            strTemp = UCase(arrCode(1))
            If strTemp = "A" Or strTemp = "A��" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 1
            ElseIf strTemp = "B" Or strTemp = "B��" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 2
            ElseIf strTemp = "AB" Or strTemp = "AB��" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 4
            ElseIf strTemp = "O" Or strTemp = "O��" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 3
            ElseIf strTemp = "����" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 5
            ElseIf strTemp = "δ��" Then
                cboInfo(cbo��ѪѪ��).ListIndex = 6
            Else
                If strTemp = "" Then
                    If blnIsAbo = True Then '����(����LISδ�����)
                        cboInfo(cbo��ѪѪ��).ListIndex = 5
                    Else 'δ�� (����δ��Ѫ�ͼ��)
                        cboInfo(cbo��ѪѪ��).ListIndex = 6
                    End If
                Else
                    cboInfo(cbo��ѪѪ��).ListIndex = 0
                End If
            End If
        End If
        
        If ArrResult(1) <> "" Then
            arrCode = Split(ArrResult(1), "'")
            strTemp = arrCode(1)
            If strTemp = "-" Or strTemp Like "����*" Then
                cboInfo(cboRHD).ListIndex = 1
            ElseIf strTemp = "+" Or strTemp Like "����*" Then
                cboInfo(cboRHD).ListIndex = 2
            Else
                cboInfo(cboRHD).ListIndex = 0
            End If
        End If
    Else
        '���һ���Լ��
        For i = 0 To 1
            If ArrResult(i) <> "" Then
                arrCode = Split(ArrResult(i), "'")
                If i = 0 Then
                    strTemp = UCase(cboInfo(cbo��ѪѪ��).Text)
                    strResult = UCase(arrCode(1))
                Else
                    strTemp = cboInfo(cboRHD).Text
                    If arrCode(1) Like "����*" Then
                        strResult = "+"
                    ElseIf arrCode(1) Like "����*" Then
                        strResult = "-"
                    Else
                        strResult = arrCode(1)
                    End If
                End If
                If strResult <> "" And strTemp <> strResult Then
                    blnMsg = True
                    If i = 0 Then
                        If strResult Like "*��" And Trim(strTemp) <> "" And strTemp & "��" = strResult Then
                            blnMsg = False
                        End If
                    End If
                    If blnMsg = True Then
                        If i = 0 Then
                            If intAboType = 0 Then
                                If MsgBox("���뵥�е�Ѫ�ͺͼ�������ָ��[" & arrCode(0) & "]�Ľ���������������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Function
                                End If
                            Else
                                MsgBox "���뵥�е�Ѫ�ͺͼ�������ָ��[" & arrCode(0) & "]�Ľ�����������飡", vbInformation, gstrSysName
                                Exit Function
                            End If
                        Else
                            If intRHType = 0 Then
                                If MsgBox("���뵥�е�RHD�ͼ�������ָ��[" & arrCode(0) & "]�Ľ���������������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Function
                                End If
                            Else
                                MsgBox "���뵥�е�RHD�ͼ�������ָ��[" & arrCode(0) & "]�Ľ�����������飡", vbInformation, gstrSysName
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
'���뷵��ѡ���ѪҺ��Ϣ����ʽ��������ĿID,������,����Ѫ��,����RH;������ĿID,������,����Ѫ��,����RH
    Dim lngRow As Long
    Dim strRow As String, strTmp As String
    Dim strIDs As String

    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                'ҽ��ѡ��Ѫ����Ϣ������Ҫ��¼ѡ����շ�ID
                strIDs = ""
                If mblnSelectBlood Then
                    strTmp = .TextMatrix(lngRow, COL_P_���)
                    If InStr(1, strTmp, "<Split3>") <> 0 Then
                        strIDs = Split(strTmp, "<Split3>")(1) '��ȡ��ѡ�е�ѪҺ��Ϣ
                    End If
                End If
                If blnABO = True Then
                    strRow = strRow & ";" & .TextMatrix(lngRow, COL_P_ID) & "," & .TextMatrix(lngRow, COL_P_������) & "," & .TextMatrix(lngRow, COL_P_����Ѫ��) & "," & .TextMatrix(lngRow, COL_P_����RH) & IIF(strIDs <> "", "," & strIDs, "")
                Else
                    strRow = strRow & ";" & .TextMatrix(lngRow, COL_P_ID) & "," & .TextMatrix(lngRow, COL_P_������) & IIF(strIDs <> "", "," & strIDs, "")
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
    
    lngLeft = txtInfo(txt��ע).Left
    lngTop = txtInfo(txt��ע).Top
    strName = "�������С�"
    
    lngLeft = lngLeft + Me.Left
    lngTop = lngTop + Me.Top - 2700
    
    strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, 2)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "û���ҵ����õ�" & strName, vbInformation, Me.Caption
            End If
        Else
            txtInfo(txt��ע).Text = strRetrun
        End If
    End If
End Sub

Private Sub Get��Ѫִ�п���()
'ѪҺ��Ŀ��ִ�п��ң��϶���Ѫ�����
    Dim bln�ϰల�� As Boolean
    Dim strSQL As String, bytDay As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    bln�ϰల�� = Check�ϰల��(False)
    If Not bln�ϰల�� Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� C" & _
            " Where  A.ID=C.����ID" & _
            " And C.������� IN([1],3) And C.��������='Ѫ��'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by ����"
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
        strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����" & _
            " From ���Ű��� B,���ű� C,��������˵�� D" & _
            " Where  B.����ID=C.ID And B.����=[2]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
            " And C.ID=D.����ID And D.������� IN([1],3) And C.��������='Ѫ��'" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " Order by ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "", IIF(mlng�������� = 1, 1, 2), bytDay)
    cboInfo(cboִ�п���).Clear
    For i = 1 To rsTmp.RecordCount
        cboInfo(cboִ�п���).AddItem rsTmp!���� & "-" & rsTmp!����
        cboInfo(cboִ�п���).ItemData(cboInfo(cboִ�п���).NewIndex) = CLng(rsTmp!ID)
'        If lngDeptID = rsTmp!ID Then
'            Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
'        End If
        rsTmp.MoveNext
    Next
    If cboInfo(cboִ�п���).ListIndex = -1 And cboInfo(cboִ�п���).ListCount > 0 Then
        Call zlControl.CboSetIndex(cboInfo(cboִ�п���).hwnd, 0)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub BloodSum()
'���ܣ���Ѫ����������������ܶ��Ʒ�ֵĵ�λ��һ������Ҫ����ת����
    Dim lngRow As Integer, i As Integer
    Dim str��λ As String, dblNum As Double, dbl����ϵ�� As Double
    Dim dblSum As Double '����
    Dim strCur��λ As String, dblCur����ϵ�� As Double
    
    txt������Ϣ.Text = "Ʒ��:"
    txt������.Text = ""
    For i = vsfBlood.FixedRows To vsfBlood.Rows - 1
        If Val(vsfBlood.Cell(flexcpData, i, COL_P_ѡ��)) = 1 Then
            'iif(mid("" & 0.5,1,1)=".","0","") & 0.5������д����Ϊ�˱�֤С�ڵ�1��ֵ��������ʾǰ׺0
            txt������Ϣ.Text = txt������Ϣ.Text & "[" & vsfBlood.TextMatrix(i, COL_P_����) & IIF(vsfBlood.TextMatrix(i, COL_P_������) <> "", "-" & IIF(Mid("" & vsfBlood.TextMatrix(i, COL_P_������), 1, 1) = ".", "0", "") & vsfBlood.TextMatrix(i, COL_P_������) & vsfBlood.TextMatrix(i, COL_P_��λ), "") & "]"
        End If
    Next
    
    If cboInfo(cbo��λ).ListIndex >= 0 Then
        strCur��λ = UCase(cboInfo(cbo��λ).List(cboInfo(cbo��λ).ListIndex))
    End If
    
    With vsfBlood
        If strCur��λ <> "" Then
            For lngRow = .FixedRows To .Rows - 1
                If strCur��λ = UCase(.TextMatrix(lngRow, COL_P_��λ)) Then
                    dblCur����ϵ�� = Val(.TextMatrix(lngRow, COL_P_����ϵ��))
                    Exit For
                End If
            Next
        End If
        If dblCur����ϵ�� <= 0 Then dblCur����ϵ�� = 1
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                str��λ = UCase(.TextMatrix(lngRow, COL_P_��λ))
                dbl����ϵ�� = Val(.TextMatrix(lngRow, COL_P_����ϵ��))
                If dbl����ϵ�� <= 0 Then dbl����ϵ�� = 1
                dblNum = Val(.TextMatrix(lngRow, COL_P_������))
                If strCur��λ = "" Then
                    For i = 0 To cboInfo(cbo��λ).ListCount - 1
                        If str��λ = UCase(cboInfo(cbo��λ).List(i)) Then
                            Call zlControl.CboSetIndex(cboInfo(cbo��λ).hwnd, i)
                            strCur��λ = str��λ
                            dblCur����ϵ�� = dbl����ϵ��
                            cboInfo(cbo��λ).Tag = i
                            Exit For
                        End If
                    Next
                End If
                If UCase(strCur��λ) = UCase(str��λ) Then
                    dblSum = dblSum + dblNum
                Else
                    If str��λ <> "ML" Then
                        dblNum = dblNum * dbl����ϵ��
                    End If
                    dblSum = dblSum + Format(dblNum / dblCur����ϵ��, "#0.00;-#0.00")
                End If
            End If
        Next
    End With
    If Val(dblSum) = 0 Then
        txtInfo(txtԤ����Ѫ��) = ""
        txt������.Text = ""
    Else
        txtInfo(txtԤ����Ѫ��).Text = zl9ComLib.FormatEx(dblSum, 5)
        txt������.Text = zl9ComLib.FormatEx(dblSum, 5)
    End If
    txtInfo(txt��λ).Text = strCur��λ
End Sub

Private Sub RsetBreedUnit()
'���ܣ���λ�л����ҽ��Ĭ��Ʒ�ֵĵ�λ��ѡ��λ�����������Ʒ���л�(��Ѫ��ѡ����Ʒ�֣���ҽ����¼ֻ�ܼ�¼һ��Ʒ��)
    Dim lngRow As Long
    Dim strCur��λ, str��λ As String
    If mlng��Ѫ��ĿID <= 0 Then Exit Sub
    If cboInfo(cbo��λ).ListIndex >= 0 Then
        strCur��λ = UCase(cboInfo(cbo��λ).List(cboInfo(cbo��λ).ListIndex))
    Else
        Exit Sub
    End If
    
    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If mlng��Ѫ��ĿID = Val(.TextMatrix(lngRow, COL_P_ID)) Then
                str��λ = UCase(.TextMatrix(lngRow, COL_P_��λ))
                Exit For
            End If
        Next
        If strCur��λ <> str��λ Then
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                    If strCur��λ = UCase(.TextMatrix(lngRow, COL_P_��λ)) Then
                        mlng��Ѫ��ĿID = Val(.TextMatrix(lngRow, COL_P_ID))
                        txtInfo(txt��λ).Text = .TextMatrix(lngRow, COL_P_��λ)
                        mlng¼������ = Val(.TextMatrix(lngRow, COL_P_¼������ID))
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Function GetBloodTotalByML() As Double
    '���ܣ�����������Ѫ����(ת����ΪML)
    Dim lngRow As Long
    Dim str��λ As String, dbl����ϵ�� As Double, dblNum As Double
    Dim dblTotal As Double
    With vsfBlood
        For lngRow = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                str��λ = UCase(.TextMatrix(lngRow, COL_P_��λ))
                dbl����ϵ�� = Val(.TextMatrix(lngRow, COL_P_����ϵ��))
                If dbl����ϵ�� <= 0 Then dbl����ϵ�� = 1
                dblNum = Val(.TextMatrix(lngRow, COL_P_������))
                If str��λ <> "ML" Then
                    dblNum = dblNum * dbl����ϵ��
                End If
                dblTotal = dblTotal + dblNum
            End If
        Next
    End With
    GetBloodTotalByML = dblTotal
End Function

Private Function GetPatiHisBloodItem() As String
'���ܣ���ȡ���˱��ξ������ʷ����ѪҺƷ����Ϣ
    Dim strSQL As String, strWhere As String
    Dim rsTmp As New ADODB.Recordset
    Dim strRetrun As String
    On Error GoTo ErrHand
     
    If mint���� = 0 Then
        strWhere = " And b.����id = [1] And b.��ҳid = [2] And Nvl(b.Ӥ��, 0) = [3] And B.id<>[4]"
    Else
        strWhere = " And b.����id = [1] And b.�Һŵ� = [2] And Nvl(b.Ӥ��, 0) = [3] And B.id<>[4]"
    End If
    strSQL = _
        " Select ����" & vbNewLine & _
        " From ������ĿĿ¼ a," & vbNewLine & _
        "     (Select ������Ŀid, Min(��ʼִ��ʱ��) ��ʼִ��ʱ��" & vbNewLine & _
        "       From (Select Decode(Nvl(c.ҽ��id, 0), 0, b.������Ŀid, c.������Ŀid) ������Ŀid, b.��ʼִ��ʱ��" & vbNewLine & _
        "              From ��Ѫ������Ŀ c, ����ҽ����¼ b" & vbNewLine & _
        "              Where c.ҽ��id(+) = b.Id And b.������� = 'K' And b.ҽ��״̬ In (1, 3, 8) " & strWhere & " And Exists" & vbNewLine & _
        "               (Select 1" & vbNewLine & _
        "                     From ������ĿĿ¼ p, ����ҽ����¼ q" & vbNewLine & _
        "                     Where p.Id = q.������Ŀid And (p.�������� = 9 Or p.�������� = 8 And p.ִ�з��� = 0) And q.���id = b.Id And q.������� = 'E'))" & vbNewLine & _
        "       Group By ������Ŀid) b" & vbNewLine & _
        " Where a.Id = b.������Ŀid" & vbNewLine & _
        " Order By ��ʼִ��ʱ��"
    If mint���� = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mlng��ҳID, mbytBaby, mlngUpdateAdvice)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ѫ��Ϣ", mlng����ID, mstr�Һŵ�, mbytBaby, mlngUpdateAdvice)
    End If
    Do While Not rsTmp.EOF
        strRetrun = strRetrun & "'" & rsTmp!����
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

Private Function GetBloodApplyCode(ByVal intģʽ As Integer) As String
'intģʽ:0=�����Ƿ������޸�ABO��RH��1=����ABO��RHָ����𣬱��ڸ��ݼ���������ABO��RH���Լ������Ǽ��ABO��RH�Ƿ�ͼ�����һ��(��Ѫ���뵥ʱ��Ч)
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    gstrSQL = "Select Zl_Fun_BloodApplyCode([1],[2],[3]) as ָ�� from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Zl_Fun_BloodApplyCode", IIF(mblnSpareBloood = True, 1, 2), cboInfo(cbo��Ѫ����).ListIndex, intģʽ)
    If Not rsTemp.EOF Then
        strValue = "" & rsTemp!ָ��
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
    If intIndex = cbo��ѪѪ�� Or intIndex = cboRHD Then
        With vsfBlood
            If .ColHidden(IIF(intIndex = cbo��ѪѪ��, COL_P_����Ѫ��, COL_P_����RH)) = True Then Exit Sub
            For lngRow = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, COL_P_ѡ��)) = 1 Then
                    If intIndex = cbo��ѪѪ�� Then
                        If InStr(1, ",A,AB,B,O,", "," & cboInfo(intIndex).Text & ",") <> 0 Then
                            .TextMatrix(lngRow, COL_P_����Ѫ��) = cboInfo(intIndex).Text
                        Else
                            .TextMatrix(lngRow, COL_P_����Ѫ��) = ""
                        End If
                    Else
                        .TextMatrix(lngRow, COL_P_����RH) = cboInfo(intIndex).Text
                    End If
                End If
            Next lngRow
        End With
    End If
End Sub
