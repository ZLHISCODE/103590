VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDoctorManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҽ����Ȩ����"
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
   StartUpPosition =   1  '����������
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "סԺ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�ļ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "һ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��Ȩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "סԺ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�д���Ȩ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�޴���Ȩ"
            BeginProperty Font 
               Name            =   "����"
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
      Begin VB.Image img����Ȩ 
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����I��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ļ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "һ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ҽ��Ȩ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����Ȩ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ҽ����Ϣ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����ȼ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ҩ��Ȩ��"
         BeginProperty Font 
            Name            =   "����"
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
            Begin VB.Image Pic����Ȩ 
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
               Caption         =   "ְ��"
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
               Caption         =   "��"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
            Begin VB.Image Pic����Ȩ 
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
               Caption         =   "ְ��"
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
               Caption         =   "��"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
            Begin VB.Image Pic����Ȩ 
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
               Caption         =   "ְ��"
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
               Caption         =   "��"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
               Caption         =   "����"
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
               Caption         =   "��"
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
               Caption         =   "ְ��"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   1200
               TabIndex        =   48
               Top             =   720
               Width           =   360
            End
            Begin VB.Image Pic����Ȩ 
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
               Caption         =   "����"
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
               Caption         =   "��"
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
               Caption         =   "ְ��"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   38
               Top             =   720
               Width           =   360
            End
            Begin VB.Image Pic����Ȩ 
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
            Begin VB.Image Pic����Ȩ 
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
               Caption         =   "ְ��"
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
               Caption         =   "��"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "סԺ��"
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
               Caption         =   "סԺ��"
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
               Caption         =   "���"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "����(&F)"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����(&D)"
      BeginProperty Font 
         Name            =   "����"
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
Private mstrPrivs As String               '��ǰ�û�ӵ��Ȩ�޵��ַ���
Private mobjBar As CommandBar             'Comandbar����
Private mstrUserDept As String            '��ǰҽ�����ڿ���
Private mrsTmp As ADODB.Recordset     '��ǰҽ����Ϣ�������ݼ�
Private mlngFindNum As Long               '�Ƿ����ò��ҹ���
Private mlngFilter As Long               '�Ƿ����ù��˹���
Private mlngIndex As Long                 '������Pic��Index
Private mblnPic As Boolean                '�Ƿ������Ƭ
Private mblnMove As Boolean               '�Ƿ�رչ���ͼƬpic
Private mstrFdsql As String               '����������sql
Private Const conMenu_View_Photo = 999    '����ҽ����Ƭ�İ�ťid
Private mcolCtl As New Collection         '����ؼ����󼯺�



Private Const conCtl = ",picInfo,lblName,lblSex,lblZw,cboLevel,Pic����Ȩ,PicMZFxz,PicMZXz,PicMZTs,PicZYFxz,PicZYXz," & _
                        "PicZYTs,PicSS1,PicSS2,PicSS3,PicSS4,PicMZDl,PicMZMz,PicMZJs,PicMZGz,PicZYDl,PicZYMz,PicZYJs,PicZYGz"

Private Enum mEnumPrivsType        'ҽ��Ȩ�����ͣ�=1 ����Ȩ,=2 ���￹��ҩ��Ȩ��,=3 סԺ����ҩ��Ȩ��,=4 �����ȼ�Ȩ��,=5 ��������ҽ��Ȩ��,=6 סԺ����ҽ��Ȩ��,=7 ҽ���ȼ�
    Priv����Ȩ = 1
    Priv���￹�� = 2
    PrivסԺ���� = 3
    Priv�����ȼ� = 4
    Priv�������� = 5
    PrivסԺ���� = 6
    Privҽ���ȼ� = 7
End Enum

Private Enum imgType
    imgType����Ȩ = 0
    imgType����ҩ�� = 1
    imgType�����ȼ� = 2
    imgType���� = 3
    imgType���� = 4
    imgType���� = 5
    imgType���� = 6
End Enum

Private Enum FilterType
    Filter����Ȩ = 0
    Filter����ҽ�� = 1
    Filter�����ȼ� = 2
    Filter����ҩ�� = 3
End Enum

Private Enum TsType
    Ts���� = 1
    Ts���� = 2
    Ts���� = 3
    Ts���� = 4
End Enum

Private Sub cboDept_Click()
    '���ص�ǰѡ��Ĳ���
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
    '����ҽ���ȼ�
    Dim strSql As String
    Dim intType As Integer
    Dim strCheck As String
On Error GoTo errH
    strCheck = cboLevel(Index).Text
    strSql = "Zl_ҽ��Ȩ��_Update(" & Val(picInfo(Index).Tag) & "," & Privҽ���ȼ� & ",'" & strCheck & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Call CacheData(Val(picInfo(Index).Tag), Privҽ���ȼ�, strCheck)
    picUser.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboLevel_KeyPress(Index As Integer, KeyAscii As Integer)
    '�������ֶ�����ҽ���ȼ�
    KeyAscii = 0
End Sub

Private Sub LoadCbo()
    '����ҽ���ȼ�����������
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim i As Integer
    
On Error GoTo errH
    strSql = "select ���� from רҵ����ְ�� where decode( substr(����,1,2),����,null ,substr(����,1,2))=23"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    For i = 0 To 5
        rsTmp.MoveFirst
        cboLevel.Item(i).AddItem ""
        Do While Not rsTmp.EOF
            cboLevel.Item(i).AddItem rsTmp!���� & ""
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
    '�رչ���ҳ��
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

'��֯����SQL
Private Function FilterSQL() As Boolean
    Dim i As Integer
    Dim strSql As String
    Dim strFilter As String
    Dim j As Integer
    Dim strTmp As String
    Dim strFind As String
    On Error GoTo errH
    
    '�Ƿ����ò���
    strFind = Replace(Replace(Replace(txtFind.Text, "'", ""), "*", ""), "%", "")
    If strFind = "" Then mlngFindNum = 0
    If mlngFindNum = 1 Then
        If zlCommFun.IsCharChinese(strFind) Then
            strSql = strSql & " And ���� like '*" & strFind & "*'"
        ElseIf IsNumeric(strFind) Then
            strSql = strSql & " And ��� like '*" & strFind & "*'"
        Else
            strSql = strSql & " And ���� like '*" & strFind & "*'"
        End If
    End If
    

    mlngFilter = 0
    For i = 0 To imgSentence.Count - 1
        If imgSentence(i).Tag <> "" Then
            '��������
            mlngFilter = 1
            strTmp = ""
            Select Case i
                '����Ȩ
                Case imgType����Ȩ
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And ����Ȩ��־ <> 1"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And ����Ȩ��־ <> 0"
                '����סԺ����ҩ��
                Case imgType����ҩ��
                  '����
                  For j = 0 To 3
                    strTmp = decode(j, 0, "������ʹ��", 1, "����ʹ��", 2, "����ʹ��", 3, "null")
                    If Split(imgSentence(i).Tag, "|")(j) = "0" Then strSql = strSql & " And ���￹��ҩ��Ȩ�� <> '" & strTmp & "'"
                  Next
                  'סԺ
                  For j = 0 To 3
                    strTmp = decode(j, 0, "������ʹ��", 1, "����ʹ��", 2, "����ʹ��", 3, "null")
                    If Split(imgSentence(i).Tag, "|")(j + 4) = "0" Then strSql = strSql & " And סԺ����ҩ��Ȩ�� <> '" & strTmp & "'"
                  Next
                '�����ȼ�
                Case imgType�����ȼ�
                  For j = 0 To 4
                    strTmp = decode(j, 0, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, "null")
                    If Split(imgSentence(i).Tag, "|")(j) = "0" Then strSql = strSql & " And �����ȼ� <> '" & strTmp & "'"
                  Next
                '����
                Case imgType����
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And ���ﶾ�� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And ���ﶾ�� <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And סԺ���� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And סԺ���� <> '0'"
                '����
                Case imgType����
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And �������� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And �������� <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And סԺ���� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And סԺ���� <> '0'"
                 '����
                Case imgType����
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And ���ﾫ�� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And ���ﾫ�� <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And סԺ���� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And סԺ���� <> '0'"
                  '����
                Case imgType����
                  If Split(imgSentence(i).Tag, "|")(0) = "0" Then strSql = strSql & " And ������� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(1) = "0" Then strSql = strSql & " And ������� <> '0'"
                  If Split(imgSentence(i).Tag, "|")(2) = "0" Then strSql = strSql & " And סԺ���� <> '1'"
                  If Split(imgSentence(i).Tag, "|")(3) = "0" Then strSql = strSql & " And סԺ���� <> '0'"
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
'��������������ȫ��ȡ��
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
    '����֧�ֹ���
    picBack.SetFocus
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FmgFlexScroll
End Sub

Private Sub Form_Deactivate()
    '����֧�ֹ���
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    On Error GoTo errH
    mlngIndex = 999
    mlngFindNum = 0
    mstrPrivs = gstrPrivs
    mstrFdsql = ""
    mstrUserDept = GetUser����IDs(True)
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    Call RestoreWinState(Me, App.ProductName)
    '��ʼ��Ȩ�޿ؼ�����
    Set mcolCtl = GetControls
    '��ʼ��������Ϣ
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
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "������Ȩ����(&N)")
            objControl.IconId = 9002
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Photo, "����ҽ����Ƭ(&P)")
        objControl.Checked = False
        mblnPic = objControl.Checked
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("������", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "������Ȩ����")
        objControl.IconId = 9002
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
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
        Case conMenu_Tool_OPSEmpower '������Ȩ����
            'Ȩ���жϺʹ���
            gstrPrivs = GetPrivFunc(glngSys, 1080)
            If gstrPrivs = "" Then
                MsgBox "��û��������Ȩ����Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Sub
            End If
            frmOPSEmpower.Show , Me
        Case conMenu_View_Refresh 'ˢ��
            Call vscBar_Change
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Function GetControls() As Collection
    '���ConCtl�ַ��������пؼ��Ķ���
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
    '���ؽ�����Ϣ
    Dim str����Ȩ As String
    Dim strKSSMz As String
    Dim strKSSZy As String
    Dim strOPS As String
    Dim strTsMz As String
    Dim strTsZy As String
    Dim strTempFile As String
    Dim i As Integer
    Dim objTmp As Object
    
    On Error GoTo errH
    '�������
    clearInfo intIndex
    
    str����Ȩ = Val(mrsTmp!����Ȩ��־ & "")
    strKSSMz = mrsTmp!���￹��ҩ��Ȩ�� & ""
    strKSSZy = mrsTmp!סԺ����ҩ��Ȩ�� & ""
    strOPS = mrsTmp!�����ȼ� & ""
    strTsMz = IIf(mrsTmp!��������ҽ��Ȩ�� & "" <> "", mrsTmp!��������ҽ��Ȩ�� & "", "0000")
    strTsZy = IIf(mrsTmp!סԺ����ҽ��Ȩ�� & "" <> "", mrsTmp!סԺ����ҽ��Ȩ�� & "", "0000")
    
    '����ҽ����Ϣ
    cboLevel(intIndex).Text = mrsTmp!רҵ����ְ�� & ""
    lblName.Item(intIndex).Caption = mrsTmp!���� & ""
    lblSex.Item(intIndex).Caption = mrsTmp!�Ա� & ""
    lblZw.Item(intIndex).Caption = mrsTmp!����ְ�� & ""
    picInfo.Item(intIndex).Tag = mrsTmp!ID & ""
    
    '���ش���Ȩ
    If str����Ȩ = 1 Then
        Pic����Ȩ(intIndex).Picture = img16.ListImages.Item("Black2").Picture: Pic����Ȩ(intIndex).Tag = "True"
    End If
    
    For i = 1 To 4
        '����סԺ����ҽ��Ȩ��
        Set objTmp = mcolCtl(decode(i, Ts����, "PicZYDl", Ts����, "PicZYMz", Ts����, "PicZYJs", Ts����, "PicZYGz") & intIndex)
        If Val(Mid(strTsZy, i, 1)) = 1 Then objTmp.Picture = img16.ListImages.Item("Green2").Picture: objTmp.Tag = "True"
        '������������ҽ��Ȩ��
        Set objTmp = mcolCtl(decode(i, Ts����, "PicMZDl", Ts����, "PicMZMz", Ts����, "PicMZJs", Ts����, "PicMZGz") & intIndex)
        If Val(Mid(strTsMz, i, 1)) = 1 Then objTmp.Picture = img16.ListImages.Item("Green2").Picture: objTmp.Tag = "True"
    Next
    
    '���������ȼ�
    If strOPS <> "null" Then
        Set objTmp = mcolCtl(decode(strOPS, "һ��", "PicSS1", "����", "PicSS2", "����", "PicSS3", "�ļ�", "PicSS4") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Blue2").Picture: objTmp.Tag = "True"
    End If

    '�������￹��ҩ��Ȩ��
    If strKSSMz <> "null" Then
        Set objTmp = mcolCtl(decode(strKSSMz, "������ʹ��", "PicMZFxz", "����ʹ��", "PicMZXz", "����ʹ��", "PicMZTs") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Red2").Picture: objTmp.Tag = "True"
    End If

    '����סԺ����ҩ��Ȩ��
    If strKSSZy <> "null" Then
        Set objTmp = mcolCtl(decode(strKSSZy, "������ʹ��", "PicZYFxz", "����ʹ��", "PicZYXz", "����ʹ��", "PicZYTs") & intIndex)
        objTmp.Picture = img16.ListImages.Item("Red2").Picture: objTmp.Tag = "True"
    End If
    
    '��ʾ��Ƭ
    If mblnPic = True Then
        If mrsTmp!��Ƭ��С & "" <> "" And Val(mrsTmp!��Ƭ��С & "") < 10000 Then
            strTempFile = sys.Readlob(100, 16, Val(picInfo.Item(intIndex).Tag))
            If strTempFile <> "" Then
                imgUser(intIndex).Picture = LoadPicture(strTempFile)
                  'ɾ������ʱ�ļ�
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
    '�ؼ���ʼ��,��ս�����Ϣ
    Dim objTmp As Object
    Dim i As Integer
On Error GoTo errH
    lblName.Item(intIndex).Caption = ""
    lblSex.Item(intIndex).Caption = ""
    lblZw.Item(intIndex).Caption = ""
    picInfo.Item(intIndex).Tag = ""
    imgUser(intIndex).Picture = imgUserTmp.Picture
    cboLevel(intIndex).Text = ""
    '���ͼƬ��ѡ����
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
'���ز���Ա��������
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select B.ID,B.����,B.���� " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", ",A.ȱʡ") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", "������Ա A, ") & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", " And a.����id = B.Id And A.��ԱID = [1] ") & vbNewLine & _
            "  And C.�������� = '�ٴ�' And C.������� <> 0  And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"
    On Error GoTo errH
    cboDept.Clear
    '���в���
    If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 Then
        cboDept.AddItem "���в���"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '����ȱʡ
        If InStr(";" & mstrPrivs & ";", ";���в���;") = 0 Then
            If rsTmp!ȱʡ = 1 Then
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

Private Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    'û��ǿ�������ٴ�,����ҽ��������
    If blnNew Then
        strSql = "Select 1 as ���,����ID From ������Ա Where ��ԱID=[1] Union" & _
                " Select Distinct 2 as ���,B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
                " Where A.����ID=B.����ID And A.��ԱID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", UserInfo.ID)
    End If
    If bln���� = False Then
        rsTmp.Filter = "��� = 1"
    Else
        rsTmp.Filter = ""
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadPrss() As Recordset
'���ܣ����ؿ���Ȩ���û�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long, y As Long
    Dim clsRs As New clsRecordset
    
    If cboDept.ListIndex = -1 Then Exit Function
    strSql = "Select a.Id, c.����id, a.����, a.���, a.�Ա�, a.����, b.���� As ��������, b.���� As ���ű���, a.רҵ����ְ��,nvl(a.����Ȩ��־,0) as ����Ȩ��־ ,nvl(a.��������ҽ��Ȩ��,'0000') as ��������ҽ��Ȩ��,nvl(a.סԺ����ҽ��Ȩ��,'0000') as סԺ����ҽ��Ȩ��,a.����ְ��," & vbNewLine & _
                "       Decode(d.����, 1, '������ʹ��', 2, '����ʹ��', 3, '����ʹ��', 'null') As סԺ����ҩ��Ȩ��," & vbNewLine & _
                "       Decode(e.����, 1, '������ʹ��', 2, '����ʹ��', 3, '����ʹ��', 'null') As ���￹��ҩ��Ȩ��, nvl(a.�����ȼ�,'null') as �����ȼ�,dbms_lob.getlength(G.��Ƭ) as ��Ƭ��С," & vbNewLine & _
                "nvl(Substr(A.��������ҽ��Ȩ��, 1, 1),'0') as ���ﶾ��,nvl(Substr(A.��������ҽ��Ȩ��, 2, 1),'0') as ��������,nvl(Substr(A.��������ҽ��Ȩ��, 3, 1),'0') as ���ﾫ��,nvl(Substr(A.��������ҽ��Ȩ��, 4, 1),'0') as �������," & vbNewLine & _
                "nvl(Substr(A.סԺ����ҽ��Ȩ��, 1, 1),'0') as סԺ����,nvl(Substr(A.סԺ����ҽ��Ȩ��, 2, 1),'0') as סԺ����,nvl(Substr(A.סԺ����ҽ��Ȩ��, 3, 1),'0') as סԺ����,nvl(Substr(A.סԺ����ҽ��Ȩ��, 4, 1),'0') as סԺ����" & vbNewLine & _
                "From ��Ա�� A, ���ű� B, ������Ա C, ��Ա����ҩ��Ȩ�� D, ��Ա����ҩ��Ȩ�� E, ��Ա����˵�� F,��Ա��Ƭ G" & vbNewLine & _
                "Where a.Id = c.��Աid And c.����id = b.Id And d.��Աid(+) = a.Id And a.Id = f.��Աid  And A.id=G.��Աid(+) And (d.��¼״̬ = 1 Or d.��¼״̬ Is Null) And d.����(+) = 1  And" & vbNewLine & _
                "      e.��Աid(+) = a.Id And (e.��¼״̬ = 1 Or e.��¼״̬ Is Null) And e.����(+) = 2 And" & vbNewLine & _
                "      b.Id In (Select ID From ���ű� Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) And" & vbNewLine & _
                "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  And f.��Ա���� = 'ҽ��'"
    On Error GoTo errH
    '�ж��Ƿ������в���Ȩ��
    If cboDept.ItemData(cboDept.ListIndex) = -1 Then
        If InStr(mstrPrivs, ";���в���;") = 0 Then
            strSql = strSql & " And Instr([1],','|| B.ID || ',')>0 Order By a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "," & mstrUserDept & ",")
        Else
            strSql = strSql & " And C.ȱʡ = 1 Order By a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        End If
    Else
        strSql = strSql & " And c.����id=[1]  Order By a.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(cboDept.ItemData(cboDept.ListIndex)))
    End If
    
    Set LoadPrss = clsRs.CopyNew(rsTmp)
    
    Call FilterSQL
    LoadPrss.Filter = mstrFdsql
    
    '����״̬������
    If mlngFindNum = 1 Then
        stbThis.Panels(2).Text = "��ǰ����һ����" & LoadPrss.RecordCount & "��ҽ��"
    Else
        stbThis.Panels(2).Text = "��ǰ����һ����" & LoadPrss.RecordCount & "��ҽ��"
    End If
    If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(�����ù���)"
    
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


'�򿪻�رչ���ҳǩ
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
        '�򿪹���ҳ��
        Call setPicMove(Index, decode(Index, imgType����Ȩ, Filter����Ȩ, imgType����ҩ��, Filter����ҩ��, imgType�����ȼ�, Filter�����ȼ�, Filter����ҽ��))
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
        '����Ȩ
        Case Filter����Ȩ
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
        '������סԺ
        Case Filter����ҽ��
            If imgSentence(Index).Tag = "" Then
                ChkFdmz(0).Value = 1: ChkFdmz(1).Value = 1: ChkFdZy(0).Value = 1: ChkFdZy(1).Value = 1
            Else
                ChkFdmz(0).Value = IIf(Split(imgSentence(Index).Tag, "|")(0) = 1, 1, 0)
                ChkFdmz(1).Value = IIf(Split(imgSentence(Index).Tag, "|")(1) = 1, 1, 0)
                ChkFdZy(0).Value = IIf(Split(imgSentence(Index).Tag, "|")(2) = 1, 1, 0)
                ChkFdZy(1).Value = IIf(Split(imgSentence(Index).Tag, "|")(3) = 1, 1, 0)
            End If
            ChkFdmz(0).SetFocus
        '�����ȼ�
        Case Filter�����ȼ�
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
        '����ҩ��Ȩ��
        Case Filter����ҩ��
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


'���ù��˰�ťͼ��,�����������
Private Sub setUseIco(Index As Integer)
    Dim i As Integer
    Dim blnChk As Boolean
    Dim strTmp As String
    
    On Error GoTo errH
    
    blnChk = True
    Select Case Index
        '����Ȩ
        Case imgType����Ȩ
            blnChk = ChkFd(0).Value And ChkFd(1).Value
            strTmp = "|" & IIf(ChkFd(0).Value, "1", "0") & "|" & IIf(ChkFd(1).Value, "1", "0")
        '����ҩ��Ȩ��
        Case imgType����ҩ��
            For i = 0 To ChkFdKSSmz.Count - 1
                blnChk = blnChk And ChkFdKSSmz(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdKSSmz(i).Value, "1", "0")
            Next
            For i = 0 To ChkFdKSSzy.Count - 1
                blnChk = blnChk And ChkFdKSSzy(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdKSSzy(i).Value, "1", "0")
            Next
        '�����ȼ�
        Case imgType�����ȼ�
            For i = 0 To ChkFdSS.Count - 1
                blnChk = blnChk And ChkFdSS(i).Value
                strTmp = strTmp & "|" & IIf(ChkFdSS(i).Value, "1", "0")
            Next
        '����ҽ��Ȩ��
        Case imgType����, imgType����, imgType����, imgType����
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
        '����״̬������
        If mlngFindNum = 1 Then
            stbThis.Panels(2).Text = "��ǰ����һ����" & mrsTmp.RecordCount & "��ҽ��"
        Else
            stbThis.Panels(2).Text = "��ǰ����һ����" & mrsTmp.RecordCount & "��ҽ��"
        End If
        If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(�����ù���)"
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


'�б������Ч��
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



'��ѡȨ�޸�ѡ��,��ִ��
Private Sub PicCheck(obj As Image)
    Dim strSql As String
    Dim intIndex As Integer
    Dim intType As Integer  'Ȩ������
    Dim strCheck As String
    Dim int���� As Integer  '=0 ȫ��, =1 סԺ,=2 ����
    Dim curDate As Date
    Dim i As Integer
    Dim objTmp As Object
    
    On Error GoTo errH
    intIndex = obj.Index
    DoEvents
    Select Case obj.Name
        '����Ȩ��ѡ
        Case "Pic����Ȩ"
            If obj.Tag = "True" Then
                obj.Picture = img16.ListImages.Item("None").Picture
                obj.Tag = "False"
            ElseIf obj.Tag = "False" Or obj.Tag = "" Then
                obj.Picture = img16.ListImages.Item("Black2").Picture
                obj.Tag = "True"
            End If
            intType = Priv����Ȩ
            strCheck = Val(IIf(obj.Tag = "True", 1, 0))
        '���￹��ҩ��Ȩ�޹�ѡ
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
            int���� = 2
            intType = Priv���￹��
            strCheck = IIf(obj.Tag = "True", IIf(obj.Name = "PicMZFxz", 1, IIf(obj.Name = "PicMZXz", 2, 3)), "")
        'סԺ����ҩ��Ȩ�޹�ѡ
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
            int���� = 1
            intType = PrivסԺ����
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
            intType = Priv�����ȼ�
            
            If obj.Tag = "True" Then
                strCheck = decode(Val(Mid(obj.Name, Len(obj.Name), 1)), 1, "һ��", 2, "����", 3, "����", 4, "�ļ�")
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
            intType = Priv��������
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
            intType = PrivסԺ����
            For i = 1 To 4
                Set objTmp = mcolCtl(decode(i, 1, "PicZYDl", 2, "PicZYMz", 3, "PicZYJs", 4, "PicZYGz") & intIndex)
                strCheck = strCheck & IIf(objTmp.Tag = "True", "1", "0")
            Next
    End Select
    
    If intType <> 0 Then
        If intType = Priv���￹�� Or intType = PrivסԺ���� Then
            curDate = sys.Currentdate
            strSql = "Zl_ҽ��Ȩ��_Update(" & Val(picInfo(intIndex).Tag) & "," & intType & ",'" & strCheck & "'," & int���� & ",'" & UserInfo.���� & "',to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Else
            strSql = "Zl_ҽ��Ȩ��_Update(" & Val(picInfo(intIndex).Tag) & "," & intType & ",'" & strCheck & "')"
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
    '���汾������
    Dim lngDw As Long
    Dim strFilter As String
    Dim strTmp As String
On Error GoTo errH
    '�����¼����
    strFilter = IIf(mrsTmp.Filter <> 0, mrsTmp.Filter, "")
    lngDw = mrsTmp.AbsolutePosition
    
    strTmp = decode(intType, Priv����Ȩ, "����Ȩ��־", Priv���￹��, "���￹��ҩ��Ȩ��", PrivסԺ����, "סԺ����ҩ��Ȩ��", Priv�����ȼ�, "�����ȼ�", Priv��������, "��������ҽ��Ȩ��", PrivסԺ����, "סԺ����ҽ��Ȩ��", Privҽ���ȼ�, "רҵ����ְ��")
    If intType = PrivסԺ���� Or intType = Priv���￹�� Then
        strCheck = decode(Val(strCheck), 1, "������ʹ��", 2, "����ʹ��", 3, "����ʹ��", "null")
    End If
    If strCheck = "" Then
        If intType = Priv�����ȼ� Then
            strCheck = "null"
        End If
    End If
    mrsTmp.Filter = "Id=" & uid
    If Not mrsTmp.EOF Then
        mrsTmp.Update strTmp, strCheck
        If intType = Priv�������� Then
            mrsTmp.Update "���ﶾ��", Val(Mid(strCheck, 1, 1))
            mrsTmp.Update "��������", Val(Mid(strCheck, 2, 1))
            mrsTmp.Update "���ﾫ��", Val(Mid(strCheck, 3, 1))
            mrsTmp.Update "�������", Val(Mid(strCheck, 4, 1))
        ElseIf intType = PrivסԺ���� Then
            mrsTmp.Update "סԺ����", Val(Mid(strCheck, 1, 1))
            mrsTmp.Update "סԺ����", Val(Mid(strCheck, 2, 1))
            mrsTmp.Update "סԺ����", Val(Mid(strCheck, 3, 1))
            mrsTmp.Update "סԺ����", Val(Mid(strCheck, 4, 1))
        End If
    End If
    '�ָ���¼��
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
    
    '����״̬������
    If mlngFindNum = 1 Then
        stbThis.Panels(2).Text = "��ǰ����һ����" & mrsTmp.RecordCount & "��ҽ��"
    Else
        stbThis.Panels(2).Text = "��ǰ����һ����" & mrsTmp.RecordCount & "��ҽ��"
    End If
    If mlngFilter = 1 Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & "(�����ù���)"
    
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
    Call CloseFilter(Filter����ҩ��)
End Sub

Private Sub ChkFdKSSmz_LostFocus(Index As Integer)
    Call CloseFilter(Filter����ҩ��)
End Sub

Private Sub ChkFd_LostFocus(Index As Integer)
    Call CloseFilter(Filter����Ȩ)
End Sub

Private Sub ChkFdmz_LostFocus(Index As Integer)
    Call CloseFilter(Filter����ҽ��)
End Sub

Private Sub ChkFdZy_LostFocus(Index As Integer)
    Call CloseFilter(Filter����ҽ��)
End Sub

Private Sub ChkFdSS_LostFocus(Index As Integer)
    Call CloseFilter(Filter�����ȼ�)
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

Private Sub Pic����Ȩ_Click(Index As Integer)
    PicCheck Pic����Ȩ(Index)
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

