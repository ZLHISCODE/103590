VERSION 5.00
Begin VB.Form frmScreenKeyboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "屏幕键盘"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frmScreenKeyboard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraOp 
      Height          =   2260
      Left            =   5460
      TabIndex        =   50
      Top             =   -90
      Width           =   960
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "→"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   39
         Left            =   475
         MouseIcon       =   "frmScreenKeyboard.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "←"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   37
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":1194
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  End"
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   35
         Left            =   475
         MouseIcon       =   "frmScreenKeyboard.frx":1A5E
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Home"
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   36
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":2328
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   111
         Left            =   475
         MouseIcon       =   "frmScreenKeyboard.frx":2BF2
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   106
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":34BC
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   109
         Left            =   475
         MouseIcon       =   "frmScreenKeyboard.frx":3D86
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   107
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":4650
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ins"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   45
         Left            =   475
         MouseIcon       =   "frmScreenKeyboard.frx":4F1A
         MousePointer    =   99  'Custom
         TabIndex        =   52
         ToolTipText     =   "Insert"
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblOp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   46
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":57E4
         MousePointer    =   99  'Custom
         TabIndex        =   51
         ToolTipText     =   "Delete"
         Top             =   120
         Width           =   440
      End
   End
   Begin VB.Frame fraAlph 
      Height          =   2260
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   5465
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   110
         Left            =   4560
         MouseIcon       =   "frmScreenKeyboard.frx":60AE
         MousePointer    =   99  'Custom
         TabIndex        =   49
         ToolTipText     =   "小数点"
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   161
         Left            =   4450
         MouseIcon       =   "frmScreenKeyboard.frx":6978
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   1800
         Width           =   965
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   191
         Left            =   4000
         MouseIcon       =   "frmScreenKeyboard.frx":7242
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   """"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   1
         Left            =   3550
         MouseIcon       =   "frmScreenKeyboard.frx":7B0C
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   2
         Left            =   3100
         MouseIcon       =   "frmScreenKeyboard.frx":83D6
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "空格"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   32
         Left            =   1300
         MouseIcon       =   "frmScreenKeyboard.frx":8CA0
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   1800
         Width           =   1790
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   222
         Left            =   850
         MouseIcon       =   "frmScreenKeyboard.frx":956A
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   186
         Left            =   400
         MouseIcon       =   "frmScreenKeyboard.frx":9E34
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1800
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   190
         Left            =   4110
         MouseIcon       =   "frmScreenKeyboard.frx":A6FE
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ","
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   188
         Left            =   3660
         MouseIcon       =   "frmScreenKeyboard.frx":AFC8
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   77
         Left            =   3210
         MouseIcon       =   "frmScreenKeyboard.frx":B892
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   78
         Left            =   2760
         MouseIcon       =   "frmScreenKeyboard.frx":C15C
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   66
         Left            =   2310
         MouseIcon       =   "frmScreenKeyboard.frx":CA26
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   86
         Left            =   1860
         MouseIcon       =   "frmScreenKeyboard.frx":D2F0
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   67
         Left            =   1410
         MouseIcon       =   "frmScreenKeyboard.frx":DBBA
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   88
         Left            =   960
         MouseIcon       =   "frmScreenKeyboard.frx":E484
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   90
         Left            =   510
         MouseIcon       =   "frmScreenKeyboard.frx":ED4E
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   1380
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   13
         Left            =   4320
         MouseIcon       =   "frmScreenKeyboard.frx":F618
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   960
         Width           =   1100
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   76
         Left            =   3870
         MouseIcon       =   "frmScreenKeyboard.frx":FEE2
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   75
         Left            =   3420
         MouseIcon       =   "frmScreenKeyboard.frx":107AC
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   74
         Left            =   2970
         MouseIcon       =   "frmScreenKeyboard.frx":11076
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   72
         Left            =   2520
         MouseIcon       =   "frmScreenKeyboard.frx":11940
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   71
         Left            =   2070
         MouseIcon       =   "frmScreenKeyboard.frx":1220A
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   70
         Left            =   1620
         MouseIcon       =   "frmScreenKeyboard.frx":12AD4
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   68
         Left            =   1170
         MouseIcon       =   "frmScreenKeyboard.frx":1339E
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   83
         Left            =   720
         MouseIcon       =   "frmScreenKeyboard.frx":13C68
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   65
         Left            =   270
         MouseIcon       =   "frmScreenKeyboard.frx":14532
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   960
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   80
         Left            =   4245
         MouseIcon       =   "frmScreenKeyboard.frx":14DFC
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   79
         Left            =   3795
         MouseIcon       =   "frmScreenKeyboard.frx":156C6
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   73
         Left            =   3345
         MouseIcon       =   "frmScreenKeyboard.frx":15F90
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   85
         Left            =   2895
         MouseIcon       =   "frmScreenKeyboard.frx":1685A
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   89
         Left            =   2445
         MouseIcon       =   "frmScreenKeyboard.frx":17124
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   84
         Left            =   1995
         MouseIcon       =   "frmScreenKeyboard.frx":179EE
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   82
         Left            =   1545
         MouseIcon       =   "frmScreenKeyboard.frx":182B8
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   69
         Left            =   1095
         MouseIcon       =   "frmScreenKeyboard.frx":18B82
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   87
         Left            =   645
         MouseIcon       =   "frmScreenKeyboard.frx":1944C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   81
         Left            =   195
         MouseIcon       =   "frmScreenKeyboard.frx":19D16
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   540
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BS ←"
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   8
         Left            =   4985
         MouseIcon       =   "frmScreenKeyboard.frx":1A5E0
         MousePointer    =   99  'Custom
         TabIndex        =   12
         ToolTipText     =   "BackSpace 回格"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   96
         Left            =   4535
         MouseIcon       =   "frmScreenKeyboard.frx":1AEAA
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   105
         Left            =   4085
         MouseIcon       =   "frmScreenKeyboard.frx":1B774
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   104
         Left            =   3635
         MouseIcon       =   "frmScreenKeyboard.frx":1C03E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   103
         Left            =   3185
         MouseIcon       =   "frmScreenKeyboard.frx":1C908
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   102
         Left            =   2735
         MouseIcon       =   "frmScreenKeyboard.frx":1D1D2
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   101
         Left            =   2285
         MouseIcon       =   "frmScreenKeyboard.frx":1DA9C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   100
         Left            =   1835
         MouseIcon       =   "frmScreenKeyboard.frx":1E366
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   99
         Left            =   1375
         MouseIcon       =   "frmScreenKeyboard.frx":1EC30
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   98
         Left            =   935
         MouseIcon       =   "frmScreenKeyboard.frx":1F4FA
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   97
         Left            =   485
         MouseIcon       =   "frmScreenKeyboard.frx":1FDC4
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   120
         Width           =   440
      End
      Begin VB.Label lblAlph 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   27
         Left            =   35
         MouseIcon       =   "frmScreenKeyboard.frx":2068E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         ToolTipText     =   "退出"
         Top             =   120
         Width           =   440
      End
   End
End
Attribute VB_Name = "frmScreenKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 转移输入焦点的声明
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' 窗口置顶的声明
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' 模拟按键声明
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

' 禁止本窗体拥有输入焦点的常数
Private Const HWND_NOTOPMOST = -2
Private Const WS_DISABLED = &H8000000
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

' 窗口置顶的常数
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
' 模拟按钮常数
Private Const KEYEVENTF_KEYUP = &H2

Private Const DEF_PRESSDOWN = &HFF00&
Private Const DEF_BK = &H80000004

Private Sub Form_Initialize()
    If App.PrevInstance Then End '必须用 End 才有用，保证只生成一个实例
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height - 1000
End Sub

Private Sub Form_Load()
'功能：设置窗口位置将，窗口置顶
    Me.Left = Screen.Width - Me.ScaleWidth - 120
    Me.Top = Screen.Height - Me.ScaleHeight - 1400
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_DISABLED
End Sub

' 鼠标移动到窗体上时，窗体置顶
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lblOp_Click(Index As Integer)
    keybd_event Index, 0, 0, 0
    keybd_event Index, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub lblAlph_Click(Index As Integer)
'在输入标点的时候有组合按键 冒号 = 2 双引号 = 1 问号 = 191

    If Index = 1 Then
        keybd_event 161, 0, 0, 0
        keybd_event 222, 0, 0, 0
        keybd_event 222, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 2 Then
        keybd_event 161, 0, 0, 0
        keybd_event 186, 0, 0, 0
        keybd_event 186, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 191 Then
        keybd_event 161, 0, 0, 0
        keybd_event 191, 0, 0, 0
        keybd_event 191, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    Else
        keybd_event Index, 0, 0, 0
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub

Private Sub lblOp_DblClick(Index As Integer)
    keybd_event Index, 0, 0, 0
    keybd_event Index, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub lblAlph_DblClick(Index As Integer)
    If Index = 1 Then
        keybd_event 161, 0, 0, 0
        keybd_event 222, 0, 0, 0
        keybd_event 222, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 2 Then
        keybd_event 161, 0, 0, 0
        keybd_event 186, 0, 0, 0
        keybd_event 186, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 191 Then
        keybd_event 161, 0, 0, 0
        keybd_event 191, 0, 0, 0
        keybd_event 191, 0, KEYEVENTF_KEYUP, 0
        keybd_event 161, 0, KEYEVENTF_KEYUP, 0
    Else
        keybd_event Index, 0, 0, 0
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub

Private Sub lblOp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOp(Index).BackColor = DEF_PRESSDOWN
End Sub

Private Sub lblOp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOp(Index).BackColor = DEF_BK
End Sub

Private Sub lblAlph_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAlph(Index).BackColor = DEF_PRESSDOWN
End Sub

Private Sub lblAlph_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAlph(Index).BackColor = DEF_BK
End Sub

