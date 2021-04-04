VERSION 5.00
Begin VB.UserControl ctlKeyBoard 
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ScaleHeight     =   6765
   ScaleWidth      =   6015
   Begin VB.PictureBox picNumeric 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   135
      ScaleHeight     =   3090
      ScaleWidth      =   4005
      TabIndex        =   39
      Top             =   3480
      Width           =   4005
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   38
         Left            =   45
         TabIndex        =   40
         Top             =   2340
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "0"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   39
         Left            =   45
         TabIndex        =   41
         Top             =   1560
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "1"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   40
         Left            =   1380
         TabIndex        =   42
         Top             =   1560
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "2"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   41
         Left            =   2700
         TabIndex        =   43
         Top             =   1560
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "3"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   42
         Left            =   45
         TabIndex        =   44
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "4"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   43
         Left            =   1380
         TabIndex        =   45
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "5"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   44
         Left            =   2700
         TabIndex        =   46
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "6"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   45
         Left            =   45
         TabIndex        =   47
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "7"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   46
         Left            =   1380
         TabIndex        =   48
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "8"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   47
         Left            =   2700
         TabIndex        =   49
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "9"
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   48
         Left            =   1380
         TabIndex        =   50
         Top             =   2340
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "确定"
         BackColor       =   16777215
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   720
         Index           =   49
         Left            =   2700
         TabIndex        =   51
         Top             =   2340
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1270
         Caption         =   "清除"
         BackColor       =   16777215
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
   End
   Begin VB.PictureBox picKey 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   165
      ScaleHeight     =   3090
      ScaleWidth      =   5310
      TabIndex        =   0
      Top             =   165
      Width           =   5310
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   9
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  0 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  1 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   1
         Left            =   45
         TabIndex        =   3
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  2 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   2
         Left            =   45
         TabIndex        =   4
         Top             =   1845
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  3 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   3
         Left            =   45
         TabIndex        =   5
         Top             =   2460
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  4 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   4
         Left            =   705
         TabIndex        =   6
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  5 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   5
         Left            =   705
         TabIndex        =   7
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  6 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   6
         Left            =   705
         TabIndex        =   8
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  7 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   7
         Left            =   705
         TabIndex        =   9
         Top             =   1845
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  8 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   8
         Left            =   705
         TabIndex        =   10
         Top             =   2460
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  9 "
         BackColor       =   16777215
         ForeColor       =   255
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   10
         Left            =   1365
         TabIndex        =   11
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  A "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   11
         Left            =   1365
         TabIndex        =   12
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  B "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   12
         Left            =   1365
         TabIndex        =   13
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  C "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   13
         Left            =   1365
         TabIndex        =   14
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  D "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   14
         Left            =   1365
         TabIndex        =   15
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  E "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   15
         Left            =   2025
         TabIndex        =   16
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  F "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   16
         Left            =   2025
         TabIndex        =   17
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  G "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   17
         Left            =   2025
         TabIndex        =   18
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  H "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   18
         Left            =   2025
         TabIndex        =   19
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  I "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   19
         Left            =   2025
         TabIndex        =   20
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  J "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   20
         Left            =   2685
         TabIndex        =   21
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1005
         Caption         =   "   确定 "
         BackColor       =   16777215
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   21
         Left            =   2685
         TabIndex        =   22
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  K "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   22
         Left            =   2685
         TabIndex        =   23
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  L "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   23
         Left            =   2685
         TabIndex        =   24
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  M "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   24
         Left            =   2685
         TabIndex        =   25
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  N "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   25
         Left            =   4005
         TabIndex        =   26
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1005
         Caption         =   "   清除 "
         BackColor       =   16777215
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   26
         Left            =   3345
         TabIndex        =   27
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  O "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   27
         Left            =   3345
         TabIndex        =   28
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  P "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   28
         Left            =   3345
         TabIndex        =   29
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  Q "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   29
         Left            =   3345
         TabIndex        =   30
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  R "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   30
         Left            =   4005
         TabIndex        =   31
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  S "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   31
         Left            =   4005
         TabIndex        =   32
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  T "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   32
         Left            =   4005
         TabIndex        =   33
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  U "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   33
         Left            =   4005
         TabIndex        =   34
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  V "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   34
         Left            =   4665
         TabIndex        =   35
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  W "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   35
         Left            =   4665
         TabIndex        =   36
         Top             =   1245
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  X "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   36
         Left            =   4665
         TabIndex        =   37
         Top             =   1860
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  Y "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   37
         Left            =   4665
         TabIndex        =   38
         Top             =   2475
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1005
         Caption         =   "  Z "
         BackColor       =   16777215
         FontSize        =   10.5
         FontBold        =   -1  'True
         AutoSize        =   0   'False
         ButtonHeight    =   450
         TextAligment    =   0
      End
   End
End
Attribute VB_Name = "ctlKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CommandClick(Caption As String)

Private mbytKeyMode As Byte

Public Property Let KeyMode(ByVal bytData As Byte)

    mbytKeyMode = bytData
    
    If mbytKeyMode = 1 Then
        picNumeric.Visible = False
        picKey.Move 0, 0, picKey.Width, picKey.Height
        picKey.Visible = True
    Else
        picNumeric.Move 0, 0, picNumeric.Width, picNumeric.Height
        picKey.Visible = False
        picNumeric.Visible = True
    End If
    
'    UserControl.Refresh
    Call UserControl_Resize
    
End Property

Private Sub picKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picKey_Paint()
    Call RaisEffect(picKey, -1)
    Call DrawColorToColor(picKey, picKey.BackColor, &HFFC0C0)
End Sub

Private Sub picNumeric_Paint()
    Call RaisEffect(picNumeric, -1)
    Call DrawColorToColor(picNumeric, picNumeric.BackColor, &HFFC0C0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    If mbytKeyMode = 1 Then
        UserControl.Width = picKey.Width
        UserControl.Height = picKey.Height
    Else
        UserControl.Width = picNumeric.Width
        UserControl.Height = picNumeric.Height
    End If
    
End Sub

Private Sub UserControl_Show()
    Dim intLoop As Integer
    
    For intLoop = 0 To 49
        UsrCmd(intLoop).ShowPicture = False
    Next
    
'    mbytKeyMode = 2
'    picKey.Visible = False
'    picNumeric.Top = 0
'    picNumeric.Visible = True

End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    RaiseEvent CommandClick(Trim(UsrCmd(Index).Caption))
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
