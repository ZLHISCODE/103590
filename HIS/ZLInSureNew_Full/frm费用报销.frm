VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm费用报销 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "费用报销"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frm费用报销.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6390
      TabIndex        =   63
      Top             =   1410
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6390
      TabIndex        =   62
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmd分档明细 
      Caption         =   "分档明细"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6390
      TabIndex        =   64
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmd虚拟结算 
      Caption         =   "预结算(&Y)"
      Height          =   350
      Left            =   6390
      TabIndex        =   61
      Top             =   480
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "门诊"
      Height          =   5295
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5955
      Begin VB.ComboBox cbo就医 
         Height          =   300
         Index           =   0
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txt帐户支付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   29
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txt统筹支付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   27
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txt首先自付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   25
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt全自付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   23
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt进入统筹 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   21
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   19
         Top             =   4020
         Width           =   1365
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   1995
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   1920
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3519
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt余额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   16
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   5190
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   750
         Width           =   255
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Index           =   0
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txt中心 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt状态 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   3840
         TabIndex        =   14
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1605
      End
      Begin VB.CommandButton cmd卡号 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt卡号 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lbl就医 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就医(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl帐户支付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户支付(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   28
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lbl统筹支付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   26
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lbl首先自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "首先自付(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   24
         Top             =   4470
         Width           =   990
      End
      Begin VB.Label lbl全自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "全自付(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   22
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label lbl进入统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "进入统筹(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   20
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl费用总额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "余额(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   15
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl病种 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病种(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl中心 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中心(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl状态 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "状态(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   13
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "住院"
      Height          =   5295
      Index           =   1
      Left            =   180
      TabIndex        =   31
      Top             =   120
      Width           =   5955
      Begin VB.ComboBox cbo就医 
         Height          =   300
         Index           =   1
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txt起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   43
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txt住院次数 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   39
         Top             =   720
         Width           =   1635
      End
      Begin VB.CommandButton cmd卡号 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   37
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txt状态 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   45
         Top             =   1500
         Width           =   1605
      End
      Begin VB.TextBox txt中心 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   41
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt余额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   47
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   50
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txt进入统筹 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   52
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txt全自付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   54
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt首先自付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   56
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt统筹支付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   58
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txt统筹自付 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   60
         Top             =   4800
         Width           =   1365
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   1995
         Index           =   1
         Left            =   450
         TabIndex        =   48
         Top             =   1920
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3519
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt卡号 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   32
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lbl起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "起付线(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2940
         TabIndex        =   42
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lbl就医 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就医(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   34
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "次数(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   38
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   30
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   36
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl状态 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "状态(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   44
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl中心 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中心(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   40
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "余额(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   46
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl费用总额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   49
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl进入统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "进入统筹(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   51
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl全自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "全自付(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   53
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label lbl首先自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "首先自付(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   55
         Top             =   4470
         Width           =   990
      End
      Begin VB.Label lbl统筹支付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   57
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lbl统筹自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹自付(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   59
         Top             =   4860
         Width           =   990
      End
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "劳司职工"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   390
      Left            =   6180
      TabIndex        =   65
      Top             =   3150
      Width           =   1440
   End
End
Attribute VB_Name = "frm费用报销"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FrameIndex
    int门诊 = 0
    int住院 = 1
End Enum
Private blnModify As Boolean                '是否修改，修改则禁止确定按钮
Private mint编辑模式 As Integer             '1-报销;2-查阅
Private mlng记录ID As Long
Private mbln住院 As Boolean                 '门诊/住院
Private mblnOK As Boolean
Private Const strFormat_金额 As String = "#####0.00;-#####0.00; ;"
Private Const strFormat_数量 As String = "#####0.000;-#####0.000; ;"

Public Function ShowME(int编辑模式 As Integer, Optional bln住院 As Boolean = False, Optional ByVal lng记录ID As Long = 0) As Boolean
    On Error Resume Next
    mint编辑模式 = int编辑模式
    mlng记录ID = lng记录ID
    mbln住院 = bln住院
    mblnOK = False
    blnModify = False
    
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub Bill_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    With Bill(Index)
        If Col = 0 Then Exit Sub
        
        If Row = .Rows - 1 Then
            .ColData(Col) = 0
        Else
            .ColData(Col) = IIf(Col = .Cols - 1, 5, 4)
        End If
    End With
End Sub

Private Sub Bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim StrInput As String
    Dim lngRow As Long
    Dim curMoney As Currency, curCount As Currency
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
    
    With Bill(Index)
        If .TxtVisible = False Then
            .Text = .TextMatrix(.Row, .Col)
            If .Text = "" Then .Text = " "
        Else
            StrInput = Format(.Text, IIf(.Col = 2 And mbln住院, strFormat_数量, strFormat_金额))
            
            If Trim(StrInput) = "" Then
                StrInput = " "
            Else
                If Not IsNumeric(StrInput) Then
                    MsgBox "输入中含有非法字符！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .RowData(.Row) = 1 And .Col = 2 Then
                    If Not (Val(StrInput) >= 0 And Val(StrInput) < 100) Then
                        MsgBox "实际报销比例不能小于零或大于等于100%！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            End If
            .Text = StrInput
        End If
        .TextMatrix(.Row, .Col) = .Text
        
        curMoney = 0: curCount = 0
        For lngRow = 1 To .Rows - 2
            curMoney = curMoney + Val(.TextMatrix(lngRow, 1))
            If mbln住院 Then
                If .RowData(lngRow) = 0 Then
                    curCount = curCount + Val(.TextMatrix(lngRow, 2))
                Else
                    '表示保存的是比例
                End If
            Else
                curCount = curCount + Val(.TextMatrix(lngRow, 2))
            End If
        Next
        .TextMatrix(.Rows - 1, 1) = Format(curMoney, strFormat_金额)
        .TextMatrix(.Rows - 1, 2) = Format(curCount, IIf(mbln住院, strFormat_数量, strFormat_金额))
    End With
End Sub

Private Sub cbo就医_Click(Index As Integer)
    Dim intMax As Integer   '保存最后一次住院次数
    Dim rsTemp As New ADODB.Recordset
    '获取指定就医方式的起付线标准
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
    
    If Index <> int住院 Then Exit Sub
    If Val(txt中心(int住院).Tag) = 0 Then Exit Sub

    '再读出帐户年度信息
    If Val(txt卡号(Index).Tag) = 0 Then Exit Sub
    gstrSQL = "select * from 帐户年度信息 where 险类=" & TYPE_四川眉山 & _
        " and 病人ID=" & Val(txt卡号(Index).Tag) & " and 年度=" & Format(zlDatabase.Currentdate, "yyyy")
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.EOF = False Then
        '设置帐户情况
        If cbo就医(Index).ListIndex = 0 Then
            txt住院次数(Index).Text = Nvl(rsTemp("住院次数累计"), 0) + IIf(mint编辑模式 = 1, 1, 0)
        Else
            txt住院次数(Index).Text = Nvl(rsTemp("外院住院次数"), 0) + IIf(mint编辑模式 = 1, 1, 0)
        End If
    Else
        txt住院次数(Index).Text = 1
    End If
    
    gstrSQL = "select 金额 from 保险报销政策 " & _
             " Where 性质=2 And 中心=" & Val(txt中心(int住院).Tag) & " And 本院=" & cbo就医(int住院).ListIndex + 1 & _
             " And 人群=" & Val(txt状态(int住院).Tag) & " And 类别='" & Val(txt住院次数(int住院)) & "'" & _
             " And 险类=" & TYPE_四川眉山 & " And 年度=" & Format(zlDatabase.Currentdate, "yyyy")
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not rsTemp.EOF Then
        txt起付线(int住院) = Format(rsTemp!金额, strFormat_金额)
    Else
        '取最大住院次数时的起付金额做为本次住院起付线
        gstrSQL = "select Max(类别) 住院次数 from 保险报销政策 " & _
                 " Where 性质=2 And 中心=" & Val(txt中心(int住院).Tag) & " And 本院=" & cbo就医(int住院).ListIndex + 1 & _
                 " And 人群=" & Val(txt状态(int住院).Tag) & " And 类别<>'A'" & _
                 " And 险类=" & TYPE_四川眉山 & " And 年度=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        intMax = rsTemp!住院次数
        
        gstrSQL = "select 金额 from 保险报销政策 " & _
                 " Where 性质=2 And 中心=" & Val(txt中心(int住院).Tag) & " And 本院=" & cbo就医(int住院).ListIndex + 1 & _
                 " And 人群=" & Val(txt状态(int住院).Tag) & " And 类别='" & intMax & "'" & _
                 " And 险类=" & TYPE_四川眉山 & " And 年度=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        txt起付线(int住院) = Format(rsTemp!金额, strFormat_金额)
    End If
    
    gComInfo_眉山.起付线 = Val(txt起付线(int住院).Text)
End Sub

Private Sub cmd病种_Click(Index As Integer)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_四川眉山
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种(Index).Text)
    If rsTemp.State <> 1 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt病种(Index).Text = rsTemp("名称")
        txt病种(Index).Tag = rsTemp("ID")
        zlControl.TxtSelAll txt病种(Index)
    End If
    txt病种(Index).SetFocus
    
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
End Sub

Private Sub cmd分档明细_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If mint编辑模式 = 1 Then
        Call frm费用报销_分档明细.ShowME
        txt统筹支付(int住院).Text = Format(gComInfo_眉山.统筹支付, strFormat_金额)
        txt统筹自付(int住院).Text = Format(gComInfo_眉山.统筹自付, strFormat_金额)
    Else
        '读取分档报销数据
        gstrSQL = " Select B.名称,A.比例,A.档次,A.进入统筹金额 进入统筹,A.统筹报销金额 统筹报销 " & _
                  " From 保险结算计算 A,(Select * From 保险费用档 Where 险类=" & TYPE_四川眉山 & " And 中心=" & Val(txt中心(int住院).Tag) & " And 档次<>0) B" & _
                  " Where A.档次=B.档次 And A.结帐ID=" & mlng记录ID
        Call OpenRecordset(rsTemp, Me.Caption)
        Call frm费用报销_分档明细.ShowME(True, rsTemp)
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    If gComInfo_眉山.费用总额 = 0 Then
        MsgBox "未输入任何数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mbln住院 Then
        If Not SaveData(cbo就医(int住院).ListIndex = 0) Then Exit Sub
    Else
        If Not SaveData(cbo就医(int门诊).ListIndex = 0) Then Exit Sub
    End If
    
    '打印票据
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605_" & IIf(mbln住院, 2, 1), Me, "险类=" & 25, "记录ID=" & mlng记录ID, 2)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd虚拟结算_Click()
    Dim cur金额 As Currency, cur统筹 As Currency, cur进入统筹 As Currency
    Dim lngRow As Long, intIndex As Integer, intCol As Integer 'intCol表示当前报销金额列
    Dim msfObj As BillEdit, sin比例 As Single, intBound As Integer
    
    '初始化
    Call Init_大类_米易
    Call Init_结构体_米易
    intIndex = IIf(mbln住院, int住院, int门诊)
    gComInfo_眉山.病人ID = Val(txt卡号(intIndex).Tag)
    gComInfo_眉山.人群 = txt状态(intIndex).Text
    If Not mbln住院 Then
        gComInfo_眉山.病种ID = Val(txt病种(intIndex).Tag)
        gComInfo_眉山.病种名称 = txt病种(intIndex).Text
    End If
    gComInfo_眉山.中心 = Val(txt中心(intIndex).Tag)
    If mbln住院 Then
        gComInfo_眉山.帐户余额 = Val(Split(txt余额(intIndex).Text, "/")(0))
    Else
        gComInfo_眉山.帐户余额 = Val(txt余额(intIndex).Text)
    End If
    If gComInfo_眉山.病人ID = 0 Then
        MsgBox "请选择医保病人，再输入各大类的汇总金额后，才可运行预结算！", vbInformation, gstrSysName
        txt卡号(intIndex).SetFocus
        Exit Sub
    End If
    Call 实际报销比例
    
    '处理大类汇总记录集
    cur金额 = 0: cur统筹 = 0: cur进入统筹 = 0
    Set msfObj = Bill(intIndex)
    For lngRow = 1 To msfObj.Rows - 2
        cur金额 = cur金额 + Val(msfObj.TextMatrix(lngRow, 1))
        rs大类_米易.MoveFirst
        rs大类_米易.Find "大类名称='" & msfObj.TextMatrix(lngRow, 0) & "'"
        If mbln住院 Then rs大类_米易!数量 = Val(msfObj.TextMatrix(lngRow, 2))
        rs大类_米易!费用总额 = Val(msfObj.TextMatrix(lngRow, 1))
        If Nvl(rs大类_米易!统筹比额, 0) <> 0 Then
            rs大类_米易!报销总额 = Val(msfObj.TextMatrix(lngRow, 1))
            cur统筹 = cur统筹 + Val(msfObj.TextMatrix(lngRow, 1))
        End If
        rs大类_米易.Update
    Next
    gComInfo_眉山.费用总额 = cur金额
    gComInfo_眉山.全自付 = cur金额 - cur统筹
    
    '计算进入统筹金额
    With rs大类_米易
        .MoveFirst
        Do While Not .EOF
            If !大类名称 = "浮动比例" Then
                '如果用户调整了比例，按用户输入的为准
                sin比例 = 100
                If InStr(1, gstr实际报销比例_米易, "|" & !大类名称 & ";") <> 0 Then
                    intBound = UBound(Split(Mid(gstr实际报销比例_米易, 2), "|"))
                    For lngRow = 0 To intBound
                        If Split(Split(Mid(gstr实际报销比例_米易, 2), "|")(lngRow), ";")(0) = !大类名称 Then
                            sin比例 = Val(Split(Split(Mid(gstr实际报销比例_米易, 2), "|")(lngRow), ";")(1))
                            Exit For
                        End If
                    Next
                End If
                !报销总额 = !报销总额 * sin比例 / 100
            Else
                If !特准定额 = 0 And !特准天数 = 0 Then
                    !报销总额 = !报销总额 * Nvl(!统筹比额, 0) / 100
                Else
                    If !数量 > !特准天数 Then
                        '如果住院日超过特准天数，那么金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                        !报销总额 = !特准定额 * !特准天数 + _
                            (!数量 - IIf(!特准定额 = 0 Or !特准天数 = 0, 0, !特准天数)) * !统筹比额
                    Else
                        If !特准定额 = 0 Or !特准天数 = 0 Then
                            '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                            '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                            !报销总额 = !报销总额 * !统筹比额 / 100
                        Else
                            !报销总额 = !数量 * !特准定额
                        End If
                    End If
                End If
            End If
            cur进入统筹 = cur进入统筹 + !报销总额
            .Update
            .MoveNext
        Loop
    End With
    gComInfo_眉山.首先自付 = cur统筹 - cur进入统筹
    gComInfo_眉山.进入统筹 = cur进入统筹
    
    Call Calc_实际报销_米易
    
    '计算统筹及帐户支付
    If Not mbln住院 Then
        Call Calc_门诊报销计算_米易(cbo就医(int门诊).ListIndex = 0)
    Else
        Call Calc_住院报销计算_米易(cbo就医(int住院).ListIndex = 0)
    End If
    
    Call 显示结算结果
    
    cmd确定.Enabled = True
    If mbln住院 Then cmd分档明细.Enabled = True
End Sub

Private Sub 显示结算结果()
    Dim intCol As Integer, intIndex As Integer, lngRow As Long
    Dim cur金额 As Currency
    Dim msfObj As BillEdit
    
    '将各大类的可报销金额写回界面，供操作员校验
    intCol = IIf(mbln住院, 3, 2)
    intIndex = IIf(mbln住院, int住院, int门诊)
    cur金额 = 0
    Set msfObj = Bill(intIndex)
    
    For lngRow = 1 To msfObj.Rows - 2
        rs大类_米易.MoveFirst
        rs大类_米易.Find "大类名称='" & msfObj.TextMatrix(lngRow, 0) & "'"
        msfObj.TextMatrix(lngRow, intCol) = Format(Nvl(rs大类_米易!报销总额, 0), strFormat_金额)
        cur金额 = cur金额 + Nvl(rs大类_米易!报销总额, 0)
    Next
    msfObj.TextMatrix(msfObj.Rows - 1, intCol) = Format(cur金额, strFormat_金额)
    
    '将计算结果写回界面，供操作员校验
    txt费用总额(intIndex).Text = Format(gComInfo_眉山.费用总额, strFormat_金额)
    txt进入统筹(intIndex).Text = Format(gComInfo_眉山.进入统筹, strFormat_金额)
    txt全自付(intIndex).Text = Format(gComInfo_眉山.全自付, strFormat_金额)
    txt首先自付(intIndex).Text = Format(gComInfo_眉山.首先自付, strFormat_金额)
    txt统筹支付(intIndex).Text = Format(gComInfo_眉山.统筹支付, strFormat_金额)
    If Not mbln住院 Then
        txt帐户支付(intIndex).Text = Format(gComInfo_眉山.帐户支付, strFormat_金额)
    Else
        txt统筹自付(intIndex).Text = Format(gComInfo_眉山.统筹自付, strFormat_金额)
    End If
End Sub

Private Sub txt病种_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt病种(Index)
End Sub

Private Sub txt病种_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种(Index).Text = ""
        txt病种(Index).Tag = ""
    End If
    
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
End Sub

Private Sub cmd卡号_Click(Index As Integer)
    gstrSQL = "Select A.病人ID as ID,A.卡号,A.医保号,'******' 密码,B.姓名,B.性别,B.出生日期,B.身份证号,A.中心 中心,C.名称 中心名称  " & _
             "  ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,E.序号 人群,E.名称 状态,A.退休证号,A.帐户余额 " & _
             "  From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D,保险人群 E " & _
             "   where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & TYPE_四川眉山 & _
             "   and A.险类=C.险类 and A.中心=C.序号 And A.在职=E.序号 and A.险类=E.险类 and A.病种ID=D.ID(+)"
    Call Get帐户情况(Index)
    
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Call InitFace
    
    gstrSQL = " Select * From 保险支付大类 " & _
              " Where 险类=" & TYPE_四川眉山 & _
              " Order by 编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Call WriteBill(rsTemp, Bill(int门诊))
    Call WriteBill(rsTemp, Bill(int住院))
    Bill(int门诊).AllowAddRow = False
    Bill(int住院).AllowAddRow = False
    
    fra(int门诊).Visible = (Not mbln住院)
    fra(int住院).Visible = (mbln住院)
    If mbln住院 Then
        Call 显示实际报销比例
        fra(int住院).ZOrder
    End If
    
    If mint编辑模式 = 2 Then
        '读取数据
        If Not ReadData Then
            MsgBox "未找到结算数据！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        '设置各控件的状态
        Dim objCon As Object
        For Each objCon In Me.Controls
            If InStr(1, "TEXTBOX;COMBOBOX", UCase(TypeName(objCon))) <> 0 Then
                objCon.Enabled = False
            ElseIf InStr(1, "CMD卡号,CMD病种", UCase(objCon.Name)) <> 0 Then
                objCon.Enabled = False
            End If
        Next
        Bill(IIf(mbln住院, int住院, int门诊)).Active = False
        cmd分档明细.Enabled = mbln住院
        cmd虚拟结算.Enabled = False
        cmd确定.Enabled = False
        cmd取消.Caption = "确定(&O)"
        cmd取消.Top = cmd确定.Top
    End If
End Sub

Private Function ReadData() As Boolean
    Dim intIndex As Integer
    Dim intCol As Integer, lngRow As Long
    Dim cur费用金额 As Currency, cur报销总额 As Currency
    Dim msfObj As BillEdit
    Dim rsTemp As New ADODB.Recordset
    '先读取医保病人的数据
    intIndex = IIf(mbln住院, int住院, int门诊)
    gstrSQL = " Select 医保号 From 保险帐户 Where 险类=" & TYPE_四川眉山 & _
              " And 病人ID = (" & _
              "        Select 病人ID From 保险结算记录 Where 记录ID=" & mlng记录ID & " And 性质=" & IIf(mbln住院, "2", "1") & " And 险类=" & TYPE_四川眉山 & ")"
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp.EOF Then Exit Function
    txt卡号(intIndex) = rsTemp!医保号
    Call txt卡号_KeyPress(intIndex, vbKeyReturn)
    
    '将计算结果写回界面，供操作员校验
    '读取结算记录信息
    gstrSQL = " Select A.*,B.名称 病种名称 From 保险结算记录 A,(Select ID 病种ID,名称 From 保险病种 Where 险类=" & TYPE_四川眉山 & ") B" & _
              " Where A.记录ID=" & mlng记录ID & " And A.性质=" & IIf(mbln住院, "2", "1") & _
              " And A.险类=" & TYPE_四川眉山 & " And A.病种ID=B.病种ID(+)"
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not mbln住院 Then
        txt病种(intIndex).Text = Nvl(rsTemp!病种名称, "")
        txt病种(intIndex).Tag = Nvl(rsTemp!病种ID, 0)
    Else
        txt起付线(int住院).Text = Nvl(rsTemp!起付线, 0)
        txt住院次数(int住院).Text = Nvl(rsTemp!住院次数, 1)
    End If
    txt费用总额(intIndex).Text = Format(rsTemp!发生费用金额, strFormat_金额)
    txt进入统筹(intIndex).Text = Format(rsTemp!进入统筹金额, strFormat_金额)
    txt全自付(intIndex).Text = Format(rsTemp!全自付金额, strFormat_金额)
    txt首先自付(intIndex).Text = Format(rsTemp!首先自付金额, strFormat_金额)
    txt统筹支付(intIndex).Text = Format(rsTemp!统筹报销金额, strFormat_金额)
    If Not mbln住院 Then
        txt帐户支付(intIndex).Text = Format(rsTemp!个人帐户支付, strFormat_金额)
    Else
        txt统筹自付(intIndex).Text = Format(Nvl(rsTemp!进入统筹金额, 0) - Nvl(rsTemp!统筹报销金额, 0), strFormat_金额)
    End If
    
    '读取大类结算数据
    gstrSQL = "Select 性质,大类名称,费用总额,报销总额 From 保险报销记录 Where 记录ID=" & mlng记录ID & " Order by 大类编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    cbo就医(intIndex).ListIndex = rsTemp!性质 - 1
    intCol = IIf(mbln住院, 3, 2)
    cur费用金额 = 0: cur报销总额 = 0
    Set msfObj = Bill(intIndex)
    msfObj.Rows = 2 + rsTemp.RecordCount
    Do While Not rsTemp.EOF
        msfObj.TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!大类名称
        msfObj.TextMatrix(rsTemp.AbsolutePosition, 1) = Format(rsTemp!费用总额, strFormat_金额)
        msfObj.TextMatrix(rsTemp.AbsolutePosition, intCol) = Format(rsTemp!报销总额, strFormat_金额)
        cur费用金额 = cur费用金额 + Nvl(rsTemp!费用总额, 0)
        cur报销总额 = cur报销总额 + Nvl(rsTemp!报销总额, 0)
        rsTemp.MoveNext
    Loop
    msfObj.TextMatrix(msfObj.Rows - 1, 0) = "合计"
    msfObj.TextMatrix(msfObj.Rows - 1, 1) = Format(cur费用金额, strFormat_金额)
    msfObj.TextMatrix(msfObj.Rows - 1, msfObj.Cols - 1) = Format(cur报销总额, strFormat_金额)
    
    ReadData = True
End Function

Private Sub InitFace()
    With cbo就医(int门诊)
        .Clear
        .AddItem "本院"
        .AddItem "外院"
        .ListIndex = 0
    End With
    With cbo就医(int住院)
        .Clear
        .AddItem "本院"
        .AddItem "外院"
        .ListIndex = 0
    End With

    Call InitBill(Bill(int门诊))
    Call InitBill(Bill(int住院))
End Sub

Private Sub InitBill(ByVal msfObj As BillEdit)
    With msfObj
        .ClearBill
        .Active = True
        .Rows = 2
        .Cols = IIf(mbln住院, 4, 3)
        .msfObj.FixedCols = 1
        
        .TextMatrix(0, 0) = "大类"
        .TextMatrix(0, 1) = "费用总额"
        If mbln住院 Then
            .TextMatrix(0, 2) = "数量"
            .TextMatrix(0, 3) = "允许报销"
        Else
            .TextMatrix(0, 2) = "允许报销"
        End If
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        If mbln住院 Then .ColWidth(3) = 1200: .ColWidth(0) = 1200: .ColWidth(2) = 800
        
        .ColData(0) = 5
        .ColData(1) = 4
        If mbln住院 Then
            .ColData(2) = 4
            .ColData(3) = 5
        Else
            .ColData(2) = 5
        End If
        .PrimaryCol = 1
        .LocateCol = 1
    End With
End Sub

Private Sub WriteBill(ByVal rsTemp As ADODB.Recordset, ByVal msfObj As BillEdit)
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !名称 = "浮动比例" And msfObj.Index = 0 Then
                '门诊不显示该大类
                .MoveNext
                If .EOF Then Exit Do
            End If
            msfObj.TextMatrix(.AbsolutePosition, 0) = !名称
            msfObj.RowData(.AbsolutePosition) = 0
            msfObj.Rows = msfObj.Rows + 1
            .MoveNext
        Loop
        
        msfObj.TextMatrix(msfObj.Rows - 1, 0) = "合计"
    End With
End Sub

Private Sub txt卡号_GotFocus(Index As Integer)
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt卡号(Index)
End Sub

Private Sub txt卡号_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    
    If Len(txt卡号(Index).Text) = txt卡号(Index).MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = UCase(Replace(Trim(txt卡号(Index).Text), "'", ""))
        If Len(strCode) = 0 Then Exit Sub
        
        If IsNumeric(Mid(strCode, 1, Len(strCode) - 1)) Then '刷卡
            str条件 = " and A.卡号='" & strCode & "'"
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
            str条件 = " and A.病人ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号(对住(过)院的病人)
            str条件 = " and B.住院号=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号(仅对门诊病人)
            str条件 = " and B.门诊号=" & Mid(strCode, 2)
        Else '当作姓名
            str条件 = " and A.卡号='" & strCode & "'"
        End If
    
        gstrSQL = "Select A.病人ID as ID,A.卡号,A.医保号,'******' 密码,B.姓名,B.性别,B.出生日期,B.身份证号,A.中心 中心,C.名称 中心名称  " & _
                 "  ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,E.序号 人群,E.名称 状态,A.退休证号,A.帐户余额 " & _
                 "  From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D,保险人群 E " & _
                 "   where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & TYPE_四川眉山 & _
                 "   and A.险类=C.险类 and A.中心=C.序号 And A.在职=E.序号 and A.险类=E.险类 and A.病种ID=D.ID(+)" & str条件
        Call Get帐户情况(Index)
    End If
    
    If mint编辑模式 = 1 Then cmd确定.Enabled = False
End Sub

Private Sub txt卡号_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub Get帐户情况(Index As Integer)
'从已经存在的记录中读出帐户信息
    Dim rsTemp As New ADODB.Recordset
    Dim rs帐户 As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , txt卡号(Index).Text, "", False, True)
    If rs帐户.State <> 1 Then Exit Sub
    If Not rs帐户 Is Nothing Then
        '检查病人状态
        gstrSQL = "select nvl(当前状态,0) as 状态,灰度级,备注 from 保险帐户 where 险类=25 and 医保号='" & Trim(txt卡号(Index).Text) & "'"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            Select Case Nvl(rsTemp!灰度级, 0)
            Case 1
                MsgBox "该医保卡已经封锁，不能使用！" & IIf(Nvl(rsTemp!备注) <> "", "（" & rsTemp!备注 & "）", ""), vbInformation, gstrSysName
                Exit Sub
            Case 9
                MsgBox "该医保卡已经撤销，不能使用！", vbInformation, gstrSysName
                Exit Sub
            End Select
        End If
        
        '检查数据库中的数据是否正确
        If Not 检查帐户信息_米易(txt卡号(Index).Text) Then Exit Sub
        
        txt卡号(Index).Text = rs帐户("卡号")
        txt卡号(Index).Tag = rs帐户("ID")
        txt姓名(Index).Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txt状态(Index).Text = rs帐户("状态")
        txt状态(Index).Tag = rs帐户("人群")
        If mbln住院 = False Then
            txt病种(Index).Text = Nvl(rs帐户("病种"))
            txt病种(Index).Tag = Nvl(rs帐户("病种ID"), 0)
        End If
        txt中心(Index).Text = rs帐户("中心名称")
        txt中心(Index).Tag = rs帐户("中心")
        txt余额(Index).Text = Format(rs帐户!帐户余额, strFormat_金额)
        lblNote.Caption = Nvl(rs帐户!退休证号)
        If Index = int住院 Then
            '再读出帐户年度信息
            gstrSQL = "select * from 帐户年度信息 where 险类=" & TYPE_四川眉山 & _
                " and 病人ID=" & rs帐户("ID") & " and 年度=" & Format(zlDatabase.Currentdate, "yyyy")
            Call OpenRecordset(rsTemp, Me.Caption)
            
            If rsTemp.EOF = False Then
                '设置帐户情况
                If cbo就医(Index).ListIndex = 0 Then
                    txt住院次数(Index).Text = Nvl(rsTemp("住院次数累计"), 0) + IIf(mint编辑模式 = 1, 1, 0)
                Else
                    txt住院次数(Index).Text = Nvl(rsTemp("外院住院次数"), 0) + IIf(mint编辑模式 = 1, 1, 0)
                End If
                txt余额(Index).Text = txt余额(Index).Text & "/" & Format(rsTemp!进入统筹累计, strFormat_金额)
            Else
                txt住院次数(Index).Text = "1"
            End If
            
            Call cbo就医_Click(Index)
        End If
    End If
End Sub

Private Function SaveData(Optional ByVal bln本院 As Boolean = True) As Boolean
    '保存门诊结算和住院结算数据
    Dim lng结帐ID As Long, str医保号 As String, intIndex As Integer
    Dim blnExecute As Boolean
    
    '取结帐流水号及病人医保卡号
    intIndex = IIf(mbln住院, int住院, int门诊)
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    str医保号 = txt卡号(intIndex)
    
    '保存
    gcnOracle.BeginTrans
    If mbln住院 Then
        blnExecute = 住院结算(lng结帐ID, bln本院)
    Else
        blnExecute = 门诊结算_眉山(lng结帐ID, Val(txt帐户支付(int门诊)), str医保号, bln本院)
    End If
    If Not blnExecute Then
        gcnOracle.RollbackTrans
    Else
        gcnOracle.CommitTrans
        mlng记录ID = lng结帐ID
    End If
    
    SaveData = blnExecute
End Function

Private Function 住院结算(lng结帐ID As Long, Optional ByVal bln本院治疗 As Boolean = True) As Boolean
    Dim int性质 As Integer
    Dim lng年度 As Long, int本院 As Integer, int外院 As Integer
    Dim cur帐户余额 As Currency, cur统筹累计 As Currency
    Dim rsTemp As New ADODB.Recordset
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    On Error GoTo ErrHand
    int性质 = IIf(bln本院治疗, 1, 2)
    
    '将结算信息保存到保险结算记录中
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_四川眉山 & "," & gComInfo_眉山.病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0," & Val(txt住院次数(int住院).Text) & "," & gComInfo_眉山.起付线 & ",0," & gComInfo_眉山.本次起付线 & "," & _
        gComInfo_眉山.费用总额 & "," & gComInfo_眉山.全自付 & "," & gComInfo_眉山.首先自付 & "," & gComInfo_眉山.进入统筹 & "," & gComInfo_眉山.统筹支付 & ",0," & _
        0 & "," & 0 & ",null,null,null,null,null,'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    '保险各大类的报销明细
    With rs大类_米易
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!费用总额, 0) <> 0 Then
                gstrSQL = "ZL_保险报销记录_INSERT(" & int性质 & "," & lng结帐ID & "," & _
                "'" & !大类编码 & "','" & !大类名称 & "'," & !统筹比额 & "," & _
                "" & !特准定额 & "," & !特准天数 & "," & !费用总额 & "," & !报销总额 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存大类报销数据")
            End If
            .MoveNext
        Loop
    End With
    
    '保险分档报销明细
    With rs分档支付_米易
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !进入统筹 <> 0 Then
                gstrSQL = "ZL_保险结算计算_INSERT(" & lng结帐ID & "," & !档次 & "," & !进入统筹 & "," & !统筹报销 & "," & !比例 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存分档报销明细")
            End If
            .MoveNext
        Loop
    End With
    
    '更新住院次数
    lng年度 = Format(zlDatabase.Currentdate, "yyyy")
    cur帐户余额 = 0: cur统筹累计 = gComInfo_眉山.进入统筹
    gstrSQL = " Select Nvl(帐户增加累计,0) 帐户余额,nvl(进入统筹累计,0) 统筹累计" & _
              " ,Nvl(住院次数累计,0) 本院,Nvl(外院住院次数,0) 外院" & _
              " From 帐户年度信息" & _
              " Where 年度=" & lng年度 & " And 病人ID=" & gComInfo_眉山.病人ID
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not rsTemp.EOF Then
        cur帐户余额 = rsTemp!帐户余额: cur统筹累计 = cur统筹累计 + rsTemp!统筹累计
        If bln本院治疗 Then
            int本院 = Val(txt住院次数(int住院).Text)
            int外院 = rsTemp!外院
        Else
            int本院 = rsTemp!本院
            int外院 = Val(txt住院次数(int住院).Text)
        End If
    Else
        If bln本院治疗 Then
            int本院 = 1
            int外院 = 0
        Else
            int本院 = 0
            int外院 = 1
        End If
    End If
    
    gstrSQL = "zl_帐户年度信息_Insert(" & gComInfo_眉山.病人ID & ",25," & lng年度 & _
              "," & cur帐户余额 & ",0," & cur统筹累计 & ",0," & int本院 & "," & int外院 & "," & Val(txt起付线(int住院).Text) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新住院次数")
    
    住院结算 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub 显示实际报销比例()
    Dim lngRow As Long
    Dim sin比例 As Single
    Dim rsTemp As New ADODB.Recordset
    
    '如果是住院预结算，发现有实际报销比例的大类不等于100%，将其报销比例显示在数量列
    If mbln住院 Then
        gstrSQL = "Select 参数名,参数值 From 保险参数 Where 中心=1 And 序号>=10"
        Call OpenRecordset(rsTemp, Me.Caption)
        With Bill(int住院)
            For lngRow = 1 To .Rows - 2
                sin比例 = 100
                If rsTemp.RecordCount <> 0 Then
                    rsTemp.MoveFirst
                    rsTemp.Find "参数名='" & .TextMatrix(lngRow, 0) & "'"
                    If Not rsTemp.EOF Then sin比例 = Nvl(rsTemp!参数值, 100)
                End If
                If sin比例 <> 100 Or .TextMatrix(lngRow, 0) = "浮动比例" Then
                    '比例不等于100%才允许调整
                    .TextMatrix(lngRow, 2) = Format(sin比例, strFormat_金额)
                    .RowData(lngRow) = 1
                End If
            Next
        End With
    End If
End Sub

Private Sub 实际报销比例()
    Dim lngRow As Long
    gstr实际报销比例_米易 = ""
    If mbln住院 = False Then Exit Sub
    
    With Bill(int住院)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) = 1 Then
                gstr实际报销比例_米易 = gstr实际报销比例_米易 & "|" & .TextMatrix(lngRow, 0) & ";" & Val(.TextMatrix(lngRow, 2))
            End If
        Next
    End With
End Sub
