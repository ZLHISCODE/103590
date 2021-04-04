VERSION 5.00
Begin VB.Form Frm交易提示框 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "交易提示框"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "Frm交易提示框.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6840
      TabIndex        =   12
      Top             =   4650
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   11
      Top             =   4380
      Width           =   8835
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   4
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3690
      Width           =   3825
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   3
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3006
      Width           =   3825
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   2
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2324
      Width           =   3825
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   1
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1642
      Width           =   3825
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   3825
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "消费后个人帐户余额"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   5
      Left            =   330
      TabIndex        =   5
      Top             =   3750
      Width           =   4185
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "付现金"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   3120
      TabIndex        =   4
      Top             =   3066
      Width           =   1395
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "其中：个人帐户支付"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   3
      Left            =   330
      TabIndex        =   3
      Top             =   2384
      Width           =   4185
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "本次消费金额"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   1725
      TabIndex        =   2
      Top             =   1702
      Width           =   2790
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   390
      TabIndex        =   1
      Top             =   270
      Width           =   1395
   End
   Begin VB.Label lblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "消费前个人帐户余额"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   1020
      Width           =   4185
   End
End
Attribute VB_Name = "Frm交易提示框"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdbl消费前余额 As Double
Private mdbl消费后余额 As Double
Private mdbl消费总额 As Double
Private mdbl个人帐户 As Double
Private mdbl个人自付 As Double
Private mstr姓名 As String

Public Sub ShowME(ByVal str姓名 As String, ByVal dbl消费前余额 As Double, ByVal dbl消费后余额 As Double, _
    ByVal dbl消费总额 As Double, ByVal dbl个人帐户 As Double, ByVal dbl个人自付 As Double)
    mdbl消费后余额 = dbl消费后余额
    mdbl消费前余额 = dbl消费前余额
    mdbl消费总额 = dbl消费总额
    mdbl个人帐户 = dbl个人帐户
    mdbl个人自付 = dbl个人自付
    mstr姓名 = str姓名
    Me.Show 1
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblNote(1).Caption = Me.lblNote(1).Caption & mstr姓名
    Me.txtNote(0).Text = Format(mdbl消费前余额, "#####0.00;-#####0.00;0;")
    Me.txtNote(1).Text = Format(mdbl消费总额, "#####0.00;-#####0.00;0;")
    Me.txtNote(2).Text = Format(mdbl个人帐户, "#####0.00;-#####0.00;0;")
    Me.txtNote(3).Text = Format(mdbl个人自付, "#####0.00;-#####0.00;0;")
    Me.txtNote(4).Text = Format(mdbl消费后余额, "#####0.00;-#####0.00;0;")
End Sub
