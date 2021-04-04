VERSION 5.00
Begin VB.Form frm门诊找零贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊结算信息"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frm门诊找零贵阳.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   1770
   End
   Begin VB.Frame Frame2 
      Caption         =   "当前共#张门诊收费单据:"
      Height          =   1635
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   4935
      Begin VB.TextBox txt差额记帐 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txt现金支付 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txt大病基金 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txt医保基金 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "差额记帐"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现金支付"
         Height          =   180
         Index           =   5
         Left            =   2460
         TabIndex        =   15
         Top             =   1297
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大病基金"
         Height          =   180
         Index           =   7
         Left            =   2460
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保基金"
         Height          =   180
         Index           =   8
         Left            =   2460
         TabIndex        =   8
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -90
      TabIndex        =   1
      Top             =   1770
      Width           =   5085
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   3540
      TabIndex        =   0
      Top             =   1890
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注意：本窗口在30秒钟后自动关闭"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   1935
      Width           =   2700
   End
End
Attribute VB_Name = "frm门诊找零贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdat As Date
Private Sub cmdOK_Click()
    If gblnLED And Val(txt现金支付.Text) > 0 Then
        zl9LedVoice.Reset Nothing
        zl9LedVoice.Speak "#21" & Val(txt现金支付.Text)
    End If
    Unload Me
End Sub
Public Sub ShowForm(ByVal intCOUNT As Long)
    Frame2.Caption = Replace(Frame2.Caption, "#", intCOUNT)
    txt费用总额.Text = Format(g门诊数据.dbl费用总额, "0.00")
    txt个人帐户.Text = Format(g门诊数据.dbl个人帐户, "0.00")
    txt医保基金.Text = Format(g门诊数据.dbl医保基金, "0.00")
    txt医疗补助.Text = Format(g门诊数据.dbl公务员补助, "0.00")
    txt大病基金.Text = Format(g门诊数据.dbl大病基金, "0.00")
    txt现金支付.Text = Format(g门诊数据.dbl现金, "0.00")
    txt差额记帐.Text = Format(g门诊数据.dbl差额记帐, "0.00")
    Label1.Caption = "注意：本窗口在" & mlngCloseTime & "秒钟后自动关闭"
    mdat = Now
    If gblnLED Then
        Call zl9LedVoice.DisplayBank("单据张数:" & intCOUNT & " 费用总额:" & txt费用总额.Text, "个人帐户:" & txt个人帐户.Text & " 医保基金:" & txt医保基金.Text, _
                "医疗补助:" & txt医疗补助.Text & " 大病基金:" & txt大病基金.Text, "现金支付:" & txt现金支付.Text, "差额记帐:" & txt差额记帐.Text)
    End If
    Me.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Timer1_Timer()
    '窗口停留时间超过30秒钟时，自动关闭本窗口，要不然HIS事务一直无法提交，引起全院死锁
    If Abs(DateDiff("s", mdat, Now)) > mlngCloseTime Then Unload Me
End Sub

