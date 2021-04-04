VERSION 5.00
Begin VB.Form frmChangeStat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "调整状态"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangeStat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3090
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1230
      TabIndex        =   6
      Top             =   1680
      Width           =   1100
   End
   Begin VB.OptionButton Opt 
      Caption         =   "3-退号"
      Height          =   225
      Index           =   5
      Left            =   3975
      TabIndex        =   5
      Top             =   472
      Width           =   1050
   End
   Begin VB.OptionButton Opt 
      Caption         =   "2-弃号"
      Height          =   225
      Index           =   4
      Left            =   3975
      TabIndex        =   4
      Top             =   90
      Width           =   1050
   End
   Begin VB.OptionButton Opt 
      Caption         =   "4-结束"
      Height          =   195
      Index           =   3
      Left            =   3975
      TabIndex        =   3
      Top             =   855
      Width           =   1050
   End
   Begin VB.OptionButton Opt 
      Caption         =   "7-执行中"
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   472
      Width           =   1050
   End
   Begin VB.OptionButton Opt 
      Caption         =   "5-待穿刺"
      Height          =   195
      Index           =   1
      Left            =   1305
      TabIndex        =   1
      Top             =   472
      Width           =   1050
   End
   Begin VB.OptionButton Opt 
      Caption         =   "1-待配液"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   472
      Width           =   1050
   End
   Begin VB.Line Line9 
      X1              =   3585
      X2              =   3985
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line8 
      X1              =   3585
      X2              =   3985
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line7 
      X1              =   3570
      X2              =   3570
      Y1              =   180
      Y2              =   975
   End
   Begin VB.Line Line6 
      X1              =   3570
      X2              =   3970
      Y1              =   165
      Y2              =   165
   End
   Begin VB.Line Line4 
      X1              =   1095
      X2              =   3630
      Y1              =   525
      Y2              =   555
   End
End
Attribute VB_Name = "frmChangeStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrStat As String
Private mblnOk As Boolean
Private mblnLiquid As Boolean

Public Function ShowMe(ByRef strStat As String, ByVal blnLiquid As Boolean) As Boolean
    mblnOk = False
    mstrStat = strStat
    mblnLiquid = blnLiquid
    
    Me.Show vbModal
    ShowMe = mblnOk
    If mblnOk Then
        strStat = mstrStat
    End If
    
End Function

Private Sub cmdCancle_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    For i = Opt.LBound To Opt.UBound
        If Opt.Item(i).Value = True Then
            mstrStat = Opt.Item(i).Caption
            Exit For
        End If
    Next
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = Opt.LBound To Opt.UBound
        If Opt.Item(i).Caption = mstrStat Then
            Opt.Item(i).Value = True
            Exit For
        End If
    Next
    
    Opt.Item(0).Enabled = mblnLiquid
    
End Sub
