VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWindowCustom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "自定义调窗"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3120
      TabIndex        =   3
      Top             =   1050
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Height          =   350
      Left            =   960
      TabIndex        =   2
      Top             =   1050
      Width           =   1100
   End
   Begin MSComCtl2.UpDown udWindow 
      Height          =   300
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtWindow 
      Height          =   300
      Index           =   2
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtWindow 
      Height          =   300
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSComCtl2.UpDown udWindow 
      Height          =   300
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "窗位"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "窗宽"
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   210
      Width           =   360
   End
End
Attribute VB_Name = "frmWindowCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngWindow As Long
Public lngLevel As Long
Public bApply As Boolean

Private Sub cmdApply_Click()
    bApply = True
    lngWindow = Val(Me.txtWindow(1).Text)
    lngLevel = Val(Me.txtWindow(2).Text)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    bApply = False
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtWindow(1).Text = lngWindow
    Me.txtWindow(2).Text = lngLevel
    txtWindow(1).SelStart = 0
    txtWindow(1).SelLength = Len(txtWindow(1).Text)
End Sub
Private Sub txtWindow_GotFocus(Index As Integer)
    txtWindow(Index).SelStart = 0
    txtWindow(Index).SelLength = Len(txtWindow(Index).Text)
End Sub

Private Sub txtWindow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 1 Then
            txtWindow(2).SetFocus
        Else
            cmdApply.SetFocus
        End If
    End If
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub udWindow_DownClick(Index As Integer)
    Me.txtWindow(Index).Text = Val(Me.txtWindow(Index).Text) - 1
End Sub

Private Sub udWindow_UpClick(Index As Integer)
    Me.txtWindow(Index).Text = Val(Me.txtWindow(Index).Text) + 1
End Sub
