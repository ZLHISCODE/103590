VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmZoomCustom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "自定义缩放"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   4
      Top             =   900
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Default         =   -1  'True
      Height          =   350
      Left            =   720
      TabIndex        =   3
      Top             =   900
      Width           =   1100
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtRatio 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Text            =   "1.00"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "自定义缩放比率"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   420
      Width           =   1260
   End
End
Attribute VB_Name = "frmZoomCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bApply As Boolean
Public sRatio As Single

Private Sub cmdApply_Click()
    bApply = True
    sRatio = Val(txtRatio.Text)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    bApply = False
    Unload Me
End Sub

Private Sub Form_Load()
    sRatio = Format(sRatio, ".##")
    If sRatio < 1 Then
        txtRatio.Text = "0" & sRatio
    Else
        txtRatio.Text = sRatio
    End If
    txtRatio.SelStart = 0
    txtRatio.SelLength = Len(txtRatio.Text)
End Sub
Private Sub txtRatio_GotFocus()
    txtRatio.SelStart = 0
    txtRatio.SelLength = Len(txtRatio.Text)
End Sub

Private Sub txtRatio_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub UpDown1_DownClick()
    txtRatio.Text = Val(txtRatio.Text) - 0.01
    If txtRatio.Text < 0.1 Then
        txtRatio.Text = "0.1"
    ElseIf txtRatio.Text < 1 Then
        txtRatio.Text = "0" & txtRatio.Text
    End If
End Sub

Private Sub UpDown1_UpClick()
    txtRatio.Text = Val(txtRatio.Text) + 0.01
    If txtRatio.Text > 16 Then
        txtRatio.Text = "16"
    ElseIf txtRatio.Text < 1 Then
        txtRatio.Text = "0" & txtRatio.Text
    End If
End Sub
