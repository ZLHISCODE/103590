VERSION 5.00
Begin VB.Form frmDeptGo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位条件"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   1440
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2955
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1245
         MaxLength       =   18
         TabIndex        =   3
         Top             =   615
         Width           =   1275
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1245
         MaxLength       =   15
         TabIndex        =   5
         Top             =   990
         Width           =   1275
      End
      Begin VB.TextBox txt床号 
         Height          =   300
         Left            =   1245
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&2)"
         Height          =   180
         Left            =   375
         TabIndex        =   2
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&3)"
         Height          =   180
         Left            =   555
         TabIndex        =   4
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号(&1)"
         Height          =   180
         Left            =   555
         TabIndex        =   0
         Top             =   300
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDeptGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If txt住院号.Text = "" And txt姓名.Text = "" And txt床号.Text = "" Then
        MsgBox "请至少设定一个条件！", vbInformation, gstrSysName
        txt床号.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
   txt床号.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt床号_GotFocus()
    zlControl.TxtSelAll txt床号
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
