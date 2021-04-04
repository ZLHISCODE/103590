VERSION 5.00
Begin VB.Form frmPatiFeeFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找条件 (查找后,按F3查找下一个)"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3030
      TabIndex        =   7
      Top             =   675
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3030
      TabIndex        =   6
      Top             =   225
      Width           =   1100
   End
   Begin VB.Frame fra条件 
      Caption         =   "条件范围："
      Height          =   1530
      Left            =   135
      TabIndex        =   8
      Top             =   90
      Width           =   2700
      Begin VB.TextBox txtBed 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   1
         Top             =   270
         Width           =   1500
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1080
         Width           =   1500
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   18
         TabIndex        =   3
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床  号(&1)"
         Height          =   180
         Left            =   135
         TabIndex        =   0
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓  名(&3)"
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lbl病员号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&2)"
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmPatiFeeFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Sub cmdCancel_Click()
    gblnOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txt姓名.Text = "" And txt住院号.Text = "" And txtBed.Text = "" Then
        MsgBox "查找至少需要输入一个条件！", vbInformation, gstrSysName
        txtBed.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt姓名.Text) > txt姓名.MaxLength Then MsgBox "病人姓名超长！", vbInformation, gstrSysName: Exit Sub
    gblnOK = True
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtBed_GotFocus()
    zlControl.TxtSelAll txtBed
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
