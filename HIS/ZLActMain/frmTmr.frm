VERSION 5.00
Begin VB.Form frmTmr 
   Caption         =   "BH融合父窗体置后"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   3660
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   480
   End
End
Attribute VB_Name = "frmTmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    tmrThis.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrThis.Enabled = False
End Sub


Public Sub SetTimr(ByVal blnEnabled As Boolean)
    tmrThis.Enabled = blnEnabled
End Sub

Private Sub tmrThis_Timer()
    Call SetWindowPos(glngMain, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
