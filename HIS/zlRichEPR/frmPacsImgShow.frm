VERSION 5.00
Begin VB.Form frmPacsImgShow 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ëõ·Å±¨¸æÍ¼"
   ClientHeight    =   4170
   ClientLeft      =   2055
   ClientTop       =   3330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgShow 
      Height          =   3855
      Left            =   120
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmPacsImgShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DblClick(ByRef pic As StdPicture, ByVal strUid As String)

Private Sub Form_Load()
    Me.imgShow.Stretch = True
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.imgShow.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgShow_DblClick()
    Err = 0: On Error GoTo Errhand
    If Me.imgShow.Picture Is Nothing Then Exit Sub
    If Me.imgShow.Picture.Handle = 0 Then Exit Sub
    RaiseEvent DblClick(Me.imgShow.Picture, Me.imgShow.Tag)
Errhand:
End Sub
