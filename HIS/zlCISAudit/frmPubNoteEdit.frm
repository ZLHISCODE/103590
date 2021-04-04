VERSION 5.00
Begin VB.Form frmPubNoteEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4410
   Icon            =   "frmPubNoteEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt 
      Height          =   1860
      Index           =   0
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   30
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3225
      TabIndex        =   1
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2010
      TabIndex        =   0
      Top             =   1950
      Width           =   1100
   End
End
Attribute VB_Name = "frmPubNoteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrNote As String

Public Function ShowNoteEdit(ByVal frmMain As Object, ByVal strCaption As String, ByRef strNote As String) As Boolean
    mblnOK = False
    Me.Caption = strCaption
    Me.Show 1, frmMain
    ShowNoteEdit = mblnOK
    strNote = mstrNote
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txt(0).Text = "" Then Exit Sub
    mstrNote = txt(0).Text
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
        Case 0
            zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
        Case 0
            zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


