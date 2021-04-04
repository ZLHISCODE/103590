VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5580
   Icon            =   "frmInputBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   5730
      Left            =   4185
      TabIndex        =   5
      Top             =   -915
      Width           =   30
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   1005
      TabIndex        =   1
      Top             =   1035
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   3
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4380
      TabIndex        =   2
      Top             =   105
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmInputBox.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   180
      Left            =   825
      TabIndex        =   0
      Top             =   1095
      Width           =   90
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   180
      Left            =   825
      TabIndex        =   4
      Top             =   255
      Width           =   3135
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrInput As String
Private mbytDataType As Byte

Public Function ShowInputBox(ByVal frmMain As Object, _
                            ByVal strCaption As String, _
                            ByVal strNote As String, _
                            ByVal strLable As String, _
                            ByRef strInput As String, _
                            Optional ByVal bytDataType As Byte = 1, _
                            Optional ByVal lngMaxLen As Long = 0) As Boolean
    
    mstrInput = strInput
    mbytDataType = bytDataType
    
    lbl.Caption = strLable
    
    txt.Text = strInput
    txt.Left = lbl.Left + lbl.Width + 30
    Me.Caption = strCaption
    lblNote.Caption = strNote
    
    txt.MaxLength = lngMaxLen
    
    Me.Show 1, frmMain
    
    ShowInputBox = mblnOK
    If mblnOK Then strInput = mstrInput
    
End Function

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Select Case mbytDataType
    Case 1
        mstrInput = txt.Text
    Case 2
        mstrInput = txt.Text
    End Select
    
    mblnOK = True
    
    Unload Me
    
End Sub

Private Sub txt_GotFocus()
    Call zlControl.TxtSelAll(txt)
    zlCommFun.OpenIme True
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
End Sub
