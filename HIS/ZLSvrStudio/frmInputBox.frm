VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4770
   Icon            =   "frmInputBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1740
      MaxLength       =   12
      TabIndex        =   1
      Top             =   930
      Width           =   2280
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1935
      TabIndex        =   2
      Top             =   1575
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   3
      Top             =   1575
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -450
      TabIndex        =   4
      Top             =   1425
      Width           =   5310
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   180
      Left            =   930
      TabIndex        =   0
      Top             =   990
      Width           =   90
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "frmInputBox.frx":000C
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Caption         =   "#"
      Height          =   630
      Left            =   945
      TabIndex        =   5
      Top             =   180
      Width           =   3525
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   180
      Picture         =   "frmInputBox.frx":1E06
      Top             =   210
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mstrInput As String
Private mblnOK As Boolean

Public Function ShowEdit(ByVal frmMain As Object, ByRef strInput As String, _
                        ByVal strCaption As String, _
                        ByVal strDescrible As String, _
                        ByVal strItemName As String, _
                        Optional ByVal lngMaxLength As Long = 0) As Boolean
    
    mblnOK = False
    
    Me.Caption = strCaption
    
    lblTitle.Caption = strDescrible
    lblName.Caption = strItemName
    txtName.MaxLength = lngMaxLength
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    strInput = mstrInput
    ShowEdit = mblnOK
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If cmdOK.Tag <> "" And Trim(txtName.Text) <> "" Then
        mstrInput = txtName.Text
        mblnOK = True
    End If
    
    Unload Me
End Sub

Private Sub txtName_Change()
    cmdOK.Tag = "Changed"
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txtName.Text, txtName.MaxLength)
End Sub
