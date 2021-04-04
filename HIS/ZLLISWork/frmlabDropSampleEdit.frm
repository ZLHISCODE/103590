VERSION 5.00
Begin VB.Form frmlabDropSampleUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "标本销毁"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Top             =   210
      Width           =   2415
   End
   Begin VB.TextBox txtEnvironment 
      Height          =   1845
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   750
      Width           =   5505
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   -900
      TabIndex        =   4
      Top             =   2880
      Width           =   8445
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5790
      TabIndex        =   3
      Top             =   3090
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   3990
      TabIndex        =   2
      Top             =   3090
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "销毁方式:"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   810
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "销毁人:"
      Height          =   180
      Left            =   540
      TabIndex        =   5
      Top             =   255
      Width           =   630
   End
End
Attribute VB_Name = "frmlabDropSampleUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrVal As String

Private Sub CmdCancel_Click()
    mstrVal = ""
    Unload Me
End Sub
Public Function ShowMe(Objfrm As Object) As String
    Me.Show vbModal, Objfrm
    ShowMe = mstrVal
End Function

Private Sub cmdSave_Click()
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "必须要输入保存人后才能保存!", vbInformation, "标本保存"
        Me.txtName.SetFocus
    End If
    mstrVal = txtName & "|" & txtEnvironment
    Unload Me
End Sub
Private Sub Form_Load()
    Me.txtName = UserInfo.姓名
End Sub

Private Sub txtEnvironment_GotFocus()
    txtEnvironment.SelStart = 0
    txtEnvironment.SelLength = Len(txtEnvironment)
End Sub

Private Sub txtEnvironment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub







Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0
    Me.txtName.SelLength = Len(Me.txtName)
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEnvironment.SetFocus
    End If
End Sub

