VERSION 5.00
Begin VB.Form frmlabONSampleEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "标本保存"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   3930
      TabIndex        =   3
      Top             =   3420
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5730
      TabIndex        =   4
      Top             =   3420
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   -960
      TabIndex        =   8
      Top             =   3270
      Width           =   8445
   End
   Begin VB.TextBox txtEnvironment 
      Height          =   1845
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   5505
   End
   Begin VB.TextBox txtlocation 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   3675
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "保存人:"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   285
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "保存环境:"
      Height          =   180
      Left            =   300
      TabIndex        =   6
      Top             =   1260
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "保存位置:"
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   772
      Width           =   810
   End
End
Attribute VB_Name = "frmlabONSampleEdit"
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
    mstrVal = txtName & "|" & txtlocation & "|" & txtEnvironment
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

Private Sub txtlocation_GotFocus()
    txtlocation.SelStart = 0
    txtlocation.SelLength = Len(txtlocation)
End Sub



Private Sub txtlocation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEnvironment.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0
    Me.txtName.SelLength = Len(Me.txtName)
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtlocation.SetFocus
    End If
End Sub
