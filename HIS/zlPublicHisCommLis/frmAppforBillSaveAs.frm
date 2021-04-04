VERSION 5.00
Begin VB.Form frmAppforBillSaveAs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "保存常用检验"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2460
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   810
      TabIndex        =   3
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -300
      TabIndex        =   2
      Top             =   1170
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入常用检验名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   2160
   End
End
Attribute VB_Name = "frmAppforBillSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrName As String
Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Function ShowMe(objfrm As Object) As String
    Me.Show vbModal, objfrm
    ShowMe = mstrName
End Function

Private Sub cmdOK_Click()
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "请输入名称后才能保存!", vbInformation, "提示"
        Me.txtName.SetFocus
    End If
    mstrName = Me.txtName.Text
    Unload Me
End Sub

Private Sub Form_Load()
    mstrName = ""
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub
