VERSION 5.00
Begin VB.Form frmPacsInterfaceParEdit 
   Caption         =   "语句构造"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8445
   Icon            =   "frmPacsInterfaceParEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8445
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdInsertPara 
      Caption         =   "插入参数(&I)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   5940
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退 出(&C)"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtPara 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmPacsInterfaceParEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPara As String
Private mblnIsOk As Boolean
Private mblnEdit As Boolean

Public Function EditPara(ByVal strPara As String, objParent As Object, ByVal blnEdit As Boolean) As String
    EditPara = strPara
    
    mblnEdit = blnEdit
    
    txtPara.Text = strPara
    Me.Show 1, objParent
    
    If mblnIsOk Then EditPara = mstrPara
End Function

Private Sub cmdCancel_Click()
    mblnIsOk = False
    Unload Me
End Sub

Private Sub cmdInsertPara_Click()
    txtPara.SelText = frmPacsInterfacePar.ShowMe(Me)
End Sub

Private Sub cmdSure_Click()
    mblnIsOk = True
    mstrPara = txtPara.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtPara.Enabled = mblnEdit
    cmdInsertPara.Enabled = mblnEdit
    cmdSure.Enabled = mblnEdit
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtPara.Left = 120
    txtPara.Top = 120
    txtPara.Width = Me.ScaleWidth - txtPara.Left * 2
    txtPara.Height = Me.ScaleHeight - txtPara.Top * 3 - cmdCancel.Height
    
    cmdInsertPara.Left = txtPara.Left
    cmdInsertPara.Top = txtPara.Top * 2 + txtPara.Height
    
    cmdSure.Left = Me.ScaleWidth - cmdCancel.Width - cmdSure.Width - 480
    cmdSure.Top = cmdInsertPara.Top
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdCancel.Top = cmdSure.Top
End Sub

Private Sub txtPara_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
