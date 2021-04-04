VERSION 5.00
Begin VB.Form frmInputBoxV2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "文字标注设置"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtInput 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入新的文字标注配置，格式为“简码1=说明1|简码2=说明2|..."
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmInputBoxV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrInput As String
Private mObjOwner As Form
'"请输入新的文字标注配置，格式为“简码1=说明1|简码2=说明2|...”。", "文字标注设置", Replace(mstrTemp, "[+]", "|"))


Private Sub cmdCancel_Click()
    mstrInput = ""
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mstrInput = txtInput.Text
    Me.Hide
End Sub

Private Sub Form_Load()
    txtInput.Text = ""
End Sub

Public Function ZlShowMe(objForm As Form, ByVal strtext As String) As String
    Dim lngLeft As Long, lngTop As Long
    
    Set mObjOwner = objForm
    
    lngLeft = objForm.Left + (objForm.Width - Me.Width) / 2
    lngTop = objForm.Top + (objForm.Height - Me.Height) / 2
    
    If lngLeft <= 0 Then lngLeft = 0
    If lngTop <= 0 Then lngTop = 0
    
    Call Me.Move(lngLeft, lngTop)
    
    txtInput.Text = strtext
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '将窗口置顶
    Me.Show 1, mObjOwner
    
    
    ZlShowMe = mstrInput
End Function


Private Sub Form_Unload(Cancel As Integer)
    Set mObjOwner = Nothing
End Sub
