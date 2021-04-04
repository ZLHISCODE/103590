VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加组别"
   ClientHeight    =   2235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "frmName.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2565
      TabIndex        =   2
      Top             =   1725
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   3
      Top             =   1725
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -1260
      TabIndex        =   5
      Top             =   1545
      Width           =   6780
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   1980
      MaxLength       =   30
      TabIndex        =   1
      Top             =   975
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "组别(&N)"
      Height          =   180
      Index           =   1
      Left            =   1335
      TabIndex        =   0
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请在下面输入新的体检组别名称："
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   4
      Top             =   315
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   315
      Picture         =   "frmName.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrName As String

Public Function ShowName(ByVal frmMain As Object) As Boolean
    
    Set mfrmMain = frmMain
    
    Me.Show 1, frmMain
    
    ShowName = mblnOK
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Trim(txt.Text) = "" Then Exit Sub
    
    On Error Resume Next
    
    If mfrmMain.NameRefresh(txt.Text) Then
        ShowSimpleMsg "添加组别成功，继续添加！"
        txt.Text = ""
        txt.SetFocus
    End If
    
End Sub

Private Sub txt_GotFocus()
    
    zlControl.TxtSelAll txt
    zlCommFun.OpenIme True
    
End Sub

Private Sub txt_LostFocus()

    zlCommFun.OpenIme False
    
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

