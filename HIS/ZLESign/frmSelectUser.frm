VERSION 5.00
Begin VB.Form frmSelectUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户选择"
   ClientHeight    =   1410
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "frmSelectUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1695
      TabIndex        =   2
      Top             =   840
      Width           =   1100
   End
   Begin VB.ComboBox cboUser 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   285
      Width           =   2955
   End
   Begin VB.Label lblUser 
      Caption         =   "用户名"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "frmSelectUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUser As String

Public Function ShowMe(ByVal strUser As String) As String
    mstrUser = strUser
    
    Me.Show 1
    
    ShowMe = mstrUser
End Function

Private Sub cmdOK_Click()
    If cboUser.ListIndex = -1 Then
        MsgBoxEx "请选择用户。", vbInformation, "注意"
        cboUser.SetFocus: Exit Sub
    End If
    
    mstrUser = cboUser.Text
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim arrUsers() As String
    
    arrUsers = Split(mstrUser, "&&&")
    
    For i = 0 To UBound(arrUsers)
        cboUser.AddItem arrUsers(i)
    Next
    
    If cboUser.ListCount > 0 Then cboUser.ListIndex = 0
    
    mstrUser = ""
End Sub
