VERSION 5.00
Begin VB.Form frmProcEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改说明填写"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdNo 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   4560
      TabIndex        =   5
      Top             =   3240
      Width           =   990
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   990
   End
   Begin VB.TextBox txtStat 
      Height          =   2415
      Left            =   840
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VB.TextBox txtPerson 
      Height          =   350
      Left            =   840
      MaxLength       =   10
      TabIndex        =   1
      Top             =   282
      Width           =   1815
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "修改说明"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblPerson 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "修改人"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmProcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSave As Boolean
Private mstrStat As String
Private mstrPerson As String

Public Function ShowMe(strPerson As String, strStatement As String) As Boolean
        
    txtStat.Text = strStatement
    mstrStat = strStatement
    txtPerson.Text = strPerson
    mstrPerson = strPerson
    
    Me.Show 1
    strPerson = mstrPerson
    strStatement = mstrStat
    ShowMe = mblnSave
End Function

Private Sub cmdNo_Click()
    mblnSave = False
    Unload Me
End Sub


Private Sub cmdYes_Click()
    If txtPerson.Text = "" Or txtStat.Text = "" Then
        MsgBox "修改人和修改说明必须填写!", , "提示"
        Exit Sub
    End If
    mblnSave = True
    mstrPerson = txtPerson.Text
    mstrStat = txtStat.Text
    Unload Me
End Sub

Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtStat_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
