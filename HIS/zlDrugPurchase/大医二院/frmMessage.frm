VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MESS"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   840
      Width           =   990
   End
   Begin VB.TextBox txtContents 
      Appearance      =   0  'Flat
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton cmdContents 
      Caption         =   "详细内容(&T)"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbrMess 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblMess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mess"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdContents_Click()
    Call SetDetail
End Sub

Private Sub SetDetail(Optional ByVal blnInit As Boolean)
    If InStr(cmdContents.Caption, "详细") = 0 Or blnInit = True Then
        Me.Height = Me.Height - Me.ScaleHeight + cmdContents.Top + cmdContents.Height + 80
        cmdContents.Caption = "详细内容(&T)"
    Else
        Me.Height = 6050
        cmdContents.Caption = "简略内容(&T)"
    End If
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Load()
    Call SetDetail(True)
End Sub
