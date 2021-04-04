VERSION 5.00
Begin VB.Form frmListSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "通知修改"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   990
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtInterval 
      Height          =   300
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   3
      Top             =   187
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   187
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblInterval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(s)"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   195
   End
   Begin VB.Label lblInterval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "通知间隔"
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmListSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngInterval As Long
Private mlngNoticeCode As Long

Public Function ShowEdit(ByVal lngNoticeCode As Long, ByVal strName As String, ByVal strInfo As String, ByVal lngInterval As Long) As String
    txtName.Text = strName
    txtInfo.Text = strInfo
    txtInterval.Text = lngInterval
    mlngInterval = lngInterval
    mlngNoticeCode = lngNoticeCode
    
    Me.Show 1
    
    ShowEdit = mlngInterval
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    mlngInterval = Val(txtInterval.Text)
    UpdateNoticeInterval mlngNoticeCode, mlngInterval
    Unload Me
End Sub


Private Sub txtInterval_GotFocus()
    txtInterval.SelStart = 0
    txtInterval.SelLength = Len(txtInterval.Text)
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

