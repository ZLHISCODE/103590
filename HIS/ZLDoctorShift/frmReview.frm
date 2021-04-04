VERSION 5.00
Begin VB.Form frmReview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医生交接班审阅"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   Icon            =   "frmReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3885
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtContent 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmReview.frx":6852
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdContentOK 
      Appearance      =   0  'Flat
      Caption         =   "确认(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdContentCanc 
      Appearance      =   0  'Flat
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1100
   End
   Begin VB.Label lblContent 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "审阅说明"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrContent As String

Public Function ShowMe() As String

    Me.Show 1
    ShowMe = mstrContent
End Function

Private Sub cmdContentCanc_Click()
    If txtContent.Text <> "" Then
        If MsgBox("您已填写审阅说明，确定要退出吗？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            mstrContent = "取消JM"
            Unload Me
            Exit Sub
        Else
            Call zlcontrol.ControlSetFocus(txtContent)
        End If
    Else
        mstrContent = "取消JM"
        Unload Me
    End If
End Sub

Private Sub cmdContentOK_Click()
    mstrContent = Replace(txtContent.Text, "'", "")
    Unload Me
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlcontrol.ControlSetFocus(cmdContentOK)
    End If
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
End Sub
