VERSION 5.00
Begin VB.Form frmEarnRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "启用\停用原因"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmEarnRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txt原因 
      Height          =   1695
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lbl原因 
      Caption         =   "启用原因(最多能录入50个汉字，含标点)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmEarnRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr原因 As String

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.txt原因.Text = "" Then
        Exit Sub
    End If
    
    If zlCommFun.ActualLen(Me.txt原因.Text) > 100 Then
        MsgBox "当前输入的原因超出50个汉字（含标点）！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    mstr原因 = Me.txt原因.Text
    
    Unload Me
End Sub

Public Sub ShowMe(ByVal intType As Integer, ByRef str原因 As String)
    If intType = 1 Then
        Me.lbl原因.Caption = "启用原因(最多能录入50个汉字，含标点)"
    Else
        Me.lbl原因.Caption = "停用原因(最多能录入50个汉字，含标点)"
    End If
    
    mstr原因 = ""
    
    Me.Show 1
    
    str原因 = mstr原因
End Sub
