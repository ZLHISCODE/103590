VERSION 5.00
Begin VB.Form Frm关闭 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关闭"
   ClientHeight    =   2412
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   Icon            =   "Frm关闭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2412
   ScaleWidth      =   4284
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   2
      Top             =   1950
      Width           =   1100
   End
   Begin VB.ComboBox Cbo关闭 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2685
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   1110
      Width           =   2625
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "希望计算机做什么:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   3
      Top             =   420
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   300
      Picture         =   "Frm关闭.frx":27A2
      Top             =   240
      Width           =   192
   End
End
Attribute VB_Name = "Frm关闭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintStyle As Integer         '-1=关闭;0=注销

Public Function ShowMe(ByRef intStyle As Integer) As Boolean
'参数：
'返回：intStyle=-1=关闭;0=注销
'      是否点击确定关闭
    mintStyle = 0
    Me.Show vbModal
    ShowMe = mblnOK
    intStyle = mintStyle
End Function

Private Sub Cbo关闭_Click()
    With Cbo关闭
        Select Case .ItemData(.ListIndex)
        Case -1
            LblNote.Caption = "关闭系统，回到Windows界面。"
            mintStyle = -1
        Case 0
            LblNote.Caption = "以其他用户的身份重新登录。"
            mintStyle = 0
        End Select
    End With
End Sub

Private Sub Cmd取消_Click()
    mblnOK = False
    mintStyle = 0
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    mblnOK = True
    With Cbo关闭
        Select Case .ItemData(.ListIndex)
        Case -1
            mintStyle = -1
        Case 0
            mintStyle = 0
        End Select
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    With Cbo关闭
        .Clear
        .AddItem "关闭系统"
        .ItemData(.NewIndex) = -1
        .AddItem "注销"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
    End With
End Sub
