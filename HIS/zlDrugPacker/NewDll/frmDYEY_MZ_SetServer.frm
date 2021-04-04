VERSION 5.00
Begin VB.Form frmDYEY_MZ_SetServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置WebService服务地址"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   Icon            =   "frmDYEY_MZ_SetServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5550
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtURL 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   300
      Width           =   4335
   End
   Begin VB.Label lblUrl 
      Caption         =   "服务地址"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmDYEY_MZ_SetServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrURL As String

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrURL = Me.txtURL.Text
    Unload Me
End Sub

Public Sub ShowMe(ByRef strUrl As String)
    mstrURL = ""
    Me.Show 1
    strUrl = mstrURL
End Sub

Private Sub Form_Load()
    mstrURL = GetSetting("ZLSOFT", "公共模块\WebService路径", "WebUrl")
    If Trim(mstrURL) = "" Then
        mstrURL = GetINIInfo("WebService路径")
    End If
    txtURL.Text = mstrURL
End Sub


