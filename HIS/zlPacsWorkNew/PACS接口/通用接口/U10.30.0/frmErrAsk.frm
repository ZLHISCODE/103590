VERSION 5.00
Begin VB.Form frmErrAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提示"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtHelp 
      Height          =   1260
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmErrAsk.frx":0000
      Top             =   1845
      Width           =   4275
   End
   Begin VB.CommandButton cmdCopyScreen 
      Caption         =   "截图(&S)"
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2175
      TabIndex        =   5
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "重试(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1230
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2940
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   8
      Top             =   1605
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblAsk 
      AutoSize        =   -1  'True
      Caption         =   "再试一次吗？"
      Height          =   180
      Left            =   975
      TabIndex        =   3
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label lblNote 
      Caption         =   "    可能是其他用户的独占或重新安装了操作系统带来的错误，排除独占使用因素仍不能运行，则需部分重装本系统。"
      Height          =   585
      Left            =   975
      TabIndex        =   2
      Top             =   360
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "说明："
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "序号："
      Height          =   180
      Left            =   3150
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmErrAsk.frx":0065
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytReturn As Byte

Private Sub cmdCancel_Click()
    mbytReturn = 0
    Unload Me
End Sub

Private Sub cmdCopyScreen_Click()
'    Call SaveScreen(txtHelp.Text, picS)
End Sub

Private Sub cmdRetry_Click()
    mbytReturn = 1
    Unload Me
End Sub

Public Function ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String) As Byte
'功能：显示错误提示窗口，可以选择重试
'参数：lngErrNum   错误编号
'      strNote     错误内容
'      strErrInfo  详细的错误信息
'返回：下一步操作的批示。1-再试；0-取消
    mbytReturn = 0
        
    lblNumber.Caption = "序号：" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo
    
    frmErrAsk.Show vbModal
    ShowEdit = mbytReturn
End Function
