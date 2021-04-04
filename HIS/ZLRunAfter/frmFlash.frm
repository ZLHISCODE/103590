VERSION 5.00
Begin VB.Form frmFlash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9360
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrThis 
      Interval        =   100
      Left            =   2880
      Top             =   840
   End
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6120
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   9015
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   9030
      TabIndex        =   1
      Top             =   1430
      Visible         =   0   'False
      Width           =   9030
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   9180
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   180
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   9180
         X2              =   9180
         Y1              =   0
         Y2              =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   9180
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblDo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Tag             =   "█"
         Top             =   30
         Width           =   9000
      End
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Left            =   6720
      Picture         =   "frmFlash.frx":27A2
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5秒后会自动隐藏到任务栏，若要查看详情，请点击任务栏图标。"
      Height          =   180
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   5130
   End
   Begin VB.Label lblComent 
      Caption         =   $"frmFlash.frx":346C
      Height          =   480
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   8415
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   120
      Picture         =   "frmFlash.frx":34DA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器：#"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   810
   End
   Begin VB.Label lblPer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4725
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "进  度：#"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1075
      Width           =   810
   End
   Begin VB.Menu mnuShow 
      Caption         =   "显示进度"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           frmFlash
'说明
'==================================================================================================
Private Const mstrCurModule     As String = "frmFlash"           '当前模块名称
Private mblnFirst               As Boolean

Private Sub Form_Activate()
    lblTip.Visible = glngSec > 0
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    DoEvents
End Sub

Private Sub Form_Load()
    mblnFirst = True
    lblTip.Visible = glngSec > 0
    Call AddIcon(picNotify.hWnd, imgNotify.Picture, "延迟脚本执行")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveIcon(picNotify.hWnd)
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '--------------------------------------------------------------------------------------------------
    '功能:  处理picNotify的各种处理事件
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(x) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            On Error Resume Next
            gblnShow = Not gblnShow
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '
End Sub


