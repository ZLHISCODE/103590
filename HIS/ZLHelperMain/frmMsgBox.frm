VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "消息提示"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5970
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDelay 
      Caption         =   "推迟2分钟(D)"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Timer tmrExit 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   960
   End
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      Caption         =   "20秒后自动关闭！"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4320
      TabIndex        =   2
      Top             =   840
      Width           =   1440
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmMsgBox.frx":6852
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "将要进行当前客户端的功能验证，可能会影响你的正常使用，请先保存重要工作！"
      Height          =   420
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   4920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrTips        As String
Private mlngTick        As Long
Private mblnFirst       As Boolean
Private mblnDelay       As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function ShowMe(ByVal strTips As String) As Boolean
    mstrTips = strTips
    gblnMsgBox = True
    mblnFirst = True
    mblnDelay = False
    Me.Show vbModal
    ShowMe = mblnDelay
End Function

Private Sub cmdDelay_Click()
    mblnFirst = False
    gblnMsgBox = False
    mblnDelay = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mblnFirst = False
    gblnMsgBox = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        SetWindowPos Me.hwnd, -1, (Screen.Width - Me.Width) / 2 / 15, (Screen.Height - Me.Height) / 2 / 15, 0, 0, 1
        mblnFirst = False
    End If
End Sub

Private Sub Form_Load()
    mlngTick = GetTickCount
    lblInfo.Caption = mstrTips
    tmrExit.Enabled = True
End Sub

Private Sub tmrExit_Timer()
    Dim lngSec     As Long
    
    lngSec = CLng(GetTickCountDiff(mlngTick) / 1000)
    If lngSec >= 20 Then
        Call cmdOK_Click
    Else
        If lngSec > 10 Then
            lblTips.Caption = " " & (20 - lngSec) & "秒后自动关闭！"
        Else
            lblTips.Caption = (20 - lngSec) & "秒后自动关闭！"
        End If
    End If
End Sub
