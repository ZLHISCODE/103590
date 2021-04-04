VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入密码"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmPassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1215
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   570
      Width           =   2850
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   90
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -15
      X2              =   5000
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   5000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入密码："
      Height          =   180
      Left            =   1230
      TabIndex        =   0
      Top             =   300
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   300
      Picture         =   "frmPassword.frx":058A
      Top             =   345
      Width           =   720
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlngTime As Long

Private mstrPass As String
Private mbytPWDMin As Byte
Private mbytPWDMax As Byte
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '指定鼠标使用绝对坐标系，此时，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
Private Const MOUSEEVENTF_MOVE = &H1 '移动鼠标
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
Private Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起
 
Private Const SW = 1024
Private Const SH = 768
 
Private Sub Screen_Click(ByVal x As Long, ByVal y As Long)
    Dim mw As Long
    Dim mh As Long
    mw = x / SW * 65535
    mh = y / SH * 65535
    mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, mw, mh, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
 
Public Function ShowMe(strPass As String, Optional ByVal bytPwdMin As Byte, Optional ByVal bytPwdMax As Byte) As Boolean
    strPass = ""
    lblInfo.Caption = ""
    mbytPWDMin = bytPwdMin
    mbytPWDMax = bytPwdMax
    Me.Show 1
    
    If mblnOK Then
        strPass = mstrPass
        ShowMe = True
    End If
End Function

Private Sub cmdCancel_Click()
    If Timer - mlngTime < 2 Then Exit Sub
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Timer - mlngTime < 2 Then Exit Sub
    mstrPass = txtPass.Text
    If mbytPWDMin < mbytPWDMax Then
        If Len(mstrPass) < mbytPWDMin Then
            lblInfo.Caption = "密码不能低于" & mbytPWDMin & "位!" & "实际录入长度为:" & Len(mstrPass)
            Exit Sub
        ElseIf Len(mstrPass) > mbytPWDMax Then
            lblInfo.Caption = "密码不能超过" & mbytPWDMax & "位!" & "实际录入长度为:" & Len(mstrPass)
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    mlngTime = Timer
    Call Screen_Click(512, 384) '模拟鼠标点击窗体获取焦点
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    mstrPass = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Timer - mlngTime < 2 Then Cancel = 1: Exit Sub
End Sub

Private Sub txtPass_Change()
    If lblInfo.Caption <> "" Then lblInfo.Caption = ""
End Sub
