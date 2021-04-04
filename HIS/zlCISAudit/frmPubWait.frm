VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPubWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请稍候 ..."
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "frmWait"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1650
      Top             =   1020
   End
   Begin ComctlLib.ProgressBar pgb 
      Height          =   210
      Left            =   75
      TabIndex        =   1
      Top             =   1335
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frmPubWait.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   3
      Top             =   840
      Width           =   5895
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -30
      ScaleHeight     =   825
      ScaleWidth      =   5925
      TabIndex        =   2
      Top             =   0
      Width           =   5925
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中联软件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   1260
      End
   End
   Begin VB.Label lbl内容 
      AutoSize        =   -1  'True
      Caption         =   "正在处理..."
      Height          =   180
      Left            =   75
      TabIndex        =   0
      Top             =   1095
      Width           =   990
   End
End
Attribute VB_Name = "frmPubWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr内容 As String

' API Declarations
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Constants
Private Const SM_CXFULLSCREEN = 16   ' Width of window client area
Private Const SM_CYFULLSCREEN = 17   ' Height of window client area
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

Public mfrmMain As Object

Public Property Let WaitInfo(ByVal vData As String)
    
    mstr内容 = vData
    lbl内容.Caption = vData
    DoEvents
    
End Property

Public Property Let WaitProgress(ByVal vData As Single)
    
    If pgb.Visible = False Then pgb.Visible = True
    pgb.Value = vData
    
    If vData > 0 Then
        lbl内容.Caption = mstr内容 & Format(vData, "0.0") & "%"
    Else
        lbl内容.Caption = mstr内容
    End If
    
    DoEvents
    
End Property

Public Property Let ShowProgress(ByVal vData As Boolean)
    pgb.Visible = vData
End Property

Public Sub CloseWait()
    
    On Error Resume Next
    
'    avi.Stop
    Unload Me
    
End Sub

Public Function OpenWait(ByVal frmMain As Object, Optional ByVal strTitle As String, Optional ByVal ShowProgress As Boolean = False) As Object
    '---------------------------------------------------------------------------------------
    '功能： 弹出提示窗口
    '---------------------------------------------------------------------------------------
    Dim strAviPath As String
   
    If frmMain Is Nothing Then
        Me.Left = (GetSystemMetrics(SM_CXFULLSCREEN) * Screen.TwipsPerPixelX - Me.Width) / 2
        Me.Top = (GetSystemMetrics(SM_CYFULLSCREEN) * Screen.TwipsPerPixelY - Me.Height) / 2
    Else
        Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
        Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2
    End If
    
    ShowWindow Me.hWnd, 4
    SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H10 Or &H1
    
    pgb.Visible = ShowProgress
    lblCompany.Caption = strTitle
        
    On Error Resume Next
    
'    strAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
'
'    avi.Open strAviPath & "\" & "Findfile.avi"
'    avi.Play
    
    DoEvents
    
    Set OpenWait = Me
    
End Function

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Private Sub Timer1_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
