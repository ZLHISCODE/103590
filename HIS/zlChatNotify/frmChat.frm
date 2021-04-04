VERSION 5.00
Object = "{BDA06EC7-411C-485C-A7B5-52224D9809BD}#1.0#0"; "SBrowser_G.ocx"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "讨论"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13260
   DrawStyle       =   3  'Dash-Dot
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13260
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraScope 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   135
      Index           =   3
      Left            =   -120
      MousePointer    =   7  'Size N S
      TabIndex        =   15
      Top             =   0
      Width           =   12135
   End
   Begin VB.Frame fraScope 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Index           =   2
      Left            =   11520
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   1320
      Width           =   135
   End
   Begin VB.Frame fraScope 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   135
      Index           =   1
      Left            =   1680
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   6720
      Width           =   5895
   End
   Begin VB.Frame fraScope 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Index           =   0
      Left            =   0
      MousePointer    =   9  'Size W E
      TabIndex        =   12
      Top             =   1320
      Width           =   165
   End
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   11160
      Top             =   4680
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   240
      Width           =   11955
      Begin VB.PictureBox picResize 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9600
         ScaleHeight     =   735
         ScaleWidth      =   1500
         TabIndex        =   3
         Top             =   0
         Width           =   1500
         Begin VB.PictureBox picBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H00D48A00&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   1010
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   8
            Top             =   10
            Width           =   480
            Begin VB.Label lblBtn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "×"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   300
               Index           =   2
               Left            =   75
               TabIndex        =   9
               Top             =   120
               Width           =   330
            End
         End
         Begin VB.PictureBox picBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H00D48A00&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   510
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   6
            Top             =   10
            Width           =   480
            Begin VB.Label lblBtn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "□"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   180
               Index           =   4
               Left            =   180
               TabIndex        =   11
               Top             =   135
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Label lblBtn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "□"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   180
               Index           =   3
               Left            =   105
               TabIndex        =   10
               Top             =   200
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Label lblBtn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "□"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   300
               Index           =   1
               Left            =   480
               TabIndex        =   7
               Top             =   0
               Width           =   330
            End
         End
         Begin VB.PictureBox picBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H00D48A00&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   500
            Index           =   0
            Left            =   10
            ScaleHeight     =   495
            ScaleWidth      =   480
            TabIndex        =   4
            Top             =   10
            Width           =   480
            Begin VB.Label lblBtn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "－"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   300
               Index           =   0
               Left            =   45
               TabIndex        =   5
               Top             =   120
               Width           =   330
            End
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   120
      End
   End
   Begin SBrowser_G.SBrowser SBrowser 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '记录窗体移动前，窗体左上角与鼠标指针位置间的纵横距离
Private mudtRect As RECT
Private mudtRectMin As RECT
Private mudtRectMax As RECT
Private mudtRectClose As RECT

Private mudtPoint As POINTAPI

Private mblnMoveStart As Boolean '判断移动是否开始
Private mblnMove As Boolean

Private mstrKey As String

Private Enum E_Func
    E_MIN = 0
    E_MAX = 1
    E_CLOSE = 2
    E_NORMAL_1 = 3
    E_NORMAL_2 = 4
End Enum

Public Function OpenChatRoom(ByVal strUrl As String, ByVal strSubject As String, Optional ByVal strSysCode As String, _
   Optional ByVal strMainCode As String, Optional ByVal dblMainId As Double, Optional ByVal strSender As String, _
    Optional ByVal strReceivers As String, Optional ByRef strMsg As String) As Boolean
        '功能:开启讨论房间
          Dim arrTemp As Variant
          
1         On Error GoTo ErrH

2         Me.Caption = strSubject
3         mstrKey = strSysCode & "_" & strMainCode & "_" & dblMainId
4         Me.Show 0
5         If strUrl = "" Then
6             arrTemp = Split(strSender, ",")
7             If UBound(arrTemp) = 1 Then strSender = URLEncode(arrTemp(0)) & "," & arrTemp(1)
8             arrTemp = Split(strReceivers, ",")
9             If UBound(arrTemp) = 1 Then strReceivers = URLEncode(arrTemp(0)) & "," & arrTemp(1)
              
10            strUrl = gstrChatURL & "?system=" & strSysCode & "&maincode=" & URLEncode(strMainCode) & _
                         "&mainid=" & dblMainId & "&subject=" & URLEncode(strSubject) & "&name=" & strSender & _
                         "&join=" & strReceivers
11        End If
12        WriteLog "发起讨论URL：" & strUrl & vbNewLine
13        Call SBrowser.LoadURL(strUrl)
14        OpenChatRoom = True

15        Exit Function

ErrH:
16        strMsg = vbExclamation & "[,]" & "在zlChatNotify.frmChat.OpenChatRoom的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description

End Function

Public Function ShowMyChat(ByVal lngUserId As Long, ByRef strMsg As String) As Boolean
'功能:打开我参与的讨论列表
          Dim strUrl As String
          
1         On Error GoTo ErrH
2         mstrKey = "K_USER_" & lngUserId
3         Me.Caption = ""
4         Me.Show 0
5         strUrl = gstrMyChatUrl & "?uid=" & lngUserId
6         WriteLog "我的讨论URL：" & strUrl & vbNewLine
7         Call SBrowser.LoadURL(strUrl)
8         ShowMyChat = True
9         Exit Function

ErrH:
10        strMsg = vbExclamation & "[,]" & "在zlChatNotify.frmChat.ShowMyChat的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    lblBtn(3).Visible = False
    lblBtn(4).Visible = False
'    SetWindowsInTaskBar Me.hWnd, True
End Sub

Private Sub Form_Resize()
    Dim lngSplit As Long
    
    On Error Resume Next
    lngSplit = 30
    picTop.Move lngSplit, lngSplit, Me.ScaleWidth - 2 * lngSplit, 500 + lngSplit
    picTop.BackColor = conCOLOR_TITLE_BAR
    picResize.BackColor = conCOLOR_TITLE_BAR
    SBrowser.Move lngSplit, picTop.Height + lngSplit, Me.ScaleWidth - lngSplit * 2, Me.ScaleHeight - picTop.Height - lngSplit * 2
    'Left
    With fraScope(0)
        .Left = 0: .Top = 0: .Height = Me.ScaleHeight: .Width = lngSplit
        .BackColor = conCOLOR_TITLE_BAR
    End With
    'bottom
    With fraScope(1)
        .Left = 0: .Top = Me.ScaleHeight - lngSplit: .Height = lngSplit: .Width = Me.ScaleWidth
        .BackColor = conCOLOR_TITLE_BAR
    End With
    'right
    With fraScope(2)
        .Left = Me.ScaleWidth - lngSplit: .Top = 0: .Height = Me.ScaleHeight: .Width = lngSplit
        .BackColor = conCOLOR_TITLE_BAR
    End With
    'Top
    With fraScope(3)
        .Top = 0: .Left = lngSplit: .Height = lngSplit: .Width = Me.ScaleWidth - 2 * lngSplit
        .BackColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objChat As frmChat
    
    On Error Resume Next
    Set objChat = gcolChat(mstrKey)
    Call gcolChat.Remove(mstrKey)
    On Error GoTo 0
    If Not objChat Is Nothing Then
        Unload objChat
        Set objChat = Nothing
    End If
End Sub

Private Sub fraScope_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMinVal As Long
    Dim lngMinW As Long
    Dim lngMinH As Long
    
    If Button = 1 Then
        lngMinVal = 240: lngMinW = 13000: lngMinH = 8000
        Select Case Index
        Case 0 'left
            If Me.Width - X < lngMinW Then Exit Sub
            If Me.Left + X < lngMinVal Or Me.Width - X < lngMinVal Then Exit Sub
            Me.Left = Me.Left + X
            Me.Width = Me.Width - X
        Case 1 'bottom
            If Me.Height + Y < lngMinH Then Exit Sub
            If Me.Height + Y < lngMinVal Then Exit Sub
            Me.Height = Me.Height + Y
        Case 2 'right
            If Me.Width + X < lngMinW Then Exit Sub
            If Me.Width + X < lngMinVal Then Exit Sub
            Me.Width = Me.Width + X
        Case 3 'top
            If Me.Height - Y < lngMinH Then Exit Sub
            If Me.Top + Y < lngMinVal Or Me.Height - Y < lngMinVal Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        End Select
        '重新获取窗体位置
        Call GetWindowRect(picTop.hWnd, mudtRect)
        Call GetWindowRect(picBtn(E_MIN).hWnd, mudtRectMin)
        Call GetWindowRect(picBtn(E_MAX).hWnd, mudtRectMax)
        Call GetWindowRect(picBtn(E_CLOSE).hWnd, mudtRectClose)
    End If
End Sub

Private Sub lblBtn_Click(Index As Integer)
    Select Case Index
    Case E_MIN
        Me.WindowState = vbMinimized
    Case E_MAX, E_NORMAL_1, E_NORMAL_2
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
            lblBtn(E_MAX).Visible = False
            lblBtn(E_NORMAL_1).Visible = True
            lblBtn(E_NORMAL_2).Visible = True
        Else
            Me.WindowState = vbNormal
            lblBtn(E_MAX).Visible = True
            lblBtn(E_NORMAL_1).Visible = False
            lblBtn(E_NORMAL_2).Visible = False
        End If
    Case E_CLOSE
        Unload Me
    End Select
End Sub

Private Sub picBtn_Click(Index As Integer)
    Call lblBtn_Click(Index)
End Sub

Private Sub picBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GetWindowRect(picBtn(E_MIN).hWnd, mudtRectMin)
    Call GetWindowRect(picBtn(E_MAX).hWnd, mudtRectMax)
    Call GetWindowRect(picBtn(E_CLOSE).hWnd, mudtRectClose)
End Sub

Private Sub picBtn_Resize(Index As Integer)
    On Error Resume Next
    lblBtn(Index).Move (picBtn(Index).ScaleWidth - lblBtn(Index).Width) / 2, (picBtn(Index).ScaleHeight - lblBtn(Index).Height) / 2
    If Index = E_MAX Then
        lblBtn(3).Move 105, 200
        lblBtn(4).Move 180, 135
    End If
End Sub

Private Sub picResize_Resize()
    On Error Resume Next
    
    picBtn(E_MIN).Move 10, 10, 480, 480
    picBtn(E_MAX).Move picBtn(E_MIN).Left + picBtn(E_MIN).Width + 30, 10, 480, 480
    picBtn(E_CLOSE).Move picBtn(E_MAX).Left + picBtn(E_MAX).Width + 30, 10, 480, 480

    Call picBtn_Resize(E_MIN)
    Call picBtn_Resize(E_MAX)
    Call picBtn_Resize(E_CLOSE)
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMove Then
        mMoveX = mudtPoint.X - mudtRect.Left
        mMoveY = mudtPoint.Y - mudtRect.Top
        mblnMoveStart = True
    End If
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngX As Long, lngY As Long
    Dim lngLeft As Long, lngTop As Long
    
    If mblnMoveStart Then
        lngX = (mudtPoint.X - mMoveX)
        lngY = (mudtPoint.Y - mMoveY)
        lngLeft = lngX * Screen.TwipsPerPixelX
        lngTop = lngY * Screen.TwipsPerPixelY
        If lngLeft < 0 Then lngLeft = 0
        If lngLeft + Me.Width > Screen.Width Then lngLeft = Screen.Width - Me.Width
        If lngTop < 0 Then lngTop = 0
        If lngTop + Me.Height > Screen.Height Then lngTop = Screen.Height - Me.Height
        Me.Left = lngLeft
        Me.Top = lngTop
    End If
End Sub

Private Sub picTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GetWindowRect(picTop.hWnd, mudtRect)
    Call GetWindowRect(picBtn(E_MIN).hWnd, mudtRectMin)
    Call GetWindowRect(picBtn(E_MAX).hWnd, mudtRectMax)
    Call GetWindowRect(picBtn(E_CLOSE).hWnd, mudtRectClose)
    mblnMoveStart = False
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    picResize.Move picTop.ScaleWidth - 1530, 0, 1500, 480
    picResize.BackColor = picTop.BackColor
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
End Sub
 
Private Sub tmrTime_Timer()
    Dim lngRet As Long
    If tmrTime.Tag = "" Then
        Call GetWindowRect(picTop.hWnd, mudtRect)
        Call GetWindowRect(picBtn(E_MIN).hWnd, mudtRectMin)
        Call GetWindowRect(picBtn(E_MAX).hWnd, mudtRectMax)
        Call GetWindowRect(picBtn(E_CLOSE).hWnd, mudtRectClose)
        tmrTime.Tag = "1" '首次记录窗体位置
    End If
    lngRet = GetCursorPos(mudtPoint)
    '判断鼠标指针是否位于窗体拖动区
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If PtInRect(mudtRectMin, mudtPoint.X, mudtPoint.Y) Then
        picBtn(E_MIN).BackColor = "&H" & Hex(RGB(212, 64, 39))   '红色
    Else
        picBtn(E_MIN).BackColor = picTop.BackColor
    End If
    If PtInRect(mudtRectMax, mudtPoint.X, mudtPoint.Y) Then
        picBtn(E_MAX).BackColor = "&H" & Hex(RGB(212, 64, 39))   '红色
    Else
        picBtn(E_MAX).BackColor = picTop.BackColor
    End If
    If PtInRect(mudtRectClose, mudtPoint.X, mudtPoint.Y) Then
        picBtn(E_CLOSE).BackColor = "&H" & Hex(RGB(212, 64, 39))  '红色
    Else
        picBtn(E_CLOSE).BackColor = picTop.BackColor
    End If
End Sub
