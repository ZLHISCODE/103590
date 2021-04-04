VERSION 5.00
Begin VB.Form frmTipInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTipInfo.frx":0000
   ScaleHeight     =   1155
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timTip 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1050
      Top             =   105
   End
   Begin VB.Label lblOutline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提纲"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   390
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "提示信息"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   795
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTipInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngHwnd As Long
Private mstrInfo As String
Private mblnShow As Boolean
Private mblnMultiRow As Boolean '是否按多行方式显示信息
Private mblnOutline As Boolean  '是否将每行文本中字符|前的文字做为提纲换行显示
Private mlngMaxWidth As Long    '窗体最大宽度
Private mvOSVer As OSVERSIONINFO
Private Const conOutlineSpit = "|"
Private mblnChildWindow As Boolean  '是否使用ChildWindowFromPoint方法

'---------------------------------------------------------------------------------------------------------------------
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32s = 0 'Win32s on Windows 3.1.
Private Const VER_PLATFORM_WIN32_WINDOWS = 1 'Windows 95, Windows 98, or Windows Me.
Private Const VER_PLATFORM_WIN32_NT = 2 'Windows NT 3.51, Windows NT 4.0, Windows 2000, Windows XP, or Windows .NET Server.

Private Const WM_PAINT = &HF

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80

Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long, Optional strTitle As String, Optional blnChild As Boolean)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      strInfo=提示信息,当传入为空时隐藏提示
'      blnMultiRow=以一定的间距分行显示多行信息，每行按vbcrlf分隔
'      blnOutline=是否将每行文本中字符|前的文字做为提纲单独一行显示
'      lngMaxWidth=窗口的最大窗度，缺省为0表示按设计状态的窗体最大宽度为准
'      blnChild=是否使用ChildWindowFromPoint方法
    
    mlngHwnd = lngHwnd
    mblnMultiRow = blnMultiRow
    mblnOutline = blnOutline
    mlngMaxWidth = lngMaxWidth
    mblnChildWindow = blnChild
    If strTitle <> "" Then
        lblOutline(0).Caption = strTitle
    End If
    
    If strInfo <> mstrInfo Then
        mstrInfo = strInfo
        Call SetMeShape
    End If
    Me.timTip.Enabled = False


    If mlngHwnd <> 0 And mstrInfo <> "" Then
        Me.timTip.Enabled = True
    End If
End Sub

Private Sub timTip_Timer()
    Static vPrePos As POINTAPI
    Static sngBegin As Single
    
    Dim vPos As POINTAPI
    Dim sngTime As Single
    Dim lngHwnd As Long
    
    If mlngHwnd = 0 Or mstrInfo = "" Then
        Me.timTip.Enabled = False
        Call HideMe: Exit Sub
    End If
    
    GetCursorPos vPos
    If mblnChildWindow Then
        ScreenToClient mlngHwnd, vPos
        lngHwnd = ChildWindowFromPoint(mlngHwnd, vPos.x, vPos.y)
    Else
        lngHwnd = WindowFromPoint(vPos.x, vPos.y)
    End If
    
    If lngHwnd <> mlngHwnd Then
        Call HideMe
    Else
        'If vPos.X & "," & vPos.Y <> vPrePos.X & "," & vPrePos.Y Then
        If Abs(vPos.x - vPrePos.x) > 2 Or Abs(vPos.y - vPrePos.y) > 2 Then
            sngBegin = 0
            Call HideMe
        Else
            sngTime = Timer
            If sngBegin = 0 Then sngBegin = sngTime
            
            If sngTime - sngBegin >= 0.2 Then
                Call ShowMe
            End If
        End If
    End If
    vPrePos = vPos
End Sub

Private Sub SetMeShape()
    Dim lngR As Long, arrTxt As Variant, i As Long, lngWidth As Long, lngMaxWidth As Long
    Dim j As Long
    '尺寸
    If mblnMultiRow And mstrInfo <> "" Then
        arrTxt = Split(mstrInfo, vbCrLf)
        lngWidth = Me.TextWidth(arrTxt(0))
        For i = 1 To UBound(arrTxt)
            If lngWidth < Me.TextWidth(arrTxt(i)) Then lngWidth = Me.TextWidth(arrTxt(i))
        Next
        
        lngMaxWidth = IIf(mlngMaxWidth = 0, Val(Me.Tag), mlngMaxWidth)
        If lngWidth + lblInfo(0).Left * 2 <= lngMaxWidth Then
            Me.Width = lngWidth + lblInfo(0).Left * 2
        Else
            Me.Width = lngMaxWidth
        End If
        lblInfo(0).Width = Me.Width - Me.lblInfo(0).Left * 2
        lblInfo(0).Caption = ""
        lblOutline(0).Caption = ""
        For i = 1 To lblInfo.UBound
            Unload lblInfo(i)
        Next
        For i = 1 To lblOutline.UBound
            Unload lblOutline(i)
        Next
        
        If Not (mblnOutline And InStr(arrTxt(0), conOutlineSpit) > 0) Then
            lblOutline(0).Visible = False
            lblInfo(0).Top = lblOutline(0).Top
        End If
        j = 0 '提纲计数器
        For i = 0 To UBound(arrTxt)
            If i > 0 Then
                Load lblInfo(i)
                Set lblInfo(i).Container = Me
                lblInfo(i).Left = lblInfo(0).Left
                lblInfo(i).Top = lblInfo(i - 1).Top + lblInfo(i - 1).Height + Screen.TwipsPerPixelY * 6
                lblInfo(i).Visible = True
            End If
            lblInfo(i).Width = Me.Width - Me.lblInfo(0).Left * 2
            
            If mblnOutline And InStr(arrTxt(i), conOutlineSpit) > 0 Then
                If j > 0 Then
                    Load lblOutline(j)
                    Set lblOutline(j).Container = Me
                    lblOutline(j).Left = lblOutline(0).Left
                    lblOutline(j).Width = Me.Width - Me.lblOutline(0).Left * 2
                    lblOutline(j).Top = lblInfo(i).Top
                    lblInfo(i).Top = lblOutline(j).Top + lblOutline(j).Height + Screen.TwipsPerPixelY * 2
                    lblOutline(j).Visible = True
                End If
                
                lblOutline(j).Caption = Split(arrTxt(i), conOutlineSpit)(0)
                lblInfo(i).Caption = "    " & Split(arrTxt(i), conOutlineSpit)(1)
                j = j + 1
            Else
                lblInfo(i).Caption = arrTxt(i)
            End If
        Next
        Me.Height = lblInfo(UBound(arrTxt)).Top + lblInfo(UBound(arrTxt)).Height + lblInfo(0).Top
    Else
        If Me.TextWidth(mstrInfo) + lblInfo(0).Left * 2 <= Val(Me.Tag) Then
            Me.Width = Me.TextWidth(mstrInfo) + lblInfo(0).Left * 2
        Else
            Me.Width = Val(Me.Tag)
        End If
        Me.lblInfo(0).Width = Me.Width - Me.lblInfo(0).Left * 2
        Me.lblInfo(0).Caption = mstrInfo
        Me.Height = Me.lblInfo(0).Height + Me.lblInfo(0).Top * 2
    End If
    
    '背景
    Me.PaintPicture Me.Picture, 0, 0, Me.Width, Me.Height
    
    '边框：API=RoundRect
    Me.Line (Screen.TwipsPerPixelX, 0)-(Me.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    Me.Line (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (0, Screen.TwipsPerPixelY)-(0, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (Me.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    
    '形状
    lngR = CreateRoundRectRgn(0, 0, Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels) + 1, Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(Me.hwnd, lngR, False)
    
    '初始透明度
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        lngR = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, lngR Or WS_EX_LAYERED)
        SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA
    End If
End Sub

Private Sub ShowMe()
    Dim vPos As POINTAPI
    Dim vRect As RECT, vDesk As RECT
    Dim x As Long, y As Long, H As Long, W As Long
    Dim i As Integer
    
    If mblnShow Then Exit Sub
    W = Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels)
    H = Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels)
    
    GetWindowRect GetDesktopWindow, vDesk
    Call GetCursorPos(vPos)
    vRect.Left = vPos.x
    vRect.Top = vPos.y
    vRect.Right = vRect.Left + GetSystemMetrics(SM_CXCURSOR) / 2
    vRect.Bottom = vRect.Top + GetSystemMetrics(SM_CYCURSOR) / 2
    
    If vRect.Right + W < vDesk.Right Then
        x = vRect.Right
    Else
        x = vRect.Left - W - 1
    End If
    If vRect.Top + H < vDesk.Bottom Then
        y = vRect.Top
    Else
        y = vRect.Bottom - H
    End If
        
    MoveWindow Me.hwnd, x, y, W, H, 0
    
    '显示在最前面且不激活：HWND_TOPMOST,SWP_NOACTIVATE
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    '渐显
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        For i = 0 To 255 Step 6
            SetLayeredWindowAttributes Me.hwnd, 0, i, LWA_ALPHA
            Call SendMessage(Me.hwnd, WM_PAINT, 0, 0) '第一次显示时看不出效果
            Call Sleep(1)
        Next
        SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
        Call SendMessage(Me.hwnd, WM_PAINT, 0, 0)
    End If
    mblnShow = True
End Sub

Private Sub HideMe()
    Dim i As Integer
        
    If Not mblnShow Then Exit Sub
        
    '渐隐
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        For i = 255 To 0 Step -6
            SetLayeredWindowAttributes Me.hwnd, 0, i, LWA_ALPHA
            Call Sleep(1)
        Next
        SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA
    End If
    mblnShow = False
    
    '用VB的Hide会出错
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW
End Sub

Private Sub Form_Load()
    Me.Tag = Me.Width
    
    mvOSVer.dwOSVersionInfoSize = Len(mvOSVer)
    GetVersionEx mvOSVer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    mstrInfo = ""
    mlngHwnd = 0
End Sub

