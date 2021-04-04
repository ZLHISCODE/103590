VERSION 5.00
Begin VB.Form frmTipInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTipInfo.frx":0000
   ScaleHeight     =   1155
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInfo 
      Height          =   150
      Left            =   3915
      Picture         =   "frmTipInfo.frx":12922
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicMsg 
      Height          =   150
      Left            =   3225
      Picture         =   "frmTipInfo.frx":13144
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicInfobak 
      Height          =   150
      Left            =   3975
      Picture         =   "frmTipInfo.frx":13407
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   225
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicEx 
      Height          =   150
      Left            =   3225
      Picture         =   "frmTipInfo.frx":136B8
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer timTip 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1050
      Top             =   105
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ��Ϣ"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   60
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
Private mblnMultiRow As Boolean '�Ƿ񰴶��з�ʽ��ʾ��Ϣ
Private mbInfoType As Byte
Private mvOSVer As OSVERSIONINFO

'---------------------------------------------------------------------------------------------------------------------
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private vPrePos As POINTAPI
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
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional bInfoType As Byte)
'���ܣ���ʾ����������ʾ
'������lngHwnd=��ʾ����ԵĿؼ����,������Ϊ0ʱ������ʾ
'      strInfo=��ʾ��Ϣ,������Ϊ��ʱ������ʾ
'      blnMultiRow=��һ���ļ�������ʾ������Ϣ��ÿ�а�vbcrlf�ָ�
'      bInfoType =0 ��ɫ��Ϣ��=1��ɫ��Ϣ =2 ��ɫ��Ϣ
    GetCursorPos vPrePos
    mlngHwnd = lngHwnd
    mblnMultiRow = blnMultiRow
    mbInfoType = bInfoType
    If strInfo <> mstrInfo Then
        mstrInfo = strInfo
        Call SetMeShape
    End If
     Call ShowMe
    
    If mlngHwnd <> 0 And mstrInfo <> "" Then
        Me.timTip.Enabled = True
    End If
End Sub

Private Sub timTip_Timer()
    Dim vPos As POINTAPI
    
    If mlngHwnd = 0 Or mstrInfo = "" Then
        Me.timTip.Enabled = False
        Call Unload(Me): Exit Sub
    End If
    
    GetCursorPos vPos
    If Abs(vPos.x - vPrePos.x) > 20 Or Abs(vPos.y - vPrePos.y) > 20 Then
        Unload Me
    End If
End Sub

Private Sub SetMeShape()
    Dim lngR As Long, arrTxt As Variant, i As Long, lngWidth As Long
    
    '�ߴ�
    If mblnMultiRow And mstrInfo <> "" Then
        arrTxt = Split(mstrInfo, vbCrLf)
        lngWidth = Me.TextWidth(arrTxt(0))
        For i = 1 To UBound(arrTxt)
            If lngWidth < Me.TextWidth(arrTxt(i)) Then lngWidth = Me.TextWidth(arrTxt(i))
        Next
        If lngWidth + lblInfo(0).Left * 2 <= Val(Me.Tag) Then
            Me.Width = lngWidth + lblInfo(0).Left * 2
        Else
            Me.Width = Val(Me.Tag)
        End If
        lblInfo(0).Caption = ""
        For i = 1 To lblInfo.UBound
            Unload lblInfo(i)
        Next
        For i = 0 To UBound(arrTxt)
            If i > 0 Then
                Load lblInfo(i)
                Set lblInfo(i).Container = Me
                lblInfo(i).Left = lblInfo(0).Left
                lblInfo(i).Top = lblInfo(i - 1).Top + lblInfo(i - 1).Height + Screen.TwipsPerPixelY * 4
                lblInfo(i).Visible = True
            End If
            lblInfo(i).Width = Me.Width - Me.lblInfo(0).Left * 2
            lblInfo(i).Caption = arrTxt(i)
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
    
    '����
    Select Case mbInfoType
        Case 0
            Me.PaintPicture PicInfo.Picture, 0, 0, Me.Width, Me.Height
        Case 1
            Me.PaintPicture PicMsg.Picture, 0, 0, Me.Width, Me.Height
        Case 2
            Me.PaintPicture PicEx.Picture, 0, 0, Me.Width, Me.Height
    End Select
    
    '�߿�API=RoundRect
    Me.Line (Screen.TwipsPerPixelX, 0)-(Me.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    Me.Line (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (0, Screen.TwipsPerPixelY)-(0, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (Me.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    
    '��״
    lngR = CreateRoundRectRgn(0, 0, Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels) + 1, Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(Me.hWnd, lngR, False)
    
    '��ʼ͸����
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        lngR = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
        Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, lngR Or WS_EX_LAYERED)
        SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
    End If
End Sub

Private Sub ShowMe()
    Dim vPos As POINTAPI
    Dim vRect As RECT, vDesk As RECT, vdRect As RECT
    Dim x As Long, y As Long, H As Long, W As Long
    Dim i As Integer
    
    If mblnShow Then Exit Sub
    
    W = Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels)
    H = Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels)
    
    GetWindowRect GetDesktopWindow, vDesk
    GetWindowRect mlngHwnd, vdRect
''    Call GetCursorPos(vPos)
    vRect.Left = vdRect.Left + 80
    vRect.Top = vdRect.Top
'    vRect.Left = vdRect.Right
'    vRect.Top = vdRect.Top
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
        
    MoveWindow Me.hWnd, x, y, W, H, 0
    
    '��ʾ����ǰ���Ҳ����HWND_TOPMOST,SWP_NOACTIVATE
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    '����
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        For i = 0 To 255 Step 2
            SetLayeredWindowAttributes Me.hWnd, 0, i, LWA_ALPHA
            Call SendMessage(Me.hWnd, WM_PAINT, 0, 0) '��һ����ʾʱ������Ч��
            Call Sleep(1)
        Next
        SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
        Call SendMessage(Me.hWnd, WM_PAINT, 0, 0)
    End If
    mblnShow = True
End Sub

Private Sub HideMe()
    Dim i As Integer
        
    If Not mblnShow Then Exit Sub
        
    '����
    If mvOSVer.dwPlatformId >= VER_PLATFORM_WIN32_NT And mvOSVer.dwMajorVersion >= 5 Then
        For i = 255 To 0 Step -2
            SetLayeredWindowAttributes Me.hWnd, 0, i, LWA_ALPHA
            Call Sleep(1)
        Next
        SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
    End If
    mblnShow = False
    
    '��VB��Hide�����
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW
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

