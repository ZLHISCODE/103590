VERSION 5.00
Begin VB.Form frmPatiInfo 
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
   Picture         =   "frmPatiInfo.frx":0000
   ScaleHeight     =   1155
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timTip 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1050
      Top             =   105
   End
   Begin VB.Line linOut 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   165
      X2              =   1680
      Y1              =   675
      Y2              =   690
   End
   Begin VB.Label lblOutline 
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
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
Attribute VB_Name = "frmPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngHwnd As Long
Private mrsInfo As ADODB.Recordset
Private mrsInfoCopy As ADODB.Recordset
Private mblnShow As Boolean
Private mlngMaxWidth As Long  '窗体最大宽度
Private mlngMinWidth As Long   '窗体最小宽度
Private mvOSVer As OSVERSIONINFO
'---------------------------------------------------------------------------------------------------------------------
Private Type POINTAPI
        X As Long
        Y As Long
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
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal rsValue As ADODB.Recordset, Optional lngMinWidth As Long, Optional lngMaxWidth As Long)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      rsValue=包含：id、parent_id（不为空则表明为提纲，为空则为内容）、content（文本内容）、color（文本颜色）
'      lngMinWidth=窗口的最小宽度，缺省为0，不为0表示按照内容扩展的最小宽度不能小于此宽度
'      lngMaxWidth=窗口的最大窗度，缺省为0表示按设计状态的窗体最大宽度为准
    Dim blnNotRecNothing As Boolean, blnCompare As Boolean
    
    mlngHwnd = lngHwnd
    mlngMaxWidth = lngMaxWidth
    mlngMinWidth = lngMinWidth
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = adStateOpen Then blnNotRecNothing = True
    End If
    If blnNotRecNothing = True Then
        mrsInfo.Filter = ""
        blnCompare = gobjComlib.Rec.Compare(rsValue, mrsInfo)
    Else
        blnCompare = False
    End If
    If blnCompare = False Then
        rsValue.Filter = ""
        Set mrsInfo = rsValue
        Set mrsInfoCopy = gobjComlib.Rec.CopyNew(rsValue)
        Call SetMeShape
    End If
    
    If mlngHwnd <> 0 And mrsInfo.RecordCount > 0 Then
        Me.timTip.Enabled = True
    End If
End Sub

Private Sub timTip_Timer()
    Static vPrePos As POINTAPI
    Static sngBegin As Single
    
    Dim vPos As POINTAPI
    Dim sngTime As Single
    Dim lngHwnd As Long
    
    If mlngHwnd = 0 Or mrsInfo.RecordCount = 0 Then
        Me.timTip.Enabled = False
        Call HideMe: Exit Sub
    End If
    
    GetCursorPos vPos
    lngHwnd = WindowFromPoint(vPos.X, vPos.Y)
    If lngHwnd <> mlngHwnd Then
        Call HideMe
    Else
        If Abs(vPos.X - vPrePos.X) > 2 Or Abs(vPos.Y - vPrePos.Y) > 2 Then
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
    Dim lngR As Long, i As Long, lngWidth As Long, lngMaxWidth As Long
    Dim j As Long
    Dim lngFirstID As Long, lngParentId As Long
    '尺寸
    If mrsInfo.RecordCount > 0 Then
        mrsInfo.MoveFirst
        mrsInfo.Sort = "id"
        lngFirstID = mrsInfo.fields("id").Value
        Do While Not mrsInfo.EOF
            If lngWidth < Me.TextWidth(mrsInfo.fields("content").Value) Then lngWidth = Me.TextWidth(mrsInfo.fields("content").Value)
            mrsInfo.MoveNext
        Loop
        
        lngMaxWidth = IIf(mlngMaxWidth = 0, Val(Me.Tag), mlngMaxWidth)
        If lngWidth + lblInfo(0).Left * 2 <= lngMaxWidth Then
            Me.Width = lngWidth + lblInfo(0).Left * 2
        Else
            Me.Width = lngMaxWidth
        End If
        If mlngMinWidth > 0 And Me.Width < mlngMinWidth Then Me.Width = mlngMinWidth
        
        lblInfo(0).Width = Me.Width - Me.lblInfo(0).Left * 2
        lblInfo(0).Caption = ""
        lblOutline(0).Caption = ""
        For i = 1 To lblInfo.UBound
            Unload lblInfo(i)
        Next
        For i = 1 To lblOutline.UBound
            Unload lblOutline(i)
        Next
        For i = 1 To linOut.UBound
            Unload linOut(i)
        Next
        mrsInfo.Filter = "parent_id=" & lngFirstID
        If Not (mrsInfo.RecordCount > 0) Then
            lblOutline(0).Visible = False
            linOut(0).Visible = False
            lblInfo(0).Top = lblOutline(0).Top
            j = 1 '提纲计数器
        Else
            j = 0 '提纲计数器
        End If
        mrsInfo.Filter = "parent_id>0"
        mrsInfo.Sort = "id"
        lngParentId = 0
        For i = 0 To mrsInfo.RecordCount - 1
            If i > 0 Then
                Load lblInfo(i)
                Set lblInfo(i).Container = Me
                lblInfo(i).Left = lblInfo(0).Left
                lblInfo(i).Top = lblInfo(i - 1).Top + lblInfo(i - 1).Height + Screen.TwipsPerPixelY * 6
                lblInfo(i).Visible = True
            End If
            lblInfo(i).Width = Me.Width - Me.lblInfo(0).Left * 2
            
            mrsInfoCopy.Filter = "id=" & Val(mrsInfo.fields("parent_id").Value)
            If mrsInfoCopy.RecordCount > 0 And lngParentId <> Val(mrsInfo.fields("parent_id").Value) Then
                If j > 0 Then
                    Load linOut(j)
                    Set linOut(j).Container = Me
                    linOut(j).X1 = lblOutline(0).Left
                    linOut(j).X2 = Me.Width - Me.lblOutline(0).Left
                    linOut(j).Y1 = lblInfo(i).Top
                    linOut(j).Y2 = linOut(j).Y1
                    linOut(j).Visible = True
                    
                    Load lblOutline(j)
                    Set lblOutline(j).Container = Me
                    lblOutline(j).Left = lblOutline(0).Left
                    lblOutline(j).Width = Me.Width - Me.lblOutline(0).Left * 2
                    lblOutline(j).Top = lblInfo(i).Top + Screen.TwipsPerPixelY * 6
                    lblInfo(i).Top = lblOutline(j).Top + lblOutline(j).Height + Screen.TwipsPerPixelY * 2
                    lblOutline(j).Visible = True
                End If
                
                lblOutline(j).Caption = mrsInfoCopy.fields("content").Value & ""
                lblOutline(j).ForeColor = Val("" & mrsInfoCopy.fields("color").Value)
                
                If j = 0 Then
                     lblOutline(j).Width = Me.Width - Me.lblOutline(0).Left * 2
                End If
                lblInfo(i).Caption = mrsInfo.fields("content").Value & ""
                lblInfo(i).ForeColor = Val("" & mrsInfo.fields("color").Value)
                lngParentId = Val(mrsInfo.fields("parent_id").Value)
                j = j + 1
            Else
                lblInfo(i).Caption = mrsInfo.fields("content").Value & ""
                lblInfo(i).ForeColor = Val("" & mrsInfo.fields("color").Value)
            End If
            mrsInfo.MoveNext
        Next
        Me.Height = lblInfo(lblInfo.Count - 1).Top + lblInfo(lblInfo.Count - 1).Height + IIf(lblOutline(0).Visible = True, lblOutline(0).Top, lblInfo(0).Top)
    Else
        If Me.TextWidth("") + lblInfo(0).Left * 2 <= Val(Me.Tag) Then
            Me.Width = Me.TextWidth("") + lblInfo(0).Left * 2
        Else
            Me.Width = Val(Me.Tag)
        End If
        Me.lblInfo(0).Width = Me.Width - Me.lblInfo(0).Left * 2
        Me.lblInfo(0).Caption = ""
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
    Dim X As Long, Y As Long, H As Long, W As Long
    Dim i As Integer
    
    If mblnShow Then Exit Sub
    
    W = Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels)
    H = Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels)
    
    GetWindowRect GetDesktopWindow, vDesk
    Call GetCursorPos(vPos)
    vRect.Left = vPos.X
    vRect.Top = vPos.Y
    vRect.Right = vRect.Left + GetSystemMetrics(SM_CXCURSOR) / 2
    vRect.Bottom = vRect.Top + GetSystemMetrics(SM_CYCURSOR) / 2
    
    If vRect.Right + W < vDesk.Right Then
        X = vRect.Right
    Else
        X = vRect.Left - W - 1
    End If
    If vRect.Top + H < vDesk.Bottom Then
        Y = vRect.Top
    Else
        Y = vRect.Bottom - H
    End If
        
    MoveWindow Me.hwnd, X, Y, W, H, 0
    
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
    Set mrsInfo = Nothing
    Set mrsInfoCopy = Nothing
    mlngHwnd = 0
End Sub
