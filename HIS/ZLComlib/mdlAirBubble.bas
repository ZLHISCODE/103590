Attribute VB_Name = "mdlAirBubble"
Option Explicit

Public Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function FindWindow Lib "user32 " Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClientRect Lib "user32 " (ByVal hwnd As Long, lpRect As RECT) As Long

Public mvOSVer As OSVERSIONINFO
Public Const VER_PLATFORM_WIN32_NT = 2 'Windows NT 3.51, Windows NT 4.0, Windows 2000, Windows XP, or Windows .NET Server.
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&

Public Sub DrawColorToColor(picDraw As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal blnVertical As Boolean = True, Optional ByVal blnBorder As Boolean = False)
'画出从一种颜色到另一种颜色的渐变
    Dim VR, VG, VB As Single
    Dim R, G, b, R2, G2, B2 As Integer
    Dim temp As Long, Y As Long, X As Long
    Dim tmpMode As Long
    Dim blnAutoRedraw As Boolean
    
    '只有窗体和图片可以画
    If Not (TypeOf picDraw Is PictureBox Or TypeOf picDraw Is Form) Then Exit Sub
    
    
    tmpMode = picDraw.ScaleMode
    blnAutoRedraw = picDraw.AutoRedraw
    
    picDraw.ScaleMode = 3
    picDraw.AutoRedraw = True
    
    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    If blnVertical Then
        VR = Abs(R - R2) / picDraw.ScaleHeight
        VG = Abs(G - G2) / picDraw.ScaleHeight
        VB = Abs(b - B2) / picDraw.ScaleHeight
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For Y = 0 To picDraw.ScaleHeight
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = b + VB * Y
            picDraw.Line (0, Y)-(picDraw.ScaleWidth, Y), RGB(R2, G2, B2)
        Next Y
    Else
        VR = Abs(R - R2) / picDraw.ScaleWidth
        VG = Abs(G - G2) / picDraw.ScaleWidth
        VB = Abs(b - B2) / picDraw.ScaleWidth
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For X = 0 To picDraw.ScaleWidth
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X
            picDraw.Line (X, 0)-(X, picDraw.ScaleHeight), RGB(R2, G2, B2)
        Next X
    End If
    
    If blnBorder Then
        picDraw.DrawWidth = 2
        picDraw.Line (1, 1)-(picDraw.ScaleWidth - 1, picDraw.ScaleHeight - 1), &HC000&, B
        picDraw.DrawWidth = 1
    End If
    
    picDraw.Refresh
    picDraw.ScaleMode = tmpMode
    picDraw.AutoRedraw = blnAutoRedraw
End Sub

Public Function PrintContent(ByVal objMain As Object, ByVal strCaption As String, Optional lngLeftGap As Long, Optional lngRightGap As Long, Optional lngRowGap As Long)
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim i As Integer
    Dim strChar As String
    Dim strFlag As String
    Dim strTmp As String
    Dim strArr() As String
    Dim intCount As Integer
    Dim lngCount As Long
    
    For i = 1 To Len(strCaption)
        strChar = Mid(strCaption, i, 1)
        If strChar = Chr(13) Or strChar = Chr(10) Then
            If strTmp <> "" Then
                strFlag = strFlag & "|" & strTmp
                strTmp = ""
            Else
                strFlag = strFlag & "|" & ""
            End If
            i = i + 1
        Else
            lngWidth = objMain.TextWidth(strTmp & strChar)
            If lngWidth > objMain.Width - lngLeftGap - lngRightGap Then
                strFlag = strFlag & "|" & strTmp
                strTmp = strChar
            Else
                strTmp = strTmp & strChar
            End If
        End If
    Next
    If strTmp <> "" Then
        strFlag = strFlag & "|" & strTmp
        strTmp = ""
    End If
    If strFlag <> "" Then
        strFlag = Mid(strFlag, 2)
        strArr = Split(strFlag, "|")
        intCount = UBound(strArr) + 1
        For lngCount = 1 To intCount
            lngHeight = objMain.TextHeight(strArr(lngCount - 1))
            lngWidth = objMain.TextWidth(strArr(lngCount - 1))
            'objMain.CurrentX = lngLeftGap
            objMain.CurrentX = (objMain.ScaleWidth - lngWidth) / 2
            If lngCount = 1 Then
                objMain.CurrentY = (objMain.Height - (lngHeight + lngRowGap) * intCount + lngRowGap) / 2
            Else
                objMain.CurrentY = objMain.CurrentY + lngRowGap
            End If
            objMain.Print strArr(lngCount - 1)
        Next
    End If
End Function
