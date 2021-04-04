Attribute VB_Name = "mTips"
Option Explicit
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Const SWP_SHOWWINDOW = &H40&

Public Const SIDESPACE As Long = 8     '边距
Public m_WndStopoverTimeVal As Long    '窗体显示多长时间然后自动隐藏(默认为5秒)

Public Sub ShowTipInfor(ByVal strText As String, _
    Optional ByVal X1 As Long, Optional ByVal Y1 As Long, Optional ByVal lW As Long)
    
    Dim lpPoint As POINTAPI
    Dim lpRect As RECT
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lLeft As Long
    Dim lTop As Long

    frmTips.Hide
    frmTips.tmrControl.Enabled = True
    m_WndStopoverTimeVal = 0
    frmTips.Cls
    frmTips.FontName = "Arial"
    frmTips.FontSize = 8
    frmTips.FontBold = False
    frmTips.FontItalic = False
    frmTips.ForeColor = RGB(0, 0, 0)
    
    frmTips.imgArrow(0).Visible = False
    frmTips.imgArrow(1).Visible = False
    frmTips.imgArrow(2).Visible = False
    frmTips.imgArrow(3).Visible = False
    
    ' 设置窗体的高度和宽度
    frmTips.Height = (frmTips.TextHeight(strText) + 16) * Screen.TwipsPerPixelY
    frmTips.Width = (frmTips.TextWidth(strText) + 40) * Screen.TwipsPerPixelX
    If frmTips.Width > 5000 Then frmTips.Width = 5000
    
    If X1 = 0 And X1 = 0 And lW = 0 Then
        GetCursorPos lpPoint
    Else
        lpPoint.X = X1
        lpPoint.Y = Y1
    End If
    lWidth = Screen.Width / Screen.TwipsPerPixelX
    lHeight = Screen.Height / Screen.TwipsPerPixelY
      
    If lpPoint.Y <= lHeight / 2 Then '上边区域
       If ((lpPoint.X < frmTips.ScaleWidth And lpPoint.Y < frmTips.ScaleHeight)) Or _
          ((lpPoint.X < frmTips.ScaleWidth)) Then '左上边
           lLeft = lpPoint.X + 16
           lTop = lpPoint.Y + 16
           frmTips.Line (0, 0)-(16, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (1, 1)-(16 - 1, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
           frmTips.Line (16 + 1, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
           frmTips.imgArrow(1).Move 4, 4
           frmTips.imgArrow(1).Visible = True
           '显示文字
           SetRect lpRect, 16 + 8, 8, frmTips.ScaleWidth - 9, frmTips.ScaleHeight - 9
           DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       ElseIf ((lWidth - lpPoint.X) < frmTips.ScaleWidth And lpPoint.Y < frmTips.ScaleHeight) Or _
          ((lWidth - lpPoint.X) < frmTips.ScaleWidth) Then '右上边
           lLeft = (lpPoint.X - frmTips.ScaleWidth) - lW - 4
           lTop = lpPoint.Y
           frmTips.Line (frmTips.ScaleWidth - (16 + 1), 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (frmTips.ScaleWidth - 16, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
           frmTips.Line (1, 1)-(frmTips.ScaleWidth - (16 + 2), frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
           frmTips.imgArrow(2).Move frmTips.ScaleWidth - (16 - 4), 4
           frmTips.imgArrow(2).Visible = True
           '显示文字
           SetRect lpRect, 8, 8, frmTips.ScaleWidth - (16 + 9), frmTips.ScaleHeight - 9
           DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       Else
            lLeft = lpPoint.X
            lTop = lpPoint.Y
            frmTips.Line (0, 0)-(16, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
            frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
            frmTips.Line (1, 1)-(16 - 1, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
            frmTips.Line (16 + 1, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
            frmTips.imgArrow(1).Move 4, 4
            frmTips.imgArrow(1).Visible = True
            '显示文字
            SetRect lpRect, 16 + 8, 8, frmTips.ScaleWidth - 9, frmTips.ScaleHeight - 9
            DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       End If
    Else '下边区域
       If ((lpPoint.X < frmTips.ScaleWidth And (lHeight - lpPoint.Y) < frmTips.ScaleHeight)) Or _
          (lpPoint.X < frmTips.ScaleWidth) Then '左下边
           lLeft = lpPoint.X
           lTop = lpPoint.Y - frmTips.ScaleHeight
           frmTips.Line (0, 0)-(16, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (1, 1)-(16 - 1, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
           frmTips.Line (16 + 1, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
           frmTips.imgArrow(0).Move 4, frmTips.ScaleHeight - (16 - 3)
           frmTips.imgArrow(0).Visible = True
           '显示文字
           SetRect lpRect, 16 + 8, 8, frmTips.ScaleWidth - 9, frmTips.ScaleHeight - 9
           DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       ElseIf ((lWidth - lpPoint.X) < frmTips.ScaleWidth And (lHeight - lpPoint.Y) < frmTips.ScaleHeight) Or _
          ((lWidth - lpPoint.X) < frmTips.ScaleWidth) Then '右下边
           lLeft = lpPoint.X - frmTips.ScaleWidth - lW - 4
           lTop = lpPoint.Y - frmTips.ScaleHeight
           frmTips.Line (frmTips.ScaleWidth - (16 + 1), 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (frmTips.ScaleWidth - 16, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
           frmTips.Line (1, 1)-(frmTips.ScaleWidth - (16 + 2), frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
           frmTips.imgArrow(3).Move frmTips.ScaleWidth - (16 - 3), frmTips.ScaleHeight - (16 - 3)
           frmTips.imgArrow(3).Visible = True
           '显示文字
           SetRect lpRect, 8, 8, frmTips.ScaleWidth - (16 + 9), frmTips.ScaleHeight - 9
           DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       Else
           lLeft = lpPoint.X
           lTop = lpPoint.Y - frmTips.ScaleHeight
           frmTips.Line (0, 0)-(16, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (0, 0)-(frmTips.ScaleWidth - 1, frmTips.ScaleHeight - 1), RGB(0, 0, 0), B
           frmTips.Line (1, 1)-(16 - 1, frmTips.ScaleHeight - 2), RGB(255, 210, 83), BF
           frmTips.Line (16 + 1, 1)-(frmTips.ScaleWidth - 2, frmTips.ScaleHeight - 2), RGB(255, 255, 225), BF
           frmTips.imgArrow(0).Move 4, frmTips.ScaleHeight - (16 - 3)
           frmTips.imgArrow(0).Visible = True
           '显示文字
           SetRect lpRect, 16 + 8, 8, frmTips.ScaleWidth - 9, frmTips.ScaleHeight - 9
           DrawText frmTips.hdc, strText, lstrlen(strText), lpRect, DT_LEFT
       End If
    End If
    
'    '显示阴影
'    frmTips.Width = frmTips.Width + 64
'    frmTips.Height = frmTips.Height + 64
'    DrawShadow frmTips.hwnd, frmTips.hdc, lLeft, lTop
    
    SetWindowPos frmTips.hWnd, HWND_TOPMOST, lLeft, lTop, 0, 0, _
                 SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOSIZE
End Sub

Private Sub DrawShadow(ByVal hWnd As Long, ByVal hdc As Long, ByVal xOrg As Long, ByVal yOrg As Long)
    Dim hDcDsk As Long
    Dim Rec As RECT
    Dim winW As Long, winH As Long
    Dim X As Long, Y As Long, c As Long

    GetWindowRect hWnd, Rec
    winW = Rec.right - Rec.left
    winH = Rec.bottom - Rec.top
     
    hDcDsk = GetWindowDC(GetDesktopWindow)
     
    '// Simulate a shadow on right edge...
    For X = 1 To 4
        DoEvents
        For Y = 0 To 3
            c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y)
            SetPixel hdc, winW - X, Y, c
        Next Y
        For Y = 4 To 7
            c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y)
            SetPixel hdc, winW - X, Y, pMask(3 * X * (Y - 3), c)
        Next Y
        For Y = 8 To winH - 5
            c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y)
            SetPixel hdc, winW - X, Y, pMask(15 * X, c)
        Next Y
        For Y = winH - 4 To winH - 1
            c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y)
            SetPixel hdc, winW - X, Y, pMask(3 * X * -(Y - winH), c)
        Next Y
    Next X
     
    '// Simulate a shadow on the bottom edge...
    For Y = 1 To 4
        DoEvents
        For X = 0 To 3
            c = GetPixel(hDcDsk, xOrg + X, yOrg + winH - Y)
            SetPixel hdc, X, winH - Y, c
        Next X
        For X = 4 To 7
            c = GetPixel(hDcDsk, xOrg + X, yOrg + winH - Y)
            SetPixel hdc, X, winH - Y, pMask(3 * (X - 3) * Y, c)
        Next X
        For X = 8 To winW - 5
            c = GetPixel(hDcDsk, xOrg + X, yOrg + winH - Y)
            SetPixel hdc, X, winH - Y, pMask(15 * Y, c)
        Next X
    Next Y
     
    ' - Release the desktop hDC...
    ReleaseDC GetDesktopWindow, hDcDsk

End Sub

'// Function pMask splits a color into its RGB components and transforms the color using a scale 0..255
Private Function pMask(ByVal lScale As Long, ByVal lColor As Long) As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
     
    Long2RGB lColor, R, G, B
     
    R = pTransform(lScale, R)
    G = pTransform(lScale, G)
    B = pTransform(lScale, B)
     
    pMask = RGB(R, G, B)
End Function

'// Function pTransform converts a RGB subcolor using a scale  where 0 = 0 and 255 = lScale
Private Function pTransform(ByVal lScale As Long, ByVal lColor As Long) As Long
    pTransform = lColor - Int(lColor * lScale / 255)
End Function

Private Sub Long2RGB(LongColor As Long, R As Byte, G As Byte, B As Byte)
    On Error Resume Next
    '// convert to hex using vb's hex function, then use the hex2rgb function
    Hex2RGB (Hex$(LongColor)), R, G, B
End Sub

Private Sub Hex2RGB(strHexColor As String, R As Byte, G As Byte, B As Byte)
    On Error Resume Next
    Dim HexColor As String
    Dim i As Byte
   
    '//  make sure the string is 6 characters long
    '// (it may have been given in &H###### format, we want ######)                                                       FixIT90210ae-R9757-R1B8ZE
    strHexColor = right$((strHexColor), 6)
    '// however, it may also have been given as or #***** format, so add 0's in front
    For i = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next
    HexColor = HexColor & strHexColor
    '// convert each set of 2 characters into bytes, using vb's cbyte function
    R = CByte("&H" & right$(HexColor, 2))
    G = CByte("&H" & Mid$(HexColor, 3, 2))
    B = CByte("&H" & left$(HexColor, 2))
End Sub


