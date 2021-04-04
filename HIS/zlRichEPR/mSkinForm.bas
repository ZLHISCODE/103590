Attribute VB_Name = "mSkinForm"
'#########################################################################
'##模 块 名：mSkinForm.bas
'##创 建 人：吴庆伟
'##日    期：2004年3月1日
'##修 改 人：
'##日    期：
'##描    述：异形窗体创建函数
'#########################################################################

Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const HTCAPTION = 2

'################################################################################################################
'## 功能：  根据图片内容设置窗体为异形窗体
'##
'## 参数：  picSkin         :图片控件
'##
'## 说明：  透明色为图片左上角（0，0）的颜色值。
'################################################################################################################
Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function


