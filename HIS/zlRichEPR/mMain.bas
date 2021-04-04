Attribute VB_Name = "mMain"
'#########################################################################
'##模 块 名：mMain.bas
'##创 建 人：吴庆伟
'##日    期：2005年3月25日
'##修 改 人：
'##日    期：
'##描    述：全局变量和类型、函数声明等
'##版    本：
'#########################################################################

Option Explicit

'##########################################################################################
'## 全局类型
'##########################################################################################

Public Type PreDefinedKeyInfo               '保留关键字
    KeyStart As String
    KeyEnd As String
End Type

Public Type KeywordsInfo
    sKeyType As String          'E/P/T/O/C
    lID As Long
    lKSS As Long
    lKSE As Long
    lKES As Long
    lKEE As Long
    bNeeded As Boolean
End Type

'##########################################################################################
'## 全局变量
'##########################################################################################
Public gFileSource As FileSourceEnum

Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstr单位名称 As String
Public gstrSQL As String
Public glngSys As Long


Public gfrm主窗体 As frmMain            '系统主窗体
Public gfrm提纲 As frmStucture          '文档结构图
Public gfrm诊治要素 As frmInsertElement '诊治要素窗口
Public gfrm属性 As frmProperty          '属性窗口

Public gKeyWords(1 To 5) As PreDefinedKeyInfo   '预定义关键字

Public gblnAgentFirstShow As Boolean    '是否第一次显示助手

'##########################################################################################
'## API声明
'##########################################################################################
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetTempFileName Lib "KERNEL32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'绘制透明图形
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'VB Errors
Private Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Private Const DSna = &H220326

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态
    ' Virtual key values
Public Const VK_TAB = &H9

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)  '系统暂停

'##########################################################################################
'## 图片转化为RTF
'##########################################################################################

Private Type Size
    cx As Long
    cy As Long
End Type
' Used to create the metafile
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Used for creating the temporary WMF file
Private Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

'##########################################################################################
'## 全局函数
'##########################################################################################

'Public Sub OpenDB()
'    Set gcnOracle = New ADODB.Connection
'    gcnOracle.Provider = "MSDataShape"
'    gcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=ORCL", "zlepr", "his"
'End Sub
'
'Public Sub CloseDB()
'    gcnOracle.Close
'    Set gcnOracle = Nothing
'End Sub

'#########################################################
'# 名称：NVL( )
'# 功能：相当于Oracle的NVL，将Null值改成另外一个预设值
'#########################################################
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'#########################################################
'# 名称：SetOpacityForm
'# 功能：透明窗体设置
'# 用法：SetOpacityForm(窗体名，透明度)
'#########################################################
Public Sub SetOpacityForm(obj As Object, Opacity As Byte)
    'On Error Resume Next
    Dim ret As Long
    'Set the window style to 'Layered'
    ret = GetWindowLong(obj.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong obj.hWnd, GWL_EXSTYLE, ret
    'Set the opacity of the layered window to 128
    SetLayeredWindowAttributes obj.hWnd, 0, Opacity, LWA_ALPHA
End Sub

'#########################################################
'# 名称：Num2Chi
'# 功能：数字到中文字符
'# 用法：Num2Chi(数值)
'#########################################################
Public Function Num2Chi(lngNum As Long) As String
    Dim strTMP As String, i As Long
    Dim strIn As String
    Dim strOut As String
    strIn = CStr(lngNum)
    For i = 1 To Len(strIn)
        Select Case Mid(strIn, i, 1)
        Case 0
            strTMP = "〇"
        Case 1
            strTMP = "一"
        Case 2
            strTMP = "二"
        Case 3
            strTMP = "三"
        Case 4
            strTMP = "四"
        Case 5
            strTMP = "五"
        Case 6
            strTMP = "六"
        Case 7
            strTMP = "七"
        Case 8
            strTMP = "八"
        Case 9
            strTMP = "九"
        End Select
        strOut = strOut & strTMP
    Next
    Num2Chi = strOut
End Function

'#########################################################
'# 名称：GetTempName
'# 功能：获取WIndows临时目录
'# 用法：GetTempName(strName)
'#########################################################
Public Function GetTempName(TmpFilePrefix As String) As String
     Dim TempFileName As String * 256
     Dim X As Long
     Dim DriveName As String
     DriveName = "c:"
       X = GetTempFileName(DriveName, TmpFilePrefix, 0, TempFileName)
       GetTempName = Left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)
End Function


'###############################################################################################
'   绘制透明图片到指定HDC上（指定透明色）。
'###############################################################################################

Public Sub PaintTransparentStdPic(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal XSrc As Long, _
                                    ByVal YSrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, XSrc, YSrc, clrMask, hPal
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle
            'Draw Transparent image
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub

PaintTransparentStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

'Provided with comments by Microsoft
Public Sub PaintTransparentDC(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    ByVal XSrc As Long, _
                                    ByVal YSrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hBmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long

    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        'Create halftone palette
        hPal = CreateHalftonePalette(hdcScreen)
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy

    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hBmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, XSrc, YSrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hBmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    DeleteObject hPal
End Sub

'判断是否为编辑键
Public Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function


