Attribute VB_Name = "mPublic"
'######################################################################################
'##描    述：公共函数或者常量声明
'######################################################################################

Option Explicit

'文本对齐方式
Public Enum ECGTextAlignFlags
   DT_TOP = &H0&                    '靠上
   DT_LEFT = &H0&                   '靠左
   DT_CENTER = &H1&                 '居中
   DT_RIGHT = &H2&                  '靠右
   DT_VCENTER = &H4&                '垂直居中
   DT_BOTTOM = &H8&                 '靠下
   DT_WORDBREAK = &H10&             '换行
   DT_SINGLELINE = &H20&            '单行
   DT_EXPANDTABS = &H40&            '扩展制表位
   DT_TABSTOP = &H80&               '制表位
   DT_NOCLIP = &H100&               '不裁剪，稍快
   DT_EXTERNALLEADING = &H200&      '包含字体前导高度
   DT_CALCRECT = &H400&             '计算高度、宽度，而不绘制文本
   DT_NOPREFIX = &H800&             '不处理前缀字符
   DT_INTERNAL = &H1000&            '利用系统字体来计算字体规格
'#if(WINVER >= =&H0400)
   DT_EDITCONTROL = &H2000&         '编辑控件特性，不会出现末行显示半行的情况
   DT_PATH_ELLIPSIS = &H4000&       '文字太长时，中间显示省略号
   DT_END_ELLIPSIS = &H8000&        '文字太长时，末尾显示省略号
   DT_MODIFYSTRING = &H10000        '编辑指定文本来适应显示文本
   DT_RTLREADING = &H20000          '从右到左阅读
   DT_WORD_ELLIPSIS = &H40000       '单词长度太长时，加上省略号
End Enum
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type
Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, LParam As Any) As Long

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, LParam As Long) As Long
    Dim FaceName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    frmProperty.cmbFontName.AddItem left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function

'#########################################################################
'扩展的 Shell 命令
Public Function ShellEx( _
        ByVal sFile As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        eShowCmd = 0    '隐藏
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then   'Shell 调用
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "内存溢出"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "文件没有找到"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "路径没有找到"
        Case ERROR_BAD_FORMAT
            sErr = "无效的可执行文件或者已经损坏"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "路径/文件存取错误"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "该文件没有有效的文件关联"
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "文件无法打开，目标程序忙！请稍后再试。"
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "文件无法打开，DDE传输忙！请稍后再试。"
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "文件无法打开，超时！请稍后再试。"
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "没有找到指定的动态链接库。"
        Case SE_ERR_FNF
            lErr = 53: sErr = "文件没有找到。"
        Case SE_ERR_NOASSOC
            sErr = "没有与之关联的应用程序。"
        Case SE_ERR_OOM
            lErr = 7: sErr = "内存溢出"
        Case SE_ERR_PNF
            lErr = 76: sErr = "路径没有找到"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "共享违例"
        Case Else
            sErr = "在打开或者打印该文件时发生错误。"
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function

'获取Shift按键状态
Public Function giGetShiftState() As Integer
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
    iR = iR Or (-vbShiftMask * gbKeyIsPressed(VK_SHIFT))
    iR = iR Or (-vbAltMask * gbKeyIsPressed(VK_MENU))
    iR = iR Or (-vbCtrlMask * gbKeyIsPressed(VK_CONTROL))
    giGetShiftState = iR

End Function

'获取鼠标按键状态
Public Function giGetMouseButton() As Integer
Dim iR As Integer
   iR = iR Or (-vbLeftButton * gbKeyIsPressed(vbKeyLButton))
   iR = iR Or (-vbMiddleButton * gbKeyIsPressed(vbKeyMButton))
   iR = iR Or (-vbRightButton * gbKeyIsPressed(vbKeyRButton))
   giGetMouseButton = iR
   
End Function

'判断某个键是否被按下
Public Function gbKeyIsPressed( _
        ByVal nVirtKeyCode As KeyCodeConstants _
    ) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        gbKeyIsPressed = True
    End If
End Function

'*************************************************************************
'**函 数 名：HIWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的高16位
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' 取出32位值的高16位
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**函 数 名：LOWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的低16位
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' 取出32位值的低16位
     LOWORD = LongIn And &HFFFF&
End Function

Public Sub PressKey(bytKey As Byte)
    '功能：向键盘发送一个键,类似SendKey
    '参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
   
Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
    '功能:打开中文输入法，或关闭输入法
    '     根据zlComlib中同名函数更改，并利用ZLHIS软件中的本地参数控制
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    Dim strUser As String
    
    strUser = GetSetting("ZLSOFT", "注册信息\登陆信息", "USER", "")
    '用户没进行设置，就不处理
    strIme = GetSetting("ZLSOFT", "私有全局\" & strUser, "输入法", "")
    If strIme = "" And blnOpen = True Then Exit Function                 '要求打开输入法，但是又没有设置
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否批定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                    Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '不是输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

'######################################################################################
'   绘制彩色进度条
'######################################################################################
Public Sub DrawProgress( _
      lPercent As Single, _
      ByVal lhDC As Long, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long _
   )
Dim hBr As Long
Dim tR As RECT
Dim tProgR As RECT

   tR.left = lLeft + 1
   tR.top = lTop + 1
   tR.right = lRight - 1
   tR.bottom = lBottom - 1

   ' Draw the progress bar
   LSet tProgR = tR
   tProgR.right = tProgR.left + (tProgR.right - tProgR.left) * lPercent
   GradientFillRect lhDC, tProgR, RGB(234, 94, 45), RGB(238, 164, 36), GRADIENT_FILL_RECT_H
   
   ' Draw the text in front of the progress bar
'   DrawTextA lHDC, Format(lPercent, "0%"), -1, tR, DT_CENTER

   ' Frame the progress bar:
   hBr = CreateSolidBrush(&H0&)
   FrameRect lhDC, tR, hBr
   DeleteObject hBr
End Sub

Public Sub GradientFillRect( _
      ByVal lhDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).X = tR.left
   tTV(0).Y = tR.top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).X = tR.right
   tTV(1).Y = tR.bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lhDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lhDC, tR, hBrush
      DeleteObject hBrush
   End If
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub

Public Function GetCharPosFromByteValue(ByVal S As String, ByVal Pos As Long) As Long
'已知字符串的字节长度，求其字符长度！
    Dim iLoop As Long
    Dim iChinese As Long
    iChinese = 0
    For iLoop = Pos To 1 Step -1
        If Asc(StrConv(MidB(StrConv(S, vbFromUnicode), iLoop, 1), vbUnicode)) = 0 Then
            iChinese = iChinese + 1
        End If
    Next iLoop
    GetCharPosFromByteValue = Pos - iChinese \ 2
End Function


Public Function GetTempName(TmpFilePrefix As String) As String
'获取WIndows临时目录
     Dim TempFileName As String * 256
     Dim X As Long
     Dim DriveName As String
     DriveName = "c:"
       X = GetTempFileName(DriveName, TmpFilePrefix, 0, TempFileName)
       GetTempName = left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)
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
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
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
            udtRect.bottom = Height
            udtRect.right = Width
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
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
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
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
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

'######################################################################################
'## 繁简转换
'######################################################################################
Public Function J2F(ByVal strText As String) As String
    '简体转繁体
    Dim strF As String      '繁体字符串
    Dim strJ As String      '简体字符串
    Dim STlen As Long       '待转换字串长度
    
    strJ = strText
    STlen = lstrlen(strJ)
    strF = Space(STlen)
    LCMapString &H804, &H4000000, strJ, STlen, strF, STlen
    J2F = strF
End Function

Public Function F2J(ByVal strText As String) As String
    '繁体转简体
    Dim strF As String      '繁体字符串
    Dim strJ As String      '简体字符串
    Dim STlen As Long       '待转换字串长度
    strF = strText
    STlen = lstrlen(strF)
    strJ = Space(STlen)
    LCMapString &H804, &H2000000, strF, STlen, strJ, STlen
    F2J = strJ
End Function


