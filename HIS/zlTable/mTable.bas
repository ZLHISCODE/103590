Attribute VB_Name = "mTable"
'#########################################################################
'##描    述：表格相关声明
'#########################################################################
Option Explicit

Public p_lIconWidth As Long         '图标宽度
Public p_lIconHeight As Long        '图标高度
Public p_TPPX As Long               'screen.TwipsPerPixelX
Public p_TPPY As Long               'screen.TwipsPerPixelY

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

' Use default rgb colour:
Public Const CLR_NONE = -1
    
' Private variables
Private m_bIsXp As Boolean
Private m_bIsNt As Boolean
Private m_bInit As Boolean
Public Const MAGIC_END_EDIT_IGNORE_WINDOW_PROP As String = "zlTable:Table"

Public Type ColInfoType
    ColWidth As Long
    LeftX As Long
    FixedWidth As Boolean
    Visible As Boolean
End Type

Public Type RowInfoType
    RowHeight As Long           '最终行高
    FixedHeight As Boolean
    TopY As Long
End Type

Public ColInfo() As ColInfoType
Public RowInfo() As RowInfoType

'#########################################################################################################
'## 功能：  返回指定句柄的窗体的类名称
'## 参数：  hWnd: 窗体句柄
'## 返回：  类名称
'#########################################################################################################
Public Function WindowClassName(ByVal hWnd As Long) As String
    Dim szBuf As String
    Dim lR As Long
    szBuf = String$(260, 0)
    lR = GetClassName(hWnd, szBuf, 260)
    lR = InStr(szBuf, vbNullChar)
    If (lR > 0) Then
        WindowClassName = Left$(szBuf, lR - 1)
    Else
        WindowClassName = szBuf
    End If
End Function

'#########################################################################################################
'## 功能：  采用XOR技术在列被Resize的时候绘制拖动的图像
'## 参数：  rcNew:  绘制拖动图像的矩形
'##         bFirst: 是否是第一次绘制
'##         bLast:  是否是最后一次绘制
'## 返回：  无
'#########################################################################################################
Public Sub DrawDragImage(ByRef rcNew As RECT, ByVal bFirst As Boolean, ByVal bLast As Boolean)
    Static rcCurrent As RECT    '静态变量，便于存储位置
    Dim hDC As Long
       
    '首先获取桌面DC
    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    '设置绘制模式为XOR
    SetROP2 hDC, R2_NOTXORPEN
    
    '覆盖和清除旧矩形
    If Not (bFirst) Then
       Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
    End If
    
    If Not (bLast) Then
       '绘制新矩形
       Rectangle hDC, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
    End If
    
    '存储这个位置，便于我们下次擦除它
    LSet rcCurrent = rcNew  '用户自定义变量的赋值用 LSet
    
    '必须释放桌面DC（记得要做这一步！）
    DeleteDC hDC
End Sub

'#########################################################################################################
'## 功能：  采用指定选项来绘制ImageList中的一个图形
'## 参数：
'##         hIml:               一个VB6 ImageList的hImageList
'##         iIndex:             从0开始的图像索引
'##         hDC:                目标设备场景句柄
'##         xPixels:            X轴位置
'##         yPixels:            Y轴位置
'##         lIconSizeX:         水平尺寸
'##         lIconSizeY:         垂直尺寸
'##         bSelected:          选中效果
'##         bDisabled:          无效图标
'## 返回：  无
'#########################################################################################################
Public Sub DrawImageIcon( _
    ByVal lhIml As Long, _
    ByVal iIndex As Long, _
    ByVal hDC As Long, _
    ByVal xPixels As Integer, _
    ByVal yPixels As Integer, _
    ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
    Optional ByVal bSelected As Boolean = False, _
    Optional ByVal bDisabled As Boolean = False)
      
    If iIndex > -1 Then
        Dim lFlags As Long
        lFlags = ILD_TRANSPARENT
        If (bSelected) Then
            lFlags = lFlags Or ILD_SELECTED
        End If
        If (bDisabled) Then
            Dim hIcon As Long
            hIcon = ImageList_GetIcon(lhIml, iIndex, 0)
                           
            If Not (hIcon = 0) And Not (hIcon = -1) Then
                DrawState hDC, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
                ' Clear up the icon:
                DestroyIcon hIcon
            End If
        Else
            ImageList_Draw lhIml, iIndex, hDC, xPixels, yPixels, lFlags
        End If
    End If
End Sub

'#########################################################################################################
'## 功能：  在分组行的前面绘制一个展开/关闭的符号。
'##         如果是XP效果下，会绘制一个TreeView的展开/关闭符号；否则是一个按钮加上+和-号。
'##
'## 参数：  hWnd:       用于检测主题的窗体句柄
'##         lHDC:       目标设备场景句柄
'##         tTR:        符号绘制的边界矩形
'##         bCollapsed: True表示折叠；False表示展开
'## 返回：  无
'#########################################################################################################
Public Sub DrawOpenCloseGlyph( _
    ByVal hWnd As Long, _
    ByVal lhDC As Long, _
    tTR As RECT, _
    ByVal bCollapsed As Boolean)
      
    Dim tGR As RECT
    Dim bDone As Boolean
   
    LSet tGR = tTR
    tGR.Left = tGR.Left + 2
    tGR.Right = tGR.Left + 12
    tGR.Top = tGR.Top + (tGR.Bottom - tGR.Top - 12) \ 2
    tGR.Bottom = tGR.Top + 12
    
    If (IsXp) Then
        'XP系统
        Dim hTheme As Long
        hTheme = OpenThemeData(hWnd, StrPtr("TREEVIEW"))    '获取TreeView的主题
        If Not (hTheme = 0) Then
            '获取成功，则显示TreeView主题的展开/关闭符号
            DrawThemeBackground hTheme, lhDC, 2, IIf(bCollapsed, 1, 2), tGR, tGR
            CloseThemeData hTheme   '关闭主题
            bDone = True
        End If
    End If
    
    If Not (bDone) Then
        '绘制按钮边框
        Dim hBr As Long
        hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        FillRect lhDC, tGR, hBr
        DeleteObject hBr
        
        Dim hPen As Long
        Dim hPenOld As Long
        Dim tJ As POINTAPI
        
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Left + 1, tGR.Bottom - 2, tJ
        LineTo lhDC, tGR.Right - 2, tGR.Bottom - 2
        LineTo lhDC, tGR.Right - 2, tGR.Top
        SelectObject lhDC, hPenOld
        DeleteObject hPen
        
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Right - 2, tGR.Top, tJ
        LineTo lhDC, tGR.Left, tGR.Top
        LineTo lhDC, tGR.Left, tGR.Bottom - 1
        SelectObject lhDC, hPenOld
        DeleteObject hPen
                          
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Left, tGR.Bottom - 1, tJ
        LineTo lhDC, tGR.Right - 1, tGR.Bottom - 1
        LineTo lhDC, tGR.Right - 1, tGR.Top
        
        ' Draw collapse/expand glyph
        MoveToEx lhDC, tGR.Left + 3, tGR.Top + 5, tJ
        LineTo lhDC, tGR.Left + 8, tGR.Top + 5
        If (bCollapsed) Then
            MoveToEx lhDC, tGR.Left + 5, tGR.Top + 3, tJ
            LineTo lhDC, tGR.Left + 5, tGR.Top + 8
        End If
        SelectObject lhDC, hPenOld
        DeleteObject hPen
    End If
End Sub

'#########################################################################################################
'## 功能：  将一个OLE_COLOR颜色转换为一个RGB长整型值
'## 参数：  oClr:   转换的颜色
'##         hPal:   在转换时采用的调色板
'## 返回：  RGB等价值，或者-1表示没有
'#########################################################################################################
Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    '将自动颜色转换为Windws颜色
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

'#########################################################################################################
'## 功能：  使用指定的Alpha值混合颜色
'##
'## 参数：  oColorFrom:     基础颜色
'##         oColorTo:       混合颜色
'##         alpha:          alpha值（0～255）
'## 返回：  混合后的RGB颜色
'#########################################################################################################
Public Property Get BlendColor( _
    ByVal oColorFrom As OLE_COLOR, _
    ByVal oColorTo As OLE_COLOR, _
    Optional ByVal Alpha As Long = 128) As Long
    
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)

    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
   
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
     
    BlendColor = RGB( _
       ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
       ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
       ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
       )
End Property

'#########################################################################################################
'## 功能：  将一个 StdFont 对象转换为相等的 Windows GDI LOGFONT 结构体
'##
'## 参数：  fntThis:    需要转换的字体
'##         hDC:        用于获取DPI信息的设备场景句柄
'##         tLF:        产生的LOGFONT结构体
'## 返回：  无
'#########################################################################################################
Public Sub OLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 1 To UBound(temp) + 1
            .lfFaceName(iChar - 1) = temp(iChar - 1)
        Next iChar
        
        ' Based on the Win32SDK documentation:
        .lfHeight = GetPixcelHeightByPoint(hDC, fntThis.Size)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
        If (IsXp) Then
          '如果为XP系统，则采用ClearType质量
           .lfQuality = CLEARTYPE_QUALITY
        Else
          '否则采用抗锯齿质量
           .lfQuality = ANTIALIASED_QUALITY
        End If
    End With
End Sub
       
'#########################################################################################################
'## 功能：  返回指定磅值的逻辑字体高度
'## 参数：  hDC:            目标设备句柄
'##         lPointValue:    字体磅值
'## 返回：  返回逻辑字体高度
'#########################################################################################################
Public Function GetPixcelHeightByPoint(hDC As Long, ByVal lPointValue As Double) As Double
    GetPixcelHeightByPoint = -MulDiv((lPointValue), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
End Function
       
'#########################################################################################################
'## 功能：  返回指定磅值的逻辑字体宽度
'## 参数：  hDC:            目标设备句柄
'##         lPointValue:    字体磅值
'## 返回：  返回逻辑字体高度
'#########################################################################################################
Public Function GetPixcelWidthByPoint(hDC As Long, ByVal lPointValue As Double) As Double
    GetPixcelWidthByPoint = -MulDiv((lPointValue), (GetDeviceCaps(hDC, LOGPIXELSX)), 72)
End Function

'#########################################################################################################
'## 功能：  修正的绘制文本方法
'## 参数：  hdc:        目标设备场景句柄
'##         sString:    字符串
'##         lCount:     字符串长度
'##         tR:         矩形框
'##         lFlags:     标志（附加属性）
'## 返回：  无
'#########################################################################################################
Public Sub DrawText(ByVal hDC As Long, ByVal sString As String, ByVal lCount As Long, tR As RECT, ByVal lFlags As Long)
    Dim lPtr As Long
    Dim tIR As RECT
    
    LSet tIR = tR       '拷贝结构体
    If (IsNt) Then
        '如果是NT以上版本的系统，调用DrawTextW
        lPtr = StrPtr(sString)
        If Not (lPtr = 0) Then
            DrawTextW hDC, lPtr, -1, tIR, lFlags
        End If
        lPtr = 0
    Else
        '否则调用DrawTextA
        DrawTextA hDC, sString, -1, tIR, lFlags
    End If
    If (lFlags And DT_CALCRECT) = DT_CALCRECT Then
        '如果只计算文本尺寸
        LSet tR = tIR
    End If
End Sub

'#########################################################################################################
'## 功能：  将一幅位图平铺到选定区域
'##
'## 参数：  hDC:        用于绘制的设备场景句柄
'##         x:          起始X坐标
'##         y:          起始Y坐标
'##         Width:      区域宽度
'##         Height:     区域高度
'##         lSrcDC:     包含图像的设备场景句柄
'##         lBitmapW:   源位图宽度
'##         lBitmapH:   源位图高度
'##         lSrcOffsetX:    X轴偏移量
'##         lSrcOffsetY:    Y轴偏移量
'## 返回：  无
'#########################################################################################################
Public Sub TileArea( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal lSrcDC As Long, _
        ByVal lBitmapW As Long, _
        ByVal lBitmapH As Long, _
        ByVal lSrcOffsetX As Long, _
        ByVal lSrcOffsetY As Long)
        
    Dim lSrcX As Long
    Dim lSrcY As Long
    Dim lSrcStartX As Long
    Dim lSrcStartY As Long
    Dim lSrcStartWidth As Long
    Dim lSrcStartHeight As Long
    Dim lDstX As Long
    Dim lDstY As Long
    Dim lDstWidth As Long
    Dim lDstHeight As Long

    lSrcStartX = ((X + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((Y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hDC, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop
End Sub

'#########################################################################################################
'## 功能：  传入一个通过ObjPtr()返回的COM对象的非引用指针，返回一个对象引用
'## 参数：  lPtr:   一个COM对象的非引用指针
'## 返回：  这个指针指向的COM对象
'#########################################################################################################
Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    ' 先将指针指向一个非法的、未计数的接口
    CopyMemory oTemp, lPtr, 4
    ' 不要在这里终止程序，会崩溃掉！！
    ' 分配一个合法的引用
    Set ObjectFromPtr = oTemp
    ' 不要在这里终止程序，会崩溃掉！！
    ' 破坏该合法引用
    CopyMemory oTemp, 0&, 4
    'OK！在这里终止程序，仍然会崩溃掉！不过这是因为子类而不是未计数的接口！
End Property

'#########################################################################################################
'## 功能：  判断系统是否是XP（或者更高版本）
'##
'## 返回：  True表示是XP系统或者更高版本
'#########################################################################################################
Public Property Get IsXp() As Boolean
    If Not (m_bInit) Then
        VerInitialise   '获取Windows版本信息
    End If
    IsXp = m_bIsXp
End Property

'#########################################################################################################
'## 功能：  判断系统是否是NT系统
'##
'## 返回：  True表示是NT/2000/XP or above
'#########################################################################################################
Public Property Get IsNt() As Boolean
    If Not (m_bInit) Then
        VerInitialise
    End If
    IsNt = m_bIsNt
End Property

'#########################################################################################################
'## 功能：  获取Windows版本信息
'#########################################################################################################
Private Sub VerInitialise()
    Dim lMajor As Long
    Dim lMinor As Long
    GetWindowsVersion lMajor, lMinor
    If (lMajor > 5) Then
        m_bIsXp = True  '是XP系统
    ElseIf (lMajor = 5) And (lMinor >= 1) Then
        m_bIsXp = True  '是XP系统
    End If
    m_bInit = True
End Sub

'#########################################################################################################
'## 功能：  返回当前的Windows版本信息
'#########################################################################################################
Private Sub GetWindowsVersion( _
    Optional ByRef lMajor = 0, _
    Optional ByRef lMinor = 0, _
    Optional ByRef lRevision = 0, _
    Optional ByRef lBuildNumber = 0)
      
    Dim lR As Long
    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)
    m_bIsNt = ((lR And &H80000000) = 0)     '是否是NT系统
End Sub

'#########################################################################################################
'## 功能：  设置窗体透明度
'#########################################################################################################
Public Sub SetTransparentForm(OBJ As Object, ByVal TransNum As Long)
    Dim Ret As Long
    'Set the window style to 'Layered'
    Ret = GetWindowLong(OBJ.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong OBJ.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes OBJ.hWnd, 0, TransNum, LWA_ALPHA
End Sub


