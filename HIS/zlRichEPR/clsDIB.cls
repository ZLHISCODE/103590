VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Private Type SAFEARRAY
    cDims As Integer         '这个数组有几维？
    fFeatures As Integer     '这个数组有什么特性？
    cbElements As Long       '数组的每个元素有多大？
    cLocks As Long           '这个数组被锁定过几次？
    pvData As Long           '这个数组里的数据放在什么地方？
    'rgsabound() As SFArrayBOUND
End Type

Private Type SAFEARRAYBOUND
    cElements As Long      '这一维有多少个元素？
    lLbound As Long        '它的索引从几开始？
End Type

Private Type SAFEARRAY1
    cDims As Integer         '这个数组有几维？
    fFeatures As Integer     '这个数组有什么特性？
    cbElements As Long       '数组的每个元素有多大？
    cLocks As Long           '这个数组被锁定过几次？
    pvData As Long           '这个数组里的数据放在什么地方？
    CE0 As Long              '这一维有多少个元素？
    LB0 As Long              '它的索引从几开始？
End Type

Private Type SAFEARRAY2
    cDims As Integer         '这个数组有几维？
    fFeatures As Integer     '这个数组有什么特性？
    cbElements As Long       '数组的每个元素有多大？
    cLocks As Long           '这个数组被锁定过几次？
    pvData As Long           '这个数组里的数据放在什么地方？
    CE0 As Long              '这一维有多少个元素？
    LB0 As Long              '它的索引从几开始？
    CE1 As Long
    LB1 As Long
End Type

Private Const pvDataPtrAdd As Long = 12&


'-- 关键变量 --------------------------------------
Private InitPtrFlag As Boolean

Private pLongAll(0 To 0) As Long
Private pLongAllPtr(0 To 0) As Long
Private OldpLongAll As Long
Private OldpLongAllPtr As Long

'-- 2个3Byte指针 ----------------------------------
Private p3Byte0(0 To 2) As Byte
Private p3Byte0Ptr(0 To 0) As Long
Private Oldp3Byte0 As Long
Private Oldp3Byte0Ptr As Long

Private p3Byte1(0 To 2) As Byte
Private p3Byte1Ptr(0 To 0) As Long
Private Oldp3Byte1 As Long
Private Oldp3Byte1Ptr As Long

'-- 2个Long指针 -----------------------------------
Private pLong0(0 To 0) As Long
Private pLong0Ptr(0 To 0) As Long
Private OldpLong0 As Long
Private OldpLong0Ptr As Long

Private pLong1(0 To 0) As Long
Private pLong1Ptr(0 To 0) As Long
Private OldpLong1 As Long
Private OldpLong1Ptr As Long

'######################################################################################################################
'## 设备场景函数 ###############################
'== Fun ========================================
'CreateCompatibleDC 创建一个与特定设备场景一致的内存设备场景。返回值：Long，新设备场景句柄，若出错则为零
'DeleteDC           删除专用设备场景或信息场景，释放所有相关窗口资源。不要将它用于GetDC函数取回的设备场景。返回值：Long，执行成功则为非零，失败则为零
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'== Const ======================================
'== Type =======================================


'## 光栅运算函数 ###############################
'== Fun ========================================
'BitBlt            将一幅位图从一个设备场景复制到另一个。源和目标DC相互间必须兼容。返回值：Long，非零表示成功，零表示失败。会设置GetLastError
'SetStretchBltMode 指定StretchBlt 和 StretchDIBits函数的伸缩模式。这种伸缩模式定义了Windows如何对伸缩过程中剔除的扫描线进行控制。对于VB窗体和控件，倘若在API绘图过程中使用这个函数，建议恢复原来的StretchBlt模式
'StretchBlt        将一幅位图从一个设备场景复制到另一个。源和目标DC相互间必须兼容。这个函数会在设备场景中定义一个目标矩形，并在位图中定义一个源图象。源矩形会根据需要进行伸缩，以便与目标矩形的大小相符。返回值：Long，非零表示成功，零表示失败。会设置GetLastError
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'== Const ======================================
'== Type =======================================
Private Const STRETCH_ANDSCANS = 1    '默认设置。剔除的线段与剩下的线段进行AND运算。这个模式通常应用于采用了白色背景的单色位图
Private Const STRETCH_ORSCANS = 2     '剔除的线段被简单的清除。这个模式通常用于彩色位图
Private Const STRETCH_DELETESCANS = 3 '剔除的线段与剩下的线段进行OR运算。这个模式通常应用于采用了白色背景的单色位图
Private Const STRETCH_HALFTONE = 4    '目标位图上的像素块被设为源位图上大致近似的块。这个模式要明显慢于其他模式


'## 绘图函数 ###################################
'== Fun ========================================
'DeleteObject 用这个函数删除GDI对象，比如画笔、刷子、字体、位图、区域以及调色板等等。对象使用的所有系统资源都会被释放。返回值：Long，非零表示成功，零表示失败
'GetObject    取得对指定对象进行说明的一个结构。windows手册建议用GetObject这个名字来引用该函数。GetObjectAPI在vb中用于避免与GetObject关键字混淆。返回值：Long，如lpObject设为零（用ByVal As Long在这种情况下传递一个NULL参数），则必须设置缓冲区的长度。如执行成功，返回载入结构内部的实际字节数；如失败，返回零值
'SelectObject 每个设备场景都可能有选入其中的图形对象。其中包括位图、刷子、字体、画笔以及区域等等。一次选入设备场景的只能有一个对象。选定的对象会在设备场景的绘图操作中使用。例如，当前选定的画笔决定了在设备场景中描绘的线段颜色及样式。返回值：Long，与以前选入设备场景的相同hObject类型的一个对象的句柄，零表示出错。如选定的对象是一个区域（Region），结果就是下列常数之一：SIMPLEREGION， COMPLEXREGION 或 NULLREGION 对区域进行描述， GDI_ERROR 表示出错
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'== Const ======================================
'== Type =======================================
Private Type BitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


'## 硬件与系统函数 #############################
'== Fun ========================================
'CopyMemory  复制内存
'timeGetTime 取得时间。单位为 毫秒
'ZeroMemory  清空内存
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)


'######################################################################################################################

Private Const CW_USEDEFAULT As Long = &H80000000

'## 位图创建 ###################################
'== Fun ========================================
'CreateDIBSection 创建一个DIBSection。这是一个GDI对象，可象一幅与设备有关位图那样使用。但是，它在内部作为一幅与设备无关位图保存。返回值：Long，执行成功返回DIBSection位图的句柄，零表示失败。会设置GetLastError
'LoadImage        载入一个位图、图标或指针
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'== Const ======================================
Private Const DIB_RGB_COLORS = 0 'BITMAPINFO包含了一个颜色表，其中保存有32位颜色（RGBQUAD）
Private Const DIB_PAL_COLORS = 1 'BITMAPINFO包含了一个16位调色板索引的数组


Private Const IMAGE_BITMAP      As Long = 0 '
Private Const IMAGE_ICON        As Long = 1 '
Private Const IMAGE_CURSOR      As Long = 2 '
Private Const IMAGE_ENHMETAFILE As Long = 3 '

Private Const LR_DEFAULTCOLOR        As Long = &H0 '以常规方式载入图象
Private Const LR_MONOCHROME          As Long = &H1 '将图象转换成单色
Private Const LR_COLOR               As Long = &H2 '
Private Const LR_COPYRETURNORG       As Long = &H4 'Creates an exact copy of the image, ignoring the cxDesired and cyDesired parameters.
Private Const LR_COPYDELETEORG       As Long = &H8 'Deletes the original image after creating the copy.
Private Const LR_LOADFROMFILE       As Long = &H10 '如hInst为零，lpsz就代表要载入适当类型的一个文件的名字，仅适用于Win95
Private Const LR_LOADTRANSPARENT    As Long = &H20 '与图象中第一个像素相符的所有像素都由系统替换
Private Const LR_DEFAULTSIZE        As Long = &H40 '不对图象进行缩放处理。忽略n1和n2的设置
Private Const LR_VGACOLOR           As Long = &H80 '使用真彩色？Uses true VGA colors.
Private Const LR_LOADMAP3DCOLORS  As Long = &H1000 '将图象中的深灰、灰、以及浅灰像素都替换成COLOR_3DSHADOW，COLOR_3DFACE以及COLOR_3DLIGHT的当前设置
Private Const LR_CREATEDIBSECTION As Long = &H2000 '如果指定了IMAGE_BITMAP，就返回DIBSection的句柄，而不是位图的句柄
Private Const LR_COPYFROMRESOURCE As Long = &H4000 'Tries to reload an icon or cursor resource from the original resource file rather than simply copying the current image. This is useful for creating a different-sized copy when the resource file contains multiple sizes of the resource. Without this flag, CopyImage stretches the original image to the new size. If this flag is set, CopyImage uses the size in the resource file closest to the desired size.
Private Const LR_SHARED           As Long = &H8000 '将图象作为一个共享资源载入。在NT 4.0中装载固有资源时要用到这个设置

'## 位图属性 ###################################
'== Fun ========================================
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQuad) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQuad) As Long

'## Other ######################################
Private Type BITMAPFILEHEADER
    bfType(1 To 2) As Byte
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BitMapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Const BI_RGB       As Long = 0&
Private Const BI_RLE8      As Long = 1&
Private Const BI_RLE4      As Long = 2&
Private Const BI_BitFields As Long = 3&
Private Const BI_JPEG      As Long = 4&
Private Const BI_PNG       As Long = 5&

Private Type RGBQuad
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BitMapInfo
    bmiHeader As BitMapInfoHeader
    bmiColors(0 To &HFF) As RGBQuad
End Type

'##################################################

Private Enum ScaleModeConst
    SMC_Min = 0
    SMC_Nearest = 0    '最邻近插值
    SMC_StretchBlt = 1 'StretchBlt
    SMC_BiliNear = 2   '双线性内插值
    SMC_ThreeOrder = 3 '三次卷积法
    SMC_Max = 3
End Enum

'######################################################################################################################
Private MyhDC As Long
Private MyBMI As BitMapInfo
Private MyhDIB As Long
Private MyPtr As Long
Private m_WidthBytes As Long
Private hOldMap As Long

'######################################################################################################################
Private Sub Class_initialize()
    MyhDC = 0
End Sub

Private Sub Class_Terminate()
    Free
End Sub

Public Property Get DIBDC() As Long
    DIBDC = MyhDC
End Property

Public Property Get DataSize() As Long
    DataSize = MyBMI.bmiHeader.biSizeImage
End Property

Public Property Get Width() As Long
    Width = MyBMI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = MyBMI.bmiHeader.biHeight
End Property

Public Property Get WidthBytes() As Long
    WidthBytes = m_WidthBytes
End Property

Public Property Get ColorBit() As Long
    ColorBit = MyBMI.bmiHeader.biBitCount
End Property

Public Property Get DataPtr() As Long
    DataPtr = MyPtr
End Property

'创建图片
Public Function Create(ByVal Width As Long, ByVal Height As Long, Optional ByVal CBit As Long = 32, Optional ClsMap As Boolean = True) As Boolean

    If Width <= 0 Or Height <= 0 Then
        Create = False
        Exit Function
    End If
    
    Select Case CBit
    Case 1, 4, 8, 16, 24, 32
    Case Else
        Create = False
        Exit Function
    End Select
    
    Free
    
    With MyBMI.bmiHeader
        .biSize = Len(MyBMI.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = CBit
        .biCompression = BI_RGB
        m_WidthBytes = GetWidthBytes(.biWidth, .biBitCount)
        m_WidthBytes = (m_WidthBytes + 3) And &H7FFFFFFC
        .biSizeImage = m_WidthBytes * .biHeight
        
    End With
    MyhDC = CreateCompatibleDC(0)
    MyhDIB = CreateDIBSection(MyhDC, MyBMI, IIf(CBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS), MyPtr, 0, 0)
    If MyhDIB Then
        hOldMap = SelectObject(MyhDC, MyhDIB)
        If ClsMap Then Cls ': SetRectAlpha &HFF
        Create = True
    End If
    
End Function

'释放图片
Public Sub Free()
    If Me.DataPtr Then
        If hOldMap Then DeleteObject SelectObject(MyhDC, hOldMap)
        hOldMap = 0
        DeleteDC MyhDC
        MyhDC = 0
        MyBMI.bmiHeader.biBitCount = 0
        MyBMI.bmiHeader.biWidth = 0
        MyBMI.bmiHeader.biHeight = 0
        MyBMI.bmiHeader.biSizeImage = 0
        MyhDIB = 0
        MyPtr = 0
        
    End If
    
End Sub

'初始化图片
Public Sub Cls()
    If Me.DataPtr Then ZeroMemory ByVal MyPtr, MyBMI.bmiHeader.biSizeImage
End Sub


Public Function PutTo(ByVal hDestDC As Long, _
        Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
        Optional ByVal nWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nHeight As Long = CW_USEDEFAULT, _
        Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0) As Boolean
    If Me.DataPtr = 0 Then Exit Function
    If nWidth = CW_USEDEFAULT Then nWidth = Me.Width
    If nHeight = CW_USEDEFAULT Then nHeight = Me.Height
    
    PutTo = SetDIBitsToDevice(hDestDC, X, Y, Width, Height, _
            xSrc, ySrc, _
            0, Me.Height, _
            ByVal Me.DataPtr, MyBMI, _
            IIf(Me.ColorBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS))
    
End Function

Public Function PutToEx(ByVal hDestDC As Long, _
        Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
        Optional ByVal nWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nHeight As Long = CW_USEDEFAULT, _
        Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
        Optional ByVal nSrcWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nSrcHeight As Long = CW_USEDEFAULT, _
        Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Boolean
    If Me.DataPtr = 0 Then Exit Function
    If nWidth = CW_USEDEFAULT Then nWidth = Me.Width
    If nHeight = CW_USEDEFAULT Then nHeight = Me.Height
    If nSrcWidth = CW_USEDEFAULT Then nSrcWidth = Me.Width
    If nSrcHeight = CW_USEDEFAULT Then nSrcHeight = Me.Height
    
    PutToEx = StretchDIBits(hDestDC, X, Y, nWidth, nHeight, _
            xSrc, ySrc, nSrcWidth, nSrcHeight, _
            ByVal Me.DataPtr, MyBMI, _
            IIf(Me.ColorBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS), _
            dwRop)
    
End Function


Private Function ChkFileWrite(FileName As String) As Boolean
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    On Error Resume Next
    
    Open FileName For Output As #FileNum
    
    If Err.Number Then
        '
    Else
        Close #FileNum
        ChkFileWrite = True
    End If
    
    On Error GoTo 0
    
End Function

Public Function SaveBMP(FileName As String) As Boolean
    Dim FileNum As Integer
    Dim TempBMFH As BITMAPFILEHEADER
    Dim TCB() As RGBQuad
    Dim TempBytes() As Byte
    
    If (ChkFileWrite(FileName) = False) Or (MyPtr = 0) Then Exit Function
    
    TempBMFH.bfType(1) = Asc("B")
    TempBMFH.bfType(2) = Asc("M")
    TempBMFH.bfOffBits = Len(TempBMFH) + Len(MyBMI.bmiHeader)
    If MyBMI.bmiHeader.biBitCount <= 8 Then
        TempBMFH.bfOffBits = TempBMFH.bfOffBits + 4 * 2 ^ MyBMI.bmiHeader.biBitCount
        
        ReDim TCB(1 To 2 ^ MyBMI.bmiHeader.biBitCount)
        GetDIBColorTable MyhDC, 0, 2 ^ MyBMI.bmiHeader.biBitCount, TCB(1)
        
    End If
    
    TempBMFH.bfSize = TempBMFH.bfOffBits + MyBMI.bmiHeader.biSizeImage
    
    ReDim TempBytes(1 To MyBMI.bmiHeader.biSizeImage)
    CopyMemory TempBytes(1), ByVal MyPtr, MyBMI.bmiHeader.biSizeImage
    
    FileNum = FreeFile
    
    Open FileName For Binary As #FileNum
    
    Put #FileNum, , TempBMFH
    Put #FileNum, , MyBMI.bmiHeader
    If MyBMI.bmiHeader.biBitCount <= 8 Then Put #FileNum, , TCB
    Put #FileNum, , TempBytes
    
    Close #FileNum
    
    SaveBMP = True
    
End Function


'######################################################################################################################
Public Function GetWidthBytes(ByVal Width As Long, ByVal ColorBits As Long) As Long
    If Width <= 0 Then
        GetWidthBytes = 0
        Exit Function
    End If
    
    Select Case ColorBits
    Case 1, 4, 8, 16, 24, 32
        Dim TempNum As Long
        
        If ColorBits < 8 Then
            TempNum = (Width * ColorBits + 7) \ 8
        Else
            TempNum = Width * (ColorBits \ 8)
        End If
        GetWidthBytes = (TempNum + 3) And &H7FFFFFFC
        
    Case Else
        GetWidthBytes = 0
    End Select
    
End Function

Public Function DIBLoadMap(FileName As String, Optional ByVal AutoSize As Boolean = True, Optional ByVal CBit As Long = 24) As Boolean
    Dim hMap As Long
    Dim hDC As Long
    Dim hOldMap As Long
    Dim TempPic As IPictureDisp
    Dim BM As BitMap
    
'    If dDIB Is Nothing Then Set dDIB = New clsDIB
    If DIBDC = 0 Then AutoSize = True
    
    If AutoSize Then
        If CBit = 16 Or CBit = 24 Or CBit = 32 Then
        Else
            Exit Function
        End If
        
    End If
        
    hMap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_CREATEDIBSECTION)
    If hMap Then
        '
    Else
        On Error Resume Next
        Set TempPic = LoadPicture(FileName)
        On Error GoTo 0
        If TempPic Is Nothing Then
        Else
            If TempPic.Type = vbPicTypeBitmap Then
                hMap = TempPic.handle
            Else
                Set TempPic = Nothing
            End If
        End If
        
    End If
    
    If hMap Then
        hDC = CreateCompatibleDC(0)
        If hDC = 0 Then GoTo FreeMe
        
        GetObjectAPI hMap, Len(BM), BM
        If AutoSize Then
            If Create(BM.bmWidth, BM.bmHeight, CBit, False) Then
            Else
                GoTo FreeMe
                
            End If
            
        End If
        
        hOldMap = SelectObject(hDC, hMap)
        SetStretchBltMode DIBDC, STRETCH_HALFTONE
        StretchBlt DIBDC, 0, 0, Width, Height, hDC, 0, 0, BM.bmWidth, BM.bmHeight, vbSrcCopy
        Call SelectObject(hDC, hOldMap)
        
        DIBLoadMap = True
        
    End If
    
    
FreeMe:
    If hDC Then DeleteDC hDC
    If TempPic Is Nothing Then DeleteObject hMap
    Set TempPic = Nothing
    
End Function

Public Function ScaleModeName(ByVal ScaleMode As Byte) As String
    Dim s As String
    
    Select Case ScaleMode
    Case ScaleModeConst.SMC_Nearest
        s = "最邻近插值"
    Case ScaleModeConst.SMC_StretchBlt
        s = "StretchBlt"
    Case ScaleModeConst.SMC_BiliNear
        s = "双线性内插值"
    Case ScaleModeConst.SMC_ThreeOrder
        s = "三次卷积法"
    End Select
    
    ScaleModeName = s
    
End Function



Private Function Sinxx(ByVal X As Single) As Single
    X = Abs(X)
    If X < 1 Then
        Sinxx = 1 - 2 * X * X + X * X * X
    ElseIf X < 2 Then
        Sinxx = 4 - 8 * X + 5 * X * X - X * X * X
    Else
        Sinxx = 0
    End If
    
End Function

Public Function DIBScale(ByVal sDIB As clsDIB, ByVal W As Long, ByVal H As Long) As Boolean
                
'    If sDIB Is Nothing Then Exit Function
    If sDIB.DataPtr = 0 Then Exit Function
    If sDIB.ColorBit <> 24 Then Exit Function
    
    If W < 2 Then Exit Function
    If H < 2 Then Exit Function
    
    'If dDIB Is Nothing Then Set dDIB = New clsDIB
    
    If Create(W, H, 24, False) = False Then Exit Function
    
    If W = sDIB.Width And H = sDIB.Height Then
        CopyMemory ByVal DataPtr, ByVal sDIB.DataPtr, sDIB.DataSize
        DIBScale = True
        Exit Function
    End If
    
    Dim pMapData() As Byte
    Dim pMapDataInfo As SAFEARRAY1
    
    pMapDataInfo.cDims = 1
    pMapDataInfo.fFeatures = 0
    pMapDataInfo.cbElements = 1
    pMapDataInfo.cLocks = 0
    pMapDataInfo.pvData = 0
    pMapDataInfo.CE0 = 12
    pMapDataInfo.LB0 = 0
    pLongAllPtr(0) = VarPtrArray(pMapData)
    pLongAll(0) = VarPtr(pMapDataInfo)
    
    
    Dim sMaxX As Long, sMaxY As Long
    Dim dMaxX As Long, dMaxY As Long
    
    sMaxX = sDIB.Width - 1
    sMaxY = sDIB.Height - 1
    dMaxX = Width - 1
    dMaxY = Height - 1
    
    Dim sPtr As Long, dPtr As Long
    Dim sLineBytes As Long, dLineBytes As Long
    
    sPtr = sDIB.DataPtr
    dPtr = DataPtr
    sLineBytes = sDIB.WidthBytes
    dLineBytes = WidthBytes
    
    Dim sPixelAdd As Long, sLineAdd As Long
    Dim dPixelAdd As Long, dLineAdd As Long
    
    sPixelAdd = 3
    sLineAdd = sLineBytes - sDIB.Width * sPixelAdd
    dPixelAdd = 3
    dLineAdd = dLineBytes - Width * dPixelAdd
    
    Dim X As Single, Y As Single
    Dim nX As Long, nY As Long
    Dim u As Single, v As Single
    Dim u_ As Single, v_ As Single
    
    Dim LinenX() As Long
    Dim Lineu() As Single
    
    Dim I As Long, J As Long
    Dim fI As Long, fJ As Long
    
    Dim LinePtr(0 To 3) As Long
    Dim CurPtr As Long, CurIdx As Long
    
    Dim OldX As Long, OldY As Long
    Dim TempLng As Long

    
    SetStretchBltMode DIBDC, STRETCH_HALFTONE
    StretchBlt DIBDC, 0, 0, Width, Height, sDIB.DIBDC, 0, 0, sDIB.Width, sDIB.Height, vbSrcCopy
        
    
    pLongAllPtr(0) = VarPtrArray(pMapData)
    pLongAll(0) = 0
    
    DIBScale = True
    
End Function

'######################################################################################################################
Public Sub MakePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByRef OldArrPtr As Long, ByRef OldpArrPtr As Long)
    Dim TempLng As Long
    Dim TempPtr As Long
    
    If InitPtrFlag Then
        Dim OldPtr As Long
        
        OldPtr = pLongAllPtr(0)
        pLongAllPtr(0) = DataArrPtr
        TempLng = pLongAll(0) + pvDataPtrAdd
        pLongAllPtr(0) = pDataArrPtr
        TempPtr = pLongAll(0) + pvDataPtrAdd
        pLongAllPtr(0) = TempPtr
        OldpArrPtr = pLongAll(0)
        pLongAll(0) = TempLng
        pLongAllPtr(0) = TempLng
        OldArrPtr = pLongAll(0)
        pLongAllPtr(0) = OldPtr
        
    Else
        CopyMemory TempLng, ByVal DataArrPtr, 4 '得到DataArrPtr的SAFEARRAY结构的地址
        TempLng = TempLng + pvDataPtrAdd '这个指针偏移12个字节后就是pvData指针
        CopyMemory TempPtr, ByVal pDataArrPtr, 4 '得到pDataArrPtr的SAFEARRAY结构的地址
        TempPtr = TempPtr + pvDataPtrAdd '这个指针偏移12个字节后就是pvData指针
        CopyMemory OldpArrPtr, ByVal TempPtr, 4  '保存旧地址
        CopyMemory ByVal TempPtr, TempLng, 4  '使pDataArrPtr指向DataArrPtr的SAFEARRAY结构的pvData指针
        CopyMemory OldArrPtr, ByVal TempLng, 4 '保存旧地址
        
    End If
    
End Sub

Public Sub FreePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByVal OldArrPtr As Long, ByVal OldpArrPtr As Long)
    Dim TempPtr As Long
    
    If InitPtrFlag Then
        pLongAllPtr(0) = DataArrPtr
        pLongAllPtr(0) = pLongAll(0) + pvDataPtrAdd
        pLongAll(0) = OldArrPtr
        pLongAllPtr(0) = pDataArrPtr
        pLongAllPtr(0) = pLongAll(0) + pvDataPtrAdd
        pLongAll(0) = OldpArrPtr
        
    Else
        CopyMemory TempPtr, ByVal DataArrPtr, 4 '得到DataArrPtr的SAFEARRAY结构的地址
        CopyMemory ByVal (TempPtr + pvDataPtrAdd), OldArrPtr, 4 '恢复旧地址
        CopyMemory TempPtr, ByVal pDataArrPtr, 4 '得到pDataArrPtr的SAFEARRAY结构的地址
        CopyMemory ByVal (TempPtr + pvDataPtrAdd), OldpArrPtr, 4 '恢复旧地址
        
    End If
    
End Sub

Public Sub PointInit()
    If InitPtrFlag Then Exit Sub
    
    MakePoint VarPtrArray(pLongAll), VarPtrArray(pLongAllPtr), OldpLongAll, OldpLongAllPtr
    InitPtrFlag = True
    
    MakePoint VarPtrArray(p3Byte0), VarPtrArray(p3Byte0Ptr), Oldp3Byte0, Oldp3Byte0Ptr
    MakePoint VarPtrArray(p3Byte1), VarPtrArray(p3Byte1Ptr), Oldp3Byte1, Oldp3Byte1Ptr
    
    MakePoint VarPtrArray(pLong0), VarPtrArray(pLong0Ptr), OldpLong0, OldpLong0Ptr
    MakePoint VarPtrArray(pLong1), VarPtrArray(pLong1Ptr), OldpLong1, OldpLong1Ptr
    
End Sub

Public Sub PointFree()
    If InitPtrFlag = False Then Exit Sub
    
    FreePoint VarPtrArray(p3Byte0), VarPtrArray(p3Byte0Ptr), Oldp3Byte0, Oldp3Byte0Ptr
    FreePoint VarPtrArray(p3Byte1), VarPtrArray(p3Byte1Ptr), Oldp3Byte1, Oldp3Byte1Ptr
    
    FreePoint VarPtrArray(pLong0), VarPtrArray(pLong0Ptr), OldpLong0, OldpLong0Ptr
    FreePoint VarPtrArray(pLong1), VarPtrArray(pLong1Ptr), OldpLong1, OldpLong1Ptr
    
    InitPtrFlag = False
    FreePoint VarPtrArray(pLongAll), VarPtrArray(pLongAllPtr), OldpLongAll, OldpLongAllPtr
    
End Sub









