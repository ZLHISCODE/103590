Attribute VB_Name = "mAPI"
'#########################################################################
'##描    述：通用 Windows API 声明
'#########################################################################

Option Explicit

'#########################################################################
'矩形
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'点
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'窗体位置信息
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI       '最大化尺寸
    ptMaxPosition As POINTAPI   '最大化位置
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'#########################################################################
' 消息常数:
Public Const WM_ACTIVATE = &H6              '窗体状态常数：WA_INACTIVE（未激活） / WM_ACTIVATE（激活） / WA_CLICKACTIVE（鼠标激活）
Public Const WM_SETFOCUS = &H7              '具备焦点，应配合游标指针函数使用
Public Const WM_KILLFOCUS = &H8F            '去掉键盘焦点，应删除相关游标指针
Public Const WM_SETREDRAW = &HB             '强制刷新
Public Const WM_GETTEXTLENGTH = &HE         '返回文本字符长度，配合 GetWindowText() / WM_GETTEXT / LB_GETTEXT / CB_GETLBTEXT
Public Const WM_PAINT = &HF                 '绘制窗体
Public Const WM_ERASEBKGND = &H14           '清除窗体背景
Public Const WM_SETCURSOR = &H20            '设置游标
Public Const WM_MOUSEACTIVATE = &H21        '窗体由鼠标激活
Public Const WM_GETMINMAXINFO = &H24        '用于处理窗体最大化尺寸及位置
Public Const WM_WINDOWPOSCHANGING = &H46    '窗体状态改变
Public Const WM_NOTIFY = &H4E               '发生事件时，提示主窗体
Public Const WM_NCHITTEST = &H84            '光标移动或者鼠标点击、放开事件
Public Const WM_NCPAINT = &H85              '窗体框架绘制消息，可以通过捕获来自定义绘制框架，但一定是矩形的。
Public Const WM_KEYDOWN = &H100             '键盘焦点窗体中的非Alt^的按键按下。
Public Const WM_KEYUP = &H101               '键盘焦点窗体中的非Alt^的按键放开。
Public Const WM_CHAR = &H102                '返回WM_KEYDOWN的按键值。
Public Const WM_COMMAND = &H111             '菜单点击、控件向父窗体发送Notify信息或者快捷键按键时产生
Public Const WM_HSCROLL = &H114             '水平滚动条
Public Const WM_VSCROLL = &H115             '垂直滚动条
Public Const WM_SYSCOMMAND = &H112          '系统菜单、控件菜单等的事件
Public Const WM_MOUSEMOVE = &H200           '鼠标移动事件
Public Const WM_LBUTTONDOWN = &H201         '鼠标左键按下
Public Const WM_LBUTTONUP = &H202           '鼠标左键放开
Public Const WM_LBUTTONDBLCLK = &H203       '鼠标左键双击
Public Const WM_RBUTTONDOWN = &H204         '鼠标右键按下
Public Const WM_RBUTTONUP = &H205           '鼠标右键放开
Public Const WM_RBUTTONDBLCLK = &H206       '鼠标右键双击
Public Const WM_MBUTTONDOWN = &H207         '鼠标中键按下
Public Const WM_MBUTTONUP = &H208           '鼠标中键放开
Public Const WM_PARENTNOTIFY = &H210        '子窗体创建、销毁
Public Const WM_EXITSIZEMOVE = &H232        'Resize完毕
Public Const WM_UNDO = &H304&               'UNDO操作
Public Const WM_CUT = &H300&                '剪切
Public Const WM_COPY = &H301&               '复制
Public Const WM_PASTE = &H302&              '粘贴
Public Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_SETTINGCHANGE = &H1A&
Public Const WM_DISPLAYCHANGE = &H7E&

'#########################################################################
' 内存操作函数:

'在堆栈中分配指定字节数的内存，只用于16进制版本的Windows兼容。
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'释放内存，只用于16进制版本的Windows兼容。
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'锁定并返回指向对象内存区域的第一个字节的指针，只用于16进制版本的Windows兼容。
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'改变内存区域大小，只用于16进制版本的Windows兼容。
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
'返回当前对象内存尺寸大小，只用于16进制版本的Windows兼容。
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'减少锁定对象数目，只用于16进制版本的Windows兼容。
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'内存分派属性
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const GMEM_LOWER = GMEM_NOT_BANKED

'将一块内存从一个地方拷贝到另一个地方
'函数原型：
'VOID CopyMemory(
'  PVOID Destination,  // 目标拷贝的地址指针。
'  CONST VOID *Source, // 源拷贝的地址指针。
'  DWORD Length        // 源拷贝的字节大小。
')
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
'作用同上，只是源为一个字符串
Public Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
    
'作用同上，只是目标为一个字符串
Public Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)

'#########################################################################
' 普通 WinAPI 函数:

' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！
'函数原型：
'LRESULT SendMessage(
'  HWND hWnd,      // 目标窗体的句柄。
'  UINT Msg,       // 待发送的消息。
'  WPARAM wParam,  // 消息第一参数。
'  LPARAM lParam   // 消息第二参数。
');
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
'作用同上，不过第二参数为Long型。
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
'作用同上，不过第二参数为String型。
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As String) As Long

'设置窗体状态（最大化、最下化、隐藏等）
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'移动窗体
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'要求窗体刷新
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
'锁定/解锁窗体的刷新
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'销毁窗体及相关资源
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'屏蔽/恢复鼠标及键盘的输入
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'搜索指定条件的窗体
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'改变指定窗体的父窗体
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'获取当前对象所在窗体：
'窗体层次有5层：Frame、Document、Pane、Parent、In-place
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'获取指定窗体的边界矩形尺寸
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'获取客户区域矩形
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'获取窗体属性
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'设置窗体属性
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'移除窗体属性
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'返回包含了指定点的窗口的句柄。
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'将屏幕上某个点的屏幕坐标转换为客户区域坐标
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'将一个窗体相关的坐标空间映射到另一个窗体的坐标空间
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'取消鼠标捕获
Public Declare Function ReleaseCapture Lib "user32" () As Long
'获取鼠标屏幕坐标位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'指定客户区域的一个即将被刷新的矩形区域
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'同上，不过参数2是一个指针了
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'创建指定属性的窗体
Public Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'将消息传送到指定的窗体过程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'设置当前线程消息队列中的窗体获取键盘焦点
Public Declare Function GetFocus Lib "user32" () As Long

'SetWindowPos参数参数：
'表示强制发送 WM_NCCALCSIZE 消息到窗口
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
'在窗口外部绘制一个框架
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'非激活状态
Public Const SWP_NOACTIVATE = &H10
'保持当前位置
Public Const SWP_NOMOVE = &H2
'保持当前尺寸
Public Const SWP_NOSIZE = &H1
'保持当前Z-Order
Public Const SWP_NOZORDER = &H4
'保存父窗体Z-Order
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering


'获取焦点
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'将指定的可执行模块（.DLL/.EXE）映射到调用过程的地址空间
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'减少DLL的引用数目
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


'#########################################################################
' 图形函数分类

'获取窗体显示元素的当前颜色值
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'绘制矩形的一条或者多条边
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'将一个 OLE_COLOR 类型转换为一个 COLORREF 类型。
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'调入一个图标、动态光标或者位图。
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'同上，不过第二参数为一个整形值。
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'表示一个Windows位图格式。
Public Const CF_BITMAP = 2
'3D效果颜色
Public Const LR_LOADMAP3DCOLORS = &H1000
'图片从文件lpsz中调入，而非从资源文件中调入。
Public Const LR_LOADFROMFILE = &H10
'调入透明色
Public Const LR_LOADTRANSPARENT = &H20
'生成 设备无关 DIB 位图，而非设备相关位图。
Public Const IMAGE_BITMAP = 0

'获取显示器或者打印机的信息
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'
'Public Const HORZRES = 8            '  Horizontal width in pixels
'
'Public Const VERTRES = 10           '  Vertical width in pixels
'
'Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
'
'Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
'
'Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
'
'Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
'
'Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units
'
'Public Const PHYSICALWIDTH = 110 '  Physical Width in device units

'设置指定画布的映射模式
Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
'开始一个打印作业
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
'通知打印设备准备接收数据。
Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
'通知打印机停止接收数据，通常用于换页
Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
'完成一次打印作业
Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
'删除指定设备场景（画布）
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'保存当前设备场景状态到上下文堆栈中。
Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
'恢复设备场景状态。
Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
'使用指定坐标指定设备场景视口的原点
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long

'每个逻辑单位为1个设备象素。正X向右，正Y向下。用于SetMapMode()
Public Const MM_TEXT = 1

'乘以两个32位的数，然后用其64位结果除以第三个数，最后四舍五入。
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long


'#########################################################################
' 打印支持:

' VB API Viewer 版本的 DocInfo 结构说明是错误的！！！！
' VB API VIEWER VERSION OF DOCINFO STRUCTURE IS WRONG!
'用于存储 StartDoc() 中文件名及其他信息
Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

'用于初始化打印对话框及返回值
Type PrintDlg
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'调用打印对话框
Public Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As PrintDlg) As Long

'用于 PrintDlg 的对话框的属性描述
Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

'用户点击系统菜单中的“移动”菜单事件
Public Const SC_MOVE = &HF012

'系统默认颜色
Public Const COLOR_WINDOWFRAME = 6  '窗体框架
Public Const COLOR_BTNFACE = 15     '按钮表明
Public Const COLOR_BTNTEXT = 18     '按钮普通文本

'用于程序客户区域绘图信息结构体
Public Type PAINTSTRUCT
   hDC As Long
   fErase As Long
   rcPaint As RECT
   fRestore As Long
   fIncUpdate As Long
   rgbReserved(0 To 31) As Byte
End Type

'定义位图的类型、宽度、高度、颜色格式和位数据块
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'指定窗体进行绘图准备，通过PAINTSTRUCT结构体来初始化。
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'在绘图完成后，标记窗体绘图结束。
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'用于获取给定绘图对象的信息。
'取决于绘图对象的不同，可以在给定缓冲区中填入BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT 或者 LOGPEN 结构
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'将一个对象选入指定的设备场景（画布）中，该对象自动替换掉同一类型的前一对象。
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'删除一个逻辑画笔、画刷、字体、位图、区域或者调色板
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'获取给定窗口或者整个屏幕的画布，用于在上面绘图。
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'释放标准Windows设备场景资源。
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'创建兼容的内存设备场景
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'创建设备相关位图
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'创建指定纯色的逻辑画刷
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'使用指定画刷填充矩形区域
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'从源画布到目标画布的比特块传送其彩色数据
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'返回桌面窗体（屏幕）的句柄
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'获取系统度量单位和系统设置，所有尺寸均以点 Pixel 表示
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'水平滚动条上的矢量位图宽度
Public Const SM_CYHSCROLL = 3
'水平滚动条上的矢量位图高度
Public Const SM_CXVSCROLL = 2


'#########################################################################
'窗体样式:
Public Const WS_CHILD = &H40000000          '子窗体
Public Const WS_HSCROLL = &H100000          '具备水平滚动条
Public Const WS_VSCROLL = &H200000          '具备垂直滚动条
Public Const WS_VISIBLE = &H10000000        '可视
Public Const WS_CLIPCHILDREN = &H2000000    '出去子窗体的父窗体绘图区域
Public Const WS_CLIPSIBLINGS = &H4000000    '更新子窗体时，排除重叠的其他子窗体
Public Const WS_BORDER = &H800000           '具备边框
Public Const WS_TABSTOP = &H10000           'Tab停留
Public Const WS_POPUP = &H80000000          '弹出窗体
Public Const WS_SYSMENU = &H80000           '在标题栏是否具备系统菜单
Public Const WS_THICKFRAME = &H40000        '厚边框
Public Const WS_MINIMIZEBOX = &H20000       '具备最小化按钮
Public Const WS_MAXIMIZEBOX = &H10000       '具备最大化按钮
Public Const WS_DLGFRAME = &H400000         '双边框但是无标题栏的窗体
Public Const WS_EX_TOPMOST = &H8&           '最前面
Public Const WS_EX_CLIENTEDGE = &H200&      '3D效果
Public Const WS_EX_Transparent = &H20&      '窗体透明
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_STATICEDGE = &H20000

Public Const GWL_STYLE = (-16)              'Set the window style
Public Const GWL_EXSTYLE = (-20)            'Set the extended window style
Public Const GWL_USERDATA = (-21)           'Sets the 32-bit value associated with the window.
Public Const GWL_WNDPROC = -4               'Sets a new address for the window procedure.
Public Const GWL_HWNDPARENT = (-8)          'Sets a new application instance handle.

Public Const HWND_TOPMOST = -1              '最前面
Public Const SW_SHOW = 5                    '激活窗体并显示
Public Const SW_HIDE = 0                    '隐藏
Public Const SW_SHOWNORMAL = 1              '还原
Public Const GW_CHILD = 5                   '获取主窗体句柄
Public Const GW_HWNDNEXT = 2                '获取指定窗体Z-Order下的下一窗体的句柄
Public Const CW_USEDEFAULT  As Long = &H80000000        '默认值
Public Const GDI_ERROR = &HFFFF             '出现GDI错误！
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'#########################################################################
' 鼠标激活响应
Public Const MA_ACTIVATE = 1                '激活CWnd
Public Const MA_ACTIVATEANDEAT = 2          '激活CWnd，屏蔽鼠标事件
Public Const MA_NOACTIVATE = 3              '不激活CWnd
Public Const MA_NOACTIVATEANDEAT = 4        '不激活CWnd，屏蔽鼠标事件

Public Const H_MAX As Long = &HFFFF + 1     '最大值
 
'Shell调用
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'同上，不过第4、5参数为Any类型
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long

Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum

Public Const ERROR_FILE_NOT_FOUND = 2&     '文件没有找到
Public Const ERROR_PATH_NOT_FOUND = 3&     '路径没有找到
Public Const ERROR_BAD_FORMAT = 11&        '无效命令
Public Const SE_ERR_ACCESSDENIED = 5       '拒绝存取
Public Const SE_ERR_ASSOCINCOMPLETE = 27   '文件名不完整或无效
Public Const SE_ERR_DDEBUSY = 30           'DDE忙
Public Const SE_ERR_DDEFAIL = 29           'DDE失败
Public Const SE_ERR_DDETIMEOUT = 28        'DDE超时
Public Const SE_ERR_DLLNOTFOUND = 32       '动态链接库没有找到
Public Const SE_ERR_FNF = 2                '文件没有找到
Public Const SE_ERR_NOASSOC = 31           '没有关联程序
Public Const SE_ERR_PNF = 3                '路径没有找到
Public Const SE_ERR_OOM = 8                '内存溢出
Public Const SE_ERR_SHARE = 26             '共享违例

'判断当前是否某个虚拟键按下或者放开
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' 虚拟键编码常数
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt

'人工合成鼠标动作和点击事件，新标准应该使用 SendInput() 函数！
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000  '绝对移动
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4       '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1         '鼠标移动
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10     '  right button up

'用于异步输入/输出 I/O
Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

'最长路径名
Public Const OFS_MAXPATHNAME = 128

'用于 OpenFile 的文件信息
Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

'#########################################################################
'流的支持:

'写入文件
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'打开文件
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'读取文件
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'关闭对象句柄
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const OF_CANCEL = &H800
Public Const OF_CREATE = &H1000
Public Const OF_DELETE = &H200
Public Const OF_EXIST = &H4000
Public Const OF_PARSE = &H100
Public Const OF_PROMPT = &H2000
Public Const OF_REOPEN = &H8000
Public Const OF_SHARE_COMPAT = &H0
Public Const OF_SHARE_DENY_NONE = &H40
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_VERIFY = &H400
Public Const OF_WRITE = &H1
Public Const OF_READ = &H0
Public Const OF_READWRITE = &H2


'#########################################################################
'API作图
'#########################################################################
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

'样式
Public Const BS_HATCHED = 2
Public Const BS_NULL = 1
Public Const BS_SOLID = 0

'底纹
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||

Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const PS_DOT = 2
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_INSIDEFRAME = 6
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086

'判断矩形与矩形、矩形与椭圆是否相交
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long


Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4

Public Const NULLREGION = 1
Public Const SIMPLEREGION = 2
Public Const COMPLEXREGION = 3

Public Const ALTERNATE = 1
Public Const WINDING = 2

Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

'######################################################################################
'获取中英文混合字符串长度
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'繁简转换
Public Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

'######################################################################################

Public Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

Public Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Public Declare Function GradientFill Lib "msimg32" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_TRIANGLE, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'鼠标位置信息
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Type Size
    cx As Long
    cy As Long
End Type
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Used for creating the temporary WMF file
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MM_ANISOTROPIC = 8 ' Map mode anisotropic
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
'VB Errors
Public Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Public Const DSna = &H220326

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态
    ' Virtual key values
Public Const VK_TAB = &H9

Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, LParam As Any) As Long

'######################################################################################
'   鼠标滚动钩子
'######################################################################################

Public Type POINTL
    X As Long
    Y As Long
End Type

Public Const WM_MOUSEWHEEL = &H20A

Public lpPrevWndProc As Long

Public sngX As Single, sngY As Single   '鼠标坐标
Public intShift As Integer              '鼠标按键
Public bWay As Boolean                  '鼠标方向
Public bMouseFlag As Boolean            '鼠标事件激活标志

'######################################################################################
'   获取字符屏幕位置
'######################################################################################
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const S_FALSE = &H1
Public Const S_OK = &H0

Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

'######################################################################################
'   直接发送按键的函数
'######################################################################################
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'######################################################################################
'   输入法处理函数
'######################################################################################
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'######################################################################################
'   释放内存
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

'######################################################################################
'表格相关GDI声明和辅助函数
'######################################################################################
Public Const BITSPIXEL = 12         '  Number of bits per pixel
Public Const DT_NOFULLWIDTHCHARBREAK = &H80000
Public Const DT_HIDEPREFIX = &H100000
Public Const DT_PREFIXONLY = &H200000
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Public Const LF_FACESIZE = 32

Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const BALTIC_CHARSET = 186
Public Const RUSSIAN_CHARSET = 204
Public Const THAI_CHARSET = 222
Public Const EE_CHARSET = 238
Public Const OEM_CHARSET = 255

'字体属性
Public Type LOGFONT
    lfHeight As Long         ' 字体尺寸 (见下面)
    lfWidth As Long          ' 通常你无需设置,让Windows创建默认的
    lfEscapement As Long     ' 角度,采用0.1度为单位
    lfOrientation As Long    ' 请采用默认值
    lfWeight As Long         ' 粗体、超粗、常规等        FW_DONTCARE/FW_THIN/FW_EXTRALIGHT/FW_ULTRALIGHT/FW_LIGHT/...
    lfItalic As Byte         ' 斜体
    lfUnderline As Byte      ' 下划线
    lfStrikeOut As Byte      ' 删除线
    lfCharSet As Byte        ' 字符集        ANSI_CHARSET/CHINESEBIG5_CHARSET/GB2312_CHARSET/SYMBOL_CHARSET/...
    lfOutPrecision As Byte   ' 请采用默认值
    lfClipPrecision As Byte  ' 请采用默认值
    lfQuality As Byte        ' 请采用默认值
    lfPitchAndFamily As Byte ' 请采用默认值
    lfFaceName(LF_FACESIZE) As Byte  ' 转化为数组的字体名称
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImagelist As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImagelist As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long '判断指定点是否在指定矩形中！！！


Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
   ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
     Public Const R2_BLACK = 1 ' 0
     Public Const R2_COPYPEN = 13 ' P
     Public Const R2_LAST = 16
     Public Const R2_MASKNOTPEN = 3 ' DPna
     Public Const R2_MASKPEN = 9 ' DPa
     Public Const R2_MASKPENNOT = 5 ' PDna
     Public Const R2_MERGENOTPEN = 12    ' DPno
     Public Const R2_MERGEPEN = 15 ' DPo
     Public Const R2_MERGEPENNOT = 14    ' PDno
     Public Const R2_NOP = 11    ' D
     Public Const R2_NOT = 6 ' Dn
     Public Const R2_NOTCOPYPEN = 4 ' PN
     Public Const R2_NOTMASKPEN = 8 ' DPan
     Public Const R2_NOTMERGEPEN = 2 ' DPon
     Public Const R2_NOTXORPEN = 10 ' DPxn
     Public Const R2_WHITE = 16 ' 1
     Public Const R2_XORPEN = 7 ' DPx

Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0

Public Const DEFAULT_QUALITY = 0           ' 默认字体质量
Public Const DRAFT_QUALITY = 1             ' Appearance is less important that PROOF_QUALITY.
Public Const PROOF_QUALITY = 2             ' 最佳字符质量
Public Const NONANTIALIASED_QUALITY = 3    ' Don't smooth font edges even if system is set to smooth font edges
Public Const ANTIALIASED_QUALITY = 4       ' Ensure font edges are smoothed if system is set to smooth font edges
Public Const CLEARTYPE_QUALITY = 5

Public Const DEFAULT_PITCH = 0

Public Const CLR_INVALID = -1

'修正的 DrawState 函数声明：
'DrawState:用于显示不同状态的图像，比如“浮雕”、“抖动”、“单色”等效果
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hDC As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal LParam As Long, _
   ByVal wParam As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long
   
'DrawStateString:同上
Public Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hDC As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' 绘图状态常数声明：
'/* 图像类型 */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* 状态类型 */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

' 创建一个 ImageList
Public Declare Function ImageList_Create Lib "comctl32.dll" ( _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal fMask As Long, _
        ByVal cInitial As Long, _
        ByVal cGrow As Long _
    ) As Long
Public Const ILC_MASK = 1&
Public Const ILC_COLOR = 0&
Public Const ILC_COLORDDB = &HFE&
Public Const ILC_COLOR4 = &H4&
Public Const ILC_COLOR8 = &H8&
Public Const ILC_COLOR16 = &H10&
Public Const ILC_COLOR24 = &H18&
Public Const ILC_COLOR32 = &H20&
Public Const ILC_PALETTE = &H800&

Public Declare Function ImageList_Destroy Lib "comctl32.dll" ( _
        ByVal hIml As Long _
    ) As Long

' 添加一幅遮罩位图到ImageList
Public Declare Function ImageList_AddMasked Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal crMask As Long _
    ) As Long
    
' 基于一个ImageList图标创建一个新的图标
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
    
' 在一个ImageList中绘制一个项目
Public Declare Function ImageList_Draw Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal fStyle As Long _
    ) As Long
    
' 同上，不过在位置和颜色上有更多控制
Public Declare Function ImageList_DrawEx Lib "comctl32.dll" ( _
      ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal rgbBk As Long, _
      ByVal rgbFg As Long, _
      ByVal fStyle As Long _
   ) As Long
   
' ImageList_Draw 方法的常量：
Public Const ILD_NORMAL = 0            '采用ImageList的背景色绘图
Public Const ILD_TRANSPARENT = 1       '采用遮罩色来绘制透明位图
Public Const ILD_BLEND25 = 2           '采用遮罩色来绘制25%透明度的位图
Public Const ILD_SELECTED = 4          '采用遮罩色来绘制50%透明度的位图
Public Const ILD_FOCUS = 4             '同25%透明度
Public Const ILD_OVERLAYMASK = 3840    '重叠图像
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' 标准的 GDI 绘制图标或者光标的函数：
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Const DI_MASK = &H1         '采用遮罩来绘制图标或光标
Public Const DI_IMAGE = &H2        '采用图像来绘制图标或光标
Public Const DI_NORMAL = &H3       '组合图像和遮罩
Public Const DI_COMPAT = &H4       '用系统默认图像
Public Const DI_DEFAULTSIZE = &H8  '采用系统默认大小

Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
' XP检测
Public Declare Function GetVersion Lib "kernel32" () As Long   '获取当前系统版本号

Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lhDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCMBUTTONDOWN = &HA7
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

' Edit 控件消息：
Public Const EM_GETSEL = &HB0&              '获取当前选中区域的开始和结束字符位置。不能大于65, 535。
Public Const EM_SETSEL = &HB1&              '选择某一范围内容。
Public Const EM_GETRECT = &HB2&             '获取一个Edit控件的格式化矩形区域。
Public Const EM_SETRECT = &HB3&             '设置Edit控件的格式化矩形区域，同时重绘文本。
Public Const EM_SETRECTNP = &HB4&           '同上，但是不重绘文本。
Public Const EM_SCROLL = &HB5&              '垂直滚动消息。
Public Const EM_LINESCROLL = &HB6&          '水平或垂直滚动文本。
Public Const EM_SCROLLCARET = &HB7&         '光标滚动为可视。
Public Const EM_GETMODIFY = &HB8&           '判断是否内容被修改了。
Public Const EM_SETMODIFY = &HB9&           '设置或清除内容修改标志。
Public Const EM_GETLINECOUNT = &HBA&        '获取行数。
Public Const EM_LINEINDEX = &HBB&           '获取某行的字符索引值（从文本头开始）。
Public Const EM_SETHANDLE = &HBC&           '设置多行Edit控件的内存句柄。
Public Const EM_GETHANDLE = &HBD&           '获取当前Edit控件的内存句柄。
Public Const EM_GETTHUMB = &HBE&            '获取当前滚动条位置。
Public Const EM_LINELENGTH = &HC1&          '获取某行的字符长度。
Public Const EM_REPLACESEL = &HC2&          '替换当前选中区域文本。
Public Const EM_GETLINE = &HC4&             '发送一行文本到指定缓冲区。
Public Const EM_LIMITTEXT = &HC5&           '限制用户输入的文本总数。
Public Const EM_CANUNDO = &HC6&             '是否可以响应 EM_UNDO 消息。
Public Const EM_UNDO = &HC7&                'Undo消息。
Public Const EM_FMTLINES = &HC8&            '设置软回车符是否启用。
Public Const EM_LINEFROMCHAR = &HC9&        '获取指定字符索引值的行数。
Public Const EM_SETTABSTOPS = &HCB&         '设置制表符位置数组。
Public Const EM_SETPASSWORDCHAR = &HCC&     '设置密码屏蔽字符。
Public Const EM_EMPTYUNDOBUFFER = &HCD&     '清空Undo队列。
Public Const EM_GETFIRSTVISIBLELINE = &HCE& '最上面的可视行的行索引（多行），或者最左边字符索引（单行）。
Public Const EM_SETREADONLY = &HCF&         '只读。
Public Const EM_SETWORDBREAKPROC = &HD0&    '自定义断字处理过程。
Public Const EM_GETWORDBREAKPROC = &HD1&    '获取当前断字处理过程地址。
Public Const EM_GETPASSWORDCHAR = &HD2&     '获取密码屏蔽字符。
'#if(WINVER >= =&H0400)
Public Const EM_SETMARGINS = &HD3&          '设置左、右间距，并刷新。
Public Const EM_GETMARGINS = &HD4&          '获取...
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT '设置字符最大长度。 ' /* ;win40 Name change */
Public Const EM_GETLIMITTEXT = &HD5&        '获取字符最大长度。
Public Const EM_POSFROMCHAR = &HD6&         '获取指定字符的坐标(X,Y)。
Public Const EM_CHARFROMPOS = &HD7&         '获取指定坐标点附近的字符。

Public Const EC_LEFTMARGIN = &H1            '表示是设置左边界。
Public Const EC_RIGHTMARGIN = &H2           '表示是设置右边界。
Public Const EC_USEFONTINFO = &HFFFF&       '边界采用字符宽度。

Public Const EM_EXGETSEL = (WM_USER + 52)       '获取选中的起始与终止字符位置。
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetTextJustification Lib "gdi32" (ByVal hDC As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Public Const ES_SUNKEN = &H4000&                '凹下效果
Public Const ES_NOHIDESEL = &H100&      '失去焦点时保持选择内容。

'################################################################################################################
'## 图片缩放模式设置
'######################################################################################
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE

Public Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Public Const AC_SRC_OVER = &H0
' AlphaFormat:
Public Const AC_SRC_ALPHA = &H1

Public Declare Function AlphaBlend Lib "MSIMG32.dll" ( _
  ByVal hDCDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long
