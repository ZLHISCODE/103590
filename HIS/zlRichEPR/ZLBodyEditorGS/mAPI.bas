Attribute VB_Name = "mAPI"
'#########################################################################
'##模 块 名：mAPI.bas
'##创 建 人：吴庆伟
'##日    期：2005年3月25日
'##修 改 人：
'##日    期：
'##描    述：通用 Windows API 声明
'##版    本：
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
    x As Long
    y As Long
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
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'作用同上，不过第二参数为Long型。
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'作用同上，不过第二参数为String型。
Public Declare Function SendMessageStr Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'设置窗体状态（最大化、最下化、隐藏等）
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'移动窗体
Public Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'要求窗体刷新
Public Declare Function UpdateWindow Lib "User32" (ByVal hWnd As Long) As Long
'锁定/解锁窗体的刷新
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
'销毁窗体及相关资源
Public Declare Function DestroyWindow Lib "User32" (ByVal hWnd As Long) As Long
'屏蔽/恢复鼠标及键盘的输入
Public Declare Function EnableWindow Lib "User32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'搜索指定条件的窗体
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'改变指定窗体的父窗体
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'获取当前对象所在窗体：
'窗体层次有5层：Frame、Document、Pane、Parent、In-place
Public Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'获取指定窗体的边界矩形尺寸
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
'获取客户区域矩形
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
'获取窗体属性
Public Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'设置窗体属性
Public Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'移除窗体属性
Public Declare Function RemoveProp Lib "User32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'返回包含了指定点的窗口的句柄。
Public Declare Function WindowFromPointXY Lib "User32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'将屏幕上某个点的屏幕坐标转换为客户区域坐标
Public Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'将一个窗体相关的坐标空间映射到另一个窗体的坐标空间
Public Declare Function MapWindowPoints Lib "User32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
'取消鼠标捕获
Public Declare Function ReleaseCapture Lib "User32" () As Long
'获取鼠标屏幕坐标位置
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
'指定客户区域的一个即将被刷新的矩形区域
Public Declare Function InvalidateRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'同上，不过参数2是一个指针了
Public Declare Function InvalidateRectAsNull Lib "User32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'创建指定属性的窗体
Public Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'将消息传送到指定的窗体过程
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'设置当前线程消息队列中的窗体获取键盘焦点
Public Declare Function GetFocus Lib "User32" () As Long

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
Public Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'将指定的可执行模块（.DLL/.EXE）映射到调用过程的地址空间
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'减少DLL的引用数目
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Integer, ByVal crColor As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'#########################################################################
' 图形函数分类

'获取窗体显示元素的当前颜色值
Public Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
'绘制矩形的一条或者多条边
Public Declare Function DrawEdge Lib "User32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'将一个 OLE_COLOR 类型转换为一个 COLORREF 类型。
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'调入一个图标、动态光标或者位图。
Public Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'同上，不过第二参数为一个整形值。
Public Declare Function LoadImageLong Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

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
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const HORZRES = 8            '  Horizontal width in pixels

Public Const VERTRES = 10           '  Vertical width in pixels

Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X

Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin

Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin

Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units

'设置指定画布的映射模式
Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
'开始一个打印作业
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
'通知打印设备准备接收数据。
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
'通知打印机停止接收数据，通常用于换页
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
'完成一次打印作业
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
'删除指定设备场景（画布）
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'保存当前设备场景状态到上下文堆栈中。
Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
'恢复设备场景状态。
Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
'使用指定坐标指定设备场景视口的原点
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long

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
    hdc As Long
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
   hdc As Long
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
Public Declare Function BeginPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'在绘图完成后，标记窗体绘图结束。
Public Declare Function EndPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'用于获取给定绘图对象的信息。
'取决于绘图对象的不同，可以在给定缓冲区中填入BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT 或者 LOGPEN 结构
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'将一个对象选入指定的设备场景（画布）中，该对象自动替换掉同一类型的前一对象。
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'删除一个逻辑画笔、画刷、字体、位图、区域或者调色板
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'获取给定窗口或者整个屏幕的画布，用于在上面绘图。
Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
'释放标准Windows设备场景资源。
Public Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'创建兼容的内存设备场景
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'创建设备相关位图
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'创建指定纯色的逻辑画刷
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'使用指定画刷填充矩形区域
Public Declare Function FillRect Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'从源画布到目标画布的比特块传送其彩色数据
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'返回桌面窗体（屏幕）的句柄
Public Declare Function GetDesktopWindow Lib "User32" () As Long
'获取系统度量单位和系统设置，所有尺寸均以点 Pixel 表示
Public Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
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
Public Const WS_DISABLED = &H8000000        '不可用

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
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer

' 虚拟键编码常数
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt

'人工合成鼠标动作和点击事件，新标准应该使用 SendInput() 函数！
Declare Sub mouse_event Lib "User32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
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
    Offset As Long
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
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

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
'In a module
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

Declare Function DrawTextEx Lib "User32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal N As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

'######################################################################################
'   释放内存
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "Psapi" (ByVal hProcess As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
'################################################################################################################
'## 图片缩放模式设置
'######################################################################################
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

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

'颜色转换
Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
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







