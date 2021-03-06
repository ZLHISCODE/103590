Attribute VB_Name = "mAPI"
'##描    述： Windows API 声明
'##版    本：
Option Explicit
Public Enum eilIconState
  Normal = 0
  Disabled = 1
End Enum

Public Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
End Enum

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

Public Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

'GetDeviceCaps()函数的参数常数
Public Const DRIVERVERSION = 0      '设备驱动程序版本
Public Const TECHNOLOGY = 2         '设备工艺
Public Const HORZSIZE = 4           '物理屏幕宽度，单位：毫米。
Public Const VERTSIZE = 6           '物理屏幕高度，单位：毫米。
Public Const HORZRES = 8            '屏幕宽度，单位：象素（pixels）。
Public Const VERTRES = 10           '屏幕高度，单位：（光栅）行。
Public Const BITSPIXEL = 12         '每个象素点的相邻颜色位数。
Public Const PLANES = 14            '颜色平面数。
Public Const NUMBRUSHES = 16        '设备相关画刷(BRUSH)数目。
Public Const NUMPENS = 18           '设备相关画笔(PEN)数目。
Public Const NUMMARKERS = 20        '设备相关标记数目。
Public Const NUMFONTS = 22          '设备相关字体数目。
Public Const NUMCOLORS = 24         '设备颜色表的入口总数，如果设备的颜色深度小于每象素8位时可用。大于该色深时，返回-1。
Public Const PDEVICESIZE = 26       '保留。
Public Const CURVECAPS = 28         '设备的曲线特性。
Public Const LINECAPS = 30          '设备的线条特性。
Public Const POLYGONALCAPS = 32     '设备的多边形特性。
Public Const TEXTCAPS = 34          '设备的文本特性。
Public Const CLIPCAPS = 36          '设备剪切性能标志，如果设备可以剪切为矩形，返回1；否则为0。
Public Const RASTERCAPS = 38        '设备的光栅特性。
Public Const ASPECTX = 40           '绘制线条时的相对象素宽度。
Public Const ASPECTY = 42           '绘制线条时的相对象素高度。
Public Const ASPECTXY = 44          '绘制线条时的相对对角线象素宽度。
Public Const SHADEBLENDCAPS = 45    '设备的阴影及混合特性。
'下面3个索引值只能在设备驱动在RASTERCAPS等于RC_PALETTE位并且在兼容16位Windows时才可用。
Public Const SIZEPALETTE = 104      '设备调色板的入口总数。
Public Const NUMRESERVED = 106      '系统调色板的保留入口总数。
Public Const COLORRES = 108         '设备的实际颜色分辨率，单位：BPP（位/象素）。
'打印相关常量，这些值将替换对应的转移符
Public Const PHYSICALWIDTH = 110    '对于打印设备而言，表示物理页宽，采用设备单位。注：物理页总是大于页面的可打印区域，不会小于它。
Public Const PHYSICALHEIGHT = 111   '对于打印设备而言，表示物理页高，采用设备单位。
Public Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Public Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Public Const SCALINGFACTORX = 114   '打印机的X－轴缩放比例。
Public Const SCALINGFACTORY = 115   '打印机的Y－轴缩放比例。
'显示设备相关常量
Public Const VREFRESH = 116         '对显示设备而言，表示当前的垂直刷新率，单位：Hz。
Public Const DESKTOPVERTRES = 117   '整个桌面的宽度，单位：Pixels
Public Const DESKTOPHORZRES = 118   '整个桌面的高度，单位：Pixels
Public Const BLTALIGNMENT = 119     '默认 blt 对齐方式
'设备工艺
Public Const DT_PLOTTER = 0         '矢量绘图仪
Public Const DT_RASDISPLAY = 1      '光栅显示器
Public Const DT_RASPRINTER = 2      '光栅打印机
Public Const DT_RASCAMERA = 3       '光栅照相机
Public Const DT_CHARSTREAM = 4      '字符流
Public Const DT_METAFILE = 5        '图元文件
Public Const DT_DISPFILE = 6        '显示文件
'设备的曲线特性。
Public Const CC_NONE = 0            '设备不支持曲线。
Public Const CC_CIRCLES = 1         '设备可以绘制弦弧。
Public Const CC_PIE = 2             '设备可以绘制圆。
Public Const CC_CHORD = 4           '设备可以绘制椭圆。
Public Const CC_ELLIPSES = 8        '设备可以绘制椭圆。
Public Const CC_WIDE = 16           '设备可以绘制宽边框。
Public Const CC_STYLED = 32         '设备可以绘制样式边框。
Public Const CC_WIDESTYLED = 64     '设备可以绘制宽样式边框。
Public Const CC_INTERIORS = 128     '设备可以绘制内部区域。
Public Const CC_ROUNDRECT = 256     '设备可以绘制圆角矩形。
'设备的线条特性。
Public Const LC_NONE = 0            '设备不支持线条。
Public Const LC_POLYLINE = 2        '设备可以绘制折线。
Public Const LC_MARKER = 4          '设备可以绘制一个标记。
Public Const LC_POLYMARKER = 8      '设备可以绘制多个标记。
Public Const LC_WIDE = 16           '设备可以绘制宽线条。
Public Const LC_STYLED = 32         '设备可以绘制样式线条。
Public Const LC_WIDESTYLED = 64     '设备可以绘制宽样式线条。
Public Const LC_INTERIORS = 128     '设备可以绘制内部区域。
'设备的多边形特性。
Public Const PC_NONE = 0            '设备不支持多边形。
Public Const PC_POLYGON = 1         '设备可以绘制交替填充的多边形。
Public Const PC_RECTANGLE = 2       '设备可以绘制矩形。
Public Const PC_WINDPOLYGON = 4     '设备可以绘制螺旋填充的多边形。
Public Const PC_TRAPEZOID = 4       '设备可以绘制不规则四边形。
Public Const PC_SCANLINE = 8        '设备可以绘制设备可以绘制单扫描线。
Public Const PC_WIDE = 16           '设备可以绘制宽边框。
Public Const PC_STYLED = 32         '设备可以绘制样式边框。
Public Const PC_WIDESTYLED = 64     '设备可以绘制宽样式边框。
Public Const PC_INTERIORS = 128     '设备可以绘制内部区域。
Public Const PC_POLYPOLYGON = 256   '设备可以绘制多个多边形。
Public Const PC_PATHS = 512         '设备可以绘制路径。
'裁剪特性
Public Const CP_NONE = 0            '输出不裁剪
Public Const CP_RECTANGLE = 1       '输出裁剪至矩形
Public Const CP_REGION = 2          '作废
'文本特性
Public Const TC_OP_CHARACTER = &H1  '设备满足字符输出精度。
Public Const TC_OP_STROKE = &H2     '设备满足笔画输出精度。
Public Const TC_CP_STROKE = &H4     '设备满足笔画裁剪精度。
Public Const TC_CR_90 = &H8         '设备可以90度字符旋转。
Public Const TC_CR_ANY = &H10       '设备可以任意字符旋转。
Public Const TC_SF_X_YINDEP = &H20  '设备可以在X轴和Y轴独立缩放。
Public Const TC_SA_DOUBLE = &H40    '设备支持2倍字符缩放。
Public Const TC_SA_INTEGER = &H80   '设备只能采用字符的整数倍缩放。
Public Const TC_SA_CONTIN = &H100   '设备可以采用字符的任意倍数缩放。
Public Const TC_EA_DOUBLE = &H200   '设备可以绘制双倍磅值的字符。
Public Const TC_IA_ABLE = &H400     '设备支持斜体。
Public Const TC_UA_ABLE = &H800     '设备支持下划线。
Public Const TC_SO_ABLE = &H1000    '设备支持删除线。
Public Const TC_RA_ABLE = &H2000    '设备支持光栅字体。
Public Const TC_VA_ABLE = &H4000    '设备支持矢量字体。
Public Const TC_RESERVED = &H8000   '保留；必须为0。
Public Const TC_SCROLLBLT = &H10000 '文本不允许卷动。
'光栅特性
Public Const RC_NONE = 0                '
Public Const RC_BITBLT = 1              '可以传递位图。
Public Const RC_BANDING = 2             '需要条带化(Banding)支持。
Public Const RC_SCALING = 4             '支持缩放。
Public Const RC_BITMAP64 = 8            '可以支持大于64KB的位图。
Public Const RC_GDI20_OUTPUT = &H10     '
Public Const RC_GDI20_STATE = &H20      '
Public Const RC_SAVEBITMAP = &H40       '
Public Const RC_DI_BITMAP = &H80        '支持SetDIBits和GetDIBits函数。
Public Const RC_PALETTE = &H100         '指定一个基于调色板的设备。
Public Const RC_DIBTODEV = &H200        '支持SetDIBitsToDevice函数。
Public Const RC_BIGFONT = &H400         '支持大于64K的字体。
Public Const RC_STRETCHBLT = &H800      '支持StretchBlt函数。
Public Const RC_FLOODFILL = &H1000      '可以执行flood fills填充操作。
Public Const RC_STRETCHDIB = &H2000     '支持StretchDIBits函数。
Public Const RC_OP_DX_OUTPUT = &H4000
Public Const RC_DEVBITS = &H8000
'设备的阴影及混合特性。
'#define SB_PREMULT_ALPHA    0x00000004
Public Const SB_NONE = &H0              '
Public Const SB_CONST_ALPHA = &H1       '
Public Const SB_PIXEL_ALPHA = &H2       '
Public Const SB_PREMULT_ALPHA = &H4     '
Public Const SB_GRAD_RECT = &H10              '
Public Const SB_GRAD_TRI = &H20              '
'WinNT自定义纸张控制================================================================
'注意以dmFields是Long型,as Long或尾部加&符
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Public Const IDOK = 1
Public Const IDCANCEL = 2
'Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32

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
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_CONTEXTMENU = &H7B&     '通知窗体的右键点击事件
Public Const WM_PRINTCLIENT = &H318&    '请求绘制其客户区域到一个指定的设备上下文中，通常是指打印机。
Public Const EM_CANPASTE = (WM_USER + 50)       '决定是否可以粘贴指定格式的剪贴板内容。
Public Const EM_DISPLAYBAND = (WM_USER + 51)    '显示RTB内容的一部分矩形区域，该区域由 EM_FORMATRANGE 消息格式化一个设备来设置。裁剪区域由该矩形决定。
Public Const EM_EXGETSEL = (WM_USER + 52)       '获取选中的起始与终止字符位置。
Public Const EM_EXLIMITTEXT = (WM_USER + 53)    '设置用户可以敲入或者粘贴进RTB中的文本总数上限。OLE对象视为一个字符，默认为32K。
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54) '判断是哪一行包含指定字符。
Public Const EM_EXSETSEL = (WM_USER + 55)       '选中一定范围的字符或者OLE对象。
Public Const EM_FINDTEXT = (WM_USER + 56)       '查找文本。
Public Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Public Const EM_GETCHARFORMAT = (WM_USER + 58)  '判断默认字符格式或者当前范围第一个字符的格式。
Public Const EM_GETEVENTMASK = (WM_USER + 59)   '获取事件掩码。
Public Const EM_GETOLEINTERFACE = (WM_USER + 60) '获取一个OLE对象，客户端用来访问该OLE对象的功能。此时会先调用AddRef() 增加一个引用，用户需要在用完后调用Release() 函数。
Public Const EM_GETPARAFORMAT = (WM_USER + 61)  '获取当前区域的第一个段落的段落属性。
Public Const EM_GETSELTEXT = (WM_USER + 62)     '获取当前选中的文本。请确保缓冲区可以容纳该文本。
Public Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Public Const EM_PASTESPECIAL = (WM_USER + 64)   '选择性粘贴。
Public Const EM_REQUESTRESIZE = (WM_USER + 65)  '通知父窗体改变尺寸，对无底控件很有用！
Public Const EM_SELECTIONTYPE = (WM_USER + 66)  '判断选中区域的类型，是文本还是OLE对象，或者多个OLE/文本对象。
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)  '设置RTB背景色。
Public Const EM_SETCHARFORMAT = (WM_USER + 68)  '设置字符格式。
Public Const EM_SETEVENTMASK = (WM_USER + 69)   '设置事件掩码。
Public Const EM_SETOLECALLBACK = (WM_USER + 70) '提供一个IRichEditOleCallback 对象给RTB，用于从客户端获取OLE相关资源和信息。
Public Const EM_SETPARAFORMAT = (WM_USER + 71)  '设置段落格式。
Public Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Public Const EM_STREAMIN = (WM_USER + 73)       '流式输入（读取）。使用应用程序提供的EditStreamCallback回调函数提供的数据流替换RTB内容。
Public Const EM_STREAMOUT = (WM_USER + 74)      '流式输出（写入）到某一文件或指定位置。
Public Const EM_GETTEXTRANGE = (WM_USER + 75)   '返回一个指定文本的选择区域。
Public Const EM_FINDWORDBREAK = (WM_USER + 76)  '获取前一/后一断字位置，或者获取当前位置字符信息。
Public Const EM_SETOPTIONS = (WM_USER + 77)     'RTB选项设置。如“双击自动选中单词”、“自动滚动条”等。
Public Const EM_GETOPTIONS = (WM_USER + 78)     '获取RTB选项。
Public Const EM_FINDTEXTEX = (WM_USER + 79)     '查找文本。
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80) '获取当前注册的扩展断字处理过程的地址。
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81) '设置当前扩展断字处理过程。0则恢复为默认。
Public Const EM_OUTLINE = (WM_USER + 220)
Public Const EM_GETZOOM = (WM_USER + 224)
Public Const EM_SETZOOM = (WM_USER + 225)
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
Public Const EM_SETMARGINS = &HD3&          '设置左、右间距，并刷新。
Public Const EM_GETMARGINS = &HD4&          '获取...
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT '设置字符最大长度。 ' /* ;win40 Name change */
Public Const EM_GETLIMITTEXT = &HD5&        '获取字符最大长度。
Public Const EM_POSFROMCHAR = &HD6&         '获取指定字符的坐标(X,Y)。
Public Const EM_CHARFROMPOS = &HD7&         '获取指定坐标点附近的字符。
Public Const EC_LEFTMARGIN = &H1            '表示是设置左边界。
Public Const EC_RIGHTMARGIN = &H2           '表示是设置右边界。
Public Const EC_USEFONTINFO = &HFFFF&       '边界采用字符宽度。
' * Edit 控件样式
Public Const ES_LEFT = &H0&             '左对齐
Public Const ES_CENTER = &H1&           '居中
Public Const ES_RIGHT = &H2&            '右对齐
Public Const ES_MULTILINE = &H4&        '多行
Public Const ES_UPPERCASE = &H8&        '大写
Public Const ES_LOWERCASE = &H10&       '小写
Public Const ES_PASSWORD = &H20&        '密码
Public Const ES_AUTOVSCROLL = &H40&     '自动垂直滚动
Public Const ES_AUTOHSCROLL = &H80&     '自动水平滚动10个字符
Public Const ES_NOHIDESEL = &H100&      '失去焦点时保持选择内容。
Public Const ES_OEMCONVERT = &H400&     '
Public Const ES_READONLY = &H800&       '只读
Public Const ES_WANTRETURN = &H1000&    '回车键换行。否则回车等同于窗体中默认按钮事件。
Public Const ES_NUMBER = &H2000&        '只允许输入数字。
'/* Edit 控件通知消息 */
Public Const EN_CHANGE = &H300          '内容改变。父窗体通过 WM_COMMAND 消息获取该通知。
Public Const EN_ERRSPACE = &H500        '内容不足以分配该操作。
Public Const EN_HSCROLL = &H601         '水平滚动事件。
Public Const EN_KILLFOCUS = &H200       '失去焦点事件。
Public Const EN_MAXTEXT = &H501         '输入的文本超过最大字符数。或者在非自动滚动时超出控件可视区域。
Public Const EN_SETFOCUS = &H100        '获得键盘输入焦点。
Public Const EN_UPDATE = &H400          '在用户改变内容但是还没有刷新显示时发出该通知。用户可以用于调节控件尺寸以适应内容。
Public Const EN_VSCROLL = &H602         '垂直滚动事件。

'补充消息：2006/5/28
Public Const EM_GETSCROLLPOS = WM_USER + 221
Public Const EM_SETSCROLLPOS = WM_USER + 222


Public Const LF_FACESIZE = 32   '字体名称字节长度。
Public Const RICHEDIT_VER = &H210    '当前Rich Edit控件版本号
Public Const cchTextLimitDefault = 32767&       '默认文本长度限制
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0
Public Const RICHEDIT_CLASSW = "RichEdit20W"
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSW       'UNICODE版本！


' /* Richedit v2.0 消息 */
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)   '设置Undo数量上限。
Public Const EM_REDO = (WM_USER + 84)           'Redo操作。
Public Const EM_CANREDO = (WM_USER + 85)        '判断Redo队列中是否有任何动作，用而决定是否可以Redo。
Public Const EM_GETUNDONAME = (WM_USER + 86)    '给出下一个Undo操作的名称。该名称由 UNDONAMEID 枚举常量定义！
Public Const EM_GETREDONAME = (WM_USER + 87)    '给出下一个Redo操作的名称。
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)    '停止当前Undo队列的字符搜集。任何击键记入下一队列。

Public Const EM_SETTEXTMODE = (WM_USER + 89)    '设置文本模式和Undo等级。如果RTB包含任何字符，则该消息不起作用！
Public Const EM_GETTEXTMODE = (WM_USER + 90)    '获取当前文本模式和Undo等级。

Public Const EM_FINDTEXTW = (WM_USER + 123)     '查找Unicode的文本。
Public Const EM_FINDTEXTEXW = (WM_USER + 124)   '同上。

' /* enum for use with EM_GET/SETTEXTMODE */    文本模式
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                 ' /* 默认行为 */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8           ' /* 默认行为 */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32           ' /* 默认行为 */
End Enum

Public Const EM_AUTOURLDETECT = (WM_USER + 91)      '启用/禁用自动URL检测。
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)   '判断是否启用了自动URL检测。
Public Const EM_SETPALETTE = (WM_USER + 93)         '改变调色板。
Public Const EM_GETTEXTEX = (WM_USER + 94)          '获取指定代码页的文本。
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)    '采用不同方式计算文本长度。

' /* 远东特殊消息 */
Public Const EM_SETPUNCTUATION = (WM_USER + 100)    '设置标点符号。仅用于亚洲语言的操作系统。
Public Const EM_GETPUNCTUATION = (WM_USER + 101)    '获取标点符号。仅用于亚洲语言的操作系统。
Public Const EM_SETWORDWRAPMODE = (WM_USER + 102)   '设置自动换行与断字选项。仅用于亚洲语言的操作系统。
Public Const EM_GETWORDWRAPMODE = (WM_USER + 103)   '获取自动换行与断字选项。仅用于亚洲语言的操作系统。
Public Const EM_SETIMECOLOR = (WM_USER + 104)       '设置IME组合颜色。仅用于亚洲语言的操作系统。
Public Const EM_GETIMECOLOR = (WM_USER + 105)       '获取IME组合颜色。仅用于亚洲语言的操作系统。
Public Const EM_SETIMEOPTIONS = (WM_USER + 106)     '设置IME选项。仅用于亚洲语言的操作系统。
Public Const EM_GETIMEOPTIONS = (WM_USER + 107)     '获取IME选项。仅用于亚洲语言的操作系统。
Public Const EM_CONVPOSITION = (WM_USER + 108)      '仅用于RTB v1.0 的亚洲语言的操作系统。RTB 2.0不支持！

Public Const EM_SETLANGOPTIONS = (WM_USER + 120)    '设置IME和远东语言支持选项。
Public Const EM_GETLANGOPTIONS = (WM_USER + 121)    '获取IME和远东语言支持选项。
Public Const EM_GETIMECOMPMODE = (WM_USER + 122)    '获取当前IME模式。


' /* BiDi 双向语言支持 特殊消息 */
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)    '设置当前双向语言支持选项。
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)    '获取当前双向语言支持选项。

' /* Options for EM_SETLANGOPTIONS and EM_GETLANGOPTIONS */
Public Const IMF_AUTOKEYBOARD = &H1             '自动键盘布局
Public Const IMF_AUTOFONT = &H2                 '自动字体
Public Const IMF_IMECANCELCOMPLETE = &H4      '// high completes the comp string when aborting, low cancels.
Public Const IMF_IMEALWAYSSENDNOTIFY = &H8

' /* EM_GETIMECOMPMODE 的取值 */
Public Const ICM_NOTOPEN = &H0          'Input Method Editor (IME) is not open.
Public Const ICM_LEVEL3 = &H1           'True inline mode.
Public Const ICM_LEVEL2 = &H2           'Level 2.
Public Const ICM_LEVEL2_5 = &H3         'Level 2.5
Public Const ICM_LEVEL2_SUI = &H4       'Special user interface (UI).

' /* 新的通知消息 */

Public Const EN_MSGFILTER = &H700&      'RTB控件通过 WM_NOTIFY 消息通知父窗体有鼠标或者键盘事件产生。
Public Const EN_REQUESTRESIZE = &H701&  'RTB控件通过 WM_NOTIFY 消息通知父窗体尺寸有改变。
Public Const EN_SELCHANGE = &H702&      'RTB控件通过 WM_NOTIFY 消息通知父窗体当前选择区域发生变化。
Public Const EN_DROPFILES = &H703&      'RTB控件在接受到 WM_DROPFILES 消息后通过 WM_NOTIFY 消息通知父窗体用户试图放下一个文件。
Public Const EN_PROTECTED = &H704&      'RTB控件通过 WM_NOTIFY 消息通知父窗体用户试图改变受保护文本。
Public Const EN_CORRECTTEXT = &H705&    '一个EN_CORRECTTEXT 手势。   /* PenWin specific */
Public Const EN_STOPNOUNDO = &H706&     'RTB控件通过 WM_NOTIFY 消息通知父窗体某个操作无法分配足够内存来记录其状态。
Public Const EN_IMECHANGE = &H707&      'IME 改变。                  /* Far East specific */
Public Const EN_SAVECLIPBOARD = &H708&  '通知父窗体，RTB在关闭时剪贴板中还有数据。
Public Const EN_OLEOPFAILED = &H709&    '通知父窗体，一个对OLE对象的操作失败。
Public Const EN_OBJECTPOSITIONS = &H70A&    '通知父窗体，RTB读入一个OLE对象。
Public Const EN_LINK = &H70B&               'RTB控件通过 WM_NOTIFY 消息通知父窗体用户在超链接效果文本上的多种鼠标事件。
Public Const EN_DRAGDROPDONE = &H70C&       'RTB控件通过 WM_NOTIFY 消息通知父窗体一个拖放操作完成。

' /* BiDi 双向语言支持 特殊通知消息 */

Public Const EN_ALIGN_LTR = &H710&      'RTB控件通过 WM_COMMAND 消息通知父窗体段落方向改为从左至右。
Public Const EN_ALIGN_RTL = &H711&      'RTB控件通过 WM_COMMAND 消息通知父窗体段落方向改为从右至左。

' /* 事件通知掩码 */

Public Const ENM_NONE = &H0             '默认值。表示不会向父窗体发送任何消息。
Public Const ENM_CHANGE = &H1           '可以发送 EN_CHANGE 消息。
Public Const ENM_UPDATE = &H2           '可以发送 EN_UPDATE 消息。
Public Const ENM_SCROLL = &H4           '可以发送 EN_HSCROLL 消息。
Public Const ENM_KEYEVENTS = &H10000    '可以发送 EN_MSGFILTER 消息。
Public Const ENM_MOUSEEVENTS = &H20000  '可以发送 EN_MSGFILTER 消息。
Public Const ENM_REQUESTRESIZE = &H40000    '可以发送 EN_REQUESTRESIZE 消息。
Public Const ENM_SELCHANGE = &H80000        '可以发送 EN_SELCHANGE 消息。
Public Const ENM_DROPFILES = &H100000       '可以发送 EN_DROPFILES 消息。
Public Const ENM_PROTECTED = &H200000       '可以发送 EN_PROTECTED 消息。
Public Const ENM_CORRECTTEXT = &H400000     ' /* PenWin specific */
Public Const ENM_SCROLLEVENTS = &H8         '可以发送 EN_MSGFILTER 中的鼠标滚轮事件消息。
Public Const ENM_DRAGDROPDONE = &H10        '可以发送 EN_DRAGDROPDONE 消息。

' /* 远东特定通知掩码 */
Public Const ENM_IMECHANGE = &H800000           ' /* RE2.0 不支持！，只用于1.0版本！*/
Public Const ENM_LANGCHANGE = &H1000000         ' ？？
Public Const ENM_OBJECTPOSITIONS = &H2000000    '可以发送 EN_OBJECTPOSITIONS 消息。
Public Const ENM_LINK = &H4000000               '可以发送 EN_LINK 消息。

' /* 新的 Edit 控件样式 */

Public Const ES_SAVESEL = &H8000&               '在失去焦点时保持选择区域高亮显示！！！Useful！
Public Const ES_SUNKEN = &H4000&                '凹下效果
Public Const ES_DISABLENOSCROLL = &H2000&       '在不需要滚动条时将其置灰，而非隐藏
' /* same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value */
Public Const ES_SELECTIONBAR = &H1000000
' /* same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop */
Public Const ES_NOOLEDRAGDROP = &H8

' /* 新的 Edit 控件扩展样式 */
' #ifdef  _WIN32
Public Const ES_EX_NOCALLOLEINIT = &H1000000
' #End If

' /* These flags are used in FE Windows */
Public Const ES_VERTICAL = &H400000     '垂直绘制文本和对象。
Public Const ES_NOIME = &H80000         '禁用IME。
Public Const ES_SELFIME = &H40000       '应用程序来控制IME操作。

' /* 新的断字处理动作 */
Public Const WB_CLASSIFY = 3&           '
Public Const WB_MOVEWORDLEFT = 4&       '
Public Const WB_MOVEWORDRIGHT = 5&      '
Public Const WB_LEFTBREAK = 6&          '
Public Const WB_RIGHTBREAK = 7&         '

' /* 远东特殊标志位 */
Public Const WB_MOVEWORDPREV = 4&
Public Const WB_MOVEWORDNEXT = 5&
Public Const WB_PREVBREAK = 6&
Public Const WB_NEXTBREAK = 7&

Public Const PC_FOLLOWING = 1&
Public Const PC_LEADING = 2&
Public Const PC_OVERFLOW = 3&
Public Const PC_DELIMITER = 4&
Public Const WBF_WORDWRAP = &H10&
Public Const WBF_WORDBREAK = &H20&
Public Const WBF_OVERFLOW = &H40&
Public Const WBF_LEVEL1 = &H80&
Public Const WBF_LEVEL2 = &H100&
Public Const WBF_CUSTOM = &H200&

' /* 远东特殊标志位 */
Public Const IMF_FORCENONE = &H1
Public Const IMF_FORCEENABLE = &H2
Public Const IMF_FORCEDISABLE = &H4
Public Const IMF_CLOSESTATUSWINDOW = &H8
Public Const IMF_VERTICAL = &H20
Public Const IMF_FORCEACTIVE = &H40
Public Const IMF_FORCEINACTIVE = &H80
Public Const IMF_FORCEREMEMBER = &H100
Public Const IMF_MULTIPLEEDIT = &H400

' /* 断字标志位（用于WB_CLASSIFY） */
Public Const WBF_CLASS = &HF          '((BYTE) =&H0F)
Public Const WBF_ISWHITE = &H10       '((BYTE) =&H10)
Public Const WBF_BREAKLINE = &H20     '((BYTE) =&H20)
Public Const WBF_BREAKAFTER = &H40    '((BYTE) =&H40)



' /* CHARFORMAT 掩码 */
Public Const CFM_BOLD = &H1             '粗体有效。
Public Const CFM_ITALIC = &H2           '斜体有效。
Public Const CFM_UNDERLINE = &H4        '下划线有效。
Public Const CFM_STRIKEOUT = &H8        '删除线有效。
Public Const CFM_PROTECTED = &H10       '保护有效。
Public Const CFM_LINK = &H20&           '超链接有效。  ' /* Exchange hyperlink extension */
Public Const CFM_SIZE = &H80000000      '字符高度有效，单位：缇。
Public Const CFM_COLOR = &H40000000     '文本颜色有效。
Public Const CFM_FACE = &H20000000      '字体名称有效。
Public Const CFM_OFFSET = &H10000000    '字符偏移有效。指基线上或下的偏移量（上标/下标）。
Public Const CFM_CHARSET = &H8000000    '字符集有效。

' /* CHARFORMAT 效果 */
Public Const CFE_BOLD = &H1&            '粗体
Public Const CFE_ITALIC = &H2&          '斜体
Public Const CFE_UNDERLINE = &H4&       '下划线
Public Const CFE_STRIKEOUT = &H8&       '删除线
Public Const CFE_PROTECTED = &H10&      '保护
Public Const CFE_LINK = &H20&           '超链接
Public Const CFE_AUTOCOLOR = &H40000000 '采用系统自动颜色。' /* NOTE: this corresponds to */
                                        ' /* CFM_COLOR, which controls it */
Public Const yHeightCharPtsMost = 1638& '最大字体尺寸值，仅指Y坐标尺寸，单位：磅（点）。

' /* EM_SETCHARFORMAT wParam 参数掩码 */
Public Const SCF_SELECTION = &H1&   '应用于当前选中区域。
Public Const SCF_WORD = &H2&        '应用于当前选中单词。
Public Const SCF_DEFAULT = &H0&            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4&                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8&         '// modifier for SCF_SELECTION; says that
                                   ' // the format came from a toolbar, etc. and
                                   ' // therefore UI formatting rules should be
                                   ' // used instead of strictly formatting the
                                   ' // selection.

' /* 流的格式 */
Public Const SF_TEXT = &H1         'Text格式
Public Const SF_RTF = &H2          'RTF格式
Public Const SF_RTFNOOBJS = &H3    '输出时用空格代替对象，仅用于输出！
Public Const SF_TEXTIZED = &H4     '输出时采用文本表示对象，仅用于输出！
Public Const SF_UNICODE = &H10            ' /* Unicode file of some kind */

' /* Flag telling stream operations to operate on the selection only */
' /* EM_STREAMIN will replace the current selection */
' /* EM_STREAMOUT will stream out the current selection */
Public Const SFF_SELECTION = &H8000&    '输入输出只对当前选择区域有效！

' /* Flag telling stream operations to operate on the common RTF keyword only */
' /* EM_STREAMIN will accept the only common RTF keyword */
' /* EM_STREAMOUT will stream out the only common RTF keyword */
Public Const SFF_PLAINRTF = &H4000&     '只使用通用RTF关键字，对于与语言相关的RTF关键字予以忽略！
' /* 所有段落度量单位均为：缇 */

Public Const MAX_TAB_STOPS = 32&    '绝对制表符的最大数目。
Public Const lDefaultTab = 720&     '默认绝对制表符位置。
' /* PARAFORMAT 掩码值 */
Public Const PFM_STARTINDENT = &H1& '首行缩进值有效。
Public Const PFM_RIGHTINDENT = &H2& '右缩进值有效。
Public Const PFM_OFFSET = &H4&      '缩进或者悬挂有效！负值表示缩进，正值表示悬挂！
Public Const PFM_ALIGNMENT = &H8&   '水平对齐方式有效。
Public Const PFM_TABSTOPS = &H10&   '绝对制表符位置有效。
Public Const PFM_NUMBERING = &H20&  '编号与项目符号有效。
Public Const PFM_OFFSETINDENT = &H80000000  '首行缩进值有效，并且给出一个相对值。

' /* PARAFORMAT 编号选项 */
Public Const PFN_BULLET = &H1&      '

' /* PARAFORMAT 对齐选项 */
Public Const PFA_LEFT = &H1&        '
Public Const PFA_RIGHT = &H2&       '
Public Const PFA_CENTER = &H3&      '
'映射为所有掩码有效。
Public Const CFM_EFFECTS = (CFM_BOLD Or CFM_ITALIC Or CFM_UNDERLINE Or CFM_COLOR Or _
                     CFM_STRIKEOUT Or CFE_PROTECTED Or CFM_LINK)
Public Const CFM_ALL = (CFM_EFFECTS Or CFM_SIZE Or CFM_FACE Or CFM_OFFSET Or CFM_CHARSET)

' /* 新的掩码和效果 － (*)表示数据在RichEdit 2.0中保存，但是不会显示！

Public Const CFM_SMALLCAPS = &H40&                 ' /* (*)  */
Public Const CFM_ALLCAPS = &H80&                   ' /* (*)  */
Public Const CFM_HIDDEN = &H100&                   ' /* (*)  */
Public Const CFM_OUTLINE = &H200&                  ' /* (*)  */
Public Const CFM_SHADOW = &H400&                   ' /* (*)  */
Public Const CFM_EMBOSS = &H800&                   ' /* (*)  */
Public Const CFM_IMPRINT = &H1000&                 ' /* (*)  */
Public Const CFM_DISABLED = &H2000&
Public Const CFM_REVISED = &H4000&

Public Const CFM_BACKCOLOR = &H4000000
Public Const CFM_LCID = &H2000000
Public Const CFM_UNDERLINETYPE = &H800000         ' /* (*)  */
Public Const CFM_WEIGHT = &H400000
Public Const CFM_SPACING = &H200000               ' /* (*)  */
Public Const CFM_KERNING = &H100000               ' /* (*)  */
Public Const CFM_STYLE = &H80000                  ' /* (*)  */
Public Const CFM_ANIMATION = &H40000              ' /* (*)  */
Public Const CFM_REVAUTHOR = &H8000&

Public Const CFE_SUBSCRIPT = &H10000                ' /*  上标和下标是互斥的！      */
Public Const CFE_SUPERSCRIPT = &H20000              ' /*  上标和下标是互斥的！      */

Public Const CFM_SUBSCRIPT = CFE_SUBSCRIPT Or CFE_SUPERSCRIPT
Public Const CFM_SUPERSCRIPT = CFM_SUBSCRIPT

'映射为所有掩码有效。
Public Const CFM_EFFECTS2 = (CFM_EFFECTS Or CFM_DISABLED Or CFM_SMALLCAPS Or CFM_ALLCAPS _
                    Or CFM_HIDDEN Or CFM_OUTLINE Or CFM_SHADOW Or CFM_EMBOSS _
                    Or CFM_IMPRINT Or CFM_DISABLED Or CFM_REVISED _
                    Or CFM_SUBSCRIPT Or CFM_SUPERSCRIPT Or CFM_BACKCOLOR)

Public Const CFM_ALL2 = (CFM_ALL Or CFM_EFFECTS2 Or CFM_BACKCOLOR Or CFM_LCID _
                    Or CFM_UNDERLINETYPE Or CFM_WEIGHT Or CFM_REVAUTHOR _
                    Or CFM_SPACING Or CFM_KERNING Or CFM_STYLE Or CFM_ANIMATION)

Public Const CFE_SMALLCAPS = CFM_SMALLCAPS
Public Const CFE_ALLCAPS = CFM_ALLCAPS
Public Const CFE_HIDDEN = CFM_HIDDEN
Public Const CFE_OUTLINE = CFM_OUTLINE
Public Const CFE_SHADOW = CFM_SHADOW
Public Const CFE_EMBOSS = CFM_EMBOSS
Public Const CFE_IMPRINT = CFM_IMPRINT
Public Const CFE_DISABLED = CFM_DISABLED
Public Const CFE_REVISED = CFM_REVISED

' /* NOTE: CFE_AUTOCOLOR and CFE_AUTOBACKCOLOR correspond to CFM_COLOR and
'   CFM_BACKCOLOR, respectively, which control them */
Public Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR

' /* Underline types */
Public Const CFU_CF1UNDERLINE = &HFF&      ' /* map charformat's bit underline to CF2.*/
Public Const CFU_INVERT = &HFE&            ' /* For IME composition fake a selection.*/
Public Const CFU_UNDERLINEDOTTED = &H4&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEDOUBLE = &H3&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEWORD = &H2&      ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINE = &H1&
Public Const CFU_UNDERLINENONE = 0&
' /* PARAFORMAT 2.0 掩码和效果 */

Public Const PFM_SPACEBEFORE = &H40&
Public Const PFM_SPACEAFTER = &H80&
Public Const PFM_LINESPACING = &H100&
Public Const PFM_STYLE = &H400&
Public Const PFM_BORDER = &H800&                   ' /* (*)  */
Public Const PFM_SHADING = &H1000&                 ' /* (*)  */
Public Const PFM_NUMBERINGSTYLE = &H2000&          ' /* (*)  */
Public Const PFM_NUMBERINGTAB = &H4000&            ' /* (*)  */
Public Const PFM_NUMBERINGSTART = &H8000&         ' /* (*)  */

Public Const PFM_DIR = &H10000
Public Const PFM_RTLPARA = &H10000                ' /* (Version 1.0 flag) */
Public Const PFM_KEEP = &H20000                   ' /* (*)  */
Public Const PFM_KEEPNEXT = &H40000               ' /* (*)  */
Public Const PFM_PAGEBREAKBEFORE = &H80000        ' /* (*)  */
Public Const PFM_NOLINENUMBER = &H100000          ' /* (*)  */
Public Const PFM_NOWIDOWCONTROL = &H200000        ' /* (*)  */
Public Const PFM_DONOTHYPHEN = &H400000           ' /* (*)  */
Public Const PFM_SIDEBYSIDE = &H800000            ' /* (*)  */

Public Const PFM_TABLE = &HC0000000               ' /* (*)  */

' /* Note: PARAFORMAT has no effects */
Public Const PFM_EFFECTS = (PFM_DIR Or PFM_KEEP Or PFM_KEEPNEXT Or PFM_TABLE _
                    Or PFM_PAGEBREAKBEFORE Or PFM_NOLINENUMBER _
                    Or PFM_NOWIDOWCONTROL Or PFM_DONOTHYPHEN Or PFM_SIDEBYSIDE _
                    Or PFM_TABLE)

Public Const PFM_ALL = (PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET Or _
                 PFM_ALIGNMENT Or PFM_TABSTOPS Or PFM_NUMBERING Or _
                 PFM_OFFSETINDENT Or PFM_DIR)

Public Const PFM_ALL2 = (PFM_ALL Or PFM_EFFECTS Or PFM_SPACEBEFORE Or PFM_SPACEAFTER _
                    Or PFM_LINESPACING Or PFM_STYLE Or PFM_SHADING Or PFM_BORDER _
                    Or PFM_NUMBERINGTAB Or PFM_NUMBERINGSTART Or PFM_NUMBERINGSTYLE)

'public const PFE_RTLPARA  =           (PFM_DIR             >> 16)
'public const PFE_RTLPAR              (PFM_RTLPARA         >> 16) ' /* (Version 1.0 flag) */
'public const PFE_KEEP                (PFM_KEEP            >> 16) ' /* (*)  */
'public const PFE_KEEPNEXT            (PFM_KEEPNEXT        >> 16) ' /* (*)  */
'public const PFE_PAGEBREAKBEFORE     (PFM_PAGEBREAKBEFORE >> 16) ' /* (*)  */
'public const PFE_NOLINENUMBER        (PFM_NOLINENUMBER    >> 16) ' /* (*)  */
'public const PFE_NOWIDOWCONTROL      (PFM_NOWIDOWCONTROL  >> 16) ' /* (*)  */
'public const PFE_DONOTHYPHEN         (PFM_DONOTHYPHEN     >> 16) ' /* (*)  */
'public const PFE_SIDEBYSIDE          (PFM_SIDEBYSIDE      >> 16) ' /* (*)  */'

Public Const PFE_TABLEROW = &HC000&                ' /* These 3 options are mutually */
Public Const PFE_TABLECELLEND = &H8000&            ' /*  exclusive and each imply    */
Public Const PFE_TABLECELL = &H4000&               ' /*  段落为表格的一部分 */

Public Const PFA_JUSTIFY = 4          ' /* 两端对齐，为了兼容TOM模型接口。 (*)  */

' 用于在 IRichEditOleCallback::GetContextMenu 函数中请求应用程序提供一个右键菜单。
Public Const GCM_RIGHTMOUSEDROP = &H8000&
Public Const OLEOP_DOVERB = 1

' 剪贴板格式，用于 RegisterClipboardFormat() 注册有效的剪贴板格式。
Public Const CF_RTF = "Rich Text Format"
Public Const CF_RTFNOOBJS = "Rich Text Format Without Objects"
Public Const CF_RETEXTOBJ = "RichEdit Text and Objects"

' /* 用于下面的 GETTEXTEX 数据结构 */
Public Const GT_DEFAULT = 0&    '不使用CR转换
Public Const GT_USECRLF = 1&    '表示在每次拷贝文本时，将CR转换为CRLF。
' GETTEXTLENGTHEX 数据结构的标志位
Public Const GTL_DEFAULT = 0&      ' /* 默认值，返回字符数目。                      */
Public Const GTL_USECRLF = 1&      ' /* 使用段落 CR/LF 计算                         */
Public Const GTL_PRECISE = 2&      ' /* 精确计算，较慢                              */
Public Const GTL_CLOSE = 4&        ' /* 近似计算，较快，常用于提前分配内存空间      */
Public Const GTL_NUMCHARS = 8&     ' /* 返回字符数目                                */
Public Const GTL_NUMBYTES = 16&    ' /* 返回字节数目                                */
' /* BIDIOPTIONS masks */
' #if (_RICHEDIT_VER == =&H0100)
Public Const BOM_DEFPARADIR = &H1&             ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOM_PLAINTEXT = &H2&              ' /* Use plain text layout (obsolete) */
Public Const BOM_NEUTRALOVERRIDE = &H4&        ' /* Override neutral layout (obsolete) */
' #endif ' /* _RICHEDIT_VER == =&H0100 */
Public Const BOM_CONTEXTREADING = &H8&         ' /* Context reading order */
Public Const BOM_CONTEXTALIGNMENT = &H10&      ' /* Context alignment */
' /* BIDIOPTIONS effects */
' #if (_RICHEDIT_VER == =&H0100)
Public Const BOE_RTLDIR = &H1&                 ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOE_PLAINTEXT = &H2&              ' /* Use plain text layout (obsolete) */
Public Const BOE_NEUTRALOVERRIDE = &H4&        ' /* Override neutral layout (obsolete) */
' #endif ' /* _RICHEDIT_VER == =&H0100 */
Public Const BOE_CONTEXTREADING = &H8&         ' /* Context reading order */
Public Const BOE_CONTEXTALIGNMENT = &H10&      ' /* Context alignment */
' /* 新增的 EM_FINDTEXT[EX] 标志 */
Public Const FR_MATCHDIAC = &H20000000          ' 阿拉伯与希伯来语用
Public Const FR_MATCHKASHIDA = &H40000000       ' 阿拉伯与希伯来语用
Public Const FR_MATCHALEFHAMZA = &H80000000     ' 阿拉伯与希伯来语用
' /* UNICODE 嵌入字符 */
Public Const WCH_EMBEDDING = &HFFFC&

Public Const SB_BOTH = 3
Public Const SB_BOTTOM = 7
Public Const SB_CTL = 2
Public Const SB_ENDSCROLL = 8
Public Const SB_HORZ = 0
Public Const SB_LEFT = 6
Public Const SB_LINEDOWN = 1
Public Const SB_LINELEFT = 0
Public Const SB_LINERIGHT = 1
Public Const SB_LINEUP = 0
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGELEFT = 2
Public Const SB_PAGERIGHT = 3
Public Const SB_PAGEUP = 2
Public Const SB_RIGHT = 7
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_VERT = 1

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_ENABLE_BOTH = &H0

Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&
Public Const SBS_SIZEGRIP = &H10&

' Flat scroll bars:
Public Const WSB_PROP_CYVSCROLL = &H1&
Public Const WSB_PROP_CXHSCROLL = &H2&
Public Const WSB_PROP_CYHSCROLL = &H4&
Public Const WSB_PROP_CXVSCROLL = &H8&
Public Const WSB_PROP_CXHTHUMB = &H10&
Public Const WSB_PROP_CYVTHUMB = &H20&
Public Const WSB_PROP_VBKGCOLOR = &H40&
Public Const WSB_PROP_HBKGCOLOR = &H80&
Public Const WSB_PROP_VSTYLE = &H100&
Public Const WSB_PROP_HSTYLE = &H200&
Public Const WSB_PROP_WINSTYLE = &H400&
Public Const WSB_PROP_PALETTE = &H800&
Public Const WSB_PROP_MASK = &HFFF&

Public Const FSB_FLAT_MODE = 2&
Public Const FSB_ENCARTA_MODE = 1&
Public Const FSB_REGULAR_MODE = 0&

Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
' Show window styles
Public Const SW_ERASE = &H4

Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_RESTORE = 9
Public Const SW_PARENTOPENING = 3
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
' Button messages:
Public Const BM_GETCHECK = &HF0&
Public Const BM_SETCHECK = &HF1&
Public Const BM_GETSTATE = &HF2&
Public Const BM_SETSTATE = &HF3&
Public Const BM_SETSTYLE = &HF4&
Public Const BM_CLICK = &HF5&
Public Const BM_GETIMAGE = &HF6&
Public Const BM_SETIMAGE = &HF7&

Public Const BST_UNCHECKED = &H0&
Public Const BST_CHECKED = &H1&
Public Const BST_INDETERMINATE = &H2&
Public Const BST_PUSHED = &H4&
Public Const BST_FOCUS = &H8&

' Button notifications:
Public Const BN_CLICKED = 0&
Public Const BN_PAINT = 1&
Public Const BN_HILITE = 2&
Public Const BN_UNHILITE = 3&
Public Const BN_DISABLE = 4&
Public Const BN_DOUBLECLICKED = 5&
Public Const BN_PUSHED = BN_HILITE
Public Const BN_UNPUSHED = BN_UNHILITE
Public Const BN_DBLCLK = BN_DOUBLECLICKED
Public Const BN_SETFOCUS = 6&
Public Const BN_KILLFOCUS = 7&

' Button Styles:
Public Const BS_3STATE = &H5&
Public Const BS_AUTO3STATE = &H6&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BS_AUTORADIOBUTTON = &H9&
Public Const BS_CHECKBOX = &H2&
Public Const BS_DEFPUSHBUTTON = &H1&
Public Const BS_GROUPBOX = &H7&
Public Const BS_LEFTTEXT = &H20&
Public Const BS_OWNERDRAW = &HB&
Public Const BS_PUSHBUTTON = &H0&
Public Const BS_RADIOBUTTON = &H4&
Public Const BS_USERBUTTON = &H8&
Public Const BS_ICON = &H40&
Public Const BS_BITMAP = &H80&
Public Const BS_LEFT = &H100&
Public Const BS_RIGHT = &H200&
Public Const BS_CENTER = &H300&
Public Const BS_TOP = &H400&
Public Const BS_BOTTOM = &H800&
Public Const BS_VCENTER = &HC00&
Public Const BS_PUSHLIKE = &H1000&
Public Const BS_MULTILINE = &H2000&
Public Const BS_NOTIFY = &H4000&
Public Const BS_FLAT = &H8000&
Public Const BS_RIGHTBUTTON = BS_LEFTTEXT

' Built in ImageList drawing methods:
Public Const ILD_NORMAL = 0
Public Const ILD_Transparent = 1
Public Const ILD_BLEND25 = 2
Public Const ILD_SELECTED = 4
Public Const ILD_FOCUS = 4
Public Const ILD_OVERLAYMASK = 3840
Public Const ILD_MASK = &H10&
Public Const ILD_IMAGE = &H20&
Public Const ILD_ROP = &H40&

' Use default rgb colour:
Public Const CLR_NONE = -1
Public Const CLR_INVALID = -1
Public Const CLR_DEFAULT = -16777216
Public Const CLR_HILIGHT = -16777216

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Const WM_MEASUREITEM = &H2C
Public Const WM_DRAWITEM = &H2B
Public Const WM_SIZE = &H5
Public Const WM_CTLCOLORSCROLLBAR = &H137

' Missing Draw State constants declarations:
'/* Image type */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* State type */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10 ' Dither
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80 ' Draw in colour of brush specified in hBrush
Public Const DSS_RIGHT = &H8000

Public Const BF_LEFT = 1
Public Const BF_TOP = 2
Public Const BF_RIGHT = 4
Public Const BF_BOTTOM = 8
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const BF_MIDDLE = 2048

Public Const BDR_RAISEDOUTER = 1
Public Const BDR_SUNKENOUTER = 2
Public Const BDR_RAISEDINNER = 4
Public Const BDR_SUNKENINNER = 8
Public Const BDR_BUTTONPRESSED = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Public Const BDR_BUTTONNORMAL = BDR_RAISEDINNER Or BDR_RAISEDOUTER
'#########################################################################

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

Public Const CB_SHOWDROPDOWN = &H14F            'Cbo弹出选项
Public Const CB_SETDROPPEDWIDTH As Long = &H160 'Cbo下拉宽度

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

'表示一个Windows位图格式。
Public Const CF_BITMAP = 2
'3D效果颜色
Public Const LR_LOADMAP3DCOLORS = &H1000
'图片从文件lpsz中调入，而非从资源文件中调入。
Public Const LR_LOADFROMFILE = &H10
'调入透明色
Public Const LR_LOADTransparent = &H20
Public Const LR_COPYRETURNORG = &H4
'用户点击系统菜单中的“移动”菜单事件
Public Const SC_MOVE = &HF012

'系统默认颜色
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLORONCOLOR = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6  '窗体框架
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15         '按钮表明
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18         '按钮普通文本
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_ADJ_MAX = 100

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
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000  '绝对移动
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4       '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1         '鼠标移动
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10     '  right button up

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
 
Public Const GWW_HINSTANCE = (-6)
Public Const LOGPIXELSX = 88        '沿屏幕宽度的每个逻辑英寸的象素值。在多显示器系统中，所有显示器的这个值均相同。
Public Const LOGPIXELSY = 90        '沿屏幕高度的每个逻辑英寸的象素值。在多显示器系统中，所有显示器的这个值均相同。
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4
Public Const SM_CYFRAME = 33
Public Const SM_CYBORDER = 6
Public Const SM_CXBORDER = 5

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Const OPAQUE = 2
Public Const Transparent = 1

Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062

Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12
Public Const SYSTEM_FONT = 13
Public Const DEFAULT_GUI_FONT = 9 'win95 only

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1



Public Const ILC_MASK = &H1&
Public Const ILCF_MOVE = &H0&
Public Const ILCF_SWAP = &H1&
'#########################################################################

Public Const HSTEP = 50         '滚动条水平步长
Public Const VSTEP = 50         '滚动条垂直步长
Public Const PAGEMARGIN = 200   '页面视图下控件与容器的边距
Public Const SHADOWOFFSET = 30  '阴影偏移量
Public Const WHEELNUMBER = 20   '鼠标滚动系数
Public Const MM_ANISOTROPIC = 8 ' Map mode anisotropic
Public Const giINVALID_PICTURE As Integer = 481        'VB Errors Error code used by Transparent Picture copy routines
Public Const DSna = &H220326 'Raster Operation Codes
Public Const VK_TAB = &H9 ' Virtual key values

Public Const EMO_EXIT = 0                     ' // enter normal mode,  lparam ignored
Public Const EMO_ENTER = 1                    ' // enter outline mode, lparam ignored
Public Const EMO_PROMOTE = 2                  ' // LOWORD(lparam) == 0 ==>
                                        ' // promote  to body-text
                                        ' // LOWORD(lparam) != 0 ==>
                                        ' // promote/demote current selection
                                        ' // by indicated number of levels
Public Const EMO_EXPAND = 3                   ' // HIWORD(lparam) = EMO_EXPANDSELECTION
                                        ' // -> expands selection to level
                                        ' // indicated in LOWORD(lparam)
                                        ' // LOWORD(lparam) = -1/+1 corresponds
                                        ' // to collapse/expand button presses
                                        ' // in winword (other values are
                                        ' // equivalent to having pressed these
                                        ' // buttons more than once)
                                        ' // HIWORD(lparam) = EMO_EXPANDDOCUMENT
                                        ' // -> expands whole document to
                                        ' // indicated level
Public Const EMO_MOVESELECTION = 4            ' // LOWORD(lparam) != 0 -> move current
                                        ' // selection up/down by indicated
                                        ' // amount
Public Const EMO_GETVIEWMODE = 5          ' // Returns VM_NORMAL or VM_OUTLINE
'   是否展开
Public Const EMO_EXPANDSELECTION = 0
Public Const EMO_EXPANDDOCUMENT = 1
Public Const VM_NORMAL = 4             ' // Agrees with RTF \viewkindN
Public Const VM_OUTLINE = 2
'######################################################################################
'   获取字符屏幕位置
'######################################################################################
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24

Public Const S_FALSE = &H1
Public Const S_OK = &H0

'######################################################################################
'   直接发送按键的函数
'######################################################################################
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

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


' VB API Viewer 版本的 DocInfo 结构说明是错误的！！！！
' VB API VIEWER VERSION OF DOCINFO STRUCTURE IS WRONG!
'用于存储 StartDoc() 中文件名及其他信息
Public Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

'用于初始化打印对话框及返回值
Public Type PrintDlg
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
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

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

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Type SIZEAPI
    cx As Long
    cy As Long
End Type

Public Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type
Public Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Public Type GUID '  GUID,  IID,  CLSID,  etc
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' /* 所有的字符格式度量单位均为：缇 */
' 已经纠正！！！...
Public Type CHARFORMAT
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58           '？？？？WCHAR
    wPad2 As Integer ' 60
End Type

'字符范围：
Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

'文本范围：通过 EM_GETTEXTRANGE 消息填充！
Public Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String    ' /* allocated by caller, zero terminated by RichEdit */
End Type
'用于存储 EM_STREAMIN 或者 EM_STREAMOUT 消息传递的数据信息。
Public Type EDITSTREAM
    dwCookie As Long     ' /* user value passed to callback as first parameter */
    dwError As Long      ' /* last error */
    pfnCallback As Long  'EDITSTREAMCALLBACK
End Type

'段落格式
Public Type PARAFORMAT
    cbSize As Integer       '
    wPad1 As Integer        '
    dwMask As Long          '
    wNumbering As Integer   '
    wEffects As Integer     ' Note reserved in RichEdit 32
    dxStartIndent As Long   '
    dxRightIndent As Long   '
    dxOffset As Long        '
    wAlignment As Integer   '
    cTabCount As Integer    '
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long   '
End Type

'用于 EM_FINDTEXT 消息的查找文本的相关信息
Public Type FindText
    chrg As CHARRANGE   '字符范围
    lpstrText As Long   '需要查找的文本
End Type

'扩展的文本查找消息结构体
Public Type FINDTEXTEX_A
    chrg As CHARRANGE       '字符范围
    lpstrText As Long       '需要查找的文本
    chrgText As CHARRANGE   '查找到的文本范围
End Type

'同上
Public Type FINDTEXTEX_W
    chrg As CHARRANGE
    lpstrText As Long
    chrgText As CHARRANGE
End Type

'包含用于格式化指定设备的相关信息
Public Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type



Public Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    'RICHEDIT20 支持的新成员
    wWeight As Integer              ' /* 字体磅值（参见LOGFONT值）      */
    sSpacing As Integer             ' /* 水平字符间隔，用于兼容TOM接口  */
    crBackColor As Long             ' /* 背景色                         */
    lLCID As Long                   ' /* 32位的本地 ID                  */
    dwReserved As Long              ' /* 保留，必须为0                  */
    sStyle As Integer               ' /* 样式指针，用于兼容TOM接口      */
    wKerning As Integer             ' /* 字符压缩最小宽度，用于兼容TOM接口 */
    bUnderlineType As Byte          ' /* 下划线类型                     */
    bAnimation As Byte              ' /* 动态文本效果，用于兼容TOM接口  */
    bRevAuthor As Byte              ' /* 修订作者索引，用不同颜色显示不同作者的修订信息 */
    bReserved1 As Byte              ' /* 保留，必须为0                  */
End Type


Public Type PARAFORMAT2
    cbSize As Integer               '指定该结构的字节大小。
    wPad1 As Integer                '
    dwMask As Long                  '掩码组合
    wNumbering As Integer           '项目符号与编号
    wReserved As Integer            '
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    'rgxTabs(0 To MAX_TAB_STOPS - 1) As Byte
    'lPtrRgxTabs As Long
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long
    dySpaceBefore As Long          ' /* Vertical spacing before para         */
    dySpaceAfter As Long           ' /* Vertical spacing after para          */
    dyLineSpacing As Long          ' /* Line spacing depending on Rule       */
    sStyle As Integer                  ' /* Style handle                         */
    bLineSpacingRule As Byte       ' /* Rule for line spacing (see tom.doc)  */
    bCRC As Byte                   ' /* Reserved for CRC for rapid searching *
    wShadingWeight As Integer          ' /* Shading in hundredths of a per cent  */
    wShadingStyle As Integer           ' /* Nibble 0: style, 1: cfpat, 2: cbpat  */
    wNumberingStart As Integer         ' /* Starting value for numbering         */
    wNumberingStyle As Integer        ' /* Alignment, roman/arabic, (), ), ., etc.*/
    wNumberingTab As Integer           ' /* Space bet 1st indent and 1st-line text*/
    wBorderSpace As Integer            ' /* Space between border and text (twips)*/
    wBorderWidth As Integer           ' /* Border pen width (twips)             */
    wBorders As Integer                ' /* Byte 0: bits specify which borders   */
                                    ' /* Nibble 2: border style, 3: color index*/
End Type

' #endif ' /* C++   */



' /* 通知的结构 */
Public Type NMHDR
    hwndFrom As Long        '消息发送的目标窗体
    wPad1 As Integer        '-
    idfrom As Integer       '发送消息的控件ID
    code As Integer         '消息代码
    wPad2 As Integer        '-
End Type
' #endif  ' /* !WM_NOTIFY */

'用于 EN_MSGFILTER 消息，存储鼠标、键盘事件。
Public Type MSGFILTER
    NMHDR As NMHDR '通知头
    Msg As Integer          '键盘或者鼠标标识符
    wPad1 As Integer        '-
    wParam As Integer       '消息的wParam值，指的是RTB的ID
    wPad2 As Integer        '-
    lParam As Long          '消息的lParam值，指的是该消息的 MSGFILTER 结构体的指针。
End Type

Public Type REQRESIZE
    NMHDR As NMHDR     '通知头
    rc As RECT                  '请求的新尺寸！
End Type

Public Type SelChange
    NMHDR As NMHDR     '通知头
    chrg As CHARRANGE           '新的选择范围
    seltyp As Long              '新的选择范围的内容（文本、对象、多个对象等）
End Type

'包含拽下的文件信息
Public Type ENDROPFILES
    NMHDR As NMHDR     '通知头
    hDrop As Long               '放下的文件列表句柄（同 WM_DROPFILES）
    cP As Long                  '将被插入的字符位置
    fProtected As Long          '指定该字符位置是否受保护
End Type

'用户试图修改受保护文档是的信息内容
Public Type ENPROTECTED
    NMHDR As NMHDR     '通知头
    Msg As Long                 '触发该通知的原始消息
    wPad1 As Integer            '-
    wParam As Long              '该消息的wParam值
    wPad2 As Integer            '-
    lParam As Long              '该消息的lParam值
    chrg As CHARRANGE           '当前选择内容
End Type

'剪贴板中的对象和文本的内容
Public Type ENSAVECLIPBOARD
    NMHDR As NMHDR     '通知头
    cObjectCount As Long        '剪贴板中对象数目
    cch As Long                 '剪贴板中字符数目
End Type

'失败的OLE操作相关信息
' #ifndef MACPORT
Public Type ENOLEOPFAILED
    NMHDR As NMHDR     '通知头
    iob As Long                 '对象索引值
    lOper As Long               '失败的OLE操作，取值为 OLEOP_DOVERB 常数
    hr As Long                  '返回的错误代码
End Type
' #End If

'对象定位信息，在对象被读入RTB时产生该通知
Public Type OBJECTPOSITIONS
    NMHDR As NMHDR     '通知头
    cObjectCount As Long        '对象数量
        ' !!!POINTER to long value!!!
    pcpPositions As Long        '对象位置指针。注意：是长整形的指针！！！！
End Type

Public Type ENLINK
    NMHDR As NMHDR     '通知头
    Msg As Integer              '触发本通知的消息
    wPad1 As Integer            '-
    wParam As Integer           '该消息的wParam值
    wPad2 As Integer            '-
    lParam As Integer           '该消息的lParam值
    chrg As CHARRANGE           '超链接文本范围
End Type

' /* PenWin specific */
Public Type ENCORRECTTEXT
    NMHDR As NMHDR     '通知头
    chrg As CHARRANGE           '当前选择范围
    seltyp As Integer           '范围中内容的类型
End Type

' 选择性粘贴
Public Type REPASTESPECIAL
    dwAspect As Long    '显示特性。取值：DVASPECT_CONTENT 或者 DVASPECT_ICON
    dwParam As Long     '如果为DVASPECT_ICON，则本参数包含一个指向该对象视图的一个图元文件句柄
End Type



' /* EM_GETTEXTEX 消息 wParam 参数 */
Public Type GETTEXTEX
    cb As Long              ' /* 读取的字符串字节数             */
    flags As Long           ' /* 文本转换操作选项               */
    codepage As Long        ' /* 转换的代码页，默认为CP_ACP，Unicode为1200
    lpDefaultChar As Long   ' /* 在Unicode模式下无法表示该字符时的替代字符，为NULL则使用系统默认值。 */
    lpUsedDefChar As Long   ' /* 是否启用替换字符   */
End Type


' /* EM_GETTEXTLENGTHEX 获取文本长度消息的 wParam 参数 */
Public Type GETTEXTLENGTHEX
    flags As Long                   ' 如上
    codepage As Long                ' 代码页
End Type
    
' /* BiDi specific features */
Public Type BIDIOPTIONS
    cbSize As Long
    wPad1 As Integer
    wMask As Integer
    wEffects As Integer
End Type
Public Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

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

Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Type POINTL
    X As Long
    Y As Long
End Type

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


'将一块内存从一个地方拷贝到另一个地方
'函数原型：
'VOID CopyMemory(
'  PVOID Destination,  // 目标拷贝的地址指针。
'  CONST VOID *Source, // 源拷贝的地址指针。
'  DWORD Length        // 源拷贝的字节大小。
')
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'作用同上，只是源为一个字符串
Public Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
'作用同上，只是目标为一个字符串
Public Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)

'#########################################################################

' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！
'函数原型：
'LRESULT SendMessage(
'  HWND hWnd,      // 目标窗体的句柄。
'  UINT Msg,       // 待发送的消息。
'  WPARAM wParam,  // 消息第一参数。
'  LPARAM lParam   // 消息第二参数。
');
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'作用同上，不过第二参数为Long型。
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'作用同上，不过第二参数为String型。
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

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
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'设置当前线程消息队列中的窗体获取键盘焦点
Public Declare Function GetFocus Lib "user32" () As Long
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

'获取显示器或者打印机的信息
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'设置指定画布的映射模式
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
'开始一个打印作业
Public Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
'通知打印设备准备接收数据。
Public Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
'通知打印机停止接收数据，通常用于换页
Public Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
'完成一次打印作业
Public Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
'删除指定设备场景（画布）
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'保存当前设备场景状态到上下文堆栈中。
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
'恢复设备场景状态。
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
'使用指定坐标指定设备场景视口的原点
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
'每个逻辑单位为1个设备象素。正X向右，正Y向下。用于SetMapMode()
Public Const MM_TEXT = 1
'乘以两个32位的数，然后用其64位结果除以第三个数，最后四舍五入。
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'#########################################################################
' 打印支持:
'调用打印对话框
Public Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As PrintDlg) As Long
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
'Shell调用
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'同上，不过第4、5参数为Any类型
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long

'#########################################################################
Public Declare Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
'流的支持:
'写入文件
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'打开文件
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'读取文件
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'关闭对象句柄
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
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
'判断矩形与矩形、矩形与椭圆是否相交
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Declare Function FlatSB_EnableScrollBar Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Public Declare Function FlatSB_ShowScrollBar Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Public Declare Function FlatSB_GetScrollRange Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, ByVal LPINT1 As Long, ByVal LPINT2 As Long) As Long
Public Declare Function FlatSB_GetScrollInfo Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Public Declare Function FlatSB_GetScrollPos Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long) As Long
Public Declare Function FlatSB_GetScrollProp Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal propIndex As Long, ByVal LPINT As Long) As Long
Public Declare Function FlatSB_SetScrollPos Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, ByVal Pos As Long, ByVal fRedraw As Boolean) As Long
Public Declare Function FlatSB_SetScrollInfo Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Public Declare Function FlatSB_SetScrollRange Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal code As Long, ByVal Min As Long, ByVal Max As Long, ByVal fRedraw As Boolean) As Long
Public Declare Function FlatSB_SetScrollProp Lib "COMCTL32.DLL" (ByVal hWnd As Long, ByVal Index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long
Public Declare Function InitializeFlatSB Lib "COMCTL32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function UninitializeFlatSB Lib "COMCTL32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, _
                                                                    ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Public Declare Function DrawStateString Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, _
                                                                        ByVal lpString As String, ByVal cbStringLen As Long, ByVal X As Long, _
                                                                        ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Public Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT, pContentRect As RECT) As Long
Public Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Public Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long



Public Declare Function PaintRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Public Declare Function DrawTextExA Lib "user32" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lp As Long) As Long
Public Declare Function DrawTextExAsNull Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function ImageList_GetBkColor Lib "COMCTL32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_Convert Lib "COMCTL32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Public Declare Function ImageList_Replace Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hBmMask As Long) As Long
Public Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, hBmMask As Long) As Long
Public Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long
Public Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Long
Public Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Public Declare Function ImageList_SetImageCount Lib "COMCTL32" (ByVal hImageList As Long, uNewCount As Long)
Public Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_SetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, prcImage As RECT) As Long
Public Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_LoadImage Lib "COMCTL32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)
Public Declare Function ImageList_SetBkColor Lib "COMCTL32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
Public Declare Function ImageList_Copy Lib "COMCTL32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
'繁简转换
Public Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Public Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'鼠标位置信息
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态
Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'释放内存
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long

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
Public Function HIWORD(LongIn As Long) As Long
   ' 取出32位值的高16位
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**函 数 名：LOWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的低16位
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Long
   ' 取出32位值的低16位
     LOWORD = LongIn And &HFFFF&
End Function





