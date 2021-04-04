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
Public glngPro As Long
Public glngGdi As Long
Public Declare Function GetGuiResources Lib "User32.dll" (ByVal hProcess As Long, ByVal uiFlags As Integer) As Long

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
Public Const WM_MOUSEWHEEL = &H20A          '鼠标滚动
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
Public Const EN_CHANGE = &H300          '内容改变。父窗体通过 WM_COMMAND 消息获取该通知。
Public Const EN_ERRSPACE = &H500        '内容不足以分配该操作。
Public Const EN_HSCROLL = &H601         '水平滚动事件。
Public Const EN_KILLFOCUS = &H200       '失去焦点事件。
Public Const EN_MAXTEXT = &H501         '输入的文本超过最大字符数。或者在非自动滚动时超出控件可视区域。
Public Const EN_SETFOCUS = &H100        '获得键盘输入焦点。
Public Const EN_UPDATE = &H400          '在用户改变内容但是还没有刷新显示时发出该通知。用户可以用于调节控件尺寸以适应内容。
Public Const EN_VSCROLL = &H602         '垂直滚动事件。
Public Const WM_CONTEXTMENU = &H7B&     '通知窗体的右键点击事件
Public Const WM_PRINTCLIENT = &H318&    '请求绘制其客户区域到一个指定的设备上下文中，通常是指打印机。
Public Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Public Const LF_FACESIZE = 32   '字体名称字节长度。
Public Const RICHEDIT_VER = &H210    '当前Rich Edit控件版本号
Public Const cchTextLimitDefault = 32767&       '默认文本长度限制
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0
Public Const RICHEDIT_CLASSW = "RichEdit20W"
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSW       'UNICODE版本！
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
Public Const EM_AUTOURLDETECT = (WM_USER + 91)      '启用/禁用自动URL检测。
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)   '判断是否启用了自动URL检测。
Public Const EM_SETPALETTE = (WM_USER + 93)         '改变调色板。
Public Const EM_GETTEXTEX = (WM_USER + 94)          '获取指定代码页的文本。
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)    '采用不同方式计算文本长度。
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
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)    '设置当前双向语言支持选项。
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)    '获取当前双向语言支持选项。
Public Const IMF_AUTOKEYBOARD = &H1             '自动键盘布局
Public Const IMF_AUTOFONT = &H2                 '自动字体
Public Const IMF_IMECANCELCOMPLETE = &H4      '// high completes the comp string when aborting, low cancels.
Public Const IMF_IMEALWAYSSENDNOTIFY = &H8
Public Const ICM_NOTOPEN = &H0          'Input Method Editor (IME) is not open.
Public Const ICM_LEVEL3 = &H1           'True inline mode.
Public Const ICM_LEVEL2 = &H2           'Level 2.
Public Const ICM_LEVEL2_5 = &H3         'Level 2.5
Public Const ICM_LEVEL2_SUI = &H4       'Special user interface (UI).
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
Public Const EN_ALIGN_LTR = &H710&      'RTB控件通过 WM_COMMAND 消息通知父窗体段落方向改为从左至右。
Public Const EN_ALIGN_RTL = &H711&      'RTB控件通过 WM_COMMAND 消息通知父窗体段落方向改为从右至左。
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
Public Const ENM_IMECHANGE = &H800000           ' /* RE2.0 不支持！，只用于1.0版本！*/
Public Const ENM_LANGCHANGE = &H1000000         ' ？？
Public Const ENM_OBJECTPOSITIONS = &H2000000    '可以发送 EN_OBJECTPOSITIONS 消息。
Public Const ENM_LINK = &H4000000               '可以发送 EN_LINK 消息。
Public Const ES_SAVESEL = &H8000&               '在失去焦点时保持选择区域高亮显示！！！Useful！
Public Const ES_SUNKEN = &H4000&                '凹下效果
Public Const ES_DISABLENOSCROLL = &H2000&       '在不需要滚动条时将其置灰，而非隐藏
Public Const ES_SELECTIONBAR = &H1000000 ' /* same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value */
Public Const ES_NOOLEDRAGDROP = &H8 ' /* same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop */
Public Const ES_EX_NOCALLOLEINIT = &H1000000 ' /* 新的 Edit 控件扩展样式 */
Public Const ES_VERTICAL = &H400000     '垂直绘制文本和对象。
Public Const ES_NOIME = &H80000         '禁用IME。
Public Const ES_SELFIME = &H40000       '应用程序来控制IME操作。
Public Const WB_CLASSIFY = 3&           '
Public Const WB_MOVEWORDLEFT = 4&       '
Public Const WB_MOVEWORDRIGHT = 5&      '
Public Const WB_LEFTBREAK = 6&          '
Public Const WB_RIGHTBREAK = 7&         '
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
Public Const IMF_FORCENONE = &H1
Public Const IMF_FORCEENABLE = &H2
Public Const IMF_FORCEDISABLE = &H4
Public Const IMF_CLOSESTATUSWINDOW = &H8
Public Const IMF_VERTICAL = &H20
Public Const IMF_FORCEACTIVE = &H40
Public Const IMF_FORCEINACTIVE = &H80
Public Const IMF_FORCEREMEMBER = &H100
Public Const IMF_MULTIPLEEDIT = &H400
Public Const WBF_CLASS = &HF          '((BYTE) =&H0F)
Public Const WBF_ISWHITE = &H10       '((BYTE) =&H10)
Public Const WBF_BREAKLINE = &H20     '((BYTE) =&H20)
Public Const WBF_BREAKAFTER = &H40    '((BYTE) =&H40)
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
Public Const CFE_BOLD = &H1&            '粗体
Public Const CFE_ITALIC = &H2&          '斜体
Public Const CFE_UNDERLINE = &H4&       '下划线
Public Const CFE_STRIKEOUT = &H8&       '删除线
Public Const CFE_PROTECTED = &H10&      '保护
Public Const CFE_LINK = &H20&           '超链接
Public Const CFE_AUTOCOLOR = &H40000000 '采用系统自动颜色。' /* NOTE: this corresponds to */Public Const SCF_SELECTION = &H1&   '应用于当前选中区域。
Public Const SCF_WORD = &H2&        '应用于当前选中单词。
Public Const SCF_DEFAULT = &H0&            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4&                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8&         '// modifier for SCF_SELECTION; says that
Public Const SF_TEXT = &H1         'Text格式
Public Const SF_RTF = &H2          'RTF格式
Public Const SF_RTFNOOBJS = &H3    '输出时用空格代替对象，仅用于输出！
Public Const SF_TEXTIZED = &H4     '输出时采用文本表示对象，仅用于输出！
Public Const SF_UNICODE = &H10            ' /* Unicode file of some kind */
Public Const SFF_SELECTION = &H8000&    '输入输出只对当前选择区域有效！
Public Const SFF_PLAINRTF = &H4000&     '只使用通用RTF关键字，对于与语言相关的RTF关键字予以忽略！
Public Const MAX_TAB_STOPS = 32&    '绝对制表符的最大数目。
Public Const lDefaultTab = 720&     '默认绝对制表符位置。
Public Const PFM_STARTINDENT = &H1& '首行缩进值有效。
Public Const PFM_RIGHTINDENT = &H2& '右缩进值有效。
Public Const PFM_OFFSET = &H4&      '缩进或者悬挂有效！负值表示缩进，正值表示悬挂！
Public Const PFM_ALIGNMENT = &H8&   '水平对齐方式有效。
Public Const PFM_TABSTOPS = &H10&   '绝对制表符位置有效。
Public Const PFM_NUMBERING = &H20&  '编号与项目符号有效。
Public Const PFM_OFFSETINDENT = &H80000000  '首行缩进值有效，并且给出一个相对值。
Public Const PFN_BULLET = &H1&   ' /* PARAFORMAT 编号选项 */
Public Const PFA_LEFT = &H1&   ' /* PARAFORMAT 对齐选项 */     '
Public Const PFA_RIGHT = &H2&       '
Public Const PFA_CENTER = &H3&      '
Public Const CFM_EFFECTS = (CFM_BOLD Or CFM_ITALIC Or CFM_UNDERLINE Or CFM_COLOR Or CFM_STRIKEOUT Or CFE_PROTECTED Or CFM_LINK)
Public Const CFM_ALL = (CFM_EFFECTS Or CFM_SIZE Or CFM_FACE Or CFM_OFFSET Or CFM_CHARSET)
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
Public Const CFM_EFFECTS2 = (CFM_EFFECTS Or CFM_DISABLED Or CFM_SMALLCAPS Or CFM_ALLCAPS Or CFM_HIDDEN Or CFM_OUTLINE Or CFM_SHADOW Or CFM_EMBOSS Or CFM_IMPRINT Or CFM_DISABLED Or CFM_REVISED Or CFM_SUBSCRIPT Or CFM_SUPERSCRIPT Or CFM_BACKCOLOR)
Public Const CFM_ALL2 = (CFM_ALL Or CFM_EFFECTS2 Or CFM_BACKCOLOR Or CFM_LCID Or CFM_UNDERLINETYPE Or CFM_WEIGHT Or CFM_REVAUTHOR Or CFM_SPACING Or CFM_KERNING Or CFM_STYLE Or CFM_ANIMATION)
Public Const CFE_SMALLCAPS = CFM_SMALLCAPS
Public Const CFE_ALLCAPS = CFM_ALLCAPS
Public Const CFE_HIDDEN = CFM_HIDDEN
Public Const CFE_OUTLINE = CFM_OUTLINE
Public Const CFE_SHADOW = CFM_SHADOW
Public Const CFE_EMBOSS = CFM_EMBOSS
Public Const CFE_IMPRINT = CFM_IMPRINT
Public Const CFE_DISABLED = CFM_DISABLED
Public Const CFE_REVISED = CFM_REVISED
Public Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR
Public Const CFU_CF1UNDERLINE = &HFF&      ' /* map charformat's bit underline to CF2.*/
Public Const CFU_INVERT = &HFE&            ' /* For IME composition fake a selection.*/
Public Const CFU_UNDERLINEDOTTED = &H4&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEDOUBLE = &H3&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEWORD = &H2&      ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINE = &H1&
Public Const CFU_UNDERLINENONE = 0&
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
Public Const PFM_EFFECTS = (PFM_DIR Or PFM_KEEP Or PFM_KEEPNEXT Or PFM_TABLE Or PFM_PAGEBREAKBEFORE Or PFM_NOLINENUMBER Or PFM_NOWIDOWCONTROL Or PFM_DONOTHYPHEN Or PFM_SIDEBYSIDE Or PFM_TABLE)
Public Const PFM_ALL = (PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET Or PFM_ALIGNMENT Or PFM_TABSTOPS Or PFM_NUMBERING Or PFM_OFFSETINDENT Or PFM_DIR)
Public Const PFM_ALL2 = (PFM_ALL Or PFM_EFFECTS Or PFM_SPACEBEFORE Or PFM_SPACEAFTER Or PFM_LINESPACING Or PFM_STYLE Or PFM_SHADING Or PFM_BORDER Or PFM_NUMBERINGTAB Or PFM_NUMBERINGSTART Or PFM_NUMBERINGSTYLE)
Public Const PFE_TABLEROW = &HC000&                ' /* These 3 options are mutually */
Public Const PFE_TABLECELLEND = &H8000&            ' /*  exclusive and each imply    */
Public Const PFE_TABLECELL = &H4000&               ' /*  段落为表格的一部分 */
Public Const PFA_JUSTIFY = 4          ' /* 两端对齐，为了兼容TOM模型接口。 (*)  */
Public Const GCM_RIGHTMOUSEDROP = &H8000& ' 用于在 IRichEditOleCallback::GetContextMenu 函数中请求应用程序提供一个右键菜单。
Public Const CF_RTF = "Rich Text Format" ' 剪贴板格式，用于 RegisterClipboardFormat() 注册有效的剪贴板格式。
Public Const CF_RTFNOOBJS = "Rich Text Format Without Objects"
Public Const CF_RETEXTOBJ = "RichEdit Text and Objects"
Public Const BOM_DEFPARADIR = &H1&             ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOM_PLAINTEXT = &H2&              ' /* Use plain text layout (obsolete) */
Public Const BOM_NEUTRALOVERRIDE = &H4&        ' /* Override neutral layout (obsolete) */
Public Const BOM_CONTEXTREADING = &H8&         ' /* Context reading order */
Public Const BOM_CONTEXTALIGNMENT = &H10&      ' /* Context alignment */
Public Const BOE_RTLDIR = &H1&                 ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOE_PLAINTEXT = &H2&              ' /* Use plain text layout (obsolete) */
Public Const BOE_NEUTRALOVERRIDE = &H4&        ' /* Override neutral layout (obsolete) */
Public Const BOE_CONTEXTREADING = &H8&         ' /* Context reading order */
Public Const BOE_CONTEXTALIGNMENT = &H10&      ' /* Context alignment */
Public Const FR_MATCHDIAC = &H20000000          ' 阿拉伯与希伯来语用
Public Const FR_MATCHKASHIDA = &H40000000       ' 阿拉伯与希伯来语用
Public Const FR_MATCHALEFHAMZA = &H80000000     ' 阿拉伯与希伯来语用
Public Const WCH_EMBEDDING = &HFFFC&

Public Const MB_OK = &H0&

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
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE

Public Const SM_CYHSCROLL = 3 '水平滚动条上的矢量位图宽度
Public Const SM_CXVSCROLL = 2 '水平滚动条上的矢量位图高度
Public Const SC_MOVE = &HF012 '用户点击系统菜单中的“移动”菜单事件
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
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80&


Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOREDRAW = &H8
Public Const SWP_SHOWWINDOW = &H40
Public Const SPI_GETWORKAREA = 48&

Public Const OLEOP_DOVERB = 1
' 用于 GETTEXTEX 数据结构
Public Const GT_DEFAULT = 0&    '不使用CR转换
Public Const GT_USECRLF = 1&    '表示在每次拷贝文本时，将CR转换为CRLF。
' GETTEXTLENGTHEX 数据结构的标志位
Public Const GTL_DEFAULT = 0&      ' /* 默认值，返回字符数目。                      */
Public Const GTL_USECRLF = 1&      ' /* 使用段落 CR/LF 计算                         */
Public Const GTL_PRECISE = 2&      ' /* 精确计算，较慢                              */
Public Const GTL_CLOSE = 4&        ' /* 近似计算，较快，常用于提前分配内存空间      */
Public Const GTL_NUMCHARS = 8&     ' /* 返回字符数目                                */
Public Const GTL_NUMBYTES = 16&    ' /* 返回字节数目

'每个逻辑单位为1个设备象素。正X向右，正Y向下。用于SetMapMode()
Public Const MM_TEXT = 1
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version
Public Const DC_PAPERNAMES = 16 '纸张名称(每64字符为一段,以Chr(0)结束)
Public Const DC_PAPERS = 2 '纸张编号(Array or Word)
Public Const DC_BINNAMES = 12 '进纸方式(每24字符为一段,以Chr(0)结束)
Public Const DC_BINS = 6 '进纸编号(Array or Word)
'打印纸张常量(256=自定义)
Public Const PageSize1 = "信笺， 8 1/2 x 11 英寸"
Public Const PageSize2 = "+A611 小型信笺， 8 1/2 x 11 英寸"
Public Const PageSize3 = "小型报， 11 x 17 英寸"
Public Const PageSize4 = "分类帐， 17 x 11 英寸"
Public Const PageSize5 = "法律文件， 8 1/2 x 14 英寸"
Public Const PageSize6 = "声明书，5 1/2 x 8 1/2 英寸"
Public Const PageSize7 = "行政文件，7 1/2 x 10 1/2 英寸"
Public Const PageSize8 = "A3, 297 x 420 毫米"
Public Const PageSize9 = "A4, 210 x 297 毫米"
Public Const PageSize10 = "A4小号， 210 x 297 毫米"
Public Const PageSize11 = "A5, 148 x 210 毫米"
Public Const PageSize12 = "B4, 250 x 354 毫米"
Public Const PageSize13 = "B5, 182 x 257 毫米"
Public Const PageSize14 = "对开本， 8 1/2 x 13 英寸"
Public Const PageSize15 = "四开本， 215 x 275 毫米"
Public Const PageSize16 = "10 x 14 英寸"
Public Const PageSize17 = "11 x 17 英寸"
Public Const PageSize18 = "便条，8 1/2 x 11 英寸"
Public Const PageSize19 = "#9 信封， 3 7/8 x 8 7/8 英寸"
Public Const PageSize20 = "#10 信封， 4 1/8 x 9 1/2 英寸"
Public Const PageSize21 = "#11 信封， 4 1/2 x 10 3/8 英寸"
Public Const PageSize22 = "#12 信封， 4 1/2 x 11 英寸"
Public Const PageSize23 = "#14 信封， 5 x 11 1/2 英寸"
Public Const PageSize24 = "C 尺寸工作单"
Public Const PageSize25 = "D 尺寸工作单"
Public Const PageSize26 = "E 尺寸工作单"
Public Const PageSize27 = "DL 型信封， 110 x 220 毫米"
Public Const PageSize28 = "C5 型信封， 162 x 229 毫米"
Public Const PageSize29 = "C3 型信封， 324 x 458 毫米"
Public Const PageSize30 = "C4 型信封， 229 x 324 毫米"
Public Const PageSize31 = "C6 型信封， 114 x 162 毫米"
Public Const PageSize32 = "C65 型信封，114 x 229 毫米"
Public Const PageSize33 = "B4 型信封， 250 x 353 毫米"
Public Const PageSize34 = "B5 型信封，176 x 250 毫米"
Public Const PageSize35 = "B6 型信封， 176 x 125 毫米"
Public Const PageSize36 = "信封， 110 x 230 毫米"
Public Const PageSize37 = "信封大王， 3 7/8 x 7 1/2 英寸"
Public Const PageSize38 = "信封， 3 5/8 x 6 1/2 英寸"
Public Const PageSize39 = "U.S. 标准复写簿， 14 7/8 x 11 英寸"
Public Const PageSize40 = "德国标准复写簿， 8 1/2 x 12 英寸"
Public Const PageSize41 = "德国法律复写簿， 8 1/2 x 13 英寸"

'WinNT自定义纸张控制================================================================
' Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
' Constants for PRINTER_DEFAULTS.DesiredAccess
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
' Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2

Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
'VB Errors
Public Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Public Const DSna = &H220326

Public Const MM_LOMETRIC = 2
Public Const MM_HIMETRIC = 3
Public Const MM_LOENGLISH = 4
Public Const MM_HIENGLISH = 5
Public Const MM_TWIPS = 6
Public Const MM_ISOTROPIC = 7
Public Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
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

'获取窗体显示元素的当前颜色值
Public Const COLOR_SCROLLBAR = 0 '滚动条
Public Const COLOR_BACKGROUND = 1 'windows桌面
Public Const COLOR_ACTIVECAPTION = 2 '活动窗口的标题
Public Const COLOR_INACTIVECAPTION = 3 '不活动窗口的标题。
Public Const COLOR_MENU = 4 '菜单。
Public Const COLOR_WINDOW = 5 '窗口背景
Public Const COLOR_WINDOWFRAME = 6 '窗框
Public Const COLOR_MENUTEXT = 7 '菜单正文。
Public Const COLOR_WINDOWTEXT = 8 '窗口正文
Public Const COLOR_CAPTIONTEXT = 9 '窗口标题中的文字
Public Const COLOR_ACTIVEBORDER = 10 '活动窗口的边框
Public Const COLOR_INACTIVEBORDER = 11 '不活动窗口的边框。
Public Const COLOR_APPWORKSPACE = 12 'MDI桌面的背景
Public Const COLOR_HIGHLIGHT = 13 '选定的项目背景。
Public Const COLOR_HIGHLIGHTTEXT = 14 '选定的项目文字。
Public Const COLOR_BTNFACE = 15 '按钮
Public Const COLOR_BTNSHADOW = 16 '按钮的3D阴影
Public Const COLOR_GRAYTEXT = 17 '灰色文字；如使用了抖动技术则为零。
Public Const COLOR_BTNTEXT = 18 '按钮文字
Public Const COLOR_INACTIVECAPTIONTEXT = 19 '不活动窗口的文字。
Public Const COLOR_BTNHIGHLIGHT = 20 '按钮的3D加亮区
Public Const COLOR_3DDKSHADOW = 21 '3D深阴影。
Public Const COLOR_3DLIGHT = 22 '3D阴影化对象的浅色。
Public Const COLOR_INFOTEXT = 23 '工具提示的文本色。
Public Const COLOR_INFOBK = 24 '工具提示的背景色。
Public Const COLOR_HOTLIGHT = 26
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100  'shorts
Public Const COLOR_3DFACE = COLOR_BTNFACE '3D阴影化对象的正面颜色。
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT '3D加亮颜色（win95）。
Public Const COLOR_DESKTOP = COLOR_BACKGROUND '桌面颜色。
Public Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT

Public Const WS_CAPTION = &HC00000
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'3D效果颜色
Public Const LR_LOADMAP3DCOLORS = &H1000
'图片从文件lpsz中调入，而非从资源文件中调入。
Public Const LR_LOADFROMFILE = &H10
'调入透明色
Public Const LR_LOADTRANSPARENT = &H20
'生成 设备无关 DIB 位图，而非设备相关位图。
Public Const IMAGE_BITMAP = 0

Public Const MAX_PATH = 260
Public Const DI_NORMAL = &H3
Public Const SHGFI_LARGEICON = &H0                      '获取大图标
Public Const SHGFI_SMALLICON = &H1                      '获取小图标
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_ICON = &H100                         '获取图标句柄

'Clipboard------------------------------------------------
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_OEMTEXT = 7
Public Const CF_ENHMETAFILE = 14
Public Const CF_HDROP = 15
Public Const CF_DSPTEXT = &H81
Public Const CF_DSPBITMAP = &H82
Public Const CF_DSPMETAFILEPICT = &H83

'共享文件夹
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

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
' 虚拟键编码常数
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt
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

Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

'矩形
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Size
    cx As Long
    cy As Long
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

Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                 ' /* 默认行为 */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8           ' /* 默认行为 */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32           ' /* 默认行为 */
End Enum

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

' /* 所有段落度量单位均为：缇 */
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

Public Type NMHDR ' /* 通知的结构 */
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
    msg As Integer          '键盘或者鼠标标识符
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
    msg As Long                 '触发该通知的原始消息
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
Public Type ENOLEOPFAILED
    NMHDR As NMHDR     '通知头
    iob As Long                 '对象索引值
    lOper As Long               '失败的OLE操作，取值为 OLEOP_DOVERB 常数
    hr As Long                  '返回的错误代码
End Type

'对象定位信息，在对象被读入RTB时产生该通知
Public Type OBJECTPOSITIONS
    NMHDR As NMHDR     '通知头
    cObjectCount As Long        '对象数量
        ' !!!POINTER to long value!!!
    pcpPositions As Long        '对象位置指针。注意：是长整形的指针！！！！
End Type

Public Type ENLINK
    NMHDR As NMHDR     '通知头
    msg As Integer              '触发本通知的消息
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

Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As Long  ' ACL
    Dacl As Long  ' ACL
End Type

' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
    flags As Long
    pName As Long   ' String
    Size As Size
    ImageableArea As RECT
End Type

Public Type sFORM_INFO_1
    flags As Long
    pName As String
    Size As Size
    ImageableArea As RECT
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

Public Type PRINTER_DEFAULTS
    pDatatype As String
    pDevMode As Long    ' DEVMODE
    DesiredAccess As Long
End Type

Public Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As DEVMODE
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type
Public Type SHFILEINFO
        hIcon As Long                       '  out: icon
        iIcon As Long                       '  out: icon index
        dwAttributes As Long                '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH  '  out: display name (or path)
        szTypeName As String * 80           '  out: type name
End Type
Public Type ENHMETAHEADER
   iType As Long
   nSize As Long
   rclBounds As RECT
   rclFrame As RECT
   dSignature As Long
   nVersion As Long
   nBytes As Long
   nRecords As Long
   nHandles As Integer
   sReserved As Integer
   nDescription As Long
   offDescription As Long
   nPalEntries As Long
   szlDevice As Size
   szlMillimeters As Size
End Type

Public Type NETRESOURCE ' 网络资源
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

'绘制半透明矩形框相关函数与声明
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

' cAlphaDibSection functions:
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

'Window版本函数
Public Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' 打印支持:
' VB API Viewer 版本的 DocInfo 结构说明是错误的！！！！
'用于存储 StartDoc() 中文件名及其他信息
Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

'用于初始化打印对话框及返回值
Type PrintDlg
    lStructSize As Long
    hwndOwner As Long
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
Public Type BitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

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
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
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

Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
                ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, _
                ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long


Public Declare Function GetTickCount Lib "kernel32" () As Long

''''''设置窗体最大化最小化关闭按钮
Public Declare Function GetSystemMenu Lib "user32 " (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32 " (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32 " (ByVal hwnd As Long) As Long
Public Declare Function RemoveMenu Lib "user32 " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'扫描函数声明
Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwnd As Long, ByVal wPixTypes As Integer) As Integer
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwnd As Long) As Integer
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Public Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Public Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, lpRect As RECT) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function CountClipboardFormats Lib "user32" () As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'绘制透明图形
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '系统暂停
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'注册表直接操作
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long
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
'将一块内存从一个地方拷贝到另一个地方 lpvDest目标拷贝的地址指针。lpvSource源拷贝的地址指针。cbCopy源拷贝的字节大小。
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'作用同上，只是源为一个字符串
Public Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
'作用同上，只是目标为一个字符串
Public Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)
'用于分屏显示消息
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！HWND hWnd 目标窗体的句柄。wMsg待发送的消息。wParam消息第一参数。lParam消息第二参数。
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'作用同上，不过第二参数为Long型。
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'作用同上，不过第二参数为String型。
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'设置窗体状态（最大化、最下化、隐藏等）
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'移动窗体
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'要求窗体刷新
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'锁定/解锁窗体的刷新
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'销毁窗体及相关资源
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'屏蔽/恢复鼠标及键盘的输入
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'搜索指定条件的窗体
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
'改变指定窗体的父窗体
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'获取当前对象所在窗体：
'窗体层次有5层：Frame、Document、Pane、Parent、In-place
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'获取指定窗体的边界矩形尺寸
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'获取客户区域矩形
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'获取窗体属性
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'设置窗体属性
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'移除窗体属性
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'返回包含了指定点的窗口的句柄。
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'将屏幕上某个点的屏幕坐标转换为客户区域坐标
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'将一个窗体相关的坐标空间映射到另一个窗体的坐标空间
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'取消鼠标捕获
Public Declare Function ReleaseCapture Lib "user32" () As Long
'获取鼠标屏幕坐标位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'指定客户区域的一个即将被刷新的矩形区域
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'同上，不过参数2是一个指针了
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'创建指定属性的窗体
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'将消息传送到指定的窗体过程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function SetWindowLongL Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'设置当前线程消息队列中的窗体获取键盘焦点
Public Declare Function GetFocus Lib "user32" () As Long
'获取焦点
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
'将指定的可执行模块（.DLL/.EXE）映射到调用过程的地址空间
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'减少DLL的引用数目
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Integer, ByVal crColor As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'获取窗体显示元素的当前颜色值
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
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
'纸张打印边界控制
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
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
'乘以两个32位的数，然后用其64位结果除以第三个数，最后四舍五入。
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'指定窗体进行绘图准备，通过PAINTSTRUCT结构体来初始化。
Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
'在绘图完成后，标记窗体绘图结束。
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
'用于获取给定绘图对象的信息。
'取决于绘图对象的不同，可以在给定缓冲区中填入BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT 或者 LOGPEN 结构
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'将一个对象选入指定的设备场景（画布）中，该对象自动替换掉同一类型的前一对象。
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'删除一个逻辑画笔、画刷、字体、位图、区域或者调色板
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'获取给定窗口或者整个屏幕的画布，用于在上面绘图。
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'释放标准Windows设备场景资源。
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'创建兼容的内存设备场景
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'创建设备相关位图
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'创建指定纯色的逻辑画刷
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'使用指定画刷填充矩形区域
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'从源画布到目标画布的比特块传送其彩色数据
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'返回桌面窗体（屏幕）的句柄
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'获取系统度量单位和系统设置，所有尺寸均以点 Pixel 表示
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Shell调用
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'同上，不过第4、5参数为Any类型
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
'判断当前是否某个虚拟键按下或者放开
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'人工合成鼠标动作和点击事件，新标准应该使用 SendInput() 函数！
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'写入文件
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'打开文件
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'读取文件
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'关闭对象句柄
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'API作图
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DrawStatusText Lib "COMCTL32" Alias "DrawStatusTextA" (ByVal hDC As Long, lprc As RECT, ByVal pszText As String, ByVal uFlags As Long) As Long
'判断矩形与矩形、矩形与椭圆是否相交
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal N As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "Psapi" (ByVal hProcess As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'扩展的 Shell 命令
Public Function ShellEx( _
        ByVal sFIle As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFIle), ".EXE") <> 0) Then
        eShowCmd = 0    '隐藏
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then   'Shell 调用
        lR = ShellExecuteForExplore(Owner, sOperation, sFIle, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFIle, sParameters, sDefaultDir, eShowCmd)
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

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    lngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(ByRef frmParent As Form, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, Optional Title As String = "") As Long

    MsgBoxD = MessageBox(frmParent.hwnd, Prompt, Title, Buttons)

End Function

