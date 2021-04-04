Attribute VB_Name = "mRTBSDK"
'#########################################################################
'##模 块 名：mRTBSDK.bas
'##创 建 人：吴庆伟
'##日    期：2005年3月25日
'##修 改 人：
'##日    期：
'##描    述：通用的 RTB SDK API 声明 (2.0版本)
'##版    本：
'#########################################################################

Option Explicit

Public Const LF_FACESIZE = 32   '字体名称字节长度。
Public Const RICHEDIT_VER = &H210    '当前Rich Edit控件版本号
Public Const cchTextLimitDefault = 32767&       '默认文本长度限制
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0
Public Const RICHEDIT_CLASSW = "RichEdit20W"
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSW       'UNICODE版本！
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
' #ifdef _WIN32
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80) '获取当前注册的扩展断字处理过程的地址。
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81) '设置当前扩展断字处理过程。0则恢复为默认。
' #End If

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


' /* 新的数据类型 */

' #ifdef _WIN32
' /* extended edit word break proc (character set aware) */
'typedef LONG (*EDITWORDBREAKPROCEX)(char *pchText, LONG cchText, BYTE bCharSet, INT action);
' #End If

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

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const CHARFORMAT CHARFORMATW
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* UNICODE */
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

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

'typedef struct _textrangew
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;   ' /* allocated by caller, zero terminated by RichEdit */
'} TEXTRANGEW;

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const TEXTRANGE   TEXTRANGEW
' #Else
'public const TEXTRANGE   TEXTRANGEA
' #endif ' /* UNICODE */
' #Else
'public const TEXTRANGE   TEXTRANGEA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */


'typedef DWORD (CALLBACK *EDITSTREAMCALLBACK)(DWORD dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb);

'用于存储 EM_STREAMIN 或者 EM_STREAMOUT 消息传递的数据信息。
Public Type EDITSTREAM
    dwCookie As Long     ' /* user value passed to callback as first parameter */
    dwError As Long      ' /* last error */
    pfnCallback As Long  'EDITSTREAMCALLBACK
End Type

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

'用于 EM_FINDTEXT 消息的查找文本的相关信息
Public Type FindText
    chrg As CHARRANGE   '字符范围
    lpstrText As Long   '需要查找的文本
End Type

'typedef struct _findtextw
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;
'} FINDTEXTW;'

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const FINDTEXT    FINDTEXTW
' #Else
'public const FINDTEXT    FINDTEXTA
' #endif ' /* UNICODE */
' #Else
'public const FINDTEXT    FINDTEXTA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

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

'typedef struct _findtextexw
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;
'    CHARRANGE chrgText;
'} FINDTEXTEXW;'

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const FINDTEXTEX  FINDTEXTEXW
' #Else
'public const FINDTEXTEX  FINDTEXTEXA
' #endif ' /* UNICODE */
' #Else
'public const FINDTEXTEX  FINDTEXTEXA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

'包含用于格式化指定设备的相关信息
Public Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

' /* 所有段落度量单位均为：缇 */

Public Const MAX_TAB_STOPS = 32&    '绝对制表符的最大数目。
Public Const lDefaultTab = 720&     '默认绝对制表符位置。

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

' /* CHARFORMAT2 and PARAFORMAT2 structures */

' #ifdef __cplusplus

'struct CHARFORMAT2W : _charformatw
'{
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'};

'struct CHARFORMAT2A : _charformat
'{
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'};

' #else   ' /* regular C-style  */

'type C
'{
'    UINT        cbSize;
''    _WPAD       _wPad1;
 '   DWORD       dwMask;
 '   DWORD       dwEffects;
 '   LONG        yHeight;
 ''   LONG        yOffset;            ' /* > 0 for superscript, < 0 for subscript */
'    COLORREF    crTextColor;
'    BYTE        bCharSet;
'    BYTE        bPitchAndFamily;
'    WCHAR       szFaceName[LF_FACESIZE];
'    _WPAD       _wPad2;
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'    BYTE        bReserved1;
'} CHARFORMAT2W;


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

' #endif ' /* C++ */

' #ifdef UNICODE
'public const CHARFORMAT2 CHARFORMAT2W
' #Else
'public const CHARFORMAT2 CHARFORMAT2A
' #End If

'public Const CHARFORMATDELTA = (Len(CHARFORMAT2) - Len(CHARFORMAT))


' /* CHARFORMAT and PARAFORMAT "ALL" masks
'   CFM_COLOR mirrors CFE_AUTOCOLOR, a little hack to easily deal with autocolor*/

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

' #ifdef __cplusplus
'struct PARAFORMAT2 : _paraformat
'{
'    LONG    dySpaceBefore;          ' /* Vertical spacing before para         */
'    LONG    dySpaceAfter;           ' /* Vertical spacing after para          */
'    LONG    dyLineSpacing;          ' /* Line spacing depending on Rule       */
'    SHORT   sStyle;                 ' /* Style handle                         */
'    BYTE    bLineSpacingRule;       ' /* Rule for line spacing (see tom.doc)  */
'    BYTE    bCRC;                   ' /* Reserved for CRC for rapid searching */
'    WORD    wShadingWeight;         ' /* Shading in hundredths of a per cent  */
'    WORD    wShadingStyle;          ' /* Nibble 0: style, 1: cfpat, 2: cbpat  */
'    WORD    wNumberingStart;        ' /* Starting value for numbering         */
'    WORD    wNumberingStyle;        ' /* Alignment, roman/arabic, (), ), ., etc.*/
'    WORD    wNumberingTab;          ' /* Space bet FirstIndent and 1st-line text*/
'    WORD    wBorderSpace;           ' /* Space between border and text (twips)*/
'    WORD    wBorderWidth;           ' /* Border pen width (twips)             */
'    WORD    wBorders;               ' /* Byte 0: bits specify which borders   */
'                                    ' /* Nibble 2: border style, 3: color index*/
'};

' #else   ' /* regular C-style  */

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

' /*
' *  PARAFORMAT numbering options (values for wNumbering):
' *
' *      Numbering Type      Value   Meaning
' *      tomNoNumbering        0     Turn off paragraph numbering
' *      tomNumberAsLCLetter   1     a, b, c, ...
' *      tomNumberAsUCLetter   2     A, B, C, ...
' *      tomNumberAsLCRoman    3     i, ii, iii, ...
' *      tomNumberAsUCRoman    4     I, II, III, ...
' *      tomNumberAsSymbols    5     default is bullet
' *      tomNumberAsNumber     6     0, 1, 2, ...
' *      tomNumberAsSequence   7     tomNumberingStart is first Unicode to use
' *
' *  Other valid Unicode chars are Unicodes for bullets.
' */


Public Const PFA_JUSTIFY = 4          ' /* 两端对齐，为了兼容TOM模型接口。 (*)  */


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

' /* used with IRichEditOleCallback::GetContextMenu, this flag will be
'   passed as a "selection type".  It indicates that a context menu for
'   a right-mouse drag drop should be generated.  The IOleObject parameter
'   will really be the IDataObject for the drop
' */
' 用于在 IRichEditOleCallback::GetContextMenu 函数中请求应用程序提供一个右键菜单。
Public Const GCM_RIGHTMOUSEDROP = &H8000&

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

Public Const OLEOP_DOVERB = 1

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

' /* Far East specific */
'typedef struct _punctuation
'{
'    UINT    iSize;
'    LPSTR   szPunctuation;
'} PUNCTUATION;

' /* Far East specific */
'typedef struct _compcolor
'{
'    COLORREF crText;
'    COLORREF crBackground;
'    DWORD dwEffects;
'}COMPCOLOR;


' 剪贴板格式，用于 RegisterClipboardFormat() 注册有效的剪贴板格式。
Public Const CF_RTF = "Rich Text Format"
Public Const CF_RTFNOOBJS = "Rich Text Format Without Objects"
Public Const CF_RETEXTOBJ = "RichEdit Text and Objects"

' 选择性粘贴
Public Type REPASTESPECIAL
    dwAspect As Long    '显示特性。取值：DVASPECT_CONTENT 或者 DVASPECT_ICON
    dwParam As Long     '如果为DVASPECT_ICON，则本参数包含一个指向该对象视图的一个图元文件句柄
End Type


' /* 用于下面的 GETTEXTEX 数据结构 */
Public Const GT_DEFAULT = 0&    '不使用CR转换
Public Const GT_USECRLF = 1&    '表示在每次拷贝文本时，将CR转换为CRLF。

' /* EM_GETTEXTEX 消息 wParam 参数 */
Public Type GETTEXTEX
    cb As Long              ' /* 读取的字符串字节数             */
    flags As Long           ' /* 文本转换操作选项               */
    codepage As Long        ' /* 转换的代码页，默认为CP_ACP，Unicode为1200
    lpDefaultChar As Long   ' /* 在Unicode模式下无法表示该字符时的替代字符，为NULL则使用系统默认值。 */
    lpUsedDefChar As Long   ' /* 是否启用替换字符   */
End Type

' GETTEXTLENGTHEX 数据结构的标志位
Public Const GTL_DEFAULT = 0&      ' /* 默认值，返回字符数目。                      */
Public Const GTL_USECRLF = 1&      ' /* 使用段落 CR/LF 计算                         */
Public Const GTL_PRECISE = 2&      ' /* 精确计算，较慢                              */
Public Const GTL_CLOSE = 4&        ' /* 近似计算，较快，常用于提前分配内存空间      */
Public Const GTL_NUMCHARS = 8&     ' /* 返回字符数目                                */
Public Const GTL_NUMBYTES = 16&    ' /* 返回字节数目                                */

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
' #ifndef WCH_EMBEDDING
Public Const WCH_EMBEDDING = &HFFFC&
' #endif ' /* WCH_EMBEDDING */
        

' #undef _WPAD

' #ifdef _WIN32
' #include <poppack.h>
' #elif !defined(RC_INVOKED)
' #pragma pack()
' #End If

' #ifdef __cplusplus
'}
' #endif  ' /* __cplusplus */

' #endif ' /* !_RICHEDIT_ */


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
'#End If ' /* WINVER >= =&H0400 */
'/*
' * Edit 控件样式
' */
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
'#if(WINVER >= =&H0400)
Public Const ES_NUMBER = &H2000&        '只允许输入数字。
'#endif /* WINVER >= =&H0400 */

'/* Edit 控件通知消息 */
Public Const EN_CHANGE = &H300          '内容改变。父窗体通过 WM_COMMAND 消息获取该通知。
Public Const EN_ERRSPACE = &H500        '内容不足以分配该操作。
Public Const EN_HSCROLL = &H601         '水平滚动事件。
Public Const EN_KILLFOCUS = &H200       '失去焦点事件。
Public Const EN_MAXTEXT = &H501         '输入的文本超过最大字符数。或者在非自动滚动时超出控件可视区域。
Public Const EN_SETFOCUS = &H100        '获得键盘输入焦点。
Public Const EN_UPDATE = &H400          '在用户改变内容但是还没有刷新显示时发出该通知。用户可以用于调节控件尺寸以适应内容。
Public Const EN_VSCROLL = &H602         '垂直滚动事件。



