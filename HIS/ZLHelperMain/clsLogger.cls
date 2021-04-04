VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@模块 clsLogger-2019/6/12
'@编写 lshuo
'@功能
'   日志记录模块，可以根据日志级别选择性记录日志
'@引用
'   Microsoft Scripting Runtime scrrun.dll
'@备注
'   1、日志级别设置的优先级。部件内模块日志级别>部件日志级别>进程日志级别
'       当进程日志级别为LogLevel_AllLog(6),则进程日志级别获取最高优先级。即程序内所有日志均记录
'   2、源码中的缺省日志级别，请使用AddComponentLogLevel方法设置，不要使用SetComponentLogLevel，除非特殊需要
'       这样能保证注册表日志级别优先大于源码缺省日志优先级
'   3、注册表日志级别设置
'       HHKEY_CURRENT_USER\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT\公共模块\
'            1）进程日志级别
'                日志跟踪\进程
'                    进程文件名 = 日志级别   （键值名=键值的值）
'            2)部件日志级别
'                日志跟踪\部件\部件名
'                   default=日志级别 （部件的日志级别）
'                    模块1=日志级别  （模块的日志级别）
'   4、缺省日志是打开的，但是日志缓存只有128行，超过，则自动清空缓存。在打开日志后请尽量早的指定日志名，否则会产生日志一楼。
'   5、1）每个方法中日志的写法
'        On Error GoTo ErrH
'        Call Logger.PushMethod("{PROJECT_NAME}.{MODULE_NAME}.{PROCEDURE_NAME}")
'        '个人代码
'        {PROCEDURE_BODY}
'        Call Logger.PopMethod("{PROJECT_NAME}.{MODULE_NAME}.{PROCEDURE_NAME}")
'        Exit {PROCEDURE_TYPE}
'   ErrH:
'        If Logger.ErrCenter("{PROJECT_NAME}.{MODULE_NAME}.{PROCEDURE_NAME}") = 1 Then
'            Resume
'        End If
'        '书写个人错误处理
'        Call Logger.PopMethod("{PROJECT_NAME}.{MODULE_NAME}.{PROCEDURE_NAME}")
'       2）每个Exit Sub,Exit Function、Exit Property 前均要加代码 Call Logger.PopMethod("{PROJECT_NAME}.{MODULE_NAME}.{PROCEDURE_NAME}")
'       3)可以调用Error、Warn、Info、DebugEx、Trace书写特定的级别的日志，也可以调用Log指定书写的日志级别
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'@原型
'    DWORD GetCurrentProcessId(
'
'    );
'@功能
'    检索调用进程的进程标识符。
'@参数
'    这个函数没有参数?
'@返回值
'    返回值是调用进程的进程标识符?
'@备注
'    在进程终止之前，进程标识符在整个系统中唯一地标识进程。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Declare Function GetTickCount Lib "kernel32" () As Long
'功能：检索自系统启动以来已运行的毫秒数，最多可达49.7天。
'返回：返回值是自系统启动以来运行的毫秒数。
'注意事项：GetTickCount函数的解析仅限于系统计时器的精度，通常在10毫秒到16毫秒之间。
'        GetTickCount函数的解析不会受到getsystemtime调适函数的调整的影响?
'        经过的时间存储为DWORD的值?
'        因此，如果系统连续运行49.7天，那么时间将会是零。
'        为了避免这个问题，请使用GetTickCount64函数。
'        否则，在比较时检查溢出条件。
'        如果你需要一个更高的分辨率计时器，可以使用多媒体定时器或高分辨率计时器。
'        为了获得计算机启动后的时间，在注册表关键hkeyperformance cedata的性能数据中检索系统的时间计数器。
'        返回的值是一个8字节的值?
'        要了解更多信息，请参见性能计数器。
'        Note：获得时间系统在工作状态自启动 , 使用QueryUnbiasedInterruptTime函数?
'        调试注意QueryUnbiasedInterruptTime函数产生不同的结果(“checked”)构建Windows,因为中断时间计数是由大约49天。
'        这有助于识别在系统运行很长时间之前可能不会发生的错误?
'        通过Microsoft Developer Network(MSDN)Web站点可以对MSDN的用户进行检查。
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
'@功能
'    将字符串映射到UTF-16(宽字符)字符串。字符串不一定来自多字节字符集。
'    错误地使用MultiByteToWideChar函数会损害应用程序的安全性。调用这个函数很容易导致缓冲区溢出，因为lpMultiByteStr表示的输入缓冲区的大小等于字符串中的字节数，而lpWideCharStr表示的输出缓冲区的大小等于字符数。为了避免缓冲区溢出，应用程序必须为缓冲区接收的数据类型指定适当的缓冲区大小。有关更多信息，请参见安全性考虑:国际特性。
'    注意，ANSI代码页可以在不同的计算机上不同，也可以针对一台计算机进行更改，从而导致数据损坏。为了获得最一致的结果，应用程序应该使用Unicode，如UTF-8或UTF-16，而不是特定的代码页，除非遗留标准或数据格式阻止使用Unicode。如果无法使用Unicode，应用程序应该在协议允许的情况下使用适当的编码名称标记数据流。HTML和XML文件允许标记，但是文本文件不允许。
'@原型
'    int MultiByteToWideChar(
'      _In_      UINT   CodePage,
'      _In_      DWORD  dwFlags,
'      _In_      LPCSTR lpMultiByteStr,
'      _In_      int    cbMultiByte,
'      _Out_opt_ LPWSTR lpWideCharStr,
'      _In_      int    cchWideChar
'    );
'@参数
'    CodePage
'    执行转换时使用的代码页。此参数可以设置为操作系统中已安装或可用的任何代码页的值。有关代码页列表，请参见代码页标识符。您的应用程序还可以指定下表中所示的值之一。
Private Const CP_ACP        As Long = 0
'    系统默认的Windows ANSI代码页?
'    注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
Private Const CP_MACCP      As Long = 1
'    当前系统Macintosh代码页?
'    注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
'    注意，这个值主要用于遗留代码，由于现代Macintosh计算机使用Unicode进行编码，所以通常不需要这个值。
Private Const CP_OEMCP      As Long = 2
'    当前系统OEM代码页?
'    注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
Private Const CP_SYMBOL     As Long = 42
'    符号代码页(42)。
Private Const CP_THREAD_ACP As Long = 3
'    当前线程的Windows ANSI代码页?
'    注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
Private Const CP_UTF7       As Long = 65000
'    utf - 7。只有在7位传输机制强制时才使用此值。最好使用UTF-8。
Private Const CP_UTF8       As Long = 65001
'    utf - 8。
'        037 IBM037  IBM EBCDIC US-Canada
'        437 IBM437  OEM United States
'        500 IBM500  IBM EBCDIC International
'        708 ASMO-708    Arabic (ASMO 708)
'        709     Arabic (ASMO-449+, BCON V4)
'        710     Arabic - Transparent Arabic
'        720 DOS-720 Arabic (Transparent ASMO); Arabic (DOS)
'        737 ibm737  OEM Greek (formerly 437G); Greek (DOS)
'        775 ibm775  OEM Baltic; Baltic (DOS)
'        850 ibm850  OEM Multilingual Latin 1; Western European (DOS)
'        852 ibm852  OEM Latin 2; Central European (DOS)
'        855 IBM855  OEM Cyrillic (primarily Russian)
'        857 ibm857  OEM Turkish; Turkish (DOS)
'        858 IBM00858    OEM Multilingual Latin 1 + Euro symbol
'        860 IBM860  OEM Portuguese; Portuguese (DOS)
'        861 ibm861  OEM Icelandic; Icelandic (DOS)
'        862 DOS-862 OEM Hebrew; Hebrew (DOS)
'        863 IBM863  OEM French Canadian; French Canadian (DOS)
'        864 IBM864  OEM Arabic; Arabic (864)
'        865 IBM865  OEM Nordic; Nordic (DOS)
'        866 cp866   OEM Russian; Cyrillic (DOS)
'        869 ibm869  OEM Modern Greek; Greek, Modern (DOS)
'        870 IBM870  IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
'        874 windows-874 ANSI/OEM Thai (ISO 8859-11); Thai (Windows)
'        875 cp875   IBM EBCDIC Greek Modern
'        932 shift_jis   ANSI/OEM Japanese; Japanese (Shift-JIS)
'        936 gb2312  ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
'        949 ks_c_5601-1987  ANSI/OEM Korean (Unified Hangul Code)
'        950 big5    ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
'        1026    IBM1026 IBM EBCDIC Turkish (Latin 5)
'        1047    IBM01047    IBM EBCDIC Latin 1/Open System
'        1140    IBM01140    IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
'        1141    IBM01141    IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
'        1142    IBM01142    IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
'        1143    IBM01143    IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
'        1144    IBM01144    IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
'        1145    IBM01145    IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
'        1146    IBM01146    IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
'        1147    IBM01147    IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
'        1148    IBM01148    IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
'        1149    IBM01149    IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
'        1200    utf-16  Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
'        1201    unicodeFFFE Unicode UTF-16, big endian byte order; available only to managed applications
'        1250    windows-1250    ANSI Central European; Central European (Windows)
'        1251    windows-1251    ANSI Cyrillic; Cyrillic (Windows)
'        1252    windows-1252    ANSI Latin 1; Western European (Windows)
'        1253    windows-1253    ANSI Greek; Greek (Windows)
'        1254    windows-1254    ANSI Turkish; Turkish (Windows)
'        1255    windows-1255    ANSI Hebrew; Hebrew (Windows)
'        1256    windows-1256    ANSI Arabic; Arabic (Windows)
'        1257    windows-1257    ANSI Baltic; Baltic (Windows)
'        1258    windows-1258    ANSI/OEM Vietnamese; Vietnamese (Windows)
'        1361            Johab Korean(Johab)
'        10000   macintosh   MAC Roman; Western European (Mac)
'        10001   x-mac-japanese  Japanese (Mac)
'        10002   x-mac-chinesetrad   MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
'        10003   x-mac-korean    Korean (Mac)
'        10004   x-mac-arabic    Arabic (Mac)
'        10005   x-mac-hebrew    Hebrew (Mac)
'        10006   x-mac-greek Greek (Mac)
'        10007   x-mac-cyrillic  Cyrillic (Mac)
'        10008   x-mac-chinesesimp   MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
'        10010   x-mac-romanian  Romanian (Mac)
'        10017   x-mac-ukrainian Ukrainian (Mac)
'        10021   x-mac-thai  Thai (Mac)
'        10029   x-mac-ce    MAC Latin 2; Central European (Mac)
'        10079   x-mac-icelandic Icelandic (Mac)
'        10081   x-mac-turkish   Turkish (Mac)
'        10082   x-mac-croatian  Croatian (Mac)
'        12000   utf-32  Unicode UTF-32, little endian byte order; available only to managed applications
'        12001   utf-32BE    Unicode UTF-32, big endian byte order; available only to managed applications
'        20000   x-Chinese_CNS   CNS Taiwan; Chinese Traditional (CNS)
'        20001   x-cp20001   TCA Taiwan
'        20002   x_Chinese-Eten  Eten Taiwan; Chinese Traditional (Eten)
'        20003   x-cp20003   IBM5550 Taiwan
'        20004   x-cp20004   TeleText Taiwan
'        20005   x-cp20005   Wang Taiwan
'        20105   x-IA5   IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
'        20106   x-IA5-German    IA5 German (7-bit)
'        20107   x-IA5-Swedish   IA5 Swedish (7-bit)
'        20108   x-IA5-Norwegian IA5 Norwegian (7-bit)
'        20127   us-ascii    US-ASCII (7-bit)
'        20261   x-cp20261   T.61
'        20269   x-cp20269   ISO 6937 Non-Spacing Accent
'        20273   IBM273  IBM EBCDIC Germany
'        20277   IBM277  IBM EBCDIC Denmark-Norway
'        20278   IBM278  IBM EBCDIC Finland-Sweden
'        20280   IBM280  IBM EBCDIC Italy
'        20284   IBM284  IBM EBCDIC Latin America-Spain
'        20285   IBM285  IBM EBCDIC United Kingdom
'        20290   IBM290  IBM EBCDIC Japanese Katakana Extended
'        20297   IBM297  IBM EBCDIC France
'        20420   IBM420  IBM EBCDIC Arabic
'        20423   IBM423  IBM EBCDIC Greek
'        20424   IBM424  IBM EBCDIC Hebrew
'        20833   x-EBCDIC-KoreanExtended IBM EBCDIC Korean Extended
'        20838   IBM-Thai    IBM EBCDIC Thai
'        20866   koi8-r  Russian (KOI8-R); Cyrillic (KOI8-R)
'        20871   IBM871  IBM EBCDIC Icelandic
'        20880   IBM880  IBM EBCDIC Cyrillic Russian
'        20905   IBM905  IBM EBCDIC Turkish
'        20924   IBM00924    IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
'        20932   EUC-JP  Japanese (JIS 0208-1990 and 0212-1990)
'        20936   x-cp20936   Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
'        20949   x-cp20949   Korean Wansung
'        21025   cp1025  IBM EBCDIC Cyrillic Serbian-Bulgarian
'        21027       (deprecated)
'        21866   koi8-u  Ukrainian (KOI8-U); Cyrillic (KOI8-U)
'        28591   iso-8859-1  ISO 8859-1 Latin 1; Western European (ISO)
'        28592   iso-8859-2  ISO 8859-2 Central European; Central European (ISO)
'        28593   iso-8859-3  ISO 8859-3 Latin 3
'        28594   iso-8859-4  ISO 8859-4 Baltic
'        28595   iso-8859-5  ISO 8859-5 Cyrillic
'        28596   iso-8859-6  ISO 8859-6 Arabic
'        28597   iso-8859-7  ISO 8859-7 Greek
'        28598   iso-8859-8  ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
'        28599   iso-8859-9  ISO 8859-9 Turkish
'        28603   iso-8859-13 ISO 8859-13 Estonian
'        28605   iso-8859-15 ISO 8859-15 Latin 9
'        29001   x-Europa    Europa 3
'        38598   iso-8859-8-i    ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
'        50220   iso-2022-jp ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
'        50221   csISO2022JP ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
'        50222   iso-2022-jp ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
'        50225   iso-2022-kr ISO 2022 Korean
'        50227   x-cp50227   ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
'        50229       ISO 2022 Traditional Chinese
'        50930       EBCDIC Japanese (Katakana) Extended
'        50931               EBCDIC US - Canada And Japanese
'        50933       EBCDIC Korean Extended and Korean
'        50935       EBCDIC Simplified Chinese Extended and Simplified Chinese
'        50936       EBCDIC Simplified Chinese
'        50937       EBCDIC US-Canada and Traditional Chinese
'        50939       EBCDIC Japanese (Latin) Extended and Japanese
'        51932   euc-jp  EUC Japanese
'        51936   EUC-CN  EUC Simplified Chinese; Chinese Simplified (EUC)
'        51949   euc-kr  EUC Korean
'        51950       EUC Traditional Chinese
'        52936   hz-gb-2312  HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
'        54936   GB18030 Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
'        57002   x-iscii-de  ISCII Devanagari
'        57003   x-iscii-be  ISCII Bangla
'        57004   x-iscii-ta  ISCII Tamil
'        57005   x-iscii-te  ISCII Telugu
'        57006   x-iscii-as  ISCII Assamese
'        57007   x-iscii-or  ISCII Odia
'        57008   x-iscii-ka  ISCII Kannada
'        57009   x-iscii-ma  ISCII Malayalam
'        57010   x-iscii-gu  ISCII Gujarati
'        57011   x-iscii-pa  ISCII Punjabi
'        65000   utf-7   Unicode (UTF-7)
'        65001   utf-8   Unicode (UTF-8)
'    dwFlags
'    指示转换类型的标志。应用程序可以指定以下值的组合，默认值为mb_precomposition。预组合和复合是互斥的。可以设置MB_USEGLYPHCHARS和MB_ERR_INVALID_CHARS，而不管其他标志的状态如何。
Private Const MB_COMPOSITE          As Long = &H2
'    始终使用分解的字符，即基本字符和一个或多个非间距字符都具有不同的代码点值的字符。例如，A由A +¨表示:拉丁大写字母A (U+0041) +结合DIAERESIS (U+0308)。注意，此标志不能与mb_precomposition一起使用。
Private Const MB_ERR_INVALID_CHARS As Long = &H8
'    如果遇到无效的输入字符，则失败。
'    从Windows Vista开始，如果应用程序没有设置此标志，该函数不会删除非法代码点，而是用U+FFFD替换非法序列(根据指定的代码页进行适当编码)。
'    如果没有设置此标志，该函数将自动删除非法代码点。对GetLastError的调用返回ERROR_NO_UNICODE_TRANSLATION。
Private Const MB_PRECOMPOSED        As Long = &H1
'    违约;不要与MB_COMPOSITE一起使用。始终使用预组合字符，即对于基字符或非间距字符组合具有单个字符值的字符。例如，在字符e中，e是基本字符，重音符号是无间距字符。如果为一个字符定义了一个Unicode编码点，应用程序应该使用它，而不是单独的基本字符和非间距字符。例如，A由单Unicode编码点拉丁大写字母A和DIAERESIS (U+00C4)表示。
Private Const MB_USEGLYPHCHARS      As Long = &H4
'    使用字形字符而不是控制字符?
'    对于下面列出的代码页，dwFlags必须设置为0。否则，函数将使用ERROR_INVALID_FLAGS失败。
'       50220
'       50221
'       50222
'       50225
'       50227
'       50229
'       57002     through 57011
'       65000 (UTF-7)
'       42 (Symbol)
'    注意:对于UTF-8或代码页54936 (GB18030，从Windows Vista开始)，dwFlags必须设置为0或MB_ERR_INVALID_CHARS。否则，函数将使用ERROR_INVALID_FLAGS失败。
'    lpMultiByteStr [
'       指向要转换的字符串的指针?
'    cbMultiByte
'    lpMultiByteStr参数表示的字符串的大小(以字节为单位)。或者，如果字符串以null结尾，可以将该参数设置为-1。注意，如果cbMultiByte为0，函数将失败。
'    如果该参数为-1，则函数处理整个输入字符串，包括终止null字符。因此，得到的Unicode字符串有一个终止null字符，函数返回的长度包含这个字符。
'    如果将此参数设置为正整数，则函数将精确处理指定的字节数。如果提供的大小不包含终止null字符，则生成的Unicode字符串不是以null结尾的，返回的长度也不包含此字符。
'    lpWideCharStr(,可选)
'    指向接收转换字符串的缓冲区的指针?
'    cchWideChar [在]
'    lpWideCharStr表示的缓冲区的大小(以字符为单位)。如果该值为0，函数将返回所需的缓冲区大小，以字符为单位，包括任何终止null字符，并且不使用lpWideCharStr缓冲区。
'@返回值
'    如果成功，返回写入lpWideCharStr指示的缓冲区的字符数。如果函数成功且cchWideChar为0，则返回值是lpWideCharStr所指示的缓冲区所需的大小(以字符为单位)。有关MB_ERR_INVALID_CHARS标志在输入无效序列时如何影响返回值的信息，请参阅dwFlags。
'    如果函数没有成功，则返回0。要获得扩展的错误信息，应用程序可以调用GetLastError，它可以返回以下错误代码之一:
'    ERROR_INSUFFICIENT_BUFFER。所提供的缓冲区大小不够大，或者被错误地设置为NULL。
'    ERROR_INVALID_FLAGS?为标志提供的值无效?
'    ERROR_INVALID_PARAMETER?任何参数值都无效?
'    ERROR_NO_UNICODE_TRANSLATION?在字符串中发现无效的Unicode?
'@备注
'    此函数的默认行为是转换为输入字符串的预组合形式。如果不存在预组合形式，该函数将尝试转换为组合形式。
'    使用mb_precomposedflag对大多数代码页影响很小，因为大多数输入数据已经被组合好了。考虑在使用MultiByteToWideChar进行转换后调用NormalizeString。NormalizeString提供了更准确、标准和一致的数据，而且速度更快。注意，对于传递给NormalizeString的NORM_FORM枚举，NormalizationC对应于mb_precomposition, NormalizationD对应于MB_COMPOSITE。
'    如上面的警告所述，如果不首先调用此函数，并将cchWideChar设置为0以获得所需的大小，则很容易溢出输出缓冲区。如果使用MB_COMPOSITE标志，每个输入字符的输出长度可以是三个或更多字符。
'    lpMultiByteStr和lpWideCharStr指针不能相同。如果它们是相同的，则函数失败，GetLastError返回值ERROR_INVALID_PARAMETER。
'    如果显式指定输入字符串长度而没有终止空字符，则MultiByteToWideChar不会为空终止输出字符串。若要空终止此函数的输出字符串，应用程序应传入-1或显式计算输入字符串的终止空字符。
'    如果设置了MB_ERR_INVALID_CHARS，并且在源字符串中遇到无效字符，则该函数将失败。无效字符是下列字符之一:
'    不是源字符串中的默认字符，但在未设置MB_ERR_INVALID_CHARS时转换为默认字符的字符
'    对于DBCS字符串，具有前导字节但没有有效跟踪字节的字符
'    从Windows Vista开始，这个函数完全符合Unicode 4.1的UTF-8和UTF-16规范。在早期操作系统中使用的函数编码或解码单独代理程序的一半或不匹配的代理程序对。在早期版本的Windows中编写的依赖于这种行为来编码随机非文本二进制数据的代码可能会遇到问题。但是，在有效的UTF-8字符串上使用该函数的代码的行为将与在早期Windows操作系统上相同。
'    Windows XP:为了防止UTF-8字符的非短格式版本的安全问题，MultiByteToWideChar删除了这些字符。
'    从Windows 8开始:MultiByteToWideChar是用stringapi .h声明的。在Windows 8之前，它是在Winnls.h中声明的。
'@Requirements
'Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'Minimum supported phone    Windows Phone 8
'Header                     Stringapiset.h (include Windows.h)
'Library                    kernel32.lib
'dll                        kernel32.dll
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
'@功能
'    将UTF-16(宽字符)字符串映射到新字符串。新的字符串不一定来自多字节字符集。
'    错误地使用WideCharToMultiByte函数会损害应用程序的安全性。调用这个函数很容易导致缓冲区溢出，因为lpWideCharStr表示的输入缓冲区的大小等于Unicode字符串中的字符数，而lpMultiByteStr表示的输出缓冲区的大小等于字节数。为了避免缓冲区溢出，应用程序必须为缓冲区接收的数据类型指定适当的缓冲区大小。
'    从UTF-16转换为非Unicode编码的数据可能会丢失数据，因为代码页可能无法表示特定Unicode数据中使用的每个字符。有关更多信息，请参见安全性考虑:国际特性。
'    注意，ANSI代码页可以在不同的计算机上不同，也可以针对一台计算机进行更改，从而导致数据损坏。为了获得最一致的结果，应用程序应该使用Unicode，如UTF-8或UTF-16，而不是特定的代码页，除非遗留标准或数据格式阻止使用Unicode。如果无法使用Unicode，应用程序应该在协议允许的情况下使用适当的编码名称标记数据流。HTML和XML文件允许标记，但是文本文件不允许。
'@原型
'    int WideCharToMultiByte(
'      _In_      UINT    CodePage,
'      _In_      DWORD   dwFlags,
'      _In_      LPCWSTR lpWideCharStr,
'      _In_      int     cchWideChar,
'      _Out_opt_ LPSTR   lpMultiByteStr,
'      _In_      int     cbMultiByte,
'      _In_opt_  LPCSTR  lpDefaultChar,
'      _Out_opt_ LPBOOL  lpUsedDefaultChar
'    );
'@参数
'    CodePage
'        执行转换时使用的代码页。此参数可以设置为操作系统中已安装或可用的任何代码页的值。有关代码页列表，请参见代码页标识符。您的应用程序还可以指定下表中所示的值之一。
'        价值意义
'        CP_ACP
'        系统默认的Windows ANSI代码页?
'        注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
'        CP_MACCP
'        当前系统Macintosh代码页?
'        注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
'        注意，这个值主要用于遗留代码，由于现代Macintosh计算机使用Unicode进行编码，所以通常不需要这个值。
'        CP_OEMCP
'        当前系统OEM代码页?
'        注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
'        CP_SYMBOL
'        视窗2000:符号代码页(42)。
'        CP_THREAD_ACP
'        Windows 2000: 当前线程的Windows ANSI代码页?
'        注意，这个值在不同的计算机上是不同的，甚至在相同的网络上也是不同的。它可以在同一台计算机上更改，从而导致存储的数据不可恢复地损坏。此值仅用于临时使用，如果可能，永久存储应该使用UTF-16或UTF-8。
'        CP_UTF7
'        utf - 7。只有在7位传输机制强制时才使用此值。最好使用UTF-8。使用这个值集，lpDefaultChar和lpUsedDefaultChar必须设置为NULL。
'        CP_UTF8
'        utf - 8。使用这个值集，lpDefaultChar和lpUsedDefaultChar必须设置为NULL。
'    dwFlags [在]
'    指示转换类型的标志。应用程序可以指定以下值的组合。当没有设置这些标志时，函数的执行速度会更快。应用程序应该指定WC_NO_BEST_FIT_CHARS和WC_COMPOSITECHECK，并使用特定的值WC_DEFAULTCHAR检索所有可能的转换结果。如果没有提供这三个值，就会丢失一些结果。
Private Const WC_COMPOSITECHECK             As Long = &H200
'    转换组合字符，包括基本字符和非间距字符，每个字符具有不同的字符值。将这些字符转换为预组合字符，预组合字符具有一个用于基非间距字符组合的字符值。例如，在字符e中，e是基本字符，重音符号是无间距字符。
'    注意:Windows通常使用预组合数据表示Unicode字符串，因此没有必要使用WC_COMPOSITECHECK标志。
'    您的应用程序可以将WC_COMPOSITECHECK与以下任何一个标志组合起来，缺省值为WC_SEPCHARS。当Unicode字符串中没有用于基非间距字符组合的预组合映射时，这些标志将决定函数的行为。如果没有提供这些标志，函数的行为就像设置了WC_SEPCHARS标志一样。有关更多信息，请参见备注部分中的WC_COMPOSITECHECK和相关标志。
'    在转换期间使用默认字符替换异常?
'    转换过程中丢弃非间距字符?
Private Const WC_SEPCHARS                   As Long = &H20
'        Default?在转换期间生成单独的字符?
Private Const WC_ERR_INVALID_CHARS          As Long = &H80
'    Windows Vista及以后版本:如果遇到无效输入字符，则失败(返回0并将last-error代码设置为ERROR_NO_UNICODE_TRANSLATION)。您可以通过调用GetLastError检索最后一个错误代码。如果未设置此标志，则函数将使用U+FFFD替换非法序列(根据指定的代码页进行适当编码)，并通过返回转换字符串的长度成功。注意，此标志仅适用于将代码页指定为CP_UTF8或54936时。它不能与其他代码页值一起使用。
Private Const WC_NO_BEST_FIT_CHARS          As Long = &H400
'    翻译任何没有直接翻译成多字节等值的Unicode字符到lpDefaultChar指定的默认字符。换句话说，如果将Unicode转换为多字节并再次转换回Unicode不能生成相同的Unicode字符，则该函数使用默认字符。此标志可以单独使用，也可以与其他已定义的标志组合使用。
'    对于需要验证的字符串，如文件、资源和用户名，应用程序应该始终使用WC_NO_BEST_FIT_CHARS标志。此标志防止函数将字符映射到看起来相似但语义非常不同的字符。在某些情况下，语义变化可能是极端的。例如，在某些代码页中，“∞”(∞)的符号映射到8(8)。
'    对于下面列出的代码页，dwFlags必须设置为0。否则，函数将使用ERROR_INVALID_FLAGS失败。
'       50220
'       50221
'       50222
'       50225
'       50227
'       50229
'       57002     through 57011
'       65000 (UTF-7)
'       42 (Symbol)
'    注意:对于UTF-8或代码页54936 (GB18030，从Windows Vista开始)，dwFlags必须设置为0或MB_ERR_INVALID_CHARS。否则，函数将使用ERROR_INVALID_FLAGS失败。
'    lpWideCharStr [在]
'       指向要转换的Unicode字符串的指针?
'    cchWideChar [在]
'       lpWideCharStr表示的字符串的大小(以字符为单位)。或者，如果字符串以null结尾，可以将该参数设置为-1。如果将cchWideChar设置为0，则函数将失败。
'       如果该参数为-1，则函数处理整个输入字符串，包括终止null字符。因此，得到的字符串有一个终止null字符，函数返回的长度包含这个字符。
'       如果将此参数设置为正整数，则函数将精确处理指定的字符数。如果提供的大小不包含终止null字符，则生成的字符串不以null结尾，返回的长度也不包含此字符。
'       lpMultiByteStr(,可选)
'       指向接收转换字符串的缓冲区的指针?
'    cbMultiByte [在]
'       lpMultiByteStr表示的缓冲区的大小(以字节为单位)。如果将该参数设置为0，该函数将返回lpMultiByteStr所需的缓冲区大小，并且不使用输出参数本身。
'    lpDefaultChar(,可选)
'       如果无法在指定的代码页中表示字符，则指向要使用的字符的指针。如果函数要使用系统默认值，应用程序将该参数设置为NULL。要获得系统默认字符，应用程序可以调用GetCPInfo或GetCPInfoEx函数。
'       对于CodePage的CP_UTF7和CP_UTF8设置，必须将该参数设置为NULL。否则，函数将使用ERROR_INVALID_PARAMETER失败。
'    lpUsedDefaultChar(,可选)
'       指向一个标志的指针，该标志指示函数在转换中是否使用了默认字符。如果源字符串中的一个或多个字符不能在指定的代码页中表示，则将该标志设置为TRUE。否则，将标志设置为FALSE。这个参数可以设置为NULL。
'       对于CodePage的CP_UTF7和CP_UTF8设置，必须将该参数设置为NULL。否则，函数将使用ERROR_INVALID_PARAMETER失败。
'@返回值
'    如果成功，返回lpMultiByteStr指向的写入缓冲区的字节数。如果函数成功且cbMultiByte为0，则返回值为lpMultiByteStr所指示的缓冲区所需的大小(以字节为单位)。有关输入无效序列时WC_ERR_INVALID_CHARS标志如何影响返回值的信息，请参阅dwFlags。
'    如果函数没有成功，则返回0。要获得扩展的错误信息，应用程序可以调用GetLastError，它可以返回以下错误代码之一:
'    ERROR_INSUFFICIENT_BUFFER。所提供的缓冲区大小不够大，或者被错误地设置为NULL。
'    ERROR_INVALID_FLAGS?为标志提供的值无效?
'    ERROR_INVALID_PARAMETER?任何参数值都无效?
'    ERROR_NO_UNICODE_TRANSLATION?在字符串中发现无效的Unicode?
'@备注
'    lpMultiByteStr和lpWideCharStr指针不能相同。如果它们是相同的，则函数失败，GetLastError返回ERROR_INVALID_PARAMETER。
'    WideCharToMultiByte不会为空——如果在没有终止空字符的情况下显式指定输入字符串长度，则终止输出字符串。若要空终止此函数的输出字符串，应用程序应传入-1或显式计算输入字符串的终止空字符。
'    如果cbMultiByte小于cchWideChar，这个函数将cbMultiByte指定的字符数写入lpMultiByteStr指定的缓冲区。但是，如果将CodePage设置为CP_SYMBOL，并且cbMultiByte小于cchWideChar，则该函数不向lpMultiByteStr写入字符。
'    当lpDefaultChar和lpUsedDefaultChar都设置为NULL时，WideCharToMultiByte函数的运行效率最高。下表显示了这些参数的四种可能组合的函数行为。
'    lpDefaultChar lpuseddefaultchar        结果
'    NULL           NULL                    没有默认检查?这些参数设置是与此函数一起使用的最有效的设置?
'    非空字符       null                    使用指定的默认字符，但不设置lpUsedDefaultChar。
'    null           非空字符                使用系统默认字符，并在必要时设置lpUsedDefaultChar。
'    非空字符       非空字符                使用指定的默认字符，并在必要时设置lpUsedDefaultChar。
'@Requirements
'Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'Minimum supported phone    Windows Phone 8
'Header                     Stringapiset.h (include Windows.h)
'Library                    kernel32.lib
'dll                        kernel32.dll

Public Enum LogLevel
    LogLevel_UnDefined = -1                 '尚未设置，应用于模块部件级别
    LogLevel_LogOFF = 0                     '不记录日志
    LogLevel_Error = 1                      '只记录错误
    LogLevel_Warn = 2                       '记录警告
    LogLevel_Info = 3                       '记录重要信息
    LogLevel_Debug = 4                      '记录调试信息
    LogLevel_Trace = 5                      '记录所有的跟踪日志
    LogLevel_AllLog = 6                     '记录所有的详细日志,该级别自动屏蔽部件的内部日志详细设置
End Enum

Public Enum LogType
    LogType_EveryProcess = 0                '每个进程一个
    LogType_SingleInstance = 1              '只有一份日志，进程是单实例运行
End Enum


Private Const M_SYS_NAME                                As String = "中联软件"
Private Const M_MAX_LINE_COUNT_INIT                     As Long = 1024           '日志初始长度以及
Private Const M_CACHELOG_MAX_COUNT                      As Long = 100           '日志缓存的最大值
Private Const M_MAX_STACK_LENGTH                        As Long = 40            '调用堆栈的长度
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mstrLogLevelName(7)                             As String               '日志级别名称
Private mobjLog                                         As TextStream
Private mobjFSO                                         As New FileSystemObject
Private mlngStart                                       As Long                 '日志生成时间
Private mdtStart                                        As Date                 '日志生成时间
'日志跟踪

Private mcllMethodStack                                 As New Collection       '调用堆栈集合


Private mstrCacheLog(M_CACHELOG_MAX_COUNT)              As String               '缓存的日志行
Private mbytCacheLogLevel(M_CACHELOG_MAX_COUNT)         As Byte                 '缓存的当前日志等级
Private mstrLogModule(M_CACHELOG_MAX_COUNT)             As String               '日志对应的模块
Private mlngCacheCount                                  As Long                 '缓存的日志行数

Private mcllAllLogLevel                                 As New Collection       '所有部件与模块的日志级别。
'部件级日志级别处理临时变量
Private mstrLastComponentModule                         As String
Private mllLastLogLevel                                 As LogLevel             '临时日志级别
Private mstrLastMethod                                  As String
Private marrTmpComponentModule                          As Variant

'临时变量
Private mlngCur                                         As Long                 '当前时间
Private mstrLogTiltle                                   As String
Private mstrTmp                                         As String
Private mlngIndex                                       As Long
Private mblnDoing                                       As Boolean


'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
'IsAppend是否添加方式(R)
Private mblnAppend                      As Boolean
'CurrentLogLevel当前日志级别(R/W)
Private mllCurrentLogLevel              As LogLevel
'LogFile日志文件文件路径(R)
Private mstrLogFile                     As String
'MaxLogCount日志文件文件最大行数(R/W)
Private mlngMaxLogCount                 As Long
'CurrentLogCount当前已经记录的行数(R)
Private mlngCurLogCount                 As Long
'IsLoopTimeFormat是否是轮训时间(R/W)
Private mblnLoopTimeFormat              As Boolean
'IsErrMessage 是否弹出错误消息提示R/W)
Private mblnErrMsg                      As Boolean
'IsLogOpened 设置日志状态，不需要日志设置为False，加快速度(R/W)
Private mblnLogOpened                   As Boolean
'LogType,日志生成类型
Private mltCurrent                      As LogType
'LogCacheSize  设置缓冲区大小（R/W）
'Private mlngLogCacheSize                As Long
'IsAppend日志是否是追加方式R/W)
Public Property Get IsAppend() As Boolean
    IsAppend = mblnAppend
End Property
'CurrentLogLevel当前日志级别(R/W)
Public Property Get CurrentLogLevel() As LogLevel
    CurrentLogLevel = mllCurrentLogLevel
End Property

Public Property Let CurrentLogLevel(ByVal llNewLogLevel As LogLevel)
    Dim blnAppend       As Boolean
    
    If llNewLogLevel < LogLevel_UnDefined Then
        llNewLogLevel = LogLevel_UnDefined
    ElseIf llNewLogLevel > LogLevel_AllLog Then
        llNewLogLevel = LogLevel_AllLog
    End If
    If mllCurrentLogLevel <> llNewLogLevel Then
        If llNewLogLevel = LogLevel_LogOFF Then
            Call CloseEx
            mllCurrentLogLevel = llNewLogLevel
        ElseIf mllCurrentLogLevel = LogLevel_LogOFF And llNewLogLevel > LogLevel_LogOFF Then
            blnAppend = mblnAppend
            Call OpenEx(mstrLogFile, , , True, mlngMaxLogCount, llNewLogLevel, mblnLoopTimeFormat, mblnErrMsg)
            mblnAppend = blnAppend
        Else
            mllCurrentLogLevel = llNewLogLevel
        End If
    End If
End Property
'IsLoopTimeFormat是否是轮训时间(R/W)
Public Property Get IsLoopTimeFormat() As Boolean
    IsLoopTimeFormat = mblnLoopTimeFormat
End Property

Public Property Let IsLoopTimeFormat(ByVal blnLoopTimeFormat As Boolean)
    mblnLoopTimeFormat = blnLoopTimeFormat
End Property
'LogFile日志文件文件名称(R)
Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'MaxLogCount日志文件文件最大行数(R/W)
Public Property Get MaxLogCount() As Long
    MaxLogCount = mlngMaxLogCount
End Property

Public Property Let MaxLogCount(ByVal lngMaxLogCount As Long)
    mlngMaxLogCount = lngMaxLogCount
    If mlngMaxLogCount < M_MAX_LINE_COUNT_INIT Then
        mlngMaxLogCount = M_MAX_LINE_COUNT_INIT
    End If
End Property
'CurrentLogCount当前已经记录的行数(R)
Public Property Get CurrentLogCount() As Long
    CurrentLogCount = mlngCurLogCount
End Property
'IsNoErrMessage 是否不弹出错误消息提示(R/W)
Public Property Get IsErrMessage() As Boolean
    IsErrMessage = mblnErrMsg
End Property

Public Property Let IsErrMessage(ByVal blnErrMsg As Boolean)
    mblnErrMsg = blnErrMsg
End Property

'IsLogOpened 设置日志状态，不需要日志设置为False，加快速度(R/W)
Public Property Get IsLogOpened() As Boolean
    IsLogOpened = mblnLogOpened
End Property

Public Property Let IsLogOpened(ByVal blnLogOpened As Boolean)
    If mblnLogOpened <> blnLogOpened Then
        If Not blnLogOpened Then
            Call CloseEx
        End If
    End If
    mblnLogOpened = blnLogOpened
End Property
'LogType,日志生成类型
Public Property Get LogType() As LogType
    LogType = mltCurrent
End Property
'LogCacheSize  设置日志缓存大小（R/W）,最小128行(R/W)
'Public Property Get LogCacheSize() As Long
'    LogCacheSize = mlngLogCacheSize
'End Property
'
'Public Property Let LogCacheSize(ByVal lngLogCacheSize As Long)
'    mlngLogCacheSize = lngLogCacheSize
'    If mlngMaxLogCount < M_MAX_LINE_COUNT_INIT \ 4 Then
'        mlngMaxLogCount = M_MAX_LINE_COUNT_INIT \ 4
'    End If
'End Property

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    OpenEx
'   打开日志。
'@返回值  Boolean
'   是否成功打开日志
'@参数:
'strLogName     String(IN)
'   日志名称，可以使全路径或者只有名称，生成的文件名为，strLogName_strLogSubFix.log
'strLogSubFix String In
'   日志名的后缀
'blnAppend      Boolean(In,opt)
'   是否是追加方式写日志
'lngMaxLogCount Long(In,opt,default= 2 ^ 20)
'   日志行数上线
'llLogLevel     LogLevel(In,Opt,default=LogLevel_LogOFF)
'   准备记录的日志格式
'blnLoopTimeFormat Boolean(In,opt)
'   日志头部时间格式是否是轮训时间
'blnErrMsg        Boolean(In,opt,default=True)
'   ErrorCenter弹出消息提示
'blnAutoCreate Boolean(In,opt,default=True)
'   是否自动生成日志文件，否则只缓存，等待指定文件名才写入。
'lngLogCache Long(In,opt,default=512)
'   设置日志缓存大小
'@备注
'   日志总是产生的，不管是否开启，但是是否记录日志根据日志级别控制。生成的日志文件名格式：Appsoft\Log\日志跟踪\进程文件名_SessionID_当前进程的操作系统会话ID_strLogSubFix.log
Public Function OpenEx(Optional ByVal strLogName As String, Optional ByVal strLogSubFix As String, Optional ByVal ltLogType As LogType = LogType_EveryProcess, Optional ByVal blnAppend As Boolean, Optional ByVal lngMaxLogCount As Long = M_MAX_LINE_COUNT_INIT, Optional ByVal llLogLevel As LogLevel = LogLevel_LogOFF, Optional ByVal blnLoopTimeFormat As Boolean, Optional ByVal blnErrMsg As Boolean) As Boolean
    Dim strLogFile      As String
    Dim arrTmp          As Variant, i           As Long
    Dim strTmp          As String
    Dim llLogLevelTmp   As LogLevel
    On Error GoTo ErrH
    '全路径模式
    If InStr(strLogName, "\") > 0 Then
        strLogFile = strLogName
    Else
        If Len(strLogName) <> 0 Then
            strLogFile = AppsoftPath & "\Log\日志跟踪\" & strLogName
        Else
            If Environment.IsDesinMode Then
                strLogFile = AppsoftPath & "\Log\日志跟踪\" & App.EXEName & "_SessionID_" & Environment.SessionID
            Else
                strLogFile = AppsoftPath & "\Log\日志跟踪\" & mobjFSO.GetBaseName(Environment.StartExePath) & "_SessionID_" & Environment.SessionID
            End If
        End If
        
        If ltLogType = LogType_EveryProcess Then
            strLogFile = strLogFile & "_PID_" & GetCurrentProcessId
        End If
        If LenB(strLogSubFix) <> 0 Then
            strLogFile = strLogFile & "_" & strLogSubFix & ".log"
        Else
            strLogFile = strLogFile & ".log"
        End If
    End If
    If mobjFSO.FolderExists(mobjFSO.GetParentFolderName(strLogFile)) Then
        arrTmp = Split(strLogFile, "\")
        strTmp = arrTmp(0) & "\"
        For i = 1 To UBound(arrTmp) - 1
            strTmp = strTmp & arrTmp(i) & "\"
            If Not mobjFSO.FolderExists(strTmp) Then
                Call mobjFSO.CreateFolder(strTmp)
            End If
        Next
    End If
    If Not mobjLog Is Nothing Then
        Call CloseEx_Sub(True)
    End If
    If LenB(strLogFile) <> 0 Then
        If blnAppend Then
            Set mobjLog = mobjFSO.OpenTextFile(strLogFile, ForAppending, True)
        Else
            Set mobjLog = mobjFSO.OpenTextFile(strLogFile, ForWriting, True)
        End If
        mblnLoopTimeFormat = blnLoopTimeFormat
        mstrLogFile = strLogFile
        mblnAppend = blnAppend
        mllCurrentLogLevel = llLogLevel
        mlngMaxLogCount = lngMaxLogCount
        mblnErrMsg = blnErrMsg
        mltCurrent = ltLogType
        Call FlushLogCache
        mblnLogOpened = True
        OpenEx = True
    End If
    Exit Function
ErrH:
    Err.Clear
    OpenEx = False
End Function
'@方法    CloseEx
'   关闭日志
'@返回值  Boolean
Public Function CloseEx() As Boolean
    CloseEx = CloseEx_Sub
End Function

'@方法    AddComponentLogLevel
'   添加部件模块的日志级别,只添加，已经存在返回False
'@返回值
'@参数:
'strComponentModule      String(In,opt)
'   部件名称.模块名称，或部件名称
'llLogLevel     LogLevel(In，opt,defualt=LogLevel_Warn)
'   部件模块采取的日志的级别,当部件为空时为模块的日志级别
Public Function AddComponentLogLevel(strComponentModule As String, Optional ByVal llLogLevel As LogLevel = LogLevel_Warn) As Boolean
    If Not IsInCollection(mcllAllLogLevel, strComponentModule) Then
        mcllAllLogLevel.Add llLogLevel, strComponentModule
        AddComponentLogLevel = True
    End If
End Function
'@方法    SetModuleLogLevel
'   添加部件模块的日志级别,不存在添加，存在则变更
'@返回值
'@参数:
'strComponentModule      String(In,opt)
'   部件名称.模块名称，或部件名称
'llLogLevel     LogLevel(In，opt,defualt=LogLevel_Warn)
'   部件模块采取的日志的级别,当部件为空时为模块的日志级别
Public Sub SetComponentLogLevel(strComponentModule As String, Optional ByVal llLogLevel As LogLevel = LogLevel_Warn)
    If IsInCollection(mcllAllLogLevel, strComponentModule) Then
        mcllAllLogLevel.Remove strComponentModule
    End If
    mcllAllLogLevel.Add llLogLevel, strComponentModule
End Sub
'@方法    GetComponentLogLevel
'   返回部件模块的日志级别
'@返回值    LogLevel
'@参数:
'strComponentModule      String(In,opt)
'   部件名称.模块名称.方法名称，或部件名称.模块名称，或部件名称
Public Function GetComponentLogLevel(ByVal strComponentModule As String) As LogLevel
    GetComponentLogLevel = GetComponentLogLevelEx(strComponentModule, True)
End Function
'@方法    PushMethod
'   将调用方法推入堆栈
'@返回值
'
'@参数:
'strMethod      String(In)
'   部件名称.模块名称.方法名称
'arrPars        ParamArray
'   方法的参数列表
'@备注
'   该方法必须和PopMethod以及ErrorCenter匹配调用
Public Sub PushMethod(strMethod As String, ParamArray arrPars() As Variant)
    Dim strText     As String
    mllLastLogLevel = GetComponentLogLevelEx(strMethod)
    If (mllLastLogLevel >= LogLevel_Trace Or mobjLog Is Nothing) And mblnLogOpened Then
        strText = ""
        For mlngIndex = LBound(arrPars) To UBound(arrPars)
            strText = strText & "," & DisPlayOneValue(arrPars(mlngIndex), mllLastLogLevel >= LogLevel_AllLog)
        Next
        strText = Mid(strText, 2)
    End If
    With mcllMethodStack
        If .Count = 0 Then
            If Len(strText) = 0 Then
                .Add strMethod
            Else
                .Add strMethod & "(" & strText & ")"
            End If
        Else
            If Len(strText) = 0 Then
                .Add strMethod, , 1
            Else
                .Add strMethod & "(" & strText & ")", , 1
            End If
        End If
        If .Count > M_MAX_STACK_LENGTH Then .Remove .Count
    End With
    If (mllLastLogLevel >= LogLevel_Trace Or mobjLog Is Nothing) And mblnLogOpened Then
        Call GetLogTiltle(LogLevel_Trace)
        If Len(strText) = 0 Then
            WriteLine mstrLogTiltle & "┏" & strMethod, strMethod, LogLevel_Trace
        Else
            WriteLine mstrLogTiltle & "┏" & strMethod & "(" & strText & ")", strMethod, LogLevel_Trace
        End If
    End If
    mstrLastMethod = strMethod
End Sub
'@方法    ErrCenter
'   错误处理中心
'@返回值  Integer
'   0-忽略继续执行，1-重试(Resume),2-中止程序
'@参数:
'strMethod String In
'   错误发生的过程
'@备注
'
Public Function ErrCenter(strMethod As String) As Integer
    Dim strText     As String
    mllLastLogLevel = GetComponentLogLevelEx(strMethod)
    If (mllLastLogLevel >= LogLevel_Error Or mobjLog Is Nothing) And mblnLogOpened Then
        Call GetLogTiltle(LogLevel_Error)
        strText = mstrLogTiltle & "┣" & strMethod & "  " & Err.Number & "-" & Err.Description
        strText = strText & vbNewLine & String(Len(mstrLogTiltle), " ") & "┣" & "调用堆栈："
        For mlngIndex = 1 To mcllMethodStack.Count
            strText = strText & vbNewLine & String(Len(mstrLogTiltle), " ") & "┣" & mcllMethodStack(mlngIndex)
        Next
        If mblnErrMsg Then
            Select Case MsgBox(strText & vbNewLine & "。是否继续？", vbInformation + vbAbortRetryIgnore, M_SYS_NAME)
                Case vbAbort
                    ErrCenter = 2
                Case vbRetry
                    ErrCenter = 1
                Case vbIgnore
                    Err.Clear
                    ErrCenter = 0
            End Select
        Else
            Err.Clear
        End If
        WriteLine strText, strMethod, LogLevel_Error
    Else
        Err.Clear
    End If
End Function
'@方法    PopMethod
'   将最近的入栈的方法移除，或者将指定方法之前入堆栈的方法移除（包含指定的方法）
'@返回值
'
'@参数:
'strMethod String In
'   方法名称，不传时弹出最近入堆栈的方法,部件.模块.方法名
'@备注
'
Public Sub PopMethod(strMethod As String, ParamArray arrPars() As Variant)
    Dim strText     As String
    mllLastLogLevel = GetComponentLogLevelEx(strMethod)
    If (mllLastLogLevel >= LogLevel_Trace Or mobjLog Is Nothing) And mblnLogOpened Then
        strText = ""
        For mlngIndex = LBound(arrPars) To UBound(arrPars)
            strText = strText & "," & DisPlayOneValue(arrPars(mlngIndex), mllLastLogLevel >= LogLevel_AllLog)
        Next
        strText = Mid(strText, 2)
        Call GetLogTiltle(LogLevel_Trace)
        If Len(strText) = 0 Then
            WriteLine mstrLogTiltle & "┗" & strMethod, strMethod, LogLevel_Trace
        Else
            WriteLine mstrLogTiltle & "┗" & strMethod & "(" & strText & ")", strMethod, LogLevel_Trace
        End If
    End If
    With mcllMethodStack
        If Len(strMethod) <> 0 Then
            For mlngIndex = 1 To .Count
                If mcllMethodStack(mlngIndex) Like strMethod & "*" Then
                    Exit For
                End If
            Next
            If mlngIndex > .Count Then
                If .Count > 0 Then  '没有找到任何匹配，则删除一个即可
                    mlngIndex = 1
                Else                '没有数据则不删除
                    mlngIndex = 0
                End If
            End If
        Else
            mlngIndex = 1  '传空则只删除一个
        End If
        
        Do While mlngIndex > 0
            .Remove 1
            mlngIndex = mlngIndex - 1
        Loop
        If .Count > 0 Then
            mstrLastMethod = .Item(1)
        Else
            mstrLastMethod = ""
        End If
        mlngIndex = 1
    End With
End Sub
'@方法    Log
'   记录日志信息，可以指定日志级别
'@返回值    Boolean
'
'@参数:
'llLogLevel     LogLevel(In)
'   当前日志记录内容的日志级别
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function Log(ByVal llLogLevel As LogLevel, strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    Log = Log_Sub(llLogLevel, strInfo, arrParsInfo)
End Function
'@方法    Trace
'   记录跟踪日志
'@返回值    Boolean
'
'@参数:
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function Trace(strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    Trace = Log_Sub(LogLevel_Trace, strInfo, arrParsInfo)
End Function
'@方法    DebugEx
'   记录调试日志
'@返回值    Boolean
'
'@参数:
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function DebugEx(strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    DebugEx = Log_Sub(LogLevel_Debug, strInfo, arrParsInfo)
End Function
'@方法    Info
'   记录信息日志
'@返回值    Boolean
'
'@参数:
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function Info(strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    Info = Log_Sub(LogLevel_Info, strInfo, arrParsInfo)
End Function
'@方法    Warn
'   记录警告信息日志
'@返回值    Boolean
'
'@参数:
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function Warn(strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    Warn = Log_Sub(LogLevel_Warn, strInfo, arrParsInfo)
End Function
'@方法    Error
'   记录错误日志
'@返回值    Boolean
'
'@参数:
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Public Function Error(strInfo As String, ParamArray arrPars() As Variant) As Boolean
    Dim arrParsInfo() As Variant
    arrParsInfo = arrPars
    Error = Log_Sub(LogLevel_Error, strInfo, arrParsInfo)
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'@方法    FlushLogCache
'   将缓存的日志刷新的到日志
'@返回值
'
'@参数:
'@备注
'
Private Sub FlushLogCache()
    Dim llLogLevelTmp       As LogLevel
    '将缓存的日志写入日志文件
    For mlngIndex = 1 To mlngCacheCount
        llLogLevelTmp = GetComponentLogLevelEx(mstrLogModule(mlngIndex))
        If llLogLevelTmp >= mbytCacheLogLevel(mlngIndex) Then
            WriteLine mstrCacheLog(mlngIndex)
        End If
    Next
    Erase mstrLogModule
    Erase mbytCacheLogLevel
    Erase mstrCacheLog
    mlngCacheCount = 0
End Sub
'@方法    CloseEx_Sub
'   关闭日志CloseEx的字方法
'@返回值  Boolean
'@参数
'blnInsideCall Boolean In (Optional)
'   是否内部调用
Private Function CloseEx_Sub(Optional ByVal blnInsideCall As Boolean) As Boolean
    On Error Resume Next
    Call FlushLogCache
    If Not blnInsideCall Then
        mblnLogOpened = False
    End If
    If Not mobjLog Is Nothing Then
        mobjLog.Close
        Set mobjLog = Nothing
    End If
    If Err.Number <> 0 Then Err.Clear
    CloseEx_Sub = True
End Function
'@方法    GetComponentLogLevelEx
'   返回部件模块的日志级别
'@返回值    LogLevel
'@参数:
'strComponentModule      String(In,opt)
'   部件名称.模块名称.方法名称，或部件名称.模块名称，或部件名称
'blnOutsideCall      Boolean In(Opt)
'   是否外部调用，外部调用不需要判断日志是否存在
Private Function GetComponentLogLevelEx(ByVal strComponentModule As String, Optional ByVal blnOutsideCall As Boolean) As LogLevel
    If Not mobjLog Is Nothing Or blnOutsideCall Then
        If LenB(strComponentModule) = 0 Then
            GetComponentLogLevelEx = mllCurrentLogLevel
        Else
            '最高日志级别时，自动屏蔽部件模块级设置
            If mllCurrentLogLevel >= LogLevel_AllLog Then
                GetComponentLogLevelEx = LogLevel_AllLog
            Else
                GetComponentLogLevelEx = LogLevel_UnDefined
                marrTmpComponentModule = Split(strComponentModule & ".", ".")
                mstrTmp = marrTmpComponentModule(0) & "." & marrTmpComponentModule(1)
                If mstrTmp = mstrLastComponentModule Then
                    GetComponentLogLevelEx = mllLastLogLevel
                Else
                    
                    If UBound(marrTmpComponentModule) > 1 Then
                        '直接获取模块的
                        mllLastLogLevel = GetComponentLogLevel_Sub(mstrTmp)
                        If mllLastLogLevel = LogLevel_UnDefined Then
                            '模块没有则判断部件
                            mllLastLogLevel = GetComponentLogLevel_Sub(marrTmpComponentModule(0) & "")
                        End If
                    Else
                        '直接获取部件的
                        mllLastLogLevel = GetComponentLogLevel_Sub(marrTmpComponentModule(0) & "")
                    End If
                    If mllLastLogLevel = LogLevel_UnDefined Then
                        mllLastLogLevel = mllCurrentLogLevel
                    End If
                    mstrLastComponentModule = mstrTmp
                    GetComponentLogLevelEx = mllLastLogLevel
                End If
            End If
        End If
    End If
End Function
'@方法    GetComponentLogLevel_Sub
'   返回部件或部件模块的日志级别，仅供内部调用,没有日志级别时，返回LogLevel_UnDefined
'@返回值    LogLevel
'@参数:
'strComponentModule      String(In,opt)
'   部件名称.模块名称，或部件名称
Private Function GetComponentLogLevel_Sub(strComponentModule As String) As LogLevel
    On Error GoTo ErrH
    GetComponentLogLevel_Sub = LogLevel_UnDefined
    GetComponentLogLevel_Sub = mcllAllLogLevel.Item(strComponentModule)
    Exit Function
ErrH:
    Err.Clear
End Function
'@方法    Log_Sub
'   记录日志信息，可以指定日志级别
'@返回值    Boolean
'
'@参数:
'llLogLevel     LogLevel(In)
'   当前日志记录内容的日志级别
'strInfo        String(In)
'   日志内容名称
'arrPars        ParamArray
'   日志记录项,项目名称arrPars(0)=项目值arrPars(1)
'@备注
'
Private Function Log_Sub(ByVal llLogLevel As LogLevel, strInfo As String, arrPars() As Variant) As Boolean
    Dim strText         As String
    On Error GoTo ErrH
    mllLastLogLevel = GetComponentLogLevelEx(mstrLastMethod)
    If llLogLevel < LogLevel_UnDefined Then
        llLogLevel = LogLevel_UnDefined
    ElseIf llLogLevel > LogLevel_AllLog Then
        llLogLevel = LogLevel_AllLog
    End If
    If (mllLastLogLevel >= llLogLevel Or mobjLog Is Nothing) And mblnLogOpened Then
        strText = ""
        Call GetLogTiltle(llLogLevel)
        For mlngIndex = LBound(arrPars) To UBound(arrPars) Step 2
            If mlngIndex = UBound(arrPars) Then
                strText = strText & arrPars(mlngIndex)
            Else
                strText = strText & arrPars(mlngIndex) & "=" & arrPars(mlngIndex + 1) & "    "
            End If
        Next
        If Len(strText) <> 0 Then
            strText = strInfo & ":  " & strText
        Else
            strText = strInfo
        End If
        If mllLastLogLevel < LogLevel_Trace Then
            strText = mstrLastMethod & "-" & strText
        End If
        WriteLine mstrLogTiltle & "┣" & strText, mstrLastMethod, llLogLevel
    End If
    Log_Sub = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    If (mllLastLogLevel >= LogLevel_Error Or mobjLog Is Nothing) And mblnLogOpened Then
        WriteLine mstrLogTiltle & "┣" & "LogInfo发生错误：" & Err.Number & "-" & Err.Description, "", LogLevel_Error
    End If
End Function

'@方法    GetLogTiltle
'   获取日志头，并进行行控制
'@返回值
Private Sub GetLogTiltle(ByVal llLogLevel As LogLevel)
    Const OFFSET_4          As Double = 4294967296#         '无符号整形的最大值
    Const M_DAY_MS          As Double = 86400000            '一天的毫秒数
    Const M_HOUR_MS         As Double = 3600000             '一小时的毫秒数
    Const M_MIN_MS          As Double = 60000               '一分钟的毫秒数
    Dim dblTmp              As Double
    Dim lngTmp              As Long
    
    If mlngCurLogCount >= mlngMaxLogCount Or mlngCurLogCount = 0 Then
        Call LogWriteHeader
    End If
    mlngCurLogCount = mlngCurLogCount + 1
    If Not mblnLoopTimeFormat Then
        mstrLogTiltle = mstrLogLevelName(llLogLevel + 1) & Format(Now, "YYYY-MM-DD hh:mm:ss") & GetLogDistance
    Else
        mlngCur = GetTickCount
        If mlngCur < mlngStart Then
            dblTmp = OFFSET_4 - LongToUnsigned(mlngStart) + LongToUnsigned(mlngCur)
        Else
            dblTmp = mlngCur - mlngStart
        End If
        mstrLogTiltle = ""
        '天数计算
        dblTmp = dblTmp / M_DAY_MS
        lngTmp = Int(dblTmp)
        mstrLogTiltle = Format(lngTmp, "00")
        '小时计算
        dblTmp = (dblTmp - lngTmp) * M_DAY_MS / M_HOUR_MS
        lngTmp = Int(dblTmp)
        mstrLogTiltle = mstrLogTiltle & " " & Format(lngTmp, "00")
        '分钟计算
        dblTmp = (dblTmp - lngTmp) * M_HOUR_MS / M_MIN_MS
        lngTmp = Int(dblTmp)
        mstrLogTiltle = mstrLogTiltle & ":" & Format(lngTmp, "00")
        '秒计算
        dblTmp = (dblTmp - lngTmp) * M_MIN_MS / 1000
        lngTmp = Int(dblTmp)
        mstrLogTiltle = mstrLogLevelName(llLogLevel + 1) & mstrLogTiltle & ":" & Format(lngTmp, "00") & "." & Format(Int((dblTmp - lngTmp) * 1000), "000") & GetLogDistance
    End If
End Sub

'@方法    LogWriteHeader
'   书写日志头
'@返回值  Boolean
Private Sub LogWriteHeader()
    If mblnDoing Then Exit Sub
    mblnDoing = True
    If mlngCurLogCount <> 0 Then
        If Not mobjLog Is Nothing Then
            Call CloseEx_Sub(True)
            Call mobjFSO.DeleteFile(mstrLogFile, True)
            Call OpenEx(mstrLogFile, , , mblnAppend, mlngMaxLogCount, mllCurrentLogLevel, mblnLoopTimeFormat, mblnErrMsg)
        Else
            Erase mstrLogModule
            Erase mbytCacheLogLevel
            Erase mstrCacheLog
            mlngCacheCount = 0
        End If
        mlngCurLogCount = 0
    Else
        mlngCurLogCount = mlngCurLogCount + 1
    End If
    WriteLine String(80, "-"), "", LogLevel_Error
    WriteLine Format(Now, "yyyy-mm-dd  hh:mm:ss") & " 进程ID=" & GetCurrentProcessId() & " ProcessUser=" & Environment.ProcessUserFullName & " SessionUser=" & Environment.SessionUserFullName, "", LogLevel_Error
    WriteLine String(80, "-"), "", LogLevel_Error
    mblnDoing = False
End Sub
'@方法    GetLogDistance
'   获取日志距离时间标志的长度
'@返回值  String
Private Function GetLogDistance() As String
    If mcllMethodStack.Count = 0 Then
        GetLogDistance = " "
    Else
        GetLogDistance = String((mcllMethodStack.Count - 1) * 2, " ")
    End If
End Function

'@方法    LoadLogLevelSetting
'   从注册表加载日志配置
'@返回值
'
'@参数:
'@备注
'
Private Sub LoadLogLevelSetting()
    Dim arrSubKey       As Variant
    Dim strParentKey    As String
    Dim i               As Long, j          As Long
    Dim varNames        As Variant
    Dim varValues       As Variant
    Dim strZLSOFTRegKey As String
    Dim strRet          As String
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsLogger.LoadLogLevelSetting")
    If LenB(Environment.SessionUserSID) <> 0 Then
        If Environment.Is64bitOS Then
            strZLSOFTRegKey = "HKEY_USERS\" & Environment.SessionUserSID & "\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT"
        Else
            strZLSOFTRegKey = "HKEY_USERS\" & Environment.SessionUserSID & "\Software\VB and VBA Program Settings\ZLSOFT"
        End If
    End If
    If LenB(strZLSOFTRegKey) <> 0 Then
        '读取指定进程的日志级别
        strRet = Registry.GetRegValue(strZLSOFTRegKey & "\公共模块\日志跟踪\进程", gobjFSO.GetFileName(Environment.StartExePath), LogLevel_Trace)
        If strRet = "" Then
            mllCurrentLogLevel = LogLevel_Trace
        Else
            mllCurrentLogLevel = Val(strRet)
        End If
        mblnLoopTimeFormat = Val(Registry.GetRegValue(strZLSOFTRegKey & "\公共模块\日志跟踪", "时间格式", "0")) = 1

        arrSubKey = Registry.GetRegSubKeys(strZLSOFTRegKey & "\公共模块\日志跟踪\部件\")
        If TypeName(arrSubKey) <> "Empty" Then
            For i = LBound(arrSubKey) To UBound(arrSubKey)
                If Registry.GetRegKeyValues(strZLSOFTRegKey & "\公共模块\日志跟踪\部件\" & arrSubKey(i), varNames, varValues) Then
                    If TypeName(varNames) <> "Empty" Then
                        For j = LBound(varNames) To UBound(varNames)
                            If varNames(j) = "" Then
                                Logger.SetComponentLogLevel arrSubKey(i) & "", Val(varValues(j))
                                Logger.DebugEx arrSubKey(i) & "", Val(varValues(j))
                            Else
                                Logger.SetComponentLogLevel arrSubKey(i) & "." & varNames(j), Val(varValues(j))
                                Logger.DebugEx arrSubKey(i) & "." & varNames(j), Val(varValues(j))
                            End If
                        Next
                    End If
                End If
            Next
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsLogger.LoadLogLevelSetting")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsLogger.LoadLogLevelSetting") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsLogger.LoadLogLevelSetting")
End Sub

'@方法    WriteLine
'   TextStream的方法封装，支持缓存
'@返回值
'
'@参数:
'strLine  String    In
'   书写的内容
'strModule  String  In
'   内容对应的方法
'llLogLevel LogLevel    In
'   书写内容的日志级别
'@备注
'
Private Sub WriteLine(strLine As String, Optional strModule As String, Optional llLogLevel As LogLevel)
    If Not mobjLog Is Nothing Then
        mobjLog.WriteLine strLine
    Else
        mlngCacheCount = mlngCacheCount + 1
        If mlngCacheCount > M_CACHELOG_MAX_COUNT Then
            mlngCacheCount = 1
        End If
        mstrCacheLog(mlngCacheCount) = strLine
        mbytCacheLogLevel(mlngCacheCount) = llLogLevel
        mstrLogModule(mlngCacheCount) = strModule
    End If
End Sub
'@方法    DisPlayOneValue
'   展示对象或值
'@返回值  String
'
'@参数:
'valValue  Variant(In)
'   转化为字符串的值
'blnSerializeObject Boolean(In,opt,defualt=True)
'@备注
'   对象类型可能不一定支持序列化
Private Function DisPlayOneValue(valValue As Variant, Optional ByVal blnSerializeObject As Boolean = True) As String
    Dim strTmp  As String
    
    If IsArray(valValue) Then
        Dim i    As Long
        strTmp = "["
        For i = LboundEx(valValue) To UboundEx(valValue)
            strTmp = strTmp & DisPlayOneValue(valValue(i), blnSerializeObject) & ","
        Next
        If Len(strTmp) = 1 Then
            strTmp = TypeName(valValue) & strTmp & "]"
        Else
            strTmp = TypeName(valValue) & Mid(strTmp, 1, Len(strTmp) - 1) & "]"
        End If
    ElseIf IsNull(valValue) Then
        strTmp = "{NULL}"
    ElseIf IsEmpty(valValue) Then
        strTmp = "{EMPTY}"
    ElseIf IsObject(valValue) Then
        If valValue Is Nothing Then
            strTmp = "{NOTHING}"
        Else
            If blnSerializeObject Then
                strTmp = "{OBJECT(" + TypeName(valValue) + ")=" & Serialize(valValue) & "}"
            Else
                strTmp = "{OBJECT(" + TypeName(valValue) + ")}"
            End If
        End If
    Else
        If VarType(valValue) = vbString Then
            strTmp = """" & valValue & """"
        Else
            strTmp = CStr(valValue)
        End If
    End If
    DisPlayOneValue = strTmp
End Function

'@方法    UboundEx
'   获取  Ubound
'@返回值  Long
'
'@参数:
'varArray Variant In
'   传入的数组
'@备注
'
Private Function UboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    UboundEx = UBound(varArray)
    Exit Function
ErrH:
    UboundEx = -1
End Function

'@方法    LboundEx
'   获取  Lbound
'@返回值  Long
'
'@参数:
'varArray Variant In
'   传入的数组
'@备注
'
Private Function LboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    LboundEx = LBound(varArray)
    Exit Function
ErrH:
    LboundEx = 0
End Function

'@方法    AppsoftPath
'   获取APPSOFT路径
'@返回值  String
Private Function AppsoftPath() As String
    If IsDesinMode Then
        AppsoftPath = "C:\APPSOFT"
    Else
        AppsoftPath = Mid(App.Path & "\", 1, InStr(5, App.Path & "\", "\"))
        If Right(AppsoftPath, 1) = "\" Then AppsoftPath = Mid(AppsoftPath, 1, Len(AppsoftPath) - 1)
    End If
End Function
'@方法    IsDesinMode
'   当前是否是源码环境
'@返回值  Boolean
Private Function IsDesinMode() As Boolean
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function
'@方法    LongToUnsigned
'   无符号长整形处理
'@返回值  Double
'   按double返回
'@参数:
'Value  Long(In)
'   待转换的长整型
Private Function LongToUnsigned(Value As Long) As Double
    Const OFFSET_4        As Double = 4294967296#         '无符号整形的最大值
    If Value < 0 Then LongToUnsigned = Value + OFFSET_4 Else LongToUnsigned = Value
End Function
'@方法    IsInCollection
'   检查集合中是否存在某元素
'@返回值  Boolean
'@参数:
'cllTest    Collection(In)
'   要检查的集合
'strKey     String(In)
'   要检查的Key
Private Function IsInCollection(cllTest As Collection, strKey As String) As Boolean
    On Error GoTo ErrorH
    IsInCollection = IsObject(cllTest.Item(strKey))
    IsInCollection = True
    Exit Function
ErrorH:
    IsInCollection = False
End Function
'@方法    Serialize
'   将对象或值序列化为字符串
'@返回值  String
'
'@参数:
'objInfo  Variant(In)
'   序列化的对象
'@备注
'
Private Function Serialize(ByVal objInfo As Variant) As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long
    Const KEY_DEFAULT_NAME = "K0"
    On Error Resume Next
    If IsArray(objInfo) Then
        objBag.WriteProperty "KL", UBound(objInfo)
        For i = LBound(objInfo) To UBound(objInfo)
            If IsArray(objInfo(i)) Then
                objBag.WriteProperty "A" & i, 1
                objBag.WriteProperty "K" & i, Serialize(objInfo(i))
            Else
                objBag.WriteProperty "K" & i, objInfo(i)
                If Err.Number = 330 Then
                    '非法参数。  因为不支持持久性不能写对象。
                    Err.Clear
                    objBag.WriteProperty "K" & i, Nothing
                End If
            End If
        Next
        bytData = objBag.Contents
        Serialize = EncodeBase64(bytData())
    Else
        objBag.WriteProperty KEY_DEFAULT_NAME, objInfo
        If Err.Number = 330 Then
            '非法参数。  因为不支持持久性不能写对象。
            Serialize = "{NotPersistable}"
            Err.Clear
        Else
            bytData = objBag.Contents
            Serialize = EncodeBase64(bytData())
        End If
    End If

End Function
'@方法    UnSerialize
'   将字符串反序列化为对象或具体的值
'@返回值  Variant
'
'@参数:
'strSource  String In
'   序列化字符串
'@备注
'
Private Function UnSerialize(ByVal strSource As String) As Variant
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long, lngLen     As Long
    Dim arrVar()    As Variant
    Const KEY_DEFAULT_NAME = "K0"
    Const KEY_LLENGTH = "KL"
    
    On Error Resume Next
    If Len(strSource) = 0 Then Exit Function
    If strSource = "{NotPersistable}" Then
         Set UnSerialize = Nothing
    Else
        bytData = DecodeBase64(strSource, True)
        objBag.Contents = bytData
        lngLen = objBag.ReadProperty(KEY_LLENGTH, -1)
        '仅有单个值序列化
        If lngLen = -1 Then
            If Not IsObject(objBag.ReadProperty(KEY_DEFAULT_NAME)) Then
                UnSerialize = objBag.ReadProperty(KEY_DEFAULT_NAME)
            Else
                Set UnSerialize = objBag.ReadProperty(KEY_DEFAULT_NAME)
            End If
        Else
            ReDim Preserve arrVar(lngLen)
            For i = 0 To lngLen
                If Not IsObject(objBag.ReadProperty("K" & i)) Then
                    If objBag.ReadProperty("A" & i, 0) = 1 Then
                        arrVar(i) = UnSerialize(objBag.ReadProperty("K" & i))
                    Else
                        arrVar(i) = objBag.ReadProperty("K" & i)
                    End If
                Else
                    Set arrVar(i) = objBag.ReadProperty("K" & i)
                End If
            Next
            UnSerialize = arrVar()
        End If
    End If
End Function

'@方法    SerializeEx
'   按顺序序列化多个信息
'@返回值  String
'
'@参数:
'arrInfo  ParamArray  In
'   多个序列化的对象
'@备注
'
Private Function SerializeEx(ParamArray arrInfo() As Variant) As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long
    On Error Resume Next
    If UBound(arrInfo) < 0 Then Exit Function
    If UBound(arrInfo) = 0 Then
        SerializeEx = Serialize(arrInfo(0))
    Else
        objBag.WriteProperty "KL", UBound(arrInfo)
        For i = 0 To UBound(arrInfo)
            If IsArray(arrInfo(i)) Then
                objBag.WriteProperty "A" & i, 1
                objBag.WriteProperty "K" & i, Serialize(arrInfo(i))
            Else
                objBag.WriteProperty "K" & i, arrInfo(i)
            End If
            If Err.Number = 330 Then
                '非法参数。  因为不支持持久性不能写对象。
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
        bytData = objBag.Contents
        SerializeEx = EncodeBase64(bytData())
    End If
End Function
'@方法    StringToUTF8Bytes
'   将字符串转换为UTF-8编码的字节数组
'@返回值  Variant
'  字符串转换的字节组
'@参数:
'strInput  String In
'   Unicode字符串
'@备注
'
Private Function StringToUTF8Bytes(strInput As String) As Byte()
    Const CP_UTF8  As Long = 65001
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '先计算需求字节数
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    '然后转换
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function
'@方法    UTF8BytesToString
'   将UTF-8编码的字节数组转换为字符串
'@返回值  String
'   转换后的字符串
'@参数:
'bytInpu  Byte() In
'   字节数组
'@备注
'
Private Function UTF8BytesToString(bytInpu() As Byte) As String
    Const CP_UTF8  As Long = 65001
    Dim lngBytesRequired As Long

    '先计算需求字节数
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    '然后转换
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function
'@方法    EncBase64Char
'   将6-bit字节转换为Base64字符
'@返回值  Byte
'   字符数值
'@参数:
'bytValue  Byte In
'   转换的字节
'@备注
'   Base64是将三个字节，每6位分割为四个字节处理的
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26个大写英文字母
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26个小写英文字母
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10个数字
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function
'@方法    DecBase64Char
'   将Base64字符转换为6 bit字节
'@返回值  Byte
'   字符数值
'@参数:
'bytValue  Byte In
'   待解码的字节
'@备注
'   Base64是将三个字节，每6位分割为四个字节处理的
Private Function DecBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue >= &H41 And bytValue <= &H5A Then
        DecBase64Char = bytValue - &H41
    ElseIf bytValue >= &H61 And bytValue <= &H7A Then
        DecBase64Char = bytValue - &H61 + 26
    ElseIf bytValue >= &H30 And bytValue <= &H39 Then
        DecBase64Char = bytValue - &H30 + 52
    ElseIf bytValue = &H2B Then
        DecBase64Char = 62
    ElseIf bytValue = &H2F Then
        DecBase64Char = 63
    End If
End Function
'@方法    EncodeBase64
'   进行Base64编码，返回Base64的字符串
'@返回值  String
'   Base64编码结果
'@参数:
'varInput  Variant
'   需要进行Base64编码的字符串或者字节数组，字符串采取UTF-8编码。Byte()类型前面的数组，元素个数传3的倍数，最后一次传递所有剩下的即可。
'@备注
'   Base64是将三个字节，每6位分割为四个字节处理的
Private Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        '原始内容,先将原文以UTF-8的方式编码
        bytInput = StringToUTF8Bytes(CStr(varInput))
    ElseIf VarType(varInput) = vbArray + vbByte Then
        If UBound(varInput) < 0 Then Exit Function
        bytInput = varInput
    Else
        Exit Function
    End If
    lngInputLen = UBound(bytInput) + 1
 
    lngOutLen = lngInputLen + (lngInputLen - 1) \ 3 + 1
    ReDim bytOut(lngOutLen - 1)
    '将8-bit字节数组转换为6-bit字节数组
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) '转换为Base64字符
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") '原文剩余内容不足3个字节需要补齐
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'@方法    DecodeBase64
'   将Base64的字符串解码为原文。
'@返回值  Variant
'   原始字符或者原始的字节组
'@参数:
'strInput  String In
'   Base64编码字符串
'blnByteArray  Boolean In,opt
'   True:返回Byte(),False-返回string
'@备注
'   Base64是将三个字节，每6位分割为四个字节处理的
Private Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo ErrH
    If Len(strInput) = 0 Then Exit Function
    lngModLen = InStr(strInput, "=")
    If lngModLen > 0 Then
        '编码后的内容
        lngModLen = Len(strInput) - lngModLen + 1
        bytInput = StrConv(strInput, vbFromUnicode)
    Else
        lngModLen = 0
        '编码后的内容
        bytInput = StrConv(strInput, vbFromUnicode)
    End If
    lngInputLen = UBound(bytInput) + 1
 
    '原始内容
    lngOutLen = lngInputLen - lngInputLen \ 4
    lngOutLen = lngOutLen - lngModLen
    ReDim bytOut(lngOutLen - 1)
 
    For j = 0 To lngInputLen - 1
        bytInput(j) = DecBase64Char(bytInput(j)) '从Base64字符转换为6-bit字节
    Next
    '将6-bit字节数组转换为8-bit字节数组
    For j = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(j) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两字节
            bytOut(j) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四字节
            bytOut(j) = (bytInput(i) And &H3) * &H40
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or bytInput(i)
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            lngBit = 0
        End If
    Next
    If blnByteArray Then
        DecodeBase64 = bytOut
    Else
        '最后将转换得到的UTF-8字符串转换为VB支持的Unicode字符串以便于显示。
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
ErrH:
    Err.Clear
End Function
'---------------------------------------------------------------------------
'                5、类初始化与销毁方法
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    mlngStart = GetTickCount
    mlngMaxLogCount = M_MAX_LINE_COUNT_INIT
    mblnLogOpened = True
    mllCurrentLogLevel = LogLevel_Warn
    mblnLoopTimeFormat = False
    mblnAppend = False
    mdtStart = Now
    mstrLogLevelName(LogLevel_UnDefined + 1) = "[UnDefine]"
    mstrLogLevelName(LogLevel_LogOFF + 1) = "[ LogOFF ]"
    mstrLogLevelName(LogLevel_Error + 1) = "[  Error ]"
    mstrLogLevelName(LogLevel_Warn + 1) = "[  Warn  ]"
    mstrLogLevelName(LogLevel_Info + 1) = "[  Info  ]"
    mstrLogLevelName(LogLevel_Debug + 1) = "[  Debug ]"
    mstrLogLevelName(LogLevel_Trace + 1) = "[  Trace ]"
    mstrLogLevelName(LogLevel_AllLog + 1) = "[ AllLog ]"
    Call LoadLogLevelSetting
End Sub

Private Sub Class_Terminate()
    Call CloseEx
    Erase mstrLogLevelName
    Set mcllAllLogLevel = Nothing
    Set mcllAllLogLevel = Nothing
    Set mobjFSO = Nothing
End Sub
