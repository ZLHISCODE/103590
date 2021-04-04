Attribute VB_Name = "mdlPubMethod"
'@模块 mdlPubMethod-2019/6/26
'@编写 lshuo
'@功能
'   公共函数方法
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
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
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。

'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------

'@方法    TruncZero
'   去掉字符串中\0以后的字符，常用于API返回字符串处理
'@返回值  String
'
'@参数:
'strInput String In
'   待处理的字符串
'@备注
'
Public Function TruncZero(ByVal strInput As String) As String
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

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
Public Function DisPlayOneValue(valValue As Variant, Optional ByVal blnSerializeObject As Boolean = True) As String
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
'@方法    Serialize
'   将对象或值序列化为字符串
'@返回值  String
'
'@参数:
'objInfo  Variant(In)
'   序列化的对象
'@备注
'
Public Function Serialize(ByVal objInfo As Variant) As String
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
Public Function UnSerialize(ByVal strSource As String) As Variant
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
Public Function SerializeEx(ParamArray arrInfo() As Variant) As String
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
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Const CP_UTF8           As Long = 65001
    Dim bytUTF8Bytes()      As Byte
    Dim lngBytesRequired    As Long
    
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
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Const CP_UTF8  As Long = 65001
    Dim lngBytesRequired As Long

    '先计算需求字节数
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    '然后转换
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
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
Public Function EncodeBase64(varInput As Variant) As String
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
Public Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
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

'@方法    DecodeEx
'   模拟Oracle的Decode函数
'@返回值  Variant
'
'@参数:
'arrPar ParamArray  In
'   当前值,判定值1,返回值1,判定值2,返回值1,...,判定值n,返回值n,缺省返回值
'   若当前值=判定值i,则返回返回值i,若没任何一个匹配，则返回缺省返回值
'   缺省值可以不传，则返回EMPTY
'@备注
'
Public Function DecodeEx(ParamArray arrPar() As Variant) As Variant
'功能：
    Dim varValue    As Variant, i   As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            DecodeEx = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            DecodeEx = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

'@方法    FromatSQL
'   去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'@返回值  String
'
'@参数:
'strText String In
'   处理字符
'blnCrlf Boolean In (Optional)
'   是否去掉换行符
'@备注
'
Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

'@方法    GetTickCountDiff
'   计算GetTickCcout的差值。由于 GetTickCountVB会产生负值以及归零现象因此需要单独处理
'@返回值  Double
'
'@参数:
'lngStart Long In
'   起始时间
'lngEnd Long In (Optional)
'   结束时间
'blnInputEnd Boolean In  (Optional)
'   标识是否传入了lngEnd
'@备注
'
Public Function GetTickCountDiff(ByVal lngStart As Long, Optional ByVal lngEnd As Long, Optional ByVal blnInputEnd As Boolean) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    If blnInputEnd Then
        lngCur = lngEnd
    Else
        lngCur = GetTickCount
    End If
    If lngCur < lngStart Then
        GetTickCountDiff = M_OFFSET_4 - LongToUnsigned(lngStart) + LongToUnsigned(lngCur)
    Else
        GetTickCountDiff = lngCur - lngStart
    End If
End Function
'@方法    IsDesinMode
'   当前是否是源码环境
'@返回值  Boolean
Public Function IsDesinMode() As Boolean
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function

'@方法    NvlEx
'   相当于Oracle的NVL，将Null值改成另外一个预设值
'@返回值  Variant
'
'@参数:
'varValue Variant In
'   判断的值
'DefaultValue Variant In (Optional,Default="")
'   缺省值
'@备注
'
Public Function NvlEx(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NvlEx = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
'@方法    IP2String
'   将IP转换为String
'@返回值  String
'
'@参数:
'lngIP Long In
'   IP数值
'@备注
'
Public Function IP2String(ByVal lngIP As Long) As String
    Dim arrByt(3)       As Byte
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.IP2String", lngIP)
    RtlMoveMemory ByVal VarPtr(arrByt(0)), lngIP, 4
    IP2String = arrByt(0) & "." & arrByt(1) & "." & arrByt(2) & "." & arrByt(3)
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IP2String", IP2String)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.IP2String") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.IP2String")
End Function

'@方法    String2IP
'   将字符串IP转换为数值
'@返回值  String
'
'@参数:
'strIP String In
'   字符串IP
'@备注
'
Public Function String2IP(ByVal strIp As String) As Long
    Dim arrByt(3)       As Byte
    Dim lngIP           As Long
    Dim arrTmp          As Variant
    Dim i               As Long
    Dim blnOK           As Boolean
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.String2IP", strIp)
    arrTmp = Split(strIp, ".")
    blnOK = True
    If UBound(arrTmp) = 3 Then
        For i = LBound(arrTmp) To UBound(arrTmp)
            If IsNumeric(i) Then
                If Val(arrTmp(i)) < 0 Or Val(arrTmp(i)) > 255 Then
                    blnOK = False
                End If
            Else
                blnOK = False
            End If
            If blnOK Then
                arrByt(i) = Val(arrTmp(i))
            Else
                Exit For
            End If
        Next
    Else
        blnOK = False
    End If
    If blnOK Then
        RtlMoveMemory lngIP, ByVal VarPtr(arrByt(0)), 4
    End If
    String2IP = lngIP
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.String2IP", String2IP)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.String2IP") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.String2IP")
End Function

'@方法    FormatIpString
'   对分段补足三位数
'@返回值  String
'
'@参数:
'strIp String In
'   IP地址
'@备注
'
Public Function FormatIpString(ByVal strIp As String) As String
    Dim arrByt(3)       As Byte
    Dim arrTmp          As Variant
    Dim i               As Long
    Dim blnOK           As Boolean
    

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.FormatIpString", strIp)
    arrTmp = Split(strIp, ".")
    blnOK = True
    If UBound(arrTmp) = 3 Then
        For i = LBound(arrTmp) To UBound(arrTmp)
            If IsNumeric(i) Then
                If Val(arrTmp(i)) < 0 Or Val(arrTmp(i)) > 255 Then
                    blnOK = False
                End If
            Else
                blnOK = False
            End If
            If blnOK Then
                arrByt(i) = Val(arrTmp(i))
            Else
                Exit For
            End If
        Next
    Else
        blnOK = False
    End If
    
    If blnOK Then
         FormatIpString = Format(arrByt(0), "000") & "." & Format(arrByt(1), "000") & "." & Format(arrByt(2), "000") & "." & Format(arrByt(3), "000")
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.FormatIpString", FormatIpString)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.FormatIpString") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.FormatIpString")
End Function

'@方法    NormalIpString
'   从格式化的IP返回原始IP
'@返回值  String
'
'@参数:
'strIp String In
'   IP地址
'@备注
'
Public Function NormalIpString(ByVal strIp As String) As String
    Dim arrByt(3)       As Byte
    Dim arrTmp          As Variant
    Dim i               As Long
    Dim blnOK           As Boolean
    

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsEnvironment.NormalIpString", strIp)
    arrTmp = Split(strIp, ".")
    blnOK = True
    If UBound(arrTmp) = 3 Then
        For i = LBound(arrTmp) To UBound(arrTmp)
            If IsNumeric(i) Then
                If Val(arrTmp(i)) < 0 Or Val(arrTmp(i)) > 255 Then
                    blnOK = False
                End If
            Else
                blnOK = False
            End If
            If blnOK Then
                arrByt(i) = Val(arrTmp(i))
            Else
                Exit For
            End If
        Next
    Else
        blnOK = False
    End If
    
    If blnOK Then
         NormalIpString = arrByt(0) & "." & arrByt(1) & "." & arrByt(2) & "." & arrByt(3)
    End If
    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.NormalIpString", NormalIpString)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsEnvironment.NormalIpString") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsEnvironment.NormalIpString")
End Function

'@方法    VerFull
'   返回VB最大支持的版本号形式:9999.9999.9999.9999,最小版本号0000.0000.0000.0000
'@返回值  String
'
'@参数:
'strVer String In
'   原始版本号
'blnMax Boolean In (Optional)
'   True=若果为空，则返回最大支持版本。False=若果为空，则返回最小支持版本
'@备注
'
Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '增加一段，以兼容特殊SP版本号
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

'@方法    IsVerSion
'   判断字符串是否是版本号
'@返回值  Boolean
'
'@参数:
'strVer String In
'   原始版本号
'blnOnlyCheckSpecial Boolean In
'   检查版本号是否是特殊SP版本号
'@备注
'
Public Function IsVerSion(ByVal strVer As String, Optional ByVal blnOnlyCheckSpecial As Boolean) As Boolean
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    If blnOnlyCheckSpecial And UBound(arrVer) <> 3 Then Exit Function
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val("1" & arrVer(i)) & "" <> Trim("1" & arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

'@方法    IsEmptyArray
'   判断对象是否是空数组
'@返回值  Boolean
'
'@参数:
'varAnyArray Variant In
'   判断的数组
'@备注
'
Public Function IsEmptyArray(varAnyArray As Variant) As Boolean
    Dim lngUbound               As Long
    On Error GoTo ErrH
    
    If IsEmpty(varAnyArray) Then
        IsEmptyArray = True
    ElseIf IsArray(varAnyArray) Then
        lngUbound = UBound(varAnyArray)
        IsEmptyArray = (lngUbound - LBound(varAnyArray)) < 0
    Else
        IsEmptyArray = True
    End If
    Exit Function
ErrH:
    IsEmptyArray = True
End Function

'密码加密程序
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '获取随机种子
    '随机种子的随机数为999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    Rnd (-1)
    Randomize (Val(strSeed))
    intlen = Len(strText)
    For i = 1 To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        ElseIf intChr < 0 Then
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Cipher = strDeText
End Function

Public Function DeCipher(ByVal strText As String) As String
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        Else
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    DeCipher = strDeText
End Function

'@方法    To_DateEx
'   获取ORACLE Date类型串
'@返回值  String
'   ORACLE Date类型串
'@参数:
'strDate String In
'   时间字符串
'strType String In
'   格式字符串类型，ymd-年月日（yyyy-mm-dd)，ymdhm-（yyyy-mm-dd hh:mm),ymdhms-（yyyy-mm-dd hh:mm:ss)
'@备注
'
Public Function To_DateEx(ByVal strDate As String, Optional ByVal strType As String = "YMDHMS") As String
    If Not IsDate(strDate) Then To_DateEx = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           To_DateEx = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           To_DateEx = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           To_DateEx = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           To_DateEx = "Null"
    End Select
End Function

'@方法    InCollection
'   检查集合中是否存在某元素
'@返回值  Boolean
'
'@参数:
'cllTest Collection In
'   要检查的集合
'strKey String In
'   要检查的Key
'@备注
'
Public Function InCollection(cllTest As Collection, strKey As String) As Boolean
    On Error GoTo ErrorH
    If VarType(cllTest.Item(strKey)) = vbObject Then
    End If
    InCollection = True
    Exit Function
ErrorH:
    InCollection = False
End Function

'@方法    TrimEx
'   去除strTrim两边的strTrmChar,功能类似Trim
'@返回值  String
'
'@参数:
'strTrim String In
'   需要格式化的字符
'strTrmChar String In (Optional,Default=" ")
'   不传strTrmChar或者传空格时，相当Trim
'@备注
'
Public Function TrimEx(ByVal strTrim As String, Optional ByVal strTrmChar As String = " ") As String
    Dim i As Integer, intB As Integer, intE As Integer
    
    If strTrim = "" Or strTrmChar = "" Then TrimEx = strTrim: Exit Function
    If strTrmChar = " " Then TrimEx = Trim(strTrim): Exit Function
    
    intB = 1
    For i = 1 To Len(strTrim)
        If Mid(strTrim, i, 1) <> strTrmChar Then intB = i: Exit For
    Next
    intE = Len(strTrim)
    For i = Len(strTrim) To 1 Step -1
        If Mid(strTrim, i, 1) <> strTrmChar Then intE = i: Exit For
    Next
    TrimEx = Mid(strTrim, intB, intE - intB + 1)
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
Public Function UboundEx(varArray As Variant) As Long
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
Public Function LboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    LboundEx = LBound(varArray)
    Exit Function
ErrH:
    LboundEx = 0
End Function
'@方法    AppsoftPath
'   获取APPSOFT路径
'@返回值  String
Public Function AppsoftPath() As String
    Static strAPPSOFT       As String
    
    If LenB(strAPPSOFT) = 0 Then
        If IsDesinMode Then
            strAPPSOFT = "C:\APPSOFT"
        Else
            strAPPSOFT = Mid(App.Path & "\", 1, InStr(5, App.Path & "\", "\"))
            If Right(AppsoftPath, 1) = "\" Then strAPPSOFT = Mid(AppsoftPath, 1, Len(AppsoftPath) - 1)
        End If
    End If
    AppsoftPath = strAPPSOFT
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------
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

'@方法    LongToUnsigned
'   将有符号Long转换为无符号值
'@返回值  Double
'
'@参数:
'Value Long In
'   有符号Long
'@备注
'
Private Function LongToUnsigned(Value As Long) As Double
    Const M_OFFSET_4    As Double = 4294967296#         '无符号整形的最大值
    If Value < 0 Then LongToUnsigned = Value + M_OFFSET_4 Else LongToUnsigned = Value
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------



