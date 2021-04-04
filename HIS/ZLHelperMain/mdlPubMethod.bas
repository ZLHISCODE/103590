Attribute VB_Name = "mdlPubMethod"
'@ģ�� mdlPubMethod-2019/6/26
'@��д lshuo
'@����
'   ������������
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
'@����
'    ���ַ���ӳ�䵽UTF-16(���ַ�)�ַ������ַ�����һ�����Զ��ֽ��ַ�����
'    �����ʹ��MultiByteToWideChar��������Ӧ�ó���İ�ȫ�ԡ�����������������׵��»������������ΪlpMultiByteStr��ʾ�����뻺�����Ĵ�С�����ַ����е��ֽ�������lpWideCharStr��ʾ������������Ĵ�С�����ַ�����Ϊ�˱��⻺���������Ӧ�ó������Ϊ���������յ���������ָ���ʵ��Ļ�������С���йظ�����Ϣ����μ���ȫ�Կ���:�������ԡ�
'    ע�⣬ANSI����ҳ�����ڲ�ͬ�ļ�����ϲ�ͬ��Ҳ�������һ̨��������и��ģ��Ӷ����������𻵡�Ϊ�˻����һ�µĽ����Ӧ�ó���Ӧ��ʹ��Unicode����UTF-8��UTF-16���������ض��Ĵ���ҳ������������׼�����ݸ�ʽ��ֹʹ��Unicode������޷�ʹ��Unicode��Ӧ�ó���Ӧ����Э������������ʹ���ʵ��ı������Ʊ����������HTML��XML�ļ������ǣ������ı��ļ�������
'@ԭ��
'    int MultiByteToWideChar(
'      _In_      UINT   CodePage,
'      _In_      DWORD  dwFlags,
'      _In_      LPCSTR lpMultiByteStr,
'      _In_      int    cbMultiByte,
'      _Out_opt_ LPWSTR lpWideCharStr,
'      _In_      int    cchWideChar
'    );
'@����
'    CodePage
'    ִ��ת��ʱʹ�õĴ���ҳ���˲�����������Ϊ����ϵͳ���Ѱ�װ����õ��κδ���ҳ��ֵ���йش���ҳ�б���μ�����ҳ��ʶ��������Ӧ�ó��򻹿���ָ���±�����ʾ��ֵ֮һ��
Private Const CP_ACP        As Long = 0
'    ϵͳĬ�ϵ�Windows ANSI����ҳ?
'    ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
Private Const CP_MACCP      As Long = 1
'    ��ǰϵͳMacintosh����ҳ?
'    ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
'    ע�⣬���ֵ��Ҫ�����������룬�����ִ�Macintosh�����ʹ��Unicode���б��룬����ͨ������Ҫ���ֵ��
Private Const CP_OEMCP      As Long = 2
'    ��ǰϵͳOEM����ҳ?
'    ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
Private Const CP_SYMBOL     As Long = 42
'    ���Ŵ���ҳ(42)��
Private Const CP_THREAD_ACP As Long = 3
'    ��ǰ�̵߳�Windows ANSI����ҳ?
'    ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
Private Const CP_UTF7       As Long = 65000
'    utf - 7��ֻ����7λ�������ǿ��ʱ��ʹ�ô�ֵ�����ʹ��UTF-8��
Private Const CP_UTF8       As Long = 65001
'    utf - 8��
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
'    ָʾת�����͵ı�־��Ӧ�ó������ָ������ֵ����ϣ�Ĭ��ֵΪmb_precomposition��Ԥ��Ϻ͸����ǻ���ġ���������MB_USEGLYPHCHARS��MB_ERR_INVALID_CHARS��������������־��״̬��Ρ�
Private Const MB_COMPOSITE          As Long = &H2
'    ʼ��ʹ�÷ֽ���ַ����������ַ���һ�������Ǽ���ַ������в�ͬ�Ĵ����ֵ���ַ������磬A��A +����ʾ:������д��ĸA (U+0041) +���DIAERESIS (U+0308)��ע�⣬�˱�־������mb_precompositionһ��ʹ�á�
Private Const MB_ERR_INVALID_CHARS As Long = &H8
'    ���������Ч�������ַ�����ʧ�ܡ�
'    ��Windows Vista��ʼ�����Ӧ�ó���û�����ô˱�־���ú�������ɾ���Ƿ�����㣬������U+FFFD�滻�Ƿ�����(����ָ���Ĵ���ҳ�����ʵ�����)��
'    ���û�����ô˱�־���ú������Զ�ɾ���Ƿ�����㡣��GetLastError�ĵ��÷���ERROR_NO_UNICODE_TRANSLATION��
Private Const MB_PRECOMPOSED        As Long = &H1
'    ΥԼ;��Ҫ��MB_COMPOSITEһ��ʹ�á�ʼ��ʹ��Ԥ����ַ��������ڻ��ַ���Ǽ���ַ���Ͼ��е����ַ�ֵ���ַ������磬���ַ�e�У�e�ǻ����ַ��������������޼���ַ������Ϊһ���ַ�������һ��Unicode����㣬Ӧ�ó���Ӧ��ʹ�����������ǵ����Ļ����ַ��ͷǼ���ַ������磬A�ɵ�Unicode�����������д��ĸA��DIAERESIS (U+00C4)��ʾ��
Private Const MB_USEGLYPHCHARS      As Long = &H4
'    ʹ�������ַ������ǿ����ַ�?
'    ���������г��Ĵ���ҳ��dwFlags��������Ϊ0�����򣬺�����ʹ��ERROR_INVALID_FLAGSʧ�ܡ�
'       50220
'       50221
'       50222
'       50225
'       50227
'       50229
'       57002     through 57011
'       65000 (UTF-7)
'       42 (Symbol)
'    ע��:����UTF-8�����ҳ54936 (GB18030����Windows Vista��ʼ)��dwFlags��������Ϊ0��MB_ERR_INVALID_CHARS�����򣬺�����ʹ��ERROR_INVALID_FLAGSʧ�ܡ�
'    lpMultiByteStr [
'       ָ��Ҫת�����ַ�����ָ��?
'    cbMultiByte
'    lpMultiByteStr������ʾ���ַ����Ĵ�С(���ֽ�Ϊ��λ)�����ߣ�����ַ�����null��β�����Խ��ò�������Ϊ-1��ע�⣬���cbMultiByteΪ0��������ʧ�ܡ�
'    ����ò���Ϊ-1�������������������ַ�����������ֹnull�ַ�����ˣ��õ���Unicode�ַ�����һ����ֹnull�ַ����������صĳ��Ȱ�������ַ���
'    ������˲�������Ϊ����������������ȷ����ָ�����ֽ���������ṩ�Ĵ�С��������ֹnull�ַ��������ɵ�Unicode�ַ���������null��β�ģ����صĳ���Ҳ���������ַ���
'    lpWideCharStr(,��ѡ)
'    ָ�����ת���ַ����Ļ�������ָ��?
'    cchWideChar [��]
'    lpWideCharStr��ʾ�Ļ������Ĵ�С(���ַ�Ϊ��λ)�������ֵΪ0����������������Ļ�������С�����ַ�Ϊ��λ�������κ���ֹnull�ַ������Ҳ�ʹ��lpWideCharStr��������
'@����ֵ
'    ����ɹ�������д��lpWideCharStrָʾ�Ļ��������ַ�������������ɹ���cchWideCharΪ0���򷵻�ֵ��lpWideCharStr��ָʾ�Ļ���������Ĵ�С(���ַ�Ϊ��λ)���й�MB_ERR_INVALID_CHARS��־��������Ч����ʱ���Ӱ�췵��ֵ����Ϣ�������dwFlags��
'    �������û�гɹ����򷵻�0��Ҫ�����չ�Ĵ�����Ϣ��Ӧ�ó�����Ե���GetLastError�������Է������´������֮һ:
'    ERROR_INSUFFICIENT_BUFFER�����ṩ�Ļ�������С�����󣬻��߱����������ΪNULL��
'    ERROR_INVALID_FLAGS?Ϊ��־�ṩ��ֵ��Ч?
'    ERROR_INVALID_PARAMETER?�κβ���ֵ����Ч?
'    ERROR_NO_UNICODE_TRANSLATION?���ַ����з�����Ч��Unicode?
'@��ע
'    �˺�����Ĭ����Ϊ��ת��Ϊ�����ַ�����Ԥ�����ʽ�����������Ԥ�����ʽ���ú���������ת��Ϊ�����ʽ��
'    ʹ��mb_precomposedflag�Դ��������ҳӰ���С����Ϊ��������������Ѿ�����Ϻ��ˡ�������ʹ��MultiByteToWideChar����ת�������NormalizeString��NormalizeString�ṩ�˸�׼ȷ����׼��һ�µ����ݣ������ٶȸ��졣ע�⣬���ڴ��ݸ�NormalizeString��NORM_FORMö�٣�NormalizationC��Ӧ��mb_precomposition, NormalizationD��Ӧ��MB_COMPOSITE��
'    ������ľ�����������������ȵ��ô˺���������cchWideChar����Ϊ0�Ի������Ĵ�С��������������������������ʹ��MB_COMPOSITE��־��ÿ�������ַ���������ȿ���������������ַ���
'    lpMultiByteStr��lpWideCharStrָ�벻����ͬ�������������ͬ�ģ�����ʧ�ܣ�GetLastError����ֵERROR_INVALID_PARAMETER��
'    �����ʽָ�������ַ������ȶ�û����ֹ���ַ�����MultiByteToWideChar����Ϊ����ֹ����ַ�������Ҫ����ֹ�˺���������ַ�����Ӧ�ó���Ӧ����-1����ʽ���������ַ�������ֹ���ַ���
'    ���������MB_ERR_INVALID_CHARS��������Դ�ַ�����������Ч�ַ�����ú�����ʧ�ܡ���Ч�ַ��������ַ�֮һ:
'    ����Դ�ַ����е�Ĭ���ַ�������δ����MB_ERR_INVALID_CHARSʱת��ΪĬ���ַ����ַ�
'    ����DBCS�ַ���������ǰ���ֽڵ�û����Ч�����ֽڵ��ַ�
'    ��Windows Vista��ʼ�����������ȫ����Unicode 4.1��UTF-8��UTF-16�淶�������ڲ���ϵͳ��ʹ�õĺ����������뵥����������һ���ƥ��Ĵ������ԡ������ڰ汾��Windows�б�д��������������Ϊ������������ı����������ݵĴ�����ܻ��������⡣���ǣ�����Ч��UTF-8�ַ�����ʹ�øú����Ĵ������Ϊ����������Windows����ϵͳ����ͬ��
'    Windows XP:Ϊ�˷�ֹUTF-8�ַ��ķǶ̸�ʽ�汾�İ�ȫ���⣬MultiByteToWideCharɾ������Щ�ַ���
'    ��Windows 8��ʼ:MultiByteToWideChar����stringapi .h�����ġ���Windows 8֮ǰ��������Winnls.h�������ġ�
'@Requirements
'Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'Minimum supported phone    Windows Phone 8
'Header                     Stringapiset.h (include Windows.h)
'Library                    kernel32.lib
'dll                        kernel32.dll
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
'@����
'    ��UTF-16(���ַ�)�ַ���ӳ�䵽���ַ������µ��ַ�����һ�����Զ��ֽ��ַ�����
'    �����ʹ��WideCharToMultiByte��������Ӧ�ó���İ�ȫ�ԡ�����������������׵��»������������ΪlpWideCharStr��ʾ�����뻺�����Ĵ�С����Unicode�ַ����е��ַ�������lpMultiByteStr��ʾ������������Ĵ�С�����ֽ�����Ϊ�˱��⻺���������Ӧ�ó������Ϊ���������յ���������ָ���ʵ��Ļ�������С��
'    ��UTF-16ת��Ϊ��Unicode��������ݿ��ܻᶪʧ���ݣ���Ϊ����ҳ�����޷���ʾ�ض�Unicode������ʹ�õ�ÿ���ַ����йظ�����Ϣ����μ���ȫ�Կ���:�������ԡ�
'    ע�⣬ANSI����ҳ�����ڲ�ͬ�ļ�����ϲ�ͬ��Ҳ�������һ̨��������и��ģ��Ӷ����������𻵡�Ϊ�˻����һ�µĽ����Ӧ�ó���Ӧ��ʹ��Unicode����UTF-8��UTF-16���������ض��Ĵ���ҳ������������׼�����ݸ�ʽ��ֹʹ��Unicode������޷�ʹ��Unicode��Ӧ�ó���Ӧ����Э������������ʹ���ʵ��ı������Ʊ����������HTML��XML�ļ������ǣ������ı��ļ�������
'@ԭ��
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
'@����
'    CodePage
'        ִ��ת��ʱʹ�õĴ���ҳ���˲�����������Ϊ����ϵͳ���Ѱ�װ����õ��κδ���ҳ��ֵ���йش���ҳ�б���μ�����ҳ��ʶ��������Ӧ�ó��򻹿���ָ���±�����ʾ��ֵ֮һ��
'        ��ֵ����
'        CP_ACP
'        ϵͳĬ�ϵ�Windows ANSI����ҳ?
'        ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
'        CP_MACCP
'        ��ǰϵͳMacintosh����ҳ?
'        ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
'        ע�⣬���ֵ��Ҫ�����������룬�����ִ�Macintosh�����ʹ��Unicode���б��룬����ͨ������Ҫ���ֵ��
'        CP_OEMCP
'        ��ǰϵͳOEM����ҳ?
'        ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
'        CP_SYMBOL
'        �Ӵ�2000:���Ŵ���ҳ(42)��
'        CP_THREAD_ACP
'        Windows 2000: ��ǰ�̵߳�Windows ANSI����ҳ?
'        ע�⣬���ֵ�ڲ�ͬ�ļ�������ǲ�ͬ�ģ���������ͬ��������Ҳ�ǲ�ͬ�ġ���������ͬһ̨������ϸ��ģ��Ӷ����´洢�����ݲ��ɻָ����𻵡���ֵ��������ʱʹ�ã�������ܣ����ô洢Ӧ��ʹ��UTF-16��UTF-8��
'        CP_UTF7
'        utf - 7��ֻ����7λ�������ǿ��ʱ��ʹ�ô�ֵ�����ʹ��UTF-8��ʹ�����ֵ����lpDefaultChar��lpUsedDefaultChar��������ΪNULL��
'        CP_UTF8
'        utf - 8��ʹ�����ֵ����lpDefaultChar��lpUsedDefaultChar��������ΪNULL��
'    dwFlags [��]
'    ָʾת�����͵ı�־��Ӧ�ó������ָ������ֵ����ϡ���û��������Щ��־ʱ��������ִ���ٶȻ���졣Ӧ�ó���Ӧ��ָ��WC_NO_BEST_FIT_CHARS��WC_COMPOSITECHECK����ʹ���ض���ֵWC_DEFAULTCHAR�������п��ܵ�ת����������û���ṩ������ֵ���ͻᶪʧһЩ�����
Private Const WC_COMPOSITECHECK             As Long = &H200
'    ת������ַ������������ַ��ͷǼ���ַ���ÿ���ַ����в�ͬ���ַ�ֵ������Щ�ַ�ת��ΪԤ����ַ���Ԥ����ַ�����һ�����ڻ��Ǽ���ַ���ϵ��ַ�ֵ�����磬���ַ�e�У�e�ǻ����ַ��������������޼���ַ���
'    ע��:Windowsͨ��ʹ��Ԥ������ݱ�ʾUnicode�ַ��������û�б�Ҫʹ��WC_COMPOSITECHECK��־��
'    ����Ӧ�ó�����Խ�WC_COMPOSITECHECK�������κ�һ����־���������ȱʡֵΪWC_SEPCHARS����Unicode�ַ�����û�����ڻ��Ǽ���ַ���ϵ�Ԥ���ӳ��ʱ����Щ��־��������������Ϊ�����û���ṩ��Щ��־����������Ϊ����������WC_SEPCHARS��־һ�����йظ�����Ϣ����μ���ע�����е�WC_COMPOSITECHECK����ر�־��
'    ��ת���ڼ�ʹ��Ĭ���ַ��滻�쳣?
'    ת�������ж����Ǽ���ַ�?
Private Const WC_SEPCHARS                   As Long = &H20
'        Default?��ת���ڼ����ɵ������ַ�?
Private Const WC_ERR_INVALID_CHARS          As Long = &H80
'    Windows Vista���Ժ�汾:���������Ч�����ַ�����ʧ��(����0����last-error��������ΪERROR_NO_UNICODE_TRANSLATION)��������ͨ������GetLastError�������һ��������롣���δ���ô˱�־��������ʹ��U+FFFD�滻�Ƿ�����(����ָ���Ĵ���ҳ�����ʵ�����)����ͨ������ת���ַ����ĳ��ȳɹ���ע�⣬�˱�־�������ڽ�����ҳָ��ΪCP_UTF8��54936ʱ������������������ҳֵһ��ʹ�á�
Private Const WC_NO_BEST_FIT_CHARS          As Long = &H400
'    �����κ�û��ֱ�ӷ���ɶ��ֽڵ�ֵ��Unicode�ַ���lpDefaultCharָ����Ĭ���ַ������仰˵�������Unicodeת��Ϊ���ֽڲ��ٴ�ת����Unicode����������ͬ��Unicode�ַ�����ú���ʹ��Ĭ���ַ����˱�־���Ե���ʹ�ã�Ҳ�����������Ѷ���ı�־���ʹ�á�
'    ������Ҫ��֤���ַ��������ļ�����Դ���û�����Ӧ�ó���Ӧ��ʼ��ʹ��WC_NO_BEST_FIT_CHARS��־���˱�־��ֹ�������ַ�ӳ�䵽���������Ƶ�����ǳ���ͬ���ַ�����ĳЩ����£�����仯�����Ǽ��˵ġ����磬��ĳЩ����ҳ�У����ޡ�(��)�ķ���ӳ�䵽8(8)��
'    ���������г��Ĵ���ҳ��dwFlags��������Ϊ0�����򣬺�����ʹ��ERROR_INVALID_FLAGSʧ�ܡ�
'       50220
'       50221
'       50222
'       50225
'       50227
'       50229
'       57002     through 57011
'       65000 (UTF-7)
'       42 (Symbol)
'    ע��:����UTF-8�����ҳ54936 (GB18030����Windows Vista��ʼ)��dwFlags��������Ϊ0��MB_ERR_INVALID_CHARS�����򣬺�����ʹ��ERROR_INVALID_FLAGSʧ�ܡ�
'    lpWideCharStr [��]
'       ָ��Ҫת����Unicode�ַ�����ָ��?
'    cchWideChar [��]
'       lpWideCharStr��ʾ���ַ����Ĵ�С(���ַ�Ϊ��λ)�����ߣ�����ַ�����null��β�����Խ��ò�������Ϊ-1�������cchWideChar����Ϊ0��������ʧ�ܡ�
'       ����ò���Ϊ-1�������������������ַ�����������ֹnull�ַ�����ˣ��õ����ַ�����һ����ֹnull�ַ����������صĳ��Ȱ�������ַ���
'       ������˲�������Ϊ����������������ȷ����ָ�����ַ���������ṩ�Ĵ�С��������ֹnull�ַ��������ɵ��ַ�������null��β�����صĳ���Ҳ���������ַ���
'       lpMultiByteStr(,��ѡ)
'       ָ�����ת���ַ����Ļ�������ָ��?
'    cbMultiByte [��]
'       lpMultiByteStr��ʾ�Ļ������Ĵ�С(���ֽ�Ϊ��λ)��������ò�������Ϊ0���ú���������lpMultiByteStr����Ļ�������С�����Ҳ�ʹ�������������
'    lpDefaultChar(,��ѡ)
'       ����޷���ָ���Ĵ���ҳ�б�ʾ�ַ�����ָ��Ҫʹ�õ��ַ���ָ�롣�������Ҫʹ��ϵͳĬ��ֵ��Ӧ�ó��򽫸ò�������ΪNULL��Ҫ���ϵͳĬ���ַ���Ӧ�ó�����Ե���GetCPInfo��GetCPInfoEx������
'       ����CodePage��CP_UTF7��CP_UTF8���ã����뽫�ò�������ΪNULL�����򣬺�����ʹ��ERROR_INVALID_PARAMETERʧ�ܡ�
'    lpUsedDefaultChar(,��ѡ)
'       ָ��һ����־��ָ�룬�ñ�־ָʾ������ת�����Ƿ�ʹ����Ĭ���ַ������Դ�ַ����е�һ�������ַ�������ָ���Ĵ���ҳ�б�ʾ���򽫸ñ�־����ΪTRUE�����򣬽���־����ΪFALSE�����������������ΪNULL��
'       ����CodePage��CP_UTF7��CP_UTF8���ã����뽫�ò�������ΪNULL�����򣬺�����ʹ��ERROR_INVALID_PARAMETERʧ�ܡ�
'@����ֵ
'    ����ɹ�������lpMultiByteStrָ���д�뻺�������ֽ�������������ɹ���cbMultiByteΪ0���򷵻�ֵΪlpMultiByteStr��ָʾ�Ļ���������Ĵ�С(���ֽ�Ϊ��λ)���й�������Ч����ʱWC_ERR_INVALID_CHARS��־���Ӱ�췵��ֵ����Ϣ�������dwFlags��
'    �������û�гɹ����򷵻�0��Ҫ�����չ�Ĵ�����Ϣ��Ӧ�ó�����Ե���GetLastError�������Է������´������֮һ:
'    ERROR_INSUFFICIENT_BUFFER�����ṩ�Ļ�������С�����󣬻��߱����������ΪNULL��
'    ERROR_INVALID_FLAGS?Ϊ��־�ṩ��ֵ��Ч?
'    ERROR_INVALID_PARAMETER?�κβ���ֵ����Ч?
'    ERROR_NO_UNICODE_TRANSLATION?���ַ����з�����Ч��Unicode?
'@��ע
'    lpMultiByteStr��lpWideCharStrָ�벻����ͬ�������������ͬ�ģ�����ʧ�ܣ�GetLastError����ERROR_INVALID_PARAMETER��
'    WideCharToMultiByte����Ϊ�ա��������û����ֹ���ַ����������ʽָ�������ַ������ȣ�����ֹ����ַ�������Ҫ����ֹ�˺���������ַ�����Ӧ�ó���Ӧ����-1����ʽ���������ַ�������ֹ���ַ���
'    ���cbMultiByteС��cchWideChar�����������cbMultiByteָ�����ַ���д��lpMultiByteStrָ���Ļ����������ǣ������CodePage����ΪCP_SYMBOL������cbMultiByteС��cchWideChar����ú�������lpMultiByteStrд���ַ���
'    ��lpDefaultChar��lpUsedDefaultChar������ΪNULLʱ��WideCharToMultiByte����������Ч����ߡ��±���ʾ����Щ���������ֿ�����ϵĺ�����Ϊ��
'    lpDefaultChar lpuseddefaultchar        ���
'    NULL           NULL                    û��Ĭ�ϼ��?��Щ������������˺���һ��ʹ�õ�����Ч������?
'    �ǿ��ַ�       null                    ʹ��ָ����Ĭ���ַ�����������lpUsedDefaultChar��
'    null           �ǿ��ַ�                ʹ��ϵͳĬ���ַ������ڱ�Ҫʱ����lpUsedDefaultChar��
'    �ǿ��ַ�       �ǿ��ַ�                ʹ��ָ����Ĭ���ַ������ڱ�Ҫʱ����lpUsedDefaultChar��
'@Requirements
'Minimum supported client   Windows 2000 Professional [desktop apps | UWP apps]
'Minimum supported server   Windows 2000 Server [desktop apps | UWP apps]
'Minimum supported phone    Windows Phone 8
'Header                     Stringapiset.h (include Windows.h)
'Library                    kernel32.lib
'dll                        kernel32.dll
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'˵�������ڴ���һ��λ���ƶ�����һ��λ��
'Destination:ָ���ƶ�Ŀ�ĵ���ʼ��ַ��ָ�롣
'Source:ָ��Ҫ�ƶ����ڴ����ʼ��ַ��ָ�롣
'Length:�ڴ��Ĵ�С���ֽ�Ϊ��λ�ƶ���
'ע����������������ΪRtlMoveMemory����������ʵ���������ġ��йظ�����Ϣ����μ�WinBase��h��Winnt.h��Դ��Ŀ�����ܻ��ص���
'           ��һ��������Ŀ�ĵأ������㹻�������ɳ����ֽڵ�Դ;���򣬿��ܻ���ֻ��������������ܵ��¾ܾ����񹥻�������з���Υ�����������������£��������������Ľ���ע���ִ�д��롣���Ŀ�ĵ���һ�����ڶ�ջ�Ļ���������������ˡ�Ҫע�⣬���һ�����������ȣ��ǽ��ֽڸ��Ƶ�Ŀ�ĵص�������������Ŀ�ĵصĴ�С��

'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------

'@����    TruncZero
'   ȥ���ַ�����\0�Ժ���ַ���������API�����ַ�������
'@����ֵ  String
'
'@����:
'strInput String In
'   ��������ַ���
'@��ע
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

'@����    DisPlayOneValue
'   չʾ�����ֵ
'@����ֵ  String
'
'@����:
'valValue  Variant(In)
'   ת��Ϊ�ַ�����ֵ
'blnSerializeObject Boolean(In,opt,defualt=True)
'@��ע
'   �������Ϳ��ܲ�һ��֧�����л�
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
'@����    Serialize
'   �������ֵ���л�Ϊ�ַ���
'@����ֵ  String
'
'@����:
'objInfo  Variant(In)
'   ���л��Ķ���
'@��ע
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
                    '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
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
            '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
            Serialize = "{NotPersistable}"
            Err.Clear
        Else
            bytData = objBag.Contents
            Serialize = EncodeBase64(bytData())
        End If
    End If

End Function
'@����    UnSerialize
'   ���ַ��������л�Ϊ���������ֵ
'@����ֵ  Variant
'
'@����:
'strSource  String In
'   ���л��ַ���
'@��ע
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
        '���е���ֵ���л�
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

'@����    SerializeEx
'   ��˳�����л������Ϣ
'@����ֵ  String
'
'@����:
'arrInfo  ParamArray  In
'   ������л��Ķ���
'@��ע
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
                '�Ƿ�������  ��Ϊ��֧�ֳ־��Բ���д����
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
        bytData = objBag.Contents
        SerializeEx = EncodeBase64(bytData())
    End If
End Function
'@����    StringToUTF8Bytes
'   ���ַ���ת��ΪUTF-8������ֽ�����
'@����ֵ  Variant
'  �ַ���ת�����ֽ���
'@����:
'strInput  String In
'   Unicode�ַ���
'@��ע
'
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Const CP_UTF8           As Long = 65001
    Dim bytUTF8Bytes()      As Byte
    Dim lngBytesRequired    As Long
    
    '�ȼ��������ֽ���
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    'Ȼ��ת��
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function
'@����    UTF8BytesToString
'   ��UTF-8������ֽ�����ת��Ϊ�ַ���
'@����ֵ  String
'   ת������ַ���
'@����:
'bytInpu  Byte() In
'   �ֽ�����
'@��ע
'
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Const CP_UTF8  As Long = 65001
    Dim lngBytesRequired As Long

    '�ȼ��������ֽ���
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    'Ȼ��ת��
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function

'@����    EncodeBase64
'   ����Base64���룬����Base64���ַ���
'@����ֵ  String
'   Base64������
'@����:
'varInput  Variant
'   ��Ҫ����Base64������ַ��������ֽ����飬�ַ�����ȡUTF-8���롣Byte()����ǰ������飬Ԫ�ظ�����3�ı��������һ�δ�������ʣ�µļ��ɡ�
'@��ע
'   Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        'ԭʼ����,�Ƚ�ԭ����UTF-8�ķ�ʽ����
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
    '��8-bit�ֽ�����ת��Ϊ6-bit�ֽ�����
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д����λ
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) 'ת��ΪBase64�ַ�
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") 'ԭ��ʣ�����ݲ���3���ֽ���Ҫ����
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'@����    DecodeBase64
'   ��Base64���ַ�������Ϊԭ�ġ�
'@����ֵ  Variant
'   ԭʼ�ַ�����ԭʼ���ֽ���
'@����:
'strInput  String In
'   Base64�����ַ���
'blnByteArray  Boolean In,opt
'   True:����Byte(),False-����string
'@��ע
'   Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
Public Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo ErrH
    If Len(strInput) = 0 Then Exit Function
    lngModLen = InStr(strInput, "=")
    If lngModLen > 0 Then
        '����������
        lngModLen = Len(strInput) - lngModLen + 1
        bytInput = StrConv(strInput, vbFromUnicode)
    Else
        lngModLen = 0
        '����������
        bytInput = StrConv(strInput, vbFromUnicode)
    End If
    lngInputLen = UBound(bytInput) + 1
 
    'ԭʼ����
    lngOutLen = lngInputLen - lngInputLen \ 4
    lngOutLen = lngOutLen - lngModLen
    ReDim bytOut(lngOutLen - 1)
 
    For j = 0 To lngInputLen - 1
        bytInput(j) = DecBase64Char(bytInput(j)) '��Base64�ַ�ת��Ϊ6-bit�ֽ�
    Next
    '��6-bit�ֽ�����ת��Ϊ8-bit�ֽ�����
    For j = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)δ��д��
            bytOut(j) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)�ѱ�д�����ֽ�
            bytOut(j) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)�ѱ�д�����ֽ�
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
        '���ת���õ���UTF-8�ַ���ת��ΪVB֧�ֵ�Unicode�ַ����Ա�����ʾ��
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

'@����    DecodeEx
'   ģ��Oracle��Decode����
'@����ֵ  Variant
'
'@����:
'arrPar ParamArray  In
'   ��ǰֵ,�ж�ֵ1,����ֵ1,�ж�ֵ2,����ֵ1,...,�ж�ֵn,����ֵn,ȱʡ����ֵ
'   ����ǰֵ=�ж�ֵi,�򷵻ط���ֵi,��û�κ�һ��ƥ�䣬�򷵻�ȱʡ����ֵ
'   ȱʡֵ���Բ������򷵻�EMPTY
'@��ע
'
Public Function DecodeEx(ParamArray arrPar() As Variant) As Variant
'���ܣ�
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

'@����    FromatSQL
'   ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'@����ֵ  String
'
'@����:
'strText String In
'   �����ַ�
'blnCrlf Boolean In (Optional)
'   �Ƿ�ȥ�����з�
'@��ע
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

'@����    GetTickCountDiff
'   ����GetTickCcout�Ĳ�ֵ������ GetTickCountVB�������ֵ�Լ��������������Ҫ��������
'@����ֵ  Double
'
'@����:
'lngStart Long In
'   ��ʼʱ��
'lngEnd Long In (Optional)
'   ����ʱ��
'blnInputEnd Boolean In  (Optional)
'   ��ʶ�Ƿ�����lngEnd
'@��ע
'
Public Function GetTickCountDiff(ByVal lngStart As Long, Optional ByVal lngEnd As Long, Optional ByVal blnInputEnd As Boolean) As Double
    Dim lngCur          As Long
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
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
'@����    IsDesinMode
'   ��ǰ�Ƿ���Դ�뻷��
'@����ֵ  Boolean
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

'@����    NvlEx
'   �൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
'@����ֵ  Variant
'
'@����:
'varValue Variant In
'   �жϵ�ֵ
'DefaultValue Variant In (Optional,Default="")
'   ȱʡֵ
'@��ע
'
Public Function NvlEx(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NvlEx = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
'@����    IP2String
'   ��IPת��ΪString
'@����ֵ  String
'
'@����:
'lngIP Long In
'   IP��ֵ
'@��ע
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

'@����    String2IP
'   ���ַ���IPת��Ϊ��ֵ
'@����ֵ  String
'
'@����:
'strIP String In
'   �ַ���IP
'@��ע
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

'@����    FormatIpString
'   �Էֶβ�����λ��
'@����ֵ  String
'
'@����:
'strIp String In
'   IP��ַ
'@��ע
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

'@����    NormalIpString
'   �Ӹ�ʽ����IP����ԭʼIP
'@����ֵ  String
'
'@����:
'strIp String In
'   IP��ַ
'@��ע
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

'@����    VerFull
'   ����VB���֧�ֵİ汾����ʽ:9999.9999.9999.9999,��С�汾��0000.0000.0000.0000
'@����ֵ  String
'
'@����:
'strVer String In
'   ԭʼ�汾��
'blnMax Boolean In (Optional)
'   True=����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
'@��ע
'
Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '����һ�Σ��Լ�������SP�汾��
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

'@����    IsVerSion
'   �ж��ַ����Ƿ��ǰ汾��
'@����ֵ  Boolean
'
'@����:
'strVer String In
'   ԭʼ�汾��
'blnOnlyCheckSpecial Boolean In
'   ���汾���Ƿ�������SP�汾��
'@��ע
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

'@����    IsEmptyArray
'   �ж϶����Ƿ��ǿ�����
'@����ֵ  Boolean
'
'@����:
'varAnyArray Variant In
'   �жϵ�����
'@��ע
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

'������ܳ���
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '��ȡ�������
    '������ӵ������Ϊ999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) 'ȡ��ĸת���ASCII��
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
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '��ȡ�������
        '������ӵ������Ϊ999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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
        
    '���ݽ��ܵ�����
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
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

'@����    To_DateEx
'   ��ȡORACLE Date���ʹ�
'@����ֵ  String
'   ORACLE Date���ʹ�
'@����:
'strDate String In
'   ʱ���ַ���
'strType String In
'   ��ʽ�ַ������ͣ�ymd-�����գ�yyyy-mm-dd)��ymdhm-��yyyy-mm-dd hh:mm),ymdhms-��yyyy-mm-dd hh:mm:ss)
'@��ע
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

'@����    InCollection
'   ��鼯�����Ƿ����ĳԪ��
'@����ֵ  Boolean
'
'@����:
'cllTest Collection In
'   Ҫ���ļ���
'strKey String In
'   Ҫ����Key
'@��ע
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

'@����    TrimEx
'   ȥ��strTrim���ߵ�strTrmChar,��������Trim
'@����ֵ  String
'
'@����:
'strTrim String In
'   ��Ҫ��ʽ�����ַ�
'strTrmChar String In (Optional,Default=" ")
'   ����strTrmChar���ߴ��ո�ʱ���൱Trim
'@��ע
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


'@����    UboundEx
'   ��ȡ  Ubound
'@����ֵ  Long
'
'@����:
'varArray Variant In
'   ���������
'@��ע
'
Public Function UboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    UboundEx = UBound(varArray)
    Exit Function
ErrH:
    UboundEx = -1
End Function

'@����    LboundEx
'   ��ȡ  Lbound
'@����ֵ  Long
'
'@����:
'varArray Variant In
'   ���������
'@��ע
'
Public Function LboundEx(varArray As Variant) As Long
    On Error GoTo ErrH
    LboundEx = LBound(varArray)
    Exit Function
ErrH:
    LboundEx = 0
End Function
'@����    AppsoftPath
'   ��ȡAPPSOFT·��
'@����ֵ  String
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
'                4��˽�з���
'---------------------------------------------------------------------------
'@����    EncBase64Char
'   ��6-bit�ֽ�ת��ΪBase64�ַ�
'@����ֵ  Byte
'   �ַ���ֵ
'@����:
'bytValue  Byte In
'   ת�����ֽ�
'@��ע
'   Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26����дӢ����ĸ
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26��СдӢ����ĸ
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10������
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function
'@����    DecBase64Char
'   ��Base64�ַ�ת��Ϊ6 bit�ֽ�
'@����ֵ  Byte
'   �ַ���ֵ
'@����:
'bytValue  Byte In
'   ��������ֽ�
'@��ע
'   Base64�ǽ������ֽڣ�ÿ6λ�ָ�Ϊ�ĸ��ֽڴ����
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

'@����    LongToUnsigned
'   ���з���Longת��Ϊ�޷���ֵ
'@����ֵ  Double
'
'@����:
'Value Long In
'   �з���Long
'@��ע
'
Private Function LongToUnsigned(Value As Long) As Double
    Const M_OFFSET_4    As Double = 4294967296#         '�޷������ε����ֵ
    If Value < 0 Then LongToUnsigned = Value + M_OFFSET_4 Else LongToUnsigned = Value
End Function
'---------------------------------------------------------------------------
'                5�����󷽷����¼�
'---------------------------------------------------------------------------



