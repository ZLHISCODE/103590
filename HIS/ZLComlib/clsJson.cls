VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@模块 clsJSON-2019/8/19
'@编写 lshuo
'@功能
'
'@引用
'
'@备注
'   1、在1W以下Scripting.Dictionary速度比VBA.Collection快，且可以获取Key值，因此Json Object 使用Scripting.Dictionary
'   2、数组比VBA.Collection快，因此Json Array使用 Array
'示例
'    Dim objJSON         As New clsJson
'    Dim objJSONArray    As New clsJson, objJSONObject   As New clsJson
'    Dim objJSONTmp      As New clsJson
'    Dim objDic          As New Scripting.Dictionary
'    Dim arrRet          As Variant
'    Dim rsTmp           As ADODB.Recordset
'    '1、构造JSON
'    '..1.1JSON字符串构造
'        objJSON.JSON = "[1,2,3]"        '生成一个JSON Array
'        '查看JSON内容,参数True表示格式化，不传不格式化
'        Debug.Print objJSON.JSON(True)  '"[1,2,3]"
'        objJSON.JSON = "{""A"":1,""B"":2}" '生成一个JSON Object
'        '查看JSON内容
'        Debug.Print objJSON.JSON(True)  '"{""A"":1,""B"":2}"
'    '..1.2数组构造JSON Array
'       objJSON.JSON = Array(1, 2, 3)
'    '..1.3 Scripting.Dictionary构造 JSON Object
'        objDic.Add "A", 1
'        objDic.Add "B", 2
'        Set objJSON.JSON = objDic
'    '..1.4 ADODB.Recordset构造JSON Object
'        'rsTmp通过SQL查询获得数据。
'        'Set objJSON.JSON = rsTmp
'    '..1.5 clsJSON构造JSON
'        objJSONArray.JSON = Array(1, 2, 3)
'        Set objJSON.JSON = objJSONArray  '顶层不是同一个对象，但是内部嵌套的JSON是相同的引用
'    '..1.6产生一个新JSON与原来的JSON没有任何联系
'        Set objJSONTmp = objJSON.Clone
'    '2、JSON常规属性支持
'    '..2.1 获取元素数目
'        Debug.Print objJSON.Count
'    '..2.2 JSON类型判断.未初始化的JSON,IsJsonArray与IsJsonObject均为False
'        Debug.Print objJSON.IsJsonArray
'        Debug.Print objJSON.IsJsonObject
'    '..2.3 JSON成员存在判断
'        Debug.Print objJSON.Exists(4)  '对象为数组，因此用数值
'        '修改为JSON object
'        Set objJSON.JSON = objDic
'        Debug.Print objJSON.Exists("A")
'        Debug.Print objJSON.Exists("C")
'    '..2.4获取数据
'        Set objDic = objJSON.Data   'JSON Object 返回 Scripting.Dictionary
'        '修改数组
'        objJSON.JSON = Array(1, 2, 3)
'        arrRet = objJSON.Data       '返回Variant()
'    '..2.5获取成员名称，只有JSON Object支持
'        Set objJSON.JSON = objDic
'        arrRet = objJSON.keys   '返回Variant()
'    '3、节点操作
'    '..3.1删除节点
'        objJSON.Remove "A"
'        objJSON.JSON = Array(1, 2, 3)
'        objJSON.Remove (0) '节点下标会重新计算
'    '..3.2 新增节点。
'        objJSON(10) = 7
'        Set objJSON.JSON = objDic   '修改为JSON Object
'        objJSON("FFF") = 3
'    '..3.3 新增节点的嵌套问题(支持JSON构造支持的类型)。注意嵌套情况下字符串，只是一个值，尽管符合JSON格式，要使用字符串代表的JSON进行嵌套，请使用ItemJson
'        objJSON("数组节点成员") = Array(1, 2, 3)
'        objJSON.JSON = Array(1, 2, 3)
'        objJSON(10) = Array(1, 2, 3)
'        '字符串嵌套
'        objJSON.JSON = Array(1, 2, 3)
'        objJSON(3) = "[1,2,3]"
'        Debug.Print objJSON.JSON() '[1,2,3,"[1,2,3]"]
'        objJSON.ItemJSON(3) = "[1,2,3]"
'        Debug.Print objJSON.JSON() '[1,2,3,[1,2,3]]
'        'Item属性不支持字符串表达式JSON,而ItemJSON支持字符串的JSON表达式，其余两者完全相同。
'    '..3.4获取成员的JSON表达式。objJSON.Item(3)等价于objJSON(3)。如下两种写法等价
'        Debug.Print objJSON.Item(3).JSON   '[1,2,3]
'        Debug.Print objJSON.Item(3).ItemJSON '[1,2,3]
'    '..3.5根据路径获取节点
'        Set objJSON.JSON = objDic
'        objJSON("C") = Array(4, 5, 6)
'        objJSON.Item("C").Item(3) = Array(7, 8, 9)
'        '获取元素9
'        Debug.Print objJSON.PathItem("C[3][2]")
'        '将9修改嵌套数组
'        objJSON.PathItem("C[3][2]") = Array(11, 12, 13)
'    '4、自定义JSON构造
'    '..4.1空的JSON Object与JSON Array生成,以前有数据则会清掉
'        objJSONArray.EmptyJsonArray
'        objJSONObject.EmptyJsonObject
'    '..4.1自定义对象嵌套，嵌套一个对象
'        objJSONArray.AddJsonObject (0)      'objJSONArray的新增一个成员，该成员下标为0，类型为JSON Object
'        objJSONArray.AddJsonArray (0)       '修改objJSONArray下标为0的元素，该元素类型为JSON Array
'        objJSONObject.AddJsonArray ("第一个成员") 'objJSONObject增加一个成员，该成员名称为"第一个成员"，成员的类型为JSON Array
'        objJSONObject.AddJsonObject ("第二个成员") 'objJSONObject增加一个成员，该成员名称为"第二个成员"，成员的类型为JSON Object
'        objJSONObject.AddJsonObject ("第一个成员") '以前"第一个成员"类型为JSON Array，现在为JSON Object
'    '5、类型转换，JSON、XMLType、ADODB.RecordSet、JSON Path，都必须满足一定的格式，具体参看函数说明
'    '..5.1 JSON<>XMLtype
'        'JSON.FromXmltype,JSON.ToXmltype
'    '..5.1 JSON<>ADODB.RecordSet
'        'JSON.FromRecordset , JSON.ToRecordset
'    '..5.2 JSON<>JSON Path
'        'JSON Path即PathItem的参数。对象成员用"."访问，数组成员用"[索引]"访问。这个转换为JSON Path和JSOn Array 的转换，数组下标为数字类型，成员名为字符串
'        'JSON.FromJSONPath , JSON.ToJSONPath
'    '6、快捷访问、快速构建
'    '..6.1快捷访问
'        'FromJSONPath、FromXmltype、FromRecordset、EmptyJsonArray、EmptyJsonObject、Remove均会返回 当前对象可以连续操作。
'        'AddJsonObject、AddJsonArray，均会返回当前新增对象，方便连续操作
'    '..6.2快速构建
'        objJSON.EmptyJsonObject
'        '根据路径规则快速生成JSON对象以及其嵌套对象
'        objJSON.PathItem("cdss_in.main_info[0].condition_info[0].value_group[0].skey") = 3

'!!!!!!!!!!!!!!!!!!!!!!!!!!
'Item属性是缺省属性，拷贝后请使用类管理器，将该属性设置为缺省属性，否则会出错。
'!!!!!!!!!!!!!!!!!!!!!!!!!!
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
'@功能
'    为标识符指定的区域设置比较两个字符串?
'    如果不正确使用CompareString，可能会损害应用程序的安全性。没有正确比较的字符串可能会产生无效的输入。例如，当用于非语言比较时，函数可能会引发安全问题，因为二进制表示中不同的两个字符串在语言上是等价的。应用程序应该在使用字符串之前测试它们的有效性，并提供错误处理程序。有关更多信息，请参见安全性考虑:国际特性。
'    注意，为了与Unicode兼容，应用程序应该选择CompareStringEx或CompareString的Unicode版本。选择CompareStringEx的另一个原因是，出于互操作性的原因，微软正在向使用地区名称而不是新地区的地区标识符的方向迁移。任何只能在Windows Vista上运行的应用程序都应该使用CompareStringEx。
'@原型
'    int CompareString(
'      LCID    Locale,
'      DWORD   dwCmpFlags,
'      LPCTSTR lpString1,
'      int     cchCount1,
'      LPCTSTR lpString2,
'      int     cchCount2
'    );
'@参数
'Locale
'    用于比较的区域设置的区域设置标识符。您可以使用MAKELCID宏来创建区域设置标识符，或者使用以下预定义值之一。
Private Const LOCALE_CUSTOM_DEFAULT         As Long = &HC00
'    Windows Vista及以后版本:默认的自定义语言环境。当NLS函数必须为当前用户的补充地区返回地区标识符时，该函数将返回此值，而不是LOCALE_USER_DEFAULT。LOCALE_CUSTOM_DEFAULT的值是0x0C00。
Private Const LOCALE_CUSTOM_UI_DEFAULT      As Long = &H1400
'    Windows Vista及以后版本:MUI的默认自定义语言环境。用户首选的UI语言和系统首选的UI语言最多可以包含由语言接口包(LIP)实现的一种语言，语言标识符对应于该语言的补充区域设置。如果列表中有这样一种语言，则常量用于在特定上下文中引用该语言。LOCALE_CUSTOM_UI_DEFAULT的值是0x1400。
Private Const LOCALE_CUSTOM_UNSPECIFIED     As Long = &H1000
'    Windows Vista及以后版本:一个未指定的自定义语言环境，用于标识除当前用户的语言环境外的所有补充语言环境。不能通过区域标识符彼此区分补充区域，但可以通过区域名称来区分。某些NLS函数可以返回此常量，以指示它们不能为特定地区提供有用的标识符。值是0x1000。
Private Const LOCALE_INVARIANT              As Long = &H7F
'    Windows XP:用于操作系统级函数的语言环境，这些函数需要一致的、与语言环境无关的结果。例如，当应用程序使用CompareString函数比较字符串并期望得到一致的结果(无论用户的区域设置是什么)时，使用不变的区域设置。不变区域设置的设置类似于英语(美国)的设置，但不应用于显示格式化数据。通常，应用程序不使用LOCALE_INVARIANT，因为它期望操作的结果依赖于控制每个地区的规则。LOCALE_INVARIANT的值是0x007f。
Private Const LOCALE_SYSTEM_DEFAULT         As Long = &H800
'   操作系统的默认区域设置。这个常量的值是0x0800。
Private Const LOCALE_USER_DEFAULT           As Long = &H400
'    用户或进程的默认区域设置。这个常量的值是0x0400。
'dwCmpFlags
'    指示函数如何比较两个字符串的标志。有关详细定义，此参数可以指定下列值的组合，也可以将其设置为0以获得默认行为。
Private Const LINGUISTIC_IGNORECASE         As Long = &H10
'    忽略大小写，在语言上是适当的。
Private Const LINGUISTIC_IGNOREDIACRITIC    As Long = &H20
'    忽略非间距字符，这在语言上是合适的。
'    注意，当与分解后的字符一起使用时，此标志并不总是产生可预测的结果，即基本字符和一个或多个非间距字符都具有不同的代码点值的字符。
Private Const NORM_IGNORECASE               As Long = &H1
'    忽略大小写。对于许多脚本(尤其是拉丁脚本)，NORM_IGNORECASE与LINGUISTIC_IGNORECASE是一致的。
'    注意，NORM_IGNORECASE忽略了任何第三级区别，不管它是否是语言大小写。例如，在阿拉伯语和印度语脚本中，这区分了字符的不同形式，但这些不同并不对应于语言情况。LINGUISTIC_IGNORECASE使函数只忽略实际的语言大小写，而不是忽略第三个排序权重。
'    注意，通过设置这个标志，函数忽略了CJK兼容性字符的宽格式和窄格式之间的区别。
Private Const NORM_IGNOREKANATYPE           As Long = &H10000
'    不要区分平假名和片假名字符?对应的平假名和片假名字符比较起来是相等的?
Private Const NORM_IGNORENONSPACE           As Long = &H2
'    忽略是非空间的角色。对于许多脚本(尤其是拉丁脚本)，NORM_IGNORENONSPACE与LINGUISTIC_IGNOREDIACRITIC重合。
'    注意，NORM_IGNORENONSPACE忽略了任何次要的区别，不管它是否是一个变音符号。朝鲜语、日语、汉语和印度语等语言的脚本除使用变音符号外，还将此区别用于其他目的。LINGUISTIC_IGNOREDIACRITIC使函数只忽略实际的变音符号，而不是忽略第二个排序权重。
'    注意，NORM_IGNORENONSPACE只对重读字符在主字符的第二次传递中排序的区域具有作用。通常，首先比较字符串中的所有字符而不考虑重音，如果字符串相等，则执行第二次遍历字符串以比较重音。此标志将导致不执行第二遍。对于第一次传递中排序重音字符的区域性，此标志没有作用。
Private Const NORM_IGNORESYMBOLS            As Long = &H4
'    忽略符号和标点符号?
Private Const NORM_IGNOREWIDTH              As Long = &H20000
'    忽略半宽度和全宽度字符之间的差异，例如，cat == cat。全宽格式是中文和日文脚本中使用的格式区别。
Private Const NORM_LINGUISTIC_CASING        As Long = &H8000000
'    使用默认的语言规则，而不是文件系统规则。注意，CompareStringEx的大多数场景都使用这个标志。当应用程序调用CompareStringOrdinal时，不必使用此标志。
Private Const SORT_DIGITSASNUMBERS          As Long = &H8
'    Windows 7:在排序时将数字视为数字，例如，在“10”之前对“2”排序。
Private Const SORT_STRINGSORT               As Long = &H1000
'    把标点符号当作符号来对待?
'lpString1
'    指向要比较的第一个字符串的指针?
'cchCount1
'    由lpString1表示的字符串长度，不包括终止空字符。这个值表示函数的ANSI版本的字节数，Unicode版本的宽字符数。如果字符串以null结尾，应用程序可以提供一个负值。在本例中，函数自动确定长度。
'lpString2
'    指针指向要比较的第二个字符串?
'cchCount2
'    由lpString2表示的字符串长度，不包括终止空字符。这个值表示函数的ANSI版本的字节数，Unicode版本的宽字符数。如果字符串以null结尾，应用程序可以提供一个负值。在本例中，函数自动确定长度。
'@返回值
'    如果成功，返回下列值之一。为了维护C运行时比较字符串的约定，可以从非零返回值中减去值2。然后，<0、==0、>的含义与C运行时一致。
Private Const CSTR_LESS_THAN                As Long = 1
'        CSTR_LESS_THAN?lpString1表示的字符串的词法值小于lpString2表示的字符串?
Private Const CSTR_EQUAL                    As Long = 2
'        CSTR_EQUAL。lpString1表示的字符串在词法值上与lpString2表示的字符串等效。对于排序的目的，这两个字符串是等价的，尽管不一定相同。
Private Const CSTR_GREATER_THAN             As Long = 3
'        CSTR_GREATER_THAN?lpString1所表示的字符串的词法值要大于lpString2所表示的字符串?
'    如果函数没有成功，则返回0。要获得扩展的错误信息，应用程序可以调用GetLastError，它可以返回以下错误代码之一:
'        ERROR_INVALID_FLAGS?为标志提供的值无效?
'        ERROR_INVALID_PARAMETER?任何参数值都无效?
'@备注
'    CompareString和CompareStringEx优化运行在最高速度设置为0或NORM_IGNORECASE dwCmpFlags时,cchCount1 cchCount2设置为1,和地区不支持任何语言按压,当传统的西班牙语排序将“ch”视为一个字符。
'    CompareString和CompareStringEx在比较过程中都忽略了阿拉伯语kashidas。因此，如果两个字符串是相同的，除了kashidas之外，函数返回CSTR_EQUAL。
'    当应用程序使用带有排序函数的NORM_IGNORENONSPACE和NORM_IGNORECASE标志时，这些标志有时会干扰字符串比较。对于不支持非间距字符或大小写，但使用等效权重级别处理其他重要操作的区域设置，可能会出现这种情况。在这种情况下，应用程序应该使用LINGUISTIC_IGNOREDIACRITIC和LINGUISTIC_IGNORECASE标志。这些标志为用例和变音符号标记的代码点排序提供了语言上合适的结果，并且对其他代码点没有影响。
'    从Windows Vista开始:CompareString和CompareStringEx都可以从自定义地区检索数据。从一台计算机到另一台计算机或应用程序运行之间的数据不保证相同。如果应用程序必须持久存储或传输数据，请参阅使用持久区域设置数据。
'    从Windows 8开始:如果你的应用程序将语言标签从Windows传递给这个函数。全球化命名空间，它必须首先通过调用ResolveLocaleName转换标记。
'    从Windows 8开始:CompareStringEx是用stringapi .h声明的。在Windows 8之前，它是在Winnls.h中声明的。
'    注意，排序的行为在Windows版本之间可能会发生变化。例如，可能会创建新的Unicode代码点。使用GetNlsVersionEx查看排序版本是否更改。
'    如果您的应用程序正在调用CompareString的ANSI版本，该函数将通过提供的地区的默认代码页转换参数。因此，应用程序永远不能使用CompareString来处理UTF-8文本。
'    通常，对于不区分大小写的比较，CompareString将小写的“i”映射为大写的“i”，即使语言环境是土耳其语或阿塞拜疆语。The NORM_LINGUISTIC_CASING标志覆盖了土耳其语或阿塞拜疆语的这种行为。如果此标志与土耳其语或阿塞拜疆语一起指定，则拉丁小写字母DOTLESS I (U+0131)是拉丁大写字母I (U+0049)的小写形式，而拉丁小写字母I (U+0069)是拉丁大写字母I的小写形式，上面带有圆点(U+0130)。
'    从Windows 8开始:函数的ANSI版本在Winnls中声明。， Unicode版本是用stringapi .h声明的。在Windows 8之前，这两个版本都是在Winnls.h中声明的。

'字符常量对照
Private Const M_INT_OPEN_BRACKET        As Integer = 91             '左方括号
Private Const M_STR_OPEN_BRACKET        As String = "["             '左方括号
Private Const M_INT_CLOSE_BRACKET       As Integer = 93             '右方括号
Private Const M_STR_CLOSE_BRACKET       As String = "]"             '右方括号
Private Const M_INT_OPEN_BRACE          As Integer = 123            '左大括号
Private Const M_STR_OPEN_BRACE          As String = "{"             '左大括号
Private Const M_INT_CLOSE_BRACE         As Integer = 125            '右大括号
Private Const M_STR_CLOSE_BRACE         As String = "}"             '右大括号

Private Const M_INT_QUOTE               As Integer = 34             '引号
Private Const M_STR_QUOTE               As String = """"            '引号
Private Const M_INT_APOSTROPHE          As Integer = 39             '单引号
Private Const M_STR_APOSTROPHE          As String = "'"             '单引号
Private Const M_INT_COMMA               As Integer = 44             '逗号
Private Const M_STR_COMMA               As String = ","             '逗号
Private Const M_INT_COLON               As Integer = 58             '冒号
Private Const M_STR_COLON               As String = ":"             '冒号
Private Const M_INT_DOT                 As Integer = 46             '点
Private Const M_STR_DOT                 As String = "."             '点
Private Const M_INT_PLUS                As Integer = 43             '加号
Private Const M_STR_PLUS                As String = "+"             '加号
Private Const M_INT_MINUS               As Integer = 45             '减号
Private Const M_STR_MINUS               As String = "-"             '减号
Private Const M_INT_SLASH               As Integer = 47             '斜线
Private Const M_STR_SLASH               As String = "/"             '斜线
Private Const M_INT_BACKSLASH           As Integer = 92             '反斜线
Private Const M_STR_BACKSLASH           As String = "\"             '反斜线
Private Const M_INT_PERCENT             As Integer = 37             '百分号
Private Const M_STR_PERCENT             As String = "%"             '百分号
Private Const M_INT_ASTERISK            As Integer = 42             '*号
Private Const M_STR_ASTERISK            As String = "*"             '*号
Private Const M_INT_AMPERSAND           As Integer = 38             '&符号
Private Const M_STR_AMPERSAND           As String = "&"             '&号
Private Const M_INT_TILDE               As Integer = 126            '~符号
Private Const M_STR_TILDE               As String = "~"             '~号

Private Const M_INT_CHAR_ZERO           As Integer = 0              '字符0
Private Const M_INT_BACK                As Integer = 8              '退格
Private Const M_STR_BACK                As String = vbBack          '退格
Private Const M_INT_TAB                 As Integer = 9              'Tab键
Private Const M_STR_TAB                 As String = vbTab           'Tab字符
Private Const M_INT_FORMFEED            As Integer = 12             '换页
Private Const M_STR_FORMFEED            As String = vbFormFeed      '换页
Private Const M_INT_CR                  As Integer = 13             '回车
Private Const M_STR_CR                  As String = vbCr            '回车
Private Const M_INT_LF                  As Integer = 10             '换行
Private Const M_STR_LF                  As String = vbLf            '换行
Private Const M_INT_SPACE               As Integer = 32             '空格
Private Const M_STR_SPACE               As String = " "             '空格


Private Const M_INT_ZERO                As Integer = 48             '0符号
Private Const M_STR_ZERO                As String = "0"             '0符号
Private Const M_INT_NINE                As Integer = 57             '9符号
Private Const M_STR_NINE                As String = "9"             '9符号

Private Const M_INT_UPPERCASE_A         As Integer = 65             '大写A
Private Const M_STR_UPPERCASE_A         As String = "A"             '大写A
Private Const M_INT_LOWERCASE_A         As Integer = 97             '小写A
Private Const M_STR_LOWERCASE_A         As String = "a"             '小写A
Private Const M_INT_UPPERCASE_B         As Integer = 66             '大写B
Private Const M_STR_UPPERCASE_B         As String = "B"             '大写B
Private Const M_INT_LOWERCASE_B         As Integer = 98             '小写B
Private Const M_STR_LOWERCASE_B         As String = "b"             '小写B
Private Const M_INT_UPPERCASE_E         As Integer = 69             '大写E
Private Const M_STR_UPPERCASE_E         As String = "E"             '大写E
Private Const M_INT_LOWERCASE_E         As Integer = 101            '小写E
Private Const M_STR_LOWERCASE_E         As String = "e"             '小写E
Private Const M_INT_UPPERCASE_F         As Integer = 70             '大写F
Private Const M_STR_UPPERCASE_F         As String = "F"             '大写F
Private Const M_INT_LOWERCASE_F         As Integer = 102            '小写F
Private Const M_STR_LOWERCASE_F         As String = "f"             '小写F
Private Const M_INT_UPPERCASE_H         As Integer = 72             '大写H
Private Const M_STR_UPPERCASE_H         As String = "H"             '大写H
Private Const M_INT_LOWERCASE_H         As Integer = 104            '小写H
Private Const M_STR_LOWERCASE_H         As String = "h"             '小写H
Private Const M_INT_UPPERCASE_L         As Integer = 76             '大写L
Private Const M_STR_UPPERCASE_L         As String = "L"             '大写L
Private Const M_INT_LOWERCASE_L         As Integer = 108            '小写L
Private Const M_STR_LOWERCASE_L         As String = "l"             '小写L
Private Const M_INT_UPPERCASE_N         As Integer = 78             '大写N
Private Const M_STR_UPPERCASE_N         As String = "N"             '大写N
Private Const M_INT_LOWERCASE_N         As Integer = 110            '小写N
Private Const M_STR_LOWERCASE_N         As String = "n"             '小写N
Private Const M_INT_UPPERCASE_R         As Integer = 82             '大写R
Private Const M_STR_UPPERCASE_R         As String = "R"             '大写R
Private Const M_INT_LOWERCASE_R         As Integer = 114            '小写R
Private Const M_STR_LOWERCASE_R         As String = "r"             '小写R
Private Const M_INT_UPPERCASE_S         As Integer = 83             '大写S
Private Const M_STR_UPPERCASE_S         As String = "S"             '大写S
Private Const M_INT_LOWERCASE_S         As Integer = 115            '小写S
Private Const M_STR_LOWERCASE_S         As String = "s"             '小写S
Private Const M_INT_UPPERCASE_T         As Integer = 84             '大写T
Private Const M_STR_UPPERCASE_T         As String = "T"             '大写T
Private Const M_INT_LOWERCASE_T         As Integer = 116            '小写T
Private Const M_STR_LOWERCASE_T         As String = "t"             '小写T
Private Const M_INT_UPPERCASE_U         As Integer = 85             '大写U
Private Const M_STR_UPPERCASE_U         As String = "U"             '大写U
Private Const M_INT_LOWERCASE_U         As Integer = 117            '小写U
Private Const M_STR_LOWERCASE_U         As String = "u"             '小写U
Private Const M_INT_UPPERCASE_X         As Integer = 88             '大写X
Private Const M_STR_UPPERCASE_X         As String = "X"             '大写X
Private Const M_INT_LOWERCASE_X         As Integer = 120            '小写X
Private Const M_STR_LOWERCASE_X         As String = "x"             '小写X
Private Const M_INT_UPPERCASE_Z         As Integer = 90             '大写Z
Private Const M_STR_UPPERCASE_Z         As String = "Z"             '大写Z
Private Const M_INT_LOWERCASE_Z         As Integer = 122            '小写Z
Private Const M_STR_LOWERCASE_Z         As String = "z"             '小写Z

Private Const M_LOWER_GREATER_UPPER     As Integer = &H20           '小写比大写大32

Private Const JSON_PREFIX           As String = "__json"
Private Const JSON_ARRAY            As String = JSON_PREFIX & "array"
Private Const JSON_TYPE             As String = JSON_PREFIX & "type"
Private Const JSON_NAME             As String = JSON_PREFIX & "name"
'用于XML转换
Private Enum JSONDataType
    JT_String = 0
    JT_Null = 1
    JT_Bool = 2
    JT_Number = 3
    JT_Array = 4
    JT_Object = 5
End Enum

Private Const FADF_AUTO      As Integer = &H1   'An array that is allocated on the stack.
Private Const FADF_FIXEDSIZE As Integer = &H10  'An array that may not be resized or reallocated.
'地址宽度
Private Const PTR_SIZE                  As Long = 4
'符号位
Private Const SIGN_BIT                  As Long = &H80000000
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
'JSON字符解析,不用截取字符串，加快解析速度
Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

Private Type JsonContext
    StrictMode          As Boolean
    Text()              As Integer
    Pos                 As Long
    Error               As String
    LastChar            As Integer
    TextArray           As SAFEARRAY1D
    StringBuilder       As clsStringBuilder
End Type
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mvarMembers             As Variant
Private Const M_INT_ARRAY_SZIE  As Long = 128
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'@方法    JSON
'   返回JSON字符串
'@返回值  Variant
'
'@参数:
'blnWhitespaceIndent In Boolean
'   是否对JSON进行缩进格式化
'blnUnicode In Boolean Optional
'   是否对多字节字符转义。多字节会转换成\uXXXX
'lngLevel In Long Optional
'   当前的JSON对象层数，每增加一层，增加一倍strWhiteSpace,该参数不需要传递，也不能传递，仅供内部使用。
'objString In clsStringBuilder Optional
'   用来获取JSON内容的字符串缓冲类，不需要传递。也可以传递一个进来，不会清除该类已经包含的内容
'@备注
'
Public Property Get JSON(Optional ByVal blnWhitespaceIndent As Boolean, Optional ByVal blnUnicode As Boolean, Optional objString As clsStringBuilder, Optional ByVal lngLevel As Long) As Variant
    Dim jsContent           As JsonContext  '方便字符串转换

    With jsContent
        '方便字符串转换
        With .TextArray
            .cDims = 1
            .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            .cbElements = 2
            .cLocks = 1
            .lLbound = 0
        End With
    End With
    If objString Is Nothing Then
        Set objString = New clsStringBuilder
    End If
    If Not IsEmpty(mvarMembers) Then
        Call Sub_SerializeValue(Me, lngLevel, objString, blnWhitespaceIndent, blnUnicode, jsContent)
    End If
    JSON = objString.Value
    '不清理，若有接收可以方便外面使用
    'Set objString = Nothing
End Property

'@方法    JSON
'   解析JSON生成JSON对象
'@返回值  String
'
'@参数:
'strWhiteSpace In String
'   不需要该参数，仅为了保持属性参数一致
'blnUnicode In Boolean Optional
'   不需要该参数，仅为了保持属性参数一致
'lngLevel In Long Optional
'   不需要该参数，仅为了保持属性参数一致
'objString In clsStringBuilder Optional
'   不需要该参数，仅为了保持属性参数一致
'varJsonSource In Variant
'   JSON数据源，接受String、clsJSON、AdoDB.RecordSet
'@备注
'
Public Property Let JSON(Optional ByVal blnWhitespaceIndent As Boolean, Optional ByVal blnUnicode As Boolean, Optional objString As clsStringBuilder, Optional ByVal lngLevel As Long, ByRef varJsonSource As Variant)
    Const FUNC_NAME         As String = "JSON[Let]"
    Dim jsContent           As JsonContext, i       As Long, varKeys    As Variant, varValue    As Variant, varTmp  As Variant
    Dim rsTmp               As ADODB.Recordset
    On Error GoTo ErrH
    Select Case varType(varJsonSource)
        Case vbString
            With jsContent
                Set .StringBuilder = New clsStringBuilder
                With .TextArray
                    .cDims = 1
                    .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
                    .cbElements = 2
                    .cLocks = 1
                    .lLbound = 0
                    .pvData = StrPtr(varJsonSource)         '字符串和数组共享相同的数据
                    .cElements = Len(varJsonSource) + 1     '字符串的最后一个0字符展示出来，因此需要+1，+2的话会最后一个元素不属于字符串的地址,可能不为0
                End With
                
                RtlMoveMemory ByVal ArrPtr(.Text), VarPtr(.TextArray), PTR_SIZE  '数组描述符移动过去，方便进行数组操作，数据没有移动，数组和传入的字符串指向相同的地址，只是读取方式不同，数组按2个字节读取。字符串当做整个对象
                Sub_JsonParse jsContent, lngLevel
            End With
        Case Is >= vbArray
            mvarMembers = Array()
            If UBound(varJsonSource) >= 0 Then
                ReDim Preserve mvarMembers(UBound(varJsonSource))
                For i = 0 To UBound(varJsonSource)
                    Select Case varType(varJsonSource(i))
                        Case vbObject, Is > vbArray '对象或者数组，则交给下一层解析
                            If IsObject(varJsonSource(i)) Then
                                If TypeOf varValue Is clsJson Then
                                    Set varTmp = varJsonSource(i)
                                Else
                                    Set varTmp = New clsJson
                                    varTmp.JSON = varJsonSource(i)
                                End If
                            Else
                                Set varTmp = New clsJson
                                varTmp.JSON = varJsonSource(i)
                            End If
                            Set mvarMembers(i) = varTmp
                        Case Else
                            mvarMembers(i) = varJsonSource(i)
                    End Select
                Next
            End If
        Case Else
            RaiseJsonError ERROR_JSON_CREATE, FUNC_NAME
    End Select
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Set JSON(Optional ByVal blnWhitespaceIndent As Boolean, Optional ByVal blnUnicode As Boolean, Optional objString As clsStringBuilder, Optional ByVal lngLevel As Long, ByRef varJsonSource As Object)
    Const FUNC_NAME         As String = "JSON[Set]"
    Dim jsContent           As JsonContext, i       As Long, varKeys    As Variant, varValue    As Variant, varTmp  As Variant
    Dim rsTmp               As ADODB.Recordset
    Dim objDic              As Scripting.Dictionary
    On Error GoTo ErrH
    If Not varJsonSource Is Nothing Then
        If TypeOf varJsonSource Is clsJson Then
            Call Sub_AssignVariant(mvarMembers, varJsonSource.Data)
        ElseIf TypeOf varJsonSource Is ADODB.Recordset Then
            Set rsTmp = varJsonSource
            Call FromRecordset(rsTmp)
        ElseIf TypeOf varJsonSource Is Scripting.Dictionary Then
            Set objDic = varJsonSource
            varKeys = objDic.keys
            Set mvarMembers = New Scripting.Dictionary
            For i = 0 To UBound(varKeys)
                Sub_AssignVariant varValue, objDic(varKeys(i))
                Select Case varType(varValue)
                    Case vbObject, Is > vbArray '对象或者数组，则交给下一层解析
                        If IsObject(varValue) Then
                            If TypeOf varValue Is clsJson Then
                                Set varTmp = varValue
                            Else
                                Set varTmp = New clsJson
                                varTmp.JSON = varValue
                            End If
                        Else
                            Set varTmp = New clsJson
                            varTmp.JSON = varValue
                        End If
                        Set mvarMembers(varKeys(i)) = varTmp
                    Case Else
                        mvarMembers(varKeys(i)) = varValue
                End Select
            Next
        Else
            RaiseJsonError ERROR_JSON_CREATE, FUNC_NAME
        End If
    Else
        mvarMembers = Empty
    End If
    Exit Property
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'Count R:获取元素个数
Public Property Get Count() As Long
    If IsJsonObject Then
        Count = mvarMembers.Count
    ElseIf IsEmpty(mvarMembers) Then
        Count = 0
    Else
        Count = UBound(mvarMembers) + 1
    End If
End Property
'Exists R:某个元素是否存在
Public Property Get Exists(ByVal varMemberID As Variant) As Boolean
    Const FUNC_NAME     As String = "Exists[Get]"
    Dim vtTypeOfID      As VbVarType
    On Error GoTo ErrH
    If IsNull(varMemberID) Or IsEmpty(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NULL, FUNC_NAME
    vtTypeOfID = VBA.varType(varMemberID)
    If vtTypeOfID = vbString Then
        If IsJsonArray Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        Exists = mvarMembers.Exists(varMemberID)
    Else
        If IsJsonObject Then RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
        If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        varMemberID = CLng(varMemberID)
        If varMemberID >= 0 Then
            If IsEmpty(mvarMembers) Then
                Exists = False
            Else
                Exists = varMemberID <= UBound(mvarMembers)
            End If
        End If
    End If
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property
'IsJsonArray R:当前对象是否是Json Array
Public Property Get IsJsonArray() As Boolean
    If Not IsEmpty(mvarMembers) Then
        IsJsonArray = Not VBA.IsObject(mvarMembers)
    End If
End Property
'IsJsonObject R:当前对象是否是Json Object
Public Property Get IsJsonObject() As Boolean
    If Not IsEmpty(mvarMembers) Then
        IsJsonObject = VBA.IsObject(mvarMembers)
    End If
End Property
'Keys R:获取所有Key
Public Property Get keys() As Variant
    Const FUNC_NAME         As String = "Keys[Get]"
    If IsJsonObject Then
        keys = mvarMembers.keys()
    Else
        RaiseJsonError ERROR_JSON_KEY_NOT_OBJECT, FUNC_NAME
    End If
End Property

'Data R:返回数据对象
Public Property Get Data() As Variant
    Sub_AssignVariant Data, mvarMembers
End Property

'PathItem R/W:获取或者设置指定成员。接收String以及非嵌套的JSON Array
Public Property Get PathItem(ByVal varPath As Variant) As Variant
    Const FUNC_NAME     As String = "PathItem[Get]"
    Dim objJsonPath     As clsJson
    Dim i               As Long
    Dim varData         As Variant
    Dim varTmp          As Variant
    On Error GoTo ErrH
    If varType(varPath) = vbString Then
        Set objJsonPath = New clsJson
        Call objJsonPath.FromJsonPath(varPath)
    ElseIf TypeOf varPath Is clsJson Then
        Set objJsonPath = varPath
    Else
        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
    End If
    varData = objJsonPath.Data
    Set varTmp = Me
    For i = 0 To UBound(varData)
        If IsObject(varData(i)) Then
            RaiseJsonError ERROR_JSON_PATH_NESTED, FUNC_NAME
        Else
            Call Sub_AssignVariant(varTmp, varTmp(varData(i)))
        End If
    Next
    Call Sub_AssignVariant(PathItem, varTmp)
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Let PathItem(ByVal varPath As Variant, ByVal varValue As Variant)
    Const FUNC_NAME  As String = "PathItem[Let]"
    Dim objJsonPath     As New clsJson, objTmp      As clsJson, varKey      As Variant
    Dim i               As Long
    Dim varData         As Variant
    Dim varTmp          As Variant
    
    On Error GoTo ErrH
    If varType(varPath) = vbString Then
        Set objJsonPath = New clsJson
        Call objJsonPath.FromJsonPath(CStr(varPath))
    ElseIf TypeOf varPath Is clsJson Then
        Set objJsonPath = varPath
    Else
        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
    End If
    varData = objJsonPath.Data
    Set objTmp = Me
    For i = 0 To UBound(varData)
        If IsObject(varData(i)) Then
            RaiseJsonError ERROR_JSON_PATH_NESTED, FUNC_NAME
        Else
            If i = 0 Then
                If Not Me.IsJsonArray And Not Me.IsJsonObject Then
                    If varType(varData(i)) = vbString Then
                        Call objTmp.EmptyJsonObject
                    ElseIf IsNumeric(varData(i)) Then
                        Call objTmp.EmptyJsonArray
                    Else
                        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
                    End If
                End If
            End If
            If varType(varData(i)) = vbString Then
                If Not objTmp.IsJsonObject Then
                    RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
                End If
            ElseIf IsNumeric(varData(i)) Then
                If Not objTmp.IsJsonArray Then
                    RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
                End If
            Else
                RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
            End If
            Call Sub_AssignVariant(varTmp, objTmp(varData(i)))
            If i <> UBound(varData) Then
                If Not IsObject(varTmp) Then
                    If varType(varData(i + 1)) = vbString Then
                        Set objTmp = objTmp.AddJsonObject(varData(i))
                    ElseIf IsNumeric(varData(i + 1)) Then
                        Set objTmp = objTmp.AddJsonArray(varData(i))
                    Else
                        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
                    End If
                Else
                    Set objTmp = varTmp
                End If
            End If
            varKey = varData(i)
        End If
    Next
    If Not IsEmpty(varKey) Then
        objTmp(varKey) = varValue
    End If
    Exit Property
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Set PathItem(ByVal varPath As Variant, ByVal varValue As Object)
    Const FUNC_NAME  As String = "PathItem[Set]"
    Dim objJsonPath     As New clsJson, objTmp      As clsJson, varKey      As Variant
    Dim i               As Long
    Dim varData         As Variant
    Dim varTmp          As Variant
    
    On Error GoTo ErrH
    If varType(varPath) = vbString Then
        Set objJsonPath = New clsJson
        Call objJsonPath.FromJsonPath(CStr(varPath))
    ElseIf TypeOf varPath Is clsJson Then
        Set objJsonPath = varPath
    Else
        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
    End If
    varData = objJsonPath.Data
    Set objTmp = Me
    For i = 0 To UBound(varData)
        If IsObject(varData(i)) Then
            RaiseJsonError ERROR_JSON_PATH_NESTED, FUNC_NAME
        Else
            If i = 0 Then
                If Not Me.IsJsonArray And Not Me.IsJsonObject Then
                    If varType(varData(i)) = vbString Then
                        Call objTmp.EmptyJsonObject
                    ElseIf IsNumeric(varData(i)) Then
                        Call objTmp.EmptyJsonArray
                    Else
                        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
                    End If
                End If
            End If
            If varType(varData(i)) = vbString Then
                If Not objTmp.IsJsonObject Then
                    RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
                End If
            ElseIf IsNumeric(varData(i)) Then
                If Not objTmp.IsJsonArray Then
                    RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
                End If
            Else
                RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
            End If
            Call Sub_AssignVariant(varTmp, objTmp(varData(i)))
            If i <> UBound(varData) Then
                If Not IsObject(varTmp) Then
                    If varType(varData(i + 1)) = vbString Then
                        Set objTmp = objTmp.AddJsonObject(varData(i))
                    ElseIf IsNumeric(varData(i + 1)) Then
                        Set objTmp = objTmp.AddJsonArray(varData(i))
                    Else
                        RaiseJsonError ERROR_JSON_PATH_NOT_SUPPORT, FUNC_NAME
                    End If
                Else
                    Set objTmp = varTmp
                End If
            End If
            varKey = varData(i)
        End If
    Next
    If Not IsEmpty(varKey) Then
        Set objTmp(varKey) = varValue
    End If
    Exit Property
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'Item R/W:获取或者设置指定成员
Public Property Get Item(ByVal varMemberID As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Const FUNC_NAME         As String = "Item[Get]"
    Dim vtTypeOfID          As VbVarType
    On Error GoTo ErrH
    If IsNull(varMemberID) Or IsEmpty(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NULL, FUNC_NAME
    vtTypeOfID = VBA.varType(varMemberID)
    If IsEmpty(mvarMembers) Then
        If vtTypeOfID = vbString Then
            Set mvarMembers = New Scripting.Dictionary
        Else
            If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
            mvarMembers = Array()
        End If
        Item = Null
        Exit Property
    End If
    
    If IsJsonObject Then
        If vtTypeOfID <> vbString Then RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
    Else
        If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        varMemberID = CLng(varMemberID)
        If varMemberID < 0 Then RaiseJsonError ERROR_JSON_KEY_INDEXED_ZERO, FUNC_NAME
        If varMemberID > UBound(mvarMembers) Then
            Item = Null
            Exit Property
        End If
    End If
    Sub_AssignVariant Item, mvarMembers(varMemberID)
    If VBA.IsEmpty(Item) Then Item = Null
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Let Item(ByVal varMemberID As Variant, ByVal varValue As Variant)
    Const FUNC_NAME         As String = "Item[Let]"
    Dim vtTypeOfID          As VbVarType
    Dim objJSON             As clsJson
    
    On Error GoTo ErrH
    If IsNull(varMemberID) Or IsEmpty(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NULL, FUNC_NAME
    vtTypeOfID = VBA.varType(varMemberID)
    If IsEmpty(mvarMembers) Then
        If vtTypeOfID = vbString Then
            Set mvarMembers = New Scripting.Dictionary
        Else
            If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
            mvarMembers = Array()
        End If
    End If
    
    If IsJsonObject Then
        If vtTypeOfID <> vbString Then RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
    Else
        If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        varMemberID = CLng(varMemberID)
        If varMemberID < 0 Then RaiseJsonError ERROR_JSON_KEY_INDEXED_ZERO, FUNC_NAME
        If varMemberID > UBound(mvarMembers) Then ReDim Preserve mvarMembers(varMemberID)
    End If
    
    Select Case VBA.varType(varValue)
        Case vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal, vbByte
            mvarMembers(varMemberID) = CDbl(varValue)
        Case vbString, vbNull, vbEmpty, vbBoolean
            mvarMembers(varMemberID) = varValue
        Case vbDate
            mvarMembers(varMemberID) = Format(varValue, "yyyy-mm-dd hh:mm:ss")
        Case Is > vbArray
            Set objJSON = New clsJson
            objJSON.JSON = varValue
            Set mvarMembers(varMemberID) = objJSON
        Case Else
            RaiseJsonError ERROR_JSON_CREATE, FUNC_NAME
    End Select
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Set Item(ByVal varMemberID As Variant, ByVal varValue As Object)
    Const FUNC_NAME         As String = "Item[Set]"
    Dim vtTypeOfID          As VbVarType
    Dim objJSON             As clsJson
    
    On Error GoTo ErrH
    If IsNull(varMemberID) Or IsEmpty(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NULL, FUNC_NAME
    vtTypeOfID = VBA.varType(varMemberID)
    If IsEmpty(mvarMembers) Then
        If vtTypeOfID = vbString Then
            Set mvarMembers = New Scripting.Dictionary
        Else
            If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
            mvarMembers = Array()
        End If
    End If
    
    If IsJsonObject Then
        If vtTypeOfID <> vbString Then RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
    Else
        If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        varMemberID = CLng(varMemberID)
        If varMemberID < 0 Then RaiseJsonError ERROR_JSON_KEY_INDEXED_ZERO, FUNC_NAME
        If varMemberID > UBound(mvarMembers) Then ReDim Preserve mvarMembers(varMemberID)
    End If
    If Not varValue Is Nothing Then
        If TypeOf varValue Is clsJson Then
            Set objJSON = varValue
        Else
            Set objJSON = New clsJson
            objJSON.JSON = varValue
        End If
        Set mvarMembers(varMemberID) = objJSON
    Else
        mvarMembers(varMemberID) = Null
    End If
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'ItemJSON R/W:使用JSON字符串初始化一个成员，或者JSON对象成员的字符串表达式
Public Property Get ItemJSON(Optional ByVal varMemberID As Variant) As Variant
    Const FUNC_NAME         As String = "ItemJSON[Get]"
    Dim varValue            As Variant
    On Error GoTo ErrH
    
    If IsMissing(varMemberID) Then
        Set varValue = Me
    Else
        Sub_AssignVariant varValue, Me(varMemberID)
    End If
    If IsObject(varValue) Then
        ItemJSON = varValue.JSON
    Else
        ItemJSON = CStr(varValue)
    End If
    Exit Property
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Let ItemJSON(ByVal varMemberID As Variant, varJSON As Variant)
    Const FUNC_NAME         As String = "ItemJSON[Let]"
    Dim objJSON             As New clsJson
    
    On Error GoTo ErrH
    objJSON.JSON = varJSON
    Set mvarMembers(varMemberID) = objJSON
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Set ItemJSON(ByVal varMemberID As Variant, varJSON As Object)
    Const FUNC_NAME         As String = "ItemJSON[Set]"
    Dim objJSON             As New clsJson
    
    On Error GoTo ErrH
    Set objJSON.JSON = varJSON
    Set mvarMembers(varMemberID) = objJSON
    Exit Property
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property
'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------

'@方法    EmptyJsonObject
'   将当前对象JSON类型设置为JSONObject,原有数据会清除
'@返回值
'
'@参数:
'@备注
'
Public Function EmptyJsonObject() As clsJson
    Set mvarMembers = New Scripting.Dictionary
    Set EmptyJsonObject = Me
End Function

'@方法    EmptyJsonArray
'   将当前对象JSON类型设置为JSONArray,原有数据会清除
'@返回值
'
'@参数:
'@备注
'
Public Function EmptyJsonArray() As clsJson
    mvarMembers = Array()
    Set EmptyJsonArray = Me
End Function

'@方法    AddJsonArray
'   添加一个Json Array 到指定对象
'@返回值  clsJson
'
'@参数:
'varMemberID In Variant
'   成员名称或索引
'@备注
'
Public Function AddJsonArray(ByVal varMemberID As Variant) As clsJson
    Set AddJsonArray = New clsJson
    Call AddJsonArray.EmptyJsonArray
    Set Item(varMemberID) = AddJsonArray
End Function

'@方法    AddJsonObject
'   添加一个Json Object 到指定对象
'@返回值  clsJson
'
'@参数:
'varMemberID In Variant
'   成员名称或索引
'@备注
'
Public Function AddJsonObject(ByVal varMemberID As Variant) As clsJson
    Set AddJsonObject = New clsJson
    Call AddJsonObject.EmptyJsonObject
    Set Item(varMemberID) = AddJsonObject
End Function
'@方法    Remove
'   从JSON对象中移除一个成员
'@返回值
'
'@参数:
'varMemberID In Variant
'   移除的成员名或者索引
'@备注
'
Public Function Remove(ByVal varMemberID As Variant) As clsJson
    Const FUNC_NAME      As String = "Remove"
    Dim vtTypeOfID      As VbVarType
    Dim i               As Long
    On Error GoTo ErrH
    If IsNull(varMemberID) Or IsEmpty(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NULL, FUNC_NAME
    If IsEmpty(mvarMembers) Then Exit Function
    vtTypeOfID = VBA.varType(varMemberID)
    If IsJsonObject Then
        If vtTypeOfID <> vbString Then RaiseJsonError ERROR_JSON_KEY_NOT_STRING, FUNC_NAME
        mvarMembers.Remove varMemberID
    Else
        If Not IsNumeric(varMemberID) Then RaiseJsonError ERROR_JSON_KEY_NOT_NUMBER, FUNC_NAME
        varMemberID = CLng(varMemberID)
        If varMemberID < 0 Then RaiseJsonError ERROR_JSON_KEY_INDEXED_ZERO, FUNC_NAME
        If varMemberID > UBound(mvarMembers) Then Exit Function
        If UBound(mvarMembers) = 0 Then
            mvarMembers = Array()
        Else
            For i = varMemberID To UBound(mvarMembers) - 1
                Sub_AssignVariant mvarMembers(i), mvarMembers(i + 1)
            Next
            ReDim Preserve mvarMembers(UBound(mvarMembers) - 1)
        End If
    End If
    Set Remove = Me
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@方法    Clone
'   复制一个全新的JSON对象和原JSON没有任何关联
'@返回值  clsJson
'
'@参数:
'@备注
'
Public Function Clone() As clsJson
    On Error GoTo ErrH
    Set Clone = Sub_Clone(Me)
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@方法    ToXmltype
'   将JSON转换为XML对象。若要获取XML字符串，.xml属性可以 直接获取
'@返回值  Object
'
'@参数:
'@备注
'   1、具有根节点（root）
'   2、数组用节点（__jsonarray）包裹,
'   3、节点原始名称为数字，则属性__jsonname的值为节点名称
'   4、节点增加类型__jsontype说明。0-String(可省略)。1、NULL，节点值为1，2-Bool,节点值0代表False,1代表True.3、数字
'"{""less"":[true,false,null,""测试"",{""TEST"":""AAAA""}],""More"":123}"
'<root><less><__jsonarray><__A0 __jsontype="2">1</__A0><__A1 __jsontype="2">0</__A1><__A2 __jsontype="1">1</__A2><__A3>测试</__A3><__A4><TEST>AAAA</TEST></__A4></__jsonarray></less><More __jsontype="3">123</More></root>

Public Function ToXmltype() As Object
    Dim objXmlRoot As Object, objXmlDoc As Object
    On Error GoTo ErrH
    Set objXmlDoc = CreateObject("MSXML2.DOMDocument")
    objXmlDoc.appendChild objXmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-16""")
    Set objXmlRoot = objXmlDoc.appendChild(objXmlDoc.createElement("root"))
    Call Sub_ToXmltype(Me, objXmlRoot, objXmlDoc)
    Set ToXmltype = objXmlRoot
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'@方法    FromXmltype
'   将XML转换为JSON
'@返回值
'
'@参数:
'varXml In Variant
'   XML对象或XML字符串
'@备注
'   1、必须具有根节点（root）
'   2、数组必须用节点（__jsonarray）包裹,
'   3、节点原始名称为数字，则属性__jsonname的值为节点名称
'   4、可以对节点增加类型__jsontype说明。0-String(可省略)。1、NULL，节点值为1，2-Bool,节点值0代表False,1代表True.3、数字
Public Function FromXmltype(varXml As Variant) As clsJson
    Dim objRoot     As Object
    Dim objDoc      As Object
    Dim strKey      As String, objNode          As Object
    Dim varItem     As Variant, lngIndex        As Long, varKey        As Variant
    Dim objJSON     As clsJson
    Dim blnArray    As Boolean
    
    On Error GoTo ErrH
    If IsObject(varXml) Then
        Set objRoot = varXml
    Else
        Set objDoc = CreateObject("MSXML2.DOMDocument")
        objDoc.loadXML CStr(varXml)
        Set objRoot = objDoc.documentElement
        If objRoot Is Nothing Then
            Exit Function
        End If
    End If
    
    If objRoot.firstChild Is Nothing Then
        Call EmptyJsonObject
    Else
        If objRoot.childNodes.Length = 1 And objRoot.firstChild.nodeName = JSON_ARRAY Then
            Call EmptyJsonArray
            blnArray = True
            Set objRoot = objRoot.firstChild
        Else
            Call EmptyJsonObject
        End If
    End If
    If blnArray Then
        If objRoot.childNodes.Length > 0 Then
            ReDim Preserve mvarMembers(objRoot.childNodes.Length - 1)
        End If
    End If

    For Each objNode In objRoot.childNodes
        If Not IsNull(objNode.getAttribute(JSON_NAME)) Then
            varKey = CStr(objNode.getAttribute(JSON_NAME))
        Else
            varKey = CStr(objNode.nodeName)
        End If
        If blnArray Then varKey = lngIndex
        Select Case Val(objNode.getAttribute(JSON_TYPE) & "")
            Case JT_Null
                mvarMembers(varKey) = Null
            Case JT_Number
                mvarMembers(varKey) = CDbl(objNode.nodeTypedValue)
            Case JT_Bool
                mvarMembers(varKey) = CBool(objNode.nodeTypedValue)
            Case JT_String
                
                If Not objNode.firstChild Is Nothing Then
                    If objNode.firstChild.nodeType = 1 Then ' NODE_ELEMENT Then
                        Set objJSON = New clsJson
                        objJSON.FromXmltype objNode
                        Set mvarMembers(varKey) = objJSON
                    Else
                        mvarMembers(varKey) = objNode.Text
                    End If
                Else
                     mvarMembers(varKey) = objNode.Text
                End If
        End Select
        lngIndex = lngIndex + 1
    Next
    Set FromXmltype = Me
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@方法    ToRecordset
'   将JSON转换成记录集
'@返回值  ADODB.Recordset
'
'@参数:
'varSchema In Variant
'   记录集的定义，一个JSON格式
'  [{"name":"Value", "type":"number"},
'  {"name":"Name", "type":"string"}]

'  [{"name":"Value", "type":"number"},
'  {"name":"Name", "type":"string"}]
'@备注
'   只有如下的格式的JSON才支持转换。
'   JSON格式。0-每行是一个对象，所有的组成一个数组。[{"F1":1,"F2":True},{"F1":1,"F2":True}]
'   1-JSON分类两个成员，Name和Data，格式为{"name":["F1","F2"],"data":[[1,True],[1,True]]}
Public Function ToRecordset(Optional varSchema As Variant) As ADODB.Recordset
    Dim objJsonSchema       As clsJson
    Dim i                   As Long, j                      As Long
    Dim rsReturn            As New ADODB.Recordset
    Dim objJsonPath         As New clsJson
    Dim objJSONTmpName      As clsJson, objJSONTmpData      As clsJson, objJsonFull As clsJson
    Dim varKeys             As Variant, varValue            As Variant, arrData     As Variant
    
    
    
    If Not (IsEmpty(varSchema) Or IsMissing(varSchema)) Then
        Set objJsonSchema = New clsJson
        If IsObject(varSchema) Then
            Set objJsonSchema.JSON = varSchema
        Else
            objJsonSchema.JSON = varSchema
        End If
    End If
    
    If Me.Count > 0 Then
        Set rsReturn = New ADODB.Recordset
        
        With ToRecordset
            If IsJsonArray Then
                Set objJsonFull = Me(0)
            Else
                Set objJSONTmpName = Me("name")
                Set objJSONTmpData = Me("data")
                If objJSONTmpData.Count = 0 Then
                    Exit Function
                End If
                Set objJSONTmpData = objJSONTmpData(0)
                Call objJsonFull.EmptyJsonObject
                For i = 0 To objJSONTmpData.Count
                    objJsonFull(objJSONTmpName(i)) = objJSONTmpData(i)
                Next
            End If
            varKeys = objJsonFull.keys
            For i = 0 To UBound(varKeys)
                Call Sub_AssignVariant(varValue, objJsonFull(varKeys(i)))
                Select Case varType(varValue)
                    Case vbString
                        .Fields.Append varKeys(i), adLongVarChar, 2000
                    Case vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal, vbByte
                        .Fields.Append varKeys(i), adNumeric, 38
                    Case Else
                        .Fields.Append varKeys(i), adLongVarChar, 2000
                End Select
            Next
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
            If IsJsonArray Then
                For i = 0 To Me.Count - 1
                    arrData = Me(i).Data
                    .AddNew
                    For j = 0 To UBound(arrData)
                        .Fields(j).Value = arrData(j)
                    Next
                Next
            Else
                Set objJSONTmpData = Me("data")
                For i = 0 To objJSONTmpData.Count - 1
                    arrData = objJSONTmpData(i).Data
                    .AddNew
                    For j = 0 To UBound(arrData)
                        .Fields(j).Value = arrData(j)
                    Next
                Next
            End If
        End With
    End If
    Set ToRecordset = rsReturn
End Function

'@方法    FromRecordset
'   将Recordset转换为JSON
'@返回值
'
'@参数:
'rsInput In ADODB.Recordset
'   要转换的记录集
'intJsonStyle In Integer
'   JSON格式。0-每行是一个对象，所有的组成一个数组。[{"F1":1,"F2":True},{"F1":1,"F2":True}]
'   1-JSON分类两个成员，Name和Data，格式为{"name":["F1","F2"],"data":[[1,True],[1,True]]}
'@备注

Public Function FromRecordset(rsInput As ADODB.Recordset, Optional intJsonStyle As Integer) As clsJson
    Dim objJSON     As clsJson, lngCount    As Long
    
    On Error GoTo ErrH
    If intJsonStyle = 0 Then
        Call EmptyJsonArray
        If rsInput Is Nothing Then Exit Function
        If rsInput.State = adStateClosed Then Exit Function
        If Not rsInput.EOF Then
            ReDim Preserve mvarMembers(rsInput.RecordCount - 1)
            Do While Not rsInput.EOF
                Set objJSON = New clsJson
                Call objJSON.Sub_FromRecordset_AddOneRow(rsInput, intJsonStyle)
                Set mvarMembers(lngCount) = objJSON
                lngCount = lngCount + 1
                rsInput.MoveNext
            Loop
            If lngCount = 0 Then
                Call EmptyJsonArray
            Else
                If UBound(mvarMembers) + 1 > lngCount Then
                    ReDim Preserve mvarMembers(lngCount - 1)
                End If
            End If
        End If
    Else
        Call EmptyJsonObject
        If rsInput Is Nothing Then Exit Function
        If rsInput.State = adStateClosed Then Exit Function
        Set objJSON = New clsJson
        Call objJSON.Sub_FromRecordset_AddName(rsInput, intJsonStyle)
        Set mvarMembers("name") = objJSON
        Set objJSON = New clsJson
        Call objJSON.Sub_FromRecordset_AddData(rsInput, intJsonStyle)
        Set mvarMembers("data") = objJSON
    End If
    Set FromRecordset = Me
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@方法    FromJsonPath
'
'@返回值
'
'@参数:
'strString In String
'   JsonPath表达式，成员用.元素索引用[i]
'@备注
'
Public Function FromJsonPath(ByVal strString As String) As clsJson
    Dim jsContent           As JsonContext
    Dim lngIndex            As Long
    Dim blnString           As Boolean, varValue    As String
    On Error GoTo ErrH
    Call EmptyJsonArray
    With jsContent
        Set .StringBuilder = New clsStringBuilder
        With .TextArray
            .cDims = 1
            .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            .cbElements = 2
            .cLocks = 1
            .lLbound = 0
            .pvData = StrPtr(strString)         '字符串和数组共享相同的数据
            .cElements = Len(strString) + 1     '字符串的最后一个0字符展示出来，因此需要+1，+2的话会最后一个元素不属于字符串的地址,可能不为0
        End With
        
        RtlMoveMemory ByVal ArrPtr(.Text), VarPtr(.TextArray), PTR_SIZE  '数组描述符移动过去，方便进行数组操作，数据没有移动，数组和传入的字符串指向相同的地址，只是读取方式不同，数组按2个字节读取。字符串当做整个对象
        
        .StringBuilder.AppendOneChar M_INT_OPEN_BRACKET
        Do
            .LastChar = .Text(.Pos)
            .Pos = .Pos + 1
            Select Case .LastChar
                Case M_INT_DOT
                    If .StringBuilder.Length > 1 Then .StringBuilder.AppendOneChar M_INT_COMMA
                    .StringBuilder.AppendOneChar M_INT_QUOTE
                    Do
                        .LastChar = .Text(.Pos)
                        .Pos = .Pos + 1
                        Select Case .LastChar
                            Case M_INT_DOT, M_INT_OPEN_BRACKET, M_INT_CHAR_ZERO
                                .Pos = .Pos - 1
                                .StringBuilder.AppendOneChar M_INT_QUOTE
                                Exit Do
                            Case Else
                                .StringBuilder.AppendOneChar .LastChar
                        End Select
                    Loop
                Case M_INT_OPEN_BRACKET
                    If .StringBuilder.Length > 1 Then .StringBuilder.AppendOneChar M_INT_COMMA
                    Do
                        .LastChar = .Text(.Pos)
                        .Pos = .Pos + 1
                        Select Case .LastChar
                            Case M_INT_CLOSE_BRACKET
                                Exit Do
                            Case M_INT_CHAR_ZERO
                                .Pos = .Pos - 1
                                Exit Do
                            Case Else
                                .StringBuilder.AppendOneChar .LastChar
                        End Select
                    Loop
                Case M_INT_CHAR_ZERO
                    Exit Do
                Case Else
                    .StringBuilder.AppendOneChar M_INT_QUOTE
                    .StringBuilder.AppendOneChar .LastChar
                    Do
                        .LastChar = .Text(.Pos)
                        .Pos = .Pos + 1
                        Select Case .LastChar
                            Case M_INT_DOT, M_INT_OPEN_BRACKET, M_INT_CHAR_ZERO
                                .Pos = .Pos - 1
                                .StringBuilder.AppendOneChar M_INT_QUOTE
                                Exit Do
                            Case Else
                                .StringBuilder.AppendOneChar .LastChar
                        End Select
                    Loop
            End Select
        Loop
        .StringBuilder.AppendOneChar M_INT_CLOSE_BRACKET
        Me.JSON = .StringBuilder.Value
    End With
    Set FromJsonPath = Me
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@方法    ToJsonPath
'   将JSONPath转换为字符表达式
'@返回值
'
'@参数:
'@备注
'
Public Function ToJsonPath() As String
    Const FUNC_NAME     As String = "ToJsonPath"
    Dim i               As Long
    Dim objString       As New clsStringBuilder
    On Error GoTo ErrH
    If IsJsonArray Then
        For i = 0 To UBound(mvarMembers)
            Select Case varType(mvarMembers(i))
                Case vbString
                    objString.Append CStr(mvarMembers(i))
                Case vbDouble
                    Call objString.Append("[").Append(CStr(mvarMembers(i))).Append("]")
                Case Else
                    RaiseJsonError ERROR_JSON_PATH_ITEM, FUNC_NAME
            End Select
        Next
        ToJsonPath = objString.Value
    Else
        RaiseJsonError ERROR_JSON_PATH_NESTED, FUNC_NAME
    End If
    Exit Function
ErrH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------
'@方法    Sub_JsonParse
'   Json解析的实现方法体
'@返回值
'
'@参数:
'jsContent In/Out JsonContext
'   JSON解析的结构
'lngLevel In Long Optional
'   JSON递归层级
'@备注
'
Friend Sub Sub_JsonParse(jsContent As JsonContext, Optional ByVal lngLevel As Long)
    Const FUNC_NAME      As String = "Sub_JsonParse"
    With jsContent
        Select Case Sub_JsonGetNextChar(jsContent)
            Case M_INT_OPEN_BRACE '右大括号,{
                Set mvarMembers = New Scripting.Dictionary
                Sub_ParseObject jsContent
            Case M_INT_OPEN_BRACKET '左方括号，[
                mvarMembers = Array()
                Sub_ParseArray jsContent
            Case M_INT_CHAR_ZERO
                If lngLevel = 0 Then RaiseJsonError ERROR_JSON_SYNTAX_JSON_START, FUNC_NAME, .Pos
            Case Else
                RaiseJsonError ERROR_JSON_SYNTAX_JSON_START, FUNC_NAME, .Pos
        End Select
    End With
End Sub
'@方法    Sub_JsonGetNextChar
'   获取下一个非空白、非注释的符号。
'@返回值  Integer
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'@备注
'
Private Function Sub_JsonGetNextChar(jsContent As JsonContext) As Integer
    Dim lngIdx          As Long
    Dim i               As Long

    With jsContent
        Do
            .LastChar = .Text(.Pos)
            .Pos = .Pos + 1
            Select Case .LastChar
                Case M_INT_CHAR_ZERO
                    Exit Function
                Case M_INT_TAB, M_INT_CR, M_INT_LF, M_INT_SPACE '空白符。vbTab, vbCr, vbLf, " "
                    '不做处理
                Case M_INT_SLASH '/ 注释
                    Select Case .Text(.Pos)
                        Case M_INT_SLASH        '单行注释。//
                            .Pos = .Pos + 1
                            Do
                                .LastChar = .Text(.Pos)
                                .Pos = .Pos + 1
                                If .LastChar = M_INT_CHAR_ZERO Then
                                    Exit Function
                                End If
                            Loop While Not (.LastChar = M_INT_LF Or .LastChar = M_INT_CR)  '遇到回车换行终止
                        Case M_INT_ASTERISK     '多行注释。/*
                            lngIdx = .Pos + 1
                            Do
                                .LastChar = .Text(lngIdx)
                                lngIdx = lngIdx + 1
                                If .LastChar = 0 Then
                                    .Pos = lngIdx - 1
                                    .Error = "在位置" & .Pos & "的注释,缺失注释终止符"
                                    Exit Function
                                End If
                            Loop While Not (.LastChar = M_INT_ASTERISK And .Text(lngIdx) = M_INT_SLASH) '多行注释结束。*/,由于最后一个是0字符，因此不用管理越界
                            .LastChar = .Text(lngIdx)
                            .Pos = lngIdx + 1
                        Case Else
                            Sub_JsonGetNextChar = .LastChar
                            Exit Do
                    End Select
                Case Else
                    Sub_JsonGetNextChar = .LastChar
                    Exit Do
            End Select
        Loop
    End With
End Function

'@方法    Sub_ParseObject
'   解析一个 JSON Object
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'lngLevel In Long Optional
'   JSON递归层级
'@备注
'
Private Function Sub_ParseObject(jsContent As JsonContext, Optional ByVal lngLevel As Long) As clsJson
    Const FUNC_NAME      As String = "Sub_ParseObject"
    Dim strName         As String
    Dim varValue        As Variant
    Dim lngCOMMA        As Long
    With jsContent
        Do
            Select Case Sub_JsonGetNextChar(jsContent)
                Case M_INT_APOSTROPHE, M_INT_QUOTE  ''、",成员名称开始
                    If lngCOMMA > 0 Then RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_MEMBER, FUNC_NAME, .Pos
                    strName = vbNullString
                    strName = Sub_ParseString(jsContent, .LastChar)
                    If LenB(strName) > 0 Then
                        If Sub_JsonGetNextChar(jsContent) = M_INT_COLON Then    ':成员值
                            Sub_AssignVariant varValue, Sub_ParseValue(jsContent, lngLevel)
                        Else
                            RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_KEY_VALUE, FUNC_NAME, .Pos
                            Exit Do
                        End If
                        If IsObject(varValue) Then
                            Set mvarMembers(strName) = varValue
                        Else
                            mvarMembers(strName) = varValue
                        End If
                        lngCOMMA = lngCOMMA + 1
                    Else
                        RaiseJsonError ERROR_JSON_SYNTAX_NULL_NAME, FUNC_NAME, .Pos
                        Exit Do
                    End If
                Case M_INT_COMMA            ',成员分隔符
                    If lngCOMMA = 0 Then
                        RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_START, FUNC_NAME, .Pos
                    Else
                        lngCOMMA = lngCOMMA - 1
                        If lngCOMMA > 0 Then
                            RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_MULTI, FUNC_NAME, .Pos
                            Exit Do
                        End If
                    End If
                Case M_INT_CLOSE_BRACE      '右大括号}
                    Exit Do
                Case M_INT_CHAR_ZERO
                    RaiseJsonError ERROR_JSON_SYNTAX_OBJECT_END, FUNC_NAME, .Pos
                    Exit Do
                Case Else
                    RaiseJsonError ERROR_JSON_SYNTAX_OBJECT_END, FUNC_NAME, .Pos
                    Exit Do
            End Select
        Loop
    End With
End Function
'@方法    Sub_ParseArray
'   解析一个 JSON Array
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'lngLevel In Long Optional
'   JSON递归层级
'@备注
'
Private Function Sub_ParseArray(jsContent As JsonContext, Optional ByVal lngLevel As Long) As clsJson
    Const FUNC_NAME      As String = "Sub_ParseObject"
    Dim varValue        As Variant
    Dim lngCOMMA        As Long, lngCount        As Long
    '先分配100个空间
    ReDim Preserve mvarMembers(M_INT_ARRAY_SZIE)
    With jsContent
        Do
            Select Case Sub_JsonGetNextChar(jsContent)
                Case M_INT_COMMA            ',成员分隔符
                    If lngCOMMA = 0 Then
                        RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_START, FUNC_NAME, .Pos
                    Else
                        lngCOMMA = lngCOMMA - 1
                        If lngCOMMA > 0 Then
                            RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_MULTI, FUNC_NAME, .Pos
                            Exit Do
                        End If
                    End If
                Case M_INT_CLOSE_BRACKET      '右中括号]
                    Exit Do
                Case M_INT_CHAR_ZERO
                    RaiseJsonError ERROR_JSON_SYNTAX_ARRAY_END, FUNC_NAME, .Pos
                    Exit Do
                Case Else
                    If lngCOMMA > 0 Then RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_MEMBER, FUNC_NAME, .Pos
                    .Pos = .Pos - 1
                    '每次增长原来长度的0.5倍
                    If lngCount > UBound(mvarMembers) Then ReDim Preserve mvarMembers(lngCount + lngCount \ 2)
                    Sub_AssignVariant mvarMembers(lngCount), Sub_ParseValue(jsContent)
                    lngCOMMA = lngCOMMA + 1
                    lngCount = lngCount + 1
            End Select
        Loop
    End With
    If lngCount = 0 Then
        Call EmptyJsonArray
    Else
        If UBound(mvarMembers) + 1 > lngCount Then
            ReDim Preserve mvarMembers(lngCount - 1)
        End If
    End If
End Function
'@方法    Sub_ParseValue
'   解析一个 JSON Value
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'lngLevel In Long Optional
'   JSON递归层级
'@备注
'
Private Function Sub_ParseValue(jsContent As JsonContext, Optional ByVal lngLevel As Long) As Variant
    Const FUNC_NAME      As String = "Sub_ParseValue"
    Dim objJSON         As clsJson
    
    With jsContent
        Select Case Sub_JsonGetNextChar(jsContent)
            Case M_INT_OPEN_BRACE, M_INT_OPEN_BRACKET '右大括号,{,左方括号，[,传递给子对象解析。
                .Pos = .Pos - 1
                Set objJSON = New clsJson
                Call objJSON.Sub_JsonParse(jsContent, lngLevel + 1)
                Set Sub_ParseValue = objJSON
            Case M_INT_QUOTE, M_INT_APOSTROPHE    '字符串开始"，'
                Sub_ParseValue = Sub_ParseString(jsContent, .LastChar)
            Case M_INT_UPPERCASE_T, M_INT_LOWERCASE_T 'T,t  True
                Sub_ParseValue = Sub_ParseBoolean(jsContent, True)
            Case M_INT_UPPERCASE_F, M_INT_LOWERCASE_F 'F,f false
                Sub_ParseValue = Sub_ParseBoolean(jsContent, False)
            Case M_INT_UPPERCASE_N, M_INT_LOWERCASE_N   'N,n null
                Sub_ParseValue = Sub_ParseNull(jsContent)
            Case M_INT_ZERO To M_INT_NINE, M_INT_DOT, M_INT_PLUS, M_INT_MINUS '0-9 + - .        '数字
                Sub_ParseValue = Sub_ParseNumber(jsContent)
            Case Else
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "数值", .Pos
        End Select
   End With
End Function
'@方法    Sub_ParseString
'   解析一个 获取一个字符串
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'intEndChar In Long Optional
'   字符串以何种符号结束，开始单引号，结束必须单引号（'）。开始引号，结束必须引号（"）
'@备注
'
Private Function Sub_ParseString(jsContent As JsonContext, ByVal intEndChar As Integer) As String
    Const FUNC_NAME      As String = "Sub_ParseString"
    Dim lngIdx          As Long
    Dim blnError        As Integer

    With jsContent
        For lngIdx = .Pos To UBound(.Text)
            Select Case .Text(lngIdx)
                Case intEndChar             '字符串终结"
                    .Pos = lngIdx + 1
                    .LastChar = intEndChar
                    Sub_ParseString = .StringBuilder.Value
                    Call .StringBuilder.Clear
                    Exit For
                Case M_INT_BACKSLASH        '转义字符。\
                    lngIdx = lngIdx + 1
                    Select Case .Text(lngIdx)
                        Case M_INT_LOWERCASE_B      'b
                            .StringBuilder.AppendOneChar M_INT_BACK
                        Case M_INT_LOWERCASE_F      'f
                            .StringBuilder.AppendOneChar M_INT_FORMFEED
                        Case M_INT_LOWERCASE_N      'n
                            .StringBuilder.AppendOneChar M_INT_LF
                        Case M_INT_LOWERCASE_T      't
                            .StringBuilder.AppendOneChar M_INT_TAB
                        Case M_INT_LOWERCASE_R      'r
                            .StringBuilder.AppendOneChar M_INT_CR
                        Case M_INT_QUOTE     '"
                            .StringBuilder.AppendOneChar M_INT_QUOTE
                        Case M_INT_APOSTROPHE       ''
                            .StringBuilder.AppendOneChar M_INT_APOSTROPHE
                        Case M_INT_SLASH            '/
                            .StringBuilder.AppendOneChar M_STR_SLASH
                        Case M_INT_BACKSLASH        '\
                            .StringBuilder.AppendOneChar M_INT_BACKSLASH
                        Case M_INT_LOWERCASE_U      'u
                            If lngIdx + 4 <= UBound(.Text) Then
                                If .StringBuilder.AppendOneCharBy4Hex(.Text(lngIdx + 1), .Text(lngIdx + 2), .Text(lngIdx + 3), .Text(lngIdx + 4)) Is Nothing Then
                                    RaiseJsonError ERROR_JSON_SYNTAX_ESCAPE_U, FUNC_NAME, .Pos
                                    Exit For
                                End If
                                lngIdx = lngIdx + 4
                            Else
                                RaiseJsonError ERROR_JSON_SYNTAX_ESCAPE_U, FUNC_NAME, .Pos
                                Exit For
                            End If
                        Case M_INT_LOWERCASE_X      'x
                            If lngIdx + 2 <= UBound(.Text) Then
                                If .StringBuilder.AppendOneCharBy2Hex(.Text(lngIdx + 1), .Text(lngIdx + 2)) Is Nothing Then
                                    RaiseJsonError ERROR_JSON_SYNTAX_ESCAPE_X, FUNC_NAME, .Pos
                                    Exit For
                                End If
                                lngIdx = lngIdx + 2
                            Else
                                RaiseJsonError ERROR_JSON_SYNTAX_ESCAPE_X, FUNC_NAME, .Pos
                                Exit For
                            End If
                        Case Else
                            RaiseJsonError ERROR_JSON_SYNTAX_ESCAPE, FUNC_NAME, .Pos
                            Exit For
                    End Select
                Case M_INT_CHAR_ZERO
                    RaiseJsonError ERROR_JSON_SYNTAX_SPLIT_STR_END, FUNC_NAME, ChrW(intEndChar), .Pos
                    Exit For
                Case Else
                    .StringBuilder.AppendOneChar .Text(lngIdx)
            End Select
        Next
    End With
End Function
'@方法    Sub_ParseString
'   解析一个 获取一个NuLL符号
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'@备注
'
Private Function Sub_ParseNull(jsContent As JsonContext) As Variant
    Const FUNC_NAME      As String = "Sub_ParseNull"
    With jsContent
        If UBound(.Text) >= .Pos + 2 Then           '越界检查
            If (.Text(.Pos + 0) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_U Then    '--- u
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "NULL", .Pos + 0
            End If
            If (.Text(.Pos + 1) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_L Then    '--- l
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "NULL", .Pos + 1
            End If
            If (.Text(.Pos + 2) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_L Then    '--- l
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "NULL", .Pos + 2
            End If
            .Pos = .Pos + 3
            Sub_ParseNull = Null
        Else
            RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "NULL", .Pos
        End If
    End With
End Function
'@方法    Sub_ParseString
'   解析一个 获取一个Boolean符号
'@返回值
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'blnExpectedValue In Boolean Optional
'   预期获取的True或者False
'@备注
'
Private Function Sub_ParseBoolean(jsContent As JsonContext, Optional ByVal blnExpectedValue As Boolean) As Boolean
    Const FUNC_NAME      As String = "Sub_ParseBoolean"
    With jsContent
        If blnExpectedValue Then
            If UBound(.Text) >= .Pos + 2 Then           '越界检查
                If (.Text(.Pos + 0) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_R Then    'r
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "True", .Pos + 0
                End If
                If (.Text(.Pos + 1) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_U Then    'u
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "True", .Pos + 1
                End If
                If (.Text(.Pos + 2) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_E Then    'e
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "True", .Pos + 2
                End If
                .Pos = .Pos + 3
                Sub_ParseBoolean = True
            Else
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "True", .Pos
            End If
        Else
            If UBound(.Text) >= .Pos + 3 Then           '越界检查
                If (.Text(.Pos + 0) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_A Then     '--- a
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "False", .Pos + 0
                End If
                If (.Text(.Pos + 1) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_L Then    '--- l
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "False", .Pos + 1
                End If
                If (.Text(.Pos + 2) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_S Then    '--- s
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "False", .Pos + 2
                End If
                If (.Text(.Pos + 3) Or M_LOWER_GREATER_UPPER) <> M_INT_LOWERCASE_E Then    '--- e
                    RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "False", .Pos + 3
                End If
                .Pos = .Pos + 4
                Sub_ParseBoolean = False
            Else
                RaiseJsonError ERROR_JSON_SYNTAX_UNEXPECTED_SYMBOL, FUNC_NAME, "False", .Pos
            End If
        End If
    End With
End Function

'@方法    Sub_JsonGetNumber
'   解析一个数字类型
'@返回值  Double
'
'@参数:
'jsContent In JsonContext
'   JSON解析结构
'@备注
'
Private Function Sub_ParseNumber(jsContent As JsonContext) As Double
    Dim i           As Long

    With jsContent
        If .Pos <= UBound(.Text) Then
            If .LastChar = M_INT_ZERO And (.Text(.Pos) Or M_LOWER_GREATER_UPPER) = M_INT_LOWERCASE_X Then   '0x
                .StringBuilder.AppendOneChar(M_INT_AMPERSAND).Append (M_INT_UPPERCASE_H)     '&H
                For i = .Pos + 1 To UBound(.Text)
                    Select Case .Text(i)
                        Case M_INT_ZERO To M_INT_NINE, M_INT_LOWERCASE_A To M_INT_LOWERCASE_F, M_INT_UPPERCASE_A To M_INT_UPPERCASE_F '0-9,a-f
                            .StringBuilder.AppendOneChar .Text(i)
                        Case Else
                            Exit For
                    End Select
                Next
            Else
                .StringBuilder.AppendOneChar .LastChar
                For i = .Pos To UBound(.Text)
                    Select Case .Text(i)
                        Case M_INT_ZERO To M_INT_NINE, M_INT_PLUS, M_INT_MINUS, M_INT_DOT, M_INT_LOWERCASE_E, M_INT_UPPERCASE_E '0-9 + - . e E
                            .StringBuilder.AppendOneChar .Text(i)
                        Case Else
                            Exit For
                    End Select
                Next
            End If
        End If
        Sub_ParseNumber = CDbl(.StringBuilder.Value)
        .Pos = jsContent.Pos + .StringBuilder.Length - 1
        Call .StringBuilder.Clear
    End With
End Function

'@方法    Sub_AssignVariant
'   处理可变类型赋值问题
'@返回值
'
'@参数:
'varDest In Variant
'   赋值目标
'varSrc In Variant
'   要赋值的值
'@备注
'
Private Sub Sub_AssignVariant(varDest As Variant, varSrc As Variant)
    On Error GoTo ErrH
    If IsObject(varSrc) Then
        Set varDest = varSrc
    Else
        varDest = varSrc
    End If
ErrH:
End Sub

'@方法    Sub_SerializeValue
'   将JSON对象序列化成字符串，不返回，直接写缓冲区。
'@返回值
'
'@参数:
'varValue In Variant
'   序列化的值
'lngLevel In Long
'   序列化的层级
'objString In clsStringBuilder
'   字符串缓冲区
'strWhiteSpace In String
'   格式化字符串
'blnUnicode In Boolean
'   对多字节字符使用ANS转换，转为为\uXXXX
'jsContent In JsonContext
'   方便字符串解析
'@备注
Private Sub Sub_SerializeValue(varValue As Variant, ByVal lngLevel As Long, objString As clsStringBuilder, ByVal blnWhitespaceIndent As Boolean, ByVal blnUnicode As Boolean, jsContent As JsonContext)
    Dim arrKeys     As Variant, arrData             As Variant
    Dim i           As Long
    Select Case VBA.varType(varValue)
        Case vbEmpty, vbNull
            Call objString.Append("null")
        Case vbBoolean
            If varValue Then
                Call objString.Append("true")
            Else
                Call objString.Append("false")
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Call objString.Append(CStr(varValue))
        Case vbString
            Call Sub_SerializeString(varValue, objString, blnUnicode, jsContent)
        Case vbObject
            Sub_AssignVariant arrData, varValue.Data
            If IsObject(arrData) Then
                Call objString.AppendOneChar(M_INT_OPEN_BRACE)
                If arrData.Count < 1 Then
                    Call objString.AppendOneChar(M_INT_CLOSE_BRACE) '}
                Else
                    arrKeys = arrData.keys
                    Call objString.AppendOneChar(M_INT_LF)
                    For i = LBound(arrKeys) To UBound(arrKeys)
                        'Key:value
                        Call objString.AppendOneChar(M_INT_SPACE, -blnWhitespaceIndent * lngLevel * 2)
                        Call Sub_SerializeValue(arrKeys(i), lngLevel, objString, blnWhitespaceIndent, blnUnicode, jsContent)
                        Call objString.AppendOneChar(M_INT_COLON)
                        Call Sub_SerializeValue(arrData(arrKeys(i)), lngLevel + 1, objString, blnWhitespaceIndent, blnUnicode, jsContent)
                        If i < UBound(arrKeys) Then
                            Call objString.AppendOneChar(M_INT_COMMA).AppendOneChar(M_INT_LF) ',
                        Else
                            Call objString.AppendOneChar(M_INT_LF) '
                        End If
                    Next
                    Call objString.AppendOneChar(M_INT_SPACE, -blnWhitespaceIndent * lngLevel * 2).AppendOneChar(M_INT_CLOSE_BRACE) '}
                End If
            ElseIf Not IsEmpty(arrData) Then
                Call objString.AppendOneChar(M_INT_OPEN_BRACKET)
                If UBound(arrData) < 0 Then
                    Call objString.AppendOneChar(M_INT_CLOSE_BRACKET) ']
                Else
                    Call objString.AppendOneChar(M_INT_LF)
                    For i = 0 To UBound(arrData)
                        Call objString.AppendOneChar(M_INT_SPACE, -blnWhitespaceIndent * lngLevel * 2)
                        Call Sub_SerializeValue(arrData(i), lngLevel + 1, objString, blnWhitespaceIndent, blnUnicode, jsContent) 'value
                        If i < UBound(arrData) Then
                            Call objString.AppendOneChar(M_INT_COMMA).AppendOneChar(M_INT_LF) ',
                        Else
                            Call objString.AppendOneChar(M_INT_LF) '
                        End If
                    Next
                    Call objString.AppendOneChar(M_INT_SPACE, -blnWhitespaceIndent * lngLevel * 2).AppendOneChar(M_INT_CLOSE_BRACKET) ']
                End If
            End If
    End Select
End Sub

'@方法    Sub_SerializeString
'   序列化字符串
'@返回值
'
'@参数:
'varValue In Variant
'   序列化的值
'objString In clsStringBuilder
'   字符串缓冲区
'blnUnicode In Boolean
'   对多字节字符使用ANS转换，转为为\uXXXX
'jsContent In JsonContext
'   方便字符串解析
'@备注
'
Private Sub Sub_SerializeString(varValue As Variant, objString As clsStringBuilder, ByVal blnUnicode As Boolean, jsContent As JsonContext)

    objString.AppendOneChar M_INT_QUOTE  '"
    If LenB(varValue) > 0 Then
        With jsContent
            .Pos = 0
            '方便字符串转换
            With .TextArray
                .pvData = StrPtr(varValue)         '字符串和数组共享相同的数据
                .cElements = Len(varValue) + 1     '字符串的最后一个0字符展示出来，因此需要+1，+2的话会最后一个元素不属于字符串的地址,可能不为0
            End With
            
            RtlMoveMemory ByVal ArrPtr(.Text), VarPtr(.TextArray), PTR_SIZE  '数组描述符移动过去，方便进行数组操作，数据没有移动，数组和传入的字符串指向相同的地址，只是读取方式不同，数组按2个字节读取。字符串当做整个对象
            Do
                .LastChar = .Text(.Pos)
                .Pos = .Pos + 1
                Select Case .LastChar
                    Case M_INT_CHAR_ZERO To M_INT_TILDE
                        Select Case .LastChar
                            Case M_INT_QUOTE, M_INT_APOSTROPHE, M_INT_BACKSLASH, M_INT_SLASH '"、'、\、/
                               Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(.LastChar)
                            Case M_INT_CHAR_ZERO
                                Exit Do
                            Case M_INT_BACK
                                Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(M_INT_LOWERCASE_B)  '\b
                            Case M_INT_FORMFEED
                                Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(M_INT_LOWERCASE_F)  '\f
                            Case M_INT_LF
                                Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(M_INT_LOWERCASE_N)  '\n
                            Case M_INT_CR
                                Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(M_INT_LOWERCASE_R)  '\r
                            Case M_INT_TAB
                                Call objString.AppendOneChar(M_INT_BACKSLASH).AppendOneChar(M_INT_LOWERCASE_T)  '\t
                            Case Is < M_INT_SPACE
                                Call objString.AppendOneCharUnicode4(.LastChar)
                            Case Else
                                Call objString.AppendOneChar(.LastChar)
                        End Select
                    '127-159(256),
                    Case &H7F To &H100, &H34F, &H200B To &H200F, &H2028 To &H202E, &H2060, &HFE01 To &HFE0F, &HFEFF, &HFFFD, &HD800 To &HDFFF
                        Call objString.AppendOneCharUnicode4(.LastChar)
                    Case Else
                        If blnUnicode Then
                            Call objString.AppendOneCharUnicode4(.LastChar)
                        Else
                            Call objString.AppendOneChar(.LastChar)
                        End If
                End Select
            Loop
        End With
    End If
    objString.AppendOneChar M_INT_QUOTE  '"
End Sub

'@方法    Sub_ToXmltype
'   将JSON转换为XMLType的实现方法
'@返回值  Object
'
'@参数:
'varValue Variant In
'   转换的对象或值
'objXmlRoot Object In
'   转换的根节点
'objXmlDoc Object In
'   转换的文档对象
'@备注
'
Private Function Sub_ToXmltype(varValue As Variant, Optional objXmlRoot As Object, Optional objXmlDoc As Object) As Object
    Dim objTmp              As Object, objItem      As Object
    Dim lngIndex            As Long, lngCount       As Long
    Dim varItem             As Variant
    Dim arrKeys             As Variant, arrData     As Variant
    Dim i                   As Long
    
    Select Case VBA.varType(varValue)
        Case vbEmpty, vbNull
            objXmlRoot.setAttribute JSON_TYPE, JT_Null
            objXmlRoot.nodeTypedValue = 1
        Case vbBoolean
            objXmlRoot.setAttribute JSON_TYPE, JT_Bool
            objXmlRoot.nodeTypedValue = -varValue
        Case vbDate
            objXmlRoot.nodeTypedValue = Format(varValue, "yyyy-mm-dd hh:mm:ss")
        Case vbObject
            Sub_AssignVariant arrData, varValue.Data
            If IsObject(arrData) Then
                arrKeys = arrData.keys
                For i = LBound(arrKeys) To UBound(arrKeys)
                    Sub_AssignVariant varItem, arrData(arrKeys(i))
                    If IsNumeric(arrKeys(i)) Or LenB(arrKeys(i)) = 0 Then
                        '名称另存的节点
                        Set objItem = objXmlRoot.appendChild(objXmlDoc.createElement("__K" & i))
                        objItem.setAttribute JSON_NAME, arrKeys(i)
                    Else
                        Set objItem = objXmlRoot.appendChild(objXmlDoc.createElement(arrKeys(i)))
                    End If
                    
                    If IsObject(varItem) Then
                        If varItem.IsJsonArray Then
                            Set objItem = objItem.appendChild(objXmlDoc.createElement(JSON_ARRAY))
                        End If
                    End If
                    Call Sub_ToXmltype(varItem, objItem, objXmlDoc)
                Next
            ElseIf Not IsEmpty(arrData) Then
                Set objTmp = objXmlRoot
                If objTmp Is objXmlDoc.documentElement Then
                    Set objTmp = objTmp.appendChild(objXmlDoc.createElement(JSON_ARRAY))
                End If
                For i = 0 To UBound(arrData)
                    Sub_AssignVariant varItem, arrData(i)
                    If objItem Is Nothing Then
                        Set objItem = objTmp.appendChild(objXmlDoc.createElement("__A" & i))
                        If IsObject(varItem) Then
                            If varItem.IsJsonArray Then
                                Set objItem = objItem.appendChild(objXmlDoc.createElement(JSON_ARRAY))
                            End If
                        End If
                    End If
                    Call Sub_ToXmltype(varItem, objItem, objXmlDoc)
                    Set objItem = Nothing
                Next
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            objXmlRoot.setAttribute JSON_TYPE, JT_Number
            objXmlRoot.nodeTypedValue = varValue
        Case Else
            objXmlRoot.nodeTypedValue = varValue
    End Select
End Function

'@方法    Sub_Clone
'   Clone方法的实现函数
'@返回值  Variant
'
'@参数:
'varValue In Variant
'   复制的对象
'@备注
'
Private Function Sub_Clone(varValue As Variant) As clsJson
    Dim arrKeys     As Variant, arrData             As Variant
    Dim objRet      As clsJson, varTmpValue         As Variant, arrTmpRet      As Variant
    Dim i           As Long
    If VBA.varType(varValue) = vbObject Then
        Set objRet = New clsJson
        Sub_AssignVariant arrData, varValue.Data
        If IsObject(arrData) Then
            Call objRet.EmptyJsonObject
            arrKeys = keys()
            For i = LBound(arrKeys) To UBound(arrKeys)
                'Key:value
                Call Sub_AssignVariant(varTmpValue, arrData(arrKeys(i)))
                If IsObject(varTmpValue) Then
                    Set objRet(arrKeys(i)) = varTmpValue.Clone
                Else
                    objRet(arrKeys(i)) = varTmpValue
                End If
            Next
        ElseIf Not IsEmpty(arrData) Then
            Call objRet.EmptyJsonArray
            For i = 0 To UBound(arrData)
                Call Sub_AssignVariant(varTmpValue, arrData(i))
                If IsObject(varTmpValue) Then
                    Set objRet(i) = varTmpValue.Clone
                Else
                    objRet(i) = varTmpValue
                End If
            Next
        End If
        Set Sub_Clone = objRet
    End If
End Function

'@方法    Sub_FromRecordset_AddName
'   FromRecordset子函数，用来生成列头信息
'@返回值
'
'@参数:
'@备注
'
Friend Sub Sub_FromRecordset_AddName(rsInput As ADODB.Recordset, Optional intJsonStyle As Integer)
    Dim i       As Long
    
    Call EmptyJsonArray
    With rsInput
        ReDim Preserve mvarMembers(.Fields.Count - 1)
        For i = 0 To UBound(mvarMembers)
            mvarMembers(i) = .Fields(i).Name
        Next
    End With
End Sub

'@方法    Sub_FromRecordset_AddData
'   FromRecordset子函数，用来添加数据
'@返回值
'
'@参数:
'Param 方向
'
'@备注
'
Friend Sub Sub_FromRecordset_AddData(rsInput As ADODB.Recordset, Optional intJsonStyle As Integer)
    Dim objJSON     As clsJson, lngCount    As Long
    
    If intJsonStyle = 1 Then
        Call EmptyJsonArray
        ReDim Preserve mvarMembers(rsInput.RecordCount - 1)
        Do While Not rsInput.EOF
            Set objJSON = New clsJson
            Call objJSON.Sub_FromRecordset_AddOneRow(rsInput, intJsonStyle)
            Set mvarMembers(lngCount) = objJSON
            lngCount = lngCount + 1
            rsInput.MoveNext
        Loop
        If lngCount = 0 Then
            Call EmptyJsonArray
        Else
            If UBound(mvarMembers) + 1 > lngCount Then
                ReDim Preserve mvarMembers(lngCount - 1)
            End If
        End If
    End If
End Sub

'@方法    Sub_FromRecordset_AddOneRow
'   FromRecordset子函数，用来添加一行数据
'@返回值
'
'@参数:
'Param 方向
'
'@备注
'
Friend Sub Sub_FromRecordset_AddOneRow(rsInput As ADODB.Recordset, Optional intJsonStyle As Integer)
    Dim i       As Long
    If intJsonStyle = 0 Then
        Call EmptyJsonObject
        With rsInput
            For i = 0 To .Fields.Count - 1
                If IsNull(.Fields(i).Value) Then
                    mvarMembers(.Fields(i).Name) = Null
                Else
                    Select Case Sub_FieldType(.Fields(i).Type)
                        Case vbString
                            mvarMembers(.Fields(i).Name) = .Fields(i).Value
                        Case vbLong
                            mvarMembers(.Fields(i).Name) = CDbl(.Fields(i).Value)
                        Case vbDate
                            mvarMembers(.Fields(i).Name) = Format(.Fields(i).Value, "yyyy-mm-dd hh:mm:ss")
                        Case Else
                            mvarMembers(.Fields(i).Name) = Null
                    End Select
                End If
            Next
        End With
    Else
        Call EmptyJsonArray
        With rsInput
            ReDim Preserve mvarMembers(.Fields.Count - 1)
            For i = 0 To UBound(mvarMembers)
                If IsNull(.Fields(i).Value) Then
                    mvarMembers(i) = Null
                Else
                    Select Case Sub_FieldType(.Fields(i).Type)
                        Case vbString
                            mvarMembers(i) = .Fields(i).Value
                        Case vbLong
                            mvarMembers(i) = CDbl(.Fields(i).Value)
                        Case vbDate
                            mvarMembers(i) = Format(.Fields(i).Value, "yyyy-mm-dd hh:mm:ss")
                        Case Else
                            mvarMembers(i) = Null
                    End Select
                End If
            Next
        End With
    End If
End Sub


'@方法    Sub_FieldType
'   判断字段类型
'@返回值  VBA.VbVarType
'   返回vbString（字符串）、vbLong（整形）、vbDate（时间类型）、vbByte（二进制类型）、vbUserDefinedType（暂时不支持的类型）
'@参数:
'dtFiledType In DataTypeEnum
'   字段类型
'@备注
'
Private Function Sub_FieldType(ByVal dtFiledType As DataTypeEnum) As VBA.VbVarType
    Select Case dtFiledType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            Sub_FieldType = vbString
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            Sub_FieldType = vbLong
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            Sub_FieldType = vbDate
        Case adBinary, adVarBinary, adLongVarBinary
            Sub_FieldType = vbByte
        Case Else
            Sub_FieldType = vbUserDefinedType
    End Select
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------

Private Sub Class_Initialize()
    Call InitJsonError
End Sub

