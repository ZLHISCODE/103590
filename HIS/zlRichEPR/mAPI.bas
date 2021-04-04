Attribute VB_Name = "mAPI"
'#########################################################################
'##ģ �� ����mAPI.bas
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��3��25��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ͨ�� Windows API ����
'##��    ����
'#########################################################################

Option Explicit
Public glngPro As Long
Public glngGdi As Long
Public Declare Function GetGuiResources Lib "User32.dll" (ByVal hProcess As Long, ByVal uiFlags As Integer) As Long

'#########################################################################
' ��Ϣ����:
Public Const WM_ACTIVATE = &H6              '����״̬������WA_INACTIVE��δ��� / WM_ACTIVATE����� / WA_CLICKACTIVE����꼤�
Public Const WM_SETFOCUS = &H7              '�߱����㣬Ӧ����α�ָ�뺯��ʹ��
Public Const WM_KILLFOCUS = &H8F            'ȥ�����̽��㣬Ӧɾ������α�ָ��
Public Const WM_SETREDRAW = &HB             'ǿ��ˢ��
Public Const WM_GETTEXTLENGTH = &HE         '�����ı��ַ����ȣ���� GetWindowText() / WM_GETTEXT / LB_GETTEXT / CB_GETLBTEXT
Public Const WM_PAINT = &HF                 '���ƴ���
Public Const WM_ERASEBKGND = &H14           '������屳��
Public Const WM_SETCURSOR = &H20            '�����α�
Public Const WM_MOUSEACTIVATE = &H21        '��������꼤��
Public Const WM_GETMINMAXINFO = &H24        '���ڴ�������󻯳ߴ缰λ��
Public Const WM_WINDOWPOSCHANGING = &H46    '����״̬�ı�
Public Const WM_NOTIFY = &H4E               '�����¼�ʱ����ʾ������
Public Const WM_NCHITTEST = &H84            '����ƶ�������������ſ��¼�
Public Const WM_NCPAINT = &H85              '�����ܻ�����Ϣ������ͨ���������Զ�����ƿ�ܣ���һ���Ǿ��εġ�
Public Const WM_KEYDOWN = &H100             '���̽��㴰���еķ�Alt^�İ������¡�
Public Const WM_KEYUP = &H101               '���̽��㴰���еķ�Alt^�İ����ſ���
Public Const WM_CHAR = &H102                '����WM_KEYDOWN�İ���ֵ��
Public Const WM_COMMAND = &H111             '�˵�������ؼ��򸸴��巢��Notify��Ϣ���߿�ݼ�����ʱ����
Public Const WM_HSCROLL = &H114             'ˮƽ������
Public Const WM_VSCROLL = &H115             '��ֱ������
Public Const WM_SYSCOMMAND = &H112          'ϵͳ�˵����ؼ��˵��ȵ��¼�
Public Const WM_MOUSEMOVE = &H200           '����ƶ��¼�
Public Const WM_LBUTTONDOWN = &H201         '����������
Public Const WM_LBUTTONUP = &H202           '�������ſ�
Public Const WM_LBUTTONDBLCLK = &H203       '������˫��
Public Const WM_RBUTTONDOWN = &H204         '����Ҽ�����
Public Const WM_RBUTTONUP = &H205           '����Ҽ��ſ�
Public Const WM_RBUTTONDBLCLK = &H206       '����Ҽ�˫��
Public Const WM_MBUTTONDOWN = &H207         '����м�����
Public Const WM_MBUTTONUP = &H208           '����м��ſ�
Public Const WM_PARENTNOTIFY = &H210        '�Ӵ��崴��������
Public Const WM_EXITSIZEMOVE = &H232        'Resize���
Public Const WM_UNDO = &H304&               'UNDO����
Public Const WM_CUT = &H300&                '����
Public Const WM_COPY = &H301&               '����
Public Const WM_PASTE = &H302&              'ճ��
Public Const WM_MOUSEWHEEL = &H20A          '������
Public Const EM_GETSEL = &HB0&              '��ȡ��ǰѡ������Ŀ�ʼ�ͽ����ַ�λ�á����ܴ���65, 535��
Public Const EM_SETSEL = &HB1&              'ѡ��ĳһ��Χ���ݡ�
Public Const EM_GETRECT = &HB2&             '��ȡһ��Edit�ؼ��ĸ�ʽ����������
Public Const EM_SETRECT = &HB3&             '����Edit�ؼ��ĸ�ʽ����������ͬʱ�ػ��ı���
Public Const EM_SETRECTNP = &HB4&           'ͬ�ϣ����ǲ��ػ��ı���
Public Const EM_SCROLL = &HB5&              '��ֱ������Ϣ��
Public Const EM_LINESCROLL = &HB6&          'ˮƽ��ֱ�����ı���
Public Const EM_SCROLLCARET = &HB7&         '������Ϊ���ӡ�
Public Const EM_GETMODIFY = &HB8&           '�ж��Ƿ����ݱ��޸��ˡ�
Public Const EM_SETMODIFY = &HB9&           '���û���������޸ı�־��
Public Const EM_GETLINECOUNT = &HBA&        '��ȡ������
Public Const EM_LINEINDEX = &HBB&           '��ȡĳ�е��ַ�����ֵ�����ı�ͷ��ʼ����
Public Const EM_SETHANDLE = &HBC&           '���ö���Edit�ؼ����ڴ�����
Public Const EM_GETHANDLE = &HBD&           '��ȡ��ǰEdit�ؼ����ڴ�����
Public Const EM_GETTHUMB = &HBE&            '��ȡ��ǰ������λ�á�
Public Const EM_LINELENGTH = &HC1&          '��ȡĳ�е��ַ����ȡ�
Public Const EM_REPLACESEL = &HC2&          '�滻��ǰѡ�������ı���
Public Const EM_GETLINE = &HC4&             '����һ���ı���ָ����������
Public Const EM_LIMITTEXT = &HC5&           '�����û�������ı�������
Public Const EM_CANUNDO = &HC6&             '�Ƿ������Ӧ EM_UNDO ��Ϣ��
Public Const EM_UNDO = &HC7&                'Undo��Ϣ��
Public Const EM_FMTLINES = &HC8&            '������س����Ƿ����á�
Public Const EM_LINEFROMCHAR = &HC9&        '��ȡָ���ַ�����ֵ��������
Public Const EM_SETTABSTOPS = &HCB&         '�����Ʊ��λ�����顣
Public Const EM_SETPASSWORDCHAR = &HCC&     '�������������ַ���
Public Const EM_EMPTYUNDOBUFFER = &HCD&     '���Undo���С�
Public Const EM_GETFIRSTVISIBLELINE = &HCE& '������Ŀ����е������������У�������������ַ����������У���
Public Const EM_SETREADONLY = &HCF&         'ֻ����
Public Const EM_SETWORDBREAKPROC = &HD0&    '�Զ�����ִ�����̡�
Public Const EM_GETWORDBREAKPROC = &HD1&    '��ȡ��ǰ���ִ�����̵�ַ��
Public Const EM_GETPASSWORDCHAR = &HD2&     '��ȡ���������ַ���
Public Const EM_SETMARGINS = &HD3&          '�������Ҽ�࣬��ˢ�¡�
Public Const EM_GETMARGINS = &HD4&          '��ȡ...
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT '�����ַ���󳤶ȡ� ' /* ;win40 Name change */
Public Const EM_GETLIMITTEXT = &HD5&        '��ȡ�ַ���󳤶ȡ�
Public Const EM_POSFROMCHAR = &HD6&         '��ȡָ���ַ�������(X,Y)��
Public Const EM_CHARFROMPOS = &HD7&         '��ȡָ������㸽�����ַ���
Public Const EC_LEFTMARGIN = &H1            '��ʾ��������߽硣
Public Const EC_RIGHTMARGIN = &H2           '��ʾ�������ұ߽硣
Public Const EC_USEFONTINFO = &HFFFF&       '�߽�����ַ���ȡ�
Public Const ES_LEFT = &H0&             '�����
Public Const ES_CENTER = &H1&           '����
Public Const ES_RIGHT = &H2&            '�Ҷ���
Public Const ES_MULTILINE = &H4&        '����
Public Const ES_UPPERCASE = &H8&        '��д
Public Const ES_LOWERCASE = &H10&       'Сд
Public Const ES_PASSWORD = &H20&        '����
Public Const ES_AUTOVSCROLL = &H40&     '�Զ���ֱ����
Public Const ES_AUTOHSCROLL = &H80&     '�Զ�ˮƽ����10���ַ�
Public Const ES_NOHIDESEL = &H100&      'ʧȥ����ʱ����ѡ�����ݡ�
Public Const ES_OEMCONVERT = &H400&     '
Public Const ES_READONLY = &H800&       'ֻ��
Public Const ES_WANTRETURN = &H1000&    '�س������С�����س���ͬ�ڴ�����Ĭ�ϰ�ť�¼���
Public Const ES_NUMBER = &H2000&        'ֻ�����������֡�
Public Const EN_CHANGE = &H300          '���ݸı䡣������ͨ�� WM_COMMAND ��Ϣ��ȡ��֪ͨ��
Public Const EN_ERRSPACE = &H500        '���ݲ����Է���ò�����
Public Const EN_HSCROLL = &H601         'ˮƽ�����¼���
Public Const EN_KILLFOCUS = &H200       'ʧȥ�����¼���
Public Const EN_MAXTEXT = &H501         '������ı���������ַ����������ڷ��Զ�����ʱ�����ؼ���������
Public Const EN_SETFOCUS = &H100        '��ü������뽹�㡣
Public Const EN_UPDATE = &H400          '���û��ı����ݵ��ǻ�û��ˢ����ʾʱ������֪ͨ���û��������ڵ��ڿؼ��ߴ�����Ӧ���ݡ�
Public Const EN_VSCROLL = &H602         '��ֱ�����¼���
Public Const WM_CONTEXTMENU = &H7B&     '֪ͨ������Ҽ�����¼�
Public Const WM_PRINTCLIENT = &H318&    '���������ͻ�����һ��ָ�����豸�������У�ͨ����ָ��ӡ����
Public Const WM_USER = &H400                'ͨ���� WM_USER + X ���Զ�����Ϣ
Public Const LF_FACESIZE = 32   '���������ֽڳ��ȡ�
Public Const RICHEDIT_VER = &H210    '��ǰRich Edit�ؼ��汾��
Public Const cchTextLimitDefault = 32767&       'Ĭ���ı���������
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0
Public Const RICHEDIT_CLASSW = "RichEdit20W"
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSW       'UNICODE�汾��
Public Const EM_CANPASTE = (WM_USER + 50)       '�����Ƿ����ճ��ָ����ʽ�ļ��������ݡ�
Public Const EM_DISPLAYBAND = (WM_USER + 51)    '��ʾRTB���ݵ�һ���־������򣬸������� EM_FORMATRANGE ��Ϣ��ʽ��һ���豸�����á��ü������ɸþ��ξ�����
Public Const EM_EXGETSEL = (WM_USER + 52)       '��ȡѡ�е���ʼ����ֹ�ַ�λ�á�
Public Const EM_EXLIMITTEXT = (WM_USER + 53)    '�����û������������ճ����RTB�е��ı��������ޡ�OLE������Ϊһ���ַ���Ĭ��Ϊ32K��
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54) '�ж�����һ�а���ָ���ַ���
Public Const EM_EXSETSEL = (WM_USER + 55)       'ѡ��һ����Χ���ַ�����OLE����
Public Const EM_FINDTEXT = (WM_USER + 56)       '�����ı���
Public Const EM_FORMATRANGE = (WM_USER + 57)    'Ϊĳһ�豸��ʽ��ָ����Χ���ı���
Public Const EM_GETCHARFORMAT = (WM_USER + 58)  '�ж�Ĭ���ַ���ʽ���ߵ�ǰ��Χ��һ���ַ��ĸ�ʽ��
Public Const EM_GETEVENTMASK = (WM_USER + 59)   '��ȡ�¼����롣
Public Const EM_GETOLEINTERFACE = (WM_USER + 60) '��ȡһ��OLE���󣬿ͻ����������ʸ�OLE����Ĺ��ܡ���ʱ���ȵ���AddRef() ����һ�����ã��û���Ҫ����������Release() ������
Public Const EM_GETPARAFORMAT = (WM_USER + 61)  '��ȡ��ǰ����ĵ�һ������Ķ������ԡ�
Public Const EM_GETSELTEXT = (WM_USER + 62)     '��ȡ��ǰѡ�е��ı�����ȷ���������������ɸ��ı���
Public Const EM_HIDESELECTION = (WM_USER + 63)  '��ʾ/�����ı���
Public Const EM_PASTESPECIAL = (WM_USER + 64)   'ѡ����ճ����
Public Const EM_REQUESTRESIZE = (WM_USER + 65)  '֪ͨ������ı�ߴ磬���޵׿ؼ������ã�
Public Const EM_SELECTIONTYPE = (WM_USER + 66)  '�ж�ѡ����������ͣ����ı�����OLE���󣬻��߶��OLE/�ı�����
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)  '����RTB����ɫ��
Public Const EM_SETCHARFORMAT = (WM_USER + 68)  '�����ַ���ʽ��
Public Const EM_SETEVENTMASK = (WM_USER + 69)   '�����¼����롣
Public Const EM_SETOLECALLBACK = (WM_USER + 70) '�ṩһ��IRichEditOleCallback �����RTB�����ڴӿͻ��˻�ȡOLE�����Դ����Ϣ��
Public Const EM_SETPARAFORMAT = (WM_USER + 71)  '���ö����ʽ��
Public Const EM_SETTARGETDEVICE = (WM_USER + 72) '�����������������õ�Ŀ���豸���п�
Public Const EM_STREAMIN = (WM_USER + 73)       '��ʽ���루��ȡ����ʹ��Ӧ�ó����ṩ��EditStreamCallback�ص������ṩ���������滻RTB���ݡ�
Public Const EM_STREAMOUT = (WM_USER + 74)      '��ʽ�����д�룩��ĳһ�ļ���ָ��λ�á�
Public Const EM_GETTEXTRANGE = (WM_USER + 75)   '����һ��ָ���ı���ѡ������
Public Const EM_FINDWORDBREAK = (WM_USER + 76)  '��ȡǰһ/��һ����λ�ã����߻�ȡ��ǰλ���ַ���Ϣ��
Public Const EM_SETOPTIONS = (WM_USER + 77)     'RTBѡ�����á��硰˫���Զ�ѡ�е��ʡ������Զ����������ȡ�
Public Const EM_GETOPTIONS = (WM_USER + 78)     '��ȡRTBѡ�
Public Const EM_FINDTEXTEX = (WM_USER + 79)     '�����ı���
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80) '��ȡ��ǰע�����չ���ִ�����̵ĵ�ַ��
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81) '���õ�ǰ��չ���ִ�����̡�0��ָ�ΪĬ�ϡ�
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)   '����Undo�������ޡ�
Public Const EM_REDO = (WM_USER + 84)           'Redo������
Public Const EM_CANREDO = (WM_USER + 85)        '�ж�Redo�������Ƿ����κζ������ö������Ƿ����Redo��
Public Const EM_GETUNDONAME = (WM_USER + 86)    '������һ��Undo���������ơ��������� UNDONAMEID ö�ٳ������壡
Public Const EM_GETREDONAME = (WM_USER + 87)    '������һ��Redo���������ơ�
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)    'ֹͣ��ǰUndo���е��ַ��Ѽ����κλ���������һ���С�
Public Const EM_SETTEXTMODE = (WM_USER + 89)    '�����ı�ģʽ��Undo�ȼ������RTB�����κ��ַ��������Ϣ�������ã�
Public Const EM_GETTEXTMODE = (WM_USER + 90)    '��ȡ��ǰ�ı�ģʽ��Undo�ȼ���
Public Const EM_FINDTEXTW = (WM_USER + 123)     '����Unicode���ı���
Public Const EM_FINDTEXTEXW = (WM_USER + 124)   'ͬ�ϡ�
Public Const EM_AUTOURLDETECT = (WM_USER + 91)      '����/�����Զ�URL��⡣
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)   '�ж��Ƿ��������Զ�URL��⡣
Public Const EM_SETPALETTE = (WM_USER + 93)         '�ı��ɫ�塣
Public Const EM_GETTEXTEX = (WM_USER + 94)          '��ȡָ������ҳ���ı���
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)    '���ò�ͬ��ʽ�����ı����ȡ�
Public Const EM_SETPUNCTUATION = (WM_USER + 100)    '���ñ����š��������������ԵĲ���ϵͳ��
Public Const EM_GETPUNCTUATION = (WM_USER + 101)    '��ȡ�����š��������������ԵĲ���ϵͳ��
Public Const EM_SETWORDWRAPMODE = (WM_USER + 102)   '�����Զ����������ѡ��������������ԵĲ���ϵͳ��
Public Const EM_GETWORDWRAPMODE = (WM_USER + 103)   '��ȡ�Զ����������ѡ��������������ԵĲ���ϵͳ��
Public Const EM_SETIMECOLOR = (WM_USER + 104)       '����IME�����ɫ���������������ԵĲ���ϵͳ��
Public Const EM_GETIMECOLOR = (WM_USER + 105)       '��ȡIME�����ɫ���������������ԵĲ���ϵͳ��
Public Const EM_SETIMEOPTIONS = (WM_USER + 106)     '����IMEѡ��������������ԵĲ���ϵͳ��
Public Const EM_GETIMEOPTIONS = (WM_USER + 107)     '��ȡIMEѡ��������������ԵĲ���ϵͳ��
Public Const EM_CONVPOSITION = (WM_USER + 108)      '������RTB v1.0 ���������ԵĲ���ϵͳ��RTB 2.0��֧�֣�
Public Const EM_SETLANGOPTIONS = (WM_USER + 120)    '����IME��Զ������֧��ѡ�
Public Const EM_GETLANGOPTIONS = (WM_USER + 121)    '��ȡIME��Զ������֧��ѡ�
Public Const EM_GETIMECOMPMODE = (WM_USER + 122)    '��ȡ��ǰIMEģʽ��
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)    '���õ�ǰ˫������֧��ѡ�
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)    '��ȡ��ǰ˫������֧��ѡ�
Public Const IMF_AUTOKEYBOARD = &H1             '�Զ����̲���
Public Const IMF_AUTOFONT = &H2                 '�Զ�����
Public Const IMF_IMECANCELCOMPLETE = &H4      '// high completes the comp string when aborting, low cancels.
Public Const IMF_IMEALWAYSSENDNOTIFY = &H8
Public Const ICM_NOTOPEN = &H0          'Input Method Editor (IME) is not open.
Public Const ICM_LEVEL3 = &H1           'True inline mode.
Public Const ICM_LEVEL2 = &H2           'Level 2.
Public Const ICM_LEVEL2_5 = &H3         'Level 2.5
Public Const ICM_LEVEL2_SUI = &H4       'Special user interface (UI).
Public Const EN_MSGFILTER = &H700&      'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ�������������߼����¼�������
Public Const EN_REQUESTRESIZE = &H701&  'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ������ߴ��иı䡣
Public Const EN_SELCHANGE = &H702&      'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ�����嵱ǰѡ���������仯��
Public Const EN_DROPFILES = &H703&      'RTB�ؼ��ڽ��ܵ� WM_DROPFILES ��Ϣ��ͨ�� WM_NOTIFY ��Ϣ֪ͨ�������û���ͼ����һ���ļ���
Public Const EN_PROTECTED = &H704&      'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ�������û���ͼ�ı��ܱ����ı���
Public Const EN_CORRECTTEXT = &H705&    'һ��EN_CORRECTTEXT ���ơ�   /* PenWin specific */
Public Const EN_STOPNOUNDO = &H706&     'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ������ĳ�������޷������㹻�ڴ�����¼��״̬��
Public Const EN_IMECHANGE = &H707&      'IME �ı䡣                  /* Far East specific */
Public Const EN_SAVECLIPBOARD = &H708&  '֪ͨ�����壬RTB�ڹر�ʱ�������л������ݡ�
Public Const EN_OLEOPFAILED = &H709&    '֪ͨ�����壬һ����OLE����Ĳ���ʧ�ܡ�
Public Const EN_OBJECTPOSITIONS = &H70A&    '֪ͨ�����壬RTB����һ��OLE����
Public Const EN_LINK = &H70B&               'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ�������û��ڳ�����Ч���ı��ϵĶ�������¼���
Public Const EN_DRAGDROPDONE = &H70C&       'RTB�ؼ�ͨ�� WM_NOTIFY ��Ϣ֪ͨ������һ���ϷŲ�����ɡ�
Public Const EN_ALIGN_LTR = &H710&      'RTB�ؼ�ͨ�� WM_COMMAND ��Ϣ֪ͨ��������䷽���Ϊ�������ҡ�
Public Const EN_ALIGN_RTL = &H711&      'RTB�ؼ�ͨ�� WM_COMMAND ��Ϣ֪ͨ��������䷽���Ϊ��������
Public Const ENM_NONE = &H0             'Ĭ��ֵ����ʾ�����򸸴��巢���κ���Ϣ��
Public Const ENM_CHANGE = &H1           '���Է��� EN_CHANGE ��Ϣ��
Public Const ENM_UPDATE = &H2           '���Է��� EN_UPDATE ��Ϣ��
Public Const ENM_SCROLL = &H4           '���Է��� EN_HSCROLL ��Ϣ��
Public Const ENM_KEYEVENTS = &H10000    '���Է��� EN_MSGFILTER ��Ϣ��
Public Const ENM_MOUSEEVENTS = &H20000  '���Է��� EN_MSGFILTER ��Ϣ��
Public Const ENM_REQUESTRESIZE = &H40000    '���Է��� EN_REQUESTRESIZE ��Ϣ��
Public Const ENM_SELCHANGE = &H80000        '���Է��� EN_SELCHANGE ��Ϣ��
Public Const ENM_DROPFILES = &H100000       '���Է��� EN_DROPFILES ��Ϣ��
Public Const ENM_PROTECTED = &H200000       '���Է��� EN_PROTECTED ��Ϣ��
Public Const ENM_CORRECTTEXT = &H400000     ' /* PenWin specific */
Public Const ENM_SCROLLEVENTS = &H8         '���Է��� EN_MSGFILTER �е��������¼���Ϣ��
Public Const ENM_DRAGDROPDONE = &H10        '���Է��� EN_DRAGDROPDONE ��Ϣ��
Public Const ENM_IMECHANGE = &H800000           ' /* RE2.0 ��֧�֣���ֻ����1.0�汾��*/
Public Const ENM_LANGCHANGE = &H1000000         ' ����
Public Const ENM_OBJECTPOSITIONS = &H2000000    '���Է��� EN_OBJECTPOSITIONS ��Ϣ��
Public Const ENM_LINK = &H4000000               '���Է��� EN_LINK ��Ϣ��
Public Const ES_SAVESEL = &H8000&               '��ʧȥ����ʱ����ѡ�����������ʾ������Useful��
Public Const ES_SUNKEN = &H4000&                '����Ч��
Public Const ES_DISABLENOSCROLL = &H2000&       '�ڲ���Ҫ������ʱ�����ûң���������
Public Const ES_SELECTIONBAR = &H1000000 ' /* same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value */
Public Const ES_NOOLEDRAGDROP = &H8 ' /* same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop */
Public Const ES_EX_NOCALLOLEINIT = &H1000000 ' /* �µ� Edit �ؼ���չ��ʽ */
Public Const ES_VERTICAL = &H400000     '��ֱ�����ı��Ͷ���
Public Const ES_NOIME = &H80000         '����IME��
Public Const ES_SELFIME = &H40000       'Ӧ�ó���������IME������
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
Public Const CFM_BOLD = &H1             '������Ч��
Public Const CFM_ITALIC = &H2           'б����Ч��
Public Const CFM_UNDERLINE = &H4        '�»�����Ч��
Public Const CFM_STRIKEOUT = &H8        'ɾ������Ч��
Public Const CFM_PROTECTED = &H10       '������Ч��
Public Const CFM_LINK = &H20&           '��������Ч��  ' /* Exchange hyperlink extension */
Public Const CFM_SIZE = &H80000000      '�ַ��߶���Ч����λ��羡�
Public Const CFM_COLOR = &H40000000     '�ı���ɫ��Ч��
Public Const CFM_FACE = &H20000000      '����������Ч��
Public Const CFM_OFFSET = &H10000000    '�ַ�ƫ����Ч��ָ�����ϻ��µ�ƫ�������ϱ�/�±꣩��
Public Const CFM_CHARSET = &H8000000    '�ַ�����Ч��
Public Const CFE_BOLD = &H1&            '����
Public Const CFE_ITALIC = &H2&          'б��
Public Const CFE_UNDERLINE = &H4&       '�»���
Public Const CFE_STRIKEOUT = &H8&       'ɾ����
Public Const CFE_PROTECTED = &H10&      '����
Public Const CFE_LINK = &H20&           '������
Public Const CFE_AUTOCOLOR = &H40000000 '����ϵͳ�Զ���ɫ��' /* NOTE: this corresponds to */Public Const SCF_SELECTION = &H1&   'Ӧ���ڵ�ǰѡ������
Public Const SCF_WORD = &H2&        'Ӧ���ڵ�ǰѡ�е��ʡ�
Public Const SCF_DEFAULT = &H0&            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4&                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8&         '// modifier for SCF_SELECTION; says that
Public Const SF_TEXT = &H1         'Text��ʽ
Public Const SF_RTF = &H2          'RTF��ʽ
Public Const SF_RTFNOOBJS = &H3    '���ʱ�ÿո������󣬽����������
Public Const SF_TEXTIZED = &H4     '���ʱ�����ı���ʾ���󣬽����������
Public Const SF_UNICODE = &H10            ' /* Unicode file of some kind */
Public Const SFF_SELECTION = &H8000&    '�������ֻ�Ե�ǰѡ��������Ч��
Public Const SFF_PLAINRTF = &H4000&     'ֻʹ��ͨ��RTF�ؼ��֣�������������ص�RTF�ؼ������Ժ��ԣ�
Public Const MAX_TAB_STOPS = 32&    '�����Ʊ���������Ŀ��
Public Const lDefaultTab = 720&     'Ĭ�Ͼ����Ʊ��λ�á�
Public Const PFM_STARTINDENT = &H1& '��������ֵ��Ч��
Public Const PFM_RIGHTINDENT = &H2& '������ֵ��Ч��
Public Const PFM_OFFSET = &H4&      '��������������Ч����ֵ��ʾ��������ֵ��ʾ���ң�
Public Const PFM_ALIGNMENT = &H8&   'ˮƽ���뷽ʽ��Ч��
Public Const PFM_TABSTOPS = &H10&   '�����Ʊ��λ����Ч��
Public Const PFM_NUMBERING = &H20&  '�������Ŀ������Ч��
Public Const PFM_OFFSETINDENT = &H80000000  '��������ֵ��Ч�����Ҹ���һ�����ֵ��
Public Const PFN_BULLET = &H1&   ' /* PARAFORMAT ���ѡ�� */
Public Const PFA_LEFT = &H1&   ' /* PARAFORMAT ����ѡ�� */     '
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
Public Const CFE_SUBSCRIPT = &H10000                ' /*  �ϱ���±��ǻ���ģ�      */
Public Const CFE_SUPERSCRIPT = &H20000              ' /*  �ϱ���±��ǻ���ģ�      */
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
Public Const PFE_TABLECELL = &H4000&               ' /*  ����Ϊ����һ���� */
Public Const PFA_JUSTIFY = 4          ' /* ���˶��룬Ϊ�˼���TOMģ�ͽӿڡ� (*)  */
Public Const GCM_RIGHTMOUSEDROP = &H8000& ' ������ IRichEditOleCallback::GetContextMenu ����������Ӧ�ó����ṩһ���Ҽ��˵���
Public Const CF_RTF = "Rich Text Format" ' �������ʽ������ RegisterClipboardFormat() ע����Ч�ļ������ʽ��
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
Public Const FR_MATCHDIAC = &H20000000          ' ��������ϣ��������
Public Const FR_MATCHKASHIDA = &H40000000       ' ��������ϣ��������
Public Const FR_MATCHALEFHAMZA = &H80000000     ' ��������ϣ��������
Public Const WCH_EMBEDDING = &HFFFC&

Public Const MB_OK = &H0&

'��ʽ
Public Const BS_HATCHED = 2
Public Const BS_NULL = 1
Public Const BS_SOLID = 0

'����
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

Public Const SM_CYHSCROLL = 3 'ˮƽ�������ϵ�ʸ��λͼ���
Public Const SM_CXVSCROLL = 2 'ˮƽ�������ϵ�ʸ��λͼ�߶�
Public Const SC_MOVE = &HF012 '�û����ϵͳ�˵��еġ��ƶ����˵��¼�
'SetWindowPos����������
'��ʾǿ�Ʒ��� WM_NCCALCSIZE ��Ϣ������
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
'�ڴ����ⲿ����һ�����
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'�Ǽ���״̬
Public Const SWP_NOACTIVATE = &H10
'���ֵ�ǰλ��
Public Const SWP_NOMOVE = &H2
'���ֵ�ǰ�ߴ�
Public Const SWP_NOSIZE = &H1
'���ֵ�ǰZ-Order
Public Const SWP_NOZORDER = &H4
'���游����Z-Order
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80&


Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOREDRAW = &H8
Public Const SWP_SHOWWINDOW = &H40
Public Const SPI_GETWORKAREA = 48&

Public Const OLEOP_DOVERB = 1
' ���� GETTEXTEX ���ݽṹ
Public Const GT_DEFAULT = 0&    '��ʹ��CRת��
Public Const GT_USECRLF = 1&    '��ʾ��ÿ�ο����ı�ʱ����CRת��ΪCRLF��
' GETTEXTLENGTHEX ���ݽṹ�ı�־λ
Public Const GTL_DEFAULT = 0&      ' /* Ĭ��ֵ�������ַ���Ŀ��                      */
Public Const GTL_USECRLF = 1&      ' /* ʹ�ö��� CR/LF ����                         */
Public Const GTL_PRECISE = 2&      ' /* ��ȷ���㣬����                              */
Public Const GTL_CLOSE = 4&        ' /* ���Ƽ��㣬�Ͽ죬��������ǰ�����ڴ�ռ�      */
Public Const GTL_NUMCHARS = 8&     ' /* �����ַ���Ŀ                                */
Public Const GTL_NUMBYTES = 16&    ' /* �����ֽ���Ŀ

'ÿ���߼���λΪ1���豸���ء���X���ң���Y���¡�����SetMapMode()
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
Public Const DC_PAPERNAMES = 16 'ֽ������(ÿ64�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_PAPERS = 2 'ֽ�ű��(Array or Word)
Public Const DC_BINNAMES = 12 '��ֽ��ʽ(ÿ24�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_BINS = 6 '��ֽ���(Array or Word)
'��ӡֽ�ų���(256=�Զ���)
Public Const PageSize1 = "�ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize2 = "+A611 С���ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize3 = "С�ͱ��� 11 x 17 Ӣ��"
Public Const PageSize4 = "�����ʣ� 17 x 11 Ӣ��"
Public Const PageSize5 = "�����ļ��� 8 1/2 x 14 Ӣ��"
Public Const PageSize6 = "�����飬5 1/2 x 8 1/2 Ӣ��"
Public Const PageSize7 = "�����ļ���7 1/2 x 10 1/2 Ӣ��"
Public Const PageSize8 = "A3, 297 x 420 ����"
Public Const PageSize9 = "A4, 210 x 297 ����"
Public Const PageSize10 = "A4С�ţ� 210 x 297 ����"
Public Const PageSize11 = "A5, 148 x 210 ����"
Public Const PageSize12 = "B4, 250 x 354 ����"
Public Const PageSize13 = "B5, 182 x 257 ����"
Public Const PageSize14 = "�Կ����� 8 1/2 x 13 Ӣ��"
Public Const PageSize15 = "�Ŀ����� 215 x 275 ����"
Public Const PageSize16 = "10 x 14 Ӣ��"
Public Const PageSize17 = "11 x 17 Ӣ��"
Public Const PageSize18 = "������8 1/2 x 11 Ӣ��"
Public Const PageSize19 = "#9 �ŷ⣬ 3 7/8 x 8 7/8 Ӣ��"
Public Const PageSize20 = "#10 �ŷ⣬ 4 1/8 x 9 1/2 Ӣ��"
Public Const PageSize21 = "#11 �ŷ⣬ 4 1/2 x 10 3/8 Ӣ��"
Public Const PageSize22 = "#12 �ŷ⣬ 4 1/2 x 11 Ӣ��"
Public Const PageSize23 = "#14 �ŷ⣬ 5 x 11 1/2 Ӣ��"
Public Const PageSize24 = "C �ߴ繤����"
Public Const PageSize25 = "D �ߴ繤����"
Public Const PageSize26 = "E �ߴ繤����"
Public Const PageSize27 = "DL ���ŷ⣬ 110 x 220 ����"
Public Const PageSize28 = "C5 ���ŷ⣬ 162 x 229 ����"
Public Const PageSize29 = "C3 ���ŷ⣬ 324 x 458 ����"
Public Const PageSize30 = "C4 ���ŷ⣬ 229 x 324 ����"
Public Const PageSize31 = "C6 ���ŷ⣬ 114 x 162 ����"
Public Const PageSize32 = "C65 ���ŷ⣬114 x 229 ����"
Public Const PageSize33 = "B4 ���ŷ⣬ 250 x 353 ����"
Public Const PageSize34 = "B5 ���ŷ⣬176 x 250 ����"
Public Const PageSize35 = "B6 ���ŷ⣬ 176 x 125 ����"
Public Const PageSize36 = "�ŷ⣬ 110 x 230 ����"
Public Const PageSize37 = "�ŷ������ 3 7/8 x 7 1/2 Ӣ��"
Public Const PageSize38 = "�ŷ⣬ 3 5/8 x 6 1/2 Ӣ��"
Public Const PageSize39 = "U.S. ��׼��д���� 14 7/8 x 11 Ӣ��"
Public Const PageSize40 = "�¹���׼��д���� 8 1/2 x 12 Ӣ��"
Public Const PageSize41 = "�¹����ɸ�д���� 8 1/2 x 13 Ӣ��"

'WinNT�Զ���ֽ�ſ���================================================================
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
'�ڴ��������
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

'��ȡ������ʾԪ�صĵ�ǰ��ɫֵ
Public Const COLOR_SCROLLBAR = 0 '������
Public Const COLOR_BACKGROUND = 1 'windows����
Public Const COLOR_ACTIVECAPTION = 2 '����ڵı���
Public Const COLOR_INACTIVECAPTION = 3 '������ڵı��⡣
Public Const COLOR_MENU = 4 '�˵���
Public Const COLOR_WINDOW = 5 '���ڱ���
Public Const COLOR_WINDOWFRAME = 6 '����
Public Const COLOR_MENUTEXT = 7 '�˵����ġ�
Public Const COLOR_WINDOWTEXT = 8 '��������
Public Const COLOR_CAPTIONTEXT = 9 '���ڱ����е�����
Public Const COLOR_ACTIVEBORDER = 10 '����ڵı߿�
Public Const COLOR_INACTIVEBORDER = 11 '������ڵı߿�
Public Const COLOR_APPWORKSPACE = 12 'MDI����ı���
Public Const COLOR_HIGHLIGHT = 13 'ѡ������Ŀ������
Public Const COLOR_HIGHLIGHTTEXT = 14 'ѡ������Ŀ���֡�
Public Const COLOR_BTNFACE = 15 '��ť
Public Const COLOR_BTNSHADOW = 16 '��ť��3D��Ӱ
Public Const COLOR_GRAYTEXT = 17 '��ɫ���֣���ʹ���˶���������Ϊ�㡣
Public Const COLOR_BTNTEXT = 18 '��ť����
Public Const COLOR_INACTIVECAPTIONTEXT = 19 '������ڵ����֡�
Public Const COLOR_BTNHIGHLIGHT = 20 '��ť��3D������
Public Const COLOR_3DDKSHADOW = 21 '3D����Ӱ��
Public Const COLOR_3DLIGHT = 22 '3D��Ӱ�������ǳɫ��
Public Const COLOR_INFOTEXT = 23 '������ʾ���ı�ɫ��
Public Const COLOR_INFOBK = 24 '������ʾ�ı���ɫ��
Public Const COLOR_HOTLIGHT = 26
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100  'shorts
Public Const COLOR_3DFACE = COLOR_BTNFACE '3D��Ӱ�������������ɫ��
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT '3D������ɫ��win95����
Public Const COLOR_DESKTOP = COLOR_BACKGROUND '������ɫ��
Public Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT

Public Const WS_CAPTION = &HC00000
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'3DЧ����ɫ
Public Const LR_LOADMAP3DCOLORS = &H1000
'ͼƬ���ļ�lpsz�е��룬���Ǵ���Դ�ļ��е��롣
Public Const LR_LOADFROMFILE = &H10
'����͸��ɫ
Public Const LR_LOADTRANSPARENT = &H20
'���� �豸�޹� DIB λͼ�������豸���λͼ��
Public Const IMAGE_BITMAP = 0

Public Const MAX_PATH = 260
Public Const DI_NORMAL = &H3
Public Const SHGFI_LARGEICON = &H0                      '��ȡ��ͼ��
Public Const SHGFI_SMALLICON = &H1                      '��ȡСͼ��
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_ICON = &H100                         '��ȡͼ����

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

'�����ļ���
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

'������ʽ:
Public Const WS_CHILD = &H40000000          '�Ӵ���
Public Const WS_HSCROLL = &H100000          '�߱�ˮƽ������
Public Const WS_VSCROLL = &H200000          '�߱���ֱ������
Public Const WS_VISIBLE = &H10000000        '����
Public Const WS_CLIPCHILDREN = &H2000000    '��ȥ�Ӵ���ĸ������ͼ����
Public Const WS_CLIPSIBLINGS = &H4000000    '�����Ӵ���ʱ���ų��ص��������Ӵ���
Public Const WS_BORDER = &H800000           '�߱��߿�
Public Const WS_TABSTOP = &H10000           'Tabͣ��
Public Const WS_POPUP = &H80000000          '��������
Public Const WS_SYSMENU = &H80000           '�ڱ������Ƿ�߱�ϵͳ�˵�
Public Const WS_THICKFRAME = &H40000        '��߿�
Public Const WS_MINIMIZEBOX = &H20000       '�߱���С����ť
Public Const WS_MAXIMIZEBOX = &H10000       '�߱���󻯰�ť
Public Const WS_DLGFRAME = &H400000         '˫�߿����ޱ������Ĵ���
Public Const WS_EX_TOPMOST = &H8&           '��ǰ��
Public Const WS_EX_CLIENTEDGE = &H200&      '3DЧ��
Public Const WS_EX_Transparent = &H20&      '����͸��
Public Const WS_DISABLED = &H8000000        '������

Public Const GWL_STYLE = (-16)              'Set the window style
Public Const GWL_EXSTYLE = (-20)            'Set the extended window style
Public Const GWL_USERDATA = (-21)           'Sets the 32-bit value associated with the window.
Public Const GWL_WNDPROC = -4               'Sets a new address for the window procedure.
Public Const GWL_HWNDPARENT = (-8)          'Sets a new application instance handle.

Public Const HWND_TOPMOST = -1              '��ǰ��
Public Const SW_SHOW = 5                    '����岢��ʾ
Public Const SW_HIDE = 0                    '����
Public Const SW_SHOWNORMAL = 1              '��ԭ
Public Const GW_CHILD = 5                   '��ȡ��������
Public Const GW_HWNDNEXT = 2                '��ȡָ������Z-Order�µ���һ����ľ��
Public Const CW_USEDEFAULT  As Long = &H80000000        'Ĭ��ֵ
Public Const GDI_ERROR = &HFFFF             '����GDI����
' ��꼤����Ӧ
Public Const MA_ACTIVATE = 1                '����CWnd
Public Const MA_ACTIVATEANDEAT = 2          '����CWnd����������¼�
Public Const MA_NOACTIVATE = 3              '������CWnd
Public Const MA_NOACTIVATEANDEAT = 4        '������CWnd����������¼�
Public Const H_MAX As Long = &HFFFF + 1     '���ֵ
Public Const ERROR_FILE_NOT_FOUND = 2&     '�ļ�û���ҵ�
Public Const ERROR_PATH_NOT_FOUND = 3&     '·��û���ҵ�
Public Const ERROR_BAD_FORMAT = 11&        '��Ч����
Public Const SE_ERR_ACCESSDENIED = 5       '�ܾ���ȡ
Public Const SE_ERR_ASSOCINCOMPLETE = 27   '�ļ�������������Ч
Public Const SE_ERR_DDEBUSY = 30           'DDEæ
Public Const SE_ERR_DDEFAIL = 29           'DDEʧ��
Public Const SE_ERR_DDETIMEOUT = 28        'DDE��ʱ
Public Const SE_ERR_DLLNOTFOUND = 32       '��̬���ӿ�û���ҵ�
Public Const SE_ERR_FNF = 2                '�ļ�û���ҵ�
Public Const SE_ERR_NOASSOC = 31           'û�й�������
Public Const SE_ERR_PNF = 3                '·��û���ҵ�
Public Const SE_ERR_OOM = 8                '�ڴ����
Public Const SE_ERR_SHARE = 26             '����Υ��
' ��������볣��
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt
Public Const MOUSEEVENTF_ABSOLUTE = &H8000  '�����ƶ�
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4       '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1         '����ƶ�
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

'����
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

'��
Public Type POINTAPI
    x As Long
    y As Long
End Type

'����λ����Ϣ
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI       '��󻯳ߴ�
    ptMaxPosition As POINTAPI   '���λ��
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                 ' /* Ĭ����Ϊ */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8           ' /* Ĭ����Ϊ */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32           ' /* Ĭ����Ϊ */
End Enum

' /* ���е��ַ���ʽ������λ��Ϊ��� */
' �Ѿ�����������...
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
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58           '��������WCHAR
    wPad2 As Integer ' 60
End Type

'�ַ���Χ��
Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

'�ı���Χ��ͨ�� EM_GETTEXTRANGE ��Ϣ��䣡
Public Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String    ' /* allocated by caller, zero terminated by RichEdit */
End Type

'���ڴ洢 EM_STREAMIN ���� EM_STREAMOUT ��Ϣ���ݵ�������Ϣ��
Public Type EDITSTREAM
    dwCookie As Long     ' /* user value passed to callback as first parameter */
    dwError As Long      ' /* last error */
    pfnCallback As Long  'EDITSTREAMCALLBACK
End Type

'���� EM_FINDTEXT ��Ϣ�Ĳ����ı��������Ϣ
Public Type FindText
    chrg As CHARRANGE   '�ַ���Χ
    lpstrText As Long   '��Ҫ���ҵ��ı�
End Type

 '��չ���ı�������Ϣ�ṹ��
Public Type FINDTEXTEX_A
    chrg As CHARRANGE       '�ַ���Χ
    lpstrText As Long       '��Ҫ���ҵ��ı�
    chrgText As CHARRANGE   '���ҵ����ı���Χ
End Type

'ͬ��
Public Type FINDTEXTEX_W
    chrg As CHARRANGE
    lpstrText As Long
    chrgText As CHARRANGE
End Type
 
'�������ڸ�ʽ��ָ���豸�������Ϣ
Public Type FORMATRANGE
    hDC As Long             '��Ⱦ�豸
    hdcTarget As Long       'Ŀ���豸
    rc As RECT              '��Ⱦ���򣬵�λ��羡�
    rcPage As RECT          '��Ⱦ�豸���������򣬵�λ��羡�
    chrg As CHARRANGE       '���ڸ�ʽ�����ı���Χ��
End Type

' /* ���ж��������λ��Ϊ��� */
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
    
    'RICHEDIT20 ֧�ֵ��³�Ա
    wWeight As Integer              ' /* �����ֵ���μ�LOGFONTֵ��      */
    sSpacing As Integer             ' /* ˮƽ�ַ���������ڼ���TOM�ӿ�  */
    crBackColor As Long             ' /* ����ɫ                         */
    lLCID As Long                   ' /* 32λ�ı��� ID                  */
    dwReserved As Long              ' /* ����������Ϊ0                  */
    sStyle As Integer               ' /* ��ʽָ�룬���ڼ���TOM�ӿ�      */
    wKerning As Integer             ' /* �ַ�ѹ����С��ȣ����ڼ���TOM�ӿ� */
    bUnderlineType As Byte          ' /* �»�������                     */
    bAnimation As Byte              ' /* ��̬�ı�Ч�������ڼ���TOM�ӿ�  */
    bRevAuthor As Byte              ' /* �޶������������ò�ͬ��ɫ��ʾ��ͬ���ߵ��޶���Ϣ */
    bReserved1 As Byte              ' /* ����������Ϊ0                  */
End Type

Public Type PARAFORMAT2
    cbSize As Integer               'ָ���ýṹ���ֽڴ�С��
    wPad1 As Integer                '
    dwMask As Long                  '�������
    wNumbering As Integer           '��Ŀ��������
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

Public Type NMHDR ' /* ֪ͨ�Ľṹ */
    hwndFrom As Long        '��Ϣ���͵�Ŀ�괰��
    wPad1 As Integer        '-
    idfrom As Integer       '������Ϣ�Ŀؼ�ID
    code As Integer         '��Ϣ����
    wPad2 As Integer        '-
End Type
' #endif  ' /* !WM_NOTIFY */

'���� EN_MSGFILTER ��Ϣ���洢��ꡢ�����¼���
Public Type MSGFILTER
    NMHDR As NMHDR '֪ͨͷ
    msg As Integer          '���̻�������ʶ��
    wPad1 As Integer        '-
    wParam As Integer       '��Ϣ��wParamֵ��ָ����RTB��ID
    wPad2 As Integer        '-
    lParam As Long          '��Ϣ��lParamֵ��ָ���Ǹ���Ϣ�� MSGFILTER �ṹ���ָ�롣
End Type

Public Type REQRESIZE
    NMHDR As NMHDR     '֪ͨͷ
    rc As RECT                  '������³ߴ磡
End Type

Public Type SelChange
    NMHDR As NMHDR     '֪ͨͷ
    chrg As CHARRANGE           '�µ�ѡ��Χ
    seltyp As Long              '�µ�ѡ��Χ�����ݣ��ı������󡢶������ȣ�
End Type

'����ק�µ��ļ���Ϣ
Public Type ENDROPFILES
    NMHDR As NMHDR     '֪ͨͷ
    hDrop As Long               '���µ��ļ��б�����ͬ WM_DROPFILES��
    cP As Long                  '����������ַ�λ��
    fProtected As Long          'ָ�����ַ�λ���Ƿ��ܱ���
End Type

'�û���ͼ�޸��ܱ����ĵ��ǵ���Ϣ����
Public Type ENPROTECTED
    NMHDR As NMHDR     '֪ͨͷ
    msg As Long                 '������֪ͨ��ԭʼ��Ϣ
    wPad1 As Integer            '-
    wParam As Long              '����Ϣ��wParamֵ
    wPad2 As Integer            '-
    lParam As Long              '����Ϣ��lParamֵ
    chrg As CHARRANGE           '��ǰѡ������
End Type

'�������еĶ�����ı�������
Public Type ENSAVECLIPBOARD
    NMHDR As NMHDR     '֪ͨͷ
    cObjectCount As Long        '�������ж�����Ŀ
    cch As Long                 '���������ַ���Ŀ
End Type

'ʧ�ܵ�OLE���������Ϣ
Public Type ENOLEOPFAILED
    NMHDR As NMHDR     '֪ͨͷ
    iob As Long                 '��������ֵ
    lOper As Long               'ʧ�ܵ�OLE������ȡֵΪ OLEOP_DOVERB ����
    hr As Long                  '���صĴ������
End Type

'����λ��Ϣ���ڶ��󱻶���RTBʱ������֪ͨ
Public Type OBJECTPOSITIONS
    NMHDR As NMHDR     '֪ͨͷ
    cObjectCount As Long        '��������
        ' !!!POINTER to long value!!!
    pcpPositions As Long        '����λ��ָ�롣ע�⣺�ǳ����ε�ָ�룡������
End Type

Public Type ENLINK
    NMHDR As NMHDR     '֪ͨͷ
    msg As Integer              '������֪ͨ����Ϣ
    wPad1 As Integer            '-
    wParam As Integer           '����Ϣ��wParamֵ
    wPad2 As Integer            '-
    lParam As Integer           '����Ϣ��lParamֵ
    chrg As CHARRANGE           '�������ı���Χ
End Type

' /* PenWin specific */
Public Type ENCORRECTTEXT
    NMHDR As NMHDR     '֪ͨͷ
    chrg As CHARRANGE           '��ǰѡ��Χ
    seltyp As Integer           '��Χ�����ݵ�����
End Type


' ѡ����ճ��
Public Type REPASTESPECIAL
    dwAspect As Long    '��ʾ���ԡ�ȡֵ��DVASPECT_CONTENT ���� DVASPECT_ICON
    dwParam As Long     '���ΪDVASPECT_ICON���򱾲�������һ��ָ��ö�����ͼ��һ��ͼԪ�ļ����
End Type



' /* EM_GETTEXTEX ��Ϣ wParam ���� */
Public Type GETTEXTEX
    cb As Long              ' /* ��ȡ���ַ����ֽ���             */
    flags As Long           ' /* �ı�ת������ѡ��               */
    codepage As Long        ' /* ת���Ĵ���ҳ��Ĭ��ΪCP_ACP��UnicodeΪ1200
    lpDefaultChar As Long   ' /* ��Unicodeģʽ���޷���ʾ���ַ�ʱ������ַ���ΪNULL��ʹ��ϵͳĬ��ֵ�� */
    lpUsedDefChar As Long   ' /* �Ƿ������滻�ַ�   */
End Type

' /* EM_GETTEXTLENGTHEX ��ȡ�ı�������Ϣ�� wParam ���� */
Public Type GETTEXTLENGTHEX
    flags As Long                   ' ����
    codepage As Long                ' ����ҳ
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

Public Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

'���ư�͸�����ο���غ���������
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

'Window�汾����
Public Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' ��ӡ֧��:
' VB API Viewer �汾�� DocInfo �ṹ˵���Ǵ���ģ�������
'���ڴ洢 StartDoc() ���ļ�����������Ϣ
Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

'���ڳ�ʼ����ӡ�Ի��򼰷���ֵ
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

'���ô�ӡ�Ի���
Public Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As PrintDlg) As Long

'���� PrintDlg �ĶԻ������������
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


'���ڳ���ͻ������ͼ��Ϣ�ṹ��
Public Type PAINTSTRUCT
   hDC As Long
   fErase As Long
   rcPaint As RECT
   fRestore As Long
   fIncUpdate As Long
   rgbReserved(0 To 31) As Byte
End Type

'����λͼ�����͡���ȡ��߶ȡ���ɫ��ʽ��λ���ݿ�
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
'�����첽����/��� I/O
Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

'�·����
Public Const OFS_MAXPATHNAME = 128

'���� OpenFile ���ļ���Ϣ
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

''''''���ô��������С���رհ�ť
Public Declare Function GetSystemMenu Lib "user32 " (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32 " (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32 " (ByVal hwnd As Long) As Long
Public Declare Function RemoveMenu Lib "user32 " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'ɨ�躯������
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
'����͸��ͼ��
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '��׽����״̬
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  'ϵͳ��ͣ
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'ע���ֱ�Ӳ���
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long
'�ڶ�ջ�з���ָ���ֽ������ڴ棬ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'�ͷ��ڴ棬ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'����������ָ������ڴ�����ĵ�һ���ֽڵ�ָ�룬ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'�ı��ڴ������С��ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
'���ص�ǰ�����ڴ�ߴ��С��ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'��������������Ŀ��ֻ����16���ư汾��Windows���ݡ�
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'��һ���ڴ��һ���ط���������һ���ط� lpvDestĿ�꿽���ĵ�ַָ�롣lpvSourceԴ�����ĵ�ַָ�롣cbCopyԴ�������ֽڴ�С��
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'����ͬ�ϣ�ֻ��ԴΪһ���ַ���
Public Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
'����ͬ�ϣ�ֻ��Ŀ��Ϊһ���ַ���
Public Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)
'���ڷ�����ʾ��Ϣ
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
' ����ָ����Ϣ�����壬�ȴ�������ŷ��أ��� PostMessage() ����������Ϣ���������أ�HWND hWnd Ŀ�괰��ľ����wMsg�����͵���Ϣ��wParam��Ϣ��һ������lParam��Ϣ�ڶ�������
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'����ͬ�ϣ������ڶ�����ΪLong�͡�
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����ͬ�ϣ������ڶ�����ΪString�͡�
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'���ô���״̬����󻯡����»������صȣ�
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'�ƶ�����
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Ҫ����ˢ��
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'����/���������ˢ��
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'���ٴ��弰�����Դ
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'����/�ָ���꼰���̵�����
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'����ָ�������Ĵ���
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
'�ı�ָ������ĸ�����
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'��ȡ��ǰ�������ڴ��壺
'��������5�㣺Frame��Document��Pane��Parent��In-place
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'��ȡָ������ı߽���γߴ�
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'��ȡ�ͻ��������
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'��ȡ��������
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'���ô�������
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'�Ƴ���������
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'���ذ�����ָ����Ĵ��ڵľ����
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'����Ļ��ĳ�������Ļ����ת��Ϊ�ͻ���������
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'��һ��������ص�����ռ�ӳ�䵽��һ�����������ռ�
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'ȡ����겶��
Public Declare Function ReleaseCapture Lib "user32" () As Long
'��ȡ�����Ļ����λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'ָ���ͻ������һ��������ˢ�µľ�������
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'ͬ�ϣ���������2��һ��ָ����
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'����ָ�����ԵĴ���
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'����Ϣ���͵�ָ���Ĵ������
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'�ı�ָ�����������
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function SetWindowLongL Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'��ȡָ�����������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'�ı䴰��λ�á�Zorder���ߴ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'���õ�ǰ�߳���Ϣ�����еĴ����ȡ���̽���
Public Declare Function GetFocus Lib "user32" () As Long
'��ȡ����
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
'��ָ���Ŀ�ִ��ģ�飨.DLL/.EXE��ӳ�䵽���ù��̵ĵ�ַ�ռ�
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'����DLL��������Ŀ
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Integer, ByVal crColor As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'��ȡ������ʾԪ�صĵ�ǰ��ɫֵ
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'���ƾ��ε�һ�����߶�����
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'��һ�� OLE_COLOR ����ת��Ϊһ�� COLORREF ���͡�
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'����һ��ͼ�ꡢ��̬������λͼ��
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'ͬ�ϣ������ڶ�����Ϊһ������ֵ��
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'��ȡ��ʾ�����ߴ�ӡ������Ϣ
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'ֽ�Ŵ�ӡ�߽����
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
'����ָ��������ӳ��ģʽ
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
'��ʼһ����ӡ��ҵ
Public Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
'֪ͨ��ӡ�豸׼���������ݡ�
Public Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
'֪ͨ��ӡ��ֹͣ�������ݣ�ͨ�����ڻ�ҳ
Public Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
'���һ�δ�ӡ��ҵ
Public Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
'ɾ��ָ���豸������������
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'���浱ǰ�豸����״̬�������Ķ�ջ�С�
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
'�ָ��豸����״̬��
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
'ʹ��ָ������ָ���豸�����ӿڵ�ԭ��
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
'��������32λ������Ȼ������64λ������Ե�������������������롣
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'ָ��������л�ͼ׼����ͨ��PAINTSTRUCT�ṹ������ʼ����
Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
'�ڻ�ͼ��ɺ󣬱�Ǵ����ͼ������
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
'���ڻ�ȡ������ͼ�������Ϣ��
'ȡ���ڻ�ͼ����Ĳ�ͬ�������ڸ���������������BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT ���� LOGPEN �ṹ
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'��һ������ѡ��ָ�����豸�������������У��ö����Զ��滻��ͬһ���͵�ǰһ����
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'ɾ��һ���߼����ʡ���ˢ�����塢λͼ��������ߵ�ɫ��
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��ȡ�������ڻ���������Ļ�Ļ����������������ͼ��
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'�ͷű�׼Windows�豸������Դ��
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'�������ݵ��ڴ��豸����
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'�����豸���λͼ
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'����ָ����ɫ���߼���ˢ
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'ʹ��ָ����ˢ����������
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'��Դ������Ŀ�껭���ı��ؿ鴫�����ɫ����
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�������洰�壨��Ļ���ľ��
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'��ȡϵͳ������λ��ϵͳ���ã����гߴ���Ե� Pixel ��ʾ
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Shell����
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ͬ�ϣ�������4��5����ΪAny����
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
'�жϵ�ǰ�Ƿ�ĳ����������»��߷ſ�
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'�˹��ϳ���궯���͵���¼����±�׼Ӧ��ʹ�� SendInput() ������
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'д���ļ�
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'���ļ�
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'��ȡ�ļ�
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'�رն�����
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'API��ͼ
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
'�жϾ�������Ρ���������Բ�Ƿ��ཻ
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
'��չ�� Shell ����
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
        eShowCmd = 0    '����
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then   'Shell ����
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
            lErr = 7: sErr = "�ڴ����"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "�ļ�û���ҵ�"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "·��û���ҵ�"
        Case ERROR_BAD_FORMAT
            sErr = "��Ч�Ŀ�ִ���ļ������Ѿ���"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "·��/�ļ���ȡ����"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "���ļ�û����Ч���ļ�����"
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "�ļ��޷��򿪣�Ŀ�����æ�����Ժ����ԡ�"
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "�ļ��޷��򿪣�DDE����æ�����Ժ����ԡ�"
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "�ļ��޷��򿪣���ʱ�����Ժ����ԡ�"
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "û���ҵ�ָ���Ķ�̬���ӿ⡣"
        Case SE_ERR_FNF
            lErr = 53: sErr = "�ļ�û���ҵ���"
        Case SE_ERR_NOASSOC
            sErr = "û����֮������Ӧ�ó���"
        Case SE_ERR_OOM
            lErr = 7: sErr = "�ڴ����"
        Case SE_ERR_PNF
            lErr = 76: sErr = "·��û���ҵ�"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "����Υ��"
        Case Else
            sErr = "�ڴ򿪻��ߴ�ӡ���ļ�ʱ��������"
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function

'��ȡShift����״̬
Public Function giGetShiftState() As Integer
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
    iR = iR Or (-vbShiftMask * gbKeyIsPressed(VK_SHIFT))
    iR = iR Or (-vbAltMask * gbKeyIsPressed(VK_MENU))
    iR = iR Or (-vbCtrlMask * gbKeyIsPressed(VK_CONTROL))
    giGetShiftState = iR

End Function

'��ȡ��갴��״̬
Public Function giGetMouseButton() As Integer
Dim iR As Integer
   iR = iR Or (-vbLeftButton * gbKeyIsPressed(vbKeyLButton))
   iR = iR Or (-vbMiddleButton * gbKeyIsPressed(vbKeyMButton))
   iR = iR Or (-vbRightButton * gbKeyIsPressed(vbKeyRButton))
   giGetMouseButton = iR
   
End Function

'�ж�ĳ�����Ƿ񱻰���
Public Function gbKeyIsPressed( _
        ByVal nVirtKeyCode As KeyCodeConstants _
    ) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        gbKeyIsPressed = True
    End If
End Function

'��ɫת��
Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

'*************************************************************************
'**�� �� ����HIWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĸ�16λ
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĸ�16λ
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**�� �� ����LOWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĵ�16λ
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĵ�16λ
     LOWORD = LongIn And &HFFFF&
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
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
''Ϊ�˴���˫��ʱ�Ի������ȷ��ʾλ�ã���API������д��һ��MsgBox����
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(ByRef frmParent As Form, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, Optional Title As String = "") As Long

    MsgBoxD = MessageBox(frmParent.hwnd, Prompt, Title, Buttons)

End Function

