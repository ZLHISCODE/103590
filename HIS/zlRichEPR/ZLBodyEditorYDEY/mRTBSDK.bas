Attribute VB_Name = "mRTBSDK"
'#########################################################################
'##ģ �� ����mRTBSDK.bas
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��3��25��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ͨ�õ� RTB SDK API ���� (2.0�汾)
'##��    ����
'#########################################################################

Option Explicit

Public Const LF_FACESIZE = 32   '���������ֽڳ��ȡ�
Public Const RICHEDIT_VER = &H210    '��ǰRich Edit�ؼ��汾��
Public Const cchTextLimitDefault = 32767&       'Ĭ���ı���������
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0
Public Const RICHEDIT_CLASSW = "RichEdit20W"
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSW       'UNICODE�汾��
Public Const WM_CONTEXTMENU = &H7B&     '֪ͨ������Ҽ�����¼�
Public Const WM_PRINTCLIENT = &H318&    '���������ͻ�����һ��ָ�����豸�������У�ͨ����ָ��ӡ����
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
' #ifdef _WIN32
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80) '��ȡ��ǰע�����չ���ִ�����̵ĵ�ַ��
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81) '���õ�ǰ��չ���ִ�����̡�0��ָ�ΪĬ�ϡ�
' #End If

' /* Richedit v2.0 ��Ϣ */
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

' /* enum for use with EM_GET/SETTEXTMODE */    �ı�ģʽ
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                 ' /* Ĭ����Ϊ */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8           ' /* Ĭ����Ϊ */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32           ' /* Ĭ����Ϊ */
End Enum

Public Const EM_AUTOURLDETECT = (WM_USER + 91)      '����/�����Զ�URL��⡣
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)   '�ж��Ƿ��������Զ�URL��⡣
Public Const EM_SETPALETTE = (WM_USER + 93)         '�ı��ɫ�塣
Public Const EM_GETTEXTEX = (WM_USER + 94)          '��ȡָ������ҳ���ı���
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)    '���ò�ͬ��ʽ�����ı����ȡ�

' /* Զ��������Ϣ */
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


' /* BiDi ˫������֧�� ������Ϣ */
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)    '���õ�ǰ˫������֧��ѡ�
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)    '��ȡ��ǰ˫������֧��ѡ�

' /* Options for EM_SETLANGOPTIONS and EM_GETLANGOPTIONS */
Public Const IMF_AUTOKEYBOARD = &H1             '�Զ����̲���
Public Const IMF_AUTOFONT = &H2                 '�Զ�����
Public Const IMF_IMECANCELCOMPLETE = &H4      '// high completes the comp string when aborting, low cancels.
Public Const IMF_IMEALWAYSSENDNOTIFY = &H8

' /* EM_GETIMECOMPMODE ��ȡֵ */
Public Const ICM_NOTOPEN = &H0          'Input Method Editor (IME) is not open.
Public Const ICM_LEVEL3 = &H1           'True inline mode.
Public Const ICM_LEVEL2 = &H2           'Level 2.
Public Const ICM_LEVEL2_5 = &H3         'Level 2.5
Public Const ICM_LEVEL2_SUI = &H4       'Special user interface (UI).

' /* �µ�֪ͨ��Ϣ */

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

' /* BiDi ˫������֧�� ����֪ͨ��Ϣ */

Public Const EN_ALIGN_LTR = &H710&      'RTB�ؼ�ͨ�� WM_COMMAND ��Ϣ֪ͨ��������䷽���Ϊ�������ҡ�
Public Const EN_ALIGN_RTL = &H711&      'RTB�ؼ�ͨ�� WM_COMMAND ��Ϣ֪ͨ��������䷽���Ϊ��������

' /* �¼�֪ͨ���� */

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

' /* Զ���ض�֪ͨ���� */
Public Const ENM_IMECHANGE = &H800000           ' /* RE2.0 ��֧�֣���ֻ����1.0�汾��*/
Public Const ENM_LANGCHANGE = &H1000000         ' ����
Public Const ENM_OBJECTPOSITIONS = &H2000000    '���Է��� EN_OBJECTPOSITIONS ��Ϣ��
Public Const ENM_LINK = &H4000000               '���Է��� EN_LINK ��Ϣ��

' /* �µ� Edit �ؼ���ʽ */

Public Const ES_SAVESEL = &H8000&               '��ʧȥ����ʱ����ѡ�����������ʾ������Useful��
Public Const ES_SUNKEN = &H4000&                '����Ч��
Public Const ES_DISABLENOSCROLL = &H2000&       '�ڲ���Ҫ������ʱ�����ûң���������
' /* same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value */
Public Const ES_SELECTIONBAR = &H1000000
' /* same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop */
Public Const ES_NOOLEDRAGDROP = &H8

' /* �µ� Edit �ؼ���չ��ʽ */
' #ifdef  _WIN32
Public Const ES_EX_NOCALLOLEINIT = &H1000000
' #End If

' /* These flags are used in FE Windows */
Public Const ES_VERTICAL = &H400000     '��ֱ�����ı��Ͷ���
Public Const ES_NOIME = &H80000         '����IME��
Public Const ES_SELFIME = &H40000       'Ӧ�ó���������IME������

' /* �µĶ��ִ����� */
Public Const WB_CLASSIFY = 3&           '
Public Const WB_MOVEWORDLEFT = 4&       '
Public Const WB_MOVEWORDRIGHT = 5&      '
Public Const WB_LEFTBREAK = 6&          '
Public Const WB_RIGHTBREAK = 7&         '

' /* Զ�������־λ */
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

' /* Զ�������־λ */
Public Const IMF_FORCENONE = &H1
Public Const IMF_FORCEENABLE = &H2
Public Const IMF_FORCEDISABLE = &H4
Public Const IMF_CLOSESTATUSWINDOW = &H8
Public Const IMF_VERTICAL = &H20
Public Const IMF_FORCEACTIVE = &H40
Public Const IMF_FORCEINACTIVE = &H80
Public Const IMF_FORCEREMEMBER = &H100
Public Const IMF_MULTIPLEEDIT = &H400

' /* ���ֱ�־λ������WB_CLASSIFY�� */
Public Const WBF_CLASS = &HF          '((BYTE) =&H0F)
Public Const WBF_ISWHITE = &H10       '((BYTE) =&H10)
Public Const WBF_BREAKLINE = &H20     '((BYTE) =&H20)
Public Const WBF_BREAKAFTER = &H40    '((BYTE) =&H40)


' /* �µ��������� */

' #ifdef _WIN32
' /* extended edit word break proc (character set aware) */
'typedef LONG (*EDITWORDBREAKPROCEX)(char *pchText, LONG cchText, BYTE bCharSet, INT action);
' #End If

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

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const CHARFORMAT CHARFORMATW
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* UNICODE */
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

' /* CHARFORMAT ���� */
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

' /* CHARFORMAT Ч�� */
Public Const CFE_BOLD = &H1&            '����
Public Const CFE_ITALIC = &H2&          'б��
Public Const CFE_UNDERLINE = &H4&       '�»���
Public Const CFE_STRIKEOUT = &H8&       'ɾ����
Public Const CFE_PROTECTED = &H10&      '����
Public Const CFE_LINK = &H20&           '������
Public Const CFE_AUTOCOLOR = &H40000000 '����ϵͳ�Զ���ɫ��' /* NOTE: this corresponds to */
                                        ' /* CFM_COLOR, which controls it */
Public Const yHeightCharPtsMost = 1638& '�������ߴ�ֵ����ָY����ߴ磬��λ�������㣩��

' /* EM_SETCHARFORMAT wParam �������� */
Public Const SCF_SELECTION = &H1&   'Ӧ���ڵ�ǰѡ������
Public Const SCF_WORD = &H2&        'Ӧ���ڵ�ǰѡ�е��ʡ�
Public Const SCF_DEFAULT = &H0&            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4&                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8&         '// modifier for SCF_SELECTION; says that
                                   ' // the format came from a toolbar, etc. and
                                   ' // therefore UI formatting rules should be
                                   ' // used instead of strictly formatting the
                                   ' // selection.


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

'���ڴ洢 EM_STREAMIN ���� EM_STREAMOUT ��Ϣ���ݵ�������Ϣ��
Public Type EDITSTREAM
    dwCookie As Long     ' /* user value passed to callback as first parameter */
    dwError As Long      ' /* last error */
    pfnCallback As Long  'EDITSTREAMCALLBACK
End Type

' /* ���ĸ�ʽ */

Public Const SF_TEXT = &H1         'Text��ʽ
Public Const SF_RTF = &H2          'RTF��ʽ
Public Const SF_RTFNOOBJS = &H3    '���ʱ�ÿո������󣬽����������
Public Const SF_TEXTIZED = &H4     '���ʱ�����ı���ʾ���󣬽����������
Public Const SF_UNICODE = &H10            ' /* Unicode file of some kind */

' /* Flag telling stream operations to operate on the selection only */
' /* EM_STREAMIN will replace the current selection */
' /* EM_STREAMOUT will stream out the current selection */
Public Const SFF_SELECTION = &H8000&    '�������ֻ�Ե�ǰѡ��������Ч��

' /* Flag telling stream operations to operate on the common RTF keyword only */
' /* EM_STREAMIN will accept the only common RTF keyword */
' /* EM_STREAMOUT will stream out the only common RTF keyword */
Public Const SFF_PLAINRTF = &H4000&     'ֻʹ��ͨ��RTF�ؼ��֣�������������ص�RTF�ؼ������Ժ��ԣ�

'���� EM_FINDTEXT ��Ϣ�Ĳ����ı��������Ϣ
Public Type FindText
    chrg As CHARRANGE   '�ַ���Χ
    lpstrText As Long   '��Ҫ���ҵ��ı�
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

'�������ڸ�ʽ��ָ���豸�������Ϣ
Public Type FORMATRANGE
    hDC As Long             '��Ⱦ�豸
    hdcTarget As Long       'Ŀ���豸
    rc As RECT              '��Ⱦ���򣬵�λ��羡�
    rcPage As RECT          '��Ⱦ�豸���������򣬵�λ��羡�
    chrg As CHARRANGE       '���ڸ�ʽ�����ı���Χ��
End Type

' /* ���ж��������λ��Ϊ��� */

Public Const MAX_TAB_STOPS = 32&    '�����Ʊ���������Ŀ��
Public Const lDefaultTab = 720&     'Ĭ�Ͼ����Ʊ��λ�á�

'�����ʽ
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

' /* PARAFORMAT ����ֵ */
Public Const PFM_STARTINDENT = &H1& '��������ֵ��Ч��
Public Const PFM_RIGHTINDENT = &H2& '������ֵ��Ч��
Public Const PFM_OFFSET = &H4&      '��������������Ч����ֵ��ʾ��������ֵ��ʾ���ң�
Public Const PFM_ALIGNMENT = &H8&   'ˮƽ���뷽ʽ��Ч��
Public Const PFM_TABSTOPS = &H10&   '�����Ʊ��λ����Ч��
Public Const PFM_NUMBERING = &H20&  '�������Ŀ������Ч��
Public Const PFM_OFFSETINDENT = &H80000000  '��������ֵ��Ч�����Ҹ���һ�����ֵ��

' /* PARAFORMAT ���ѡ�� */
Public Const PFN_BULLET = &H1&      '

' /* PARAFORMAT ����ѡ�� */
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

' #endif ' /* C++ */

' #ifdef UNICODE
'public const CHARFORMAT2 CHARFORMAT2W
' #Else
'public const CHARFORMAT2 CHARFORMAT2A
' #End If

'public Const CHARFORMATDELTA = (Len(CHARFORMAT2) - Len(CHARFORMAT))


' /* CHARFORMAT and PARAFORMAT "ALL" masks
'   CFM_COLOR mirrors CFE_AUTOCOLOR, a little hack to easily deal with autocolor*/

'ӳ��Ϊ����������Ч��
Public Const CFM_EFFECTS = (CFM_BOLD Or CFM_ITALIC Or CFM_UNDERLINE Or CFM_COLOR Or _
                     CFM_STRIKEOUT Or CFE_PROTECTED Or CFM_LINK)
Public Const CFM_ALL = (CFM_EFFECTS Or CFM_SIZE Or CFM_FACE Or CFM_OFFSET Or CFM_CHARSET)

' /* �µ������Ч�� �� (*)��ʾ������RichEdit 2.0�б��棬���ǲ�����ʾ��

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

'ӳ��Ϊ����������Ч��
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

' #endif ' /* C++   */

' /* PARAFORMAT 2.0 �����Ч�� */

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
Public Const PFE_TABLECELL = &H4000&               ' /*  ����Ϊ����һ���� */

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


Public Const PFA_JUSTIFY = 4          ' /* ���˶��룬Ϊ�˼���TOMģ�ͽӿڡ� (*)  */


' /* ֪ͨ�Ľṹ */
Public Type NMHDR
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
    Msg As Integer          '���̻�������ʶ��
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

' /* used with IRichEditOleCallback::GetContextMenu, this flag will be
'   passed as a "selection type".  It indicates that a context menu for
'   a right-mouse drag drop should be generated.  The IOleObject parameter
'   will really be the IDataObject for the drop
' */
' ������ IRichEditOleCallback::GetContextMenu ����������Ӧ�ó����ṩһ���Ҽ��˵���
Public Const GCM_RIGHTMOUSEDROP = &H8000&

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
    Msg As Long                 '������֪ͨ��ԭʼ��Ϣ
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
' #ifndef MACPORT
Public Type ENOLEOPFAILED
    NMHDR As NMHDR     '֪ͨͷ
    iob As Long                 '��������ֵ
    lOper As Long               'ʧ�ܵ�OLE������ȡֵΪ OLEOP_DOVERB ����
    hr As Long                  '���صĴ������
End Type
' #End If

Public Const OLEOP_DOVERB = 1

'����λ��Ϣ���ڶ��󱻶���RTBʱ������֪ͨ
Public Type OBJECTPOSITIONS
    NMHDR As NMHDR     '֪ͨͷ
    cObjectCount As Long        '��������
        ' !!!POINTER to long value!!!
    pcpPositions As Long        '����λ��ָ�롣ע�⣺�ǳ����ε�ָ�룡������
End Type

Public Type ENLINK
    NMHDR As NMHDR     '֪ͨͷ
    Msg As Integer              '������֪ͨ����Ϣ
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


' �������ʽ������ RegisterClipboardFormat() ע����Ч�ļ������ʽ��
Public Const CF_RTF = "Rich Text Format"
Public Const CF_RTFNOOBJS = "Rich Text Format Without Objects"
Public Const CF_RETEXTOBJ = "RichEdit Text and Objects"

' ѡ����ճ��
Public Type REPASTESPECIAL
    dwAspect As Long    '��ʾ���ԡ�ȡֵ��DVASPECT_CONTENT ���� DVASPECT_ICON
    dwParam As Long     '���ΪDVASPECT_ICON���򱾲�������һ��ָ��ö�����ͼ��һ��ͼԪ�ļ����
End Type


' /* ��������� GETTEXTEX ���ݽṹ */
Public Const GT_DEFAULT = 0&    '��ʹ��CRת��
Public Const GT_USECRLF = 1&    '��ʾ��ÿ�ο����ı�ʱ����CRת��ΪCRLF��

' /* EM_GETTEXTEX ��Ϣ wParam ���� */
Public Type GETTEXTEX
    cb As Long              ' /* ��ȡ���ַ����ֽ���             */
    flags As Long           ' /* �ı�ת������ѡ��               */
    codepage As Long        ' /* ת���Ĵ���ҳ��Ĭ��ΪCP_ACP��UnicodeΪ1200
    lpDefaultChar As Long   ' /* ��Unicodeģʽ���޷���ʾ���ַ�ʱ������ַ���ΪNULL��ʹ��ϵͳĬ��ֵ�� */
    lpUsedDefChar As Long   ' /* �Ƿ������滻�ַ�   */
End Type

' GETTEXTLENGTHEX ���ݽṹ�ı�־λ
Public Const GTL_DEFAULT = 0&      ' /* Ĭ��ֵ�������ַ���Ŀ��                      */
Public Const GTL_USECRLF = 1&      ' /* ʹ�ö��� CR/LF ����                         */
Public Const GTL_PRECISE = 2&      ' /* ��ȷ���㣬����                              */
Public Const GTL_CLOSE = 4&        ' /* ���Ƽ��㣬�Ͽ죬��������ǰ�����ڴ�ռ�      */
Public Const GTL_NUMCHARS = 8&     ' /* �����ַ���Ŀ                                */
Public Const GTL_NUMBYTES = 16&    ' /* �����ֽ���Ŀ                                */

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

' /* ������ EM_FINDTEXT[EX] ��־ */
Public Const FR_MATCHDIAC = &H20000000          ' ��������ϣ��������
Public Const FR_MATCHKASHIDA = &H40000000       ' ��������ϣ��������
Public Const FR_MATCHALEFHAMZA = &H80000000     ' ��������ϣ��������

' /* UNICODE Ƕ���ַ� */
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


' Edit �ؼ���Ϣ��
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
'#if(WINVER >= =&H0400)
Public Const EM_SETMARGINS = &HD3&          '�������Ҽ�࣬��ˢ�¡�
Public Const EM_GETMARGINS = &HD4&          '��ȡ...
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT '�����ַ���󳤶ȡ� ' /* ;win40 Name change */
Public Const EM_GETLIMITTEXT = &HD5&        '��ȡ�ַ���󳤶ȡ�
Public Const EM_POSFROMCHAR = &HD6&         '��ȡָ���ַ�������(X,Y)��
Public Const EM_CHARFROMPOS = &HD7&         '��ȡָ������㸽�����ַ���

Public Const EC_LEFTMARGIN = &H1            '��ʾ��������߽硣
Public Const EC_RIGHTMARGIN = &H2           '��ʾ�������ұ߽硣
Public Const EC_USEFONTINFO = &HFFFF&       '�߽�����ַ���ȡ�
'#End If ' /* WINVER >= =&H0400 */
'/*
' * Edit �ؼ���ʽ
' */
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
'#if(WINVER >= =&H0400)
Public Const ES_NUMBER = &H2000&        'ֻ�����������֡�
'#endif /* WINVER >= =&H0400 */

'/* Edit �ؼ�֪ͨ��Ϣ */
Public Const EN_CHANGE = &H300          '���ݸı䡣������ͨ�� WM_COMMAND ��Ϣ��ȡ��֪ͨ��
Public Const EN_ERRSPACE = &H500        '���ݲ����Է���ò�����
Public Const EN_HSCROLL = &H601         'ˮƽ�����¼���
Public Const EN_KILLFOCUS = &H200       'ʧȥ�����¼���
Public Const EN_MAXTEXT = &H501         '������ı���������ַ����������ڷ��Զ�����ʱ�����ؼ���������
Public Const EN_SETFOCUS = &H100        '��ü������뽹�㡣
Public Const EN_UPDATE = &H400          '���û��ı����ݵ��ǻ�û��ˢ����ʾʱ������֪ͨ���û��������ڵ��ڿؼ��ߴ�����Ӧ���ݡ�
Public Const EN_VSCROLL = &H602         '��ֱ�����¼���



