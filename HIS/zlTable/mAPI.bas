Attribute VB_Name = "mAPI"
'#########################################################################
'##��    ����ͨ�� Windows API ����
'#########################################################################

Option Explicit

'#########################################################################
'����
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'��
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'����λ����Ϣ
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI       '��󻯳ߴ�
    ptMaxPosition As POINTAPI   '���λ��
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

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
Public Const WM_USER = &H400                'ͨ���� WM_USER + X ���Զ�����Ϣ
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_SETTINGCHANGE = &H1A&
Public Const WM_DISPLAYCHANGE = &H7E&

'#########################################################################
' �ڴ��������:

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

'��һ���ڴ��һ���ط���������һ���ط�
'����ԭ�ͣ�
'VOID CopyMemory(
'  PVOID Destination,  // Ŀ�꿽���ĵ�ַָ�롣
'  CONST VOID *Source, // Դ�����ĵ�ַָ�롣
'  DWORD Length        // Դ�������ֽڴ�С��
')
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
'����ͬ�ϣ�ֻ��ԴΪһ���ַ���
Public Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
    
'����ͬ�ϣ�ֻ��Ŀ��Ϊһ���ַ���
Public Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)

'#########################################################################
' ��ͨ WinAPI ����:

' ����ָ����Ϣ�����壬�ȴ�������ŷ��أ��� PostMessage() ����������Ϣ���������أ�
'����ԭ�ͣ�
'LRESULT SendMessage(
'  HWND hWnd,      // Ŀ�괰��ľ����
'  UINT Msg,       // �����͵���Ϣ��
'  WPARAM wParam,  // ��Ϣ��һ������
'  LPARAM lParam   // ��Ϣ�ڶ�������
');
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
'����ͬ�ϣ������ڶ�����ΪLong�͡�
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
'����ͬ�ϣ������ڶ�����ΪString�͡�
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As String) As Long

'���ô���״̬����󻯡����»������صȣ�
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'�ƶ�����
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Ҫ����ˢ��
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
'����/���������ˢ��
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'���ٴ��弰�����Դ
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'����/�ָ���꼰���̵�����
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'����ָ�������Ĵ���
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'�ı�ָ������ĸ�����
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'��ȡ��ǰ�������ڴ��壺
'��������5�㣺Frame��Document��Pane��Parent��In-place
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'��ȡָ������ı߽���γߴ�
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'��ȡ�ͻ��������
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'��ȡ��������
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'���ô�������
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'�Ƴ���������
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'���ذ�����ָ����Ĵ��ڵľ����
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'����Ļ��ĳ�������Ļ����ת��Ϊ�ͻ���������
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'��һ��������ص�����ռ�ӳ�䵽��һ�����������ռ�
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'ȡ����겶��
Public Declare Function ReleaseCapture Lib "user32" () As Long
'��ȡ�����Ļ����λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'ָ���ͻ������һ��������ˢ�µľ�������
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'ͬ�ϣ���������2��һ��ָ����
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'����ָ�����ԵĴ���
Public Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'����Ϣ���͵�ָ���Ĵ������
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
'�ı�ָ�����������
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'��ȡָ�����������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'�ı䴰��λ�á�Zorder���ߴ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'���õ�ǰ�߳���Ϣ�����еĴ����ȡ���̽���
Public Declare Function GetFocus Lib "user32" () As Long

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


'��ȡ����
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'��ָ���Ŀ�ִ��ģ�飨.DLL/.EXE��ӳ�䵽���ù��̵ĵ�ַ�ռ�
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'����DLL��������Ŀ
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


'#########################################################################
' ͼ�κ�������

'��ȡ������ʾԪ�صĵ�ǰ��ɫֵ
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'���ƾ��ε�һ�����߶�����
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'��һ�� OLE_COLOR ����ת��Ϊһ�� COLORREF ���͡�
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'����һ��ͼ�ꡢ��̬������λͼ��
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'ͬ�ϣ������ڶ�����Ϊһ������ֵ��
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'��ʾһ��Windowsλͼ��ʽ��
Public Const CF_BITMAP = 2
'3DЧ����ɫ
Public Const LR_LOADMAP3DCOLORS = &H1000
'ͼƬ���ļ�lpsz�е��룬���Ǵ���Դ�ļ��е��롣
Public Const LR_LOADFROMFILE = &H10
'����͸��ɫ
Public Const LR_LOADTRANSPARENT = &H20
'���� �豸�޹� DIB λͼ�������豸���λͼ��
Public Const IMAGE_BITMAP = 0

'��ȡ��ʾ�����ߴ�ӡ������Ϣ
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'
'Public Const HORZRES = 8            '  Horizontal width in pixels
'
'Public Const VERTRES = 10           '  Vertical width in pixels
'
'Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
'
'Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
'
'Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
'
'Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
'
'Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units
'
'Public Const PHYSICALWIDTH = 110 '  Physical Width in device units

'����ָ��������ӳ��ģʽ
Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
'��ʼһ����ӡ��ҵ
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
'֪ͨ��ӡ�豸׼���������ݡ�
Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
'֪ͨ��ӡ��ֹͣ�������ݣ�ͨ�����ڻ�ҳ
Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
'���һ�δ�ӡ��ҵ
Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
'ɾ��ָ���豸������������
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'���浱ǰ�豸����״̬�������Ķ�ջ�С�
Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
'�ָ��豸����״̬��
Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
'ʹ��ָ������ָ���豸�����ӿڵ�ԭ��
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long

'ÿ���߼���λΪ1���豸���ء���X���ң���Y���¡�����SetMapMode()
Public Const MM_TEXT = 1

'��������32λ������Ȼ������64λ������Ե�������������������롣
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long


'#########################################################################
' ��ӡ֧��:

' VB API Viewer �汾�� DocInfo �ṹ˵���Ǵ���ģ�������
' VB API VIEWER VERSION OF DOCINFO STRUCTURE IS WRONG!
'���ڴ洢 StartDoc() ���ļ�����������Ϣ
Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

'���ڳ�ʼ����ӡ�Ի��򼰷���ֵ
Type PrintDlg
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

'�û����ϵͳ�˵��еġ��ƶ����˵��¼�
Public Const SC_MOVE = &HF012

'ϵͳĬ����ɫ
Public Const COLOR_WINDOWFRAME = 6  '������
Public Const COLOR_BTNFACE = 15     '��ť����
Public Const COLOR_BTNTEXT = 18     '��ť��ͨ�ı�

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
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'ָ��������л�ͼ׼����ͨ��PAINTSTRUCT�ṹ������ʼ����
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'�ڻ�ͼ��ɺ󣬱�Ǵ����ͼ������
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'���ڻ�ȡ������ͼ�������Ϣ��
'ȡ���ڻ�ͼ����Ĳ�ͬ�������ڸ���������������BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT ���� LOGPEN �ṹ
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'��һ������ѡ��ָ�����豸�������������У��ö����Զ��滻��ͬһ���͵�ǰһ����
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'ɾ��һ���߼����ʡ���ˢ�����塢λͼ��������ߵ�ɫ��
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��ȡ�������ڻ���������Ļ�Ļ����������������ͼ��
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'�ͷű�׼Windows�豸������Դ��
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'�������ݵ��ڴ��豸����
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'�����豸���λͼ
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'����ָ����ɫ���߼���ˢ
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'ʹ��ָ����ˢ����������
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'��Դ������Ŀ�껭���ı��ؿ鴫�����ɫ����
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�������洰�壨��Ļ���ľ��
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'��ȡϵͳ������λ��ϵͳ���ã����гߴ���Ե� Pixel ��ʾ
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'ˮƽ�������ϵ�ʸ��λͼ���
Public Const SM_CYHSCROLL = 3
'ˮƽ�������ϵ�ʸ��λͼ�߶�
Public Const SM_CXVSCROLL = 2


'#########################################################################
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
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_STATICEDGE = &H20000

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
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'#########################################################################
' ��꼤����Ӧ
Public Const MA_ACTIVATE = 1                '����CWnd
Public Const MA_ACTIVATEANDEAT = 2          '����CWnd����������¼�
Public Const MA_NOACTIVATE = 3              '������CWnd
Public Const MA_NOACTIVATEANDEAT = 4        '������CWnd����������¼�

Public Const H_MAX As Long = &HFFFF + 1     '���ֵ
 
'Shell����
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ͬ�ϣ�������4��5����ΪAny����
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

'�жϵ�ǰ�Ƿ�ĳ����������»��߷ſ�
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' ��������볣��
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt

'�˹��ϳ���궯���͵���¼����±�׼Ӧ��ʹ�� SendInput() ������
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000  '�����ƶ�
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4       '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1         '����ƶ�
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10     '  right button up

'�����첽����/��� I/O
Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
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

'#########################################################################
'����֧��:

'д���ļ�
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'���ļ�
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'��ȡ�ļ�
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
'�رն�����
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
'API��ͼ
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

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

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

'�жϾ�������Ρ���������Բ�Ƿ��ཻ
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

Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

'######################################################################################
'��ȡ��Ӣ�Ļ���ַ�������
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'����ת��
Public Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

'######################################################################################

Public Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

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

Public Declare Function GradientFill Lib "msimg32" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_TRIANGLE, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'���λ����Ϣ
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Type Size
    cx As Long
    cy As Long
End Type
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Used for creating the temporary WMF file
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MM_ANISOTROPIC = 8 ' Map mode anisotropic
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
'VB Errors
Public Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Public Const DSna = &H220326

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '��׽����״̬
    ' Virtual key values
Public Const VK_TAB = &H9

Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, LParam As Any) As Long

'######################################################################################
'   ����������
'######################################################################################

Public Type POINTL
    X As Long
    Y As Long
End Type

Public Const WM_MOUSEWHEEL = &H20A

Public lpPrevWndProc As Long

Public sngX As Single, sngY As Single   '�������
Public intShift As Integer              '��갴��
Public bWay As Boolean                  '��귽��
Public bMouseFlag As Boolean            '����¼������־

'######################################################################################
'   ��ȡ�ַ���Ļλ��
'######################################################################################
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const S_FALSE = &H1
Public Const S_OK = &H0

Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

'######################################################################################
'   ֱ�ӷ��Ͱ����ĺ���
'######################################################################################
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'######################################################################################
'   ���뷨������
'######################################################################################
'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'######################################################################################
'   �ͷ��ڴ�
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

'######################################################################################
'������GDI�����͸�������
'######################################################################################
Public Const BITSPIXEL = 12         '  Number of bits per pixel
Public Const DT_NOFULLWIDTHCHARBREAK = &H80000
Public Const DT_HIDEPREFIX = &H100000
Public Const DT_PREFIXONLY = &H200000
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Public Const LF_FACESIZE = 32

Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const BALTIC_CHARSET = 186
Public Const RUSSIAN_CHARSET = 204
Public Const THAI_CHARSET = 222
Public Const EE_CHARSET = 238
Public Const OEM_CHARSET = 255

'��������
Public Type LOGFONT
    lfHeight As Long         ' ����ߴ� (������)
    lfWidth As Long          ' ͨ������������,��Windows����Ĭ�ϵ�
    lfEscapement As Long     ' �Ƕ�,����0.1��Ϊ��λ
    lfOrientation As Long    ' �����Ĭ��ֵ
    lfWeight As Long         ' ���塢���֡������        FW_DONTCARE/FW_THIN/FW_EXTRALIGHT/FW_ULTRALIGHT/FW_LIGHT/...
    lfItalic As Byte         ' б��
    lfUnderline As Byte      ' �»���
    lfStrikeOut As Byte      ' ɾ����
    lfCharSet As Byte        ' �ַ���        ANSI_CHARSET/CHINESEBIG5_CHARSET/GB2312_CHARSET/SYMBOL_CHARSET/...
    lfOutPrecision As Byte   ' �����Ĭ��ֵ
    lfClipPrecision As Byte  ' �����Ĭ��ֵ
    lfQuality As Byte        ' �����Ĭ��ֵ
    lfPitchAndFamily As Byte ' �����Ĭ��ֵ
    lfFaceName(LF_FACESIZE) As Byte  ' ת��Ϊ�������������
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImagelist As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImagelist As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long '�ж�ָ�����Ƿ���ָ�������У�����


Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
   ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
     Public Const R2_BLACK = 1 ' 0
     Public Const R2_COPYPEN = 13 ' P
     Public Const R2_LAST = 16
     Public Const R2_MASKNOTPEN = 3 ' DPna
     Public Const R2_MASKPEN = 9 ' DPa
     Public Const R2_MASKPENNOT = 5 ' PDna
     Public Const R2_MERGENOTPEN = 12    ' DPno
     Public Const R2_MERGEPEN = 15 ' DPo
     Public Const R2_MERGEPENNOT = 14    ' PDno
     Public Const R2_NOP = 11    ' D
     Public Const R2_NOT = 6 ' Dn
     Public Const R2_NOTCOPYPEN = 4 ' PN
     Public Const R2_NOTMASKPEN = 8 ' DPan
     Public Const R2_NOTMERGEPEN = 2 ' DPon
     Public Const R2_NOTXORPEN = 10 ' DPxn
     Public Const R2_WHITE = 16 ' 1
     Public Const R2_XORPEN = 7 ' DPx

Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0

Public Const DEFAULT_QUALITY = 0           ' Ĭ����������
Public Const DRAFT_QUALITY = 1             ' Appearance is less important that PROOF_QUALITY.
Public Const PROOF_QUALITY = 2             ' ����ַ�����
Public Const NONANTIALIASED_QUALITY = 3    ' Don't smooth font edges even if system is set to smooth font edges
Public Const ANTIALIASED_QUALITY = 4       ' Ensure font edges are smoothed if system is set to smooth font edges
Public Const CLEARTYPE_QUALITY = 5

Public Const DEFAULT_PITCH = 0

Public Const CLR_INVALID = -1

'������ DrawState ����������
'DrawState:������ʾ��ͬ״̬��ͼ�񣬱��硰���񡱡���������������ɫ����Ч��
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hDC As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal LParam As Long, _
   ByVal wParam As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long
   
'DrawStateString:ͬ��
Public Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hDC As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' ��ͼ״̬����������
'/* ͼ������ */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* ״̬���� */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

' ����һ�� ImageList
Public Declare Function ImageList_Create Lib "comctl32.dll" ( _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal fMask As Long, _
        ByVal cInitial As Long, _
        ByVal cGrow As Long _
    ) As Long
Public Const ILC_MASK = 1&
Public Const ILC_COLOR = 0&
Public Const ILC_COLORDDB = &HFE&
Public Const ILC_COLOR4 = &H4&
Public Const ILC_COLOR8 = &H8&
Public Const ILC_COLOR16 = &H10&
Public Const ILC_COLOR24 = &H18&
Public Const ILC_COLOR32 = &H20&
Public Const ILC_PALETTE = &H800&

Public Declare Function ImageList_Destroy Lib "comctl32.dll" ( _
        ByVal hIml As Long _
    ) As Long

' ���һ������λͼ��ImageList
Public Declare Function ImageList_AddMasked Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal crMask As Long _
    ) As Long
    
' ����һ��ImageListͼ�괴��һ���µ�ͼ��
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
    
' ��һ��ImageList�л���һ����Ŀ
Public Declare Function ImageList_Draw Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal fStyle As Long _
    ) As Long
    
' ͬ�ϣ�������λ�ú���ɫ���и������
Public Declare Function ImageList_DrawEx Lib "comctl32.dll" ( _
      ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal rgbBk As Long, _
      ByVal rgbFg As Long, _
      ByVal fStyle As Long _
   ) As Long
   
' ImageList_Draw �����ĳ�����
Public Const ILD_NORMAL = 0            '����ImageList�ı���ɫ��ͼ
Public Const ILD_TRANSPARENT = 1       '��������ɫ������͸��λͼ
Public Const ILD_BLEND25 = 2           '��������ɫ������25%͸���ȵ�λͼ
Public Const ILD_SELECTED = 4          '��������ɫ������50%͸���ȵ�λͼ
Public Const ILD_FOCUS = 4             'ͬ25%͸����
Public Const ILD_OVERLAYMASK = 3840    '�ص�ͼ��
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' ��׼�� GDI ����ͼ����߹��ĺ�����
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Const DI_MASK = &H1         '��������������ͼ�����
Public Const DI_IMAGE = &H2        '����ͼ��������ͼ�����
Public Const DI_NORMAL = &H3       '���ͼ�������
Public Const DI_COMPAT = &H4       '��ϵͳĬ��ͼ��
Public Const DI_DEFAULTSIZE = &H8  '����ϵͳĬ�ϴ�С

Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
' XP���
Public Declare Function GetVersion Lib "kernel32" () As Long   '��ȡ��ǰϵͳ�汾��

Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lhDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCMBUTTONDOWN = &HA7
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

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

Public Const EM_EXGETSEL = (WM_USER + 52)       '��ȡѡ�е���ʼ����ֹ�ַ�λ�á�
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetTextJustification Lib "gdi32" (ByVal hDC As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Public Const ES_SUNKEN = &H4000&                '����Ч��
Public Const ES_NOHIDESEL = &H100&      'ʧȥ����ʱ����ѡ�����ݡ�

'################################################################################################################
'## ͼƬ����ģʽ����
'######################################################################################
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE

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

Public Declare Function AlphaBlend Lib "MSIMG32.dll" ( _
  ByVal hDCDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long
