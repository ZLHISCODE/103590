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
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'����ͬ�ϣ������ڶ�����ΪLong�͡�
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����ͬ�ϣ������ڶ�����ΪString�͡�
Public Declare Function SendMessageStr Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'���ô���״̬����󻯡����»������صȣ�
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'�ƶ�����
Public Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Ҫ����ˢ��
Public Declare Function UpdateWindow Lib "User32" (ByVal hWnd As Long) As Long
'����/���������ˢ��
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
'���ٴ��弰�����Դ
Public Declare Function DestroyWindow Lib "User32" (ByVal hWnd As Long) As Long
'����/�ָ���꼰���̵�����
Public Declare Function EnableWindow Lib "User32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'����ָ�������Ĵ���
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'�ı�ָ������ĸ�����
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'��ȡ��ǰ�������ڴ��壺
'��������5�㣺Frame��Document��Pane��Parent��In-place
Public Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'��ȡָ������ı߽���γߴ�
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
'��ȡ�ͻ��������
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
'��ȡ��������
Public Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'���ô�������
Public Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'�Ƴ���������
Public Declare Function RemoveProp Lib "User32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'���ذ�����ָ����Ĵ��ڵľ����
Public Declare Function WindowFromPointXY Lib "User32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'����Ļ��ĳ�������Ļ����ת��Ϊ�ͻ���������
Public Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'��һ��������ص�����ռ�ӳ�䵽��һ�����������ռ�
Public Declare Function MapWindowPoints Lib "User32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
'ȡ����겶��
Public Declare Function ReleaseCapture Lib "User32" () As Long
'��ȡ�����Ļ����λ��
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
'ָ���ͻ������һ��������ˢ�µľ�������
Public Declare Function InvalidateRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'ͬ�ϣ���������2��һ��ָ����
Public Declare Function InvalidateRectAsNull Lib "User32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'����ָ�����ԵĴ���
Public Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'����Ϣ���͵�ָ���Ĵ������
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'�ı�ָ�����������
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'��ȡָ�����������
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'�ı䴰��λ�á�Zorder���ߴ��
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'���õ�ǰ�߳���Ϣ�����еĴ����ȡ���̽���
Public Declare Function GetFocus Lib "User32" () As Long

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
Public Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'��ָ���Ŀ�ִ��ģ�飨.DLL/.EXE��ӳ�䵽���ù��̵ĵ�ַ�ռ�
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'����DLL��������Ŀ
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Integer, ByVal crColor As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'#########################################################################
' ͼ�κ�������

'��ȡ������ʾԪ�صĵ�ǰ��ɫֵ
Public Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
'���ƾ��ε�һ�����߶�����
Public Declare Function DrawEdge Lib "User32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'��һ�� OLE_COLOR ����ת��Ϊһ�� COLORREF ���͡�
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'����һ��ͼ�ꡢ��̬������λͼ��
Public Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'ͬ�ϣ������ڶ�����Ϊһ������ֵ��
Public Declare Function LoadImageLong Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

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
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const HORZRES = 8            '  Horizontal width in pixels

Public Const VERTRES = 10           '  Vertical width in pixels

Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X

Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin

Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin

Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units

'����ָ��������ӳ��ģʽ
Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
'��ʼһ����ӡ��ҵ
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
'֪ͨ��ӡ�豸׼���������ݡ�
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
'֪ͨ��ӡ��ֹͣ�������ݣ�ͨ�����ڻ�ҳ
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
'���һ�δ�ӡ��ҵ
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
'ɾ��ָ���豸������������
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'���浱ǰ�豸����״̬�������Ķ�ջ�С�
Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
'�ָ��豸����״̬��
Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
'ʹ��ָ������ָ���豸�����ӿڵ�ԭ��
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long

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
    hdc As Long
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
   hdc As Long
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
Public Declare Function BeginPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'�ڻ�ͼ��ɺ󣬱�Ǵ����ͼ������
Public Declare Function EndPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'���ڻ�ȡ������ͼ�������Ϣ��
'ȡ���ڻ�ͼ����Ĳ�ͬ�������ڸ���������������BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT ���� LOGPEN �ṹ
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'��һ������ѡ��ָ�����豸�������������У��ö����Զ��滻��ͬһ���͵�ǰһ����
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'ɾ��һ���߼����ʡ���ˢ�����塢λͼ��������ߵ�ɫ��
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��ȡ�������ڻ���������Ļ�Ļ����������������ͼ��
Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
'�ͷű�׼Windows�豸������Դ��
Public Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'�������ݵ��ڴ��豸����
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'�����豸���λͼ
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'����ָ����ɫ���߼���ˢ
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'ʹ��ָ����ˢ����������
Public Declare Function FillRect Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'��Դ������Ŀ�껭���ı��ؿ鴫�����ɫ����
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'�������洰�壨��Ļ���ľ��
Public Declare Function GetDesktopWindow Lib "User32" () As Long
'��ȡϵͳ������λ��ϵͳ���ã����гߴ���Ե� Pixel ��ʾ
Public Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
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
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer

' ��������볣��
Public Const VK_SHIFT = &H10&               'Shift
Public Const VK_CONTROL = &H11&             'Ctl
Public Const VK_MENU = &H12&                'Alt

'�˹��ϳ���궯���͵���¼����±�׼Ӧ��ʹ�� SendInput() ������
Declare Sub mouse_event Lib "User32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
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
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

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

Declare Function DrawTextEx Lib "User32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal N As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

'######################################################################################
'   �ͷ��ڴ�
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "Psapi" (ByVal hProcess As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
'################################################################################################################
'## ͼƬ����ģʽ����
'######################################################################################
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'#########################################################################
'��չ�� Shell ����
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
        eShowCmd = 0    '����
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then   'Shell ����
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







