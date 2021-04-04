Attribute VB_Name = "mdlPrint"
Option Explicit

'Window�汾����
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

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

Public Const conBin1 = "�ϲ�ֽ�н�ֽ"
Public Const conBin2 = "�²�ֽ�н�ֽ"
Public Const conBin3 = "�м�ֽ�н�ֽ"
Public Const conBin4 = "�ȴ��ֶ�����ÿҳֽ"
Public Const conBin5 = "�ŷ��ֽ����ֽ"
Public Const conBin6 = "�ŷ��ֽ����ֽ����Ҫ�ȴ��ֶ�����"
Public Const conBin7 = "��ǰȱʡֽ�н�ֽ"
Public Const conBin8 = "������ֽ����ֽ"
Public Const conBin9 = "С�ͽ�ֽ����ֽ"
Public Const conBin10 = "����ֽ�н�ֽ"
Public Const conBin11 = "��������ֽ����ֽ"

'ֽ�Ŵ�ӡ�߽����================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'��ͬ��ӡ���Ĵ�ӡ��Ԫ���Ȳ�ͬ

Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

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


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'***************************************************************
'�滭���API���������ṹ����
'***************************************************************
'�ṹ����

'�������
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'���ָ߶ȺͿ��
Private Type Size
    W   As Long
    H   As Long
End Type

'����
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'����
Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type
'ˢ��
Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

'��������
Public Const LF_FACESIZE = 32

Public Type LogFont
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public T_OldPoint   As POINTAPI
Public T_NewPoint   As POINTAPI
Public T_ClientRect As RECT
Public T_LableRect  As RECT      '������ı�����Ч����
Public T_Brush      As LOGBRUSH
Public T_Font       As LogFont
Public T_Size       As Size

'������õ����ж���
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, _
                            lpBits As Any) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'�������ʡ�ˢ��
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, _
                            lpStyle As Long) As Long

Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5                    '������ͼ

Public Const PS_COSMETIC = &H0
Public Const PS_GEOMETRIC = &H10000
Public Const PS_ALTERNATE = 8
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MASK = &HF000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0

'CreateSolidBrush ������ɫ��ˢ
'CreateBrushIndirect ͨ�� LOGBRUSH ���ʹ�����ˢ
'CreateHatchBrush ������Ӱ��ˢ
'CreatePatternBrush ����ͼ����ˢ
'GetSysColorBrush ����ϵͳ��׼ɫ��ˢ
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

'//lbStyle����ѡֵ:
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_MONOPATTERN = 9

'//lbHatch����ѡֵ:
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx

Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'nIndex,ͬ���溯����lbHatch
'Public Const HS_HORIZONTAL = 0              '  -----
'Public Const HS_VERTICAL = 1                '  |||||
'Public Const HS_FDIAGONAL = 2               '  \\\\\
'Public Const HS_BDIAGONAL = 3               '  /////
'Public Const HS_CROSS = 4                   '  +++++
'Public Const HS_DIAGCROSS = 5               '  xxxxx

Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'BLACK_BRUSH����ɫ����
'DKGRAY_BRUSH������ɫ����
'GRAY_BRUSH����ɫ����
'HOLLOW_BRUSH���ջ��ʣ��൱��HOLLOW_BRUSH��
'LTGRAY_BRUSH������ɫ����
'NULL_BRUSH���ջ��ʣ��൱��HOLLOW_BRUSH��
'WHITE_BRUSH����ɫ����
'BLACK_PEN����ɫ�ֱ�
'WHITE_PEN����ɫ�ֱ�
Public Const WHITE_BRUSH = 0    '��ɫ����
Public Const LTGRAY_BRUSH = 1   '����ɫ����
Public Const GRAY_BRUSH = 2     '��ɫ����
Public Const DKGRAY_BRUSH = 3   '����ɫ����
Public Const BLACK_BRUSH = 4    '��ɫ����
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH
Public Const WHITE_PEN = 6      '��ɫ�ֱ�
Public Const BLACK_PEN = 7      '��ɫ�ֱ�

'����һ������
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                            ByVal X3 As Long, ByVal Y3 As Long) As Long

'�������ͷŶ�����
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'�����ǹ��ܺ���
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Public Const ALTERNATE = 1
Public Const WINDING = 2

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetArcDirection Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Public Const TRANSPARENT = 1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
                                  lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CENTER = &H1

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, _
                                 ByVal nCount As Long) As Long

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, _
                                    ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, _
                                    ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Public Declare Function GetUpdateRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

'��ָ�����Դ���һ���߼�����
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
'��ȡ����ĸ߶�,��ȡ���ֵĿ�Ȳ�׼
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, _
                                              ByVal cbString As Long, lpSize As Size) As Long
'nNumber*nNumerator/nDenominator �Զ��������롣�޷�����ķ���-1
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Const OFFSET_LEFT = 10
Public Const OFFSET_TOP = 20
Public Const OFFSET_RIGHT = 10
Public Const OFFSET_BOTTOM = 20

Private Type Type_Printer
    intPage As Integer
    lngWidth As Long
    lngHeight As Long
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    intOrient As Integer
    intBin As Integer
End Type
Public gPrinter As Type_Printer


'***************************************************************
'����ͼ�滭��ر���
'***************************************************************

'����ͼҳ�������Ϣ
'----------------------------------
Public mintMaxPage As Integer '���ҳ��
Private mintPageCount As Integer
Private mstrTimeRange()      'ÿһҳ��ʱ�䷶Χ ��ʽ ��ʼʱ��;����ʱ��(���ݷ���ʱ��)
Private mArrPageTime()       'ÿһҳ�Ŀ�ʼʱ��(�������к�ʹ��)
'˵�������ݵ�ǰҳ��1�����Ի�ȡ��ҳ����ʱ�䷶Χ�ͱ�ҳ��ʼʱ�䣬2������ʱ�䷶Χ���˱�ҳ�����ݡ�3�����ݱ�ҳ��ʼʱ���������������
'------------------------------------
'����
Private Const gintBmpW    As Integer = 12
Private Const gintBmpH    As Integer = 12
Private Const gintRows    As Integer = 11

Private RGB_BLACK         As Long
Private RGB_RED           As Long
Private RGB_WRITE         As Long
Private RGB_BLUE          As Long
Private RGB_GRAY          As Long
Private RGB_FleetGRAY     As Long
Public sinTwipsPerPixelY As Single
Public sinTwipsPerPixelX As Single
Private msngTwips         As Single
Private mobjDraw          As Object '��ͼ�豸����
Private gstdset           As New StdFont
Private mblnPrint         As Boolean '�Ǵ�ӡԤ������չʾ
Private msnTimeH          As Single  '��������ʱ��ĸ߶�
Private mlngHwnd          As Long
Private mlngDC            As Long
Private mlngMemDC         As Long
Private mlngBitmap        As Long
Private mlngOldBitmap     As Long
Private mlngMemBitmap     As Long
Private mlngPen           As Long
Private mlngBrush         As Long
Private mlngOldPen        As Long
Private mlngOldBrush      As Long
Private mlngFont          As Long
Private mlngOldFont       As Long

Private Type DrawClient
    ƫ���� As POINTAPI
    �̶����� As RECT
    �̶ȵ�λ As Long
    �������� As RECT
    �е�λ As Long
    �е�λ As Long
    MaxX As Long
    ������ As Long
    ������� As RECT
End Type
Private T_DrawClient As DrawClient

Private Type type_Patient
    lng�ļ�ID As Long
    lng����ID As Long
    lng��ҳID As Long
    lng����ID As Long
    lng���� As Long
    lngҳ�� As Long
End Type
Private T_Info As type_Patient

'--�̶���Ŀ���
Private Type type_PartogramItem
    lng�������� As Long
    lng��¶�ߵ� As Long
    lng���� As Long
    lng���� As Long
End Type

Private T_Partogram As type_PartogramItem
'----------------------------------------------------------
'����ͼ��ر���
'----------------------------------------------------------
Private mrsItems As New ADODB.Recordset '������Ŀ��¼
Private mrsPartogram As New ADODB.Recordset '����������Ŀ
Private mrsDrawItems As New ADODB.Recordset '���������¼��
Private mrsSelItems As New ADODB.Recordset
Private gstrFields As String
Private gstrValues As String
Private mstrCatercorner As String           '�жԽ��߼���
Private mbln����ʱ��ϲ� As Boolean         '������ʱ��ϲ�
Private mblnDateAd As Boolean               '������д?
Private mstr����ʱ�� As String              '�����й��ɹ�����ʼʱ��
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private mlng��ʽID As Long                  '�ļ���ʽID
'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mobjSubFont As New StdFont  '���±�ǩ����
Private mobjTitleFont As New StdFont '��������
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageFoot As String      'ҳ��
Private mstrTitle As String         '��������
Private mstrSubHead As String       '���ϱ�ǩ
Private mstrSubEnd As String        '���±�ǩ
Private mstrOutSubHead As String    '��ȡ���ݺ���ϱ�ǩ
Private mstrOutSubEnd As String    '��ȡ���ݺ���±�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private mlngItems As Long '�����Ŀ��
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mTabForeColor As Long '����ı���ɫ
Private mTabGridColor As Long '���������ɫ
'������̼�¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String


Public Sub ShowPrintPartogram(ByVal objParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDtpID As Long, ByVal lngFileIndex As Long, ByVal lngFilePage As Long, Optional ByVal blnPrint As Boolean = True, Optional ByVal strPrintDevice As String = "")
'-----------------------------------------------------------------------------------------------------------
'��ɲ���ͼԤ������ӡ
'-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strParam As String, lngFileFormat As Long
    Dim objPrint As Object
    
    On Error GoTo errHand
    
    gstrSQL = "select ��ʽID from ���˻����ļ� where ID=[1] And ����ID=[2] And ��ҳID=[3] And nvl(Ӥ��,0)=0"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ʽID", lngFileID, lngPatiID, lngPageId)
    lngFileFormat = Val(NVL(rsTemp!��ʽID))
    
    '��ȡ��ӡ������Ϣ
    If Not PrintState(lngFileFormat, strPrintDevice) Then Exit Sub
    
    If blnPrint = True Then
        Set objPrint = Printer
    Else
        Set objPrint = frmPartogramPrint
    End If
    
    Load frmPartogramRead
    Call frmPartogramRead.InitRechBox(lngFileFormat)
    
    strParam = lngFileID & ";" & lngPatiID & ";" & lngPageId & ";" & lngDtpID & ";" & lngFileIndex & ";" & lngFilePage
    If blnPrint = True Then Printer.Print ""
    If Not PreViewOrPrintPartogram(strParam, objPrint, objParent, True) Then
        MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
        GoTo ErrEnd
    End If
    
    If blnPrint = True Then
        Printer.EndDoc
    Else
        '��ʾԤ������
        Call frmPartogramPrint.Preview(objParent, lngFileID, lngPatiID, lngPageId, lngDtpID, lngFileIndex, lngFilePage)
    End If
ErrEnd:
    Unload frmPartogramRead
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PreViewOrPrintPartogram(ByVal strParam As String, ObjDraw As Object, ByVal objParent As Object, Optional ByVal blnPrint As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------------
'��ɲ���ͼ���ݵ�չʾ(չʾ��Ԥ������ӡ)
'������strPram:�ļ�ID;����ID;��ҳID;����ID;����;ҳ��
'      objDraw:չ�ֲ���ͼ�Ķ��� (Pictrue?Printer)
'      blnPrint:True ��ʾ��ӡ��Ԥ�����ã�false����չʾ
'˵����blnPrint=trueʱ������=-1��ʾ��ӡ�����ļ���ҳ��û��;����>0ʱ��ҳ��=-1��ʾ��ӡ�˷��ļ���ҳ��>0�Ǳ�ʾ��ӡ���ļ�ĳһҳ
'      blnPrint=False ʱ��������ҳ��������С��0����ʾչʾĳ���ļ���ĳһҳ����
'---------------------------------------------------------------------------------------------------
    Dim arrParam
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset  '��ȡ�������ݣ������ط���Ҫʹ��
    Dim blnPrinter As Boolean
    Dim stdSet As StdFont
    Dim lngCurX As Long, lngCurY As Long
    Dim lngMaxRows As Long, lngFileIndexCount As Long
    Dim strPartogram As String '����������Ŀ��Ϣ����������������;��
    Dim strTmp As String
    
    'Ԥ����ӡ���
    Dim lngFileCount As Long, lngFileIndex As Long, lngPageIndex As Long, lngPicPageIndex As Long, lngԭʼҳ�� As Long, lngPrintPage As Long
    Dim i As Long, j As Long
    Dim dblSureW As Double, dblSureH As Double
    Dim blnReadData As Boolean '�Ƿ��ٴζ�ȡ����
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    Dim lngOffsetLeft As Long, lngOffsetTop As Long
    Dim intFine As Integer, intBold As Integer
    Dim lngHeight As Long
    '��ӡ�ȴ��������
    Dim sngCurOpt As Single, sngScale As Single, sngScaleFileIndex As Single, sngScaleFilePage As Single, lngCountOpt As Long
    Dim strInfo As String
    
    '�����С������ر���
    Dim lngObjHeight As Long, lngobjWidth As Long
    On Error GoTo errHand
    '��ʼ������
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then
        MsgBox "���鴫��Ĳ�����ʽ���Ƿ���ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    T_Info.lng�ļ�ID = Val(arrParam(0))
    T_Info.lng����ID = Val(arrParam(1))
    T_Info.lng��ҳID = Val(arrParam(2))
    T_Info.lng����ID = Val(arrParam(3))
    T_Info.lng���� = 1: T_Info.lngҳ�� = 1
    If UBound(arrParam) > 3 Then T_Info.lng���� = IIf(Val(arrParam(4)) = 0, 1, Val(arrParam(4)))
    If UBound(arrParam) > 4 Then T_Info.lngҳ�� = IIf(Val(arrParam(5)) = 0, 1, Val(arrParam(5)))
    
    lngPicPageIndex = 0
    Set mobjDraw = ObjDraw
    If blnPrint = True And Not ISobjPrinter Then
        Set mobjDraw = ObjDraw.picPage(lngPicPageIndex)
    Else
        Set mobjDraw = ObjDraw
    End If
    
    mblnPrint = blnPrint
    blnPrinter = ISobjPrinter
    'һ����ʼ����--------------------------------------------------------
    
    Screen.MousePointer = 11
    Call ShowFlash(blnPrint, "���ڳ�ʼ������,���Ժ�...", 0.1, objParent)
    
    '1����ʼ����������
    Call InitEnv
    If Not blnPrinter Then
        sinTwipsPerPixelX = Screen.TwipsPerPixelX
        sinTwipsPerPixelY = Screen.TwipsPerPixelY
        msngTwips = 1
        mobjDraw.Width = Printer.Width
        mobjDraw.Height = Printer.Height
        intBold = 2
        intFine = 1
    Else
        sinTwipsPerPixelX = Printer.TwipsPerPixelX
        sinTwipsPerPixelY = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        intBold = 6
        intFine = 2
    End If
    
    '��ȡ���̶̹���Ŀ��Ϣ
    strPartogram = ""
    mrsItems.Filter = 0
    mrsItems.Filter = "��Ŀ����='��������' And ������Ŀ=1"
    T_Partogram.lng�������� = mrsItems!��Ŀ���
    mrsItems.Filter = "��Ŀ����='��¶�ߵ�' And ������Ŀ=1"
    T_Partogram.lng��¶�ߵ� = mrsItems!��Ŀ���
    mrsItems.Filter = "��Ŀ����='����' And ������Ŀ=1"
    T_Partogram.lng���� = mrsItems!��Ŀ���
    mrsItems.Filter = "��Ŀ����='����' And ������Ŀ=1"
    T_Partogram.lng���� = mrsItems!��Ŀ���
    
    lngMaxRows = 0
    mstrSQL = "SELECT ��¼��,��¼��,��¼ɫ,���ֵ,��Сֵ,��λֵ,��λ FROM ���¼�¼��Ŀ WHERE ��Ŀ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "������Ŀ", T_Partogram.lng��������)
    If rsTemp.RecordCount = 0 Then
        strPartogram = "��������[LPF]��[LPF]255[LPF]10[LPF]0[LPF]1[LPF]CM"
        lngMaxRows = 11
    Else
        strPartogram = NVL(rsTemp!��¼��, "��������") & "[LPF]" & NVL(rsTemp!��¼��, "��") & "[LPF]" & NVL(rsTemp!��¼ɫ, 255) & "[LPF]" & _
            NVL(rsTemp!���ֵ, 10) & "[LPF]" & NVL(rsTemp!��Сֵ, "0") & "[LPF]" & NVL(rsTemp!��λֵ, 1) & "[LPF]" & NVL(rsTemp!��λ, "CM")
        lngMaxRows = Val(NVL(rsTemp!���ֵ, 10)) - Val(NVL(rsTemp!��Сֵ, "0"))
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "������Ŀ", T_Partogram.lng��¶�ߵ�)
    If rsTemp.RecordCount = 0 Then
        strPartogram = strPartogram & "[|LPF|]" & "��¶�ߵ�[LPF]��[LPF]10485760[LPF]5[LPF]-5[LPF]1[LPF]CM"
        lngMaxRows = 11
    Else
        strPartogram = strPartogram & "[|LPF|]" & NVL(rsTemp!��¼��, "��¶�ߵ�") & "[LPF]" & NVL(rsTemp!��¼��, "��") & "[LPF]" & NVL(rsTemp!��¼ɫ, 10485760) & "[LPF]" & _
            NVL(rsTemp!���ֵ, 5) & "[LPF]" & NVL(rsTemp!��Сֵ, "-5") & "[LPF]" & NVL(rsTemp!��λֵ, 1) & "[LPF]" & NVL(rsTemp!��λ, "CM")
        lngMaxRows = Val(NVL(rsTemp!���ֵ, 10)) - Val(NVL(rsTemp!��Сֵ, "0"))
    End If
    
    If lngMaxRows <= 0 Then lngMaxRows = gintRows
    
    If blnPrint = False Then
        T_DrawClient.ƫ����.X = 10 * msngTwips
        T_DrawClient.ƫ����.Y = 10 * msngTwips
    Else
        lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips) / sinTwipsPerPixelX
        lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips) / sinTwipsPerPixelY

        '�Ӵ�ӡ���ϱ߾����߾�
        lngLeft = (gPrinter.lngLeft * conRatemmToTwip) / sinTwipsPerPixelX + lngOffsetLeft
        lngRight = (gPrinter.lngRight * conRatemmToTwip) / sinTwipsPerPixelX
        lngTop = (gPrinter.lngTop * conRatemmToTwip) / sinTwipsPerPixelY + lngOffsetTop
        lngBottom = (gPrinter.lngBottom * conRatemmToTwip) / sinTwipsPerPixelY + lngOffsetTop

        T_DrawClient.ƫ����.X = lngLeft
        T_DrawClient.ƫ����.Y = lngTop
    End If
    T_DrawClient.�̶ȵ�λ = Format(52 * msngTwips, "0")
    T_DrawClient.�е�λ = Format(26 * msngTwips, "0")
    T_DrawClient.�е�λ = Format(26 * msngTwips, "0")
    T_DrawClient.MaxX = T_DrawClient.�е�λ * 24 + T_DrawClient.�̶ȵ�λ + T_DrawClient.ƫ����.X
    
    msnTimeH = (mobjDraw.TextHeight("1") / sinTwipsPerPixelY) + (2 * msngTwips)
    '---��ȡ�ļ�����
    If Not ReadStruDef Then GoTo ErrEnd
    '����ͼչʾʱ����PicTrue��С
    If blnPrint = False Then
        lngObjHeight = T_DrawClient.ƫ����.Y
        lngobjWidth = T_DrawClient.ƫ����.X
        '���
        If Val(zlDatabase.GetPara("��¶�ߵ���ʾλ��", glngSys, 1255, 0)) = 0 Then lngobjWidth = lngobjWidth + T_DrawClient.�̶ȵ�λ
        lngobjWidth = lngobjWidth + T_DrawClient.�̶ȵ�λ + T_DrawClient.�е�λ * 24
        '�߶�
        mobjDraw.FontSize = mobjTitleFont.Size
        T_Size.H = mobjDraw.TextHeight(mstrTitle) / sinTwipsPerPixelY
        lngObjHeight = lngObjHeight + (T_Size.H * 2) '����߶�
        mobjDraw.FontSize = mobjSubFont.Size
        lngObjHeight = lngObjHeight + (mobjDraw.TextHeight(mstrSubHead) / sinTwipsPerPixelY) + (2 * msngTwips) '�ϱ�ǩ
        lngObjHeight = lngObjHeight + lngMaxRows * T_DrawClient.�е�λ  '�������߲���
        If Val(zlDatabase.GetPara("����ͼ��ʾ����ʱ��", glngSys, 1255, 0)) = 1 Then lngObjHeight = lngObjHeight + msnTimeH
        lngObjHeight = lngObjHeight + msnTimeH
        mrsSelItems.Filter = "�̶�=0" '��񲿷�
        mrsSelItems.Sort = "��"
        Do While Not mrsSelItems.EOF
            lngObjHeight = lngObjHeight + Val(NVL(mrsSelItems!�߶�))
            mrsSelItems.MoveNext
        Loop
        mobjDraw.FontSize = mobjSubFont.Size
        lngObjHeight = lngObjHeight + (mobjDraw.TextHeight(mstrSubEnd) / sinTwipsPerPixelY) + (4 * msngTwips) '�±�ǩ
        
        mobjDraw.Width = (lngobjWidth + 12) * sinTwipsPerPixelX
        mobjDraw.Height = (lngObjHeight + 12) * sinTwipsPerPixelY
    End If
    Call ShowFlash(blnPrint, "���ڳ�ʼ������,���Ժ�...", 0.2, objParent)
    '---------------------------------------------------------------------------------------------------------------------
    '��ɲ���ͼչʾ��Ԥ���Լ���ӡ����
    '---------------------------------------------------------------------------------------------------------------------
    lngFileIndex = T_Info.lng����
    lngFileCount = lngFileIndex
    lngPicPageIndex = 0
    lngFileIndexCount = GetFileCount(T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID)
    If blnPrint = True Then
        If T_Info.lng���� < 1 Then '��ʾ��ӡ�����ļ�
            lngFileCount = lngFileIndexCount
            lngFileIndex = 1
            lngPageIndex = 1
            T_Info.lngҳ�� = -1
        End If
    Else
        If T_Info.lng���� < 1 Then T_Info.lng���� = 1
        lngFileIndex = T_Info.lng����
        lngFileCount = lngFileIndex
    End If
    lngԭʼҳ�� = T_Info.lngҳ��
    
    sngScaleFileIndex = Round(0.8 / (lngFileCount - lngFileIndex + 1), 2)
    For i = lngFileIndex To lngFileCount
        T_Info.lng���� = i
        blnReadData = True
        Set rsData = New ADODB.Recordset
        Call GetFileProperty
        If T_Info.lngҳ�� > mintPageCount Then T_Info.lngҳ�� = mintPageCount
        lngPageIndex = T_Info.lngҳ��
        If blnPrint = True Then
            If lngԭʼҳ�� < 1 Then '��ӡ��ǰ�ļ�
                lngPageIndex = 1
            Else
                mintPageCount = lngPageIndex
            End If
        Else
            If T_Info.lngҳ�� < 1 Then T_Info.lngҳ�� = 1
            lngPageIndex = T_Info.lngҳ��
            mintPageCount = lngPageIndex
        End If
        sngScale = (i - lngFileIndex) * sngScaleFileIndex
        sngScaleFilePage = Round(sngScaleFileIndex / (mintPageCount - lngPageIndex + 1), 2)
        For j = lngPageIndex To mintPageCount
            T_Info.lngҳ�� = j
            If lngPicPageIndex > 0 Then
                If blnPrint = True Then
                    If Not ISobjPrinter Then
                        Load ObjDraw.picPage(lngPicPageIndex)
                        Set mobjDraw = ObjDraw.picPage(lngPicPageIndex)
                        mobjDraw.Cls
                        mobjDraw.Width = Printer.Width
                        mobjDraw.Height = Printer.Height
                    Else
                        Printer.NewPage
                    End If
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Set mobjDraw = ObjDraw
                End If
            Else
                lngPicPageIndex = 1
            End If
            
            sngScale = sngScale + 0.2
            sngCurOpt = sngScaleFilePage * (i - lngFileIndex + 1)
            sngCurOpt = Round(sngCurOpt / 6, 2)
            
            strInfo = "���ڴ�ӡ��[" & i & "]���ļ��ĵ�[" & j & "]ҳ,���Ժ�..."
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 1) + sngScale, objParent)
            '������һ��û�����ݵĲ���ͼ-----------------------------------------
            mlngDC = mobjDraw.hDC
            
            'չʾԤ��������ջ�������
            If Not blnPrinter Then
                Call GetClientRect(mobjDraw.Hwnd, T_ClientRect)      'ȡ����Ļ����Ч����
                '������ɫˢ��
                mlngBrush = GetStockObject(WHITE_BRUSH)
                'ʹ�ø�ˢ����䱳��ɫ��ȫ�ף�
                mlngOldBrush = SelectObject(mlngDC, mlngBrush)
                Call FillRect(mlngDC, T_ClientRect, mlngBrush)
                '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
                Call SelectObject(mlngDC, mlngOldBrush)
                Call DeleteObject(mlngBrush)
            End If
            '����ҳü��Ϣ
            If blnPrint = True Then Call frmPartogramRead.PrintRTBData(mobjDraw, True, lngTop)
             '��ȡ���±���Ϣ
            If blnReadData = True Then Call GetMarkConnect
            '���������Ϣ
            Call SetFontIndirect(mobjTitleFont, mlngDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngDC, mlngFont)
            Call GetTextExtentPoint32(mlngDC, mstrTitle, Len(mstrTitle), T_Size)
            lngCurY = T_Size.H + T_DrawClient.ƫ����.Y
            Call GetTextRect(mobjDraw, 0, lngCurY, mstrTitle, mobjDraw.Width / sinTwipsPerPixelX, True, T_Size.H)
            Call DrawText(mlngDC, mstrTitle, -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            lngCurY = lngCurY + T_Size.H
            '����ϱ���Ϣ mstrOutSubHead
            lngCurX = T_DrawClient.ƫ����.X + T_DrawClient.�̶ȵ�λ
            mstrOutSubHead = Replace(mstrOutSubHead, "[ZLSOFTLPF]", "  ")
            Call DrawMarkConnect(mstrOutSubHead, lngCurX, lngCurY)
            lngCurY = lngCurY + (mobjDraw.TextHeight(mstrOutSubHead) / sinTwipsPerPixelY) + (2 * msngTwips)
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 2) + sngScale, objParent)
            '��������Ϣ
            T_DrawClient.�̶�����.Top = lngCurY
            T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����.X
            T_DrawClient.�̶�����.Bottom = T_DrawClient.�̶�����.Top + lngMaxRows * T_DrawClient.�е�λ
            T_DrawClient.�̶�����.Right = T_DrawClient.�̶ȵ�λ + T_DrawClient.�̶�����.Left
            T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
            T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
            T_DrawClient.��������.Right = T_DrawClient.��������.Left + 24 * T_DrawClient.�е�λ
            T_DrawClient.��������.Bottom = T_DrawClient.�̶�����.Bottom
            T_DrawClient.������ = lngMaxRows
            
            lngCurY = DrawPartogram(strPartogram) '�����̶̿��������������
            T_DrawClient.�������.Top = lngCurY
            T_DrawClient.�������.Left = T_DrawClient.�̶�����.Left
            T_DrawClient.�������.Right = T_DrawClient.�̶�����.Right
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 3) + sngScale, objParent)
            '��ɱ������Ļ�ͼ
            lngCurY = DrawPartogramTab
            T_DrawClient.�������.Bottom = lngCurY
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 4) + sngScale, objParent)
            '����±���Ϣ mstrOutSubend
            mstrOutSubEnd = Replace(mstrOutSubEnd, "[ZLSOFTLPF]", "  ")
            If Trim(Replace(mstrOutSubEnd, Chr(1), "")) <> "" Then 'ֻҪ���±�ǩ���ݲ�Ϊ�գ���ͷ���ƾ�Ϊ"��ע"�����߿���
                lngHeight = (mobjDraw.TextHeight(mstrOutSubEnd) / sinTwipsPerPixelY) + (4 * msngTwips)
                Call DrawLine(mlngDC, T_DrawClient.�̶�����.Left, lngCurY, T_DrawClient.�̶�����.Left, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.MaxX, lngCurY, T_DrawClient.MaxX, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.�̶�����.Left, lngCurY + lngHeight, T_DrawClient.MaxX, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                '�����ע
                strTmp = "��ע"
                strTmp = CheckConnect(strTmp, T_DrawClient.�̶ȵ�λ, lngHeight)
                lngHeight = (lngHeight - mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY) / 2
                Call GetTextRect(mobjDraw, T_DrawClient.�̶�����.Left, lngCurY + lngHeight, strTmp, T_DrawClient.�̶ȵ�λ, False)
                Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
            End If
            
            lngCurX = T_DrawClient.��������.Left
            lngCurY = lngCurY + (2 * msngTwips)
            Call DrawMarkConnect(mstrOutSubEnd, lngCurX, lngCurY)
            lngCurY = lngCurY + (mobjDraw.TextHeight(mstrOutSubEnd) / sinTwipsPerPixelY) + (2 * msngTwips)
            '������ȡ�������ߺͱ�����ݣ���������ݵ�չ��----------------------------------
            If blnReadData = True Then
                Call SQLCombination
                Call SQLDIY(mstrSQL)
                Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ϣ", T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID, 0, T_Info.lng����)
            End If
            blnReadData = False
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 5) + sngScale, objParent)
            '����������ݻ滭
            Call DrawPartogramCurData(rsData)
            '��ɱ�����ݵĻ滭
            Call DrawPartogramTabData(rsData)
            
            'ҳ��ͼ�����
            If blnPrint = True Then
                Call frmPartogramRead.PrintRTBData(mobjDraw, False, lngBottom, "�� " & j & " ҳ" & IIf(lngFileIndexCount > 1, "(" & i & ")", ""), mstrPageFoot)
            End If
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 6) + sngScale, objParent)
            
            If Not blnPrinter Then mobjDraw.Refresh
            
            If blnPrint = True And Not ISobjPrinter Then
                '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
                dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
                dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
                On Error Resume Next
                Call DrawRect(mlngDC, (mobjDraw.Width * dblSureW) / sinTwipsPerPixelX, (mobjDraw.Height * (1 - dblSureH)) / sinTwipsPerPixelY, _
                (mobjDraw.Width * (1 - dblSureW)) / sinTwipsPerPixelX, mobjDraw.Height * dblSureH / sinTwipsPerPixelY, PS_DOT, 1, RGB_FleetGRAY)
            End If
        Next j
    Next i
    PreViewOrPrintPartogram = True
ErrEnd:
    Screen.MousePointer = 0
    Call ShowFlash(blnPrint, "", 1, objParent)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call ShowFlash(blnPrint, "", 1, objParent)
    Call SaveErrLog
End Function

Private Function DrawMarkConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long) As Long
'---------------------------------------------------
'����:������±�ǩ�����
'---------------------------------------------------
    Dim strTmp As String, strFomart As String
    Dim intIndex As Integer, i As Integer, j As Integer
    Dim ArrMark, ArrCode
    Dim lngCurX As Long, lngCurY As Long, lngY1 As Long
    Dim lngHeight As Long
    Dim intSize As Single, ��¼ԭʼ������Ϣ
    
    If Trim(strText) = "" Then DrawMarkConnect = lngY: Exit Function
    strTmp = strText
    
    '���ܸ�ʽ�̶������±�ķ�ʽ���
    Do While True
        If strTmp Like "����*+*" Or strTmp Like "*��+*" Then
            intIndex = InStr(1, strTmp, "+")
            i = intIndex + 1
            For i = intIndex + 1 To Len(strTmp)
                If Not IsNumeric(Mid(strTmp, i, 1)) Then Exit For
            Next i
            
            strFomart = strFomart & Mid(strTmp, 1, intIndex - 1) & "['LPF']" & Mid(strTmp, intIndex, i - intIndex) & "['LPF']"
            strTmp = Mid(strTmp, i)
        Else
            strFomart = strFomart & strTmp
            Exit Do
        End If
    Loop
    
    '�����������ɫ
    intSize = mobjSubFont.Size
    Call SetTextColor(mlngDC, mTabForeColor)
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    lngHeight = mobjDraw.TextHeight("��") / sinTwipsPerPixelY
    ArrMark = Split(strFomart, vbCrLf)
    For i = 0 To UBound(ArrMark)
        lngCurX = lngX
        lngCurY = lngY + (i * lngHeight)
        lngY1 = lngCurY
        strFomart = CStr(ArrMark(i))
        ArrCode = Split(strFomart, "['LPF']")
        For j = 0 To UBound(ArrCode)
            strTmp = CStr(ArrCode(j))
            lngCurY = lngY1
            If j <> 0 And j < UBound(ArrCode) Then '��С���
                mobjSubFont.Size = 7
                lngCurY = lngCurY - (2 * msngTwips)
            Else '�������
                mobjSubFont.Size = intSize
            End If
            Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
            Call GetTextRect(mobjDraw, lngCurX, lngCurY, strTmp, 0, False)
            Call DrawText(mlngDC, strTmp, -1, T_LableRect, 0)
            lngCurX = lngCurX + (mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX)
        Next j
    Next i
    
    '���ԭ������Ϣ
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    DrawMarkConnect = lngY1
End Function

Private Sub DrawPartogramCurData(ByVal rsTemp As ADODB.Recordset)
'----------------------------------------------------------------------
'����:��������������ݵ�չʾ
'----------------------------------------------------------------------
    Dim rsCurData As New ADODB.Recordset
    Dim arrItemNo
    Dim lngOrder As Long, strOrder As String
    Dim sngX As Single, sngY As Single
    Dim sngOutX As Single, sngOutY As Single
    Dim i As Integer, j As Integer
    Dim strBeginDate As String, strFiled As String, strValue As String, strTime As String, strContent As String
    
    '��ŷ�����������Ϣ
    Dim rsCurInfo As New ADODB.Recordset
    Dim rsCopyCurInfo As New ADODB.Recordset 'rsCurInfo�ĸ���
    Dim strFields As String, strValues As String, strFiled1 As String
    
    '��ͼ����
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lng��Ŀ��� As Long, sinԭX As Single, sinԭY As Single
    Dim lngRGB As Long, strICon As String
    Dim blnLine As Boolean '��������ߣ�1������֮������ͨ��3CM��2������3Cm�ĵ㣩
    Dim sin3CmY As Single, sin3CmX As Single, intState As Integer
    Dim sin10CmY As Single, sin10CmX As Single
    Dim lngType As Long  '������ʽ����ʵ�߻�������
    '������Ϣ
    Dim bln��ʾ������ As Boolean, int�������� As Integer, int�쳣���� As Integer
    Dim int�������߱�־S As Integer, int��¶���߱�־S As Integer '˳��
    Dim int�������߱�־Y As Integer, int��¶���߱�־Y As Integer '�쳣��
    Dim int�������߱�־ As Integer, int��¶���߱�־ As Integer
    
    Dim int��־���� As Integer, int��־λ�� As Integer
    Dim int������� As Integer, blnZeroLine As Boolean  '��һ���ߵ���0������:0-������,1-����,2-ʵ��
    Dim sinZeroX As Single, sinZeroY As Single
    Dim blnAbnormal As Boolean '�쳣��
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '�������ݹ��˵��������
    strBeginDate = Format(DateAdd("d", IIf((T_Info.lngҳ�� - 1) < 0, 0, T_Info.lngҳ�� - 1), mstr����ʱ��), "YYYY-MM-DD HH:mm:ss")
    rsTemp.Filter = "����ʱ��>='" & strBeginDate & "' And ����ʱ��<='" & Format(DateAdd("d", 1, CDate(strBeginDate)), "YYYY-MM-DD HH:mm:ss") & "'"
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    '-----��ȡ������Ϣ
     '�����������߱�־
    strTmp = zlDatabase.GetPara("�����������߱�־", glngSys, 1255, "1;1", , True)
    int�������߱�־S = Val(Split(strTmp, ";")(0))
    int��¶���߱�־S = Val(Split(strTmp, ";")(1))
    strTmp = int�������߱�־S & ";" & int��¶���߱�־S
    strTmp = zlDatabase.GetPara("�����������߱�־(��)", glngSys, 1255, strTmp, , True)
    int�������߱�־Y = Val(Split(strTmp, ";")(0))
    int��¶���߱�־Y = Val(Split(strTmp, ";")(1))
    
    '����������ʩ��־
    strTmp = zlDatabase.GetPara("����������ʩ��־", glngSys, 1255, "1;1", , True)
    int��־���� = Val(Split(strTmp, ";")(0))
    int��־λ�� = Val(Split(strTmp, ";")(1))
    
    '���̾����߱�־
    strTmp = zlDatabase.GetPara("���̾����쳣�߱�־", glngSys, 1255, "1;1", , True)
    int�������� = Val(Split(strTmp, ";")(0))
    int�쳣���� = Val(Split(strTmp, ";")(1))
    bln��ʾ������ = (Val(zlDatabase.GetPara("����ͼ��ʾ������", glngSys, 1255, "1", , True)) = 1)
    int������� = Val(zlDatabase.GetPara("�������ߵ���0������", glngSys, 1255, "0", , True))
    If int������� < 0 Or int������� > 2 Then int������� = 0
    
    '�������߼�¼��
    gstrFields = "��Ŀ���," & adDouble & ",18|����," & adLongVarChar & ",1000|ʱ��," & adLongVarChar & ",20|X����," & adDouble & ",5|Y����," & adDouble & ",5"
    Call Record_Init(rsCurData, gstrFields)
    gstrFields = "��Ŀ���|����|ʱ��|X����|Y����"
    
    strFields = "��Ŀ���," & adDouble & ",18|��ֵ," & adLongVarChar & ",100|����," & adLongVarChar & ",1000|ʱ��," & adLongVarChar & ",20|X����," & adDouble & ",5|Y����," & adDouble & ",5|" & _
        "��ӡX����," & adDouble & ",5|��ӡY����," & adDouble & ",5|ģʽ," & adInteger & ",1|���," & adDouble & ",18|�߶�," & adDouble & ",18"
    Call Record_Init(rsCurInfo, strFields)
    Call Record_Init(rsCopyCurInfo, strFields)
    strFields = "��Ŀ���|��ֵ|����|ʱ��|X����"
     
    '---����ȡ����������¶�½���������Ϣ
    strOrder = T_Partogram.lng�������� & ";" & T_Partogram.lng��¶�ߵ� & ";" & T_Partogram.lng����
    arrItemNo = Split(strOrder, ";")
    For i = 0 To UBound(arrItemNo)
        mrsSelItems.Filter = 0
        mrsSelItems.Filter = "��=" & Val(arrItemNo(i)) & " And �̶�=1"
        If mrsSelItems.RecordCount > 0 Then
            lngOrder = Val(mrsSelItems!�������)
            strFiled = "C" & Format(lngOrder, "00")
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                strValue = Trim(NVL(rsTemp.Fields(strFiled).Value))
                If strValue <> "" Then
                    If Val(arrItemNo(i)) <> T_Partogram.lng���� Then
                        '����X��Y����
                        strTime = Format(rsTemp!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                        sngX = GetXCoordinate(strTime, strBeginDate)
                        sngY = GetYCoordinate(Val(arrItemNo(i)), strValue)
                        '��Ӽ�¼��
                        If (sngX <= T_DrawClient.MaxX And sngX >= T_DrawClient.��������.Left) And (sngY >= T_DrawClient.��������.Top And sngY <= T_DrawClient.��������.Bottom) Then
                            gstrValues = Val(arrItemNo(i)) & "|" & strValue & "|" & strTime & "|" & sngX & "|" & sngY
                            Call Record_Add(rsCurData, gstrFields, gstrValues)
                        End If
                    Else '����������ݺͱ��
                        '��ȡ������ֶ���
                        If Mid(strValue, 1, 1) = "��" Then
                            strContent = ""
                            Select Case int��־����
                                Case 1
                                    strContent = "����"
                                Case 2
                                    mrsSelItems.Filter = 0
                                    mrsSelItems.Filter = "��=" & T_Partogram.lng���� & " And �̶�=1"
                                    If mrsSelItems.RecordCount > 0 Then
                                        strFiled1 = "C" & Format(Val(mrsSelItems!�������), "00")
                                        strContent = Trim(NVL(rsTemp.Fields(strFiled1).Value))
                                    End If
                            End Select
                            strTime = Format(rsTemp!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                            sngX = GetXCoordinate(strTime, strBeginDate)
                            If (sngX <= T_DrawClient.MaxX And sngX >= T_DrawClient.��������.Left) Then
                                strValues = Val(arrItemNo(i)) & "|" & strValue & "|" & strContent & "|" & strTime & "|" & sngX
                                Call Record_Add(rsCurInfo, strFields, strValues)
                            End If
                        End If
                    End If
                End If
                rsTemp.MoveNext
            Loop
        End If
    Next i
    
    '��ʼ�������߲���
    blnPrinter = ISobjPrinter
    If blnPrinter = True Then
        intBold = 4
        intFine = 4
    Else
        intBold = 2
        intFine = 1
    End If
    
    '��ȡ��������3Cm������
    sin3CmY = GetYCoordinate(T_Partogram.lng��������, 3)
    
    rsCurData.Filter = ""
    rsCurData.Sort = "��Ŀ���,ʱ��"
    '----��ɵ����֮������ߺͷ��ŵ����
    blnLine = False
    blnZeroLine = False
    lng��Ŀ��� = -999
    sin3CmX = 0: sinԭX = 0: sinԭY = 0
    With rsCurData
        Do While Not .EOF
            blnAbnormal = False
            If NVL(!��Ŀ���) <> lng��Ŀ��� Then
                blnZeroLine = True
                sinԭX = 0
                sinԭY = 0
                lng��Ŀ��� = NVL(!��Ŀ���)
                lngRGB = Val(GetDrawItemValue(lng��Ŀ���, "��ɫ"))
                strICon = GetDrawItemValue(lng��Ŀ���, "��¼��")
                intState = GetDrawItemValue(lng��Ŀ���, "��ʾģʽ")
            End If
            
            If sinԭX <> 0 Then
                '�����Ժ�����쳣����¶�ߵʹ���һ�㵽�����㻭ֱ������
                '���һ��������ͬʱ¼������¶�ߵ͵����
                If lng��Ŀ��� = T_Partogram.lng��¶�ߵ� Then
                    rsCurInfo.Filter = "X����=" & !x����
                    If rsCurInfo.RecordCount > 0 Then
                        blnAbnormal = (rsCurInfo!��ֵ = "��(��)")
                    End If
                End If
                If blnAbnormal = True And int��¶���߱�־Y = 3 Then
                    Call DrawLine(mlngDC, sinԭX, sinԭY, !x����, sinԭY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                    Call DrawLine(mlngDC, !x����, sinԭY, !x����, !Y����, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                Else
                    Call DrawLine(mlngDC, sinԭX, sinԭY, !x����, !Y����, PS_SOLID, intFine, lngRGB)
                End If
                '�ж��Ƿ�տڿ���3Cm������֮������߻���3Cm
                If blnLine = False And lng��Ŀ��� = T_Partogram.lng�������� Then
                    If intState = 0 Then
                        If sin3CmY < sinԭY And sin3CmY > !Y���� Then blnLine = True
                    Else
                        If sin3CmY > sinԭY And sin3CmY < !Y���� Then blnLine = True
                    End If
                    '���㾭��3Cm���ĵ�
                    If blnLine = True Then
                        sin3CmX = ((!x���� - sinԭX) * (sin3CmY - sinԭY) / (!Y���� - sinԭY)) + sinԭX
                    End If
                End If
            End If
            sinԭX = NVL(!x����, 0)
            sinԭY = NVL(!Y����, 0)
            If blnLine = False And lng��Ŀ��� = T_Partogram.lng�������� Then
                If sin3CmY = sinԭY Then blnLine = True: sin3CmX = sinԭX
            End If
            '���ͼ��
            Set gstdset = New StdFont
            gstdset.Name = "����"
            gstdset.Size = 9
            gstdset.Underline = False
            gstdset.Italic = False
            Call SetFontIndirect(gstdset, mlngDC, mobjDraw)
            Call SetTextColor(mlngDC, lngRGB)
            Call GetTextRect(mobjDraw, sinԭX - (mobjDraw.TextWidth(strICon) / sinTwipsPerPixelX / 2), sinԭY, Trim(strICon))
            Call DrawText(mlngDC, Trim(strICon), -1, T_LableRect, DT_CENTER)
            If int������� > 0 And blnZeroLine = True And sinԭX > T_DrawClient.��������.Left Then
                sinZeroX = T_DrawClient.��������.Left
                mrsDrawItems.Filter = "��Ŀ���=" & lng��Ŀ���
                If mrsDrawItems.RecordCount > 0 Then
                    sinZeroY = GetYCoordinate(lng��Ŀ���, mrsDrawItems!��Сֵ)
                Else
                    sinZeroY = 0
                End If
                If int������� = 1 Then '����
                    Call DrawLine(mlngDC, sinZeroX, sinZeroY, sinԭX, sinԭY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                Else 'ʵ��
                    Call DrawLine(mlngDC, sinZeroX, sinZeroY, sinԭX, sinԭY, PS_SOLID, intFine, lngRGB)
                End If
                blnZeroLine = False
            End If
        .MoveNext
        Loop
    End With
    
    '-----�������ߺ��쳣��
    If blnLine = True And bln��ʾ������ Then
        If sin3CmX > 0 And sin3CmX < T_DrawClient.MaxX Then
            sin10CmX = sin3CmX + (T_DrawClient.�е�λ * 4)
            sin10CmY = GetYCoordinate(T_Partogram.lng��������, Val(GetDrawItemValue(T_Partogram.lng��������, "���ֵ")))
            If sin10CmX > T_DrawClient.MaxX Then GoTo ErrNext
            '��������
            Select Case int��������
                Case 0
                    lngType = IIf(blnPrinter, PS_DASH, PS_DOT)
                Case Else
                    lngType = PS_SOLID
            End Select
            Call DrawLine(mlngDC, sin3CmX, sin3CmY, sin10CmX, sin10CmY, lngType, intFine, RGB_RED)
            '���쳣��
            If sin10CmX + (T_DrawClient.�е�λ * 4) > T_DrawClient.MaxX Then GoTo ErrNext
             Select Case int�쳣����
                Case 0
                    lngType = IIf(blnPrinter, PS_DASH, PS_DOT)
                Case Else
                    lngType = PS_SOLID
            End Select
            Call DrawLine(mlngDC, sin10CmX, sin3CmY, sin10CmX + (T_DrawClient.�е�λ * 4), sin10CmY, lngType, intFine, RGB_RED)
        End If
    End If
ErrNext:
    '----���в��˷��䴦��
    Dim sinUpY As Single, sinUpX As Single '�����ı�����ʹ�ã������ط���Ҫʹ��
    strTmp = ""
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "ʱ��"
    With rsCurInfo  '"��Ŀ���|����|ʱ��|X����"
        Do While Not .EOF
            sngX = !x����
            blnAbnormal = (!��ֵ = "��(��)")
            If blnAbnormal = True Then
                int�������߱�־ = int�������߱�־Y
                int��¶���߱�־ = int��¶���߱�־Y
            Else
                int�������߱�־ = int�������߱�־S
                int��¶���߱�־ = int��¶���߱�־S
            End If
            
            If int�������߱�־ > 0 Then
                rsCurData.Filter = ""
                rsCurData.Filter = "��Ŀ���=" & T_Partogram.lng�������� & " And X����<=" & sngX
                rsCurData.Sort = "ʱ�� DESC"
                If rsCurData.RecordCount > 0 Then
                    sinԭX = rsCurData!x����
                    sinԭY = rsCurData!Y����
                    lngRGB = Val(GetDrawItemValue(T_Partogram.lng��������, "��ɫ"))
                    intState = Val(GetDrawItemValue(T_Partogram.lng��������, "��ʾģʽ"))
                    '��ʾ�ڹ�������ʱ��������Y���������Ĺ���Y������ͬ
                    Call DrawLine(mlngDC, sinԭX, sinԭY, sngX, sinԭY, PS_SOLID, intFine, lngRGB)
                    
                    If intState = 0 Then
                        sin3CmY = sinԭY + T_DrawClient.�е�λ
                    Else
                        sin3CmY = sinԭY - T_DrawClient.�е�λ
                    End If
                    If int�������߱�־ = 1 Then '��ʾ���߼�ͷ
                        Call DrawLine(mlngDC, sngX, sinԭY, sngX, sin3CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB, True)
                    Else '��ʾʵ�߼�ͷ
                        Call DrawLine(mlngDC, sngX, sinԭY, sngX, sin3CmY, PS_SOLID, intFine, lngRGB, True)
                    End If
                    If int��־λ�� = 0 Then
                       !Y���� = sinԭY
                       !��ӡY���� = sinԭY
                       .Update
                    End If
                End If
            End If
            
            If int��¶���߱�־ > 0 Then
                rsCurData.Filter = ""
                rsCurData.Filter = "��Ŀ���=" & T_Partogram.lng��¶�ߵ� & " And X����<=" & sngX
                rsCurData.Sort = "ʱ�� DESC"
                If rsCurData.RecordCount > 0 Then
                    sinԭX = rsCurData!x����
                    sinԭY = rsCurData!Y����
                    sin10CmY = GetYCoordinate(T_Partogram.lng��¶�ߵ�, Val(GetDrawItemValue(T_Partogram.lng��¶�ߵ�, "���ֵ")))
                    lngRGB = Val(GetDrawItemValue(T_Partogram.lng��¶�ߵ�, "��ɫ"))
                    intState = Val(GetDrawItemValue(T_Partogram.lng��¶�ߵ�, "��ʾģʽ"))
                    '��ʾ�ڹ�������ʱ��������Y���������Ĺ���Y������ͬ
                    If sngX = sinԭX Then
                        sin10CmY = sinԭY
                    Else
                        If blnAbnormal = True And int��¶���߱�־ = 3 Then
                            Call DrawLine(mlngDC, sinԭX, sinԭY, sngX, sinԭY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                            Call DrawLine(mlngDC, sngX, sinԭY, sngX, sin10CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                        Else
                            Call DrawLine(mlngDC, sinԭX, sinԭY, sngX, sin10CmY, PS_SOLID, intFine, lngRGB)
                        End If
                    End If
                    If intState = 0 Then
                        sin3CmY = sin10CmY - (T_DrawClient.�е�λ / 2)
                        If sin10CmY = T_DrawClient.�̶�����.Top Then
                            sin3CmY = sin10CmY - (T_DrawClient.�е�λ / 2)
                        ElseIf sin10CmY - T_DrawClient.�е�λ <= T_DrawClient.�̶�����.Top Then
                            sin3CmY = sin10CmY - T_DrawClient.�е�λ
                        ElseIf sin10CmY - T_DrawClient.�̶�����.Top >= (T_DrawClient.�е�λ / 2) Then
                            sin3CmY = T_DrawClient.�̶�����.Top
                        End If
                    Else
                        sin3CmY = sin10CmY + (T_DrawClient.�е�λ / 2)
                        If sin10CmY = T_DrawClient.�̶�����.Bottom Then
                            sin3CmY = sin10CmY + (T_DrawClient.�е�λ / 2)
                        ElseIf sin10CmY + T_DrawClient.�е�λ <= T_DrawClient.�̶�����.Bottom Then
                            sin3CmY = sin10CmY + T_DrawClient.�е�λ
                        ElseIf T_DrawClient.�̶�����.Bottom - sin10CmY >= (T_DrawClient.�е�λ / 2) Then
                            sin3CmY = T_DrawClient.�̶�����.Bottom
                        End If
                    End If
                    If int��¶���߱�־ = 1 Then '��ʾ���߼�ͷ
                        Call DrawLine(mlngDC, sngX, sin10CmY, sngX, sin3CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB, True)
                    ElseIf int��¶���߱�־ = 2 Then '��ʾʵ�߼�ͷ
                        Call DrawLine(mlngDC, sngX, sin10CmY, sngX, sin3CmY, PS_SOLID, intFine, lngRGB, True)
                    End If
                    '�ı���־λ��
                    If int��־λ�� = 1 Then
                       !Y���� = sinԭY
                       !��ӡY���� = sinԭY
                       .Update
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '----��ʼ��������ı�������Ϣ
    If int��־λ�� = 0 Then
        lngRGB = Val(GetDrawItemValue(T_Partogram.lng��������, "��ɫ"))
        intState = Val(GetDrawItemValue(T_Partogram.lng��������, "��ʾģʽ"))
    Else
        lngRGB = Val(GetDrawItemValue(T_Partogram.lng��¶�ߵ�, "��ɫ"))
        intState = Val(GetDrawItemValue(T_Partogram.lng��¶�ߵ�, "��ʾģʽ"))
    End If
    Dim intģʽ As Integer, blnInit As Boolean, blnEnd As Boolean
    Dim lngMax As Single, lngMaxY As Single, sng�в� As Single
    Dim lngCurWidth As Long, lngCurHeight As Long
    
    sng�в� = 4 * msngTwips
    
    '����rsCurInfo
    rsCurInfo.Filter = ""
    Do While Not rsCurInfo.EOF
        rsCopyCurInfo.AddNew
        For j = 0 To rsCurInfo.Fields.Count - 1
            rsCopyCurInfo.Fields(j).Value = IIf(NVL(rsCurInfo.Fields(j).Value) = "", Null, rsCurInfo.Fields(j).Value)
        Next j
        rsCopyCurInfo.Update
    rsCurInfo.MoveNext
    Loop
    
    T_Size.W = mobjDraw.TextWidth("��") / sinTwipsPerPixelX
    T_Size.H = mobjDraw.TextHeight("��") / sinTwipsPerPixelY
    lngMax = T_DrawClient.MaxX
    lngMaxY = T_DrawClient.��������.Bottom
    '�����������ݵ�X��������ģʽ(����������Ǻ������)
    rsCopyCurInfo.Filter = ""
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "X���� DESC"
    With rsCurInfo
        Do While Not .EOF
            If Val(NVL(!Y����, 0)) >= T_DrawClient.�̶�����.Top And Val(NVL(!Y����, 0)) <= T_DrawClient.�̶�����.Bottom Then
                strTmp = NVL(!����)
                If .AbsolutePosition = 1 Then
                    sngX = Val(NVL(!x����))
                Else
                    If sngX < Format(Val(NVL(!x����)) + T_Size.W + sng�в�, "0.0") Then
                        sngX = Format(sngOutX - T_Size.W - sng�в�, "0.0")
                    Else
                        sngX = Val(NVL(!x����))
                    End If
                End If
                sngY = Val(NVL(!Y����))
                If CInt(lngMax - sngX) < CInt(T_DrawClient.�е�λ * 2) Then '��������ı���Ϣ
                    intģʽ = 1
                    If CInt(sngX + T_Size.W + sng�в�) > CInt(lngMax) Then
                        rsCopyCurInfo.Filter = "X����<" & sngX
                        rsCopyCurInfo.Sort = "X���� DESC"
                        If rsCopyCurInfo.RecordCount > 0 Then
                            If CInt(sngX - Val(NVL(rsCopyCurInfo!x����))) < CInt(T_Size.W + sng�в�) Then
                                sngX = Format(lngMax - T_Size.W - sng�в�, "0.0")
                            Else
                                sngX = Format(sngX - T_Size.W - sng�в�, "0.0")
                            End If
                        Else
                            sngX = Format(sngX - T_Size.W - sng�в�, "0.0")
                        End If
                    End If
                   
                    If intState = 1 Then
                         If CInt(lngMaxY - sngY) < CInt(mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX) Then
                            sngY = Format(lngMaxY - (mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX) - T_Size.H, "0.0")
                         End If
                    End If
                    If sngY < T_DrawClient.�̶�����.Top Then sngY = T_DrawClient.�̶�����.Top
                    
                    lngCurWidth = lngMax - sngX
                    lngCurHeight = T_DrawClient.��������.Bottom - sngY
                    lngMax = sngX
                    sngOutX = sngX
                Else '��������ı���Ϣ
                    intģʽ = 0
                    lngCurWidth = lngMax - sngX
                    strTmp = CheckConnect(strTmp, lngCurWidth, T_DrawClient.��������.Bottom - sngY)
                    lngCurHeight = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
                    If CInt(lngCurHeight + sngY + sng�в�) > CInt(T_DrawClient.��������.Bottom) Then
                        sngY = Format(T_DrawClient.��������.Bottom - lngCurHeight - sng�в�, "0.0")
                    End If
                    rsCopyCurInfo.Filter = "X����<" & sngX
                    rsCopyCurInfo.Sort = "X���� DESC"
                    If rsCopyCurInfo.RecordCount > 0 Then
                        If CInt(sngX - Val(NVL(rsCopyCurInfo!x����))) < CInt(T_Size.W + sng�в�) Then
                            sngOutX = Format(sngX + T_Size.W + sng�в�, "0.0")
                            sngX = Format(Val(NVL(rsCopyCurInfo!x����)) + T_Size.W + sng�в� - (1 * msngTwips), "0.0")
                        Else
                            lngMax = sngX
                            sngOutX = sngX
                        End If
                    Else
                        lngMax = sngX
                        sngOutX = sngX
                    End If
                End If
                !�߶� = lngCurHeight
                !��� = lngCurWidth
                !��ӡX���� = Int(sngX)
                !��ӡY���� = Int(sngY)
                !ģʽ = intģʽ
                .Update
            End If
        .MoveNext
        Loop
    End With
    
    '���¸���rsCurInfo��¼��
    rsCopyCurInfo.Filter = ""
    Do While Not rsCopyCurInfo.EOF
        rsCopyCurInfo.Delete
        rsCopyCurInfo.Update
    rsCopyCurInfo.MoveNext
    Loop
    
    rsCurInfo.Filter = ""
    Do While Not rsCurInfo.EOF
        rsCopyCurInfo.AddNew
        For j = 0 To rsCurInfo.Fields.Count - 1
            rsCopyCurInfo.Fields(j).Value = IIf(NVL(rsCurInfo.Fields(j).Value) = "", Null, rsCurInfo.Fields(j).Value)
        Next j
        rsCopyCurInfo.Update
    rsCurInfo.MoveNext
    Loop
    
'    rsCurInfo.Filter = ""
'    Call OutputRsData(rsCurInfo, True)
    
    blnInit = False
    blnEnd = False
    '�����������������ݴ�ӡY����
    rsCopyCurInfo.Filter = "ģʽ=0"
    rsCopyCurInfo.Sort = "��ӡX����"
    i = 1
    With rsCopyCurInfo
        Do While Not .EOF
            If blnInit = False Then sngX = Val(NVL(!��ӡX����)): lngCurHeight = 0: blnInit = True
            If Val(NVL(!��ӡX����)) <> sngX Then
ErrEnd:
                rsCurInfo.Filter = "��ӡX����=" & sngX
                rsCurInfo.Sort = "��ӡY����"
                Do While Not rsCurInfo.EOF
                    sngY = Val(NVL(rsCurInfo!��ӡY����)) + IIf(rsCurInfo.AbsolutePosition > 1, lngCurHeight + sng�в�, 0)
                    sngOutY = lngMaxY - lngCurHeight - sng�в�
                    If sngOutY < T_DrawClient.��������.Top + sng�в� Then sngOutY = T_DrawClient.��������.Top + sng�в�
                    If sngY > sngOutY Then sngY = sngOutY
                    If Val(NVL(rsCurInfo!�߶�)) > lngMaxY - sngY Then rsCurInfo!�߶� = Format(lngMaxY - sngY, "0.00")
                    rsCurInfo!��ӡY���� = sngY
                    rsCurInfo.Update
                    
                    lngCurHeight = lngCurHeight - Val(NVL(rsCurInfo!�߶�))
                rsCurInfo.MoveNext
                Loop
            
                sngX = Val(NVL(!��ӡX����)): lngCurHeight = Val(NVL(!�߶�))
                If blnEnd = True Then GoTo ErrOutPut
            Else
                lngCurHeight = lngCurHeight + Val(NVL(!�߶�))
            End If
            If i = rsCopyCurInfo.RecordCount Then blnEnd = True: GoTo ErrEnd
            i = i + 1
        .MoveNext
        Loop
    End With
ErrOutPut:
    '��ʽ����ı���Ϣ
    Call SetTextColor(mlngDC, lngRGB)
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "X����"
    With rsCurInfo
        Do While Not .EOF
            If Val(NVL(!Y����, 0)) >= T_DrawClient.�̶�����.Top And Val(NVL(!Y����, 0)) <= T_DrawClient.�̶�����.Bottom Then
                strTmp = NVL(!����)
                If Val(NVL(!ģʽ)) = 0 Then
                    Call OutBigHConnect(strTmp, !��ӡX����, !��ӡY����, !�߶�, !���, False, False)
                Else
                    Call OutBigVConnect(strTmp, !��ӡX����, !��ӡY����, !�߶�, !���, False, False, lngRGB)
                End If
            End If
        .MoveNext
        Loop
    End With

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawPartogramTabData(ByVal rsTemp As ADODB.Recordset)
'----------------------------------------------------------------------
'����:��ɱ���������ݵ�չʾ
'----------------------------------------------------------------------
    Dim strBegin As String, strEnd As String, strTmp As String, strPageTime As String
    Dim rsTabData As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long, blnInit As Boolean
    Dim intFields As Integer, lngOrder As Long
    '---�����Ϣ
    Dim arrItemOrder, arrItemName, arrItemHeight
    Dim lngCurX As Long, lngCurY As Long, lngHeight As Long
    Dim strName As String
    Dim strLeft As String, strRight As String
    Dim intBold As Integer, intFine As Integer
    On Error GoTo errHand
    
    If ISobjPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    '������ݸ���ҳ����ȡ��ǰҳ����
    rsTemp.Filter = 0
    strTmp = mstrTimeRange(T_Info.lngҳ�� - 1) '��ǰҳ���ݷ�Χ
    strPageTime = mArrPageTime(T_Info.lngҳ�� - 1) '��ǰҳ��ʼʱ��
    strBegin = Format(Split(strTmp, ";")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Split(strTmp, ";")(1), "YYYY-MM-DD HH:mm:ss")
    rsTemp.Filter = "����ʱ��>='" & strBegin & "' And ����ʱ��<='" & strEnd & "'"
    If rsTemp.RecordCount = 0 Then Exit Sub
    rsTemp.Sort = "����ʱ��"
    '���Ƽ�¼���ֶ�
    Set rsTabData = CopyNewRec(rsTemp)
    gstrFields = ""
    For intFields = 0 To rsTabData.Fields.Count - 1
        gstrFields = gstrFields & "|" & rsTabData.Fields(intFields).Name
    Next intFields
    gstrFields = Mid(gstrFields, 2)
    '������������һ��
    With rsTemp
        Do While Not .EOF
            gstrValues = ""
            lngCol = Int((CDate(Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss")) - CDate(Format(strPageTime, "YYYY-MM-DD HH:mm:ss"))) * 24) + 1
            For intFields = 0 To rsTemp.Fields.Count - 1
                gstrValues = gstrValues & "|" & rsTemp.Fields(intFields).Value
            Next intFields
            gstrValues = Mid(gstrValues, 2) & "|" & lngCol
            Call Record_Update(rsTabData, gstrFields, gstrValues, "����ʱ��|" & Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss"))
        .MoveNext
        Loop
    End With
    
    '���������к�
    blnInit = False: lngCol = 0: lngColOld = 0
    rsTabData.Filter = ""
    rsTabData.Sort = "����ʱ��"
    With rsTabData
        Do While Not .EOF
            If lngCol <> Val(!�к�) Or blnInit = False Then
                lngCol = Val(!�к�)
                '73792:������,2014-06-23,�����λ�ü������
                If lngColOld - lngCol >= 0 Then
                    lngColOld = lngColOld + 1 ' (2 * lngColOld) - lngCol + 1
                Else
                    lngColOld = lngCol
                End If
            Else
                lngColOld = lngColOld + 1
            End If
            blnInit = True
            If lngColOld <> lngCol Then
                Call Record_Update(rsTabData, "����ʱ��|�к�", Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss") & "|" & lngColOld, "����ʱ��|" & Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss"))
            End If
        .MoveNext
        Loop
    End With
    
    arrItemOrder = Array()
    arrItemName = Array()
    arrItemHeight = Array()
    mrsSelItems.Filter = ""
    mrsSelItems.Filter = "�̶�=0"
    mrsSelItems.Sort = "��"
    Do While Not mrsSelItems.EOF
        ReDim Preserve arrItemOrder(UBound(arrItemOrder) + 1)
        ReDim Preserve arrItemName(UBound(arrItemName) + 1)
        ReDim Preserve arrItemHeight(UBound(arrItemHeight) + 1)
        lngOrder = Val(mrsSelItems!�������)
        arrItemOrder(UBound(arrItemOrder)) = lngOrder & ";" & NVL(mrsSelItems!Ҫ������)
        arrItemName(UBound(arrItemName)) = "C" & Format(lngOrder, "00")
        arrItemHeight(UBound(arrItemHeight)) = Val(mrsSelItems!�߶�)
    mrsSelItems.MoveNext
    Loop
    '��ʼ��ɱ���������
    rsTabData.Filter = ""
    rsTabData.Sort = "����ʱ��"
    With rsTabData
        Call SetTextColor(mlngDC, mTabForeColor)
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        Do While Not .EOF
            lngHeight = 0
            For intFields = 0 To UBound(arrItemOrder)
                strTmp = rsTabData.Fields(CStr(arrItemName(intFields)))
                lngOrder = Val(Split(CStr(arrItemOrder(intFields)), ";")(0))
                strName = Split(CStr(arrItemOrder(intFields)), ";")(1)
                lngCol = rsTabData.Fields("�к�")
                lngCurX = T_DrawClient.��������.Left + (lngCol - 1) * T_DrawClient.�е�λ
                lngCurY = T_DrawClient.�������.Top + lngHeight
                '�жԽ��ߵ�����
                If IsDiagonal(lngOrder) And InStr(1, strTmp, "/") <> 0 Then
                    strLeft = Split(strTmp, "/")(0)
                    strRight = Mid(strTmp, InStr(1, strTmp, "/") + 1)
                    '���Խ���
                    Call DrawLine(mlngDC, lngCurX, lngCurY + Val(arrItemHeight(intFields)), lngCurX + T_DrawClient.�е�λ, lngCurY, PS_SOLID, intFine, mTabGridColor)
                    '����ı�
                    Call GetTextRect(mobjDraw, lngCurX, lngCurY, strLeft, 0, False)
                    Call DrawText(mlngDC, strLeft, -1, T_LableRect, 0)
                    T_Size.H = mobjDraw.TextHeight(strRight) / sinTwipsPerPixelY
                    T_Size.W = mobjDraw.TextWidth(strRight) / sinTwipsPerPixelY
                    Call GetTextRect(mobjDraw, IIf(T_DrawClient.�е�λ - T_Size.W > 0, lngCurX + T_DrawClient.�е�λ - T_Size.W, lngCurX), lngCurY + Val(arrItemHeight(intFields)) - T_Size.H, strRight, 0, False)
                    Call DrawText(mlngDC, strRight, -1, T_LableRect, 0)
                ElseIf isBigConnect(strName, 1) = True Then '���ڳ���>10���ı���Ŀ�������������Ϣ
                    Call OutBigVConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.�е�λ, InStr(1, ",ǩ����,��ʿ,", "," & strName & ",") <> 0, True, mTabForeColor)
                ElseIf isBigConnect(strName, 0) Then '�Ƿ�����ֵ���͵���Ŀ
                    Call OutNumConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.�е�λ, IIf(strName = "����", True, False))
                Else
                    Call OutBigHConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.�е�λ, True)
                End If
                lngHeight = lngHeight + Val(arrItemHeight(intFields))
            Next intFields
        .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub OutNumConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, Optional ByVal bln���� As Boolean = False)
'���ܣ�����������Ŀ���
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long, lngY2 As Long
    Dim strLeft As String, strRight As String
    Dim bln���� As Boolean
    
    bln���� = True
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngX1 = lngX: lngY1 = lngY
    lngTmp = lngColWidth
    If InStr(1, strText, "/") <> 0 Then
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngWidth = mobjDraw.TextWidth(strLeft) / sinTwipsPerPixelX
        If lngWidth > lngTmp Then
            sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
            intSize = Round(Round((1 - sngD), 4) * intSize - 1)
            If intSize < 7 Then intSize = 7
            If intSize > intOldSize Then intSize = intOldSize
        End If
        mobjSubFont.Size = intSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        Call GetTextRect(mobjDraw, lngX, lngY, strLeft, T_DrawClient.�е�λ, False)
        Call DrawText(mlngDC, strLeft, -1, T_LableRect, DT_CENTER)
        T_Size.H = mobjDraw.TextHeight("��") / sinTwipsPerPixelY
        '���Խ���
        lngY2 = lngY1 + T_Size.H + (T_Size.H / 2)
        If lngY2 > lngY1 + lngHeight Then
            lngY2 = lngY1 + lngHeight
        End If
        Call DrawLine(mlngDC, lngX1, lngY2, lngX1 + T_DrawClient.�е�λ, lngY1 + (T_Size.H / 2), PS_SOLID, IIf(ISobjPrinter = True, 2, 1), mTabGridColor)
        lngHeight = lngHeight + lngY1 - lngY2
        lngY = lngY2
        lngY1 = lngY
        strText = Trim(strRight)
        '��ԭ����
        mobjSubFont.Size = intOldSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        bln���� = False
        GoTo ErrNext
    Else
ErrNext:
        '��С����������
        lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
        If lngWidth > lngTmp Then
            sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
            intSize = Round(Round((1 - sngD), 4) * 9 - 1)
            If intSize < 7 Then Call OutBigHConnect(strText, lngX1, lngY1, lngHeight, lngColWidth, bln����): GoTo ErrEnd
            If intSize > intOldSize Then intSize = intOldSize
        End If
        mobjSubFont.Size = intSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        T_Size.H = mobjDraw.TextHeight("��") / sinTwipsPerPixelY
        If bln���� = True Then
            lngY = lngY1 + (lngHeight - T_Size.H) / 2
            If lngY < lngY1 Then lngY = lngY1
        Else
            lngY = lngY1
        End If
        Call GetTextRect(mobjDraw, lngX, lngY, strText, T_DrawClient.�е�λ, False)
        Call DrawText(mlngDC, strText, -1, T_LableRect, DT_CENTER)
    End If
    
ErrEnd:
     '��ԭ����
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
End Sub

Private Function OutBigVConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, _
    Optional ByVal bln����Y As Boolean = False, Optional ByVal bln����X As Boolean = True, Optional ByVal lngColor As Long = 0) As Single
'���ܣ���������ı���Ϣ
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long
    Dim strTmp As String, strConnect As String
    Dim arrTmp
    
    '��ȡ��������
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
    lngTmp = lngHeight * Int(lngColWidth / (mobjDraw.TextWidth("��") / sinTwipsPerPixelX))
    If lngTmp <= 0 Then lngTmp = lngHeight
    If lngWidth > lngTmp Then
        sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
        intSize = Round(Round((1 - sngD), 4) * intSize - 1)
        If intSize < 7 Then intSize = 7
        If intSize > intOldSize Then intSize = intOldSize
    End If
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    '��ʼ����ı���Ϣ
    lngMaxY = lngY + lngHeight
    lngMaxX = lngX + lngColWidth
    lngX1 = lngX
    lngY1 = lngY
    strConnect = ""
    arrTmp = Array()
    
    lngX = lngX + (mobjDraw.TextWidth("��") / sinTwipsPerPixelX)
    ReDim arrTmp(UBound(arrTmp) + 1)
    For i = 1 To Len(strText)
        strTmp = Mid(strText, i, 1)
        If Asc(strTmp) > 0 Then
            T_Size.W = mobjDraw.TextWidth("��") / sinTwipsPerPixelX
            T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
        Else
            T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
            T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        End If
        If lngY + T_Size.H > lngMaxY And strConnect <> "" Then
            If (lngX + T_Size.W - lngMaxX) > (T_Size.W * 0.2) Then GoTo ErrEnd
            lngX = lngX + T_Size.W
            lngY = lngY1
            strConnect = ""
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
           GoTo ErrNext
        Else
ErrNext:
            strConnect = strConnect & strTmp
            arrTmp(UBound(arrTmp)) = strConnect
            lngY = lngY + T_Size.H + (1 * msngTwips)
        End If
    Next i
ErrEnd:
    T_Size.W = mobjDraw.TextWidth("��") / sinTwipsPerPixelX
    If lngColWidth / (UBound(arrTmp) + 1) > T_Size.W Then
        lngWidth = Int(lngColWidth / (UBound(arrTmp) + 1))
    Else
        lngWidth = T_Size.W
    End If
    For i = 0 To UBound(arrTmp)
        lngX = lngX1 + (i * T_Size.W)
        If bln����Y = True Then
            lngY = 0
            For j = 1 To Len(arrTmp(i))
                strTmp = Mid(arrTmp(i), j, 1)
                If Asc(strTmp) > 0 Then
                    T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
                Else
                    T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY
                End If
                lngY = lngY + T_Size.H + (1 * msngTwips)
            Next j
            lngY = lngY1 + (lngHeight - lngY) / 2
        Else
            lngY = lngY1
        End If
        
        For j = 1 To Len(arrTmp(i))
            strTmp = Mid(arrTmp(i), j, 1)
            If Asc(strTmp) > 0 Then
                T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
            Else
                T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY
            End If
            Call DrawRotateText(mobjDraw, mlngDC, lngX, lngY, strTmp, IIf(bln����X = True, lngWidth, 0), lngColor)
            lngY = lngY + T_Size.H + (1 * msngTwips)
        Next j
    Next i
    '��ԭ����
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    
    OutBigVConnect = lngX + T_Size.W
End Function

Private Function OutBigHConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, _
    Optional ByVal bln����Y As Boolean = True, Optional ByVal bln����X As Boolean = True) As Single
'���ܣ���������ı���Ϣ
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long
    Dim strTmp As String, strConnect As String
    Dim arrTmp
    
    '��ȡ��������
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
    lngTmp = lngHeight * Int(lngColWidth / (mobjDraw.TextWidth("��") / sinTwipsPerPixelX))
    If lngTmp <= 0 Then lngTmp = lngHeight
    If lngWidth > lngTmp Then
        sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
        intSize = Round(Round((1 - sngD), 4) * intSize - 1)
        If intSize < 7 Then intSize = 7
        If intSize > intOldSize Then intSize = intOldSize
    End If
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    '��ʼ����ı���Ϣ
    lngMaxY = lngY + lngHeight
    lngMaxX = lngX + lngColWidth
    lngX1 = lngX
    lngY1 = lngY
    strConnect = ""
    arrTmp = Array()
    
    lngY = lngY + (mobjDraw.TextHeight("��") / sinTwipsPerPixelY)
    ReDim arrTmp(UBound(arrTmp) + 1)
    For i = 1 To Len(strText)
        strTmp = Mid(strText, i, 1)
        If Asc(strTmp) > 0 Then
            T_Size.W = mobjDraw.TextWidth("��") / sinTwipsPerPixelX / 2
        Else
            T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
        End If
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        If lngX + T_Size.W > lngMaxX And strConnect <> "" Then
            If (lngY + T_Size.H - lngMaxY) > (T_Size.W * 0.2) Then GoTo ErrEnd
            lngY = lngY + T_Size.H + (1 * msngTwips)
            lngX = lngX1
            strConnect = ""
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
           GoTo ErrNext
        Else
ErrNext:
            strConnect = strConnect & strTmp
            arrTmp(UBound(arrTmp)) = strConnect
            lngX = lngX + T_Size.W
        End If
    Next i
ErrEnd:
    If bln����Y = True Then
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        lngY = lngY1 + (lngHeight - T_Size.H * (UBound(arrTmp) + 1)) / 2
        If lngY < lngY1 Then lngY = lngY1
        lngY1 = lngY
    End If
    lngWidth = lngColWidth
    For i = 0 To UBound(arrTmp)
        lngX = lngX1
        lngY = lngY1 + (i * (T_Size.H + (1 * msngTwips)))
        Call GetTextRect(mobjDraw, lngX, lngY, CStr(arrTmp(i)), IIf(bln����X = True, lngColWidth, 0), False)
        Call DrawText(mlngDC, CStr(arrTmp(i)), -1, T_LableRect, DT_CENTER)
    Next i
    '��ԭ����
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    
    OutBigHConnect = lngY + T_Size.H
End Function

Private Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "��������" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        If blnAddPage Then
            .Fields.Append "�к�", adDouble, 18
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function GetXCoordinate(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln���� As Boolean = True) As String

    '����ʱ��õ�X��������X����ת��Ϊʱ�䷶Χ
    Dim sinX   As Single
    Dim sinTime As Single
    Dim strDay As String

    On Error GoTo errHand
    
    If bln���� Then
        '�������ٷ���
        sinTime = Format(DateDiff("n", CDate(strBeginDate), CDate(strInput)) / 60, "#0.0000;-#0.0000;0000")
        
        '����õ�X����(ÿ��6��,������*�е�λ�õ�����)
        sinX = Format(T_DrawClient.��������.Left + (sinTime * T_DrawClient.�е�λ), "#0.0")
        GetXCoordinate = sinX
    Else
        '����õ������ٸ��̶�
        sinX = Val(strInput)
        sinTime = Format(sinX - T_DrawClient.��������.Left, "#0.0000;-#0.0000;0000") / T_DrawClient.�е�λ
        sinTime = Format(sinTime * 60, "#0.0")
        strDay = Format(DateAdd("n", sinTime, strBeginDate), "yyyy-MM-dd HH:mm:ss")
        GetXCoordinate = strDay
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetYCoordinate(ByVal int��Ŀ��� As Integer, ByVal strInput As String, Optional ByVal bln���� As Boolean = True) As Single

    Dim sinCurX As Single, sinCurY As Single, sinScale As Single
    On Error GoTo errHand
    '����ָ���������ݵ�Y��������Y�����������
    '���Ըú�������ȷ�Կ����ڼ�������Ӹô���ʵ��(˼��:�ɸú����Լ��������ݼ���õ�Y����,��ת��Ϊ����,��ת��Ϊ���������ַ����к˶�,��ӡ������˵��ת������):
    
    mrsDrawItems.Filter = "��Ŀ���=" & int��Ŀ���
    If mrsDrawItems.RecordCount = 0 Then
        GetYCoordinate = 0
        Exit Function
    End If
    
    If bln���� Then
        '�õ���Ч������ʼ����
        sinCurX = Split(mrsDrawItems!���ֵ����, ",")(0)
        sinCurY = Split(mrsDrawItems!���ֵ����, ",")(1)
        
        '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
        If Val(mrsDrawItems!��ʾģʽ) = 1 Then '�Ǵ���Сֵ�����ֵ
            sinScale = (-1 * (mrsDrawItems!��Сֵ - Val(strInput)) / mrsDrawItems!��λֵ) * Val(Split(mrsDrawItems!��λ�̶�, ",")(0))
        Else
            sinScale = ((mrsDrawItems!���ֵ - Val(strInput)) / mrsDrawItems!��λֵ) * Val(Split(mrsDrawItems!��λ�̶�, ",")(0))
        End If
        GetYCoordinate = Format(sinCurY + sinScale, "#0.0;-#0.0;0")
    Else
        '�õ����������ֵ
        sinCurY = CDbl(strInput)
        
        '(����-���ֵ����)/��λ�̶ȵõ������ٸ��̶�
        '(���ֵ-��λ�̶�*��λֵ)�õ�ʵ������
        sinScale = (sinCurY - Split(mrsDrawItems!���ֵ����, ",")(1)) / Val(Split(mrsDrawItems!��λ�̶�, ",")(0))
        If Val(mrsDrawItems!��ʾģʽ) = 1 Then  '�Ǵ���Сֵ�����ֵ
            GetYCoordinate = Format(mrsDrawItems!��Сֵ + sinScale * mrsDrawItems!��λֵ, "#0.0;-#0.0;0")
        Else
            GetYCoordinate = Format(mrsDrawItems!���ֵ - sinScale * mrsDrawItems!��λֵ, "#0.0;-#0.0;0")
        End If
    End If

    mrsDrawItems.Filter = ""
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetDrawItemValue(ByVal lngItemNo As Long, ByVal strFields As String) As String
'-----------------------------------------------------------------------
'����:�����ֶλ�ȡ��Ӧ��Ϣ
'-----------------------------------------------------------------------
    Dim strValue As String
    
    If InStr(1, "|��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��¼��|��ɫ|", "|" & strFields & "|") = 0 Then Exit Function
    
    mrsDrawItems.Filter = ""
    mrsDrawItems.Filter = "��Ŀ���=" & lngItemNo
    If mrsDrawItems.RecordCount > 0 Then
        strValue = mrsDrawItems.Fields(strFields).Value
    End If
    GetDrawItemValue = strValue
End Function

Private Function ISobjPrinter() As Boolean
'�ж��ͷ��Ǵ�ӡ������
    Dim blnPrinter As Boolean
    blnPrinter = (TypeName(mobjDraw) = "Printer")
    ISobjPrinter = blnPrinter
End Function

Private Sub InitEnv()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ,nvl(������Ŀ,0) ������Ŀ" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    '��ȡ���в���Ҫ����Ϣ
    gstrSQL = "Select ������,�滻��,����,����,С��,��λ,��ʾ��,��ֵ��,����" & vbNewLine & _
        "From (Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('02', '03', '05', '06') And i.�滻�� = 1 And k.���� = 1" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('04', '05') And k.���� = 2)" & vbNewLine & _
        "Order By ����id, ����, �滻��"

    Set mrsPartogram = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Ҫ����Ϣ")
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintState(ByVal lngFormatID As Long, Optional ByVal strPrintDevice As String = "") As Boolean
'******************************************************************************************************************
'����:���ô�ӡ������
'******************************************************************************************************************
    Dim i As Long
    Dim strPaper As String
    Dim strPrintName As String
    Dim blnYesPrinter As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    '��ӡ���ָ�������
    If Not ExistsPrinter Then
        MsgBox "ϵͳû�а�װ�κδ�ӡ�����ܼ�����ӡ�������˳���", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select ��ʽ,ҳ�� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", lngFormatID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    gPrinter.lngLeft = OFFSET_LEFT
    gPrinter.lngRight = OFFSET_RIGHT
    gPrinter.lngTop = OFFSET_TOP
    gPrinter.lngBottom = OFFSET_BOTTOM
       
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then
        gPrinter.intPage = Val(Split(mstrPaperSet, ";")(0))
        If UBound(Split(mstrPaperSet, ";")) >= 1 Then gPrinter.intOrient = Val(Split(mstrPaperSet, ";")(1))
        If UBound(Split(mstrPaperSet, ";")) >= 2 Then gPrinter.lngHeight = Val(Split(mstrPaperSet, ";")(2))
        If UBound(Split(mstrPaperSet, ";")) >= 3 Then gPrinter.lngWidth = Val(Split(mstrPaperSet, ";")(3))
        If UBound(Split(mstrPaperSet, ";")) >= 4 Then gPrinter.lngLeft = CLng(Val(Split(mstrPaperSet, ";")(4)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 5 Then gPrinter.lngRight = CLng(Val(Split(mstrPaperSet, ";")(5)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 6 Then gPrinter.lngTop = CLng(Val(Split(mstrPaperSet, ";")(6)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 7 Then gPrinter.lngBottom = CLng(Val(Split(mstrPaperSet, ";")(7)) / conRatemmToTwip)
    End If
    
    If strPrintDevice = "" Then
        If Trim(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName)) = "" Then
            MsgBox "û�����ô�ӡ��,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
        Else
            strPrintName = Trim(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName))
        End If
    Else
        strPrintName = strPrintDevice
    End If
    
    '��ӡ��
    blnYesPrinter = False
    If Printer.DeviceName <> strPrintName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrintName Then Set Printer = Printers(i): blnYesPrinter = True: Exit For
        Next
        If blnYesPrinter = False Then
            MsgBox "���õĴ�ӡ���Ѳ�����,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
        End If
    End If
            
    gPrinter.intBin = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", ""))
    
    On Error Resume Next
    'ֽ��
    If gPrinter.intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = gPrinter.lngWidth
        Printer.Height = gPrinter.lngHeight
    Else
        Printer.PaperSize = gPrinter.intPage
    End If
    Printer.Orientation = gPrinter.intOrient
    If IsWindowsNT And gPrinter.intPage = 256 Then
        Call SetNTPrinterPaper(frmSample.Hwnd, gPrinter.lngWidth / conRatemmToTwip, gPrinter.lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
        Unload frmSample
    End If
    
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    PrintState = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ�ļ�����
    mblnDateAd = False
    gstrSQL = " Select   ��ʽID From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", T_Info.lng����ID, T_Info.lng��ҳID, 0, T_Info.lng�ļ�ID)
    If rsTemp.RecordCount > 0 Then mlng��ʽID = rsTemp!��ʽID

    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mbln����ʱ��ϲ� = False
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":   mlngItems = Val("" & !�����ı�)
            Case "��С�и�": '
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = 9 ' Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjSubFont = objFont
                
            Case "�ı���ɫ": mTabForeColor = Val("" & !�����ı�)
            Case "�����ɫ": mTabGridColor = Val("" & !�����ı�)
            
            Case "�����ı�": mstrTitle = "" & !�����ı�
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTitleFont = objFont
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������"
                '
            Case "����ʱ��ϲ�"
                mbln����ʱ��ϲ� = (Val(!�����ı�) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubHead = ""
        Do While Not .EOF
            mstrSubHead = mstrSubHead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubHead <> "" Then mstrSubHead = Replace(Mid(mstrSubHead, 2), Chr(1), " ")
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���±�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubEnd = ""
        Do While Not .EOF
            mstrSubEnd = mstrSubEnd & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubEnd <> "" Then mstrSubEnd = Replace(Mid(mstrSubEnd, 2), Chr(1), " ")
    End With
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ' And d.�����д�=1" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String, strSqlNull As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
    
    gstrSQL = "Select   d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = "": strSqlNull = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    If Mid(strSqlNull, 3) = "" Then
                        strSqlNull = "''"
                    Else
                        strSqlNull = Mid(strSqlNull, 3)
                    End If
                    mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                    
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                strSqlNull = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                blnѡ���� = False
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    blnѡ���� = (mrsItems!��Ŀ��ʾ = 5)
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    'mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    If bln�Խ��� And blnѡ���� Then
                        If strSql�� <> "" Then
                            '�ڶ���
                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
                        Else
                            '��һ��
                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        End If
                    Else
                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
                    End If
                    
                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                        mstrSQL���� = mstrSQL���� & " Or Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') Is Not Null"
                    Else
                        'mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
                        mstrSQL���� = mstrSQL���� & " Or Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') Is Not Null"
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')),  '" & !�����ı� & "'||'" & !Ҫ�ص�λ & "') As """ & !Ҫ������ & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
    End With
    '��ʼ�������Ϣ
    Call InitTabRecords
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub GetFileProperty()
    '��ȡ�ļ�����
    Dim rsTemp As New ADODB.Recordset
    Dim strEnd As String
    On Error GoTo errHand
    
    gstrSQL = " Select   ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", T_Info.lng����ID, T_Info.lng��ҳID, 0, T_Info.lng�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")
        mstr����ʱ�� = Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm")
        strEnd = DateAdd("n", -1, CDate(Format(CDate(mstr��ʼʱ��) + 1, "yyyy-MM-dd HH:mm:ss")))
        If mstr����ʱ�� = "" Then
            mstr����ʱ�� = strEnd
        Else
            If (mstr����ʱ�� <> "" And CDate(mstr����ʱ��) > CDate(strEnd)) Then mstr����ʱ�� = strEnd
        End If
    End If
    '�ڶ����ļ�������ȡ��ʼʱ��
    If T_Info.lng���� > 1 Then
        gstrSQL = "SELECT Max(B.����ʱ��) ����ʱ��" & vbNewLine & _
            "FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
            "WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] And ����ID=[2] And ��ҳID=[3] And Ӥ��=[4] AND B.�������<[5] AND C.��Ŀ���=D.��Ŀ���" & vbNewLine & _
            "AND NVL(D.��Ŀ����,'')='����' AND NVL(D.������Ŀ,1)=1 ORDER BY B.����ʱ��"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID, 0, T_Info.lng����)
        If rsTemp.RecordCount <> 0 Then
            mstr��ʼʱ�� = DateAdd("n", 1, CDate(Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")))
        End If
    End If
    
    '��ȡ�ļ�ҳ��
    Call GetPartogramPage(T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SQLCombination()
'-��ȡ����SQL
    mstrSQL = "Select to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select �������,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select l.�������,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.�ļ�ID=f.ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And l.�������=[5]" & IIf(mstrSQL���� <> "", " And (" & mstrSQL���� & ")", "") & ")" & vbCrLf & _
                "       Group By ����, ʱ��, ����ʱ��,�������,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By �������,����ʱ��,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Sub SQLCombinationPage()
'-��ȡ�����к�SQL
    mstrSQL = " SELECT ����ʱ��,FLOOR(TO_NUMBER(����ʱ��-��ʼʱ��)*24)+1 AS ��" & vbCrLf & _
                " From (Select �������,ʱ�� as ����,����ʱ��,Max(��ʼʱ��) ��ʼʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select l.�������,l.����ʱ��,F.��ʼʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.�ļ�ID=f.ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And l.�������=[5]" & IIf(mstrSQL���� <> "", " And (" & mstrSQL���� & ")", "") & ")" & vbCrLf & _
                "       Group By ����, ʱ��, ����ʱ��,�������,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By �������,����ʱ��,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Function GetPeriod() As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = " Select   ��Ժ���� AS ��ʼʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", T_Info.lng����ID, T_Info.lng��ҳID)
    GetPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(mstr����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsDiagonal(ByVal lngOrder As Long) As Boolean
    '�ж�ָ�����Ƿ��������жԽ���
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & lngOrder & ",") <> 0)
End Function

Private Function isBigConnect(ByVal strName As String, ByVal bytMode As Byte) As String
'------------------------------------------------------
'���ܣ��ж���Ŀ����
'bytMode:0-��ֵ;1-���ı�(����>10);2-����(�磺��ѡ;��ѡ;�ı�����<10���ı���Ŀ)
'------------------------------------------------------
    Dim blnOK As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "��Ŀ����='" & strName & "'"
    If mrsItems.RecordCount > 0 Then
        Select Case bytMode
        Case 0
            If Val(NVL(mrsItems!��Ŀ����, 0)) = 0 Then
                blnOK = True
            End If
        Case 1
            If Val(NVL(mrsItems!��Ŀ����, 0)) = 1 And Val(NVL(mrsItems!��Ŀ��ʾ, 0)) = 0 Then
                blnOK = Val(NVL(mrsItems!��Ŀ����, 0)) > 10
            End If
        Case Else
            blnOK = True
        End Select
    Else
        Select Case bytMode
        Case 0
            If strName = "ʱ��" Or strName = "����" Then blnOK = True
        Case 1
            If strName = "ǩ����" Or strName = "��ʿ" Then blnOK = True
        Case Else
            blnOK = True
        End Select
    End If
    isBigConnect = blnOK = True
End Function


Private Sub GetMarkConnect()
'----------------------------------------------------------------------
'����:��ȡ���±���Ϣ
'----------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim aryRow() As String, aryItem() As String, arrItemEnd() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strSubHend As String, strSubEnd As String
    Dim strTmp As String, str��λ As String
    Dim i As Integer
    
    On Error GoTo errHand
    
    strSubHend = ""
    strSubEnd = ""
    aryPeriod = Split(GetPeriod, "��")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubHead, "|")
    arrItemEnd = Split(mstrSubEnd, "|")
    For i = 0 To 1
        For lngCount = 0 To IIf(i = 0, UBound(aryItem), UBound(arrItemEnd))
            If i = 0 Then
                strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
                strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
            Else
                strPrefix = Left(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") - 1)
                strItemName = Mid(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") + 1, InStr(1, arrItemEnd(lngCount), "}") - InStr(1, arrItemEnd(lngCount), "{") - 1)
            End If
            mrsPartogram.Filter = 0
            mrsPartogram.Filter = "������='" & strItemName & "'"
            '�������Ҳ����������ֹ��޸�����
            If mrsPartogram.RecordCount = 0 Then GoTo ErrNext
            str��λ = Trim(NVL(mrsPartogram!��λ))
            If Val(NVL(mrsPartogram!�滻��)) = 1 Then
                '���̶̹�Ҫ����Ϣ
                strTmp = strPrefix
                Select Case strItemName
                Case "��ǰ����"
                
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                                
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   a.����" & vbNewLine & _
                                "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
        
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   ���� From ���ű� a Where a.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", T_Info.lng����ID)
                    
                Case "סԺҽʦ"
                    strTmpSQL = "Select   a.����ҽʦ" & vbNewLine & _
                                "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "���λ�ʿ"
                
                    strTmpSQL = "Select   a.���λ�ʿ" & vbNewLine & _
                                "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "����ȼ�"
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "������"
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, T_Info.lng����ID, T_Info.lng��ҳID, 0, CDate(aryPeriod(0)))
                Case Else
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, T_Info.lng����ID, T_Info.lng��ҳID, 0)
                End Select
            Else
                '����¼��Ҫ����Ϣ
                strTmp = strPrefix
                gstrSQL = "SELECT ���� From ����Ҫ������" & _
                    "   Where �ļ�ID = [1] And Ӥ�� = [2] And ���� =[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", T_Info.lng�ļ�ID, T_Info.lng����, strItemName)
            End If
            If rsTemp.BOF = False Then
                If i = 0 Then
                    If strTmp <> "" Then
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & strTmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & rsTemp.Fields(0).Value & str��λ
                    End If
                Else
                    If strTmp <> "" Then
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & strTmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & rsTemp.Fields(0).Value & str��λ
                    End If
                End If
            Else
                If i = 0 Then
                    If strTmp <> "" Then
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & strTmp
                    Else
                        strSubHend = strSubHend & "[ZLSOFTLPF]"
                    End If
                Else
                    If strTmp <> "" Then
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & strTmp
                    Else
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]"
                    End If
                End If
            End If
ErrNext:
        Next
    Next i
    If Left(strSubHend, 11) = "[ZLSOFTLPF]" Then strSubHend = Mid(strSubHend, 12)
    If Left(strSubEnd, 11) = "[ZLSOFTLPF]" Then strSubEnd = Mid(strSubEnd, 12)
    mstrOutSubHead = strSubHend
    mstrOutSubEnd = strSubEnd
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFileCount(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
'��ȡ�ļ�����
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo errHand
    mstrSQL = "SELECT NVL(MAX(NVL(�������,1)),1) �ļ�" & vbNewLine & _
            "FROM ���˻����ļ� A,���˻������� B" & vbNewLine & _
            "WHERE A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.ID=B.�ļ�ID"
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�ļ�����", lng�ļ�ID, lng����ID, lng��ҳID)
    If rsTemp.RecordCount = 0 Then
        lngCount = 1
    Else
        lngCount = Val(NVL(rsTemp!�ļ�, 1))
    End If
    If lngCount < 1 Then lngCount = 1
    GetFileCount = lngCount
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabRecords()
    Dim i As Integer, j As Integer, k As Integer
    Dim lngCol As Long, strName As String, lngColHeight As Long, lngOrder As Long, lngItemNo As Long, strName1 As String
    Dim arrColumn, arrItem, arrWidth, strColumns As String, strTabHead As String, strColWidth As String
    On Error GoTo errHand
    
    strColumns = mstrColumns
    strTabHead = mstrTabHead
    strColWidth = mstrColWidth
    '��ʼ������ʽ���ݼ�¼��
    gstrFields = "��," & adDouble & ",18|�������," & adDouble & ",18|����," & adLongVarChar & ",20|�߶�," & adDouble & ",18|�̶�," & adInteger & ",1|Ҫ������," & adLongVarChar & ",20"
    Call Record_Init(mrsSelItems, gstrFields)
    gstrFields = "��|�������|����|�߶�|�̶�|Ҫ������"
    
    arrColumn = Split(strColumns, "|")
    arrItem = Split(strTabHead, "|")
    arrWidth = Split(strColWidth, ",")
    j = UBound(arrColumn)
    lngCol = 1
    For i = 0 To j
        lngOrder = Split(arrColumn(i), "'")(0)
        k = UBound(Split(Split(arrColumn(i), "'")(1), ","))
        strName1 = Split(Split(arrColumn(i), "'")(1), ",")(0)
        lngColHeight = Format(CDbl((Split(arrWidth(i), "`")(0)) / sinTwipsPerPixelX), "0.00")
        '�̶���Ŀ��Ϊ������������¶�ߵ͡�����������
        mrsItems.Filter = "��Ŀ����='" & strName1 & "' And ������Ŀ=1"
        If mrsItems.RecordCount = 0 Then
ErrAdd:
            strName = Split(arrItem(i), ",")(2)
            gstrValues = lngCol & "|" & lngOrder & "|" & strName & "|" & lngColHeight & "|0|" & strName1
            Call Record_Add(mrsSelItems, gstrFields, gstrValues)
            lngCol = lngCol + 1
        Else
            Select Case strName1
                Case "��������"
                    lngItemNo = T_Partogram.lng��������
                Case "��¶�ߵ�"
                    lngItemNo = T_Partogram.lng��¶�ߵ�
                Case "����"
                    lngItemNo = T_Partogram.lng����
                Case "����"
                    lngItemNo = T_Partogram.lng����
                Case Else
                    GoTo ErrAdd
            End Select
            gstrValues = lngItemNo & "|" & lngOrder & "|" & strName1 & "|" & lngColHeight & "|1|" & strName1
            Call Record_Add(mrsSelItems, gstrFields, gstrValues)
            If strName1 = "����" Or k > 0 Then GoTo ErrAdd
        End If
    Next
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DrawPartogramTab() As Single
'-------------------------------------------------------------------------
'����:���ݹ���Ĳ���¼�룬�����̱������
'-------------------------------------------------------------------------
    On Error GoTo errHand
    Dim lngCurX As Long, lngCurY As Long, lngCount As Long
    Dim lngCurMaxY As Long, lngCurMaxX As Long, lngHeight As Long
    Dim intBold As Integer, intFine As Integer
    Dim blnPrinter As Boolean, blnInit As Boolean
    Dim strConnect As String, strConnect1 As String, strConnect2 As String
    Dim i As Integer
    
    If TypeName(mobjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '----���Ȼ���ͷ����
    lngCurX = T_DrawClient.�������.Left
    lngCurY = T_DrawClient.�������.Top
    lngCurMaxX = T_DrawClient.MaxX
    lngCurMaxY = lngCurY
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurMaxX, lngCurY, PS_SOLID, intFine, mTabGridColor)
    Call SetTextColor(mlngDC, RGB_BLACK)
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    mrsSelItems.Filter = "�̶�=0"
    mrsSelItems.Sort = "��"
    If mrsSelItems.RecordCount = 0 Then GoTo ErrOver
    lngHeight = 0
    blnInit = False
    strConnect2 = ""
    lngCount = 0
    '����ͷ���ݣ���ͷ���ݿ��ܴ��ںϲ������
    With mrsSelItems
        Do While Not .EOF
            strConnect1 = Trim(NVL(mrsSelItems!����))
            If blnInit = False Then strConnect2 = strConnect1
            If strConnect1 <> strConnect2 Then
ErrEnd:
                strConnect = CheckConnect(strConnect2, T_DrawClient.�������.Right - T_DrawClient.�������.Left, lngHeight)
                '���ݱ��Ŀ�Ⱥ͸߶ȼ������
                T_Size.H = mobjDraw.TextHeight(strConnect) / sinTwipsPerPixelY
                T_Size.H = (lngHeight - T_Size.H) / 2
                Call GetTextRect(mobjDraw, lngCurX, lngCurY + T_Size.H, strConnect, T_DrawClient.�̶ȵ�λ, False)
                Call DrawText(mlngDC, strConnect, -1, T_LableRect, DT_CENTER)
                Call DrawLine(mlngDC, lngCurX, lngCurY + lngHeight, T_DrawClient.�������.Right, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                lngCurY = lngCurY + lngHeight
                lngHeight = 0
                strConnect2 = strConnect1
            End If
            lngHeight = lngHeight + Val(mrsSelItems!�߶�)
            blnInit = True
            lngCount = lngCount + 1
            If mrsSelItems.RecordCount = lngCount Then GoTo ErrEnd
        .MoveNext
        Loop
    End With
    '������
    mrsSelItems.Filter = "�̶�=0"
    mrsSelItems.Sort = "��"
    lngCurY = T_DrawClient.�������.Top
    With mrsSelItems
        Do While Not .EOF
            lngCurY = lngCurY + Val(mrsSelItems!�߶�)
             Call DrawLine(mlngDC, T_DrawClient.�������.Right, lngCurY, lngCurMaxX, lngCurY, PS_SOLID, intFine, mTabGridColor)
        .MoveNext
        Loop
    End With
    lngCurMaxY = lngCurY
    '����ߵ�����
    lngCurX = T_DrawClient.�������.Left
    lngCurY = T_DrawClient.�������.Top
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngCurMaxY, PS_SOLID, intFine, mTabGridColor)
    
    lngCurY = T_DrawClient.�������.Top
    lngCurX = T_DrawClient.�������.Right
    '������
    For i = 0 To 24
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngCurMaxY, PS_SOLID, intFine, mTabGridColor)
        lngCurX = lngCurX + T_DrawClient.�е�λ
    Next i
ErrOver:
    DrawPartogramTab = lngCurMaxY
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckConnect(ByVal strValue As String, lngWidth As Long, lngHeight As Long) As String
'���ܣ����ݱ��Ŀ�ȼ����������
    Dim strConnect As String, strTmp As String
    Dim i As Integer
    Dim lngX As Long, lngY As Long
    For i = 1 To Len(strValue)
        strTmp = Trim(Mid(strValue, i, 1))
        T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        If lngX + T_Size.W > lngWidth Then
            lngY = lngY + T_Size.H
            If lngY > lngHeight Then Exit For
            lngX = 0
            If strConnect = "" Then
                strConnect = strTmp
            Else
                strConnect = strConnect & vbCrLf & strTmp
            End If
        Else
            lngX = lngX + T_Size.W
            strConnect = strConnect & strTmp
        End If
    Next i
    
    CheckConnect = strConnect
End Function

Private Function DrawPartogram(ByVal strPartogram As String) As Single
'----------------------------------------------------------------
'���ܣ������̶̿�����Ͳ�����������
'������strPartogram :��¼��[LPF]��¼��[LPF]��¼ɫ[LPF]���ֵ[LPF]��Сֵ[LPF]��λֵ[LPF]��λ[|LPF|]��¼��[LPF]��¼��...
'------------------------------------------------------------------
    
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Single  '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Single  '�߽�
    Dim lngRow      As Long
    Dim intLables   As Integer
    '���¶��Ǳ�׼�߶�
    Dim intLineMode   As Integer
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single
    Dim arrTemp()     As String, ArrCode() As String, strTmp As String, strConnect As String
    Dim i As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngValue As String
    Dim blnInit As Boolean '��ʾ�Ƿ��һ����ʾ�̶�
    Dim blnDesc As Boolean '���ڱ�ʾ��¶�ߵ��Ƿ�����ʾ
    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double
    Dim str���ֵ���� As String, str��Сֵ���� As String
    
    '����������ͼģʽ����¶�ߵ���ʾλ���Լ��Ƿ���ʾ����ʱ��
    Dim intģʽ As Integer, int��¶λ�� As Integer
    Dim bln����ʱ�� As Boolean
    
    On Error GoTo errHand
    If TypeName(mobjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '����������Ŀ����ͼ����(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    gstrFields = "��Ŀ���," & adDouble & ",18|���ֵ," & adDouble & ",18|��Сֵ," & adDouble & ",18|" & "��λֵ," & adDouble & _
        ",18|���ֵ����," & adLongVarChar & ",20|��Сֵ����," & adLongVarChar & ",20|" & "��λ�̶�," & adLongVarChar & ",20|" & _
        "��ʾģʽ," & adDouble & ",5|��¼��," & adLongVarChar & ",10|��ɫ," & adDouble & ",18"
    Call Record_Init(mrsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '����ֵ
    intLineMode = PS_SOLID
    lngLableStep = T_DrawClient.�̶ȵ�λ
    lngColStep = T_DrawClient.�е�λ
    sinRowStep = T_DrawClient.�е�λ
    
    '����ͼģʽ����¶�ߵ���ʾλ��
    intģʽ = Val(zlDatabase.GetPara("����ͼģʽ", glngSys, 1255, 0)) '0-����ʽ 1-����ʽ
    int��¶λ�� = Val(zlDatabase.GetPara("��¶�ߵ���ʾλ��", glngSys, 1255, 0)) '0-��ʾ���Ҳ�,1-��ʾ�����
    bln����ʱ�� = (Val(zlDatabase.GetPara("����ͼ��ʾ����ʱ��", glngSys, 1255, 0)) = 1) '0-����ʾ,1-��ʾ
    arrTemp = Split(strPartogram, "[|LPF|]")
    '�����
    intLables = UBound(arrTemp) + 1
    lngCurX = T_DrawClient.ƫ����.X
    lngCurY = T_DrawClient.�̶�����.Top
    lngMaxX = T_DrawClient.MaxX 'ƫ����+�̶�����+24*�е�λ
    lngMaxY = T_DrawClient.��������.Bottom '��ʼ����+����*�е�λ
        
    SlngMaxY = lngMaxY
    
    If int��¶λ�� = 1 Then
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY + msnTimeH + IIf(bln����ʱ�� = True, msnTimeH, 0), PS_SOLID, intFine, RGB_BLACK)
    End If
    lngValue = Val(Split(arrTemp(1), "[LPF]")(3))
    '������ͼ������
    lngCurX = T_DrawClient.��������.Left
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY + msnTimeH + IIf(bln����ʱ�� = True, msnTimeH, 0), PS_SOLID, intFine, RGB_BLACK)

    For lngRow = 0 To T_DrawClient.������
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '������ͼ��������
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, IIf(lngValue = 0, intBold, intFine), RGB_BLACK)
        lngValue = lngValue - 1
    Next
    
    lngCurY = T_DrawClient.�̶�����.Top
    
    '������ͼ������
    For lngRow = 1 To 24
        lngCurX = lngCurX + lngColStep
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intFine, RGB_BLACK)
    Next

    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
'    gstdset.Name = "����"
'    gstdset.Size = 9
'    gstdset.Bold = False
'    gstdset.Italic = False
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngDC, mlngFont)
    
    For lngRow = 0 To UBound(arrTemp)
        ArrCode = Split(arrTemp(lngRow), "[LPF]")
        '��ʾ�̶ȿ���Ŀ������
        If int��¶λ�� = 1 Then '��¶�ߵ���ʾ�����
            strTmp = ArrCode(0) & ArrCode(6)
            lngCurX = T_DrawClient.�̶�����.Left + lngRow * (T_DrawClient.�̶ȵ�λ / 2)
            lngCurY = T_DrawClient.�̶�����.Top - Val(Format((Len(strTmp) / 2), "#0")) * mobjDraw.TextHeight(strConnect) / sinTwipsPerPixelY
            For i = 1 To Val(Format((Len(strTmp) / 2), "#0"))
                '������Ŀ����
                strConnect = Trim(Mid(strTmp, i * 2 - 1, 2))
                Call SetTextColor(mlngDC, ArrCode(2))
                Call GetTextRect(mobjDraw, lngCurX, lngCurY, Trim(strConnect), T_DrawClient.�̶ȵ�λ / 2, False)
                Call DrawText(mlngDC, Trim(strConnect), -1, T_LableRect, DT_CENTER)
                lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY
            Next i
        Else '��¶�ߵ���ʾ���Ҳ�
            strTmp = ArrCode(0) & "(" & ArrCode(6) & ")" '& IIf(int��¶λ�� = 0, arrCode(1), "")
            If lngRow = 0 Then
                lngCurX = T_DrawClient.�̶�����.Left
            Else
                lngCurX = T_DrawClient.MaxX + T_DrawClient.�̶ȵ�λ / 2
            End If
            lngCurY = T_DrawClient.�̶�����.Top + ((T_DrawClient.�̶�����.Bottom - T_DrawClient.�̶�����.Top - ((Len(strTmp) - 1) * mobjDraw.TextHeight("1") / sinTwipsPerPixelY)) / 2)
            For i = 1 To Len(strTmp)
                strConnect = Mid(strTmp, i, 1)
                Call SetTextColor(mlngDC, ArrCode(2))
                Call DrawRotateText(mobjDraw, mlngDC, lngCurX, lngCurY, Trim(strConnect), T_DrawClient.�̶ȵ�λ / 2, ArrCode(2))
                If Asc(strConnect) < 0 And strConnect <> "��" Then
                    lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY
                Else
                    lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
                End If
            Next i
        End If
        
        '���п̶���ֵX�������
        If int��¶λ�� = 1 Then
            lngCurX = T_DrawClient.�̶�����.Left + lngRow * (T_DrawClient.�̶ȵ�λ / 2)
        Else
            If lngRow = 0 Then
                lngCurX = T_DrawClient.�̶�����.Left + (T_DrawClient.�̶ȵ�λ / 2)
            Else
                lngCurX = T_DrawClient.MaxX
            End If
        End If
        lngCurY = T_DrawClient.�̶�����.Top
        dbl��λֵ = 1
        sin�̶ȼ�� = 1
        
        blnDesc = False
        If lngRow = 1 And intģʽ = 1 Then blnDesc = True
        blnInit = False
        Do While True
            bln��ʾ�̶� = False
            If blnInit = False Then      '�ս���ѭ������ʱȡ�����ֵ
                sin�̶� = IIf(blnDesc = True, Val(ArrCode(4)), Val(ArrCode(3)))
                sinBegin�̶� = sin�̶�
                str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                blnInit = True
            Else                    '����õ�ÿ���̶ȵ�ֵ
                sin�̶� = sin�̶� - (IIf(blnDesc = True, -1, 1) * dbl��λֵ)
            End If
            
            '�������õĿ̶ȼ����ʾ�̶�ֵ
            If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
            If bln��ʾ�̶� = True Then sinBegin�̶� = sinBegin�̶� - (IIf(blnDesc = True, -1, 1) * sin�̶ȼ��)
            If bln��ʾ�̶� Then
                Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ / 2, _
                    IIf(IIf(blnDesc = True, Val(ArrCode(4)), Val(ArrCode(3))) = Val(Format(sin�̶�, "#0")), False, True))
                Call DrawText(mlngDC, Format(sin�̶�, "#0"), -1, T_LableRect, DT_CENTER)
            End If
            '���������Ч��Χ�ڣ����߳����������˳�
            If Val(Format(sin�̶�, "#0.00")) = Val(Format(IIf(blnDesc = True, ArrCode(3), ArrCode(4)), "#0.00")) Or lngCurY > T_DrawClient.�̶�����.Bottom Then
                str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��¼��|��ɫ"
                gstrValues = IIf(lngRow = 0, T_Partogram.lng��������, T_Partogram.lng��¶�ߵ�) & "|" & Val(ArrCode(3)) & "|" & Val(ArrCode(4)) & _
                "|" & dbl��λֵ & "|" & str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & IIf(blnDesc = True, 1, 0) & "|" & ArrCode(1) & "|" & ArrCode(2)
                Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                Exit Do
            End If
            lngCurY = lngCurY + T_DrawClient.�е�λ
        Loop
    Next lngRow
    
    '�������µ�ʱ����
    lngValue = mobjDraw.TextWidth("12") / sinTwipsPerPixelX
    lngCurX = T_DrawClient.�̶�����.Right - (lngValue / 2)
    lngCurY = T_DrawClient.��������.Bottom + (msnTimeH / 2)
    Call SetTextColor(mlngDC, RGB_BLACK)
    For lngRow = 1 To 24
        lngCurX = lngCurX + T_DrawClient.�е�λ
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, lngRow, lngValue)
        Call DrawText(mlngDC, lngRow, -1, T_LableRect, DT_CENTER)
    Next lngRow
    '�������ʱ��
    If bln����ʱ�� = True Then
        lngCurX = T_DrawClient.�̶�����.Left
        lngCurY = T_DrawClient.��������.Bottom + msnTimeH + (msnTimeH / 2)
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, "����ʱ��", T_DrawClient.�̶ȵ�λ)
        Call DrawText(mlngDC, "����ʱ��", -1, T_LableRect, DT_CENTER)
        lngCurX = T_DrawClient.��������.Left
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, mstr����ʱ��, (mobjDraw.TextWidth(mstr����ʱ��) / sinTwipsPerPixelX) + 2)
        Call DrawText(mlngDC, mstr����ʱ��, -1, T_LableRect, DT_CENTER)
    End If
    Call SelectObject(mlngDC, mlngOldFont)
    Call DeleteObject(mlngFont)

    DrawPartogram = lngCurY + (msnTimeH / 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub GetPartogramPage(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileIndex As Long)
'--------------------------------------------------------------------------
'���ܣ���ȡ�ļ�ҳ��
'����;�ļ�ID������ID����ҳID����¼���
'--------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long, blnInit As Boolean
    Dim ArrTime, ArrCode() As String
    Dim strTmp As String, intPage As Integer, strEnd As String
    Dim strBeginTime As String
    On Error GoTo errHand
    If Not mblnPrint Then mintMaxPage = 1 'Ԥ����ӡʱ���ܸ��´˱���������ᵼ�²���ͼչ��ҳ������
    intPage = 1
    lngCol = 0
    
    strBeginTime = Format(mstr����ʱ��, "YYYY-MM-DD HH:mm:ss")
    ArrTime = Array()
    mArrPageTime = Array()
    ReDim ArrTime(UBound(ArrTime) + 1)
    ReDim mArrPageTime(UBound(mArrPageTime) + 1)
    ArrTime(UBound(ArrTime)) = strBeginTime & ";" & Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss")
    mArrPageTime(UBound(mArrPageTime)) = Format(strBeginTime, "YYYY-MM-DD HH:mm:ss")
    
'    gstrSQL = _
'        " SELECT ����ʱ��,FLOOR(TO_NUMBER(����ʱ��-��ʼʱ��)*24)+1 AS ��" & vbNewLine & _
'        " FROM (" & vbNewLine & _
'        " SELECT ����ʱ��,Max(A.��ʼʱ��) ��ʼʱ�� FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & vbNewLine & _
'        " WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=0" & vbNewLine & _
'        " AND B.�������=[4]" & vbNewLine & _
'        " GROUP BY B.����ʱ�� ORDER BY B.����ʱ��)"
    Call SQLCombinationPage
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "���˻����ļ�", lngFileID, lngPatiID, lngPageId, 0, lngFileIndex)
    If rsTemp.RecordCount > 0 Then
        With rsTemp
            blnInit = False: lngCol = 0
            Do While Not .EOF
                '73792:������,2014-06-23,�����λ�ü������
                If blnInit = False Then strEnd = Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                If lngCol <> Val(!��) Then
                    lngCol = Val(!��)
                    If lngColOld - lngCol >= 0 Then
                        lngColOld = lngColOld + 1 '(2 * lngColOld) - lngCol + 1
                    Else
                        lngColOld = lngCol
                    End If
                Else
                    lngColOld = lngColOld + 1
                End If
                
                blnInit = True
                
                If intPage <> (Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)) Then
                    intPage = (Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1))
                    ReDim Preserve ArrTime(UBound(ArrTime) + 1)
                    '���ȸ�����һҳ�����ʱ��
                    ArrCode = Split(ArrTime(UBound(ArrTime) - 1), ";")
                    If CDate(strEnd) > CDate(Format(ArrCode(1), "YYYY-MM-DD HH:mm:ss")) Then
                        strEnd = CDate(Format(ArrCode(1), "YYYY-MM-DD HH:mm:ss"))
                    End If
                    ArrTime(UBound(ArrTime) - 1) = ArrCode(0) & ";" & strEnd
                    '���±�ҳ�Ŀ�ʼʱ��
                    strTmp = Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                    strEnd = Format(DateAdd("H", (intPage * 24) - lngColOld, CDate(strTmp)), "YYYY-MM-DD HH:mm:ss")
                    If CDate(strEnd) > CDate(Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss")) Then
                        strEnd = CDate(Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss"))
                    End If
                    ArrTime(UBound(ArrTime)) = strTmp & ";" & strEnd
                    '���±�ҳ��ʼʱ��
                    ReDim Preserve mArrPageTime(UBound(mArrPageTime) + 1)
                    strTmp = DateAdd("n", -1 * ((lngColOld Mod 24) - 1) * 60, CDate(Format(strTmp, "YYYY-MM-DD HH:mm:ss")))
                    mArrPageTime(UBound(mArrPageTime)) = strTmp
                Else
                    strEnd = Format(!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                End If
            .MoveNext
            Loop
        End With
    End If
    
    If lngColOld <= 0 Then lngColOld = 1
    If Not mblnPrint Then mintMaxPage = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    mstrTimeRange = ArrTime
    mintPageCount = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetFilePage(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileIndex As Long) As Long
'�����ļ��ţ���ȡҳ��
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long
    Dim intPage As Long
    
    On Error GoTo errHand
    
'    gstrSQL = _
'        " SELECT ����ʱ��,FLOOR(TO_NUMBER(����ʱ��-��ʼʱ��)*24)+1 AS ��" & vbNewLine & _
'        " FROM (" & vbNewLine & _
'        " SELECT ����ʱ��,MAX(A.��ʼʱ��) ��ʼʱ�� FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & vbNewLine & _
'        " WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=0" & vbNewLine & _
'        " AND B.�������=[4]" & vbNewLine & _
'        " GROUP BY B.����ʱ�� ORDER BY B.����ʱ��)"
    Call SQLCombinationPage
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "���˻����ļ�", lngFileID, lngPatiID, lngPageId, 0, lngFileIndex)
    If rsTemp.RecordCount > 0 Then
        With rsTemp
            Do While Not .EOF
                If lngCol <> Val(!��) Then
                    lngCol = Val(!��)
                    '73792:������,2014-06-23,�����λ�ü������
                    If lngColOld - lngCol >= 0 Then
                        lngColOld = lngColOld + 1 ' (2 * lngColOld) - lngCol + 1
                    Else
                        lngColOld = lngCol
                    End If
                Else
                    lngColOld = lngColOld + 1
                End If
            .MoveNext
            Loop
        End With
    End If
    If lngColOld <= 0 Then lngColOld = 1
    intPage = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    GetFilePage = intPage
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SetFontIndirect(ByVal stdSet As StdFont, ByVal lngDc As Long, ByVal ObjDraw As Object)

    '����:������������
    Dim BFileName() As Byte
    Dim i As Integer

    On Error GoTo errHand
    
    ObjDraw.Font.Size = stdSet.Size
    BFileName = StrConv(stdSet.Name, vbFromUnicode)
    With T_Font
        For i = 1 To Len(stdSet.Name)
            .lfFaceName(i - 1) = BFileName(i - 1)
        Next i
        .lfHeight = -MulDiv(stdSet.Size, GetDeviceCaps(lngDc, LOGPIXELSY), 71)
        .lfWidth = 0
        .lfWeight = IIf(stdSet.Bold = True, FW_BOLD, FW_NORMAL)
        .lfCharSet = stdSet.Charset
        .lfUnderline = stdSet.Underline
        .lfItalic = stdSet.Italic
        .lfStrikeOut = stdSet.Strikethrough
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub GetTextRect(ByVal ObjDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln���� As Boolean = True, Optional ByVal lngHeght As Long = 0, Optional ByVal sngScale As Single = 1)
    
    '��ȡ����������Ч����
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + 1 '��������߽绮���غ�
    
    If bln���� = True Then
        T_LableRect.Top = lngY - ObjDraw.TextHeight(strInput) / 2 / sinTwipsPerPixelY
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY + T_LableRect.Left + 2
    T_LableRect.Bottom = ObjDraw.TextHeight(strInput) / sinTwipsPerPixelY + T_LableRect.Top + 2
    
    If lngWidth <> 0 Then
        '���ı���ʾ����ʾ��ȵ��м�����
        T_LableRect.Left = T_LableRect.Left + (lngWidth - ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY - 1) / 2
        T_LableRect.Right = ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY + T_LableRect.Left + 2
    End If
    
    If lngHeght <> 0 Then
        T_LableRect.Bottom = T_LableRect.Bottom + (lngHeght - ObjDraw.TextHeight(strInput) / sinTwipsPerPixelY)
    End If
    
End Sub


Public Sub DrawLine(ByVal lngDc As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0, _
    Optional ByVal blnEndRow As Boolean = False, Optional ByVal blnPrinter As Boolean = False)
    
    Dim X As Long
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim sngX As Single, sngY As Single
    On Error GoTo errHand
    '�����»��ʽ��л���
    
    If msngTwips = 0 Then msngTwips = 1
    sngX = 3 * msngTwips
    sngY = 4 * msngTwips

    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDc, lngPen)
    '��ͼ
    Call MoveToEx(lngDc, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDc, lngDX, lngDY)
    '���������»����¼�ͷ
    If blnEndRow Then
        If lngSY > lngDY Then '���ϼ�ͷ
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDc, X, lngDY + sngY, T_OldPoint)
                Call LineTo(lngDc, lngDX, lngDY)
            Next X
        Else '���¼�ͷ
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDc, X, lngDY - sngY, T_OldPoint)
                Call LineTo(lngDc, lngDX, lngDY)
            Next X
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(lngDc, lngOldPen)
    Call DeleteObject(lngPen)
    lngPen = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub DrawRect(ByVal lngDc As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0)
    
    Dim lngPen As Long, lngOldPen As Long
    On Error GoTo errHand
    '�����»��ʽ��л�һ������
    
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDc, lngPen)
    '��ͼ
    Call Rectangle(lngDc, lngSX, lngSY, lngDX, lngDY)
    '��ԭ���ʲ�����
    Call SelectObject(lngDc, lngOldPen)
    Call DeleteObject(lngPen)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub DrawRotateText(ByVal ObjDraw As Object, ByVal lngDc As Long, ByVal X As Single, _
                          ByVal Y As Single, _
                          ByVal strText As String, _
                          ByVal sngWidth As Single, _
                          Optional ByVal ForeColor As Long = 0, _
                          Optional ByVal sglScale As Single = 1)

    '��(X,Y)�����Text�ı�
    Dim objFont    As Object
    Dim lngFont    As Long
    Dim lngOldFont As Long
    Dim X1         As Long
    Dim blnPrinter As Boolean
    
    If strText = "" Then Exit Sub
    
    If TypeName(ObjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    '�����ı���ɫ
    Call SetTextColor(lngDc, ForeColor)
    
    '�����������
    If Asc(strText) < 0 And strText <> "��" Then
    
        Call GetTextRect(ObjDraw, X, Y, strText, sngWidth, False)
        Call DrawText(lngDc, strText, -1, T_LableRect, DT_CENTER)
        
    Else '��ת90���������
        '�ڴ�ӡ�Ƿ�ת������� objDraw.TextWidth ������ڴ�������֮ǰ�������޷���ת��
        Call GetTextRect(ObjDraw, X, Y, strText, sngWidth, False)
        X1 = X + (ObjDraw.TextWidth("��") / sinTwipsPerPixelX) + (T_LableRect.Left - X) - (IIf(blnPrinter = True, 2, 1) * msngTwips)
        Set objFont = New clsRotateFont
        Set objFont.LogFont = mobjSubFont
        
        objFont.sngTwpic = msngTwips
        objFont.Rotation = -90
        lngFont = objFont.Handle
        lngOldFont = SelectObject(lngDc, lngFont)
        Call TextOut(lngDc, X1, Y, strText, LenB(StrConv(strText, vbFromUnicode)))
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
    End If
End Sub

Public Sub ShowFlash(ByVal blnPrint As Boolean, Optional strInfo As String, Optional sngPer As Single, Optional frmParent As Object)
    '���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
    '����:strInfo=������ʾ��Ϣ
    '     sngPer=����
    '     blnPrint=true ��ʾ��ʾ�˴��壬false����ʾ
    If Not blnPrint Then Exit Sub
    Static blnShow As Boolean
    If sngPer > 1 Then sngPer = 1

    If strInfo = "" Then
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
            frmFlash.lblPer.Top = frmFlash.lbl.Top
            frmFlash.lbl.Caption = strInfo
            frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)

            If sngPer > 0 Then
                frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
            Else
                frmFlash.lblPer.Caption = ""
            End If

            If frmParent Is Nothing Then
                SetWindowPos frmFlash.Hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.Hwnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.Hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.Hwnd, 5
                End If
            End If

            frmFlash.Refresh
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Function GetPaperName(intSize As Integer) As String
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '���أ� ֽ������
    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
        intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

