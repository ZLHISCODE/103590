Attribute VB_Name = "mdlPrint"
Option Explicit

'�����ơ��߶ȡ���ȡ���С�߾�(��������)��ҳü�߾ࡢҳ�ű߾ࡢ��Ӧ��ӡֽ�����е�ֽ�����ೣ��
Public Const PageSize1 = "�ż� 8 1/2��11 Ӣ��                        ,15842,12242,350,350,350,350,350,350,1"
Public Const PageSize2 = "+A611 С���ż� 8 1/2��11 Ӣ��              ,15842,12242,350,350,350,350,350,350,2"
Public Const PageSize3 = "С�ͱ� 11��17 Ӣ��                         ,24477,15842,350,350,350,350,350,350,3"
Public Const PageSize4 = "������ 17��11 Ӣ��                         ,15842,24477,350,350,350,350,350,350,4"
Public Const PageSize5 = "�����ļ� 8 1/2��14 Ӣ��                    ,20163,12242,350,350,350,350,350,350,5"
Public Const PageSize6 = "������5 1/2��8 1/2 Ӣ��                    ,12242,7919,350,350,350,350,350,350,6"
Public Const PageSize7 = "�����ļ�7 1/2��10 1/2 Ӣ��                 ,15122,10438,350,350,350,350,350,350,7"
Public Const PageSize8 = "A3 297��420 ����                           ,23814,16840,350,350,350,350,350,350,8"
Public Const PageSize9 = "A4 210��297 ����                           ,16840,11907,350,350,350,350,350,350,9"
Public Const PageSize10 = "A4С�� 210��297 ����                      ,16840,11907,350,350,350,350,350,350,9"
Public Const PageSize11 = "A5 148��210 ����                          ,11907,8392,350,350,350,350,350,350,11"
Public Const PageSize12 = "B4 250��354 ����                          ,20067,14171,350,350,350,350,350,350,12"
Public Const PageSize13 = "B5 182��257 ����                          ,14572,10319,350,350,350,350,350,350,13"
Public Const PageSize14 = "�Կ��� 8 1/2��13 Ӣ��                     ,18722,12242,350,350,350,350,350,350,14"
Public Const PageSize15 = "�Ŀ��� 215��275 ����                      ,15589,12187,350,350,350,350,350,350,15"
Public Const PageSize16 = "10��14 Ӣ��                               ,20163,14398,350,350,350,350,350,350,16"
Public Const PageSize17 = "11��17 Ӣ��                               ,24477,15842,350,350,350,350,350,350,17"
Public Const PageSize18 = "����8 1/2��11 Ӣ��                        ,15842,12242,350,350,350,350,350,350,18"
Public Const PageSize19 = "#9 �ŷ� 3 7/8��8 7/8 Ӣ��                 ,5579,12780,350,350,350,350,350,350,19"
Public Const PageSize20 = "#10 �ŷ� 4 1/8��9 1/2 Ӣ��                ,5936,13682,350,350,350,350,350,350,20"
Public Const PageSize21 = "#11 �ŷ� 4 1/2��10 3/8 Ӣ��               ,14938,6479,350,350,350,350,350,350,21"
Public Const PageSize22 = "#12 �ŷ� 4 1/2��11 Ӣ��                   ,15842,6479,350,350,350,350,350,350,22"
Public Const PageSize23 = "#14 �ŷ� 5��11 1/2 Ӣ��                   ,16558,7199,350,350,350,350,350,350,23"
Public Const PageSize24 = "C �ߴ繤����                              ,16558,7199,350,350,350,350,350,350,24"
Public Const PageSize25 = "D �ߴ繤����                              ,16558,7199,350,350,350,350,350,350,25"
Public Const PageSize26 = "E �ߴ繤����                              ,16558,7199,350,350,350,350,350,350,26"
Public Const PageSize27 = "DL ���ŷ� 110��220 ����                   ,6237,12474,350,350,350,350,350,350,27"
Public Const PageSize28 = "C5 ���ŷ� 162��229 ����                   ,9185,12984,350,350,350,350,350,350,28"
Public Const PageSize29 = "C3 ���ŷ� 324��458 ����                   ,25969,18371,350,350,350,350,350,350,29"
Public Const PageSize30 = "C4 ���ŷ� 229��324 ����                   ,18371,12981,350,350,350,350,350,350,30"
Public Const PageSize31 = "C6 ���ŷ� 114��162 ����                   ,9183,6462,350,350,350,350,350,350,31"
Public Const PageSize32 = "C65 ���ŷ�114��229 ����                   ,12981,6462,350,350,350,350,350,350,32"
Public Const PageSize33 = "B4 ���ŷ� 250��353 ����                   ,20010,14171,350,350,350,350,350,350,33"
Public Const PageSize34 = "B5 ���ŷ�176��250 ����                    ,9979,14350,350,350,350,350,350,350,34"
Public Const PageSize35 = "B6 ���ŷ� 176��125 ����                   ,7086,9976,350,350,350,350,350,350,35"
Public Const PageSize36 = "�ŷ� 110��230 ����                        ,13037,6237,350,350,350,350,350,350,36"
Public Const PageSize37 = "�ŷ���� 3 7/8��7 1/2 Ӣ��                ,5579,10801,350,350,350,350,350,350,37"
Public Const PageSize38 = "�ŷ� 3 5/8��6 1/2 Ӣ��                    ,9359,5219,350,350,350,350,350,350,38"
Public Const PageSize39 = "U.S. ��׼��д�� 14 7/8��11 Ӣ��           ,15842,21421,350,350,350,350,350,350,39"
Public Const PageSize40 = "�¹���׼��д�� 8 1/2��12 Ӣ��             ,17282,12242,350,350,350,350,350,350,40"
Public Const PageSize41 = "�¹����ɸ�д�� 8 1/2��13 Ӣ��             ,18722,12242,350,350,350,350,350,350,41"
Public Const PageSize42 = "�Զ���ֽ��                                ,22680,16443,350,350,350,350,350,350,256"



'/* Device Parameters for GetDeviceCaps() */
'#define DRIVERVERSION 0     /* Device driver version                    */
'#define TECHNOLOGY    2     /* Device classification                    */
'#define HORZSIZE      4     /* Horizontal size in millimeters           */
'#define VERTSIZE      6     /* Vertical size in millimeters             */
'#define HORZRES       8     /* Horizontal width in pixels               */
'#define VERTRES       10    /* Vertical height in pixels                */
'#define BITSPIXEL     12    /* Number of bits per pixel                 */
'#define PLANES        14    /* Number of planes                         */
'#define NUMBRUSHES    16    /* Number of brushes the device has         */
'#define NUMPENS       18    /* Number of pens the device has            */
'#define NUMMARKERS    20    /* Number of markers the device has         */
'#define NUMFONTS      22    /* Number of fonts the device has           */
'#define NUMCOLORS     24    /* Number of colors the device supports     */
'#define PDEVICESIZE   26    /* Size required for device descriptor      */
'#define CURVECAPS     28    /* Curve capabilities                       */
'#define LINECAPS      30    /* Line capabilities                        */
'#define POLYGONALCAPS 32    /* Polygonal capabilities                   */
'#define TEXTCAPS      34    /* Text capabilities                        */
'#define CLIPCAPS      36    /* Clipping capabilities                    */
'#define RASTERCAPS    38    /* Bitblt capabilities                      */
'#define ASPECTX       40    /* Length of the X leg                      */
'#define ASPECTY       42    /* Length of the Y leg                      */
'#define ASPECTXY      44    /* Length of the hypotenuse                 */

'GetDeviceCaps()�����Ĳ�������
Public Const DRIVERVERSION = 0      '�豸��������汾
Public Const TECHNOLOGY = 2         '�豸����
Public Const HORZSIZE = 4           '������Ļ��ȣ���λ�����ס�
Public Const VERTSIZE = 6           '������Ļ�߶ȣ���λ�����ס�
Public Const HORZRES = 8            '��Ļ��ȣ���λ�����أ�pixels����
Public Const VERTRES = 10           '��Ļ�߶ȣ���λ������դ���С�
Public Const BITSPIXEL = 12         'ÿ�����ص��������ɫλ����
Public Const PLANES = 14            '��ɫƽ������
Public Const NUMBRUSHES = 16        '�豸��ػ�ˢ(BRUSH)��Ŀ��
Public Const NUMPENS = 18           '�豸��ػ���(PEN)��Ŀ��
Public Const NUMMARKERS = 20        '�豸��ر����Ŀ��
Public Const NUMFONTS = 22          '�豸���������Ŀ��
Public Const NUMCOLORS = 24         '�豸��ɫ����������������豸����ɫ���С��ÿ����8λʱ���á����ڸ�ɫ��ʱ������-1��
Public Const PDEVICESIZE = 26       '������
Public Const CURVECAPS = 28         '�豸���������ԡ�
Public Const LINECAPS = 30          '�豸���������ԡ�
Public Const POLYGONALCAPS = 32     '�豸�Ķ�������ԡ�
Public Const TEXTCAPS = 34          '�豸���ı����ԡ�
Public Const CLIPCAPS = 36          '�豸�������ܱ�־������豸���Լ���Ϊ���Σ�����1������Ϊ0��
Public Const RASTERCAPS = 38        '�豸�Ĺ�դ���ԡ�
Public Const ASPECTX = 40           '��������ʱ��������ؿ�ȡ�
Public Const ASPECTY = 42           '��������ʱ��������ظ߶ȡ�
Public Const ASPECTXY = 44          '��������ʱ����ԶԽ������ؿ�ȡ�

'#if(WINVER >= 0x0500)
'#define SHADEBLENDCAPS 45   /* Shading and blending caps                */
'#endif /* WINVER >= 0x0500 */
Public Const SHADEBLENDCAPS = 45    '�豸����Ӱ��������ԡ�

'#define LOGPIXELSX    88    /* Logical pixels/inch in X                 */
'#define LOGPIXELSY    90    /* Logical pixels/inch in Y                 */
Public Const LOGPIXELSX = 88        '����Ļ��ȵ�ÿ���߼�Ӣ�������ֵ���ڶ���ʾ��ϵͳ�У�������ʾ�������ֵ����ͬ��
Public Const LOGPIXELSY = 90        '����Ļ�߶ȵ�ÿ���߼�Ӣ�������ֵ���ڶ���ʾ��ϵͳ�У�������ʾ�������ֵ����ͬ��

'#define SIZEPALETTE  104    /* Number of entries in physical palette    */
'#define NUMRESERVED  106    /* Number of reserved entries in palette    */
'#define COLORRES     108    /* Actual color resolution                  */

'����3������ֵֻ�����豸������RASTERCAPS����RC_PALETTEλ�����ڼ���16λWindowsʱ�ſ��á�
Public Const SIZEPALETTE = 104      '�豸��ɫ������������
Public Const NUMRESERVED = 106      'ϵͳ��ɫ��ı������������
Public Const COLORRES = 108         '�豸��ʵ����ɫ�ֱ��ʣ���λ��BPP��λ/���أ���

'// Printing related DeviceCaps. These replace the appropriate Escapes
'
'#define PHYSICALWIDTH   110 /* Physical Width in device units           */
'#define PHYSICALHEIGHT  111 /* Physical Height in device units          */
'#define PHYSICALOFFSETX 112 /* Physical Printable Area x margin         */
'#define PHYSICALOFFSETY 113 /* Physical Printable Area y margin         */
'#define SCALINGFACTORX  114 /* Scaling factor x                         */
'#define SCALINGFACTORY  115 /* Scaling factor y                         */

'��ӡ��س�������Щֵ���滻��Ӧ��ת�Ʒ�
Public Const PHYSICALWIDTH = 110    '���ڴ�ӡ�豸���ԣ���ʾ����ҳ�������豸��λ��ע������ҳ���Ǵ���ҳ��Ŀɴ�ӡ���򣬲���С������
Public Const PHYSICALHEIGHT = 111   '���ڴ�ӡ�豸���ԣ���ʾ����ҳ�ߣ������豸��λ��
Public Const PHYSICALOFFSETX = 112  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ�����Ե���ɴ�ӡ��������Ե�ľ��룬�����豸��λ��
Public Const PHYSICALOFFSETY = 113  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ���ϱ�Ե���ɴ�ӡ������ϱ�Ե�ľ��룬�����豸��λ��
Public Const SCALINGFACTORX = 114   '��ӡ����X�������ű�����
Public Const SCALINGFACTORY = 115   '��ӡ����Y�������ű�����

'// Display driver specific
'
'#define VREFRESH        116  /* Current vertical refresh rate of the    */
'                             /* display device (for displays only) in Hz*/
'#define DESKTOPVERTRES  117  /* Horizontal width of entire desktop in   */
'                             /* pixels                                  */
'#define DESKTOPHORZRES  118  /* Vertical height of entire desktop in    */
'                             /* pixels                                  */
'#define BLTALIGNMENT    119  /* Preferred blt alignment                 */
'
'#ifndef NOGDICAPMASKS

'��ʾ�豸��س���
Public Const VREFRESH = 116         '����ʾ�豸���ԣ���ʾ��ǰ�Ĵ�ֱˢ���ʣ���λ��Hz��
Public Const DESKTOPVERTRES = 117   '��������Ŀ�ȣ���λ��Pixels
Public Const DESKTOPHORZRES = 118   '��������ĸ߶ȣ���λ��Pixels
Public Const BLTALIGNMENT = 119     'Ĭ�� blt ���뷽ʽ

'/* Device Capability Masks: */
'�豸��������

'/* Device Technologies */
'#define DT_PLOTTER          0   /* Vector plotter                   */
'#define DT_RASDISPLAY       1   /* Raster display                   */
'#define DT_RASPRINTER       2   /* Raster printer                   */
'#define DT_RASCAMERA        3   /* Raster camera                    */
'#define DT_CHARSTREAM       4   /* Character-stream, PLP            */
'#define DT_METAFILE         5   /* Metafile, VDM                    */
'#define DT_DISPFILE         6   /* Display-file                     */
'�豸����
Public Const DT_PLOTTER = 0         'ʸ����ͼ��
Public Const DT_RASDISPLAY = 1      '��դ��ʾ��
Public Const DT_RASPRINTER = 2      '��դ��ӡ��
Public Const DT_RASCAMERA = 3       '��դ�����
Public Const DT_CHARSTREAM = 4      '�ַ���
Public Const DT_METAFILE = 5        'ͼԪ�ļ�
Public Const DT_DISPFILE = 6        '��ʾ�ļ�

'/* Curve Capabilities */
'�豸���������ԡ�

'#define CC_NONE             0   /* Curves not supported             */
'#define CC_CIRCLES          1   /* Can do circles                   */
'#define CC_PIE              2   /* Can do pie wedges                */
'#define CC_CHORD            4   /* Can do chord arcs                */
'#define CC_ELLIPSES         8   /* Can do ellipese                  */
'#define CC_WIDE             16  /* Can do wide lines                */
'#define CC_STYLED           32  /* Can do styled lines              */
'#define CC_WIDESTYLED       64  /* Can do wide styled lines         */
'#define CC_INTERIORS        128 /* Can do interiors                 */
'#define CC_ROUNDRECT        256 /*                                  */
Public Const CC_NONE = 0            '�豸��֧�����ߡ�
Public Const CC_CIRCLES = 1         '�豸���Ի����һ���
Public Const CC_PIE = 2             '�豸���Ի���Բ��
Public Const CC_CHORD = 4           '�豸���Ի�����Բ��
Public Const CC_ELLIPSES = 8        '�豸���Ի�����Բ��
Public Const CC_WIDE = 16           '�豸���Ի��ƿ�߿�
Public Const CC_STYLED = 32         '�豸���Ի�����ʽ�߿�
Public Const CC_WIDESTYLED = 64     '�豸���Ի��ƿ���ʽ�߿�
Public Const CC_INTERIORS = 128     '�豸���Ի����ڲ�����
Public Const CC_ROUNDRECT = 256     '�豸���Ի���Բ�Ǿ��Ρ�

'/* Line Capabilities */
'�豸���������ԡ�

'#define LC_NONE             0   /* Lines not supported              */
'#define LC_POLYLINE         2   /* Can do polylines                 */
'#define LC_MARKER           4   /* Can do markers                   */
'#define LC_POLYMARKER       8   /* Can do polymarkers               */
'#define LC_WIDE             16  /* Can do wide lines                */
'#define LC_STYLED           32  /* Can do styled lines              */
'#define LC_WIDESTYLED       64  /* Can do wide styled lines         */
'#define LC_INTERIORS        128 /* Can do interiors                 */
Public Const LC_NONE = 0            '�豸��֧��������
Public Const LC_POLYLINE = 2        '�豸���Ի������ߡ�
Public Const LC_MARKER = 4          '�豸���Ի���һ����ǡ�
Public Const LC_POLYMARKER = 8      '�豸���Ի��ƶ����ǡ�
Public Const LC_WIDE = 16           '�豸���Ի��ƿ�������
Public Const LC_STYLED = 32         '�豸���Ի�����ʽ������
Public Const LC_WIDESTYLED = 64     '�豸���Ի��ƿ���ʽ������
Public Const LC_INTERIORS = 128     '�豸���Ի����ڲ�����

'/* Polygonal Capabilities */
'�豸�Ķ�������ԡ�
'#define PC_NONE             0   /* Polygonals not supported         */
'#define PC_POLYGON          1   /* Can do polygons                  */
'#define PC_RECTANGLE        2   /* Can do rectangles                */
'#define PC_WINDPOLYGON      4   /* Can do winding polygons          */
'#define PC_TRAPEZOID        4   /* Can do trapezoids                */
'#define PC_SCANLINE         8   /* Can do scanlines                 */
'#define PC_WIDE             16  /* Can do wide borders              */
'#define PC_STYLED           32  /* Can do styled borders            */
'#define PC_WIDESTYLED       64  /* Can do wide styled borders       */
'#define PC_INTERIORS        128 /* Can do interiors                 */
'#define PC_POLYPOLYGON      256 /* Can do polypolygons              */
'#define PC_PATHS            512 /* Can do paths                     */
Public Const PC_NONE = 0            '�豸��֧�ֶ���Ρ�
Public Const PC_POLYGON = 1         '�豸���Ի��ƽ������Ķ���Ρ�
Public Const PC_RECTANGLE = 2       '�豸���Ի��ƾ��Ρ�
Public Const PC_WINDPOLYGON = 4     '�豸���Ի����������Ķ���Ρ�
Public Const PC_TRAPEZOID = 4       '�豸���Ի��Ʋ������ı��Ρ�
Public Const PC_SCANLINE = 8        '�豸���Ի����豸���Ի��Ƶ�ɨ���ߡ�
Public Const PC_WIDE = 16           '�豸���Ի��ƿ�߿�
Public Const PC_STYLED = 32         '�豸���Ի�����ʽ�߿�
Public Const PC_WIDESTYLED = 64     '�豸���Ի��ƿ���ʽ�߿�
Public Const PC_INTERIORS = 128     '�豸���Ի����ڲ�����
Public Const PC_POLYPOLYGON = 256   '�豸���Ի��ƶ������Ρ�
Public Const PC_PATHS = 512         '�豸���Ի���·����

'/* Clipping Capabilities */
'�ü�����
'#define CP_NONE             0   /* No clipping of output            */
'#define CP_RECTANGLE        1   /* Output clipped to rects          */
'#define CP_REGION           2   /* obsolete                         */
Public Const CP_NONE = 0            '������ü�
Public Const CP_RECTANGLE = 1       '����ü�������
Public Const CP_REGION = 2          '����

'/* Text Capabilities */
'�ı�����
'#define TC_OP_CHARACTER     0x00000001  /* Can do OutputPrecision   CHARACTER      */
'#define TC_OP_STROKE        0x00000002  /* Can do OutputPrecision   STROKE         */
'#define TC_CP_STROKE        0x00000004  /* Can do ClipPrecision     STROKE         */
'#define TC_CR_90            0x00000008  /* Can do CharRotAbility    90             */
'#define TC_CR_ANY           0x00000010  /* Can do CharRotAbility    ANY            */
'#define TC_SF_X_YINDEP      0x00000020  /* Can do ScaleFreedom      X_YINDEPENDENT */
'#define TC_SA_DOUBLE        0x00000040  /* Can do ScaleAbility      DOUBLE         */
'#define TC_SA_INTEGER       0x00000080  /* Can do ScaleAbility      INTEGER        */
'#define TC_SA_CONTIN        0x00000100  /* Can do ScaleAbility      CONTINUOUS     */
'#define TC_EA_DOUBLE        0x00000200  /* Can do EmboldenAbility   DOUBLE         */
'#define TC_IA_ABLE          0x00000400  /* Can do ItalisizeAbility  ABLE           */
'#define TC_UA_ABLE          0x00000800  /* Can do UnderlineAbility  ABLE           */
'#define TC_SO_ABLE          0x00001000  /* Can do StrikeOutAbility  ABLE           */
'#define TC_RA_ABLE          0x00002000  /* Can do RasterFontAble    ABLE           */
'#define TC_VA_ABLE          0x00004000  /* Can do VectorFontAble    ABLE           */
'#define TC_RESERVED         0x00008000
'#define TC_SCROLLBLT        0x00010000  /* Don't do text scroll with blt           */
Public Const TC_OP_CHARACTER = &H1  '�豸�����ַ�������ȡ�
Public Const TC_OP_STROKE = &H2     '�豸����ʻ�������ȡ�
Public Const TC_CP_STROKE = &H4     '�豸����ʻ��ü����ȡ�
Public Const TC_CR_90 = &H8         '�豸����90���ַ���ת��
Public Const TC_CR_ANY = &H10       '�豸���������ַ���ת��
Public Const TC_SF_X_YINDEP = &H20  '�豸������X���Y��������š�
Public Const TC_SA_DOUBLE = &H40    '�豸֧��2���ַ����š�
Public Const TC_SA_INTEGER = &H80   '�豸ֻ�ܲ����ַ������������š�
Public Const TC_SA_CONTIN = &H100   '�豸���Բ����ַ������ⱶ�����š�
Public Const TC_EA_DOUBLE = &H200   '�豸���Ի���˫����ֵ���ַ���
Public Const TC_IA_ABLE = &H400     '�豸֧��б�塣
Public Const TC_UA_ABLE = &H800     '�豸֧���»��ߡ�
Public Const TC_SO_ABLE = &H1000    '�豸֧��ɾ���ߡ�
Public Const TC_RA_ABLE = &H2000    '�豸֧�ֹ�դ���塣
Public Const TC_VA_ABLE = &H4000    '�豸֧��ʸ�����塣
Public Const TC_RESERVED = &H8000   '����������Ϊ0��
Public Const TC_SCROLLBLT = &H10000 '�ı����������

'#endif /* NOGDICAPMASKS */

'/* Raster Capabilities */
'��դ����
'#define RC_NONE
'#define RC_BITBLT           1       /* Can do standard BLT.             */
'#define RC_BANDING          2       /* Device requires banding support  */
'#define RC_SCALING          4       /* Device requires scaling support  */
'#define RC_BITMAP64         8       /* Device can support >64K bitmap   */
'#define RC_GDI20_OUTPUT     0x0010      /* has 2.0 output calls         */
'#define RC_GDI20_STATE      0x0020
'#define RC_SAVEBITMAP       0x0040
'#define RC_DI_BITMAP        0x0080      /* supports DIB to memory       */
'#define RC_PALETTE          0x0100      /* supports a palette           */
'#define RC_DIBTODEV         0x0200      /* supports DIBitsToDevice      */
'#define RC_BIGFONT          0x0400      /* supports >64K fonts          */
'#define RC_STRETCHBLT       0x0800      /* supports StretchBlt          */
'#define RC_FLOODFILL        0x1000      /* supports FloodFill           */
'#define RC_STRETCHDIB       0x2000      /* supports StretchDIBits       */
'#define RC_OP_DX_OUTPUT     0x4000
'#define RC_DEVBITS          0x8000
Public Const RC_NONE = 0                '
Public Const RC_BITBLT = 1              '���Դ���λͼ��
Public Const RC_BANDING = 2             '��Ҫ������(Banding)֧�֡�
Public Const RC_SCALING = 4             '֧�����š�
Public Const RC_BITMAP64 = 8            '����֧�ִ���64KB��λͼ��
Public Const RC_GDI20_OUTPUT = &H10     '
Public Const RC_GDI20_STATE = &H20      '
Public Const RC_SAVEBITMAP = &H40       '
Public Const RC_DI_BITMAP = &H80        '֧��SetDIBits��GetDIBits������
Public Const RC_PALETTE = &H100         'ָ��һ�����ڵ�ɫ����豸��
Public Const RC_DIBTODEV = &H200        '֧��SetDIBitsToDevice������
Public Const RC_BIGFONT = &H400         '֧�ִ���64K�����塣
Public Const RC_STRETCHBLT = &H800      '֧��StretchBlt������
Public Const RC_FLOODFILL = &H1000      '����ִ��flood fills��������
Public Const RC_STRETCHDIB = &H2000     '֧��StretchDIBits������
Public Const RC_OP_DX_OUTPUT = &H4000
Public Const RC_DEVBITS = &H8000

'#if(WINVER >= 0x0500)

'/* Shading and blending caps                */
'�豸����Ӱ��������ԡ�
'#define SB_NONE             0x00000000
'#define SB_CONST_ALPHA      0x00000001
'#define SB_PIXEL_ALPHA      0x00000002
'#define SB_PREMULT_ALPHA    0x00000004
Public Const SB_NONE = &H0              '
Public Const SB_CONST_ALPHA = &H1       '
Public Const SB_PIXEL_ALPHA = &H2       '
Public Const SB_PREMULT_ALPHA = &H4     '

'#define SB_GRAD_RECT        0x00000010
'#define SB_GRAD_TRI         0x00000020
Public Const SB_GRAD_RECT = &H10              '
Public Const SB_GRAD_TRI = &H20              '

'#endif /* WINVER >= 0x0500 */


Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'WinNT�Զ���ֽ�ſ���================================================================
'ע����dmFields��Long��,as Long��β����&��
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Public Const IDOK = 1
Public Const IDCANCEL = 2
'Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
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
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As Any) As Long
Public Const conRatemmToTwip As Single = 56.6857142857143      '������羵ı���

Public Const WM_MOUSEWHEEL = &H20A

'######################################################################################
'   �ͷ��ڴ�
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "Psapi" (ByVal hProcess As Long) As Long

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


'######################################################################################
'## �ж�ϵͳ�Ƿ���NT
'######################################################################################

Public Function IsWindowsNT() As Boolean
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Private Function GetWinPlatform() As Long
'��    �ܣ����ص�ǰ��ϵͳ�汾����
'��    ������
'��    �أ�
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

'######################################################################################
'## ���ô�ӡ�����Զ���ֽ�ųߴ�
'######################################################################################

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte
    Dim lngSize As Long
    
    Dim lngPrtDC As Long
    Dim lngHandle As Long
    Dim strPrtName As String
    
    lngPrtDC = Printer.hdc
    strPrtName = Printer.DeviceName
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE:fMode=0
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetCustomPager(ByRef lngHwnd As Long, ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'���ܣ��������Զ���ֽ��
'�����������Ϊ��λ
    If IsWindowsNT Then
        '��Ȼ����ʹ�����Ч�����ܸı�PaperSize������ֵ
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustomPager = SetNTPrinterPaper(lngHwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98ϵ�л�����ͨ����������
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function



Public Function GetPaperName(ByVal intSize As Integer) As String
'���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
'������ lngW,lngH=�Զ���ֽ�ŵĿ��(Twip)
'���أ� ֽ������
    If intSize >= 1 And intSize <= 42 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41, intSize = 42, PageSize42)
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function
Public Function ExistsPrinter() As Boolean
    Dim lngHDc As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDc = Printer.hdc
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
End Function
