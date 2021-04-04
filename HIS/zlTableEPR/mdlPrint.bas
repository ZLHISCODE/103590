Attribute VB_Name = "mdlPrint"
Option Explicit

'按名称、高度、宽度、最小边距(上下左右)、页眉边距、页脚边距、对应打印纸张排列的纸张种类常量
Public Const PageSize1 = "信笺 8 1/2×11 英寸                        ,15842,12242,350,350,350,350,350,350,1"
Public Const PageSize2 = "+A611 小型信笺 8 1/2×11 英寸              ,15842,12242,350,350,350,350,350,350,2"
Public Const PageSize3 = "小型报 11×17 英寸                         ,24477,15842,350,350,350,350,350,350,3"
Public Const PageSize4 = "分类帐 17×11 英寸                         ,15842,24477,350,350,350,350,350,350,4"
Public Const PageSize5 = "法律文件 8 1/2×14 英寸                    ,20163,12242,350,350,350,350,350,350,5"
Public Const PageSize6 = "声明书5 1/2×8 1/2 英寸                    ,12242,7919,350,350,350,350,350,350,6"
Public Const PageSize7 = "行政文件7 1/2×10 1/2 英寸                 ,15122,10438,350,350,350,350,350,350,7"
Public Const PageSize8 = "A3 297×420 毫米                           ,23814,16840,350,350,350,350,350,350,8"
Public Const PageSize9 = "A4 210×297 毫米                           ,16840,11907,350,350,350,350,350,350,9"
Public Const PageSize10 = "A4小号 210×297 毫米                      ,16840,11907,350,350,350,350,350,350,9"
Public Const PageSize11 = "A5 148×210 毫米                          ,11907,8392,350,350,350,350,350,350,11"
Public Const PageSize12 = "B4 250×354 毫米                          ,20067,14171,350,350,350,350,350,350,12"
Public Const PageSize13 = "B5 182×257 毫米                          ,14572,10319,350,350,350,350,350,350,13"
Public Const PageSize14 = "对开本 8 1/2×13 英寸                     ,18722,12242,350,350,350,350,350,350,14"
Public Const PageSize15 = "四开本 215×275 毫米                      ,15589,12187,350,350,350,350,350,350,15"
Public Const PageSize16 = "10×14 英寸                               ,20163,14398,350,350,350,350,350,350,16"
Public Const PageSize17 = "11×17 英寸                               ,24477,15842,350,350,350,350,350,350,17"
Public Const PageSize18 = "便条8 1/2×11 英寸                        ,15842,12242,350,350,350,350,350,350,18"
Public Const PageSize19 = "#9 信封 3 7/8×8 7/8 英寸                 ,5579,12780,350,350,350,350,350,350,19"
Public Const PageSize20 = "#10 信封 4 1/8×9 1/2 英寸                ,5936,13682,350,350,350,350,350,350,20"
Public Const PageSize21 = "#11 信封 4 1/2×10 3/8 英寸               ,14938,6479,350,350,350,350,350,350,21"
Public Const PageSize22 = "#12 信封 4 1/2×11 英寸                   ,15842,6479,350,350,350,350,350,350,22"
Public Const PageSize23 = "#14 信封 5×11 1/2 英寸                   ,16558,7199,350,350,350,350,350,350,23"
Public Const PageSize24 = "C 尺寸工作单                              ,16558,7199,350,350,350,350,350,350,24"
Public Const PageSize25 = "D 尺寸工作单                              ,16558,7199,350,350,350,350,350,350,25"
Public Const PageSize26 = "E 尺寸工作单                              ,16558,7199,350,350,350,350,350,350,26"
Public Const PageSize27 = "DL 型信封 110×220 毫米                   ,6237,12474,350,350,350,350,350,350,27"
Public Const PageSize28 = "C5 型信封 162×229 毫米                   ,9185,12984,350,350,350,350,350,350,28"
Public Const PageSize29 = "C3 型信封 324×458 毫米                   ,25969,18371,350,350,350,350,350,350,29"
Public Const PageSize30 = "C4 型信封 229×324 毫米                   ,18371,12981,350,350,350,350,350,350,30"
Public Const PageSize31 = "C6 型信封 114×162 毫米                   ,9183,6462,350,350,350,350,350,350,31"
Public Const PageSize32 = "C65 型信封114×229 毫米                   ,12981,6462,350,350,350,350,350,350,32"
Public Const PageSize33 = "B4 型信封 250×353 毫米                   ,20010,14171,350,350,350,350,350,350,33"
Public Const PageSize34 = "B5 型信封176×250 毫米                    ,9979,14350,350,350,350,350,350,350,34"
Public Const PageSize35 = "B6 型信封 176×125 毫米                   ,7086,9976,350,350,350,350,350,350,35"
Public Const PageSize36 = "信封 110×230 毫米                        ,13037,6237,350,350,350,350,350,350,36"
Public Const PageSize37 = "信封大王 3 7/8×7 1/2 英寸                ,5579,10801,350,350,350,350,350,350,37"
Public Const PageSize38 = "信封 3 5/8×6 1/2 英寸                    ,9359,5219,350,350,350,350,350,350,38"
Public Const PageSize39 = "U.S. 标准复写簿 14 7/8×11 英寸           ,15842,21421,350,350,350,350,350,350,39"
Public Const PageSize40 = "德国标准复写簿 8 1/2×12 英寸             ,17282,12242,350,350,350,350,350,350,40"
Public Const PageSize41 = "德国法律复写簿 8 1/2×13 英寸             ,18722,12242,350,350,350,350,350,350,41"
Public Const PageSize42 = "自定义纸张                                ,22680,16443,350,350,350,350,350,350,256"



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

'GetDeviceCaps()函数的参数常数
Public Const DRIVERVERSION = 0      '设备驱动程序版本
Public Const TECHNOLOGY = 2         '设备工艺
Public Const HORZSIZE = 4           '物理屏幕宽度，单位：毫米。
Public Const VERTSIZE = 6           '物理屏幕高度，单位：毫米。
Public Const HORZRES = 8            '屏幕宽度，单位：象素（pixels）。
Public Const VERTRES = 10           '屏幕高度，单位：（光栅）行。
Public Const BITSPIXEL = 12         '每个象素点的相邻颜色位数。
Public Const PLANES = 14            '颜色平面数。
Public Const NUMBRUSHES = 16        '设备相关画刷(BRUSH)数目。
Public Const NUMPENS = 18           '设备相关画笔(PEN)数目。
Public Const NUMMARKERS = 20        '设备相关标记数目。
Public Const NUMFONTS = 22          '设备相关字体数目。
Public Const NUMCOLORS = 24         '设备颜色表的入口总数，如果设备的颜色深度小于每象素8位时可用。大于该色深时，返回-1。
Public Const PDEVICESIZE = 26       '保留。
Public Const CURVECAPS = 28         '设备的曲线特性。
Public Const LINECAPS = 30          '设备的线条特性。
Public Const POLYGONALCAPS = 32     '设备的多边形特性。
Public Const TEXTCAPS = 34          '设备的文本特性。
Public Const CLIPCAPS = 36          '设备剪切性能标志，如果设备可以剪切为矩形，返回1；否则为0。
Public Const RASTERCAPS = 38        '设备的光栅特性。
Public Const ASPECTX = 40           '绘制线条时的相对象素宽度。
Public Const ASPECTY = 42           '绘制线条时的相对象素高度。
Public Const ASPECTXY = 44          '绘制线条时的相对对角线象素宽度。

'#if(WINVER >= 0x0500)
'#define SHADEBLENDCAPS 45   /* Shading and blending caps                */
'#endif /* WINVER >= 0x0500 */
Public Const SHADEBLENDCAPS = 45    '设备的阴影及混合特性。

'#define LOGPIXELSX    88    /* Logical pixels/inch in X                 */
'#define LOGPIXELSY    90    /* Logical pixels/inch in Y                 */
Public Const LOGPIXELSX = 88        '沿屏幕宽度的每个逻辑英寸的象素值。在多显示器系统中，所有显示器的这个值均相同。
Public Const LOGPIXELSY = 90        '沿屏幕高度的每个逻辑英寸的象素值。在多显示器系统中，所有显示器的这个值均相同。

'#define SIZEPALETTE  104    /* Number of entries in physical palette    */
'#define NUMRESERVED  106    /* Number of reserved entries in palette    */
'#define COLORRES     108    /* Actual color resolution                  */

'下面3个索引值只能在设备驱动在RASTERCAPS等于RC_PALETTE位并且在兼容16位Windows时才可用。
Public Const SIZEPALETTE = 104      '设备调色板的入口总数。
Public Const NUMRESERVED = 106      '系统调色板的保留入口总数。
Public Const COLORRES = 108         '设备的实际颜色分辨率，单位：BPP（位/象素）。

'// Printing related DeviceCaps. These replace the appropriate Escapes
'
'#define PHYSICALWIDTH   110 /* Physical Width in device units           */
'#define PHYSICALHEIGHT  111 /* Physical Height in device units          */
'#define PHYSICALOFFSETX 112 /* Physical Printable Area x margin         */
'#define PHYSICALOFFSETY 113 /* Physical Printable Area y margin         */
'#define SCALINGFACTORX  114 /* Scaling factor x                         */
'#define SCALINGFACTORY  115 /* Scaling factor y                         */

'打印相关常量，这些值将替换对应的转移符
Public Const PHYSICALWIDTH = 110    '对于打印设备而言，表示物理页宽，采用设备单位。注：物理页总是大于页面的可打印区域，不会小于它。
Public Const PHYSICALHEIGHT = 111   '对于打印设备而言，表示物理页高，采用设备单位。
Public Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Public Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Public Const SCALINGFACTORX = 114   '打印机的X－轴缩放比例。
Public Const SCALINGFACTORY = 115   '打印机的Y－轴缩放比例。

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

'显示设备相关常量
Public Const VREFRESH = 116         '对显示设备而言，表示当前的垂直刷新率，单位：Hz。
Public Const DESKTOPVERTRES = 117   '整个桌面的宽度，单位：Pixels
Public Const DESKTOPHORZRES = 118   '整个桌面的高度，单位：Pixels
Public Const BLTALIGNMENT = 119     '默认 blt 对齐方式

'/* Device Capability Masks: */
'设备性能掩码

'/* Device Technologies */
'#define DT_PLOTTER          0   /* Vector plotter                   */
'#define DT_RASDISPLAY       1   /* Raster display                   */
'#define DT_RASPRINTER       2   /* Raster printer                   */
'#define DT_RASCAMERA        3   /* Raster camera                    */
'#define DT_CHARSTREAM       4   /* Character-stream, PLP            */
'#define DT_METAFILE         5   /* Metafile, VDM                    */
'#define DT_DISPFILE         6   /* Display-file                     */
'设备工艺
Public Const DT_PLOTTER = 0         '矢量绘图仪
Public Const DT_RASDISPLAY = 1      '光栅显示器
Public Const DT_RASPRINTER = 2      '光栅打印机
Public Const DT_RASCAMERA = 3       '光栅照相机
Public Const DT_CHARSTREAM = 4      '字符流
Public Const DT_METAFILE = 5        '图元文件
Public Const DT_DISPFILE = 6        '显示文件

'/* Curve Capabilities */
'设备的曲线特性。

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
Public Const CC_NONE = 0            '设备不支持曲线。
Public Const CC_CIRCLES = 1         '设备可以绘制弦弧。
Public Const CC_PIE = 2             '设备可以绘制圆。
Public Const CC_CHORD = 4           '设备可以绘制椭圆。
Public Const CC_ELLIPSES = 8        '设备可以绘制椭圆。
Public Const CC_WIDE = 16           '设备可以绘制宽边框。
Public Const CC_STYLED = 32         '设备可以绘制样式边框。
Public Const CC_WIDESTYLED = 64     '设备可以绘制宽样式边框。
Public Const CC_INTERIORS = 128     '设备可以绘制内部区域。
Public Const CC_ROUNDRECT = 256     '设备可以绘制圆角矩形。

'/* Line Capabilities */
'设备的线条特性。

'#define LC_NONE             0   /* Lines not supported              */
'#define LC_POLYLINE         2   /* Can do polylines                 */
'#define LC_MARKER           4   /* Can do markers                   */
'#define LC_POLYMARKER       8   /* Can do polymarkers               */
'#define LC_WIDE             16  /* Can do wide lines                */
'#define LC_STYLED           32  /* Can do styled lines              */
'#define LC_WIDESTYLED       64  /* Can do wide styled lines         */
'#define LC_INTERIORS        128 /* Can do interiors                 */
Public Const LC_NONE = 0            '设备不支持线条。
Public Const LC_POLYLINE = 2        '设备可以绘制折线。
Public Const LC_MARKER = 4          '设备可以绘制一个标记。
Public Const LC_POLYMARKER = 8      '设备可以绘制多个标记。
Public Const LC_WIDE = 16           '设备可以绘制宽线条。
Public Const LC_STYLED = 32         '设备可以绘制样式线条。
Public Const LC_WIDESTYLED = 64     '设备可以绘制宽样式线条。
Public Const LC_INTERIORS = 128     '设备可以绘制内部区域。

'/* Polygonal Capabilities */
'设备的多边形特性。
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
Public Const PC_NONE = 0            '设备不支持多边形。
Public Const PC_POLYGON = 1         '设备可以绘制交替填充的多边形。
Public Const PC_RECTANGLE = 2       '设备可以绘制矩形。
Public Const PC_WINDPOLYGON = 4     '设备可以绘制螺旋填充的多边形。
Public Const PC_TRAPEZOID = 4       '设备可以绘制不规则四边形。
Public Const PC_SCANLINE = 8        '设备可以绘制设备可以绘制单扫描线。
Public Const PC_WIDE = 16           '设备可以绘制宽边框。
Public Const PC_STYLED = 32         '设备可以绘制样式边框。
Public Const PC_WIDESTYLED = 64     '设备可以绘制宽样式边框。
Public Const PC_INTERIORS = 128     '设备可以绘制内部区域。
Public Const PC_POLYPOLYGON = 256   '设备可以绘制多个多边形。
Public Const PC_PATHS = 512         '设备可以绘制路径。

'/* Clipping Capabilities */
'裁剪特性
'#define CP_NONE             0   /* No clipping of output            */
'#define CP_RECTANGLE        1   /* Output clipped to rects          */
'#define CP_REGION           2   /* obsolete                         */
Public Const CP_NONE = 0            '输出不裁剪
Public Const CP_RECTANGLE = 1       '输出裁剪至矩形
Public Const CP_REGION = 2          '作废

'/* Text Capabilities */
'文本特性
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
Public Const TC_OP_CHARACTER = &H1  '设备满足字符输出精度。
Public Const TC_OP_STROKE = &H2     '设备满足笔画输出精度。
Public Const TC_CP_STROKE = &H4     '设备满足笔画裁剪精度。
Public Const TC_CR_90 = &H8         '设备可以90度字符旋转。
Public Const TC_CR_ANY = &H10       '设备可以任意字符旋转。
Public Const TC_SF_X_YINDEP = &H20  '设备可以在X轴和Y轴独立缩放。
Public Const TC_SA_DOUBLE = &H40    '设备支持2倍字符缩放。
Public Const TC_SA_INTEGER = &H80   '设备只能采用字符的整数倍缩放。
Public Const TC_SA_CONTIN = &H100   '设备可以采用字符的任意倍数缩放。
Public Const TC_EA_DOUBLE = &H200   '设备可以绘制双倍磅值的字符。
Public Const TC_IA_ABLE = &H400     '设备支持斜体。
Public Const TC_UA_ABLE = &H800     '设备支持下划线。
Public Const TC_SO_ABLE = &H1000    '设备支持删除线。
Public Const TC_RA_ABLE = &H2000    '设备支持光栅字体。
Public Const TC_VA_ABLE = &H4000    '设备支持矢量字体。
Public Const TC_RESERVED = &H8000   '保留；必须为0。
Public Const TC_SCROLLBLT = &H10000 '文本不允许卷动。

'#endif /* NOGDICAPMASKS */

'/* Raster Capabilities */
'光栅特性
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
Public Const RC_BITBLT = 1              '可以传递位图。
Public Const RC_BANDING = 2             '需要条带化(Banding)支持。
Public Const RC_SCALING = 4             '支持缩放。
Public Const RC_BITMAP64 = 8            '可以支持大于64KB的位图。
Public Const RC_GDI20_OUTPUT = &H10     '
Public Const RC_GDI20_STATE = &H20      '
Public Const RC_SAVEBITMAP = &H40       '
Public Const RC_DI_BITMAP = &H80        '支持SetDIBits和GetDIBits函数。
Public Const RC_PALETTE = &H100         '指定一个基于调色板的设备。
Public Const RC_DIBTODEV = &H200        '支持SetDIBitsToDevice函数。
Public Const RC_BIGFONT = &H400         '支持大于64K的字体。
Public Const RC_STRETCHBLT = &H800      '支持StretchBlt函数。
Public Const RC_FLOODFILL = &H1000      '可以执行flood fills填充操作。
Public Const RC_STRETCHDIB = &H2000     '支持StretchDIBits函数。
Public Const RC_OP_DX_OUTPUT = &H4000
Public Const RC_DEVBITS = &H8000

'#if(WINVER >= 0x0500)

'/* Shading and blending caps                */
'设备的阴影及混合特性。
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
'WinNT自定义纸张控制================================================================
'注意以dmFields是Long型,as Long或尾部加&符
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
Public Const conRatemmToTwip As Single = 56.6857142857143      '毫米与缇的比率

Public Const WM_MOUSEWHEEL = &H20A

'######################################################################################
'   释放内存
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "Psapi" (ByVal hProcess As Long) As Long

'################################################################################################################
'## 图片缩放模式设置
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
'## 判断系统是否是NT
'######################################################################################

Public Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Private Function GetWinPlatform() As Long
'功    能：返回当前的系统版本代号
'参    数：无
'返    回：
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

'######################################################################################
'## 设置打印机的自定义纸张尺寸
'######################################################################################

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'功能：NT环境中，设置打印机的自定义纸张尺寸
'参数：lngWidth、lngHeight=mm(毫米)
'     intOrient=1-纵向,2-横向
'     intCopys=打印份数(如果打印机支持,1-9999,不支持时不会出错,也不影响其它设置)
'说明：除了Width,Height外，其它通过本函数设置的属性不直接反映在Printer上，
'      (取DevMode也反映不出来，可能要用GetJob才能获取最近的打印文档属性)
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
        
        '设置打印文档属性
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '高级打印功能(当取消时,Copies只支持1;但不知怎么取不了)
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
'功能：在设置自定义纸张
'参数：是以绨为单位
    If IsWindowsNT Then
        '虽然不能使宽度生效，但能改变PaperSize的属性值
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustomPager = SetNTPrinterPaper(lngHwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98系列还是以通常方法处理
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function



Public Function GetPaperName(ByVal intSize As Integer) As String
'功能： 根据当前打印机的设置，获取纸张名称
'参数： lngW,lngH=自定义纸张的宽高(Twip)
'返回： 纸张名称
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
        GetPaperName = "不可测的纸张 ..."
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
