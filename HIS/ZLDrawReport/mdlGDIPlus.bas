Attribute VB_Name = "mdlGDIPlus"

Option Explicit

Public Const NotPI = 3.14159265238 / 180


Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type POINTL
   x As Long
   y As Long
End Type

Public Type POINTF
   x As Single
   y As Single
End Type
'
''=================================
'Rectange Structure
Public Type RECTL
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RECTF
   Left As Single
   Top As Single
   Right As Single
   Bottom As Single
End Type

''=================================
'Size Structure
Public Type SIZEL
   CX As Long
   CY As Long
End Type

Public Type Clsid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Public Type EncoderParameter
   Guid As Clsid
   NumberOfValues As Long
   type As EncoderParameterValueType
   value As Long
End Type

Public Type EncoderParameters
   count As Long
   Parameter As EncoderParameter
End Type

''=================================
''== Enums                       ==
''=================================
'
''=================================
'Pixel
Public Enum GpPixelFormat
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
End Enum
'
''=================================
'Unit
Public Enum GpUnit
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

Public Enum StringAlignment
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum


'=================================
'Rotate
Public Enum RotateFlipType
   RotateNoneFlipNone = 0
   Rotate90FlipNone = 1
   Rotate180FlipNone = 2
   Rotate270FlipNone = 3

   RotateNoneFlipX = 4
   Rotate90FlipX = 5
   Rotate180FlipX = 6
   Rotate270FlipX = 7

   RotateNoneFlipY = Rotate180FlipX
   Rotate90FlipY = Rotate270FlipX
   Rotate180FlipY = RotateNoneFlipX
   Rotate270FlipY = Rotate90FlipX

   RotateNoneFlipXY = Rotate180FlipNone
   Rotate90FlipXY = Rotate270FlipNone
   Rotate180FlipXY = RotateNoneFlipNone
   Rotate270FlipxY = Rotate90FlipNone
End Enum


Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, Graphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
''==================================================
'
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus

'
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As Long, Image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
'
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal Filename As Long, clsidEncoder As Clsid, encoderParams As Any) As GpStatus

Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, Bitmap As Long) As GpStatus

''==================================================
'
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, Brush As Long) As GpStatus

Public Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements Lib "gdiplus" (ByVal matrix As Long, ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus

Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus

Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As RotateFlipType) As GpStatus

'===================================================================================
'  不怎么常用的东西
'===================================================================================

''=================================
'Meta File
Public Type PWMFRect16
   Left As Integer
   Top As Integer
   Right As Integer
   Bottom As Integer
End Type

Public Type WmfPlaceableFileHeader
   Key As Long                        ' GDIP_WMF_PLACEABLEKEY
   Hmf As Integer                     ' Metafile HANDLE number (always 0)
   boundingBox As PWMFRect16          ' Coordinates in metafile units
   Inch As Integer                    ' Number of metafile units per inch
   Reserved As Long                   ' Reserved (always 0)
   Checksum As Integer                ' Checksum value for previous 10 WORDs
End Type

Public Type ENHMETAHEADER3
   itype As Long               ' Record type EMR_HEADER
   nSize As Long               ' Record size in bytes.  This may be greater
                               ' than the sizeof(ENHMETAHEADER).
   rclBounds As RECTL        ' Inclusive-inclusive bounds in device units
   rclFrame As RECTL         ' Inclusive-inclusive Picture Frame .01mm unit
   dSignature As Long          ' Signature.  Must be ENHMETA_SIGNATURE.
   nVersion As Long            ' Version number
   nBytes As Long              ' Size of the metafile in bytes
   nRecords As Long            ' Number of records in the metafile
   nHandles As Integer         ' Number of handles in the handle table
                               ' Handle index zero is reserved.
   sReserved As Integer        ' Reserved.  Must be zero.
   nDescription As Long        ' Number of chars in the unicode desc string
                               ' This is 0 if there is no description string
   offDescription As Long      ' Offset to the metafile description record.
                               ' This is 0 if there is no description string
   nPalEntries As Long         ' Number of entries in the metafile palette.
   szlDevice As SIZEL           ' Size of the reference device in pels
   szlMillimeters As SIZEL      ' Size of the reference device in millimeters
End Type

Public Type MetafileHeader
   mType As MetafileType
   size As Long                ' Size of the metafile (in bytes)
   Version As Long             ' EMF+, EMF, or WMF version
   EmfPlusFlags As Long
   DpiX As Single
   DpiY As Single
   x As Long                   ' Bounds in device units
   y As Long
   Width As Long
   Height As Long

   EmfHeader As ENHMETAHEADER3 ' NOTE: You'll have to use CopyMemory to view the METAHEADER type
   EmfPlusHeaderSize As Long   ' size of the EMF+ header in file
   LogicalDpiX As Long         ' Logical Dpi of reference Hdc
   LogicalDpiY As Long         ' usually valid only for EMF+
End Type

'=================================
'Meta File
Public Enum MetafileType
   MetafileTypeInvalid            ' Invalid metafile
   MetafileTypeWmf                ' Standard WMF
   MetafileTypeWmfPlaceable       ' Placeable WMF
   MetafileTypeEmf                ' EMF (not EMF+)
   MetafileTypeEmfPlusOnly        ' EMF+ without dual down-level records
   MetafileTypeEmfPlusDual         ' EMF+ with dual down-level records
End Enum

Public Enum EmfType
    EmfTypeEmfOnly = MetafileTypeEmf               ' no EMF+  only EMF
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly   ' no EMF  only EMF+
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual   ' both EMF+ and EMF
End Enum

Public Enum MetafileFrameUnit
   MetafileFrameUnitPixel = UnitPixel
   MetafileFrameUnitPoint = UnitPoint
   MetafileFrameUnitInch = UnitInch
   MetafileFrameUnitDocument = UnitDocument
   MetafileFrameUnitMillimeter = UnitMillimeter
   MetafileFrameUnitGdi                        ' GDI compatible .01 MM units
End Enum


Public Enum EmfPlusRecordType
   WmfRecordTypeSetBkColor = &H10201
   WmfRecordTypeSetBkMode = &H10102
   WmfRecordTypeSetMapMode = &H10103
   WmfRecordTypeSetROP2 = &H10104
   WmfRecordTypeSetRelAbs = &H10105
   WmfRecordTypeSetPolyFillMode = &H10106
   WmfRecordTypeSetStretchBltMode = &H10107
   WmfRecordTypeSetTextCharExtra = &H10108
   WmfRecordTypeSetTextColor = &H10209
   WmfRecordTypeSetTextJustification = &H1020A
   WmfRecordTypeSetWindowOrg = &H1020B
   WmfRecordTypeSetWindowExt = &H1020C
   WmfRecordTypeSetViewportOrg = &H1020D
   WmfRecordTypeSetViewportExt = &H1020E
   WmfRecordTypeOffsetWindowOrg = &H1020F
   WmfRecordTypeScaleWindowExt = &H10410
   WmfRecordTypeOffsetViewportOrg = &H10211
   WmfRecordTypeScaleViewportExt = &H10412
   WmfRecordTypeLineTo = &H10213
   WmfRecordTypeMoveTo = &H10214
   WmfRecordTypeExcludeClipRect = &H10415
   WmfRecordTypeIntersectClipRect = &H10416
   WmfRecordTypeArc = &H10817
   WmfRecordTypeEllipse = &H10418
   WmfRecordTypeFloodFill = &H10419
   WmfRecordTypePie = &H1081A
   WmfRecordTypeRectangle = &H1041B
   WmfRecordTypeRoundRect = &H1061C
   WmfRecordTypePatBlt = &H1061D
   WmfRecordTypeSaveDC = &H1001E
   WmfRecordTypeSetPixel = &H1041F
   WmfRecordTypeOffsetClipRgn = &H10220
   WmfRecordTypeTextOut = &H10521
   WmfRecordTypeBitBlt = &H10922
   WmfRecordTypeStretchBlt = &H10B23
   WmfRecordTypePolygon = &H10324
   WmfRecordTypePolyline = &H10325
   WmfRecordTypeEscape = &H10626
   WmfRecordTypeRestoreDC = &H10127
   WmfRecordTypeFillRegion = &H10228
   WmfRecordTypeFrameRegion = &H10429
   WmfRecordTypeInvertRegion = &H1012A
   WmfRecordTypePaintRegion = &H1012B
   WmfRecordTypeSelectClipRegion = &H1012C
   WmfRecordTypeSelectObject = &H1012D
   WmfRecordTypeSetTextAlign = &H1012E
   WmfRecordTypeDrawText = &H1062F
   WmfRecordTypeChord = &H10830
   WmfRecordTypeSetMapperFlags = &H10231
   WmfRecordTypeExtTextOut = &H10A32
   WmfRecordTypeSetDIBToDev = &H10D33
   WmfRecordTypeSelectPalette = &H10234
   WmfRecordTypeRealizePalette = &H10035
   WmfRecordTypeAnimatePalette = &H10436
   WmfRecordTypeSetPalEntries = &H10037
   WmfRecordTypePolyPolygon = &H10538
   WmfRecordTypeResizePalette = &H10139
   WmfRecordTypeDIBBitBlt = &H10940
   WmfRecordTypeDIBStretchBlt = &H10B41
   WmfRecordTypeDIBCreatePatternBrush = &H10142
   WmfRecordTypeStretchDIB = &H10F43
   WmfRecordTypeExtFloodFill = &H10548
   WmfRecordTypeSetLayout = &H10149
   WmfRecordTypeResetDC = &H1014C
   WmfRecordTypeStartDoc = &H1014D
   WmfRecordTypeStartPage = &H1004F
   WmfRecordTypeEndPage = &H10050
   WmfRecordTypeAbortDoc = &H10052
   WmfRecordTypeEndDoc = &H1005E
   WmfRecordTypeDeleteObject = &H101F0
   WmfRecordTypeCreatePalette = &H100F7
   WmfRecordTypeCreateBrush = &H100F8
   WmfRecordTypeCreatePatternBrush = &H101F9
   WmfRecordTypeCreatePenIndirect = &H102FA
   WmfRecordTypeCreateFontIndirect = &H102FB
   WmfRecordTypeCreateBrushIndirect = &H102FC
   WmfRecordTypeCreateBitmapIndirect = &H102FD
   WmfRecordTypeCreateBitmap = &H106FE
   WmfRecordTypeCreateRegion = &H106FF
   EmfRecordTypeHeader = 1
   EmfRecordTypePolyBezier = 2
   EmfRecordTypePolygon = 3
   EmfRecordTypePolyline = 4
   EmfRecordTypePolyBezierTo = 5
   EmfRecordTypePolyLineTo = 6
   EmfRecordTypePolyPolyline = 7
   EmfRecordTypePolyPolygon = 8
   EmfRecordTypeSetWindowExtEx = 9
   EmfRecordTypeSetWindowOrgEx = 10
   EmfRecordTypeSetViewportExtEx = 11
   EmfRecordTypeSetViewportOrgEx = 12
   EmfRecordTypeSetBrushOrgEx = 13
   EmfRecordTypeEOF = 14
   EmfRecordTypeSetPixelV = 15
   EmfRecordTypeSetMapperFlags = 16
   EmfRecordTypeSetMapMode = 17
   EmfRecordTypeSetBkMode = 18
   EmfRecordTypeSetPolyFillMode = 19
   EmfRecordTypeSetROP2 = 20
   EmfRecordTypeSetStretchBltMode = 21
   EmfRecordTypeSetTextAlign = 22
   EmfRecordTypeSetColorAdjustment = 23
   EmfRecordTypeSetTextColor = 24
   EmfRecordTypeSetBkColor = 25
   EmfRecordTypeOffsetClipRgn = 26
   EmfRecordTypeMoveToEx = 27
   EmfRecordTypeSetMetaRgn = 28
   EmfRecordTypeExcludeClipRect = 29
   EmfRecordTypeIntersectClipRect = 30
   EmfRecordTypeScaleViewportExtEx = 31
   EmfRecordTypeScaleWindowExtEx = 32
   EmfRecordTypeSaveDC = 33
   EmfRecordTypeRestoreDC = 34
   EmfRecordTypeSetWorldTransform = 35
   EmfRecordTypeModifyWorldTransform = 36
   EmfRecordTypeSelectObject = 37
   EmfRecordTypeCreatePen = 38
   EmfRecordTypeCreateBrushIndirect = 39
   EmfRecordTypeDeleteObject = 40
   EmfRecordTypeAngleArc = 41
   EmfRecordTypeEllipse = 42
   EmfRecordTypeRectangle = 43
   EmfRecordTypeRoundRect = 44
   EmfRecordTypeArc = 45
   EmfRecordTypeChord = 46
   EmfRecordTypePie = 47
   EmfRecordTypeSelectPalette = 48
   EmfRecordTypeCreatePalette = 49
   EmfRecordTypeSetPaletteEntries = 50
   EmfRecordTypeResizePalette = 51
   EmfRecordTypeRealizePalette = 52
   EmfRecordTypeExtFloodFill = 53
   EmfRecordTypeLineTo = 54
   EmfRecordTypeArcTo = 55
   EmfRecordTypePolyDraw = 56
   EmfRecordTypeSetArcDirection = 57
   EmfRecordTypeSetMiterLimit = 58
   EmfRecordTypeBeginPath = 59
   EmfRecordTypeEndPath = 60
   EmfRecordTypeCloseFigure = 61
   EmfRecordTypeFillPath = 62
   EmfRecordTypeStrokeAndFillPath = 63
   EmfRecordTypeStrokePath = 64
   EmfRecordTypeFlattenPath = 65
   EmfRecordTypeWidenPath = 66
   EmfRecordTypeSelectClipPath = 67
   EmfRecordTypeAbortPath = 68
   EmfRecordTypeReserved_069 = 69
   EmfRecordTypeGdiComment = 70
   EmfRecordTypeFillRgn = 71
   EmfRecordTypeFrameRgn = 72
   EmfRecordTypeInvertRgn = 73
   EmfRecordTypePaintRgn = 74
   EmfRecordTypeExtSelectClipRgn = 75
   EmfRecordTypeBitBlt = 76
   EmfRecordTypeStretchBlt = 77
   EmfRecordTypeMaskBlt = 78
   EmfRecordTypePlgBlt = 79
   EmfRecordTypeSetDIBitsToDevice = 80
   EmfRecordTypeStretchDIBits = 81
   EmfRecordTypeExtCreateFontIndirect = 82
   EmfRecordTypeExtTextOutA = 83
   EmfRecordTypeExtTextOutW = 84
   EmfRecordTypePolyBezier16 = 85
   EmfRecordTypePolygon16 = 86
   EmfRecordTypePolyline16 = 87
   EmfRecordTypePolyBezierTo16 = 88
   EmfRecordTypePolylineTo16 = 89
   EmfRecordTypePolyPolyline16 = 90
   EmfRecordTypePolyPolygon16 = 91
   EmfRecordTypePolyDraw16 = 92
   EmfRecordTypeCreateMonoBrush = 93
   EmfRecordTypeCreateDIBPatternBrushPt = 94
   EmfRecordTypeExtCreatePen = 95
   EmfRecordTypePolyTextOutA = 96
   EmfRecordTypePolyTextOutW = 97
   EmfRecordTypeSetICMMode = 98
   EmfRecordTypeCreateColorSpace = 99
   EmfRecordTypeSetColorSpace = 100
   EmfRecordTypeDeleteColorSpace = 101
   EmfRecordTypeGLSRecord = 102
   EmfRecordTypeGLSBoundedRecord = 103
   EmfRecordTypePixelFormat = 104
   EmfRecordTypeDrawEscape = 105
   EmfRecordTypeExtEscape = 106
   EmfRecordTypeStartDoc = 107
   EmfRecordTypeSmallTextOut = 108
   EmfRecordTypeForceUFIMapping = 109
   EmfRecordTypeNamedEscape = 110
   EmfRecordTypeColorCorrectPalette = 111
   EmfRecordTypeSetICMProfileA = 112
   EmfRecordTypeSetICMProfileW = 113
   EmfRecordTypeAlphaBlend = 114
   EmfRecordTypeSetLayout = 115
   EmfRecordTypeTransparentBlt = 116
   EmfRecordTypeReserved_117 = 117
   EmfRecordTypeGradientFill = 118
   EmfRecordTypeSetLinkedUFIs = 119
   EmfRecordTypeSetTextJustification = 120
   EmfRecordTypeColorMatchToTargetW = 121
   EmfRecordTypeCreateColorSpaceW = 122
   EmfRecordTypeMax = 122
   EmfRecordTypeMin = 1

   EmfPlusRecordTypeInvalid = 16384 '//GDIP_EMFPLUS_RECORD_BASE
   EmfPlusRecordTypeHeader = 16385
   EmfPlusRecordTypeEndOfFile = 16386
   EmfPlusRecordTypeComment = 16387
   EmfPlusRecordTypeGetDC = 16388
   EmfPlusRecordTypeMultiFormatStart = 16389
   EmfPlusRecordTypeMultiFormatSection = 16390
   EmfPlusRecordTypeMultiFormatEnd = 16391

   EmfPlusRecordTypeObject = 16392

   EmfPlusRecordTypeClear = 16393
   EmfPlusRecordTypeFillRects = 16394
   EmfPlusRecordTypeDrawRects = 16395
   EmfPlusRecordTypeFillPolygon = 16396
   EmfPlusRecordTypeDrawLines = 16397
   EmfPlusRecordTypeFillEllipse = 16398
   EmfPlusRecordTypeDrawEllipse = 16399
   EmfPlusRecordTypeFillPie = 16400
   EmfPlusRecordTypeDrawPie = 16401
   EmfPlusRecordTypeDrawArc = 16402
   EmfPlusRecordTypeFillRegion = 16403
   EmfPlusRecordTypeFillPath = 16404
   EmfPlusRecordTypeDrawPath = 16405
   EmfPlusRecordTypeFillClosedCurve = 16406
   EmfPlusRecordTypeDrawClosedCurve = 16407
   EmfPlusRecordTypeDrawCurve = 16408
   EmfPlusRecordTypeDrawBeziers = 16409
   EmfPlusRecordTypeDrawImage = 16410
   EmfPlusRecordTypeDrawImagePoints = 16411
   EmfPlusRecordTypeDrawString = 16412

   EmfPlusRecordTypeSetRenderingOrigin = 16413
   EmfPlusRecordTypeSetAntiAliasMode = 16414
   EmfPlusRecordTypeSetTextRenderingHint = 16415
   EmfPlusRecordTypeSetTextContrast = 16416
   EmfPlusRecordTypeSetInterpolationMode = 16417
   EmfPlusRecordTypeSetPixelOffsetMode = 16418
   EmfPlusRecordTypeSetCompositingMode = 16419
   EmfPlusRecordTypeSetCompositingQuality = 16420
   EmfPlusRecordTypeSave = 16421
   EmfPlusRecordTypeRestore = 16422
   EmfPlusRecordTypeBeginContainer = 16423
   EmfPlusRecordTypeBeginContainerNoParams = 16424
   EmfPlusRecordTypeEndContainer = 16425
   EmfPlusRecordTypeSetWorldTransform = 16426
   EmfPlusRecordTypeResetWorldTransform = 16427
   EmfPlusRecordTypeMultiplyWorldTransform = 16428
   EmfPlusRecordTypeTranslateWorldTransform = 16429
   EmfPlusRecordTypeScaleWorldTransform = 16430
   EmfPlusRecordTypeRotateWorldTransform = 16431
   EmfPlusRecordTypeSetPageTransform = 16432
   EmfPlusRecordTypeResetClip = 16433
   EmfPlusRecordTypeSetClipRect = 16434
   EmfPlusRecordTypeSetClipPath = 16435
   EmfPlusRecordTypeSetClipRegion = 16436
   EmfPlusRecordTypeOffsetClip = 16437
   EmfPlusRecordTypeDrawDriverString = 16438
   EmfPlusRecordTotal = 16439
   EmfPlusRecordTypeMax = 16438
   EmfPlusRecordTypeMin = 16385
End Enum

Public Enum FlushIntention
   FlushIntentionFlush = 0         ' Flush all batched rendering operations
   FlushIntentionSync = 1          ' Flush all batched rendering operations
End Enum


Public Enum EncoderParameterValueType
   EncoderParameterValueTypeByte = 1              ' 8-bit unsigned int
   EncoderParameterValueTypeASCII = 2             ' 8-bit byte containing one 7-bit ASCII
                                                   ' code. NULL terminated.
   EncoderParameterValueTypeShort = 3             ' 16-bit unsigned int
   EncoderParameterValueTypeLong = 4              ' 32-bit unsigned int
   EncoderParameterValueTypeRational = 5          ' Two Longs. The first Long is the
                                                   ' numerator the second Long expresses the
                                                   ' denomintor.
   EncoderParameterValueTypeLongRange = 6         ' Two longs which specify a range of
                                                   ' integer values. The first Long specifies
                                                   ' the lower end and the second one
                                                   ' specifies the higher end. All values
                                                   ' are inclusive at both ends
   EncoderParameterValueTypeUndefined = 7         ' 8-bit byte that can take any value
                                                   ' depending on field definition
   EncoderParameterValueTypeRationalRange = 8      ' Two Rationals. The first Rational
                                                   ' specifies the lower end and the second
                                                   ' specifies the higher end. All values
                                                   ' are inclusive at both ends
End Enum

Public Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As GpStatus


'===================================================================================
'  公共部分 / 其他部分
'===================================================================================

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus

Public Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Public Enum GpStatus
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As Clsid) As Long

Public Enum ImageType
    bmp
    EMF
    WMF
    JPG
    PNG
    GIF
    TIF
    ICO
End Enum
'
Public Const ImageEncoderSuffix       As String = "-1A04-11D3-9A73-0000F81EF32E}"
Public Const ImageEncoderBMP          As String = "{557CF400" & ImageEncoderSuffix
Public Const ImageEncoderJPG          As String = "{557CF401" & ImageEncoderSuffix
Public Const ImageEncoderGIF          As String = "{557CF402" & ImageEncoderSuffix
Public Const ImageEncoderEMF          As String = "{557CF403" & ImageEncoderSuffix
Public Const ImageEncoderWMF          As String = "{557CF404" & ImageEncoderSuffix
Public Const ImageEncoderTIF          As String = "{557CF405" & ImageEncoderSuffix
Public Const ImageEncoderPNG          As String = "{557CF406" & ImageEncoderSuffix
Public Const ImageEncoderICO          As String = "{557CF407" & ImageEncoderSuffix
Public Const EncoderCompression       As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EncoderColorDepth        As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EncoderScanMethod        As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EncoderVersion           As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EncoderRenderMethod      As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EncoderQuality           As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EncoderTransformation    As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EncoderLuminanceTable    As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EncoderChrominanceTable  As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EncoderSaveFlag          As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"

Dim mToken As Long

Dim Pens() As Long, PenCount As Long
Dim Brushes() As Long, BrushCount As Long
Dim StrFormats() As Long, StrFormatCount As Long
Dim Matrixes() As Long, MatrixCount As Long

Public Function DeleteObjects()
    
    On Error Resume Next
    
    Dim i As Long
    For i = 1 To PenCount: GdipDeletePen Pens(i): Next
    For i = 1 To BrushCount: GdipDeleteBrush Brushes(i): Next
    For i = 1 To StrFormatCount: GdipDeleteStringFormat StrFormats(i): Next
    For i = 1 To MatrixCount: GdipDeleteMatrix Matrixes(i): Next
    PenCount = 0
    BrushCount = 0
    StrFormatCount = 0
    MatrixCount = 0
End Function

Public Function NewPen(Color As Long, Width As Single) As Long
    PenCount = PenCount + 1
    ReDim Preserve Pens(PenCount)
    
    GdipCreatePen1 Color, Width, UnitPixel, Pens(PenCount)
    NewPen = Pens(PenCount)
End Function

Public Function NewBrush(Color As Long) As Long
    BrushCount = BrushCount + 1
    ReDim Preserve Brushes(BrushCount)
    
    GdipCreateSolidFill Color, Brushes(BrushCount)
    NewBrush = Brushes(BrushCount)
End Function

Public Function NewStringFormat(Align As StringAlignment) As Long
    StrFormatCount = StrFormatCount + 1
    ReDim Preserve StrFormats(StrFormatCount)
    
    GdipCreateStringFormat 0, 0, StrFormats(StrFormatCount)
    GdipSetStringFormatAlign StrFormats(StrFormatCount), Align
    NewStringFormat = StrFormats(StrFormatCount)
End Function

Public Function NewMatrix(m11 As Single, m12 As Single, m21 As Single, m22 As Single, dx As Single, dy As Single) As Long
    MatrixCount = MatrixCount + 1
    ReDim Preserve Matrixes(MatrixCount)
    
    GdipCreateMatrix Matrixes(MatrixCount)
    GdipSetMatrixElements Matrixes(MatrixCount), m11, m12, m21, m22, dx, dy
    NewMatrix = Matrixes(MatrixCount)
End Function

Public Function NewRectF(Left As Single, Top As Single, Width As Single, Height As Single) As RECTF
    With NewRectF
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function NewRectL(Left As Single, Top As Long, Width As Long, Height As Long) As RECTL
    With NewRectL
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function InitGDIPlus(Optional OnErrShowMsg As Boolean = True, Optional OnErrEndApp As Boolean = True, _
                            Optional ErrMsgText As String = "GDI+ 初始化错误。程序即将关闭。", _
                            Optional ErrMsgStyle As VbMsgBoxStyle = vbCritical, _
                            Optional ErrMsgTitle As String = "初始化错误") As GpStatus
    
    If mToken <> 0 Then
'        Debug.Print "InitGDIPlus> GdiPlus已被初始化"
        Exit Function
    End If
    
    Dim uInput As GdiplusStartupInput
    Dim ret As GpStatus
    
    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(mToken, uInput)
    If ret <> Ok Then
        If OnErrShowMsg Then MsgBox ErrMsgText, ErrMsgStyle, ErrMsgTitle
        If OnErrEndApp Then Exit Function
    End If
    
    InitGDIPlus = ret
End Function

Public Sub TerminateGDIPlus()
    If mToken = 0 Then
'        Debug.Print "TerminateGDIPlus> GdiPlus已被结束"
        Exit Sub
    End If
    
    DeleteObjects
    
    GdiplusShutdown mToken
    
    mToken = 0
End Sub

Public Function InitGDIPlusTo(ByRef token As Long, _
                              Optional OnErrShowMsg As Boolean = True, Optional OnErrEndApp As Boolean = True, _
                              Optional ErrMsgText As String = "GDI+ 初始化错误。程序即将关闭。", _
                              Optional ErrMsgStyle As VbMsgBoxStyle = vbCritical, _
                              Optional ErrMsgTitle As String = "初始化错误") As GpStatus
    
    If token <> 0 Then
'        Debug.Print "InitGDIPlusTo> GdiPlus已被初始化"
        Exit Function
    End If
    
    Dim uInput As GdiplusStartupInput
    Dim ret As GpStatus
    
    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(token, uInput)
    If ret <> Ok Then
        If OnErrShowMsg Then MsgBox ErrMsgText, ErrMsgStyle, ErrMsgTitle
        If OnErrEndApp Then Exit Function
    End If
    
    InitGDIPlusTo = ret
End Function

Public Sub TerminateGDIPlusFrom(ByVal token As Long)
    
    On Error Resume Next
    
    'Debug.Print "TerminateGDIPlus> GdiPlus已被结束"
    If token = 0 Then Exit Sub
        
    DeleteObjects
    
    GdiplusShutdown token
    
    token = 0
End Sub

Public Function GetImageEncoderClsid(ByVal ImageType As ImageType) As Clsid
    Select Case ImageType
        Case PNG: CLSIDFromString StrPtr(ImageEncoderPNG), GetImageEncoderClsid
        Case JPG: CLSIDFromString StrPtr(ImageEncoderJPG), GetImageEncoderClsid
        Case GIF: CLSIDFromString StrPtr(ImageEncoderGIF), GetImageEncoderClsid
        Case bmp: CLSIDFromString StrPtr(ImageEncoderBMP), GetImageEncoderClsid
        Case ICO: CLSIDFromString StrPtr(ImageEncoderICO), GetImageEncoderClsid
        Case EMF: CLSIDFromString StrPtr(ImageEncoderEMF), GetImageEncoderClsid
        Case WMF: CLSIDFromString StrPtr(ImageEncoderWMF), GetImageEncoderClsid
        Case TIF: CLSIDFromString StrPtr(ImageEncoderTIF), GetImageEncoderClsid
    End Select
End Function

Public Function SaveImageToPNG(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToPNG = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEncoderClsid(PNG), ByVal 0)
End Function

Public Function SaveImageToJPG(ByVal Image As Long, ByVal Path As String, ByVal Quality As Long) As GpStatus
    Dim Params As EncoderParameters
    
    Params.count = 1
    CLSIDFromString StrPtr(EncoderQuality), Params.Parameter.Guid
    Params.Parameter.NumberOfValues = 1
    Params.Parameter.type = 4
    Params.Parameter.value = VarPtr(Quality)
    
    SaveImageToJPG = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEncoderClsid(JPG), Params)
End Function

Public Function SaveImageToGIF(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToGIF = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEncoderClsid(GIF), ByVal 0)
End Function

Public Function SaveImageToBMP(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToBMP = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEncoderClsid(bmp), ByVal 0)
End Function

Public Function CreateBitmap(ByRef Bitmap As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus
    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, Bitmap
End Function

Public Function CreateBitmapWithGraphics(ByRef Bitmap As Long, ByRef Graphics As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus
    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, Bitmap
    GdipGetImageGraphicsContext Bitmap, Graphics
End Function














