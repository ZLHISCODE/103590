Attribute VB_Name = "mdlLisDraw"
Option Explicit

'**************************
'���²��־�ΪLIS�Ļ�ͼ��ض��弰����
'**************************
Private mclsUnzip As New cUnzip

Private Const UnitPixel As Long = 2

Private Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type

Private Enum ImageLockMode
   ImageLockModeRead = &H1
   ImageLockModeWrite = &H2
   ImageLockModeUserInputBuf = &H4
End Enum

Private Type BitmapData
   Width As Long
   Height As Long
   Stride As Long
   PixelFormat As Long
   scan0 As Long
   Reserved As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type

Private Enum PixelFormat
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


Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Enum PaletteFlags
    [PaletteFlagsHasAlpha] = &H1
    [PaletteFlagsGrayScale] = &H2
    [PaletteFlagsHalftone] = &H4
End Enum

Private Type ColorPalette '(8bpp)
   flags        As PaletteFlags
   count        As Long
   Entries(255) As RGBQUAD
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GDIPlus" (ByVal Filename As String, Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal m_Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal Graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDC As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal Graphics As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipBitmapLockBits Lib "GDIPlus" (ByVal BITMAP As Long, Rct As RECT, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GDIPlus" (ByVal BITMAP As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GDIPlus" (ByVal Image As Long, PixelFormat As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long

Private Declare Function GdipGetImagePaletteSize Lib "GDIPlus" (ByVal Image As Long, Size As Long) As Long
Private Declare Function GdipGetImagePalette Lib "GDIPlus" (ByVal Image As Long, Palette As ColorPalette, ByVal Size As Long) As Long
Private Declare Function GdipSetImagePalette Lib "GDIPlus" (ByVal hImage As Long, Palette As ColorPalette) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal Filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long


' ***************************************************
' *             �ı���תģ��                        *
' *                                                 *
' ***************************************************

Public uDisplayDescript  As Boolean      'ѡ��ʱ��ʾ��ϸ����

'API ����:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

'�ṹ����:
Private Type PointAPI
    x   As Long
    Y   As Long
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

'API ����:
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'----- ����ΪJPG��ʽ��ͼƬ
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type



Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public dX As Long, dy As Long          ' distance XY = size of snapping zone
Public X1 As Long, X2 As Long          ' co�rdinates snapping zone
Public Y1 As Long, Y2 As Long

'--------------------------------------------------------------
Dim lngTime
'       ���뺯�� �� ���ش���ʽ˵��
'    ResultFromFile ����  ���ַ������鷽ʽ���ؽ�����, һ������Ԫ�ذ���һ�������;
'    Analyse        ����  ���ַ�����ʽ���ؽ�����,ÿ���������||�ָ�
'    ÿ���������Ԫ��֮����|�ָ�,������ϸ˵��



'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath) <> "" Then
        zlFileUnzip = strZipPath & Dir(strZipPath)
    Else
        zlFileUnzip = ""
    End If
End Function

'*************************************************************************
'**�� �� ����PrintRotText
'**��    �룺ByVal hDC(Long)          -
'**        ��ByVal Text(String)       -  Ҫ��ӡ������
'**        ��ByVal CenterX(Long)      -  X���ĵ����������
'**        ��ByVal CenterY(Long)      -  Y���ĵ����������
'**        ��ByVal RotDegrees(Single) -  ��ת�Ƕ�(0.0 �� 359.9999999) ��˳ʱ�룬0=ˮƽ(����ת)
'**��    ����(Boolean) -
'**������������һ��������������X,����Y���������ԽǶȻ�����ת����
'**ȫ�ֱ�����
'**����ģ�飺
'*************************************************************************
Public Function PrintRotText(ByVal hDC As Long, ByVal Text As String, ByVal CenterX As Long, ByVal CenterY As Long, ByVal RotDegrees As Single) As Boolean

Dim bOkSoFar    As Boolean      '������ʶ.
Dim hFontOld    As Long         'ԭ������
Dim hFontNew    As Long         '��������
Dim lfFont      As LOGFONT      'LOGFONT ������ṹ.
Dim ptOrigin    As PointAPI     '���ֻ���ԭ��
Dim ptCenter    As PointAPI     '�������ĵ�.
Dim szText      As SizeStruct   '���ֿ�Ⱥ͸߶�

    '���豸�еõ���ǰ LOGFONT �ṹ.
    hFontOld = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    
    '������豸�õ�������ɹ�...
    If hFontOld <> 0 Then
        
        '�������ȡ LOGFONT �ṹ
        bOkSoFar = (GetObjectAPI(hFontOld, Len(lfFont), lfFont) <> 0)
        
        '��ԭ��������
        Call SelectObject(hDC, hFontOld)
        
        '��λ�Ժ�ʹ��
        hFontOld = 0
    End If
    
    '����ɹ���� LOGFONT �ṹ������.
    If bOkSoFar Then
    
        '�ı����巽��ͳ���
        lfFont.lfEscapement = RotDegrees * 10
        lfFont.lfOrientation = lfFont.lfEscapement
        lfFont.lfQuality = ANTIALIASED_QUALITY
        
        '�� LOGFONT �ṹ�д������������
        hFontNew = CreateFontIndirect(lfFont)
        
        '���崴���ɹ�
        If hFontNew <> 0 Then
            
            'Select the neѡ���µ����嵽���豸
            hFontOld = SelectObject(hDC, hFontNew)
            
            '�ɹ�
            If hFontOld <> 0 Then
                
                '��ȡ�����߼���λ��С(����)
                bOkSoFar = (GetTextExtentPoint32(hDC, Text, LenB(StrConv(Text, vbFromUnicode)), szText) <> 0)
                
                '�ɹ�
                If bOkSoFar Then
                    
                    '��������ˮƽԭ��
                    With ptOrigin
                        .x = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    'ת�� CenterX, CenterY ����ṹ
                    '(��Ҫ���� RotatePoint).
                    With ptCenter
                        .x = CenterX
                        .Y = CenterY
                    End With
                    
                    '��ԭ��ѡ����ƥ��Ԥ��ѡ��
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    '���ڴ�ӡ��ת�ı������سɹ�/ʧ��
                    PrintRotText = (TextOut(hDC, ptOrigin.x, _
                      ptOrigin.Y, Text, LenB(StrConv(Text, vbFromUnicode))) <> 0)
                
                End If
                
                '�ָ����嵽ԭ���豸
                hFontNew = SelectObject(hDC, hFontOld)
            
            End If
            
            '����ڴ沢ɾ������������
            Call DeleteObject(hFontNew)
        
        End If
        
    End If
            
End Function


'*************************************************************************
'**    ��    �� ��    laviewpbt
'**    �� �� �� ��    SavePic
'**    ��    �� ��    pic(StdPicture)        -   ͼ����
'**             ��    FileName(String)       -   ����·��
'**             ��    Quality(Byte)          -   JPGͼ������
'**             ��    TIFF_ColorDepth(Long)  -   TTF��ʽ����ɫ���
'**             ��    TIFF_Compression(Long) -   TTF��ʽ��ѹ����
'**    ��    �� ��    ��
'**    �������� ��    ��ͼ�󱣴�ΪJPG��TIFF��PNG��GIF��BMP��ʽ
'**    ��    �� ��
'**    �� �� �� ��    laviewpbt
'**    ��    �� ��    2005-10-23 14.43.52
'**    ��    �� ��    Version 1.2.1
'*************************************************************************
Public Sub SavePic(ByVal pict As StdPicture, ByVal Filename As String, picType As String, _
                    Optional ByVal Quality As Byte = 100, _
                    Optional ByVal TIFF_ColorDepth As Long = 24, _
                    Optional ByVal TIFF_Compression As Long = 6)
100    Screen.MousePointer = vbHourglass
       Dim tSI As GdiplusStartupInput
       Dim lRes As Long
       Dim lGDIP As Long
       Dim lBitmap As Long
       Dim aEncParams() As Byte
       On Error GoTo errHandle:
102    tSI.GdiplusVersion = 1   ' ��ʼ�� GDI+
104    lRes = GdiplusStartup(lGDIP, tSI)
106    If lRes = 0 Then     ' �Ӿ������ GDI+ ͼ��
108       lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap)
110       If lRes = 0 Then
             Dim tJpgEncoder As GUID
             Dim tParams As EncoderParameters    '��ʼ����������GUID��ʶ
112          Select Case UCase(picType)
             Case ".JPG", "JPG", ".JPEG", "JPEG"
114             CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
116             tParams.count = 1                               ' ���ý���������
118             With tParams.Parameter ' Quality
120                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID    ' �õ�Quality������GUID��ʶ
122                .NumberOfValues = 1
124                .type = 4
126                .Value = VarPtr(Quality)
                End With
128             ReDim aEncParams(1 To Len(tParams))
130             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
132         Case ".PNG", "PNG"
134              CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
136              ReDim aEncParams(1 To Len(tParams))
138         Case ".GIF", "GIF"
140              CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
142              ReDim aEncParams(1 To Len(tParams))
144         Case ".TIFF", "TIFF"
146              CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
148              tParams.count = 2
150              ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
152              With tParams.Parameter
154                 .NumberOfValues = 1
156                 .type = 4
158                  CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID    ' �õ�ColorDepth������GUID��ʶ
160                 .Value = VarPtr(TIFF_Compression)
                End With
162             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
164             With tParams.Parameter
166                 .NumberOfValues = 1
168                 .type = 4
170                  CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID    ' �õ�Compression������GUID��ʶ
172                 .Value = VarPtr(TIFF_ColorDepth)
                End With
174             Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
176         Case ".BMP", "BMP"                                              '������ǰд����ΪBMP�Ĵ��룬��Ϊ��û����GDI+
178             SavePicture pict, Filename
180             Screen.MousePointer = vbDefault
                Exit Sub
            End Select
182          lRes = GdipSaveImageToFile(lBitmap, StrPtr(Filename), tJpgEncoder, aEncParams(1))             '����ͼ��
184          GdipDisposeImage lBitmap       ' ����GDI+ͼ��
          End If
186       GdiplusShutdown lGDIP              '���� GDI+
       End If
188    Screen.MousePointer = vbDefault
190    Erase aEncParams
       Exit Sub
errHandle:
192     Screen.MousePointer = vbDefault
End Sub


'*************************************************************************
'**�� �� ����RotatePoint
'**��    �룺ptAxis(PointAPI)   -
'**        ��ptRotate(PointAPI) -
'**        ��fDegrees(Single)   -
'**��    ������
'**������������ǰfdegrees��ǰ����ѡ��ptRotate���ҵ�ptAxis
'**ȫ�ֱ�����
'**����ģ�飺
'*************************************************************************
Private Sub RotatePoint(ptAxis As PointAPI, ptRotate As PointAPI, fDegrees As Single)

' ***************************************************
' *                 RotatePoint                     *
' *                                                 *
' *  Created by: Rocky Clark (Kath-Rock Software)   *
' *                                                 *
' *  Rotate ptRotate around ptAxis, fDegrees from   *
' *  its current position.                          *
' *                                                 *
' * This procedure may be used and distributed, as  *
' * is, in your code, as long as these credits and  *
' * the code itself remain unchanged.               *
' *                                                 *
' ***************************************************

Dim fDX     As Single   'X����
Dim fDY     As Single   'Y����
Dim fRads   As Single   '����
Const dPi   As Double = 3.14159265358979  'Pi Բ����


    'ת���Ƕ�Ϊ����
    fRads = fDegrees * (dPi / 180#)
    
    '�����ĵ�������
    fDX = ptRotate.x - ptAxis.x
    fDY = ptRotate.Y - ptAxis.Y
    
    '��ת��
    ptRotate.x = ptAxis.x + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub

Public Function CheckGif(ByVal strFile As String) As Boolean
    '���GIF�ļ������Ƿ�����
    'GIF��ͷ��00 3B����
    Dim intFileNo As Integer, lngFileSize As Long, arrEnd(2) As Byte, arrTitle(3) As Byte
    Dim lngCount As Long
    On Error GoTo hErr
100 If Dir(strFile) <> "" Then
102     intFileNo = FreeFile
104     Open strFile For Binary Access Read As intFileNo
106     lngFileSize = LOF(intFileNo)
108     If lngFileSize > 0 Then
110         Get intFileNo, , arrTitle
112         Seek intFileNo, lngFileSize - 1
114         Get intFileNo, , arrEnd
        End If
116     Close intFileNo
        
118     If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1)) & Chr(arrTitle(2))) = "GIF" And arrEnd(0) = 0 And arrEnd(1) = 59 Then
120         CheckGif = True
        End If
        '�ж��Ƿ��Լ�����ͼƬ����ġ�ʹ�ÿؼ�����ͼƬ��
        '����gif��ʽͼƬ����ֻ�Ǻ�׺����Ϊgif��ʽ��ʵ�ʱ����ͼƬ����bmpͼƬ��
        If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1))) = "BM" Then
            CheckGif = True
        End If
    End If
    Exit Function
hErr:

End Function


