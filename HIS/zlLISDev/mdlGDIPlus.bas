Attribute VB_Name = "mdlGDIPlus"
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

Private Declare Function GdiplusStartup Lib "GDIPlus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal Token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GDIPlus" (ByVal FileName As String, Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal m_Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal Graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal Hdc As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal Graphics As Long) As Long

Private Declare Function GdipBitmapLockBits Lib "GDIPlus" (ByVal bitmap As Long, Rct As RECT, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GDIPlus" (ByVal bitmap As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GDIPlus" (ByVal Image As Long, PixelFormat As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long

Private Declare Function GdipGetImagePaletteSize Lib "GDIPlus" (ByVal Image As Long, Size As Long) As Long
Private Declare Function GdipGetImagePalette Lib "GDIPlus" (ByVal Image As Long, Palette As ColorPalette, ByVal Size As Long) As Long
Private Declare Function GdipSetImagePalette Lib "GDIPlus" (ByVal hImage As Long, Palette As ColorPalette) As Long

Private Sub MakePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByRef OldArrPtr As Long, ByRef OldpArrPtr As Long)
    Dim Temp As Long, TempPtr As Long
    CopyMemory Temp, ByVal DataArrPtr, 4        '�õ�DataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
    Temp = Temp + 12                            '���ָ��ƫ��12���ֽں����pvDataָ��
    CopyMemory TempPtr, ByVal pDataArrPtr, 4    '�õ�pDataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
    TempPtr = TempPtr + 12                      '���ָ��ƫ��12���ֽں����pvDataָ��
    CopyMemory OldpArrPtr, ByVal TempPtr, 4     '����ɵ�ַ
    CopyMemory ByVal TempPtr, Temp, 4           'ʹpDataArrPtrָ��DataArrPtr��SAFEARRAY�ṹ��pvDataָ��
    CopyMemory OldArrPtr, ByVal Temp, 4         '����ɵ�ַ
End Sub


Private Sub FreePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByVal OldArrPtr As Long, ByVal OldpArrPtr As Long)
    Dim TempPtr As Long
    CopyMemory TempPtr, ByVal DataArrPtr, 4         '�õ�DataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
    CopyMemory ByVal (TempPtr + 12), OldArrPtr, 4   '�ָ��ɵ�ַ
    CopyMemory TempPtr, ByVal pDataArrPtr, 4        '�õ�pDataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
    CopyMemory ByVal (TempPtr + 12), OldpArrPtr, 4  '�ָ��ɵ�ַ
End Sub


Public Sub PicInvertAndSave(ByRef objPic As PictureBox, ByVal strFileName As String, strFileType As String)

    Dim Token               As Long
    Dim Gsp                 As GdiplusStartupInput
    
    Dim Img                 As Long, Graphics           As Long
    Dim lRes As Long
        
    Gsp.GdiplusVersion = 1
    GdiplusStartup Token, Gsp                       '����GDI+
    
    GdipCreateBitmapFromHBITMAP objPic.Picture.Handle, objPic.Picture.hpal, Img     '����������Խ�GDI��Stdpicture����ת��ΪGDI+��Image����
    Invert Img
    GdipCreateFromHDC objPic.Hdc, Graphics                                              '��DC�л���ͼ��
    GdipDrawImageRectRectI Graphics, Img, 0, 0, objPic.ScaleWidth, objPic.ScaleHeight, 0, 0, objPic.ScaleWidth, objPic.ScaleHeight, UnitPixel
    
    GdipDeleteGraphics Graphics

    GdipDisposeImage Img                    '�������ǵ�GDI+����
    GdiplusShutdown Token                   '�ر�GDI+
    objPic.Refresh                            'ˢ��
    Call SavePic(objPic.Image, strFileName, strFileType)
 
End Sub


Private Function Invert(Image As Long) As Boolean
    Dim PixelFormat         As Long
    
    Dim Dimensions          As RECTF, Rct               As RECT
    Dim BmpData             As BitmapData, Rtn          As Long
    
    Dim DataArr(0 To 3)     As Byte, pDataArr(0 To 0)   As Long
    Dim OldArrPtr           As Long, OldpArrPtr         As Long
    Dim LineAddBytes        As Long, PixelAddBytes      As Long
    
    Dim X                   As Long, Y                  As Long

    GdipGetImageBounds Image, Dimensions, UnitPixel                   ' �õ�ͼ��Ĵ�С��������Ϊ��λ,������������ϲ��᲻�ɹ�
    GdipGetImagePixelFormat Image, PixelFormat                        ' �õ�ͼ���������ʽ
    Select Case PixelFormat
    Case PixelFormat32bppRGB, PixelFormat24bppRGB
        Rct.Right = Dimensions.nWidth
        Rct.Bottom = Dimensions.nHeight
        GdipBitmapLockBits Image, Rct, ImageLockModeRead, PixelFormat, BmpData    '��ȡͼ�������
        MakePoint VarPtrArray(DataArr), VarPtrArray(pDataArr), OldArrPtr, OldpArrPtr    'ģ��ָ��
        pDataArr(0) = BmpData.scan0                                                     'ָ��ͼ�����ڴ��е��׵�ַ
        PixelAddBytes = IIf(PixelFormat = PixelFormat32bppRGB, 4, 3)                    'ÿ��������ռ�õ��ֽ���
        LineAddBytes = BmpData.Stride - BmpData.Width * PixelAddBytes                   'ÿ��ɨ�����ж�����ֽ���������Ҫ�����
        For Y = 1 To BmpData.Height                                                     '���ϵ���ɨ��
            For X = 1 To BmpData.Width                                                  '������ɨ��
                DataArr(0) = 255 - DataArr(0)                                           '������㷨
                DataArr(1) = 255 - DataArr(1)
                DataArr(2) = 255 - DataArr(2)
                
                If DataArr(0) = 0 And DataArr(1) = 0 And DataArr(2) = 0 Then
                    DataArr(0) = 255
                    DataArr(1) = 255
                    DataArr(2) = 255
                End If
                 
                pDataArr(0) = pDataArr(0) + PixelAddBytes                               'ָ����λ
            Next
            pDataArr(0) = pDataArr(0) + LineAddBytes                                    'һ����һ��ɨ���е���ʼλ��
        Next
        FreePoint VarPtrArray(DataArr), VarPtrArray(pDataArr), OldArrPtr, OldpArrPtr    '�ͷ�ģ��ָ��
        GdipBitmapUnlockBits Image, BmpData                                               '��������
    Case PixelFormat8bppIndexed, PixelFormat4bppIndexed, PixelFormat1bppIndexed
        Dim Palette         As ColorPalette
        Dim PaletteSize     As Long
        GdipGetImagePaletteSize Image, PaletteSize
        GdipGetImagePalette Image, Palette, PaletteSize
        For Y = 0 To (PaletteSize - 8) / 4 - 1
            Palette.Entries(Y).Blue = 255 - Palette.Entries(Y).Blue
            Palette.Entries(Y).Green = 255 - Palette.Entries(Y).Green
            Palette.Entries(Y).Red = 255 - Palette.Entries(Y).Red
            
            With Palette.Entries(Y)
            If .Blue = 255 And .Green = 255 And .Red = 255 Then
                .Blue = 0
                .Green = 0
                .Red = 0
            End If
            End With
        Next
        GdipSetImagePalette Image, Palette
    End Select
End Function
