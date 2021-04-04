Attribute VB_Name = "mdlPicFormat"
Option Explicit

Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type
  
Private Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type
  
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    left   As Integer
    top    As Integer
    Right  As Integer
    Bottom As Integer
End Type
  
Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type
  
' GDI and GDI+ constants
Private Const PLANES As Long = 14             '  Number of planes
Private Const BITSPIXEL As Long = 12          '  Number of bits per pixel
Private Const PATCOPY  As Long = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP As Long = 1      ' Bitmap type
Private Const InterpolationModeHighQualityBicubic As Long = 7
Private Const GDIP_WMF_PLACEABLEKEY As Long = &H9AC6CDD7
Private Const UnitPixel                  As Long = 2
Private Const EncoderQuality             As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID(0 To 3)        As Long
    NumberOfValues      As Long
    Type                As EncoderParameterValueType
    Value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Public Enum ImageFileFormat
    BMP = 1
    JPG = 2
    PNG = 3
    GIF = 4
End Enum

' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
  
' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, Optional GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal size As Long, Encoders As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal Bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As Long
 

Public Function InitGDIPlus() As Long
' Initialises GDI Plus
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput
      
    gdipInit.GdiplusVersion = 1    'GDI+ 1.0�汾
    GdiplusStartup Token, gdipInit, ByVal 0&    '��ʼ��GDI+
    InitGDIPlus = Token
End Function
  
Public Sub FreeGDIPlus(Token As Long)
' Frees GDI Plus
    GdiplusShutdown Token
End Sub
  

Public Function LoadPictureGDIPlus(strPicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, _
    Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
'����:��PNG�ļ����ص�VB��Picture����ʾ
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    Dim Token As Long
    
    Token = InitGDIPlus
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(strPicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & strPicFile
        Exit Function
    End If
      
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
      
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height
  
    ' Resize the picture
    gdipResize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
      
    ' Get the bitmap back
    GetBitmap hDC, hBitmap
  
    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    
    Call FreeGDIPlus(Token)
End Function
  

Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
' Initialises the hDC to draw
    Dim hBrush As Long
          
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub
  
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
' Resize the picture using GDI plus
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
      
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
      
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
          
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
          
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
'        DestX = (Width - DestWidth) / 2
'        DestY = (Height - DestHeight) / 2
  
        DestX = 0
        DestY = 0
  
        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub

Private Sub GetBitmap(hDC As Long, hBitmap As Long)
' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

Private Function CreatePicture(hBitmap As Long) As IPicture
' Creates a Picture Object from a handle to a bitmap
    Dim IID_IDispatch As GUID
    Dim Pic           As PICTDESC
    Dim IPic          As IPicture
      
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
          
    ' Fill Pic with necessary parts
    Pic.size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap
  
    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function
  
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, Width As Long, Height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
' Returns a resized version of the picture
    Dim Img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
      
    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, Img
    End Select
      
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, BackColor, Width, Height
        gdipResize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
End Function
  
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
' Fills in the wmfPlacable header
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub

'--------------------------------------------------------------------------------------------------------------------------------

Public Function SaveStdPicToFile(stdPic As StdPicture, ByVal strFileName As String, Optional ByVal FileFormat As ImageFileFormat = JPG, _
    Optional ByVal JpgQuality As Long = 80, Optional Resolution As Single) As Boolean
'����:��VB��ͼƬ��ʽת����ָ�����͵��ļ���ʽ�洢JPG,PNG,GIF,BMP
    Dim CLSID(3)        As Long
    Dim Bitmap          As Long
    Dim Token           As Long
    '��ʼ��GDI+
    Token = InitGDIPlus
    
    GdipCreateBitmapFromHBITMAP stdPic.Handle, stdPic.hPal, Bitmap
    
    If Bitmap <> 0 Then
        '˵�����ǳɹ��Ľ�StdPic����ת��ΪGDI+��Bitmap������
        GdipBitmapSetResolution Bitmap, Resolution, Resolution
        Select Case FileFormat
        Case ImageFileFormat.BMP
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(strFileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.JPG                    'JPG��ʽ�������ñ��������
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                        ' �����Զ���ı������������Ϊ1������
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                   ' ���ò���ֵ����������Ϊ������
                    Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))  ' ���ò���Ψһ��־��GUID������Ϊ����Ʒ��
                    .Value = VarPtr(JpgQuality)                                ' ���ò�����ֵ��Ʒ�ʵȼ������Ϊ100��ͼ���ļ���С��Ʒ�ʳ�����
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(strFileName), CLSID(0), aEncParams(1)) = 0)
            End If
        Case ImageFileFormat.PNG
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(strFileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.GIF
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '���ԭʼ��ͼ����24λ����������������ϵͳ�ĵ�ɫ������ͼ��ת��Ϊ8λ��ת����Ч���᲻������,��Ҳ�п���ϵͳ���Զ�ת��������ʧ��
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(strFileName), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap      'ע���ͷ���Դ
    '�ر�GDI+��
    Call FreeGDIPlus(Token)
    
End Function


Private Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, size               '�õ�����������Ĵ�С
    If size <> 0 Then
       ReDim Info(1 To Num) As ImageCodecInfo       '�����鶯̬�����ڴ�
       ReDim Buffer(1 To size) As Byte
       GdipGetImageEncoders Num, size, Buffer(1)            '�õ�������ַ�����
       CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)     '������ͷ
       For i = 1 To Num             'ѭ��������н���
           If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then         '�����ָ��ת���ɿ��õ��ַ�
               CopyMemory ClassID(0), Info(i).ClassID(0), 16  '�������ID
               GetEncoderClsID = i      '���سɹ�������ֵ
               Exit For
           End If
       Next
    End If
End Function

Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function


