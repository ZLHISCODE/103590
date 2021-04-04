Attribute VB_Name = "mPublic"
'######################################################################################
'##ģ �� ����mPublic.bas
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��4��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ���������������߳�������
'##��    ����
'######################################################################################

Option Explicit

'######################################################################################
'   ֽ����������
'######################################################################################
Public PaperKindConst(1 To 42) As String    '�����ơ��߶ȡ���ȡ���С�߾�(��������)����Ӧ��ӡֽ�����е�ֽ�����ೣ��

Public gTargetDC As Long
Public sngX As Single, sngY As Single   '�������
Public intShift As Integer              '��갴��
Public bWay As Boolean                  '��귽��
Public bMouseFlag As Boolean            '����¼������־

'#########################################################################################################
'## ���ܣ�  ����ָ����ֵ���߼�����߶�
'## ������  hDC:            Ŀ���豸���
'##         lPointValue:    �����ֵ
'## ���أ�  �����߼�����߶�
'#########################################################################################################
Public Function GetPixcelHeightByPoint(hDC As Long, ByVal lPointValue As Double) As Double
    GetPixcelHeightByPoint = -MulDiv((lPointValue), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
End Function
       
'#########################################################################################################
'## ���ܣ�  ����ָ����ֵ���߼�������
'## ������  hDC:            Ŀ���豸���
'##         lPointValue:    �����ֵ
'## ���أ�  �����߼�����߶�
'#########################################################################################################
Public Function GetPixcelWidthByPoint(hDC As Long, ByVal lPointValue As Double) As Double
    GetPixcelWidthByPoint = -MulDiv((lPointValue), (GetDeviceCaps(hDC, LOGPIXELSX)), 72)
End Function

Public Sub PressKey(bytKey As Byte)
    '���ܣ�����̷���һ����,����SendKey
    '������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
   
Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
    '����:���������뷨����ر����뷨
    '     ����zlComlib��ͬ���������ģ�������ZLHIS����еı��ز�������
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    Dim strUser As String
    
    strUser = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "USER", "")
    '�û�û�������ã��Ͳ�����
    strIme = GetSetting("ZLSOFT", "˽��ȫ��\" & strUser, "���뷨", "")
    If strIme = "" And blnOpen = True Then Exit Function                 'Ҫ������뷨��������û������
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ��������뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                    Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '�������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function
   
'############################################################################################################
'## ���ܣ�  RTF��������������ʾ�����ӡ���һ�£�
'##
'## ������  RTF         ��RTF�ؼ�
'##         MarginLeft  ����߾�
'##         MarginRight ���ұ߾�
'##         MarginTop   ���ϱ߾�
'##         MarginBottom���±߾�
'##         PaperWidth  ��ҳ��
'##         PaperHeight ��ҳ��
'##
'## ˵����  ��������ĻΪ������׼��������
'############################################################################################################
Public Sub WYSIWYG_RTF(ByRef RTF As RichTextBox, _
    ByVal MarginLeft As Long, _
    ByVal MarginRight As Long, _
    ByVal MarginTop As Long, _
    ByVal MarginBottom As Long, _
    ByVal PaperWidth As Long, _
    ByVal PaperHeight As Long)
    
    Dim lngOffsetLeft As Long   '���ƫ����
    Dim lngMarginLeft As Long   '��߾�
    Dim r As Long               '����ֵ
    PaperWidth = PaperWidth - MarginLeft - MarginRight                      '����ɴ�ӡ���ֿ��
    r = SendMessage(RTF.Hwnd, EM_SETTARGETDEVICE, gTargetDC, ByVal PaperWidth)     '�ı��п�ִ����Ⱦ
End Sub

'######################################################################################
'   ���Ʋ�ɫ������
'######################################################################################

Public Sub DrawProgress( _
      lPercent As Single, _
      ByVal lHDC As Long, _
      ByVal lLeft As Long, ByVal lTOp As Long, _
      ByVal lRight As Long, ByVal lBottom As Long _
   )
Dim hBr As Long
Dim tR As RECT
Dim tProgR As RECT

   tR.Left = lLeft + 1
   tR.Top = lTOp + 1
   tR.Right = lRight - 1
   tR.Bottom = lBottom - 1

   ' Draw the progress bar
   LSet tProgR = tR
   tProgR.Right = tProgR.Left + (tProgR.Right - tProgR.Left) * lPercent
   GradientFillRect lHDC, tProgR, RGB(234, 94, 45), RGB(238, 164, 36), GRADIENT_FILL_RECT_H
   
   ' Draw the text in front of the progress bar
'   DrawTextA lHDC, Format(lPercent, "0%"), -1, tR, DT_CENTER

   ' Frame the progress bar:
   hBr = CreateSolidBrush(&H0&)
   FrameRect lHDC, tR, hBr
   DeleteObject hBr
End Sub

Public Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).X = tR.Left
   tTV(0).Y = tR.Top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).X = tR.Right
   tTV(1).Y = tR.Bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   End If
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub

Public Function GetCharPosFromByteValue(ByVal S As String, ByVal Pos As Long) As Long
'��֪�ַ������ֽڳ��ȣ������ַ����ȣ�
    Dim iLoop As Long
    Dim iChinese As Long
    iChinese = 0
    For iLoop = Pos To 1 Step -1
        If Asc(StrConv(MidB(StrConv(S, vbFromUnicode), iLoop, 1), vbUnicode)) = 0 Then
            iChinese = iChinese + 1
        End If
    Next iLoop
    GetCharPosFromByteValue = Pos - iChinese \ 2
End Function


Public Function GetTempName(TmpFilePrefix As String) As String
'��ȡWIndows��ʱĿ¼
     Dim TempFileName As String * 256
     Dim X As Long
     Dim DriveName As String
     DriveName = "c:"
       X = GetTempFileName(DriveName, TmpFilePrefix, 0, TempFileName)
       GetTempName = Left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)
End Function


Public Function StdPicAsRTF(aStdPic As StdPicture, lWidth As Long, lHeight As Long) As String
'��ȡͼƬRTF�ַ�������������߶ȡ����

    ' ***********************************************************************
    '  Author: The Hand
    '    Date: June, 2002
    ' Company: EliteVB
    '
    '  Function: StdPicAsRTF
    ' Arguments: aStdPic - Any standard picture object from memory, a
    '                      picturebox, or other source.
    '
    ' Description:
    '    Embeds a standard picture object in a windows metafile and returns
    '    rich text format code (RTF) so it can be placed in a RichTextBox.
    '    Useful for emoticons in chat programs, pics, etc. Currently does
    '    not support icon files, but that is easy enough to add in.
    ' ***********************************************************************
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As BITMAP
    Dim aSize       As SIZE
    Dim aPt         As POINTAPI
    Dim FileName    As String
'    Dim aMetaHdr    As METAHEADER
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim byteStr     As String
    Dim bytes()     As Byte
    Dim filenum     As Integer
    Dim numBytes    As Long
    Dim i           As Long
    
    ' Create a metafile to a temporary file in the registered windows TEMP folder
    FileName = GetTempName("WMF")
    hMetaDC = CreateMetaFile(FileName)
    
    ' Set the map mode to MM_ANISOTROPIC
    SetMapMode hMetaDC, MM_ANISOTROPIC
    ' Set the metafile origin as 0, 0
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    ' Get the bitmap's dimensions
    GetObject aStdPic.Handle, Len(aBMP), aBMP
    ' Set the metafile width and height
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    ' save the new dimensions
    SaveDC hMetaDC
    ' OK. Now transfer the freakin image to the metafile
    screenDC = GetDC(aStdPic.Handle) 'GetDC(0)
    hPicDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    hOldBmp = SelectObject(hPicDC, aStdPic.Handle)
    BitBlt hMetaDC, 0, 0, aBMP.bmWidth, aBMP.bmHeight, hPicDC, 0, 0, vbSrcCopy
    SelectObject hPicDC, hOldBmp
    DeleteDC hPicDC
    DeleteObject hOldBmp
    ' "redraw" the metafile DC
    RestoreDC hMetaDC, True
    ' close it and get the metafile handle
    hMeta = CloseMetaFile(hMetaDC)
    
'    GetObject hMeta, Len(aMetaHdr), aMetaHdr
    ' delete it from memory
    DeleteMetaFile hMeta
    
    ' Do the RTF header for the object. This little bit is sometimes required on
    '  earlier versions of the rich text box and in certain operating systems
    '  (WinNT springs to mind)
    headerStr = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052\uc1 "
    ' Picture specific tag stuff
    
    If lWidth <= 0 Then lWidth = aBMP.bmWidth * Screen.TwipsPerPixelX
    If lHeight <= 0 Then lHeight = aBMP.bmHeight * Screen.TwipsPerPixelY
    
    headerStr = headerStr & _
                "{\pict\picscalex100\picscaley100" & _
                "\picw" & aStdPic.Width & "\pich" & aStdPic.Height & _
                "\picwgoal" & lWidth & _
                "\pichgoal" & lHeight & _
                "\wmetafile8"

'    lWidth = aBMP.bmWidth * Screen.TwipsPerPixelX
'    lHeight = aBMP.bmHeight * Screen.TwipsPerPixelY
    
    ' Get the size of the metafile
    numBytes = FileLen(FileName)
    ' Create our byte buffer for reading
    ReDim bytes(1 To numBytes)
    ' get a free file number
    filenum = FreeFile()
    ' open the file for input
    Open FileName For Binary Access Read As #filenum
    ' read the bytes
    Get #filenum, , bytes
    ' close the file
    Close #filenum
    ' Generate our hex encoded byte string
    byteStr = String(numBytes * 2, "0")
    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) > &HF Then
            Mid$(byteStr, 1 + (i - 1) * 2, 2) = Hex$(bytes(i))
        Else
            Mid$(byteStr, 2 + (i - 1) * 2, 1) = Hex$(bytes(i))
        End If
    Next i
    ' stick it all together
    retStr = headerStr & " " & byteStr & "}"
    ' Add in the closing RTF bit
    retStr = retStr & "}"
        
    StdPicAsRTF = retStr
    On Local Error Resume Next
    ' Kill the temporary file
    If Dir(FileName) <> "" Then Kill FileName

End Function

'###############################################################################################
'   ����͸��ͼƬ��ָ��HDC�ϣ�ָ��͸��ɫ����
'###############################################################################################

Public Sub PaintTransparentStdPic(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle
            'Draw Transparent image
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub

PaintTransparentStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

'Provided with comments by Microsoft
Public Sub PaintTransparentDC(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hBmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long

    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        'Create halftone palette
        hPal = CreateHalftonePalette(hdcScreen)
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy

    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hBmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hBmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    DeleteObject hPal
End Sub
Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'���ܣ���ItemData����ComboBox������ֵ
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.itemData(i) = lngData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function
Public Sub SetSelection(lHwnd As Long, ByVal lStart As Long, ByVal lEnd As Long)
    Dim tCR As CHARRANGE
    tCR.cpMin = lStart
    tCR.cpMax = lEnd
    SendMessage lHwnd, EM_EXSETSEL, 0, tCR
End Sub

'######################################################################################
'## ����ת��
'######################################################################################

Public Function J2F(ByVal strText As String) As String
    '����ת����
    Dim strF As String      '�����ַ���
    Dim strJ As String      '�����ַ���
    Dim STlen As Long       '��ת���ִ�����
    
    strJ = strText
    STlen = lstrlen(strJ)
    strF = Space(STlen)
    LCMapString &H804, &H4000000, strJ, STlen, strF, STlen
    J2F = strF
End Function

Public Function F2J(ByVal strText As String) As String
    '����ת����
    Dim strF As String      '�����ַ���
    Dim strJ As String      '�����ַ���
    Dim STlen As Long       '��ת���ִ�����
    strF = strText
    STlen = lstrlen(strF)
    strJ = Space(STlen)
    LCMapString &H804, &H2000000, strF, STlen, strJ, STlen
    F2J = strJ
End Function


