Attribute VB_Name = "mPublic"
'#########################################################################
'##ģ �� ����mPublic.bas
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��3��25��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ������������������
'##��    ����
'#########################################################################

Option Explicit

Public Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'���viewer�е���ʾ������ͼ���һ��ͼ��

    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim lngWhiteX As Long           '��ͼ���ɫ�ĳɰ�ɫ��X���
    Dim lngWhiteY As Long           '��ͼ���ɫ�ĳɰ�ɫ��Y�߶�
    
    If AssembleViewer.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If

    '������ͼ��Ŀ�Ⱥ͸߶�

    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '������ͼ��Ŀ�Ⱥ͸߶�

    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������

    '����ͼ����¿��
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '������������ͼ������
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '����ͼ��Ŀ�ߣ����ܴ������ֵ
    '�������intMaxWidth��intMaxHeight�򣬰���ͼ���ܳ���ȣ�ʹ��С�ڵ���intMaxWidth��intMaxHeight��Ϊ�¿��,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '�ɼ�ͼ��
    '��ͼ��ɼ�����ʱͼ��
    For i = 1 To AssembleViewer.Count
        '�������ű��� hj�޸�,�����ͼ�ϲ�ʱ���Ŵ��ͼ���޷������Ŵ������
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '�ɼ�ͼ��
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
End Function
Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MAXROWS As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    '-----------------------------------------------------------
    '���ܣ� ������Ҫ��ʾ��ͼ����������ʾ���򣬼������ʾͼ�����������
    '������ PicCount-ͼ������
    '       RegionWidth,RegionHeight-�����ȸ߶�
    '       Rows,Cols-�����Զ����е�������
    '-----------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    
    Err = 0: On Error GoTo LL
    
    
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    If iCols = 0 Then iCols = 1
    If iRows = 0 Then iRows = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MAXROWS > 0 And iRows > MAXROWS Then
        iRows = MAXROWS
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MAXROWS > 0 And iRows > MAXROWS Then iRows = MAXROWS
    
    If iRows = 1 And iCols <> ImageCount Then
        iCols = ImageCount
    ElseIf iCols = 1 And iRows <> ImageCount Then
        iRows = ImageCount
    End If
    
    Rows = iRows: Cols = iCols

LL:
End Sub
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    Rpad = zl9ComLib.zlStr.Rpad(strCode, lngLen, strChar, True)
End Function
Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �ϴ����� From zlDataMove Where ϵͳ=[1] And ���=1 And �ϴ����� is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '�ϴ�����û��ʱ��,"<"�ж���ת��������һ��
        If vDate < rsTmp!�ϴ����� Then
            MovedByDate = True
        End If
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetNumber(ByVal strInput As String) As Double
Dim i As Integer
    For i = 1 To Len(strInput)
        If i > Len(strInput) Then Exit For
        If Not IsNumeric(Mid(strInput, i, 1)) Then
            strInput = Replace(strInput, Mid(strInput, i, 1), "")
            If i > Len(strInput) Then Exit For
        End If
    Next
    GetNumber = Val(strInput)
End Function
Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function
Public Function CheckLen(txt As TextBox, intLen As Integer, Optional txtName As String) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(IIf(txtName = "", txt.Name, txtName), 4) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function
Public Function NeedNo(strList As String) As String
    If InStr(strList, "[") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "[") - 1))
    ElseIf InStr(strList, "(") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "(") - 1))
    ElseIf InStr(strList, "-") > 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "-") - 1))
    Else
        NeedNo = LTrim(strList)
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ȡWindowsĬ�ϵ���ʱ�ļ���
'##
'## ������  TmpFilePrefix   :��ʱ�ļ��ĺ�׺��
'################################################################################################################
Public Function GetTempName(TmpFilePrefix As String) As String
     Dim TempFileName As String * 256
     Dim X As Long
     Dim DriveName As String
     DriveName = "c:"   'Ĭ��ȡC��ΪĿ���̷�
       X = GetTempFileName(DriveName, TmpFilePrefix, 0, TempFileName)
       GetTempName = Left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)
End Function

'################################################################################################################
'## ���ܣ�  ����͸��ͼƬ��ָ��HDC�ϣ�ָ��͸��ɫ��
'##
'## ������  hDCDest         :Ŀ���ͼ����
'##         (xDest,yDest)   :���Ͻ�λ��
'##         (Width,Height)  :��ͼ����߶ȡ����
'##         picSource       :ԴͼƬ
'##         (XSrc,YSrc)     :ԴͼƬƫ��λ��
'##         clrMask         :͸��ɫ(MaskColor)
'##         hPal            :��ɫ��������ѡ
'##
'## �÷���  PaintTransparentStdPic UserControl.hdc, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
'################################################################################################################
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

'################################################################################################################
'## ���ܣ�  ����͸��ͼƬ��ָ��DC�ϣ�ָ��͸��ɫ��
'##
'## ˵����  ���ڵ� PaintTransparentStdPic() �������á�
'################################################################################################################
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

'################################################################################################################
'## ���ܣ�  �ж��Ƿ�Ϊ�༭��
'##
'## ������  KeyAscii        :��ǰ�༭��ʽ��
'##         AllowSubtract   :Insert���Ƿ���Ϊ�༭������ѡ
'##
'## ���أ�  ����Ǳ༭�����򷵻� True�����򣬷��� False
'################################################################################################################
Public Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ؼ���ָ����������Ļ�е�λ��
'##
'## ������  lngHwnd         :�ؼ���� hWnd
'##         (lngX,lngY)     :�ؼ��е�����λ��
'##
'## ���أ�  ���ؿؼ��е���������Ļ�е�λ�ã���λ���
'################################################################################################################
Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'################################################################################################################
'## ���ܣ�  ��VB��ϵͳ��ɫת��ΪRGBɫ
'##
'## ������  lngColor        :��Ҫת����ϵͳ��ɫ(COLORREF)
'##
'## ���أ�  ����ת�����RGB��ɫ
'################################################################################################################
Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

'################################################################################################################
'## ���ܣ�  �õ�ָ�����ֵ������ֺ�
'##
'## ������  lngNum      :�����С�����֣�
'################################################################################################################
Public Function GetFontSizeChinese(sngNum As Single) As String
    Dim lngNum As Single
    lngNum = Format(sngNum, "0.0")
    Select Case lngNum
    Case 42
        GetFontSizeChinese = "����"
    Case 36
        GetFontSizeChinese = "С��"
    Case 26
        GetFontSizeChinese = "һ��"
    Case 24
        GetFontSizeChinese = "Сһ"
    Case 22
        GetFontSizeChinese = "����"
    Case 18
        GetFontSizeChinese = "С��"
    Case 16
        GetFontSizeChinese = "����"
    Case 15
        GetFontSizeChinese = "С��"
    Case 14
        GetFontSizeChinese = "�ĺ�"
    Case 12
        GetFontSizeChinese = "С��"
    Case 10.5
        GetFontSizeChinese = "���"
    Case 9
        GetFontSizeChinese = "С��"
    Case 7.5
        GetFontSizeChinese = "����"
    Case 6.5
        GetFontSizeChinese = "С��"
    Case 5.5
        GetFontSizeChinese = "�ߺ�"
    Case 5
        GetFontSizeChinese = "�˺�"
    Case 0
        GetFontSizeChinese = ""
    Case Else
        GetFontSizeChinese = lngNum
    End Select
End Function
Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
End Function
'################################################################################################################
'## ���ܣ�  �õ�ָ�������ֺŵ����ֳߴ�
'##
'## ������  strSize     :�����ֺ�
'################################################################################################################
Public Function GetFontSizeNumber(strSize As String) As Single
    Dim sngNum As Single
    Select Case strSize
    Case "����"
        sngNum = 42
    Case "С��"
        sngNum = 36
    Case "һ��"
        sngNum = 26
    Case "Сһ"
        sngNum = 24
    Case "����"
        sngNum = 22
    Case "С��"
        sngNum = 18
    Case "����"
        sngNum = 16
    Case "С��"
        sngNum = 15
    Case "�ĺ�"
        sngNum = 14
    Case "С��"
        sngNum = 12
    Case "���"
        sngNum = 10.5
    Case "С��"
        sngNum = 9
    Case "����"
        sngNum = 7.5
    Case "С��"
        sngNum = 6.5
    Case "�ߺ�"
        sngNum = 5.5
    Case "�˺�"
        sngNum = 5
    Case Else
        sngNum = IIf(Val(strSize) <= 0, 10, Val(strSize))
    End Select
    GetFontSizeNumber = Format(sngNum, "0.0")
End Function

'################################################################################################################
'## ���ܣ�  ����׼ stdPicture ͼƬת��Ϊ Meta ͼԪ�ļ�
'##
'## ������  aStdPic         :��ת���ı�׼ͼƬ
'##         strDestFileName :ת�����Ŀ�� Meta ͼԪ�ļ���
'################################################################################################################
Public Sub StdPicToMetaFile(aStdPic As StdPicture, strDestFileName As String)
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As BitMap
    Dim aSize       As Size
    Dim aPt         As POINTAPI
    Dim Filename    As String
'    Dim aMetaHdr    As METAHEADER
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim bytes()     As Byte
    Dim FileNum     As Integer

    ' Create a metafile to a temporary file in the registered windows TEMP folder
    Filename = GetTempName("WMF")
    hMetaDC = CreateMetaFile(Filename)

    ' Set the map mode to MM_ANISOTROPIC
    SetMapMode hMetaDC, 8    'MM_ANISOTROPIC
    ' Set the metafile origin as 0, 0
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    ' Get the bitmap's dimensions
    GetObject aStdPic.Handle, Len(aBMP), aBMP
    ' Set the metafile width and height
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    ' save the new dimensions
    SaveDC hMetaDC
    ' OK. Now transfer the freakin image to the metafile
    screenDC = GetDC(0)
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

    gobjFSO.CopyFile Filename, strDestFileName, True
    If gobjFSO.FileExists(Filename) Then gobjFSO.DeleteFile Filename, True
End Sub

'################################################################################################################
'## ���ܣ�  ��Ӱ�ť
'################################################################################################################
Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, ID, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.STYLE = ButtonStyle
    Control.Category = Category

    Set AddButton = Control
End Function


'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo errHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '�����и���
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '�����и���
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

errHand:
    zlReportToVSFlexGrid = False
End Function

Public Sub ValidControlText(ByRef txtInput As Object)
    On Error Resume Next
    '�޳��ؼ����ݵ������ַ�'��%
    Dim strSource As String, i As Long, j As Long, k As Long
    Dim strDest As String, lngLen As Long
    Dim lngSelStart As Long, lngSelStart2 As Long
    strSource = txtInput.Text
    lngSelStart = txtInput.SelStart
    lngLen = Len(strSource)
    
    For i = 1 To lngLen
        If Mid(strSource, i, 1) <> "'" And Mid(strSource, i, 1) <> "%" Then
            strDest = strDest & Mid(strSource, i, 1)
            j = j + 1
        End If
        If i = lngSelStart Then lngSelStart2 = j
    Next
    txtInput.Text = strDest
    txtInput.SelStart = lngSelStart2
End Sub
'################################################################################################################
'## ���ܣ�  ��ȡһ���ڵ��ֵ
'##
'## ������  CurNode         :   ��ǰ�ڵ����
'##         SubNodeName     :   �ӽڵ�����
'##         DefaultValue    :   Ĭ��ֵ
'################################################################################################################
Public Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then 'ֻ�лس����޻��з�
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then 'ֻ�л��з��޻س���
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## ���ܣ�  ����һ��XML�ڵ㲢��ֵ
'##
'## ������  TabNumber   :   �������������ʾ�ж��ٸ�Tab�Ʊ���������Ķ���
'##         Parent      :   ���ڵ�
'##         Node_Type   :   �ڵ����ͣ�Ŀǰ֧�� NODE_ELEMENT ��NODE_CDATA_SECTION ��NODE_COMMENT ��NODE_ATTRIBUTE�ȣ�
'##         Node_Name   :   �ڵ�����
'##         Node_Value  :   �ڵ��ı�
'################################################################################################################
Public Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '�ַ�����ֵ���ã���Ӱ�����ݣ���ֻӰ���Ķ����۶�
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '�����ı��ڵ�
    '�����½ڵ�
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '�����ı�ֵ
    New_Node.Text = Node_Value
    '��ӵ����ڵ�
    Parent.appendChild New_Node
    '���ĩβ�س�����Ӱ�����ݣ���ֻӰ���Ķ����۶�
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '�����ı��ڵ�
    Set CreateNode = New_Node
End Function


'################################################################################################################
'## ���ܣ�  '���ƴ����С���رհ�ť�Ĺ���
'##
'## ������   blnEnable   :   true ���ã�false ������
'################################################################################################################
Public Sub EnableControlBar(ByRef FormObj As Object, ByVal blnEnable As Boolean)
    Dim hSysMenu  As Long, nCnt  As Long
    hSysMenu = GetSystemMenu(FormObj.hwnd, blnEnable)
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 3, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 4, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 5, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar FormObj.hwnd
        End If
    End If
End Sub

Public Sub SetFontSize(ByRef frmObj As Object, ByVal bytFontSize As Byte)
Dim CtlFont As StdFont, objCtrl As Control
    On Error Resume Next
    If bytFontSize = 0 Then bytFontSize = 9
    frmObj.FontSize = bytFontSize
    For Each objCtrl In frmObj.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = frmObj.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
            objCtrl.PaintManager.Layout = xtpTabLayoutAutoSize
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PanelPaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = frmObj.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PanelPaintManager.Font = CtlFont
            objCtrl.PanelPaintManager.Layout = xtpTabLayoutAutoSize
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = frmObj.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("VSFlexGrid")
            If objCtrl.Cols > 2 Then
                Call zlControl.VSFSetFontSize(objCtrl, bytFontSize, 3)
            Else
                Call zlControl.VSFSetFontSize(objCtrl, bytFontSize, 0)
            End If
        Case UCase("ReportControL")
            Set CtlFont = objCtrl.PaintManager.TextFont
            If CtlFont Is Nothing Then
                Set CtlFont = frmObj.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
        End Select
    Next
End Sub
Public Function DynamicCreate(ByVal strClass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

'################################################################################################################
'##  Ԥ������Ƕ�ؼ��ֳ�ʼ��
'################################################################################################################
Public Sub InitPreDefinedKeys()
    gKeyWords(1).KeyStart = "OS"
    gKeyWords(1).KeyEnd = "OE"
    gKeyWords(2).KeyStart = "PS"
    gKeyWords(2).KeyEnd = "PE"
    gKeyWords(3).KeyStart = "ES"
    gKeyWords(3).KeyEnd = "EE"
    gKeyWords(4).KeyStart = "TS"
    gKeyWords(4).KeyEnd = "TE"
    gKeyWords(5).KeyStart = "SS"
    gKeyWords(5).KeyEnd = "SE"
    gKeyWords(6).KeyStart = "DS"
    gKeyWords(6).KeyEnd = "DE"
End Sub
Public Sub GetUserInfo()
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
        
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount <> 0 Then
            gstrDBUser = .Fields("�û���").Value
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = NVL(.Fields("����").Value, "")  '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("������").Value        '��ǰ�û�
            gstrDeptName = .Fields("������").Value        '��ǰ�û�
        Else
            gstrDBUser = ""
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
    End With
    
    gstrSQL = "Select ǩ�� From ��Ա�� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ������", glngUserId)
    If Not rsTemp.EOF Then
        gstrSignName = NVL(rsTemp!ǩ��, gstrUserName)
    End If
   
   
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    Err = 0: On Error GoTo errHand
    strSQL = "Select ������ From Zlsystems Where ��� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!������
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
errHand:
    Substr = ""
End Function

Public Function GUID() As String
    GUID = Replace(Replace(Replace(Left(CreateObject("Scriptlet.TypeLib").GUID, 38), "-", ""), "{", ""), "}", "")
End Function

Public Function GetCurrentGdi() As Long
    If glngPro = 0 Then
        glngPro = GetCurrentProcess
    End If
    
    GetCurrentGdi = GetGuiResources(glngPro, 0)
End Function
