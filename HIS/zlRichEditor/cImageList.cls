VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##模 块 名：cImageList.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：继承并扩展的图像列表控件。用于HRuler控件。
'##版    本：
'######################################################################################

Option Explicit

' -----------
' ENUMS
' -----------

Public Enum eilColourDepth
    ILC_COLOR = &H0
    ILC_COLOR4 = &H4
    ILC_COLOR8 = &H8
    ILC_COLOR16 = &H10
    ILC_COLOR24 = &H18
    ILC_COLOR32 = &H20
End Enum

Public Enum eilSwapTypes
   eilCopy = ILCF_MOVE
   eilSwap = ILCF_SWAP
End Enum

' ------------------
' Private variables:
' ------------------
Private m_hIml As Long
Private m_lIconSizeX As Long
Private m_lIconSizeY As Long
Private m_eColourDepth As eilColourDepth
Private m_sKey() As String
Private m_HDC As Long

Public Property Let OwnerHDC(ByVal lHDC As Long)
   m_HDC = lHDC
End Property

Public Property Get SystemColourDepth() As eilColourDepth
Dim lR As Long
Dim lHDC As Long
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lR = GetDeviceCaps(lHDC, BITSPIXEL)
   DeleteDC lHDC
   SystemColourDepth = lR
End Property

Public Sub SwapOrCopyImage( _
      ByVal vKeySrc As Variant, _
      ByVal vKeyDst As Variant, _
      Optional ByVal eSwap As eilSwapTypes = eilSwap _
   )
Dim lDst As Long
Dim lSrc As Long
Dim sKeyDst As String
Dim sKeySrc As String

   If (m_hIml <> 0) Then
      lDst = ItemIndex(vKeySrc)
      If (lDst > -1) Then
         lSrc = ItemIndex(vKeyDst)
         If (lSrc > -1) Then
            ImageList_Copy m_hIml, lDst, m_hIml, lSrc, eSwap
            sKeyDst = m_sKey(lDst)
            sKeySrc = m_sKey(lSrc)
            m_sKey(lDst) = sKeySrc
            m_sKey(lSrc) = sKeyDst
         End If
      End If
   End If
End Sub

Public Function Create() As Boolean
     
     ' Do we already have an image list?  Kill it if we have:
    Destroy

    'Create the Imagelist:
    m_hIml = ImageList_Create(m_lIconSizeX, m_lIconSizeY, ILC_MASK Or m_eColourDepth, 4, 4)
    If (m_hIml <> 0) And (m_hIml <> -1) Then
      ' Ok
      Create = True
    Else
      m_hIml = 0
    End If
    
End Function
Public Sub Destroy()
   ' Kill the image list if we have one:
   If (hIml <> 0) Then
      ImageList_Destroy hIml
      m_hIml = 0
   End If
   Erase m_sKey
End Sub
Public Sub DrawImage( _
        ByVal vKey As Variant, _
        ByVal hDC As Long, _
        ByVal xPixels As Integer, _
        ByVal yPixels As Integer, _
        Optional ByVal bSelected = False, _
        Optional ByVal bCut = False, _
        Optional ByVal bDisabled = False, _
        Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
        Optional ByVal hExternalIml As Long = 0 _
    )
Dim hIcon As Long
Dim lFlags As Long
Dim lhIml As Long
Dim lColor As Long
Dim iImgIndex As Long

   ' Draw the image at 1 based index or key supplied in vKey.
   ' on the hDC at xPixels,yPixels with the supplied options.
   ' You can even draw an ImageList from another ImageList control
   ' if you supply the handle to hExternalIml with this function.
   
   iImgIndex = ItemIndex(vKey)
   If (iImgIndex > -1) Then
      If (hExternalIml <> 0) Then
          lhIml = hExternalIml
      Else
          lhIml = hIml
      End If
      
      lFlags = ILD_Transparent
      If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
      End If
      
      If (bCut) Then
        ' Draw dithered:
        lColor = TranslateColor(oCutDitherColour)
        If (lColor = -1) Then lColor = GetSysColor(COLOR_WINDOW)
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hDC, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, lColor, _
              lFlags
      ElseIf (bDisabled) Then
        ' extract a copy of the icon:
        hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
        ' Draw it disabled at x,y:
        DrawState hDC, 0, 0, hIcon, 0, xPixels, yPixels, m_lIconSizeX, m_lIconSizeY, DST_ICON Or DSS_DISABLED
        ' Clear up the icon:
        DestroyIcon hIcon
              
      Else
        ' Standard draw:
        ImageList_Draw _
            lhIml, _
            iImgIndex, _
            hDC, _
            xPixels, _
            yPixels, _
            lFlags
      End If
   End If
End Sub

Public Property Get IconSizeX() As Long
   ' Returns the icon width
    IconSizeX = m_lIconSizeX
End Property
Public Property Let IconSizeX(ByVal lSizeX As Long)
   ' Sets the icon width.  NB no change at runtime unless you
   ' call Create and add all the images in again.
    m_lIconSizeX = lSizeX
End Property
Public Property Get IconSizeY() As Long
   ' Returns the icon height:
    IconSizeY = m_lIconSizeY
End Property
Public Property Let IconSizeY(ByVal lSizeY As Long)
   ' Sets the icon height.  NB no change at runtime unless you
   ' call Create and add all the images in again.
    m_lIconSizeY = lSizeY
End Property
Public Property Get ColourDepth() As eilColourDepth
   ' Returns the ColourDepth:
    ColourDepth = m_eColourDepth
End Property
Public Property Let ColourDepth(ByVal eDepth As eilColourDepth)
   ' Sets the ColourDepth.  NB no change at runtime unless you
   ' call Create and rebuild the image list.
    m_eColourDepth = eDepth
End Property

Public Property Get ImageCount() As Integer
   ' Returns the number of images in the ImageList:
   If (hIml <> 0) Then
      ImageCount = ImageList_GetImageCount(hIml)
   End If
End Property
Public Sub RemoveImage(ByVal vKey As Variant)
Dim lIndex As Long
Dim i As Long
   ' Removes an image from the ImageList:
   If (hIml <> 0) Then
      lIndex = ItemIndex(vKey)
      ImageList_Remove hIml, lIndex
      ' Fix up the keys:
      For i = lIndex To ImageCount - 1
         m_sKey(i) = m_sKey(i + 1)
      Next i
      pEnsureKeys
   End If

End Sub
Public Property Get KeyExists(ByVal sKey As String) As Boolean
Dim iL As Long
Dim iU As Long
   If ImageCount > 0 Then
      On Error Resume Next
      iU = UBound(m_sKey)
      If Err.Number <> 0 Then
         iU = 0
      End If
      If (iU <> ImageCount - 1) Then
         pEnsureKeys
      End If
      For iL = 0 To ImageCount - 1
         If m_sKey(iL) = sKey Then
            KeyExists = True
            Exit For
         End If
      Next iL
   End If
End Property

Public Property Get ItemIndex(ByVal vKey As Variant) As Long
Dim lR As Long
Dim i As Long
   ' Returns the 0 based Index for the selected
   ' Image list item:
   If (IsNumeric(vKey)) Then
      lR = vKey
      If (lR > 0) And (lR <= ImageCount) Then
         ItemIndex = lR - 1
      Else
         ' error
         Err.Raise 9, App.EXEName & ".vbalImageList"
         ItemIndex = -1
      End If
   Else
      lR = -1
      For i = 0 To ImageCount - 1
         If (m_sKey(i) = vKey) Then
            lR = i
            Exit For
         End If
      Next i
      If (lR > 0) And (lR <= ImageCount) Then
         ItemIndex = lR
      Else
         Err.Raise 9, App.EXEName & ".vbalImageList"
         ItemIndex = -1
      End If
   End If
End Property
Public Property Get ItemKey(ByVal iIndex As Long) As Variant
   ' Returns the Key for an image:
   If (iIndex > 0) And (iIndex <= ImageCount) Then
      ItemKey = m_sKey(iIndex - 1)
   Else
      Err.Raise 9, App.EXEName & ".vbalImageList"
   End If
End Property
Public Property Let ItemKey(ByVal iIndex As Long, ByVal vKey As Variant)
   ' Sets the Key for the an image:
   iIndex = iIndex - 1
   If (iIndex > 0) And (iIndex < ImageCount) Then
      SetKey iIndex, vKey
   Else
      Err.Raise 9, App.EXEName & ".vbalImageList"
   End If
End Property
Public Property Get ItemPicture(ByVal vKey As Variant) As IPicture
Dim lIndex As Long
Dim hIcon As Long
   ' Returns a StdPicture for an image in the ImageList:
   lIndex = ItemIndex(vKey)
   If (lIndex > -1) Then
      hIcon = ImageList_GetIcon(m_hIml, lIndex, ILD_Transparent)
      If (hIcon <> 0) Then
         Set ItemPicture = IconToPicture(hIcon)
         ' Don't destroy the icon - it is now owned by
         ' the picture object
      End If
   End If
   
End Property
Public Property Get ItemCopyOfIcon(ByVal vKey As Variant) As Long
Dim lIndex As Long
   ' Returns a hIcon for an image in the ImageList.  User must
   ' call DestroyIcon on the returned handle.
   lIndex = ItemIndex(vKey)
   If (lIndex > -1) Then
      ItemCopyOfIcon = ImageList_GetIcon(m_hIml, lIndex, ILD_Transparent)
   End If
End Property
Public Sub Clear()
   ' Recreates the image list.
   Create
End Sub
Public Function AddFromFile( _
        ByVal sFIleName As String, _
        ByVal iType As ImageTypes, _
        Optional ByVal vKey As Variant, _
        Optional ByVal bMapSysColors As Boolean = False, _
        Optional ByVal lBackColor As OLE_COLOR = -1, _
        Optional ByVal vKeyAfter As Variant _
    ) As Long
Dim hImage As Long
Dim un2 As Long
Dim lR As Long
    
   ' Adds an image or series of images from a file:
   If (hIml <> 0) Then
      un2 = LR_LOADFROMFILE
      ' Load the image from file:
      If bMapSysColors Then
          un2 = un2 Or LR_LOADMAP3DCOLORS
      End If
      hImage = LoadImage(App.hInstance, sFIleName, iType, 0, 0, un2)
      AddFromFile = AddFromHandle(hImage, iType, vKey, lBackColor, vKeyAfter)
      Select Case iType
      Case IMAGE_ICON
         DestroyIcon hImage
      Case IMAGE_CURSOR
         DestroyCursor hImage
      Case IMAGE_BITMAP
         DeleteObject hImage
      End Select
   Else
      ' no image list...
      AddFromFile = False
   End If
                  
End Function
Public Function AddFromResourceID( _
      ByVal lID As Long, _
      ByVal hInst As Long, _
      ByVal iType As ImageTypes, _
      Optional ByVal vKey As Variant, _
      Optional ByVal bMapSysColors As Boolean = False, _
      Optional ByVal lBackColor As OLE_COLOR = -1, _
      Optional ByVal vKeyAfter As Variant _
    ) As Long
Dim hImage As Long
Dim un2 As Long
Dim lR As Long
Dim iX As Long, iY As Long
    
   ' Adds an image or series of images from a resource id.  Note this will
   ' only work when working on a resource in a compiled executable:
   If (hIml <> 0) Then
      ' Load the image from file:
      If bMapSysColors Then
          un2 = un2 Or LR_LOADMAP3DCOLORS
      End If
      ' Choose the icon closest to the image list size:
      If iType <> IMAGE_BITMAP Then
         iX = m_lIconSizeX
         iY = m_lIconSizeY
      End If
      If hInst = 0 Then
         ' Assume we're trying to pick a shared
         ' resource
         un2 = un2 Or LR_COPYRETURNORG
      End If
      hImage = LoadImageLong(hInst, lID, iType, iX, iY, un2)
      AddFromResourceID = AddFromHandle(hImage, iType, vKey, lBackColor, vKeyAfter)
      Select Case iType
      Case IMAGE_ICON
         DestroyIcon hImage
      Case IMAGE_CURSOR
         DestroyCursor hImage
      Case IMAGE_BITMAP
         DeleteObject hImage
      End Select
   Else
      ' no image list...
      AddFromResourceID = False
   End If
   
End Function

Public Function AddFromHandle( _
      ByVal hImage As Long, _
      ByVal iType As ImageTypes, _
      Optional ByVal vKey As Variant, _
      Optional ByVal lBackColor As OLE_COLOR = -1, _
      Optional ByVal vKeyAfter As Variant _
   ) As Boolean
Dim lR As Long
Dim lDst As Long
Dim bOk As Boolean
Dim bInsert As Boolean
Dim i As Long, j As Long
Dim iOrigCount As Long
Dim iCount As Long
Dim sSwapKey As String

   ' Adds an image or series of images from a GDI image handle.
   If (m_hIml <> 0) Then
      If (hImage <> 0) Then
         iOrigCount = ImageCount
         
         bOk = True
         If Not IsMissing(vKeyAfter) Then
            If (ImageCount > 0) Then
               If vKeyAfter = 0 Then
                  bInsert = False
                  lDst = 0
               Else
                  bInsert = True
                  bOk = False
                  lDst = ItemIndex(vKeyAfter)
                  If (lDst > -1) Then
                     bOk = True
                  End If
               End If
            End If
         End If
         
         If (bOk) Then
            If (iType = IMAGE_BITMAP) Then
               ' And add it to the image list:
               If (lBackColor = -1) Then
                   ' Ideally Determine the top left pixel of the
                   ' bitmap and use as back colour...
                   Dim lHDCDisp As Long, lHDC As Long, hBmpOld As Long
                   lHDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
                   If lHDCDisp <> 0 Then
                     lHDC = CreateCompatibleDC(lHDCDisp)
                     DeleteDC lHDCDisp
                     If lHDC <> 0 Then
                        hBmpOld = SelectObject(lHDC, hImage)
                        If hBmpOld <> 0 Then
                           ' Get the colour of the 0,0 pixel:
                           lBackColor = GetPixel(lHDC, 0, 0)
                           SelectObject lHDC, hBmpOld
                        End If
                        DeleteObject lHDC
                        DeleteDC lHDC
                     End If
                  End If
               End If
               lR = ImageList_AddMasked(hIml, hImage, lBackColor)
            ElseIf (iType = IMAGE_ICON) Or (iType = IMAGE_CURSOR) Then
               ' Add the icon:
               lR = ImageList_AddIcon(hIml, hImage)
            End If
         End If
         
         If (lR > -1) Then
            If (bInsert) Then
               If (lDst < ImageCount - 1) Then
                  ' We are inserting and have to swap all
                  ' the images.
                  pEnsureKeys
                  iCount = ImageCount
                  For i = iOrigCount - 1 To lDst Step -1
                     For j = i To i + iCount - iOrigCount - 1
                        ImageList_Copy m_hIml, j + 1, m_hIml, j, eilSwap
                        sSwapKey = m_sKey(j)
                        m_sKey(j) = m_sKey(j + 1)
                        m_sKey(j + 1) = sSwapKey
                     Next j
                  Next i
                  
               End If
            End If
         End If
         
      Else
          lR = -1
      End If
   Else
      lR = -1
   End If
   
   If (lR <> -1) Then
      If bInsert Then
         SetKey lDst, vKey
      Else
         SetKey lR, vKey
      End If
      AddFromHandle = (lR <> -1)
   End If
   pEnsureKeys
   
End Function
Public Function AddFromPictureBox( _
        ByVal hDC As Long, _
        pic As Object, _
        Optional ByVal vKey As Variant, _
        Optional ByVal LeftPixels As Long = 0, _
        Optional ByVal TopPixels As Long = 0, _
        Optional ByVal lBackColor As OLE_COLOR = -1 _
    ) As Long
Dim lHDC As Long
Dim lhBmp As Long, lhBmpOld As Long
Dim tBM As BITMAP
Dim lAColor As Long
Dim lW As Long, lH As Long
Dim hBrush As Long
Dim tR As RECT
Dim lR As Long
Dim lBPixel As Long
   
   ' Adds an image or series of images from an area of a PictureBox
   ' or other Device Context:
   lR = -1
   If (hIml <> 0) Then
      ' Create a DC to hold the bitmap to transfer into the image list:
      lHDC = CreateCompatibleDC(hDC)
      If (lHDC <> 0) Then
          lhBmp = CreateCompatibleBitmap(hDC, m_lIconSizeX, m_lIconSizeY)
          If (lhBmp <> 0) Then
              ' Get the backcolor to use:
              If (lBackColor = -1) Then
                  ' None specified, use the colour at 0,0:
                  lBackColor = GetPixel(pic.hDC, 0, 0)
              Else
                  ' Try to get the specified backcolor:
                  If OleTranslateColor(lBackColor, 0, lAColor) Then
                      ' Failed- use default of silver
                      lBackColor = &HC0C0C0
                  Else
                      ' Set to GDI version of OLE Color
                      lBackColor = lAColor
                  End If
              End If
              ' Select the bitmap into the DC
              lhBmpOld = SelectObject(lHDC, lhBmp)
              ' Clear the background:
              hBrush = CreateSolidBrush(lBackColor)
              tR.Right = m_lIconSizeX: tR.Bottom = m_lIconSizeY
              FillRect lHDC, tR, hBrush
              DeleteObject hBrush
              
              ' Get the source picture's dimension:
              GetObjectAPI pic.Picture.Handle, LenB(tBM), tBM
              lW = 16
              lH = 16
              If (lW + LeftPixels > tBM.bmWidth) Then
                  lW = tBM.bmWidth - LeftPixels
              End If
              If (lH + TopPixels > tBM.bmHeight) Then
                  lH = tBM.bmHeight - TopPixels
              End If
              If (lW > 0) And (lH > 0) Then
                  ' Blt from the picture into the bitmap:
                  lR = BitBlt(lHDC, 0, 0, lW, lH, hDC, LeftPixels, TopPixels, SRCCOPY)
                  Debug.Assert (lR <> 0)
              End If
              
              ' We now have the image in the bitmap, so select it out of the DC:
              SelectObject lHDC, lhBmpOld
              ' And add it to the image list:
              AddFromHandle lhBmp, IMAGE_BITMAP, vKey, lBackColor
                  
              DeleteObject lhBmp
          End If
          ' Clear up the DC:
          DeleteDC lHDC
      End If
   End If

   If (lR <> -1) Then
        SetKey lR, vKey
   End If
   
   AddFromPictureBox = lR + 1
   pEnsureKeys
   
End Function
Private Sub SetKey(ByVal lIndex As Long, ByVal vKey As Variant)
Dim sKey As String
Dim lI As Long

   If (IsEmpty(vKey) Or IsMissing(vKey)) Then
      sKey = ""
   Else
      sKey = vKey
   End If
    
   If (m_hIml <> 0) Then
      
      On Error Resume Next
      lI = UBound(m_sKey)
      If (Err.Number = 0) Then
         If (lIndex > lI) Then
            ReDim Preserve m_sKey(0 To lIndex) As String
         End If
      Else
         ReDim Preserve m_sKey(0 To lIndex) As String
      End If
      
      For lI = 0 To UBound(m_sKey)
         If Not lI = lIndex Then
            If Trim$(m_sKey(lI)) <> "" Then
               If m_sKey(lI) = vKey Then
                  Err.Raise 457
                  Exit Sub
               End If
            End If
         End If
      Next lI
      m_sKey(lIndex) = vKey
   End If
End Sub
Public Property Get hIml() As Long
   ' Returns the ImageList handle:
    hIml = m_hIml
End Property
Public Property Get ImagePictureStrip( _
      Optional ByVal vStartKey As Variant, _
      Optional ByVal vEndKey As Variant, _
      Optional ByVal oBackColor As OLE_COLOR = vbButtonFace _
   ) As IPicture
Dim iStart As Long
Dim iEnd As Long
Dim iImgIndex As Long
Dim lHDC As Long
Dim lParenthDC As Long
Dim lhBmp As Long
Dim lhBmpOld As Long
Dim lSizeX As Long
Dim hBr As Long
Dim tR As RECT
Dim lColor As Long
   
   If (m_hIml <> 0) Then
      If (IsMissing(vStartKey)) Then
         iStart = 0
      Else
         iStart = ItemIndex(vStartKey)
      End If
      If (IsMissing(vEndKey)) Then
         iEnd = ImageCount - 1
      Else
         iEnd = ItemIndex(vEndKey)
      End If
      
      If (iEnd > iStart) And (iEnd > -1) Then
         lParenthDC = m_HDC
         lHDC = CreateCompatibleDC(lParenthDC)
         If (lHDC <> 0) Then
            lSizeX = ImageCount * m_lIconSizeX
            lhBmp = CreateCompatibleBitmap(lParenthDC, lSizeX, m_lIconSizeY)
            If (lhBmp <> 0) Then
               lhBmpOld = SelectObject(lHDC, lhBmp)
               If (lhBmpOld <> 0) Then
                  lColor = TranslateColor(oBackColor)
                  tR.Bottom = m_lIconSizeY
                  tR.Right = lSizeX
                  hBr = CreateSolidBrush(lColor)
                  FillRect lHDC, tR, hBr
                  DeleteObject hBr
                  For iImgIndex = iStart To iEnd
                     ImageList_Draw m_hIml, iImgIndex, lHDC, iImgIndex * m_lIconSizeX, 0, ILD_Transparent
                  Next iImgIndex
                  SelectObject lHDC, lhBmpOld
                  Set ImagePictureStrip = BitmapToPicture(lhBmp)
               Else
                  DeleteObject lhBmp
               End If
            End If
            DeleteDC lHDC
         End If
      End If
   End If
   
End Property

Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    ' This is all magic if you ask me:
    Dim NewPic As Picture, PicConv As PictDesc, IGuid As GUID
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect PicConv, IGuid, True, NewPic
    
    Set IconToPicture = NewPic
    
End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture

   If (hBmp = 0) Then Exit Function
   
   Dim NewPic As Picture, tPicConv As PictDesc, IGuid As GUID
   
   ' Fill PictDesc structure with necessary parts:
   With tPicConv
      .cbSizeofStruct = Len(tPicConv)
      .picType = vbPicTypeBitmap
      .hImage = hBmp
   End With
   
   ' Fill in IDispatch Interface ID
   With IGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   
   ' Create a picture object:
   OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
   ' Return it:
   Set BitmapToPicture = NewPic
      

End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If
End Function
Private Sub pEnsureKeys()
Dim iCount As Long
Dim iU As Long
   If m_hIml <> 0 Then
      iCount = ImageCount
      On Error Resume Next
      iU = UBound(m_sKey)
      If (Err.Number <> 0) Then iU = -1
      Err.Clear
      If (iU <> iCount - 1) Then
         ReDim Preserve m_sKey(0 To iCount - 1) As String
      End If
   End If
End Sub

Private Sub Class_Initialize()
   m_lIconSizeX = 16
   m_lIconSizeY = 16
   m_eColourDepth = ILC_COLOR
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub



