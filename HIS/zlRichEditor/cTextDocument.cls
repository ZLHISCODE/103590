VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'#########################################################################
'##模 块 名：cTextDocument.cls
'##创 建 人：吴庆伟
'##日    期：2005年4月5日
'##描    述：实现 TOM (Text Object Model)接口
'##要    求：ISHF_Ex.tlb 与 SHLEXT.tlb 两个TLB文件
'##说    明：RTF对象化的强有力工具！
'#########################################################################

Option Explicit
 
'以下 ENUM 由ActivX接口查询工具自动生成
Public Enum tomEnum
    tomFalse = 0
    tomTrue = -1
    tomUndefined = -9999999
    tomToggle = -9999998
    tomAutoColor = -9999997
    tomDefault = -9999996
    tomSuspend = -9999995
    tomResume = -9999994
    tomApplyNow = 0
    tomApplyLater = 1
    tomTrackParms = 2
    tomCacheParms = 3
    tomApplyTmp = 4
    tomBackward = -1073741823
    tomForward = 1073741823
    tomMove = 0
    tomExtend = 1
    tomNoSelection = 0
    tomSelectionIP = 1
    tomSelectionNormal = 2
    tomSelectionFrame = 3
    tomSelectionColumn = 4
    tomSelectionRow = 5
    tomSelectionBlock = 6
    tomSelectionInlineShape = 7
    tomSelectionShape = 8
    tomSelStartActive = 1
    tomSelAtEOL = 2
    tomSelOvertype = 4
    tomSelActive = 8
    tomSelReplace = 16
    tomEnd = 0
    tomStart = 32
    tomCollapseEnd = 0
    tomCollapseStart = 1
    tomClientCoord = 256
    tomAllowOffClient = 512
    tomNone = 0
    tomWords = 2
    tomDouble = 3
    tomDotted = 4
    tomDash = 5
    tomDashDot = 6
    tomDashDotDot = 7
    tomWave = 8
    tomThick = 9
    tomHair = 10
    tomDoubleWave = 11
    tomHeavyWave = 12
    tomLongDash = 13
    tomThickDash = 14
    tomThickDashDot = 15
    tomThickDashDotDot = 16
    tomThickDotted = 17
    tomThickLongDash = 18
    tomLineSpaceSingle = 0
    tomLineSpace1pt5 = 1
    tomLineSpaceDouble = 2
    tomLineSpaceAtLeast = 3
    tomLineSpaceExactly = 4
    tomLineSpaceMultiple = 5
    tomAlignLeft = 0
    tomAlignCenter = 1
    tomAlignRight = 2
    tomAlignJustify = 3
    tomAlignDecimal = 3
    tomAlignBar = 4
    tomAlignInterWord = 3
    tomAlignInterLetter = 4
    tomAlignScaled = 5
    tomAlignGlyphs = 6
    tomAlignSnapGrid = 7
    tomSpaces = 0
    tomDots = 1
    tomDashes = 2
    tomLines = 3
    tomThickLines = 4
    tomEquals = 5
    tomTabBack = -3
    tomTabNext = -2
    tomTabHere = -1
    tomListNone = 0
    tomListBullet = 1
    tomListNumberAsArabic = 2
    tomListNumberAsLCLetter = 3
    tomListNumberAsUCLetter = 4
    tomListNumberAsLCRoman = 5
    tomListNumberAsUCRoman = 6
    tomListNumberAsSequence = 7
    tomListParentheses = 65536
    tomListPeriod = 131072
    tomListPlain = 196608
    tomCharacter = 1
    tomWord = 2
    tomSentence = 3
    tomParagraph = 4
    tomLine = 5
    tomStory = 6
    tomScreen = 7
    tomSection = 8
    tomColumn = 9
    tomRow = 10
    tomWindow = 11
    tomCell = 12
    tomCharFormat = 13
    tomParaFormat = 14
    tomTable = 15
    tomObject = 16
    tomPage = 17
    tomMatchWord = 2
    tomMatchCase = 4
    tomMatchPattern = 8
    tomUnknownStory = 0
    tomMainTextStory = 1
    tomFootnotesStory = 2
    tomEndnotesStory = 3
    tomCommentsStory = 4
    tomTextFrameStory = 5
    tomEvenPagesHeaderStory = 6
    tomPrimaryHeaderStory = 7
    tomEvenPagesFooterStory = 8
    tomPrimaryFooterStory = 9
    tomFirstPageHeaderStory = 10
    tomFirstPageFooterStory = 11
    tomNoAnimation = 0
    tomLasVegasLights = 1
    tomBlinkingBackground = 2
    tomSparkleText = 3
    tomMarchingBlackAnts = 4
    tomMarchingRedAnts = 5
    tomShimmer = 6
    tomWipeDown = 7
    tomWipeRight = 8
    tomAnimationMax = 8
    tomLowerCase = 0
    tomUpperCase = 1
    tomTitleCase = 2
    tomSentenceCase = 4
    tomToggleCase = 5
    tomReadOnly = 256
    tomShareDenyRead = 512
    tomShareDenyWrite = 1024
    tomPasteFile = 4096
    tomCreateNew = 16
    tomCreateAlways = 32
    tomOpenExisting = 48
    tomOpenAlways = 64
    tomTruncateExisting = 80
    tomRTF = 1
    tomText = 2
    tomHTML = 3
    tomWordDocument = 4
    tomBold = -2147483647
    tomItalic = -2147483646
    tomUnderline = -2147483644
    tomStrikeout = -2147483640
    tomProtected = -2147483632
    tomLink = -2147483616
    tomSmallCaps = -2147483584
    tomAllCaps = -2147483520
    tomHidden = -2147483392
    tomOutline = -2147483136
    tomShadow = -2147482624
    tomEmboss = -2147481600
    tomImprint = -2147479552
    tomDisabled = -2147475456
    tomRevised = -2147467264
    tomNormalCaret = 0
    tomKoreanBlockCaret = 1
    tomIncludeInset = 1
    tomIgnoreCurrentFont = 0
    tomMatchFontCharset = 1
    tomMatchFontSignature = 2
    tomCharset = -2147483648#
    tomRE10Mode = 1
    tomUseAtFont = 2
    tomTextFlowMask = 12
    tomTextFlowES = 0
    tomTextFlowSW = 4
    tomTextFlowWN = 8
    tomTextFlowNE = 12
    tomUsePassword = 16
    tomNoIME = 524288
    tomSelfIME = 262144
End Enum

Private m_oTextDocument As TOM.ITextDocument
Public hWnd As Long

Private Property Get Succeeded(ByVal hr As Long) As Boolean
   ' COM Success macro: check Severity is 0
   Succeeded = ((hr And &H80000000) = 0)
End Property

Public Property Get TextDocument() As TOM.ITextDocument
   ' Returns the current TextDocument interface:
   Set TextDocument = m_oTextDocument
End Property

Public Function Init(ByRef edtThis As Object, ByRef oParent As Object) As Boolean
' Generates an ITextDocument interface from

Dim iUnk As IShellFolderEx_TLB.IUnknown
Dim oTOM As TOM.ITextDocument
Dim iUnkTom As IShellFolderEx_TLB.IUnknown
Dim IID_ITextDocument As IShellFolderEx_TLB.GUID
Dim hr As Long

   ' Clear any existing TOM we have:
   ClearUp

   ' If we passed in an object (can Init Nothing to clear up):
   If Not (edtThis Is Nothing) Then

      ' Prepare the IID for ITextDocument interface:
      ' 8CC497C0-A1DF-11ce-8098-00AA0047BE5D
      With IID_ITextDocument
         .Data1 = &H8CC497C0
         .Data2 = &HA1DF
         .Data3 = &H11CE
         .Data4(0) = &H80
         .Data4(1) = &H98
         .Data4(2) = &H0
         .Data4(3) = &HAA
         .Data4(4) = &H0
         .Data4(5) = &H47
         .Data4(6) = &HBE
         .Data4(7) = &H5D
      End With
      
      ' Ask the RichEdit control to provide its OLE Interface:
      SendMessage edtThis.hWnd, EM_GETOLEINTERFACE, 0, iUnk
      
      ' Now query for ITextDocument support:
      hr = iUnk.QueryInterface(IID_ITextDocument, oTOM)
      
      ' If we got it...
      If Succeeded(hr) Then
         
         ' ... then AddRef (COM rules):
         Set iUnkTom = oTOM
         iUnkTom.AddRef
         Set iUnkTom = Nothing
         
         ' Store a reference to it:
         Set m_oTextDocument = oTOM
         
         hWnd = edtThis.hWnd
'         Set objEditControl = edtThis
'         Set objParent = oParent
         ' Done!
         Init = True
         
      Else
         Err.Raise hr, App.EXEName & ".cTextDocument", "Failed to get ITextDocument interface."
      End If
      
      ' We release the AddRef that was added by RichEdit:
      iUnk.Release
      ' But we don't want VB to try and release it again:
      CopyMemory iUnk, 0&, 4
   
   Else
      Init = True
      
   End If
   
End Function
Public Property Get RTF(ByRef oRange As ITextRange) As String
' A helper function to return the RTF for a selected range
' unfortunately setting RTF into a range is a bit more
' challenging...

Dim sName As String
Dim oData As IDataObject
Dim pf As FORMATETC
Dim pm As STGMEDIUM
Dim hGlobal As Long
Dim lPtr As Long
Dim lSize As Long
Dim iFmt As Long
Dim iIFmt As Integer

   ' Request to copy the range into an object
   ' implementing IDataObject.  The clipboard
   ' is unaffected because the data goes directly
   ' to oData.  Set oData to a null ptr to
   ' copy the range to the clipboard.
   oRange.Copy oData
                     
   ' Now get the Clipboard format ID of the Rich Text Format:
   sName = "Rich Text Format"
   iFmt = RegisterClipboardFormat(sName & Chr$(0))
   
   ' Unfortunately the format is limited to an integer, but custom
   ' formats tend to be 32767<iFmt<65535.  This causes an overflow
   ' if we try to assign directly, so CopyMemory into the integer
   ' instead:
   CopyMemory iIFmt, iFmt, 2
   
   ' Set up the FORMATETC structure specifying the format that
   ' we want:
   pf.cfFormat = iIFmt ' the RTF format.  vbCFText will work here too
   pf.dwAspect = DVASPECT_CONTENT ' irrelevant for RTF
   pf.lIndex = -1 ' entire contents
   pf.pDVTARGETDEVICE = 0 ' null - irrelevant
   pf.TYMED = TYMED_HGLOBAL ' ask for the
   
   ' Now set up the STGMEDIUM structure specifying how we'd like
   ' the contents returned. Here I ask for a hMem handle to global
   ' memory, and I specify I want to free it.
   pm.TYMED = TYMED_HGLOBAL
   pm.Data = hGlobal
   pm.pUnkForRelease = 0 'null
            
   ' Actually get the data in the requested format from the
   ' IDataObject.  If the format is not available, you'll get
   ' an Automation error here.
   oData.GetData pf, pm
   
   ' If we succeed then pm.Data contains our hMem:
   If pm.Data Then
   
      ' Standard code for reading from a handle to global memory:
      lSize = GlobalSize(pm.Data)
      If lSize Then
         lPtr = GlobalLock(pm.Data)
         If lPtr Then
            ReDim b(0 To lSize - 1) As Byte
            CopyMemory b(0), ByVal lPtr, lSize
            GlobalUnlock pm.Data
            
            ' Return the data!
            RTF = StrConv(b, vbUnicode)
            
         End If
      End If
      
      ' Don't forget to free the handle
      GlobalFree pm.Data
   End If

End Property

Public Sub ClearUp()
Dim iUnkTom As IShellFolderEx_TLB.IUnknown
Dim lR As Long
   
   If Not m_oTextDocument Is Nothing Then
      ' Release our AddRef:
      Set iUnkTom = m_oTextDocument
      iUnkTom.Release
      ' Remove the remaining references:
      Set iUnkTom = Nothing
      
      Set m_oTextDocument = Nothing
   End If
   
End Sub

Private Sub Class_Terminate()
   ' Ensure we release the class
   ClearUp
End Sub


