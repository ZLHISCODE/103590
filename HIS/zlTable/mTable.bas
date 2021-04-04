Attribute VB_Name = "mTable"
'#########################################################################
'##��    ��������������
'#########################################################################
Option Explicit

Public p_lIconWidth As Long         'ͼ����
Public p_lIconHeight As Long        'ͼ��߶�
Public p_TPPX As Long               'screen.TwipsPerPixelX
Public p_TPPY As Long               'screen.TwipsPerPixelY

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

' Use default rgb colour:
Public Const CLR_NONE = -1
    
' Private variables
Private m_bIsXp As Boolean
Private m_bIsNt As Boolean
Private m_bInit As Boolean
Public Const MAGIC_END_EDIT_IGNORE_WINDOW_PROP As String = "zlTable:Table"

Public Type ColInfoType
    ColWidth As Long
    LeftX As Long
    FixedWidth As Boolean
    Visible As Boolean
End Type

Public Type RowInfoType
    RowHeight As Long           '�����и�
    FixedHeight As Boolean
    TopY As Long
End Type

Public ColInfo() As ColInfoType
Public RowInfo() As RowInfoType

'#########################################################################################################
'## ���ܣ�  ����ָ������Ĵ����������
'## ������  hWnd: ������
'## ���أ�  ������
'#########################################################################################################
Public Function WindowClassName(ByVal hWnd As Long) As String
    Dim szBuf As String
    Dim lR As Long
    szBuf = String$(260, 0)
    lR = GetClassName(hWnd, szBuf, 260)
    lR = InStr(szBuf, vbNullChar)
    If (lR > 0) Then
        WindowClassName = Left$(szBuf, lR - 1)
    Else
        WindowClassName = szBuf
    End If
End Function

'#########################################################################################################
'## ���ܣ�  ����XOR�������б�Resize��ʱ������϶���ͼ��
'## ������  rcNew:  �����϶�ͼ��ľ���
'##         bFirst: �Ƿ��ǵ�һ�λ���
'##         bLast:  �Ƿ������һ�λ���
'## ���أ�  ��
'#########################################################################################################
Public Sub DrawDragImage(ByRef rcNew As RECT, ByVal bFirst As Boolean, ByVal bLast As Boolean)
    Static rcCurrent As RECT    '��̬���������ڴ洢λ��
    Dim hDC As Long
       
    '���Ȼ�ȡ����DC
    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    '���û���ģʽΪXOR
    SetROP2 hDC, R2_NOTXORPEN
    
    '���Ǻ�����ɾ���
    If Not (bFirst) Then
       Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
    End If
    
    If Not (bLast) Then
       '�����¾���
       Rectangle hDC, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
    End If
    
    '�洢���λ�ã����������´β�����
    LSet rcCurrent = rcNew  '�û��Զ�������ĸ�ֵ�� LSet
    
    '�����ͷ�����DC���ǵ�Ҫ����һ������
    DeleteDC hDC
End Sub

'#########################################################################################################
'## ���ܣ�  ����ָ��ѡ��������ImageList�е�һ��ͼ��
'## ������
'##         hIml:               һ��VB6 ImageList��hImageList
'##         iIndex:             ��0��ʼ��ͼ������
'##         hDC:                Ŀ���豸�������
'##         xPixels:            X��λ��
'##         yPixels:            Y��λ��
'##         lIconSizeX:         ˮƽ�ߴ�
'##         lIconSizeY:         ��ֱ�ߴ�
'##         bSelected:          ѡ��Ч��
'##         bDisabled:          ��Чͼ��
'## ���أ�  ��
'#########################################################################################################
Public Sub DrawImageIcon( _
    ByVal lhIml As Long, _
    ByVal iIndex As Long, _
    ByVal hDC As Long, _
    ByVal xPixels As Integer, _
    ByVal yPixels As Integer, _
    ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
    Optional ByVal bSelected As Boolean = False, _
    Optional ByVal bDisabled As Boolean = False)
      
    If iIndex > -1 Then
        Dim lFlags As Long
        lFlags = ILD_TRANSPARENT
        If (bSelected) Then
            lFlags = lFlags Or ILD_SELECTED
        End If
        If (bDisabled) Then
            Dim hIcon As Long
            hIcon = ImageList_GetIcon(lhIml, iIndex, 0)
                           
            If Not (hIcon = 0) And Not (hIcon = -1) Then
                DrawState hDC, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
                ' Clear up the icon:
                DestroyIcon hIcon
            End If
        Else
            ImageList_Draw lhIml, iIndex, hDC, xPixels, yPixels, lFlags
        End If
    End If
End Sub

'#########################################################################################################
'## ���ܣ�  �ڷ����е�ǰ�����һ��չ��/�رյķ��š�
'##         �����XPЧ���£������һ��TreeView��չ��/�رշ��ţ�������һ����ť����+��-�š�
'##
'## ������  hWnd:       ���ڼ������Ĵ�����
'##         lHDC:       Ŀ���豸�������
'##         tTR:        ���Ż��Ƶı߽����
'##         bCollapsed: True��ʾ�۵���False��ʾչ��
'## ���أ�  ��
'#########################################################################################################
Public Sub DrawOpenCloseGlyph( _
    ByVal hWnd As Long, _
    ByVal lhDC As Long, _
    tTR As RECT, _
    ByVal bCollapsed As Boolean)
      
    Dim tGR As RECT
    Dim bDone As Boolean
   
    LSet tGR = tTR
    tGR.Left = tGR.Left + 2
    tGR.Right = tGR.Left + 12
    tGR.Top = tGR.Top + (tGR.Bottom - tGR.Top - 12) \ 2
    tGR.Bottom = tGR.Top + 12
    
    If (IsXp) Then
        'XPϵͳ
        Dim hTheme As Long
        hTheme = OpenThemeData(hWnd, StrPtr("TREEVIEW"))    '��ȡTreeView������
        If Not (hTheme = 0) Then
            '��ȡ�ɹ�������ʾTreeView�����չ��/�رշ���
            DrawThemeBackground hTheme, lhDC, 2, IIf(bCollapsed, 1, 2), tGR, tGR
            CloseThemeData hTheme   '�ر�����
            bDone = True
        End If
    End If
    
    If Not (bDone) Then
        '���ư�ť�߿�
        Dim hBr As Long
        hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        FillRect lhDC, tGR, hBr
        DeleteObject hBr
        
        Dim hPen As Long
        Dim hPenOld As Long
        Dim tJ As POINTAPI
        
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Left + 1, tGR.Bottom - 2, tJ
        LineTo lhDC, tGR.Right - 2, tGR.Bottom - 2
        LineTo lhDC, tGR.Right - 2, tGR.Top
        SelectObject lhDC, hPenOld
        DeleteObject hPen
        
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Right - 2, tGR.Top, tJ
        LineTo lhDC, tGR.Left, tGR.Top
        LineTo lhDC, tGR.Left, tGR.Bottom - 1
        SelectObject lhDC, hPenOld
        DeleteObject hPen
                          
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
        hPenOld = SelectObject(lhDC, hPen)
        MoveToEx lhDC, tGR.Left, tGR.Bottom - 1, tJ
        LineTo lhDC, tGR.Right - 1, tGR.Bottom - 1
        LineTo lhDC, tGR.Right - 1, tGR.Top
        
        ' Draw collapse/expand glyph
        MoveToEx lhDC, tGR.Left + 3, tGR.Top + 5, tJ
        LineTo lhDC, tGR.Left + 8, tGR.Top + 5
        If (bCollapsed) Then
            MoveToEx lhDC, tGR.Left + 5, tGR.Top + 3, tJ
            LineTo lhDC, tGR.Left + 5, tGR.Top + 8
        End If
        SelectObject lhDC, hPenOld
        DeleteObject hPen
    End If
End Sub

'#########################################################################################################
'## ���ܣ�  ��һ��OLE_COLOR��ɫת��Ϊһ��RGB������ֵ
'## ������  oClr:   ת������ɫ
'##         hPal:   ��ת��ʱ���õĵ�ɫ��
'## ���أ�  RGB�ȼ�ֵ������-1��ʾû��
'#########################################################################################################
Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    '���Զ���ɫת��ΪWindws��ɫ
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

'#########################################################################################################
'## ���ܣ�  ʹ��ָ����Alphaֵ�����ɫ
'##
'## ������  oColorFrom:     ������ɫ
'##         oColorTo:       �����ɫ
'##         alpha:          alphaֵ��0��255��
'## ���أ�  ��Ϻ��RGB��ɫ
'#########################################################################################################
Public Property Get BlendColor( _
    ByVal oColorFrom As OLE_COLOR, _
    ByVal oColorTo As OLE_COLOR, _
    Optional ByVal Alpha As Long = 128) As Long
    
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)

    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
   
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
     
    BlendColor = RGB( _
       ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
       ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
       ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
       )
End Property

'#########################################################################################################
'## ���ܣ�  ��һ�� StdFont ����ת��Ϊ��ȵ� Windows GDI LOGFONT �ṹ��
'##
'## ������  fntThis:    ��Ҫת��������
'##         hDC:        ���ڻ�ȡDPI��Ϣ���豸�������
'##         tLF:        ������LOGFONT�ṹ��
'## ���أ�  ��
'#########################################################################################################
Public Sub OLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 1 To UBound(temp) + 1
            .lfFaceName(iChar - 1) = temp(iChar - 1)
        Next iChar
        
        ' Based on the Win32SDK documentation:
        .lfHeight = GetPixcelHeightByPoint(hDC, fntThis.Size)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
        If (IsXp) Then
          '���ΪXPϵͳ�������ClearType����
           .lfQuality = CLEARTYPE_QUALITY
        Else
          '������ÿ��������
           .lfQuality = ANTIALIASED_QUALITY
        End If
    End With
End Sub
       
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

'#########################################################################################################
'## ���ܣ�  �����Ļ����ı�����
'## ������  hdc:        Ŀ���豸�������
'##         sString:    �ַ���
'##         lCount:     �ַ�������
'##         tR:         ���ο�
'##         lFlags:     ��־���������ԣ�
'## ���أ�  ��
'#########################################################################################################
Public Sub DrawText(ByVal hDC As Long, ByVal sString As String, ByVal lCount As Long, tR As RECT, ByVal lFlags As Long)
    Dim lPtr As Long
    Dim tIR As RECT
    
    LSet tIR = tR       '�����ṹ��
    If (IsNt) Then
        '�����NT���ϰ汾��ϵͳ������DrawTextW
        lPtr = StrPtr(sString)
        If Not (lPtr = 0) Then
            DrawTextW hDC, lPtr, -1, tIR, lFlags
        End If
        lPtr = 0
    Else
        '�������DrawTextA
        DrawTextA hDC, sString, -1, tIR, lFlags
    End If
    If (lFlags And DT_CALCRECT) = DT_CALCRECT Then
        '���ֻ�����ı��ߴ�
        LSet tR = tIR
    End If
End Sub

'#########################################################################################################
'## ���ܣ�  ��һ��λͼƽ�̵�ѡ������
'##
'## ������  hDC:        ���ڻ��Ƶ��豸�������
'##         x:          ��ʼX����
'##         y:          ��ʼY����
'##         Width:      ������
'##         Height:     ����߶�
'##         lSrcDC:     ����ͼ����豸�������
'##         lBitmapW:   Դλͼ���
'##         lBitmapH:   Դλͼ�߶�
'##         lSrcOffsetX:    X��ƫ����
'##         lSrcOffsetY:    Y��ƫ����
'## ���أ�  ��
'#########################################################################################################
Public Sub TileArea( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal lSrcDC As Long, _
        ByVal lBitmapW As Long, _
        ByVal lBitmapH As Long, _
        ByVal lSrcOffsetX As Long, _
        ByVal lSrcOffsetY As Long)
        
    Dim lSrcX As Long
    Dim lSrcY As Long
    Dim lSrcStartX As Long
    Dim lSrcStartY As Long
    Dim lSrcStartWidth As Long
    Dim lSrcStartHeight As Long
    Dim lDstX As Long
    Dim lDstY As Long
    Dim lDstWidth As Long
    Dim lDstHeight As Long

    lSrcStartX = ((X + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((Y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hDC, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop
End Sub

'#########################################################################################################
'## ���ܣ�  ����һ��ͨ��ObjPtr()���ص�COM����ķ�����ָ�룬����һ����������
'## ������  lPtr:   һ��COM����ķ�����ָ��
'## ���أ�  ���ָ��ָ���COM����
'#########################################################################################################
Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    ' �Ƚ�ָ��ָ��һ���Ƿ��ġ�δ�����Ľӿ�
    CopyMemory oTemp, lPtr, 4
    ' ��Ҫ��������ֹ���򣬻����������
    ' ����һ���Ϸ�������
    Set ObjectFromPtr = oTemp
    ' ��Ҫ��������ֹ���򣬻����������
    ' �ƻ��úϷ�����
    CopyMemory oTemp, 0&, 4
    'OK����������ֹ������Ȼ�������������������Ϊ���������δ�����Ľӿڣ�
End Property

'#########################################################################################################
'## ���ܣ�  �ж�ϵͳ�Ƿ���XP�����߸��߰汾��
'##
'## ���أ�  True��ʾ��XPϵͳ���߸��߰汾
'#########################################################################################################
Public Property Get IsXp() As Boolean
    If Not (m_bInit) Then
        VerInitialise   '��ȡWindows�汾��Ϣ
    End If
    IsXp = m_bIsXp
End Property

'#########################################################################################################
'## ���ܣ�  �ж�ϵͳ�Ƿ���NTϵͳ
'##
'## ���أ�  True��ʾ��NT/2000/XP or above
'#########################################################################################################
Public Property Get IsNt() As Boolean
    If Not (m_bInit) Then
        VerInitialise
    End If
    IsNt = m_bIsNt
End Property

'#########################################################################################################
'## ���ܣ�  ��ȡWindows�汾��Ϣ
'#########################################################################################################
Private Sub VerInitialise()
    Dim lMajor As Long
    Dim lMinor As Long
    GetWindowsVersion lMajor, lMinor
    If (lMajor > 5) Then
        m_bIsXp = True  '��XPϵͳ
    ElseIf (lMajor = 5) And (lMinor >= 1) Then
        m_bIsXp = True  '��XPϵͳ
    End If
    m_bInit = True
End Sub

'#########################################################################################################
'## ���ܣ�  ���ص�ǰ��Windows�汾��Ϣ
'#########################################################################################################
Private Sub GetWindowsVersion( _
    Optional ByRef lMajor = 0, _
    Optional ByRef lMinor = 0, _
    Optional ByRef lRevision = 0, _
    Optional ByRef lBuildNumber = 0)
      
    Dim lR As Long
    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)
    m_bIsNt = ((lR And &H80000000) = 0)     '�Ƿ���NTϵͳ
End Sub

'#########################################################################################################
'## ���ܣ�  ���ô���͸����
'#########################################################################################################
Public Sub SetTransparentForm(OBJ As Object, ByVal TransNum As Long)
    Dim Ret As Long
    'Set the window style to 'Layered'
    Ret = GetWindowLong(OBJ.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong OBJ.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes OBJ.hWnd, 0, TransNum, LWA_ALPHA
End Sub


