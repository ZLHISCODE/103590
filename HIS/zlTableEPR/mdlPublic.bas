Attribute VB_Name = "mdlPublic"
Option Explicit
Public gTargetDC As Long

'######################################################################################
'   ȫ�ֳ���������ҳ����ʾ��������
'######################################################################################

Public Const HSTEP = 50         '������ˮƽ����
Public Const VSTEP = 50         '��������ֱ����
Public Const PAGEMARGIN = 200   'ҳ����ͼ�¿ؼ��������ı߾�
Public Const SHADOWOFFSET = 30  '��Ӱƫ����
Public Const WHEELNUMBER = 20   '������ϵ��
'######################################################################################
'��ȡ��Ӣ�Ļ���ַ�������
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'����ת��
Public Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

'######################################################################################

Public Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Public Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_TRIANGLE, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'���λ����Ϣ
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Type Size
    cx As Long
    cy As Long
End Type
' Used to create the metafile
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

' Used for creating the temporary WMF file
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8 ' Map mode anisotropic
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'VB Errors
Private Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Private Const DSna = &H220326

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '��׽����״̬

'######################################################################################
'   ���ģʽ��س�������            RTB SDK 3.0
'######################################################################################
Public Const EM_OUTLINE = (WM_USER + 220)


Public Const EMO_EXIT = 0                     ' // enter normal mode,  lparam ignored
Public Const EMO_ENTER = 1                    ' // enter outline mode, lparam ignored
Public Const EMO_PROMOTE = 2                  ' // LOWORD(lparam) == 0 ==>
                                        ' // promote  to body-text
                                        ' // LOWORD(lparam) != 0 ==>
                                        ' // promote/demote current selection
                                        ' // by indicated number of levels
Public Const EMO_EXPAND = 3                   ' // HIWORD(lparam) = EMO_EXPANDSELECTION
                                        ' // -> expands selection to level
                                        ' // indicated in LOWORD(lparam)
                                        ' // LOWORD(lparam) = -1/+1 corresponds
                                        ' // to collapse/expand button presses
                                        ' // in winword (other values are
                                        ' // equivalent to having pressed these
                                        ' // buttons more than once)
                                        ' // HIWORD(lparam) = EMO_EXPANDDOCUMENT
                                        ' // -> expands whole document to
                                        ' // indicated level
Public Const EMO_MOVESELECTION = 4            ' // LOWORD(lparam) != 0 -> move current
                                        ' // selection up/down by indicated
                                        ' // amount
Public Const EMO_GETVIEWMODE = 5          ' // Returns VM_NORMAL or VM_OUTLINE

'   �Ƿ�չ��
Public Const EMO_EXPANDSELECTION = 0
Public Const EMO_EXPANDDOCUMENT = 1

Public Const VM_NORMAL = 4             ' // Agrees with RTF \viewkindN
Public Const VM_OUTLINE = 2

'######################################################################################
'   ���ű�����س�������            RTB SDK 3.0
'######################################################################################

Public Const EM_GETZOOM = (WM_USER + 224)
Public Const EM_SETZOOM = (WM_USER + 225)
Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long

'######################################################################################
'   ����������
'######################################################################################

Public Type POINTL
    x As Long
    y As Long
End Type
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public lpPrevWndProc As Long

Public sngX As Single, sngY As Single   '�������
Public intShift As Integer              '��갴��
Public bWay As Boolean                  '��귽��
Public bMouseFlag As Boolean            '����¼������־

'######################################################################################
'   ��ȡ�ַ���Ļλ��
'######################################################################################
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const S_FALSE = &H1
Public Const S_OK = &H0

Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

'######################################################################################
'   ֱ�ӷ��Ͱ����ĺ���
'######################################################################################
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'######################################################################################
'   ���뷨������
'######################################################################################
'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'######################################################################################
'   �ͷ��ڴ�
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = Nvl(rsTmp!����ID, 0)
        UserInfo.���� = Nvl(rsTmp!����, "")
        UserInfo.���� = Nvl(rsTmp!����, "")
        UserInfo.�û��� = Nvl(rsTmp!�û���, "")
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, ID, Caption)
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    Set AddButton = Control
End Function
'################################################################################################################
'## ���ܣ�  ��������A���õ�������B��ͬһ��
'##
'## ������  BarToDock   ������Ĺ�����
'##         BarOnLeft   ��λ����ߵĹ�����
'################################################################################################################
Public Sub DockingRightOf(Controls As CommandBars, BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    Controls.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    Controls.DockToolBar BarToDock, 0, (Bottom + Top) / 2, BarOnLeft.Position
End Sub
Public Function GetAllFonts() As Collection
'�����б�
Dim sFont As String, i As Long, FontsCol As New Collection
    On Error Resume Next
    If Not ExistsPrinter Then
        For i = 0 To Screen.FontCount - 1
           sFont = Screen.Fonts(i)
           FontsCol.Add sFont, "F_" & sFont
        Next i
    Else
        For i = 0 To Printer.FontCount - 1
           sFont = Printer.Fonts(i)
           FontsCol.Add sFont, "F_" & sFont
        Next i
    End If
    Err.Clear
    Set GetAllFonts = FontsCol
End Function
Public Function UsableFont(ByVal sFont As String) As String
'����Ч����ֱ�ӷ�������
    Err.Clear
    On Error GoTo errFont
    UsableFont = gAllFont("F_" & sFont)
    Exit Function
errFont:
    UsableFont = "����"
    Err.Clear
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
 
Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'���ܣ���ItemData����ComboBox������ֵ
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
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

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
    ToVarchar = zl9Comlib.zlStr.ToVarchar(varText, lngLength)
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
    Err.Clear
End Sub
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

Public Function GetFontSizeNumber(ByVal strFontSize As String) As Integer
    On Error Resume Next
    Dim sngNum As Single
    Select Case strFontSize
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
        sngNum = IIf(Val(strFontSize) <= 0, 10, Val(strFontSize))
    End Select
    GetFontSizeNumber = Format(sngNum, "0.0")
    Err.Clear
End Function
Public Function SetFont(ByVal lngHwnd As Long, ByVal tmphdc As Long, tmpFont As StdFont, tmpColor As OLE_COLOR) As Boolean
Dim cF As CHOOSEFONT, lF As LOGFONT
    With lF
        .lfFaceName = StrConv(tmpFont.Name, vbFromUnicode) & vbNullChar '��ʼ���������ƣ���Ҫ��Unicodeת�������Կ��ַ���β
        .lfItalic = tmpFont.Italic '��ʼ���Ƿ���б��
        .lfStrikeOut = tmpFont.Strikethrough '��ʼ���Ƿ���ɾ����
        .lfUnderline = tmpFont.Underline '��ʼ���Ƿ����»���
        .lfWeight = tmpFont.Weight '��ʼ�������С
        .lfCharSet = tmpFont.Charset '��ʼ���ַ���
        .lfHeight = -MulDiv(tmpFont.Size, GetDeviceCaps(tmphdc, LOGPIXELSY), 72) '������ת��ΪlfHeight���õ���ʽ
    End With
    With cF
        .rgbColors = tmpColor '��ʼ��������ɫ
        .lStructSize = Len(cF)
        .hWndOwner = lngHwnd
        .hInstance = App.hInstance
        .flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_EFFECTS Or CF_LIMITSIZE '����������Flags�����б�
        .lpLogFont = VarPtr(lF) '����Ϊ����õ�LogFont�ṹ��ַ
        .nSizeMin = 4 '��С�����С
        .nSizeMax = 200 '��������С
    End With
    If CHOOSEFONT(cF) = 0 Then Exit Function '�������ȡ�������˳�����
    With tmpFont
        .Name = StrConv(lF.lfFaceName, vbUnicode) '������������
        .Italic = lF.lfItalic '�����Ƿ�б��
        .Strikethrough = lF.lfStrikeOut '�����Ƿ�ɾ����
        .Underline = lF.lfUnderline '�����Ƿ��»���
        .Weight = lF.lfWeight '�����Ƿ����
        .Charset = lF.lfCharSet '�����ַ���
        .Size = -lF.lfHeight - ((-lF.lfHeight) / 4) - IIf(-lF.lfHeight Mod 4 > 1, 1, 0) '���������С��lfHeight���ֺŵ�ת����Ҫ�õ���ʽ
        tmpColor = cF.rgbColors '����������ɫ
    End With
    SetFont = True
End Function
Public Function GetSaveFile(ByVal hWndOwner As Long, ByVal strFileName As String, strFileType As String, strSaveTitle As String) As String
Dim fileOpen As OPENFILENAME, strFile As String, lResult As Long
    With fileOpen
        .lStructSize = Len(fileOpen) '�ṹ����
        .hWndOwner = hWndOwner
        .flags = 0
        .lpstrFile = Rpad(strFileName, 254) '����Ĭ��Ҫ�����ļ�
        .nMaxFile = 255 '��ʾ�ļ����ĳ���
        .lpstrFileTitle = String$(255, 0) '�򿪶Ի���ı��ⳤ��
        .nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
        .lpstrInitialDir = App.Path
        .lpstrFilter = strFileType '�ļ�����
        .nFilterIndex = 1
        .lpstrTitle = strSaveTitle
        lResult = GetSaveFileName(fileOpen) 'ȡ���ļ���
        If lResult <> 0 Then
            strFile = Split(.lpstrFile, Chr(0))(0)
        Else
            strFile = ""
        End If
    End With
    GetSaveFile = strFile
End Function
Public Function GetOpenFile(ByVal hWndOwner As Long, ByVal strFileType As String, ByVal strTypeFilter As String, strOpenTitle As String) As String
'strTypeFilter ��ʽ "��ʾ����chr(0)*.��������chr(0);��ʾ����chr(0)*.��������chr(0)chr(0)
Dim fileOpen As OPENFILENAME, strFile As String, lResult As Long
    With fileOpen
        .lStructSize = Len(fileOpen) '�ṹ����
        .hWndOwner = hWndOwner
        .flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
        .lpstrFile = Rpad(strFileType, 254)
        .nMaxFile = 255 '��ʾ�ļ����ĳ���
        .lpstrFileTitle = Space(254) '�򿪶Ի���ı��ⳤ��
        .nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
        .lpstrInitialDir = App.Path
        .lpstrFilter = strTypeFilter '�򿪵��ļ�����
        .nFilterIndex = 1
        .lpstrTitle = strOpenTitle '�򿪶Ի���ı���
        lResult = GetOpenFileName(fileOpen) 'ȡ���ļ���
        If lResult <> 0 Then
            strFile = Split(.lpstrFile, Chr(0))(0)
        End If
    End With
    GetOpenFile = strFile
End Function
 
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    Rpad = zl9Comlib.zlStr.Rpad(strCode, lngLen, strChar, True)
End Function
 
Public Function ChkControl(ControlTmp As Object) As Boolean
Dim strName As String
    Err.Clear
    On Error GoTo errHand
    strName = ControlTmp.Name
    ChkControl = True
    Exit Function
errHand:
    Err.Clear
    ChkControl = False
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    Set rs = zlDatabase.OpenSQLRecord("SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlPublic")
    
    GetMaxLength = rs.Fields(0).DefinedSize
    
End Function
Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function
Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'���ܣ���ȡָ����ı�����������ֵ
'������strTable  ����;
'      strField  �ֶ���;
'      intLength �ֶγ���
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo errHand
    gstrSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""���ֵ"",max(length(" & _
         strField & ")) as ""�ֵ"" FROM " & strTable & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    With rsTemp
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("���ֵ").Value), "0", .Fields("���ֵ").Value)
        lngLengh = IIf(IsNull(.Fields("�ֵ").Value), intLength, .Fields("�ֵ").Value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            gstrSQL = "Select ZL_INCSTR([1]) As MAXVALUE From Dual"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", CStr(varTemp))
            If rsTemp.BOF = False Then
                GetMax = Trim(rsTemp("MAXVALUE").Value)
            End If
        End If
        .Close
    End With
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If bytIsWB Then
        gstrSQL = "Select zlWBcode('" & strInput & "') from dual"
    Else
        gstrSQL = "Select zlSpellcode('" & strInput & "') from dual"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    zlGetSymbol = Nvl(rsTmp.Fields(0).Value)
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function
Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    'Or InStr(strInput, ";") > 0 Or InStr(strInput, ",") > 0 Or InStr(strInput, "`") > 0 Or InStr(strInput, """") > 0
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    StrIsValid = True
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
