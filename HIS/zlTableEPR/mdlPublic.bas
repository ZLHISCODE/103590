Attribute VB_Name = "mdlPublic"
Option Explicit
Public gTargetDC As Long

'######################################################################################
'   全局常数，用于页面显示及滚动条
'######################################################################################

Public Const HSTEP = 50         '滚动条水平步长
Public Const VSTEP = 50         '滚动条垂直步长
Public Const PAGEMARGIN = 200   '页面视图下控件与容器的边距
Public Const SHADOWOFFSET = 30  '阴影偏移量
Public Const WHEELNUMBER = 20   '鼠标滚动系数
'######################################################################################
'获取中英文混合字符串长度
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'繁简转换
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

'鼠标位置信息
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

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态

'######################################################################################
'   大纲模式相关常量声明            RTB SDK 3.0
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

'   是否展开
Public Const EMO_EXPANDSELECTION = 0
Public Const EMO_EXPANDDOCUMENT = 1

Public Const VM_NORMAL = 4             ' // Agrees with RTF \viewkindN
Public Const VM_OUTLINE = 2

'######################################################################################
'   缩放比例相关常量声明            RTB SDK 3.0
'######################################################################################

Public Const EM_GETZOOM = (WM_USER + 224)
Public Const EM_SETZOOM = (WM_USER + 225)
Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long

'######################################################################################
'   鼠标滚动钩子
'######################################################################################

Public Type POINTL
    x As Long
    y As Long
End Type
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public lpPrevWndProc As Long

Public sngX As Single, sngY As Single   '鼠标坐标
Public intShift As Integer              '鼠标按键
Public bWay As Boolean                  '鼠标方向
Public bMouseFlag As Boolean            '鼠标事件激活标志

'######################################################################################
'   获取字符屏幕位置
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
'   直接发送按键的函数
'######################################################################################
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'######################################################################################
'   输入法处理函数
'######################################################################################
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'######################################################################################
'   释放内存
'######################################################################################
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = Nvl(rsTmp!部门ID, 0)
        UserInfo.简码 = Nvl(rsTmp!简码, "")
        UserInfo.姓名 = Nvl(rsTmp!姓名, "")
        UserInfo.用户名 = Nvl(rsTmp!用户名, "")
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
'## 功能：  将工具条A放置到工具条B的同一行
'##
'## 参数：  BarToDock   ：加入的工具栏
'##         BarOnLeft   ：位于左边的工具条
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
'字体列表
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
'对无效字体直接返回宋体
    Err.Clear
    On Error GoTo errFont
    UsableFont = gAllFont("F_" & sFont)
    Exit Function
errFont:
    UsableFont = "宋体"
    Err.Clear
End Function
Public Sub PressKey(bytKey As Byte)
    '功能：向键盘发送一个键,类似SendKey
    '参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
   
Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
    '功能:打开中文输入法，或关闭输入法
    '     根据zlComlib中同名函数更改，并利用ZLHIS软件中的本地参数控制
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    Dim strUser As String
    
    strUser = GetSetting("ZLSOFT", "注册信息\登陆信息", "USER", "")
    '用户没进行设置，就不处理
    strIme = GetSetting("ZLSOFT", "私有全局\" & strUser, "输入法", "")
    If strIme = "" And blnOpen = True Then Exit Function                 '要求打开输入法，但是又没有设置
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否批定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                    Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '不是输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function
 
Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'功能：由ItemData查找ComboBox的索引值
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
Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
    ToVarchar = zl9Comlib.zlStr.ToVarchar(varText, lngLength)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
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
    '剔除控件内容的特殊字符'和%
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
        GetFontSizeChinese = "初号"
    Case 36
        GetFontSizeChinese = "小初"
    Case 26
        GetFontSizeChinese = "一号"
    Case 24
        GetFontSizeChinese = "小一"
    Case 22
        GetFontSizeChinese = "二号"
    Case 18
        GetFontSizeChinese = "小二"
    Case 16
        GetFontSizeChinese = "三号"
    Case 15
        GetFontSizeChinese = "小三"
    Case 14
        GetFontSizeChinese = "四号"
    Case 12
        GetFontSizeChinese = "小四"
    Case 10.5
        GetFontSizeChinese = "五号"
    Case 9
        GetFontSizeChinese = "小五"
    Case 7.5
        GetFontSizeChinese = "六号"
    Case 6.5
        GetFontSizeChinese = "小六"
    Case 5.5
        GetFontSizeChinese = "七号"
    Case 5
        GetFontSizeChinese = "八号"
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
    Case "初号"
        sngNum = 42
    Case "小初"
        sngNum = 36
    Case "一号"
        sngNum = 26
    Case "小一"
        sngNum = 24
    Case "二号"
        sngNum = 22
    Case "小二"
        sngNum = 18
    Case "三号"
        sngNum = 16
    Case "小三"
        sngNum = 15
    Case "四号"
        sngNum = 14
    Case "小四"
        sngNum = 12
    Case "五号"
        sngNum = 10.5
    Case "小五"
        sngNum = 9
    Case "六号"
        sngNum = 7.5
    Case "小六"
        sngNum = 6.5
    Case "七号"
        sngNum = 5.5
    Case "八号"
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
        .lfFaceName = StrConv(tmpFont.Name, vbFromUnicode) & vbNullChar '初始化字体名称，需要从Unicode转换，须以空字符结尾
        .lfItalic = tmpFont.Italic '初始化是否有斜体
        .lfStrikeOut = tmpFont.Strikethrough '初始化是否有删除线
        .lfUnderline = tmpFont.Underline '初始化是否有下划线
        .lfWeight = tmpFont.Weight '初始化字体大小
        .lfCharSet = tmpFont.Charset '初始化字符集
        .lfHeight = -MulDiv(tmpFont.Size, GetDeviceCaps(tmphdc, LOGPIXELSY), 72) '把字体转换为lfHeight，用到公式
    End With
    With cF
        .rgbColors = tmpColor '初始化字体颜色
        .lStructSize = Len(cF)
        .hWndOwner = lngHwnd
        .hInstance = App.hInstance
        .flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_EFFECTS Or CF_LIMITSIZE '见下文所述Flags常数列表
        .lpLogFont = VarPtr(lF) '设置为定义好的LogFont结构地址
        .nSizeMin = 4 '最小字体大小
        .nSizeMax = 200 '最大字体大小
    End With
    If CHOOSEFONT(cF) = 0 Then Exit Function '如果按“取消”则退出过程
    With tmpFont
        .Name = StrConv(lF.lfFaceName, vbUnicode) '设置字体名称
        .Italic = lF.lfItalic '设置是否斜体
        .Strikethrough = lF.lfStrikeOut '设置是否删除线
        .Underline = lF.lfUnderline '设置是否下划线
        .Weight = lF.lfWeight '设置是否粗体
        .Charset = lF.lfCharSet '设置字符集
        .Size = -lF.lfHeight - ((-lF.lfHeight) / 4) - IIf(-lF.lfHeight Mod 4 > 1, 1, 0) '设置字体大小，lfHeight与字号得转换需要用到公式
        tmpColor = cF.rgbColors '设置字体颜色
    End With
    SetFont = True
End Function
Public Function GetSaveFile(ByVal hWndOwner As Long, ByVal strFileName As String, strFileType As String, strSaveTitle As String) As String
Dim fileOpen As OPENFILENAME, strFile As String, lResult As Long
    With fileOpen
        .lStructSize = Len(fileOpen) '结构长度
        .hWndOwner = hWndOwner
        .flags = 0
        .lpstrFile = Rpad(strFileName, 254) '设置默认要保存文件
        .nMaxFile = 255 '显示文件名的长度
        .lpstrFileTitle = String$(255, 0) '打开对话框的标题长度
        .nMaxFileTitle = 255 '打开对话框的标题的长度
        .lpstrInitialDir = App.Path
        .lpstrFilter = strFileType '文件类型
        .nFilterIndex = 1
        .lpstrTitle = strSaveTitle
        lResult = GetSaveFileName(fileOpen) '取得文件名
        If lResult <> 0 Then
            strFile = Split(.lpstrFile, Chr(0))(0)
        Else
            strFile = ""
        End If
    End With
    GetSaveFile = strFile
End Function
Public Function GetOpenFile(ByVal hWndOwner As Long, ByVal strFileType As String, ByVal strTypeFilter As String, strOpenTitle As String) As String
'strTypeFilter 格式 "显示类型chr(0)*.过滤类型chr(0);显示类型chr(0)*.过滤类型chr(0)chr(0)
Dim fileOpen As OPENFILENAME, strFile As String, lResult As Long
    With fileOpen
        .lStructSize = Len(fileOpen) '结构长度
        .hWndOwner = hWndOwner
        .flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
        .lpstrFile = Rpad(strFileType, 254)
        .nMaxFile = 255 '显示文件名的长度
        .lpstrFileTitle = Space(254) '打开对话框的标题长度
        .nMaxFileTitle = 255 '打开对话框的标题的长度
        .lpstrInitialDir = App.Path
        .lpstrFilter = strTypeFilter '打开的文件类型
        .nFilterIndex = 1
        .lpstrTitle = strOpenTitle '打开对话框的标题
        lResult = GetOpenFileName(fileOpen) '取得文件名
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
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
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
'功能：读取指定表的本级编码的最大值
'参数：strTable  表名;
'      strField  字段名;
'      intLength 字段长度
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo errHand
    gstrSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""最大值"",max(length(" & _
         strField & ")) as ""最长值"" FROM " & strTable & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    With rsTemp
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("最大值").Value), "0", .Fields("最大值").Value)
        lngLengh = IIf(IsNull(.Fields("最长值").Value), intLength, .Fields("最长值").Value)
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
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
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
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    'Or InStr(strInput, ";") > 0 Or InStr(strInput, ",") > 0 Or InStr(strInput, "`") > 0 Or InStr(strInput, """") > 0
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function
Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MAXROWS As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    '-----------------------------------------------------------
    '功能： 根据需要显示的图像数量和显示区域，计算可显示图像的行列数。
    '参数： PicCount-图像数量
    '       RegionWidth,RegionHeight-区域宽度高度
    '       Rows,Cols-返回自动排列的行列数
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
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function
