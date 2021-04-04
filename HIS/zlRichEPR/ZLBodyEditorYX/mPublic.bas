Attribute VB_Name = "mPublic"
'#########################################################################
'##模 块 名：mPublic.bas
'##创 建 人：吴庆伟
'##日    期：2005年3月25日
'##修 改 人：
'##日    期：
'##描    述：公共函数声明等
'##版    本：
'#########################################################################

Option Explicit


'##########################################################################################
'## API声明
'##########################################################################################
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'绘制透明图形
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'VB Errors
Private Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines
'Raster Operation Codes
Private Const DSna = &H220326

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer       '捕捉按键状态
    ' Virtual key values
Public Const VK_TAB = &H9

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '系统暂停

'##########################################################################################
'## 图片转化为RTF
'##########################################################################################

Private Type Size
    cx As Long
    cy As Long
End Type
' Used to create the metafile
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Used for creating the temporary WMF file
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MM_TEXT = 1
Public Const MM_LOMETRIC = 2
Public Const MM_HIMETRIC = 3
Public Const MM_LOENGLISH = 4
Public Const MM_HIENGLISH = 5
Public Const MM_TWIPS = 6
Public Const MM_ISOTROPIC = 7
Public Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

'注册表直接操作
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '计算横向和纵向图像数量
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '修正图像的宽高，不能大于最大值
    '如果大于intMaxWidth×intMaxHeight则，按照图像总长宽比，使用小于等于intMaxWidth×intMaxHeight作为新宽高,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '采集图像
    '将图像采集到临时图像集
    For i = 1 To AssembleViewer.Count
        '计算缩放比例 hj修改,解决多图合并时，放大的图象无法真正放大的问题
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
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

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
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

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
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
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = Substr(strCode, 1, lngLen)
    End If
    '取掉最后半个字符
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
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
'################################################################################################################
'## 功能：注册表的直接删除
'##
'## 参数：  Hkey        :主键，如HKEY_CLASSES_ROOT、HKEY_CURRENT_USER、HKEY_LOCAL_MACHINE、HKEY_USERS
'##         strPath     :路径
'##         strValue    :键名
'##
'## 示例：
'## Private Sub Command1_Click()
'##    Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "turkey")
'## End Sub
'################################################################################################################
Public Function RegDelete(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

'################################################################################################################
'## 功能：注册表的直接保存
'##
'## 参数：  Hkey        :主键，如HKEY_CLASSES_ROOT、HKEY_CURRENT_USER、HKEY_LOCAL_MACHINE、HKEY_USERS
'##         strPath     :路径
'##         strValue    :键名
'##         strdata     :值
'##
'## 示例：
'## Private Sub Command1_Click()
'##     Call savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "turkey", "c:\hoho.txt")
'## End Sub
'################################################################################################################
Public Sub RegSaveString(Hkey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub

'################################################################################################################
'## 功能：注册表的直接读取
'##
'## 参数：  Hkey        :主键，如HKEY_CLASSES_ROOT、HKEY_CURRENT_USER、HKEY_LOCAL_MACHINE、HKEY_USERS
'##         strPath     :路径
'##         strValue    :键名
'##
'## 示例：
'## Private Sub Command1_Click()
'##     Text1.Text = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "turkey")
'## End Sub
'################################################################################################################
Public Function RegGetString(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType As Long
    
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                RegGetString = Left$(strBuf, intZeroPos - 1)
            Else
                RegGetString = strBuf
            End If
        End If
    End If
End Function
Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'################################################################################################################
'## 功能：相当于Oracle的NVL，将Null值改成另外一个预设值
'##
'## 参数：  varValue        :变量
'##         DefaultValue    :默认值
'################################################################################################################
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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
Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function
Public Function CheckLen(txt As TextBox, intLen As Integer, Optional txtName As String) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(IIf(txtName = "", txt.Name, txtName), 4) & "只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function
'################################################################################################################
'## 功能：  设置窗体的（Alpha）不透明度
'##
'## 参数：  obj         :窗体对象
'##         Opacity     :不透明度
'################################################################################################################
Public Sub SeekIndexWithNo(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedNo(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlControl.CboSetIndex(objCbo.hWnd, -1)
    End If
End Sub
Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件,blnPreserve--如果找不到匹配项目，则保持原有项目
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedName(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnPreserve = True Then
        If blnEvent = False Then
            Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.ListIndex)
        End If
    Else
        If blnEvent Then
            objCbo.ListIndex = -1
        Else
            Call zlControl.CboSetIndex(objCbo.hWnd, -1)
        End If
    End If
    
End Sub
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
Public Function NeedName(strList As String) As String
    NeedName = zlCommFun.GetNeedName(strList)
End Function
Public Sub SetOpacityForm(obj As Object, Opacity As Byte)
    On Error Resume Next
    Dim ret As Long
    'Set the window style to 'Layered'
    ret = GetWindowLong(obj.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong obj.hWnd, GWL_EXSTYLE, ret
    'Set the opacity of the layered window to 128
    SetLayeredWindowAttributes obj.hWnd, 0, Opacity, LWA_ALPHA
End Sub

'################################################################################################################
'## 功能：  获取Windows默认的临时文件名
'##
'## 参数：  TmpFilePrefix   :临时文件的后缀名
'################################################################################################################
Public Function GetTempName(TmpFilePrefix As String) As String
     Dim TempFileName As String * 256
     Dim X As Long
     Dim DriveName As String
     DriveName = "c:"   '默认取C盘为目标盘符
       X = GetTempFileName(DriveName, TmpFilePrefix, 0, TempFileName)
       GetTempName = Left$(TempFileName, InStr(TempFileName, Chr(0)) - 1)
End Function

'################################################################################################################
'## 功能：  绘制透明图片到指定HDC上（指定透明色）
'##
'## 参数：  hDCDest         :目标绘图对象
'##         (xDest,yDest)   :左上角位置
'##         (Width,Height)  :绘图区域高度、宽度
'##         picSource       :源图片
'##         (XSrc,YSrc)     :源图片偏移位置
'##         clrMask         :透明色(MaskColor)
'##         hPal            :调色板句柄，可选
'##
'## 用法：  PaintTransparentStdPic UserControl.hdc, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
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
'## 功能：  绘制透明图片到指定DC上（指定透明色）
'##
'## 说明：  用于的 PaintTransparentStdPic() 函数调用。
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
'## 功能：  判断是否为编辑键
'##
'## 参数：  KeyAscii        :当前编辑方式；
'##         AllowSubtract   :Insert键是否作为编辑键，可选
'##
'## 返回：  如果是编辑键，则返回 True；否则，返回 False
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
'## 功能：  求控件中指定坐标在屏幕中的位置
'##
'## 参数：  lngHwnd         :控件句柄 hWnd
'##         (lngX,lngY)     :控件中的坐标位置
'##
'## 返回：  返回控件中的坐标在屏幕中的位置，单位：缇
'################################################################################################################
Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'################################################################################################################
'## 功能：  将VB的系统颜色转换为RGB色
'##
'## 参数：  lngColor        :需要转换的系统颜色(COLORREF)
'##
'## 返回：  返回转换后的RGB颜色
'################################################################################################################
Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

'################################################################################################################
'## 功能：  得到指定数字的中文字号
'##
'## 参数：  lngNum      :字体大小（数字）
'################################################################################################################
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
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function
'################################################################################################################
'## 功能：  得到指定中文字号的数字尺寸
'##
'## 参数：  strSize     :中文字号
'################################################################################################################
Public Function GetFontSizeNumber(strSize As String) As Single
    Dim sngNum As Single
    Select Case strSize
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
        sngNum = IIf(Val(strSize) <= 0, 10, Val(strSize))
    End Select
    GetFontSizeNumber = Format(sngNum, "0.0")
End Function

'################################################################################################################
'## 功能：  将标准 stdPicture 图片转换为 Meta 图元文件
'##
'## 参数：  aStdPic         :需转换的标准图片
'##         strDestFileName :转换后的目标 Meta 图元文件名
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
'## 功能：  添加按钮
'################################################################################################################
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
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
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
        
        '标题行复制
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
        
        '数据行复制
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
End Sub

