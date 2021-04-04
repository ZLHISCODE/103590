VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsOS"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"clsString"
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsCommFun
'包含对象：实现算法和特殊功能的函数
'命名规则：操作类型+操作对象
'说明    ：各段的含义与上面类似。因为这类函数功能千差万别，不一定好按上面的规则命名，特殊情况也灵活命名。
'例      ：SetRegValue，CheckStrValidity，ToVarchar，NVL
'--------------------------------------------------------------------------
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'选择文件夹（OpenDir）
Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_STATUSTEXT = &H4
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Window版本函数

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'获取文件图标(GetFIleICON)
Private Type SHFILEINFO
        hIcon As Long                       '  out: icon
        iIcon As Long                       '  out: icon index
        dwAttributes As Long                '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH  '  out: display name (or path)
        szTypeName As String * 80           '  out: type name
End Type
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const DI_NORMAL = &H3
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'该函数确定给定窗口是否是最大化的窗口。
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
'---------------------------------------------------
Private mfrmTipInfo As frmTipInfo

Private Enum ICON_SIZE
    ICON_SMALL = 16
    ICON_LARGE = 32
End Enum

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub PressKeyEx(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
'功能：向键盘发送一个键,类似SendKey
'参数：KeyCode=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4,vbKeyA等
'      Shift=vbKeyControl,vbKeyShift,vbKeyMenu(ALT)，如果是组合，用Or运算
    
    '按下组合键
    If (Shift And vbKeyControl) = vbKeyControl Then
        Call keybd_event(vbKeyControl, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyMenu) = vbKeyMenu Then
        Call keybd_event(vbKeyMenu, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyShift) = vbKeyShift Then
        Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    
    '按下普通键
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
    
    '松开普通键
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    
    '松开组合键
    If (Shift And vbKeyShift) = vbKeyShift Then
        Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
    If (Shift And vbKeyMenu) = vbKeyMenu Then
        Call keybd_event(vbKeyMenu, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
    If (Shift And vbKeyControl) = vbKeyControl Then
        Call keybd_event(vbKeyControl, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
End Sub

Public Sub ShowFlash(Optional strNote As String, Optional frmParent As Object)
    '------------------------------------------------
    '功能： 显示等待的动态窗体
    '参数：
    '   strNote:提示信息
    '   frmParent：关于窗体的父窗体
    '返回：
    '------------------------------------------------
    Dim strStyle As String   '导航台风格
    
    With frmFlash
        If strNote <> "" Then .lbl提示.Caption = strNote
        Err = 0
        On Error Resume Next
        .avi.Open gstrAviPath & "\" & "Findfile.avi"
        If Err <> 0 Then
            .lblFile.Visible = True
        End If
        .Refresh
        
        
        If frmParent Is Nothing Then
            .Show
        Else
            '10674问题 在MDI风格下, 这种显示方式会造成系统跳到另一个程序,影响用户操作.
             strStyle = UCase(gobjComLib.zlDatabase.GetPara("导航台"))
             If strStyle = "ZLMDI" Then
                .Show
             Else
                .Show , frmParent
            End If
        End If
        .Refresh
        If Not .lblFile.Visible Then .avi.Play
    End With
End Sub

Public Sub StopFlash()
    '------------------------------------------------
    '功能： 停止并关闭等待的动态窗体
    '参数：
    '返回：
    '------------------------------------------------
    On Error Resume Next
    frmFlash.avi.Stop
    Unload frmFlash
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If str项目 = "" Then str项目 = "所输入内容"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str项目 & "含有非法字符。", vbExclamation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str项目 & "不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function

Public Function InputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, _
    ByVal blnBrushPassShow As Boolean, Optional blnNumberIsCarded As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
    '入参:KeyAscii=在KeyPress事件中调用的参数
    '       blnBrushPassShow-刷卡是否卡号密文显示
    '       blnNumberIsCard-数字默认为是刷卡,缺省为true,表示数字默认为都是刷卡,false-所有都是按输入速度来判断是否刷卡
    '返回:是刷卡,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-02 00:28:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     '刷卡时含有特殊符号的由调用方取消输入
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '判断是否在刷卡
    '55456:blnNumberIsCard
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) And blnNumberIsCarded Then  '姓名输入框如果输的是全数字，认为是刷卡
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
        End If
    End If
    
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(blnBrushPassShow, "*", "")
    Else
        txtInput.PasswordChar = ""
    End If
    InputIsCard = blnCard
End Function


Public Function IntIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
'功能：检查字符串是否合法的整数。
'参数：strInput        输入的字符串
'      intMax          整数的位数
'      blnNegative     是否进行负数检查
'      blnZero         是否进行零的检查


    Dim dblValue As Double
    
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "请输入整数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax Then
        MsgBox str项目 & "数值过大，不能超过" & intMax & "位。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If Int(dblValue) <> dblValue Then
        MsgBox str项目 & "请输入整数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If blnNegative = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    IntIsValid = True
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte, Optional lngLen As Long = 10) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '          lngLen=<1或>40返回的字符串长度为10，否则返回指定长度字符串
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode([1], [2]) from dual"
    Else
        strSQL = "select zlSpellcode([1], [2]) from dual"
    End If
    On Error GoTo Errhand
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlGetSymbol", strInput, lngLen)
    zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function SpellCode(ByVal strAsk As String) As String
'-------------------------------------------------------------
'功能：返回指定字符串的简码
'说明：根据指定字符串生成简码，可以生成三种类型的简码
'        0、拼音，取每字的首字母构成简码
'        1、五笔，取每字的首字母构成简码
'        2、五笔，按五笔规则构成简码
'      在传入的参数中未发现※符号，就按用户在系统选项中设置的方式生成简码；
'        否则就按在※符号后的数字指定的方式强制生成简码，如※1表示按五笔首字母生成
'-------------------------------------------------------------
    Dim lngSplit As Long, lng方式 As Long
    
    lngSplit = InStr(strAsk, "※")
    If lngSplit = 0 Then
        '按用户在系统选项中设置的方式生成简码
        lng方式 = Val(gobjComLib.zlDatabase.GetPara("简码方式"))
    Else
        '指定的方式强制生成简码
        lng方式 = Val(Mid(strAsk, lngSplit + 1))
        strAsk = Mid(strAsk, 1, lngSplit - 1)
    End If
    
    Select Case lng方式
        Case 1, 2
            SpellCode = gobjComLib.zlStr.WBCode(strAsk, lng方式)
        Case Else
            SpellCode = gobjComLib.zlStr.PinYinCode(strAsk)
    End Select
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由数字和英文字母构成，用于允许数字
    '       和字母但不允许特殊字符的情况下的检测，isnumberic只能判断数字。
    '参数：（SSC编制）
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 47 And j < 58) Or (j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsNumOrChar = False
                Exit Function
            End If
        Next
    End If
    IsNumOrChar = True

End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由英文字母构成    '
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否含有汉字
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function UppeMoney(curMoney) As String
    '-------------------------------------------------------------
    '功能：将指定的金额数值转换为大写金额显示
    '参数：
    '       curMoney:需要转换的金额数值
    '返回：
    '-------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
   
    strSQL = "Select zlUppMoney([1]) as NUM From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "UppeMoney", Val(curMoney))
    UppeMoney = rsTmp.Fields(0).value
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function ZyMod(X1, X2) As Long
'-------------------------------------------------------------
'功能：取余函数，由于vb的Mod函数仅仅对亿位以下起作用，超过时出错
'参数：
'       X1-除数；X2-被除数
'返回：
'-------------------------------------------------------------
    ZyMod = X1 - Int(X1 / X2) * X2
End Function

Public Function SelOper(ByVal frmParent As Object) As String
'-------------------------------------------------------------
'功能：显示人员选择器，按部门进行选择
'-------------------------------------------------------------
    Dim frmTemp As New frmManSelect
    frmTemp.Show vbModal, frmParent
    SelOper = frmTemp.mstrReturn
    
    Set frmTemp = Nothing
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
'功能:打开中文输入法，或关闭输入法
'参数：strImeName-打开指定的输入法，没有指定时打开系统选项设置的缺省输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
    If strImeName = "不自动开启" Then OpenIme = True: Exit Function
    '用户没进行设置，就不处理
    If blnOpen Then
        If strImeName <> "" Then
            strIme = Trim(strImeName)
        Else
            strIme = Trim(gobjComLib.zlDatabase.GetPara("输入法"))
        End If
        If strIme = "" Or strIme = "不自动开启" Then Exit Function                '要求打开输入法，但是又没有设置
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否指定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '不是中文输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是1的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function


Public Function GetLike(ByVal strTable As String, ByVal strField As String, ByVal strInput As String) As String
'功能：产生指定字段的Like条件子句
'参数：strTable     表名，与SQL语句相关，可以为别名
'      strField     字段名
'      strInput     输入的字符串
    Dim strMatch  As String
    
    strMatch = IIf(Val(gobjComLib.zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    strInput = UCase(Replace(Trim(strInput), "'", "''"))
    
    GetLike = "Upper(" & IIf(strTable = "", "", strTable & ".") & strField & _
                   ") LIKE '" & strMatch & strInput & "%'"
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long, Optional strDefault As String = "") As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), strDefault, varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal bln时 As Boolean = False) As String
'功能：为不全的日期信息补充完整
'参数：strOrin  用户输入的原始值
'      bln时    是否增加小时部分
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            '使用 . 隔
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        '没有"-",手工加上
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If bln时 = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '处理小时
            If InStr(strTemp, " ") > 0 Then
                '输入了小时
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:mm")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:mm")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'功能：选择Windows的文件夹
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    gstrAPIPath = strInitDir & Chr(0)
    With tBrowseInfo
       .hwndOwner = hwndOwner
       .lpszTitle = lstrcat(strTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf OpenDirCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH * 2)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
       OpenDir = sBuffer
    End If
End Function

Private Function AddressOfFunction(Address As Long) As Long
'功能：OpenDir子函数
    AddressOfFunction = Address
End Function

Public Function IncStr(ByVal strVal As String, Optional ByVal blnDown As Boolean, Optional ByRef strErr As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
'         blnDown=True ,减一，False,加一
    Dim strValuse As String
    Dim intAdd As Integer
    Dim intUp As Integer
    Dim strValue As String
    Dim strValueOne As String
    Dim i  As Integer
    
    On Error GoTo ErrH
    
    strVal = UCase(strVal)
    If Not blnDown Then
        '加1
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = 1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp < 10 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "0" & strValue
                    intUp = 1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp <= Asc("Z") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "A" & strValue
                    intUp = 1
                End If
            End If
        Next
        If intUp = 1 Then
            If IsNumeric(strValueOne) Then
                strValue = "1" & strValue
            Else
                strValue = "A" & strValue
            End If
        End If
        IncStr = strValue
    Else
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = -1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp >= 0 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
'                    If intAdd = 0 Then
                        strValue = "9" & strValue
'                    End If
                    intUp = -1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp >= Asc("A") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    If intAdd = 0 Then
                        strValue = "Z" & strValue
                    End If
                    intUp = -1
                End If
            End If
        Next
        If intUp = 1 Then
            strValue = -1
        End If
        If Mid(strValue, 1, 1) = "0" Or Mid(strValue, 1, 1) = "A" Then
            strValue = Mid(strValue, 2)
        End If
        IncStr = strValue
    End If
    Exit Function
ErrH:
    strErr = "出错函数(IncStr),出错信息:" & Err.Number & " " & Err.Description
    Err.Clear
End Function

Public Function GetNeedName(strList As String, Optional ByVal strSplit As String) As String
'功能：从编码名称组合串中分离出名称
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
'          strSplit=指定的编码名称分割符，没有指定，则按默认优先级进行解析,解析符只能在如下或者其他单个中间分割符
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）>换行(Chr(10))> - > [] > ()
    Dim intType As Integer
    
    intType = gobjComLib.Decode(strSplit, "", 0, Chr(13), 1, Chr(10), 2, "-", 3, "[]", 4, "()", 5, 6)
    
    If intType = 0 Or intType = 1 Then
        '优先判断以回车符分割
        If InStr(strList, Chr(13)) > 0 Then
            GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 2 Then
        '以换行符分割
        If InStr(strList, Chr(10)) > 0 Then
            GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(10)) + 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 4 Then
        '以[]分割
        If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
                GetNeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
                Exit Function
            End If
        End If
    End If
    
    If intType = 0 Or intType = 5 Then
        '以()分割
        If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
                GetNeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
                Exit Function
            End If
        End If
    End If
    If intType = 0 Or intType = 3 Then
        '以-分割
        GetNeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    Else
        GetNeedName = LTrim(Mid(strList, InStr(strList, strSplit) + IIf(InStr(strList, strSplit) = 0, 1, Len(strSplit))))
    End If
End Function

Public Function GetFullNo(strNO As String, Optional ByVal strNum As String = -1, _
                Optional ByVal lng科室ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As String
'功能：由用户输入的部份单号，返回当年的完整单据号。
'参数：strNO=由用户输入的单据号部分,strNum=序号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If strNum = -1 Then
        If Len(strNO) >= 8 Then GetFullNo = Right(strNO, 8): Exit Function
        GetFullNo = gobjComLib.zlStr.PrefixNO & Format(strNO, "0000000")
    Else
        On Error GoTo Errhand
        'If Not IsNumeric(strNO) Then GetFullNo = strNO: Exit Function '非数字直接返回,100306
        strSQL = "Select FullNO([1],[2],[3],[4],[5]) as NO From Dual"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "GetFullNo", strNum, strNO, lng科室ID, strTag, intStep)
        
        If gcnOracle.Errors.count > 0 Then 'Select中函数出错时,在VB中不自动触发错误
            Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
        End If
        
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!NO) Then GetFullNo = rsTmp!NO
        End If
    End If
    
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetIDCardDate(strCardID As String) As String
'功能：根据身份证号返回出生日期
'参数：ID=身份证号,应该为15位或18位
'返回：格式"yyyy-MM-dd"
    Dim strTmp As String
    
    If Len(strCardID) = 15 Then
        strTmp = Mid(strCardID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(strCardID) = 18 Then
        strTmp = Mid(strCardID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDCardDate = strTmp
End Function

Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String

    osvi.dwOSVersionInfoSize = Len(osvi)

    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If

    GetWinPlatform = osvi.dwPlatformId
End Function

Public Function IsWindowsNT() As Boolean
'功能：判断是否WindowNT操作系统
    Const dwMaskNT = &H2&

    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows9X() As Boolean
'功能：判断是否Window95操作系统
    Const dwMask95 = &H1&

    IsWindows9X = (GetWinPlatform() And dwMask95)
End Function

Public Function mGetFullPY(strText As String, Optional intCapital As Integer = 0, _
    Optional blnUseSpliter As Boolean = True) As String
    '------------------------------------------------
    '功能：根据GBK编码提取汉字的全拼音
    '参数： strText         要取得全拼的字符串
    '       intCapital      大小写标记：0-大写；1-小写；2-首字母大写
    '       blnUseSpliter   拼音之间是否返回空格：True-带空格；False-不带空格
    '返回：             返回取得的全拼如果没有则为空
    '------------------------------------------------
    
    Dim arrMusicCode() As String    '拼音表
    Dim arrPYCodeIndex() As String  'ASCII转拼音映射表
    Dim strChar As String
    Dim hiByte As Integer
    Dim lowByte As Integer
    Dim strOneWord As String
    Dim intIndex As Integer     '拼音索引
    Dim strPYCode As String
    Dim strResult As String
    
    Dim i As Integer
    
    On Error Resume Next
    
    '从资源文件提取GBK字库数据
    arrMusicCode = Split(LoadResString(101), ",")
    strPYCode = LoadResString(102) & LoadResString(103) & LoadResString(104)
    arrPYCodeIndex = Split(strPYCode, ";")
        
    '分析字符串
    For i = 1 To Len(strText)
        strChar = Mid(strText, i, 1)
        
        '提取ASCII码的高地位，处理四舍五入的问题
        If Asc(strChar) And &HFF& = 0 Then
            hiByte = Asc(strChar) \ &H100 And &HFF&
        Else
            hiByte = (Asc(strChar) \ &H100 And &HFF&) - 1
        End If
    
        lowByte = Asc(strChar) And &HFF&
        
        ' 是否为 GBK 字符
        '一个汉字双字节，当两个ASCII码分别大于十六进制的81，40，才是汉字编码的区域
        If hiByte >= &H81 And lowByte >= &H40 Then
                '获得拼音索引,GBK汉字编码从81组开始
                intIndex = Split(arrPYCodeIndex(hiByte - &H81), ",")(lowByte - &H40)
                
                If intIndex = 0 Then    '无此汉字, 不能翻译的字符, GBK 保留
                    strOneWord = strChar
                Else
                    strOneWord = arrMusicCode(intIndex - 1)
                    '首字母大写
                    If intCapital = 2 Then
                        strOneWord = UCase(Left(strOneWord, 1)) & LCase(Mid(strOneWord, 2))
                    End If
                End If
                strResult = strResult & IIf(blnUseSpliter = True, " ", "") & strOneWord
        Else    '不是双字节字符
            strOneWord = strChar
            strResult = strResult & strOneWord
        End If
    Next i
    If intCapital = 0 Then
        mGetFullPY = UCase(Trim(strResult))
    ElseIf intCapital = 1 Then
        mGetFullPY = LCase(Trim(strResult))
    Else
        mGetFullPY = Trim(strResult)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lnghwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lnghwnd, GWL_EXSTYLE, lngStyle)
End Sub

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, _
    Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String, Optional ByVal strSelectCaption As String, _
    Optional ByRef strSelectInput As String, Optional ByVal strTextCaption As String, _
    Optional ByVal lngTextLength As Long, Optional ByRef strTextInput As String, Optional ByVal strSort As String = "1,2,3", _
    Optional ByVal blnSelectMust As Boolean) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=传入的日期的标题，如果<>""则显示日期控件，供用户输入日期，将DateInput返回。
'      strDateFormat=时间格式 格式""yyyy-MM-dd hh:mm:ss"其中HH为大写是24小时制"
'      strSelectCaption=选项的标题:选项1|1(1设为缺省),选项2|0|1(选择本选项时才允许填写，1代表日期，2代表文本框，0代表所有)···
'      strSelectInput=选项选择的返回值(返回选择的名称)
'      strTextCaption=文本框标题
'      lngTextLength=文本框最大录入长度
'      strTextInput=文本框的返回值
'      strSort=日期、选项、文本框的排序，日期=1，选项=2，文本=3，默认排序"1,2,3"
'      blnSelectMust=如果有单选项，则必须选择一个，否则提示。
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    ShowMsgBox = frmMsgBox.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle, strDateCaption, DateInput, strDateFormat, strSelectCaption, strSelectInput, strTextCaption, lngTextLength, strTextInput, strSort, blnSelectMust)
End Function

Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, _
    Optional ByVal strName As String, Optional ByVal strSex As String, _
    Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
'功能：对密码进行验证
'参数：frmParent=显示的父窗体
'      strPass=正确的密码
'      strName,strSex,strOld=可选参数，病人姓名、性别、年龄，当不传入时不显示这个区域。
'      blnPassEncode-strPass是否传入的加密串
'返回：True=密码验证通过,False=取消输入，或连续3次输入错误的密码
    VerifyPassWord = frmVerifyPassword.ShowMe(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
End Function

Public Sub DeviceSetup(frmParent As Object, ByVal lngSys As Long, ByVal lngModule As Long)
'功能:根据不同的系统和模块,调用不同的设备配置
    
    Call frmDeviceSetup.ShowMe(frmParent, lngSys, lngModule)
End Sub

Public Sub ShowTipInfo(ByVal lnghwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      strInfo=提示信息,当传入为空时隐藏提示
'      blnMultiRow=以一定的间距分行显示多行信息，每行按vbcrlf分隔
'      blnOutline=是否将每行文本中字符|前的文字做为提纲单独一行显示
'      lngMaxWidth=窗口的最大窗度，缺省为0表示按设计状态的窗体最大宽度为准

    Call mfrmTipInfo.ShowTipInfo(lnghwnd, strInfo, blnMultiRow, blnOutline, lngMaxWidth)
End Sub

Public Sub ShowChildWindow(ByVal lnghwnd As Long, Optional ByVal lngMainHwnd As Long, Optional ByVal blnMaximized As Boolean)
'功能：显示指定的窗体，以子窗体方式
'参数：lngHwnd=要作为子窗体显示的窗体的句柄
'      lngMainHwnd=父窗体句柄，不传时表明不以它的子窗体显示
'      blnMaximized=是否固定将窗体最大化处理；因为API设置为子窗体后，窗体原有的最大化特性将失效
'说明：该项函数主要用于在ZLBH中融合调用ZLHIS窗体显示
    Dim vRect1 As RECT, vRect2 As RECT
    Dim X As Long, Y As Long, lngHeight As Long, lngWidth As Long
    Dim blnCurZoomed As Boolean
    
    On Error Resume Next
    If lngMainHwnd <> 0 Then
        SetParent lnghwnd, lngMainHwnd
    End If
    If IsZoomed(lnghwnd) <> 0 Then
        blnCurZoomed = True
    End If
    If blnMaximized Then
        If blnCurZoomed Then '最大化窗口通过此处触发Resize
            ShowWindow lnghwnd, SW_RESTORE
        End If
        ShowWindow lnghwnd, SW_SHOWMAXIMIZED
    Else
        If lngMainHwnd <> 0 Then
            '显示在父窗体的中央
            If IsWindowVisible(lnghwnd) = 0 Then
                vRect1 = gobjComLib.zlControl.GetControlRect(lnghwnd, False)
                vRect2 = gobjComLib.zlControl.GetControlRect(lngMainHwnd, False)
                lngHeight = vRect1.Bottom - vRect1.Top + IIf(blnCurZoomed, 0, 1)
                lngWidth = vRect1.Right - vRect1.Left + IIf(blnCurZoomed, 0, 1)
                X = ((vRect2.Right - vRect2.Left) - lngWidth) / 2
                Y = ((vRect2.Bottom - vRect2.Top) - lngHeight) / 2
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                '非最大化窗口通过此处触发Resize
                SetWindowPos lnghwnd, HWND_TOP, X, Y, lngWidth, lngHeight, SWP_SHOWWINDOW
                If blnCurZoomed Then '最大化窗口通过此处触发Resize
                    ShowWindow lnghwnd, SW_RESTORE
                    ShowWindow lnghwnd, SW_SHOWMAXIMIZED
                    '最大化后，可能位置发生变化，重新设置一次
                    SetWindowPos lnghwnd, HWND_TOP, X, Y, lngWidth, lngHeight, SWP_SHOWWINDOW
                End If
            End If
        Else
            '没有主窗体的，显示在主窗体中央
            If IsWindowVisible(lnghwnd) = 0 Then
                vRect1 = gobjComLib.zlControl.GetControlRect(lnghwnd, False)
                lngHeight = vRect1.Bottom - vRect1.Top + IIf(blnCurZoomed, 0, 1)
                lngWidth = vRect1.Right - vRect1.Left + IIf(blnCurZoomed, 0, 1)
                X = (Screen.Width / Screen.TwipsPerPixelX - lngWidth) / 2
                Y = (Screen.Height / Screen.TwipsPerPixelY - lngHeight) / 2
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                '非最大化窗口通过此处触发Resize
                SetWindowPos lnghwnd, HWND_TOP, X, Y, lngWidth, lngHeight, SWP_SHOWWINDOW
                If blnCurZoomed Then '最大化窗口通过此处触发Resize
                    ShowWindow lnghwnd, SW_RESTORE
                    ShowWindow lnghwnd, SW_SHOWMAXIMIZED
                    '最大化后，可能位置发生变化，重新设置一次
                    SetWindowPos lnghwnd, HWND_TOP, X, Y, lngWidth, lngHeight, SWP_SHOWWINDOW
                End If
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function zlStringEncode(ByVal strPutString As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:字符串加密
    '入参:strPutString-需要加密的串
    '出参:
    '返回:加密串
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPutString = "" Then Exit Function
    zlStringEncode = Md5_String_Calc(strPutString)
End Function

Public Function PDFPrinterDeviceName() As String
'******************************************************************************************************************
'功能：获取当前PDF打印机名称，由于不同部件当前打印机不一致，因此需要传递。只要PDFInitialize调用成功，就可以获取到，否则返回空串
'******************************************************************************************************************
    If Not gclsPDF Is Nothing Then
        PDFPrinterDeviceName = gclsPDF.PrinterDeviceName
    End If
End Function


Public Function PDFInitialize(Optional ByRef strErrorDesc As String) As Boolean
    '******************************************************************************************************************
    '功能：初始化，并返回是否可以正常输出PDF
    '参数：strErrorDesc=不成功时返回的错误信息内容
    '返回：返回True表示能正常输出PDF文件，False表示不能正常输出PDF文件
    '******************************************************************************************************************
    On Error GoTo Errhand
    
    If gclsPDF Is Nothing Then Set gclsPDF = New clsPDF
    
    If gclsPDF Is Nothing Then
        strErrorDesc = "未能实例化PDF输出类，不能输出PDF文件！"
        Exit Function
    End If
    
    If gclsPDF.PDFInitialize = False Then
        strErrorDesc = gclsPDF.GetLastError
        Exit Function
    End If
    
    PDFInitialize = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErrorDesc = Err.Description
    Err.Clear
End Function

Public Function PDFFile(ByVal strPDFFile As String, _
                        Optional ByVal blnCopyable As Boolean = False, _
                        Optional ByVal blnEditable As Boolean = False, _
                        Optional ByVal blnPrintable As Boolean = False, _
                        Optional ByVal strPassWord As String = "", _
                        Optional ByRef strErrorDesc As String) As Boolean
    '******************************************************************************************************************
    '功能：设置要生成的PDF文件及相关属性，对于Foxit虚拟打印，必须指定文件名，并且只支持指定文件名，不支持其他参数。
    '参数：strPDFFile=输出文件名，包含文件路径和文件扩展名
    '                 文件路径必须存在，且会自动覆盖同名文件
    '                 如果未指定，则弹出文件保存对话框
    '      blnCopyable=输出的PDF文件是否允许复制内容
    '      blnEditable=输出的PDF文件是否允许编辑内容
    '      blnPrintable=输出的PDF文件是否允许打印输出
    '      strPassword=是否要求输入密码
    '      strErrorDesc=不成功时返回的错误信息内容
    '返回：设置成功返回True；否则返回False
    
    '******************************************************************************************************************
        
    On Error GoTo Errhand
    
    If gclsPDF Is Nothing Then
        strErrorDesc = "未能实例化PDF输出类，不能输出PDF文件！"
        Exit Function
    End If
    
    If gclsPDF.PDFFile(strPDFFile, blnCopyable, blnEditable, blnPrintable, strPassWord, "") = False Then
        strErrorDesc = gclsPDF.GetLastError
        Exit Function
    End If
    
    PDFFile = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErrorDesc = Err.Description
    Err.Clear
End Function
 
Public Function PDFFileSuccess(Optional ByRef strErrorDesc As String) As Boolean
    '******************************************************************************************************************
    '功能：PDF输出文件是否成功。foxit必须调用该方法，才能生成文件。
    '返回：True,文件生成成功，False-文件生成失败
    '说明：该函数在打印输出完成后调用
    '******************************************************************************************************************
        
    On Error GoTo Errhand
    If gclsPDF Is Nothing Then
        strErrorDesc = "未能实例化PDF输出类，不能输出PDF文件！"
        Exit Function
    End If
    PDFFileSuccess = gclsPDF.PDFFileSuccess
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErrorDesc = Err.Description
    Err.Clear
End Function
 
 
Private Sub Class_Initialize()
    If mfrmTipInfo Is Nothing Then
        Set mfrmTipInfo = New frmTipInfo
    End If
End Sub

Private Sub Class_Terminate()
    If Not mfrmTipInfo Is Nothing Then
        Unload mfrmTipInfo
        Set mfrmTipInfo = Nothing
    End If
End Sub

Public Function GetPubIcons() As ImageManagerIcons
    Set GetPubIcons = frmPubIcons.imgPublic.Icons
End Function

Public Function GetPaitSignImageList(Optional ByVal bytSize As Byte = 0) As Object
'功能：获取病人标记设置公共图标列表
'参数：bytSize=0小图标(16*16);bytSize=1大图标(24*24)
    Set GetPaitSignImageList = frmPubIcons.Img标记(IIf(bytSize = 0, 999, 1))
End Function

Public Function GetFileIcon(ByVal strFile As String, ByVal blnSmallSize As Boolean, ByVal lngAPPHhInstance As Long) As StdPicture
'功能：返回指定文件的大图标或小图标
'参数：strFile，包含后缀的文件名；当文件真实文件时，应该包含完整的路径名
'      blnSmallSize，是否是小图标
'      lngAPPHhInstance：调用功能的App.hInstance
    Dim fInfo As SHFILEINFO
    Dim objText As TextStream
    Dim blnExists As Boolean, lngRetu As Long
    Dim intSize As ICON_SIZE
    Dim objFile As New FileSystemObject     '文件操作对象
    Dim objPic  As PictureBox
    
    Set objPic = frmPubTmp.picDraw
    
    intSize = IIf(blnSmallSize, ICON_SMALL, ICON_LARGE)
    blnExists = objFile.FileExists(strFile)
    If Not blnExists Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder).Path & "\" & lngAPPHhInstance & Int(timer) & strFile
        Set objText = objFile.CreateTextFile(strFile)
    End If

    If intSize = ICON_LARGE Then
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_LARGEICON)
    Else
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_SMALLICON)
    End If
    If Not blnExists Then
        objText.Close
        objFile.DeleteFile strFile, True
    End If
    
    objPic.Width = intSize * Screen.TwipsPerPixelX
    objPic.Height = intSize * Screen.TwipsPerPixelY
    objPic.Cls
    If lngRetu <> 0 Then
        DrawIconEx objPic.hDC, 0, 0, fInfo.hIcon, intSize, intSize, 0, 0, DI_NORMAL
        DestroyIcon fInfo.hIcon
    End If
    Set GetFileIcon = objPic.Image
End Function

Public Function GetWeekBase(ByVal datDate As Date, Optional ByVal blnNextWeek As Boolean) As Date
'功能：获取指定时间所在星期的星期一的时间
'参数：blnNextWeek=是否获取下一周的周一
    'Oracle:Select Sysdate-(To_Number(To_Char(Sysdate,'D'))-1)+1 From Dual
    GetWeekBase = Format(datDate - (Weekday(datDate, vbMonday) - 1 + IIf(blnNextWeek, 7, 0)), "yyyy-MM-dd 00:00:00")
End Function

Public Function TruncateDate(ByVal datFull As Variant) As Date
'去掉日期中的时、分、秒
    If TypeName(datFull) = "String" Then
        If IsDate(datFull) Then
            TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
        Else
            TruncateDate = CDate(Format(CDate(0), "yyyy-MM-dd"))
        End If
    ElseIf TypeName(datFull) = "Date" Then
        TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
    Else
        TruncateDate = CDate(Format(CDate(0), "yyyy-MM-dd"))
    End If
End Function

Public Function Identity(ByRef lngCount As Long) As Long
'功能：模拟主键自增
'参数：lngCount=自增变量
    lngCount = lngCount + 1
    Identity = lngCount
End Function

Public Function DblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
 '功能:检查字符串是否合法的金额
'入参:strInput        输入的字符串
'     intMax          整数的位数
'     blnNegative     是否进行负数检查
'     blnZero         是否进行零的检查
'出参:
'返回:成功,返回true,否则返回False

    Dim dblValue As Double
    
    If blnZero = True Then
        If strInput = "" Then
            MsgBox str项目 & "未输入，请检查!", vbInformation, gstrSysName
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then DblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    DblIsValid = True
End Function


