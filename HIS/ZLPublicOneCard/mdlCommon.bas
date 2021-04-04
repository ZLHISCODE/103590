Attribute VB_Name = "mdlCommon"
Option Explicit
Public glngHook As Long
Public gdtBegin As Date
Public glngOld As Long
Public gblnOk As Boolean

'------------------------------------------------------------------------------------------------------------------------------------
'枚举普量
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum


Public Sub UnHookKBD()
    If glngHook <> 0 Then
    UnhookWindowsHookEx glngHook
    glngHook = 0
    End If
End Sub

Public Function EnableKBDHook()
    If glngHook <> 0 Then
        gdtBegin = Time
        Exit Function
    End If
    gdtBegin = Time
    glngHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHFunc, App.hInstance, App.ThreadID)
End Function

Public Function MyKBHFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Time - gdtBegin) * 60 * 60 * 24 < 0.3 Then
        MyKBHFunc = 1 '表示要处理这个讯息If wParam = vbKeySnapshot Then '侦测 有没有按到PrintScreen键MyKBHFunc = 1 '在这个Hook便吃掉这个讯息End If
    Else
        MyKBHFunc = 0
    End If
    Call CallNextHookEx(glngHook, iCode, wParam, lParam) '传给下一个HookEnd Function
End Function


Public Function NeedName(strList As String, Optional ByVal strSplit As String) As String
'功能：从编码名称组合串中分离出名称
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
'          strSplit=指定的编码名称分割符，没有指定，则按默认优先级进行解析,解析符只能在如下或者其他单个中间分割符
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）>换行(Chr(10))> - > [] > ()
    NeedName = GetNeedName(strList, strSplit)
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

Public Function SetWindowResizeWndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = gWinRect.MinW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = gWinRect.MinH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = gWinRect.MaxW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = gWinRect.MaxH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SetWindowResizeWndMessage = 1
        Exit Function
    End If
    SetWindowResizeWndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function


'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '获取指定窗体的父窗体
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function



Public Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    
    If Trim(objTxt.Text) = "" Then Exit Sub
    
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        
        If TypeName(objTxt) = "TextBox" Then
            If objTxt.MultiLine Then
                SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
            End If
        End If
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub


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


Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：等同Oracle的RPAD函数
'功能:按指定长度填制空格
 '参数：
 '       strText:填充字符串
 '       intCount:填充后的长度
 '       StrPAD:填充的字符
 '       blnAutoSub:字符串超长后自动截取
'返回:返回字串
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：等同Oracle的LPAD函数
 '功能:按指定长度填制空格
 '参数：
 '  strText:填充字符串
 '  intCount:填充后的长度
 '  StrPAD:填充的字符
 '  blnAutoSub:字符串超长后自动截取
 '返回:返回字串
 
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        LPAD = String((intCount - lngTmp) \ lngFill, StrPAD) & strText
    ElseIf lngTmp > intCount And blnAutoSub Then
        LPAD = SubB(strText, 1, intCount)
    Else
        LPAD = strText
    End If
End Function


Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, I As Integer
    
    I = 1
    varValue = arrPar(0)
    Do While I <= UBound(arrPar)
        If I = UBound(arrPar) Then
            Decode = arrPar(I): Exit Function
        ElseIf varValue = arrPar(I) Then
            Decode = arrPar(I + 1): Exit Function
        Else
            I = I + 2
        End If
    Loop
End Function

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
errHand:
End Sub

Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
errHand:
End Sub

Public Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'功能：获取指定控件在屏幕中的位置(Twip/Pixel)
'返回：blnTwip=True-返回Twip单位，False-返回像素单位
    Dim vRect As RECT
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function
Public Function GetIDCardDate(strCardID As String) As String
    '功能：根据身份证号返回出生日期
    '参数：ID=身份证号,应该为15位或18位
    '返回：格式"yyyy-MM-dd"
    Dim strTmp As String
    
    If gobjCommFun Is Nothing Then Call zlInitCommLib
    If Not gobjCommFun Is Nothing Then
       GetIDCardDate = gobjCommFun.GetIDCardDate(strCardID): Exit Function
    End If
    
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
Public Sub FormSetCaption(ByVal objForm As Variant, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：objForm=传窗体对象，可以传窗体句柄（仅隐藏时blnCaption=false)
'         blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    Dim lnghwnd As Long
    If IsObject(objForm) Then
        lnghwnd = objForm.hWnd
    Else
        lnghwnd = objForm
    End If
    
    Call GetWindowRect(lnghwnd, vRect)
    lngStyle = GetWindowLong(lnghwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong lnghwnd, GWL_STYLE, lngStyle
    SetWindowPos lnghwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub


Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由数字和英文字母构成，用于允许数字
    '       和字母但不允许特殊字符的情况下的检测，isnumberic只能判断数字。
    '参数：（SSC编制）
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If Not ((J > 47 And J < 58) Or (J > 64 And J < 91) Or (J > 96 And J < 123)) Then
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
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If Not ((J > 64 And J < 91) Or (J > 96 And J < 123)) Then
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
    Dim I As Integer, J As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For I = 1 To Len(Trim(strAsk))
            J = Asc(Mid(Trim(strAsk), I, 1))
            If J < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub


Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
    '功能:打开中文输入法，或关闭输入法
    '参数：strImeName-打开指定的输入法，没有指定时打开系统选项设置的缺省输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
    If strImeName = "不自动开启" Then OpenIme = True: Exit Function
    If gobjCommFun Is Nothing Then zlInitCommLib
    If Not gobjCommFun Is Nothing Then
        OpenIme = gobjCommFun.OpenIme(blnOpen, strImeName)
        Exit Function
    End If
    
    
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
Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

