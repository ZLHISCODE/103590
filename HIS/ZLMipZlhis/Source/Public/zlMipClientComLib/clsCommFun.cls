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
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsCommFun
'包含对象：实现算法和特殊功能的函数
'命名规则：操作类型+操作对象
'说明    ：各段的含义与上面类似。因为这类函数功能千差万别，不一定好按上面的规则命名，特殊情况也灵活命名。
'例      ：SetRegValue，CheckStrValidity，ToVarchar，NVL
'--------------------------------------------------------------------------
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_MULTI_SZ = 7
Private Const REG_SZ = 1
Private Const KEY_QUERY_VaLUE = &H1
Private Const KEY_SET_VaLUE = &H2
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

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


Public Enum emDataFormat
    IP4
    IdCard
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

'######################################################################################################################

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

Public Function GetRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String, strValue As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim strData As String
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    strData = Space(255)
    lngLength = 255
    lngType = REG_SZ
    lngReturn = RegQueryValueEx(lngKey, strValueName, 0, lngType, strData, lngLength)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    strValue = Mid(strData, 1, lngLength)
    
    If InStr(strValue, Chr(0)) > 0 Then
        strValue = Mid(strValue, 1, InStr(strValue, Chr(0)) - 1)
    End If
    RegCloseKey lngKey
    GetRegValue = True
End Function

Public Function SetRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String, ByVal strValue As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    lngLength = LenB(StrConv(strValue, vbFromUnicode))
    lngType = REG_SZ
    lngReturn = RegSetValueEx(lngKey, strValueName, 0, lngType, strValue, lngLength)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    SetRegValue = True
End Function

Public Function DeleteRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    DeleteRegValue = True
End Function

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
    ByVal blnBrushPassShow As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
    '入参:KeyAscii=在KeyPress事件中调用的参数
    '       blnBrushPassShow-刷卡是否卡号密文显示
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
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then  '姓名输入框如果输的是全数字，认为是刷卡
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

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "Select zlWBcode([1]) From Dual"
    Else
        strSQL = "Select zlSpellcode([1]) From Dual"
    End If
    On Error GoTo errHand
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlGetSymbol", strInput)
    zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)
    Exit Function
errHand:
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
            SpellCode = mWBX(strAsk, lng方式)
        Case Else
            SpellCode = mPinYin(strAsk)
    End Select
End Function

Private Function mPinYin(ByVal strAsk As String) As String
'功能：返回指定字符串的拼音简码
'参数：strAsk  待处理的字符串

    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer
    Dim strCode As String, StrBit As String

'    aryStard = Split("八;擦;哒;讹;发;噶;哈;击;;咔;垃;妈;拿;噢;啪;七;热;撒;他;挖;;挖;西;丫;匝", ";")
    aryStard = Split("八;擦;咑;屙;发;旮;铪;讥;;咔;垃;妈;拿;噢;啪;七;呥;撒;他;挖;;;西;丫;匝", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '将全角转换为半角，小写转换为大写
    
    strCode = ""
    For intBit = 1 To Len(strAsk)
        StrBit = Mid(strAsk, intBit, 1)
        If InStr(1, "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩαβγ蔫趴属哇娃夕汐仨兮拚嚓饧澶赕膪欹焘恁砉铛疋覃瞿她", StrBit) > 0 Then
            '特殊字的处理
            strCode = strCode & Switch(StrBit = "Ⅰ", "1", StrBit = "Ⅱ", "2", StrBit = "Ⅲ", "3", StrBit = "Ⅳ", "4", StrBit = "Ⅴ", "5" _
                            , StrBit = "Ⅵ", "6", StrBit = "Ⅶ", "7", StrBit = "Ⅷ", "8", StrBit = "Ⅸ", "9" _
                            , StrBit = "α", "A", StrBit = "β", "B", StrBit = "γ", "G" _
                            , StrBit = "蔫", "N", StrBit = "趴", "P", StrBit = "属", "S", StrBit = "哇", "W" _
                            , StrBit = "娃", "W", StrBit = "夕", "X", StrBit = "汐", "X", StrBit = "仨", "S" _
                            , StrBit = "兮", "X", StrBit = "拚", "P", StrBit = "嚓", "C", StrBit = "饧", "X" _
                            , StrBit = "澶", "C", StrBit = "赕", "D", StrBit = "膪", "C", StrBit = "欹", "Q" _
                            , StrBit = "焘", "T", StrBit = "恁", "N", StrBit = "砉", "H", StrBit = "铛", "D" _
                            , StrBit = "疋", "P", StrBit = "覃", "Q", StrBit = "瞿", "Q", StrBit = "她", "T")
        ElseIf Asc(StrBit) < 0 Then
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(StrBit, aryStard(iCount), vbTextCompare) = -1 Then
                        strCode = strCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        strCode = strCode & "Z"
                    End If
                End If
            Next
        Else
            If StrBit >= "A" And StrBit <= "Z" Then
                strCode = strCode & StrBit
            End If
        End If
        If Len(strCode) >= 10 Then Exit For
    Next
    mPinYin = strCode

End Function

Private Function mWBX(ByVal strAsk As String, ByVal lng方式 As Long) As String
'功能：返回指定字符串的五笔型简码
'参数：strAsk  待处理的字符串
'      lng方式 1-取首字母，2-按五笔规则
    Static blnNotFound As Boolean
    Dim lngFile As Long, strFile As String, strReturn As String
    Dim str编码表 As String, str汉字 As String, bln前字母 As Boolean, str编码 As String
    Dim intBit As Integer, StrBit As String
    
    If blnNotFound = True Then
        'wbx.txt文件未找到，不能进行编码查询
        Exit Function
    End If
    
    '打开文件
    strFile = gstrAviPath
    If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
'    strFile = "C:\AppSoft\"
    strFile = strFile & "wbx.txt"
    
    On Error Resume Next
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    If Err <> 0 Then
        blnNotFound = True
        MsgBox "未发现" & strFile & "文件。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '找到每一个字对应编码
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        If InStr(strAsk, Left(strReturn, 1)) > 0 Then
            '把这个判断放在内部，主要是为了加快速度，因为只有字经过第一个判断
            If InStr(strReturn, " ") > 0 Then
                str编码表 = str编码表 & strReturn & "|"
            End If
        End If
    Loop
    Close #lngFile
    str编码表 = UCase(str编码表)
    
    '得到字符串的中汉字
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '将全角转换为半角，将字符串文字转成小写
    If lng方式 = 1 Then
        '按首字母
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '汉字
                str编码 = str编码 & mGet编码By汉字(str编码表, StrBit, 1)
                bln前字母 = False
            ElseIf InStr(" ,.;:", StrBit) > 0 Then
                '空格
                bln前字母 = False
            Else
                If bln前字母 = False And StrBit >= "A" And StrBit <= "Z" Then
                    '只取一个字符串的首字母
                    str编码 = str编码 & StrBit
                End If
                bln前字母 = True
            End If
        Next
    Else
        '按五笔规则
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '汉字
                str汉字 = str汉字 & StrBit
            End If
        Next
        
        Select Case Len(str汉字)
            Case 0
            Case 1
               str编码 = mGet编码By汉字(str编码表, str汉字, 4)
            Case 2
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 2) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 2)
            Case 3
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 3, 1), 2)
            Case Else
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 1) & _
                         mGet编码By汉字(str编码表, Mid(str汉字, 3, 1), 1) & mGet编码By汉字(str编码表, Right(str汉字, 1), 1)
        End Select
    End If
    
    mWBX = str编码
End Function

Private Function mGet编码By汉字(ByVal str编码表 As String, ByVal str汉字 As String, ByVal lngLen As Long) As String
'功能：根据汉字得到其编码
    Dim lngStart As Long, lngEnd As Long
    Dim str编码 As String
    
    lngStart = InStr(str编码表, str汉字)
    If lngStart = 0 Then
        '未在编码表找到该字编码
        mGet编码By汉字 = "Z"
        Exit Function
    End If
    
    lngEnd = InStr(lngStart, str编码表, "|")
    str编码 = Mid(str编码表, lngStart, lngEnd - lngStart)
    mGet编码By汉字 = Mid(Split(str编码, " ")(1), 1, lngLen)
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
    
    On Error GoTo errH
   
    strSQL = "Select zlUppMoney([1]) as NUM From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "UppeMoney", Val(curMoney))
    UppeMoney = rsTmp.Fields(0).value
    Exit Function
errH:
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

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'功能:打开中文输入法，或关闭输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
    
 
    '用户没进行设置，就不处理
'    strIme = gobjComLib.zlDatabase.GetPara("输入法")
    If strIme = "" And blnOpen = True Then Exit Function                 '要求打开输入法，但是又没有设置
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否批定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '不是输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是true的输入法,因此,需要单独处理.
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

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String) As String
'功能：选择Windows的文件夹
    Const BIF_RETURNONLYFSDIRS = 1
    Const BIF_DONTGOBELOWDOMAIN = 2

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    With tBrowseInfo
       .hwndOwner = hwndOwner
       .lpszTitle = lstrcat(strTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       OpenDir = sBuffer
    End If
End Function


Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function GetNeedName(strList As String) As String
'功能：从编码名称组合串中分离出名称
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
    If InStr(strList, Chr(13)) > 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        GetNeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：根据ZLHIS的单据号前缀规则,返回大写的单据号年前缀
'参数：curDate=单据号前缀基准时间,不传时基准时间为服务器时间
'应用：该函数主要是配合GetFullNO函数应用

    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(gobjComLib.zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'功能：由用户输入的部份单号，返回当年的完整单据号。
'参数：strNO=由用户输入的单据号部分

    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
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
End Function

Public Sub SetWindowsInTaskBar(ByVal lngHwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lngHwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lngHwnd, GWL_EXSTYLE, lngStyle)
End Sub

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, _
    Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=传入的日期的标题，如果<>""则显示日期控件，供用户输入日期，将DateInput返回。
'      strDateFormat=时间格式 格式""yyyy-MM-dd hh:mm:ss"其中HH为大写是24小时制"
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    ShowMsgBox = frmMsgBox.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle, strDateCaption, DateInput, strDateFormat)
End Function

'Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, _
'    Optional ByVal strName As String, Optional ByVal strSex As String, _
'    Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
''功能：对密码进行验证
''参数：frmParent=显示的父窗体
''      strPass=正确的密码
''      strName,strSex,strOld=可选参数，病人姓名、性别、年龄，当不传入时不显示这个区域。
''      blnPassEncode-strPass是否传入的加密串
''返回：True=密码验证通过,False=取消输入，或连续3次输入错误的密码
'    VerifyPassWord = frmVerifyPassword.ShowMe(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
'End Function

Public Sub ShowChildWindow(ByVal lngHwnd As Long, Optional ByVal lngMainHwnd As Long, Optional ByVal blnMaximized As Boolean)
'功能：显示指定的窗体，以子窗体方式
'参数：lngHwnd=要作为子窗体显示的窗体的句柄
'      lngMainHwnd=父窗体句柄，不传时表明不以它的子窗体显示
'      blnMaximized=是否固定将窗体最大化处理；因为API设置为子窗体后，窗体原有的最大化特性将失效
'说明：该项函数主要用于在ZLBH中融合调用ZLHIS窗体显示
    Dim vRect1 As RECT, vRect2 As RECT
    Dim X As Long, Y As Long
    
    If lngMainHwnd <> 0 Then
        SetParent lngHwnd, lngMainHwnd
    End If
    If blnMaximized Then
        ShowWindow lngHwnd, SW_SHOWMAXIMIZED
    Else
        If lngMainHwnd <> 0 Then
            '显示在父窗体的中央
            If IsWindowVisible(lngHwnd) = 0 Then
                GetWindowRect lngHwnd, vRect1
                GetWindowRect lngMainHwnd, vRect2
                
                X = ((vRect2.Right - vRect2.Left) - (vRect1.Right - vRect1.Left)) / 2
                Y = ((vRect2.Bottom - vRect2.Top) - (vRect1.Bottom - vRect1.Top)) / 2
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                
                SetWindowPos lngHwnd, HWND_TOP, X, Y, 0, 0, SWP_SHOWWINDOW Or SWP_NOSIZE
            End If
        Else
            '没有主窗体的，显示在主窗体中央
            If IsWindowVisible(lngHwnd) = 0 Then
                GetWindowRect lngHwnd, vRect1
                X = (Screen.Width / Screen.TwipsPerPixelX - (vRect1.Right - vRect1.Left)) / 2
                Y = (Screen.Height / Screen.TwipsPerPixelY - (vRect1.Bottom - vRect1.Top)) / 2
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                                
                SetWindowPos lngHwnd, HWND_TOP, X, Y, 0, 0, SWP_SHOWWINDOW Or SWP_NOSIZE
            End If
            BringWindowToTop lngHwnd
        End If
        
        ShowWindow lngHwnd, SW_RESTORE
    End If
End Sub
Public Function zlStringEncode(ByVal strPutString As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:字符串加密
    '入参:strPutString-需要加密的串
    '出参:
    '返回:加密串
    '编制:刘兴洪
    '日期:2011-05-27 16:32:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPutString = "" Then Exit Function
    zlStringEncode = Md5_String_Calc(strPutString)
End Function
 
Public Function GetPubIcons() As ImageManagerIcons
    Set GetPubIcons = frmPubIcons.imgPublic.Icons
End Function


Public Function DockPannelInit(ByRef dkpMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long

    lngX = X / 15
    lngY = Y / 15

    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next

    lngXY = lngXY * lngY + lngX

    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String, _
                                Optional ByVal strDescriptionText As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        objControl.DescriptionText = strDescriptionText
    End With
    
    Set NewCommandBar = objControl
    
End Function


'Public Function TabControlInit(ByRef tbc As TabControl, _
'                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003, _
'                                Optional ByVal bytPosition As XTPTabPosition = xtpTabPositionTop) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    With tbc
'
'        With .PaintManager
'
'            .Appearance = bytAppearance
'            .ClientFrame = xtpTabFrameSingleLine
'            .Color = xtpTabColorOffice2003
'            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
'            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
'            .ShowIcons = True
'            .BoldSelected = True
'            .Position = bytPosition
'        End With
'
''        Set .Icons = frmPubResource.imgPublic.Icons
'    End With
'
'    TabControlInit = True
'
'End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer, _
                                Optional ByVal strDescriptionText As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText
        objControl.DescriptionText = strDescriptionText
        
        If objControl.type = xtpControlButton Or objControl.type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function CommandBarInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean, Optional ByVal objIcons As ImageManagerIcons) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization
'    If objIcons Is Nothing Then
'        Set cbsMain.Icons = imgPublic.Icons
'    Else
'        Set cbsMain.Icons = objIcons
'    End If
    
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

Public Function CreateCondition() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Fields.Append "条件名称", adVarChar, 30
        .Fields.Append "条件结果", adVarChar, 4000
        .Fields.Append "条件类型", adVarChar, 30
        .Open
    End With
    
    Set CreateCondition = rs
    
End Function

Public Function SetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String, ByVal strConditionValue As String, Optional ByVal strConditionType As String = "文本") As Boolean
    
    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("条件名称").value = strConditionName
    rs("条件结果").value = strConditionValue
    rs("条件类型").value = strConditionType
    
    SetCondition = True
    
End Function

Public Function GetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String) As String
    
    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount > 0 Then
        GetCondition = CStr(rs("条件结果").value)
    End If
    
End Function

Public Function CreateParameter() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Fields.Append "参数名称", adVarChar, 50
        .Fields.Append "参数结果", adVarChar, 4000
        .Fields.Append "参数类型", adVarChar, 30
        .Open
    End With
    
    Set CreateParameter = rs
    
End Function

Public Function SetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String, ByVal strParameterValue As String, Optional ByVal strParameterType As String = "文本") As Boolean
    
    rs.Filter = ""
    rs.Filter = "参数名称='" & strParameterName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("参数名称").value = strParameterName
    rs("参数结果").value = strParameterValue
    rs("参数类型").value = strParameterType
    
    SetParameter = True
    
End Function

Public Function GetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String) As String
    
    rs.Filter = ""
    rs.Filter = "参数名称='" & strParameterName & "'"
    If rs.RecordCount > 0 Then
        GetParameter = CStr(rs("参数结果").value)
    End If
    
End Function

Public Function CheckValid(ByVal strText As String, ByVal emFormat As emDataFormat) As Boolean
    
    Select Case emFormat
    Case emDataFormat.IP4
        
    End Select
    
    CheckValid = True
End Function


Public Function GetGUID() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim udtGUID As GUID

    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnReturnOneData As Boolean, _
                                Optional ByVal frmMain As Object = Nothing, _
                                Optional ByVal strValue As String = "") As Byte
    '******************************************************************************************************************
    '功能：打开树型+列表结构,应用于表格控件
    '参数：
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If blnReturnOneData Then
        If rsData.RecordCount = 1 Then
            Set rsResult = rsData
            ShowPubSelect = 1
            Exit Function
        End If
    End If
    
    If obj Is Nothing Then
        lngX = (Screen.Width - lngCX) / 2
        lngY = (Screen.Height - lngCY) / 2
        lngObjHeight = 0
    Else
        Call ClientToScreen(obj.hWnd, objPoint)
        
        Select Case TypeName(obj)
        Case "TextBox", "CommandButton", "ComboBox"
        
            lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
            lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
            lngObjHeight = obj.Height
            
        Case Else
            lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
            lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
            lngObjHeight = obj.CellHeight
        End Select
    End If
    If frmMain Is Nothing Then
        ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl, strValue)
    Else
        ShowPubSelect = frmMain.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl, strValue)
    End If
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          '全数字
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          '全字母
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
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
