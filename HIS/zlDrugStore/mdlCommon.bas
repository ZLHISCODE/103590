Attribute VB_Name = "mdlCommon"
Option Explicit
Private mobjVoice As Object

Private Type MousePoint
    CurX As Single
    CurY As Single
End Type
Public CurMousePoint As MousePoint          '鼠标位置

Public Const CALLSOUND_MS = 1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const CB_FINDSTRING = &H14C
Public Const GWL_HWNDPARENT = (-8)
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Const ETO_OPAQUE = 2

Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

'语音播放的函数
Public Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Public Declare Function StopPlayStr Lib "StrSound" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub zlCall_MsSoundPlay(ByVal strCall As String, ByVal intVoiceSpeed As Integer)
    Dim Token As Object
    
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
'        Set mobjVoice.Voice = mobjVoice.GetVoices("Name=Microsoft Lili").Item(0)
    End If
    
'    For Each Token In objVoice.GetVoices
'        Debug.Print Token.GetDescription()
'    Next
    
'    Microsoft Lili - Chinese(China)
'    Microsoft Anna - English (United States)
'    Microsoft Simplified Chinese
    
    '语音类型
'    Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Simplified Chinese").Item(0)
'    Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Sam").Item(0)
'    Set objVoice.Voice = objVoice.GetVoices("Name=Girl XiaoKun").Item(0)
    
    
    If intVoiceSpeed > 10 Or intVoiceSpeed < -10 Then
        intVoiceSpeed = -4
    End If
    
    mobjVoice.Rate = intVoiceSpeed   '速度:-10,10  0
    mobjVoice.Volume = 100           '声音:0,100 100
    
'    objVoice.Speak "请、" & "马政、" & "马政、" & "、到一号窗口"
'    objVoice.Speak "请、" & "余智勇、" & "余智勇、" & "、到一号窗口"
'    objVoice.Speak "请、" & "马 政、" & "马 政、" & "、到药房窗口"
    
    mobjVoice.speak strCall, 1
'    Set objVoice = Nothing
End Sub

Public Sub zlCall_SystemSoundPlay(ByVal strCall As String, ByVal intVoiceSpeed As Integer)
'    Call StartTextPlay("请、" & "余智勇、" & "余智勇、" & "、到一号窗口", 60)
    
    If intVoiceSpeed > 100 Or intVoiceSpeed < 0 Then
        intVoiceSpeed = 65
    End If
    
    Call StartTextPlay(strCall, intVoiceSpeed)
End Sub

Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '根据传入的字符串进行分解，大于指定字符长度就需要进行分解，结果保存到数组中
    '入参：strInput-输入的字符串；strSplitChar-字符串中内容的分隔符
    '返回：数组，其中数组成员的字符长度不超过指定长度
    Dim strArray As Variant
    Dim ArrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '大于指定字符时就需要分解
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '无分隔符时
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '有分隔符时
            ArrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(ArrTmp)
        
            For i = 0 To lngCount
                If ArrTmp(i) <> "" Then
                    '有分隔符的需要保持分隔符之间字符的完整性，不能把分隔符之间的字符拆开
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = ArrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    GetArrayByStr = strArray
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

Public Function mWBX(ByVal strAsk As String, ByVal lng方式 As Long) As String
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
    If err <> 0 Then
        blnNotFound = True
'        MsgBox "未发现" & strFile & "文件。", vbInformation, gstrSysName
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

Public Function mPinYin(ByVal strAsk As String) As String
'功能：返回指定字符串的拼音简码
'参数：strAsk  待处理的字符串

    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer
    Dim StrCode As String, StrBit As String

'    aryStard = Split("八;擦;哒;讹;发;噶;哈;击;;咔;垃;妈;拿;噢;啪;七;热;撒;他;挖;;挖;西;丫;匝", ";")
    aryStard = Split("八;擦;咑;屙;发;旮;铪;讥;;咔;垃;妈;拿;噢;啪;七;呥;撒;他;挖;;;西;丫;匝", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '将全角转换为半角，小写转换为大写
    
    StrCode = ""
    For intBit = 1 To Len(strAsk)
        StrBit = Mid(strAsk, intBit, 1)
        If InStr(1, "ⅠⅡⅢⅣⅤⅥⅧⅧⅨⅩαβγ蔫趴属哇娃夕汐仨兮拚嚓饧澶赕膪欹焘恁砉铛疋覃瞿她", StrBit) > 0 Then
            '特殊字的处理
            StrCode = StrCode & Switch(StrBit = "Ⅰ", "1", StrBit = "Ⅱ", "2", StrBit = "Ⅲ", "3", StrBit = "Ⅳ", "4", StrBit = "Ⅴ", "5" _
                            , StrBit = "Ⅵ", "6", StrBit = "Ⅷ", "7", StrBit = "Ⅷ", "8", StrBit = "Ⅸ", "9" _
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
                        StrCode = StrCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        StrCode = StrCode & "Z"
                    End If
                End If
            Next
        Else
            If StrBit >= "A" And StrBit <= "Z" Then
                StrCode = StrCode & StrBit
            End If
        End If
        If Len(StrCode) >= 10 Then Exit For
    Next
    mPinYin = StrCode

End Function


Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function
Public Function AviShow(FrmMain As Form, Optional ByVal blnShow As Boolean = True)
    '控制Flash窗体
    DoEvents
    
    If blnShow Then
        FS.ShowFlash "正在查找数据,请稍候...", FrmMain
    Else
        FS.StopFlash
    End If
    
    DoEvents
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDO As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDO = 1 To 12
        strSource = Mid(strOld, intDO, 1)
        strTarget = Mid(strPass, intDO, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function 相同符号(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_负数 As Boolean, blnSecond_负数 As Boolean
    相同符号 = False
    
    blnFirst_负数 = (sinFirst <= 0)
    blnSecond_负数 = (sinSecond <= 0)
    
    相同符号 = (blnFirst_负数 = blnSecond_负数)
End Function

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then '为1表示中文输入法
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(1, strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function InDesign() As Boolean
    'InDesign = False: Exit Function
    
    On Error Resume Next
    Debug.Print 1 / 0
    If err.Number <> 0 Then err.Clear: InDesign = True
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = 0 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
    If MatchIndex = -1 Then MatchIndex = 1
End Function


Public Function SelAll(txtObj As TextBox)
    With txtObj
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Function

Public Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '获取指定窗体的父窗体
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function

Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    '取指定列头的列位置
    
    Dim i As Integer
    
    On Error GoTo errH
    
    GetCol = -1
    
    If TypeName(mshFlex) = "MSHFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    ElseIf TypeName(mshFlex) = "VSFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    End If
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    On Error Resume Next
    '获取指定窗体的标题
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function

Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng序号列 As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    '从指定行开始更新序号
    
    With mshBill
        lngRows = .rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng序号列) = lngRow
        Next
    End With
End Sub
