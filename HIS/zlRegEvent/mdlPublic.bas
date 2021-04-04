Attribute VB_Name = "mdlPublic"
Option Explicit '要求变量声明
Public gclsInsure As New clsInsure          '医保接口对象
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrPrivsStation As String '当前用户在医生工作站的权限  只有通过接口调入时,才存在
Public gstrSysName As String                '系统名称
Public gstrUnitName As String
Public glngSys As Long
Public glngModul As Long
Public gstrProductName As String

Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gbytDec As Byte '费用金额的小数点位数
Public gbyt清除门诊信息 As Byte '0-不清除;1-清除;2-提示清除
Public gblnOk As Boolean
Public gstrDBUser As String '当前用户名
Public gfrmMain As Object
'用户信息------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

'系统参数
Public Type TY_Reg_Para  '挂号相关参数
    bytNODaysGeneral As Byte    '普通挂号有效天数
    bytNoDayseMergency As Byte '急诊挂号有效天数
End Type
Public Type TY_SysPara
    Sy_Reg  As TY_Reg_Para
End Type
Public gSysPara As TY_SysPara       '系统参数相关;以后可以扩展(刘兴洪)

Public gstrLike As String   '输入匹配方式
Public glngInterval As Long '挂号安排表自动刷新间隔,0表示不自动刷新

'Public gblnShowCard As Boolean '是否明文显示卡号

Public gblnSharedInvoice As Boolean '挂号使用收费票据
Public gblnBill挂号 As Boolean '是否严格控制票据

Public gbytFactLength As Byte '挂号票据号码长度
Public glng挂号ID As Long '挂号领用ID
Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbyt预存款退费验卡 As Byte '预存款退费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbln消费卡退费验卡 As Boolean '消费卡退费时是否刷卡验证
Public gbln精简界面 As Boolean

Public gstr磁卡ID As String  '就诊卡领用ID
'Public gblnBill磁卡 As Boolean '是否严格控制票据
'Public gbyt磁卡 As Byte '就诊卡号长度
Public gstrCardPass As String '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
Public gblnPrePayPriority As Boolean '优先使用预交款

Public gint预约天数 As Integer '挂号允许的预约天数
Public gstr上班时间 As String

Public gstr挂号科室ID As String   '本工作站允许挂号的科室ID
Public gstrIme As String '自动开启的输入法

Public gbytRegistMode As Byte '挂号模式
Public gdatRegistTime As Date '出诊表模式启用时间

Public Type TY_VisitPlan_ModulePara '临床出诊安排模块参数
    byt出诊表打印方式 As Byte
    str号源维护站点 As String '未区分站点的科室号源的维护站点
    byt号码比较方式  As Byte '号源号码按哪种比较方式进行排序：0-按字符比较，1-按数值比较
End Type
Public gVisitPlan_ModulePara As TY_VisitPlan_ModulePara

'可选输入项目
Public gbln病人 As Boolean '病人
Public gbln性别 As Boolean  '性别
Public gbln年龄 As Boolean  '年龄
Public gbln家庭地址 As Boolean  '家庭地址
Public gbln付款方式 As Boolean  '付款方式
Public gbln费别 As Boolean '费别
Public gbln结算方式 As Boolean '结算方式
Public gbln医生 As Boolean '医生
Public gbln电话 As Boolean

'缺省值
Public gstr付款方式 As String '缺省付款方式
Public gstr费别 As String '缺省费别
Public gstr性别 As String '缺省性别
Public gstr结算方式 As String '缺省结算方式
'刘兴洪 问题:????    日期:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000

'其它参数
Public gbln缴款结束 As Boolean
Public gbln自动门诊号 As Boolean
Public gblnAutoAddName As Boolean '发卡时自动产生临时姓名
Public gblnNewCardNoPop As Boolean '发卡时不弹出窗口若悬河
Public gbln卡费仅划价 As Boolean
Public gbln退费重打 As Boolean '退号不退卡时是否重打发票
Public gint号长 As Integer '号别长度
Public gblnLED As Boolean
Public gblnPrintFree As Boolean
Public gblnPrintCase As Boolean '打印病历标签
Public gbytInvoice As Byte   '发票打印方式
Public gByt打印病人条码 As Byte '病人条码 打印方式
Public gblnPrice As Boolean     '建档病人挂号存为划价单
Public gintNameDays As Integer  '输入姓名查找N天内的病人
Public gblnSeekName As Boolean

Public glngOld As Long
Public glngMinW As Long, glngMaxW As Long
Public glngMinH As Long, glngMaxH As Long
Public gbln身份证唯一 As Boolean
'WIN32函数

'API申明
Public Const CB_ADDSTRING = &H143
Public Const CB_FINDSTRING = &H14C
Public Const CB_SHOWDROPDOWN = &H14F

Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
'控件坐标位置获取转换
Public Const EM_EXGETSEL = (&H400 + 52)
Public Const EM_POSFROMCHAR = &HD6

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000



Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4

Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim p As KBDLLHOOKSTRUCT
Public p1 As KBDLLHOOKSTRUCT
Public gblnBegin As Boolean
Public gblnLen As Boolean
Public gblnCard As Boolean
Public gsngStartTime As Single

Public glngTXTProc As Long '保存默认的消息函数的地址
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '厚边框
Public Const WS_SYSMENU = &H80000           '在标题栏是否具备系统菜单
Public Const WS_MINIMIZEBOX = &H20000       '具备最小化按钮
Public Const WS_MAXIMIZEBOX = &H10000       '具备最大化按钮
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_RESETCONTENT = &H14B

'移动控件或无边框窗体
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF010&
Public Const HTCAPTION = 2

'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Type Ty_CardProperty
       lng卡类别ID      As Long
       str卡名称        As String
       str短名称        As String
       lng卡号长度      As Long
       lng结算方式      As String
       bln自制卡        As Boolean
       bln严格控制      As Boolean
       lng领用ID        As Long
       lng共用批次      As Long
       bln变价          As Boolean
       int密码长度      As Integer
       int密码长度限制  As Integer
       int密码规则      As Integer
       bln就诊卡        As Boolean
       str卡号密文      As String
       str特准项目      As String
       bln缺省标志      As Boolean
       blnOneCard       As Boolean '  '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费
       rs卡费           As ADODB.Recordset
       dbl应收金额      As Double
       dbl实收金额      As Double
       bln是否制卡      As Boolean
       bln是否发卡      As Boolean
       bln是否写卡      As Boolean
       lng发卡性质      As Long '0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0 问题号:57326
       bln重复使用      As Boolean
       str读卡性质      As String
       byt发卡控制      As Byte
       lng收费细目ID    As Long '医院自身调整卡费返回的收费细目ID,不与当前卡费的收费细目ID同步
End Type
Public gCurSendCard As Ty_CardProperty
Public gstrSQL  As String
Public glngMax家庭地址 As Long       '家庭地址最大允许录入长度
Public glngMax户口地址 As Long       '户口地址最大允许录入长度
Public glngMax出生地点 As Long       '出生地点最大允许录入长度
Public glngMax联系人地址 As Long    '联系人地址最大允许录入长度

Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'功能：去掉TextBox的默认右键菜单
    If msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, Hwnd, msg, wp, lp)
    End If
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
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Function FindName(cbo As ComboBox) As String
'功能：取出当前ComboBox的值(其组成为“编号-名称”)
'说明：主要为SQL语句使用
    If cbo.ListIndex = -1 Then
        FindName = "Null"
    Else
        FindName = "'" & Mid(cbo.Text, InStr(1, cbo.Text, "-") + 1) & "'"
    End If
End Function

Public Function FindText(txt As TextBox) As String
'功能：将当前TextBox的值转化为标准SQL语句
'说明：主要为SQL语句使用
    If Len(Trim(txt.Text)) = 0 Then
        FindText = "Null"
    Else
        FindText = "'" & txt.Text & "'"
    End If
End Function

Public Function NeedName(strList As String, Optional ByVal blnLast As Boolean = False, _
Optional strSplit As String = "-") As String
    If Not blnLast Then
        NeedName = Mid(strList, InStr(strList, strSplit) + 1)
    Else
        NeedName = strList
        Do While (InStr(NeedName, strSplit)) > 0
            NeedName = Mid(NeedName, InStr(NeedName, strSplit) + 1)
        Loop
    End If
End Function

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim fEatKeystroke As Boolean
    Dim sngTime As Single
    Dim sngPreTime As Timer
    
    gblnCard = False
    
    sngTime = Timer
    If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            
            CopyMemory p, ByVal lParam, Len(p)
            gblnCard = (sngTime - gsngStartTime) < 0.6
            If gblnCard = False Then gblnLen = False
             
            gsngStartTime = sngTime
            fEatKeystroke = _
            ((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
            ((p.vkCode = 91) Or (p.vkCode = 92) Or (p.vkCode = 93)) Or _
            ((p.vkCode = VK_F4) And (p.flags And LLKHF_ALTDOWN) <> 0) '加入这行代码屏弊Alt+F4
            If p.vkCode = Asc(";") Then fEatKeystroke = True
        End If
        
        If p.vkCode = vbKeyBack Then
            LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
            Exit Function
        End If
    End If
    If (fEatKeystroke Or gblnLen) Then
        LowLevelKeyboardProc = -1
    Else
        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
End Function

Public Function NeedCode(strList As String) As String
    If InStr(strList, "-") = 0 Then NeedCode = strList: Exit Function
    NeedCode = Mid(strList, 1, InStr(strList, "-") - 1)
End Function
Public Function Custom_WndMessage(ByVal Hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW \ 15
        MinMax.ptMinTrackSize.Y = glngMinH \ 15
        MinMax.ptMaxTrackSize.X = glngMaxW \ 15
        MinMax.ptMaxTrackSize.Y = glngMaxH \ 15
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, Hwnd, msg, wp, lp)
End Function

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
    If cbo.ListCount > 0 And cbo.ListIndex = -1 Then cbo.ListIndex = 0
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function OpenIme(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "不自动开启" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
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

Public Function GetBaseDict() As ADODB.Recordset
'功能：从字典中读取数据
    Dim strSQL As String, strTmp As String, arrTmp As Variant, i As Integer
    strTmp = "国籍,民族,婚姻状况,职业,社会关系"
    arrTmp = Split(strTmp, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If strSQL = "" Then
            strSQL = "Select '" & strTmp & "' 类别,编码,名称,Nvl(缺省标志,0) as 缺省 From " & strTmp
        Else
            strSQL = strSQL & " Union all Select '" & strTmp & "' 类别,编码,名称,Nvl(缺省标志,0) as 缺省 From " & strTmp
        End If
    Next
    strSQL = strSQL & " Order by 类别,编码"
    
    On Error GoTo errH
    Set GetBaseDict = zlDatabase.OpenSQLRecord(strSQL, "获取国籍,民族,婚姻状况,职业,社会关系")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "病人ID", adBigInt, , adFldIsNullable
            .Append "主页ID", adBigInt, , adFldIsNullable
            .Append "挂号ID", adBigInt, , adFldIsNullable
            .Append "门诊号", adVarChar, 18, adFldIsNullable
            .Append "住院号", adVarChar, 18, adFldIsNullable
            .Append "医保号", adVarChar, 30, adFldIsNullable
            .Append "身份证号", adVarChar, 18, adFldIsNullable
            .Append "其他证件", adVarChar, 20, adFldIsNullable
            .Append "姓名", adVarChar, 100, adFldIsNullable
            .Append "性别", adVarChar, 4, adFldIsNullable
            .Append "出生日期", adVarChar, 20, adFldIsNullable
            .Append "出生地点", adVarChar, 100, adFldIsNullable
            .Append "国籍", adVarChar, 30, adFldIsNullable
            .Append "民族", adVarChar, 20, adFldIsNullable
            .Append "学历", adVarChar, 10, adFldIsNullable
            .Append "职业", adVarChar, 80, adFldIsNullable
            .Append "工作单位", adVarChar, 100, adFldIsNullable
            .Append "邮箱", adVarChar, 30, adFldIsNullable
            .Append "婚姻状况", adVarChar, 4, adFldIsNullable
            .Append "家庭电话", adVarChar, 20, adFldIsNullable
            .Append "联系人电话", adVarChar, 20, adFldIsNullable
            .Append "单位电话", adVarChar, 20, adFldIsNullable
            .Append "家庭地址", adVarChar, 100, adFldIsNullable
            .Append "家庭地址邮编", adVarChar, 6, adFldIsNullable
            .Append "户口地址", adVarChar, 100, adFldIsNullable
            .Append "户口地址邮编", adVarChar, 6, adFldIsNullable
            .Append "单位邮编", adVarChar, 6, adFldIsNullable
            .Append "联系人地址", adVarChar, 100, adFldIsNullable
            .Append "联系人关系", adVarChar, 30, adFldIsNullable
            .Append "联系人姓名", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub
