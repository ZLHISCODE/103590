Attribute VB_Name = "mdlCommon"
Option Explicit
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrMatchMethod As String
Public gstrProductName As String
Public gstrDBUser As String   '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object
Public gstrSQL As String
Public gblnTestCardNo As Boolean  '测试
Public gintDebug As Integer
Private Type gPrecision
      ty_小数 As Integer
      ty_Fmt_Vb As String
      ty_Fmt_Ora As String
End Type
Private Type FeePrecision   '费用相关精度
        ty_单价 As gPrecision
        ty_金额 As gPrecision
End Type
Public glngOld As Long
Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect

Private Type SystemParameter
    int简码方式 As Integer
    bln个性化风格 As Boolean               '使用个性化风格
    bln全数字按编码查 As Boolean
    bln全字母按简码查 As Boolean
    bln存在站点 As Boolean      '是否存在站点管理
    ty_费用精度 As FeePrecision    '费用精度
    bln免挂号模式 As Boolean '是否免挂模式,流程：直接在分诊台取号，然后在接诊时，产生划价单
End Type
Public gSystemPara As SystemParameter
Public Enum mAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Ty_UserInfor
    id As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门名称 As String
    
End Type
Public UserInfo As Ty_UserInfor
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public Type Ty_Color
     lngGridColorSel As OLE_COLOR     '选择颜色
     lngGridColorLost As OLE_COLOR   '离开颜色
End Type
Public gSysColor As Ty_Color
Public glngHook As Long
Public gdtBegin As Date

'以下为卡对象
Public gstrComputerName As String '计算机名称
Public glngInstanceCount As Long '当前实例个数
Public gcolPrivs As Collection '权限对象

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


Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    On Error Resume Next
    '获取指定窗体的标题
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function


Public Sub zlSetWindowsBroldStyle(ByVal frmMain As Form)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改变可调窗体不不可调窗体（即设置只有关闭按钮窗口,如果窗体本身只有关闭，只会自动加上最大化、最小化等按钮)
    '入参:frmMain.hwnd-窗体的句柄
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-10 14:58:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim pt_SavePoint As POINTAPI, pt_MovePoint As POINTAPI
    Err = 0: On Error GoTo Errhand:
    With pt_MovePoint
      .X = (-1): .Y = 10
    End With
    '设置窗体的broldStyle
    Call SetWindowLong(frmMain.hWnd, GWL_STYLE, GetWindowLong(frmMain.hWnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call GetSystemMenu(frmMain.hWnd, 1&)
    '必需重画数据
    With frmMain
        .Move .Left, .Top, .Width - 15, .Height - 15
        .Move .Left, .Top, .Width + 15, .Height + 15
    End With
    Call GetCursorPos(pt_SavePoint)
    Call ClientToScreen(frmMain.hWnd, pt_MovePoint)
    Call SetCursorPos(pt_MovePoint.X, pt_MovePoint.Y)
    Call SetCursorPos(pt_SavePoint.X, pt_SavePoint.Y)
Errhand:
End Sub

Public Sub zlInitColorSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始系统颜色
    '编制:刘兴洪
    '日期:2009-11-27 17:12:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '        Public Const G_Row_COLORSEL = &H8000000D
    '        Public Const G_Row_COLORLost = &HE0E0E0
    With gSysColor
        .lngGridColorLost = &HE0E0E0   '离开颜色
        .lngGridColorSel = &HFFEBD7       '选择颜色
    End With
End Sub

Public Function zl_GetUserInfo(Optional cnOracle As ADODB.Connection) As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    Dim objDatabase As clsDataBase
    If Not cnOracle Is Nothing Then
        Set objDatabase = New clsDataBase
        Call objDatabase.InitCommon(cnOracle)
        Set rsTmp = objDatabase.GetUserInfo
        Set objDatabase = Nothing
    Else
        Set rsTmp = zlDatabase.GetUserInfo
    End If
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.id = rsTmp!id
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.部门名称 = "" & rsTmp!部门名
        UserInfo.简码 = "" & rsTmp!简码
        UserInfo.姓名 = "" & rsTmp!姓名
        zl_GetUserInfo = True
    End If
    Exit Function
Errhand:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter() = 1 Then Resume
        Call objDatabase.SaveErrLog
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub
Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    GetTaskbarHeight = os.TaskbarHeight
End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '返回:返回加匹配串%dd%,并且是大写
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function CheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:20070101或2007-01-01
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgbox strTittle & "必须为日期型,请检查！"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgbox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！"
        Exit Function
    End If
    CheckIsDate = strKey
End Function


Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '功能：对文本框的的文本选中或进入进打开输入法
    '参数:blnOpenIme-是否打开输入法
    '返回:
    '--------------------------------------------------------------------------------------------------------
    zlControl.TxtSelAll (objTxt)
    
    If blnOpenIme Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

Public Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    Err = 0
    On Error GoTo Errhand:
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    Exit Function
Errhand:
    TranNumToDate = ""
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
    On Error GoTo Errhand:
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
Errhand:
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
    On Error GoTo Errhand:
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
Errhand:
End Sub

Public Function zl_GetFieldLens(ByVal strTableName As String, ByVal strFields As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取字段的实际长度
    '入参:strTableName-表名称
    '     strFields-字段数(字段名要唯一，否则报错),如:编码,名称,简码
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-17 16:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, cllFields As New Collection
    Dim varFields As Variant, i As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "Select " & strFields & " From " & strTableName & " where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "取字段长度"
    
    varFields = Split(strFields, ",")
    With rsTemp
        For i = 0 To UBound(varFields)
            Select Case .Fields(varFields(i)).type
            Case 222
            Case Else
                cllFields.Add .Fields(varFields(i)).DefinedSize, varFields(i)
            End Select
        Next
    End With
    Set zl_GetFieldLens = cllFields
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub Init站点信息()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化站点的相关信息
    '编制:刘兴洪
    '日期:2009-03-02 17:23:24
    '-----------------------------------------------------------------------------------------------------------
    gbln存在站点控制 = gstrNodeNo <> "-"
 End Sub
Public Sub zl_加载站点信息(ByVal objcbo As ComboBox)
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载站点信息值
    '编制:刘兴洪
    '日期:2009-03-03 12:09:01
    '-----------------------------------------------------------------------------------------------------------
    With objcbo
        .Clear
        .AddItem ""
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .ListIndex = 0
    End With
End Sub
 
Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function


Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可
    '返回:初如成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub


'*********************************************************************************************************************
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlComboxLoadFromRecodeset(ByVal strFromCaption As String, ByVal rsSource As ADODB.Recordset, cboControls As Variant, Optional ByVal blnID As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:本函数的功能是从本地记录何时中装到下拉框中
    '入参:cboControls-控件数组
    '     rsSource:源记录(编码,名称,缺省标志)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-09 14:54:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cboArrays As Variant
    On Error GoTo errHandle
    
    Set rsTemp = rsSource
    '下拉框数组
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        '强行组成一个数组
        cboArrays = Array(cboControls)
    End If
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("编码")) Then
                cboArrays(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
            Else
                cboArrays(intCount).AddItem rsTemp("编码") & "." & rsTemp("名称")
            End If
            If blnID = True Then cboArrays(intCount).ItemData(cboArrays(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("缺省标志") = 1 Then
                cboArrays(intCount).ListIndex = cboArrays(intCount).NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True And cboArrays(intCount).ListIndex < 0 Then cboArrays(intCount).ListIndex = 0
    Next
    zlComboxLoadFromRecodeset = True
    Exit Function
errHandle:
    zlComboxLoadFromRecodeset = False
End Function

Public Function zlComboxLoadFromArray(ByVal varArray As Variant, cboControls As Variant, Optional blnSaveItemData As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:本函数的功能是数组中读出列表值装到下拉框中
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-09 14:53:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cboArrays As Variant
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo errHandle
    
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        '强行组成一个数组
        cboArrays = Array(cboControls)
    End If
    
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        For intArray = LBound(varArray) To UBound(varArray)
            cboArrays(intCount).AddItem varArray(intArray)
        Next
        cboArrays(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromArray = True
    Exit Function
errHandle:
    zlComboxLoadFromArray = False
End Function

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     blnNegative     是否进行负数检查
    '     blnZero         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    Dim dblValue As Double
    If blnZero = True Then
        If strInput = "" Then
            ShowMsgbox str项目 & "未输入，请检查!"
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlDblIsValid = True: Exit Function
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
    zlDblIsValid = True
End Function
Public Function zl_FromComboxGetData(cboControl As ComboBox, Optional ByVal blnID As Boolean = False, Optional strSplit As String = ".") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从Combox中获取数据
    '入参:blnID-是否读取ComboxData数据
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-11 15:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------

    If cboControl.ListIndex < 0 Then zl_FromComboxGetData = "NULL"
    If blnID = False Then
        If cboControl.Text = "" Or cboControl.Enabled = False Then
            zl_FromComboxGetData = "NULL"
        Else
            zl_FromComboxGetData = "'" & Mid(cboControl.Text, InStr(cboControl.Text, strSplit) + 1) & "'"
        End If
    Else
        zl_FromComboxGetData = cboControl.ItemData(cboControl.ListIndex)
    End If
End Function
 Public Function IsDesinMode() As Boolean
      '刘兴洪 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
  

Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
Errhand:
End Function

Public Function zlGetReDawImge(ByVal frmMain As Form, ByVal lngColor As Long, _
    ByVal strCaption As String, sngWidth As Single, sngHeight As Single, _
    Optional sngFontSize As Single = 9, _
    Optional blnFontBold As Boolean = True) As StdPicture
    Dim objPicture As PictureBox
    Set objPicture = frmMain.Controls.Add("VB.PictureBox", "objPictemp")
    With objPicture
        .Cls
        .AutoRedraw = True
        .FontSize = 9
        .Width = sngWidth: .Height = sngHeight
        objPicture.Line (20, 20)-(sngWidth, sngHeight), lngColor, BF              '一个矩形(填充)
        .ForeColor = &H80000016
        .CurrentY = 20
        .FontBold = blnFontBold
        .FontSize = sngFontSize
        If strCaption <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight("刘")) \ 2
            .CurrentX = (.ScaleWidth - .TextWidth(strCaption)) \ 2
            objPicture.Print strCaption
        End If
    End With
    Set zlGetReDawImge = objPicture.Image
    frmMain.Controls.Remove ("objPictemp")
    Set objPicture = Nothing
End Function
Public Sub zlSetStatusPanelCololor(ByVal frmMain As Form, ByVal objStatus As Object, _
    ByVal intPancelIdex As Integer, strCaption As String, _
    ByVal lngColor As Long, Optional blnTextCenter As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置单元格的颜色
    '入参：blnTextCenter-文本居中
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 15:22:18
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    With objStatus
        sngWidth = frmMain.TextWidth(strCaption) + 60
        sngHeight = frmMain.TextHeight("刘") + 60
        .Panels(intPancelIdex).Width = sngWidth
        If blnTextCenter = False Then
            .Panels(intPancelIdex).Width = sngWidth + 300
            .Panels(intPancelIdex).Text = strCaption
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, "", 300, sngHeight, 7, True)
        Else
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, strCaption, sngWidth, sngHeight, 7, True)
        End If
    End With
End Sub
 

Public Sub zlDebugTool(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:跟踪调试信息
    '入参:strInfo-调试信息
    '编制:刘兴洪
    '日期:2011-05-27 11:36:33
    '说明:
    '     gintDebug:1-表示提未调试信息,2-将调式信息写入文本；其它情况不输出调试信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFile As FileSystemObject, objText As TextStream, strFile As String
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "卡结算部件", "调试", 0))
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If gintDebug <= 0 Or gintDebug > 2 Then Exit Sub
    If gintDebug = 2 Then
        '写入文件中
        Set objFile = New FileSystemObject
        strFile = App.Path & "\Square" & Format(Now, "yyyy_MM_DD") & ".Log"
        If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfo: objText.Close
    End If
    MsgBox strInfo, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
End Sub
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Function zlAuditingWarn(ByVal strPrivs As String, _
    ByVal strNos As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核划价单时，对费用进行报警
    '入参:str序号=指定单据中要审核的行号,为空表示所有行
    '返回:
    '编制:刘兴洪
    '日期:2011-06-23 10:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWarn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, j As Long, str类别s As String
    Dim cur当日额 As Currency, cur金额 As Currency, cur余额 As Currency
    Dim strWarn As String, intWarn As Integer
    Dim bln报警包含划价费用  As Boolean
    '记帐报警包含划价费用
    bln报警包含划价费用 = zlDatabase.GetPara(98, glngSys) = "1"
    
    strSQL = "" & _
    " Select /*+ rule */ A.门诊标志, A.姓名, A.病人id , E.预交余额 - E.费用余额 As 余额, B.担保额, C.编码 As 付款码," & vbNewLine & _
    "        A.收费类别, D.名称 As 类别名称, Sum(A.实收金额) As 金额, Zl_Patiwarnscheme(A.病人id) As 适用病人" & vbNewLine & _
    " From 门诊费用记录 A, 病人信息 B, Table(f_Str2list([1])) J," & _
    "           医疗付款方式 C, 收费项目类别 D," & _
    "           (   Select 病人ID,Sum(Nvl(预交余额,0)) as 预交余额,Sum(nvl(费用余额,0))  费用余额" & _
    "               From  病人余额 " & vbNewLine & _
    "               Where   病人ID=[2]  and 性质=1 And nvl(类型,2)=1 Group by 病人ID)  E" & vbNewLine & _
    " Where A.记录性质 = 2 And A.病人ID+0=[2] And A.记录状态 = 0 " & _
    "           And A.NO = J.Column_value " & vbNewLine & _
    "           And A.收费类别 = D.编码 And A.病人id = E.病人id(+) " & vbNewLine & _
    "           And A.病人id = B.病人id And B.医疗付款方式 = C.名称(+)" & vbNewLine & _
    " Group By Nvl(A.价格父号, A.序号), A.门诊标志, A.姓名, A.病人id,  B.担保额, E.预交余额, E.费用余额, C.编码," & vbNewLine & _
    "         A.收费类别, D.名称"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos, lng病人ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If InStr(str类别s, rsTmp!收费类别 & rsTmp!类别名称) = 0 Then
                str类别s = str类别s & "," & rsTmp!收费类别 & rsTmp!类别名称
            End If
            cur金额 = cur金额 + rsTmp!金额
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
        str类别s = Mid(str类别s, 2)
        If cur金额 > 0 Then
            Set rsWarn = zlGetUnitWarn(rsTmp!适用病人, "0")
            cur当日额 = GetPatiDayMoney(rsTmp!病人ID)
            cur余额 = Nvl(rsTmp!余额, 0)
            If bln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, rsTmp!病人ID) + cur金额
            '分类报警
            For j = 0 To UBound(Split(str类别s, ","))
                intWarn = zlBillingWarn(strPrivs, rsTmp!姓名, rsTmp!适用病人, rsWarn, _
                    cur余额, cur当日额, cur金额, Nvl(rsTmp!担保额, 0), _
                    Left(Split(str类别s, ",")(j), 1), Mid(Split(str类别s, ",")(j), 2), strWarn)
                If intWarn = 2 Or intWarn = 3 Then Exit Function
            Next
        End If
    End If
    zlAuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlGetUnitWarn(Optional ByVal str适用病人 As String, Optional ByVal str病区ID As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回病区记帐报警记录集
    '入参:str适用病人-适用的病人
    '        str病区IＤ－病区ID集
    '出参:
    '返回:病区报警集
    '编制:刘兴洪
    '日期:2011-06-24 14:59:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select Nvl(病区ID,0) 病区ID,适用病人,Nvl(报警方法,1) as 报警方法," & _
            " 报警值,报警标志1,报警标志2,报警标志3" & _
            " From 记帐报警线 Where 1=1" & _
            IIf(str适用病人 = "", "", " And 适用病人 = [1]") & _
            IIf(str病区ID = "", "", " And Nvl(病区ID,0) = [2]")
    Set zlGetUnitWarn = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str适用病人, str病区ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlBillingWarn(strPrivs As String, str姓名 As String, str适用病人 As String, _
    rsWarn As ADODB.Recordset, cur余额 As Currency, cur当日额 As Currency, _
    cur单据金额 As Currency, cur担保 As Currency, str类别 As String, _
    ByVal str类别名 As String, ByRef str已报类别 As String, Optional bln多病人 As Boolean, Optional strMoneyFMT As String = "") As Integer
'功能:对病人记帐进行报警提示
'参数:
'     str姓名=病人姓名,用于提示
'     str适用病人=根据病人身份返回的记帐报警适用方案
'     rsWarn=当前病区记帐报警设置记录
'     cur余额=病人余额,用于累计报警
'     cur当日额=病人当日发生的费用额,用于每日报警
'     cur单据金额=病人单据中输入的费用
'     cur担保=病人担保费用额,用于累计报警
'     str类别=当前要检查的类别,用于分类报警
'     str类别名=类别名称,用于提示
'     strMoneyFMT-格式精度
'返回:0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
'     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    Dim i As Integer, byt标志 As Byte
    Dim bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    If strMoneyFMT = "" Then
        strMoneyFMT = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    End If
    '报警参数检查
    rsWarn.Filter = "病区ID=0 And 适用病人='" & str适用病人 & "'"
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    If bln多病人 Then
        '示例：",周:-,张:DEF,李:567,张567"
        '报警标志2示例：",周:-①,张:DEF①,李:567①,张567②"
        bln已报警 = str已报类别 & "," Like "*," & str姓名 & ":-*,*" _
            Or str已报类别 & "," Like "*," & str姓名 & ":*" & str类别 & "*,*"
    Else
        '示例："-" 或 ",ABC,567,DEF"
        '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
        bln已报警 = InStr(str已报类别, str类别) > 0 Or str已报类别 Like "-*"
    End If
    
    If bln已报警 Then
        If byt标志 = 2 Then
            If bln多病人 Then
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str姓名 & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str姓名 & ":*" & str类别 & "*,*" Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For  '说明见住院模块
                    End If
                Next
            Else
                If str已报类别 Like "-*" Then
                    byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
                Else
                    arrTmp = Split(str已报类别, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str类别) > 0 Then
                            byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                            'Exit For '说明见住院模块
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名 <> "" Then str类别名 = """" & str类别名 & """费用"
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        '先只有两种:1.强制记帐,无权限时,禁止记帐
                        Call MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐", vbInformation + vbOKOnly, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur余额 + cur担保 - cur单据金额 < 0 Then
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                            zlBillingWarn = 3
                        Else
                            MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    ElseIf cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                zlBillingWarn = 2
                            Else
                                zlBillingWarn = 1
                            End If
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur余额 + cur担保 - cur单据金额 < 0 Then
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                                MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                zlBillingWarn = 3
                            Else
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                                zlBillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        Call MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐.", vbOKOnly + vbInformation, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If zlBillingWarn = 1 Or zlBillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            GetUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSQL = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                GetUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                GetUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            GetUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人当天发生的费用总额
    '返回:获取病人当天发生的费用总额
    '编制:刘兴洪
    '日期:2011-06-23 10:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng病人ID As Long) As Double
'功能:获取指定病人的划价单金额合计
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '记帐报警包含所有住院划价费用
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara(42, glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(主页ID,0) = (Select Nvl(主页ID,0) From 病人信息 Where 病人ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志=2" & strWhere
    Else
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计 " & _
        "   From 门诊费用记录  " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]  and 门诊标志<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(划价费用合计,0)) as 划价费用合计  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取指定病人的划价总额", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
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
Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function
Public Function StrToNum(ByVal strNumber As String) As Double
    '功能:将字符串转换成数据
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function


Public Function ExistFeeInsurePatient(lng病人ID As Long) As Boolean
'功能：判断医保病人是否存在未结费用
'返回：
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSQL = "Select Nvl(sum(B.费用余额),0) 费用余额 From 病人信息 A,病人余额 B Where A.病人ID=B.病人ID And Nvl(A.险类,0)<>0 And A.病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!费用余额 <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select 编码 as ID,编码,名称,简码 From 区域" & _
                 " Where Nvl(级数,0)<3 And (编码 Like [1] Or upper(简码) Like '" & gstrLike & "'||[1]||'%' Or 名称 Like '" & gstrLike & "'||[1]||'%')"
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select 编码 as ID,编码,名称,简码 From 区域 Where Nvl(级数,0)<3"
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str病人类型 As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人类型,设置不同病人类型的显示颜色
    '入参:objPatiControl-病人控件(文本框,标签)
    '    str病人类型-病人类型
    '    lngDefaultColor-缺省病人的显示颜色
    '返回:True-设置颜色成功，False-失败
    '编制:李南春
    '日期:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str病人类型 <> "" Then
        lngColor = zlDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'功能：四舍五入方式格式化数字
'参数：intBit=最大小数位数
'问题号：94552
'说明：VB自带的Round是银行家舍入法,与实际不一致。如Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Function CentMoney(ByVal curMoney As Currency, ByVal bytMoney As Byte) As Currency
'功能：对指定金额按分币处理规则进行处理,返回处理后的金额
'参数：curMoney=要进行分币处理的金额(为应缴金额,2位小数)
'      bytMoney=
'         0.不处理
'         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
'         2.补整收法,eg:0.51=0.60,0.56=0.60
'         3.舍分收法,eg:0.51=0.50,0.56=0.50
'         4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约
'           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
'         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.29(含)以下都舍掉角，0.80(含)以上都进角，0.3-0.79处理为0.5。
'         6.五舍六入:eg:0.15=0.10:0.16=0.2:   刘兴洪 问题:34519  日期:2010-12-06 09:58:02
'91385,调整“5.三七作五、二舍八入”规则：先对分币进行四舍五入，即0.24(含)以下都舍掉角，0.75(含)以上都进角，0.25-0.74都处理为0.5
'       分币先四舍五入，那么0.00～0.24=0，0.25～0.5=0.50, 0.50～0.74=0.50，0.75～1.00=1，这样舍和入各占50%的比例
    
    Dim intSign As Integer, curTmp As Currency

    If bytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf bytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '先取两位金额,再处理分币,如:0.248 得0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf bytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf bytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf bytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf bytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf bytMoney = 6 Then
        '刘兴洪 问题:34519 五舍六入:eg:0.15=0.10:0.16=0.2:    日期:2010-12-06 09:58:02
         CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

'============================================================================================================
'单元格自动根据数据调整行列数，使其不出现横向滚动条
Public Sub ZL_vsGrid_AutoSetGridRowAndCol(vsfGrid As VSFlexGrid)
    '自动调整表格的行列，使其不出现横向滚动条
    Dim sngWidth As Single
    
    With vsfGrid
        If .Cols <= 0 Or .Rows <= 0 Then Exit Sub
        .AutoSize 0, .Cols - 1
        sngWidth = GetAllCellWidth(vsfGrid)
        If sngWidth <= .Width - 300 Or .Cols = 1 Then Exit Sub
        
        Call MoveGridCell(vsfGrid)
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfGrid)
    End With
End Sub

Private Sub MoveGridCell(vsfGrid As VSFlexGrid)
    '移动单元格
    Dim lngNewRows As Long, lngNewCols As Long
    Dim i As Long, j As Long
    Dim lngMoveCells As Long, lngMoveRows As Long
    Dim lngAfterCurRowCells As Long, lngCurRowAfterCells As Long
    Dim lngNewCol As Long
    Dim blnExitFor As Boolean
    
    With vsfGrid
        '1.新的列数
        lngNewCols = .Cols - 1
        
        '2.新的行数
        lngAfterCurRowCells = 0 '确定最后一行剩余的空白单元格
        For i = .Cols - 1 To 0 Step -1
            If Trim(.TextMatrix(.Rows - 1, i)) <> "" Then Exit For
            lngAfterCurRowCells = lngAfterCurRowCells + 1
        Next
        If .Rows - lngAfterCurRowCells > 0 Then
            lngNewRows = .Rows + (Ceil((.Rows - lngAfterCurRowCells) / lngNewCols))
        Else
            lngNewRows = .Rows
        End If
        
        '3.开始移动单元格
        .Rows = lngNewRows
        blnExitFor = False
        For i = .Rows - 1 To 0 Step -1
            For j = .Cols - 1 To 0 Step -1
                If Trim(.TextMatrix(i, j)) <> "" Then
                    '确定目标单元格的位置
                    If j = .Cols - 1 Then
                        lngMoveCells = i + 1 '移动单元格数
                        lngAfterCurRowCells = 0 '当前行后面的单元格数
                    Else
                        lngMoveCells = i  '移动单元格数
                        lngAfterCurRowCells = lngNewCols - 1 - j '当前行后面的单元格数
                    End If
                    
                    lngCurRowAfterCells = lngMoveCells - lngAfterCurRowCells
                    lngMoveRows = Ceil(lngCurRowAfterCells / lngNewCols) '移动的行数
                    If lngMoveRows = 0 Then
                        lngNewCol = j + lngMoveCells
                    Else
                        If lngCurRowAfterCells Mod lngNewCols = 0 Then
                            lngNewCol = lngNewCols - 1
                        Else
                            lngNewCol = lngMoveCells - (Floor(lngCurRowAfterCells / lngNewCols)) * lngNewCols - lngAfterCurRowCells - 1
                        End If
                    End If
                    
                    '移动数据
                    .TextMatrix(i + lngMoveRows, lngNewCol) = .TextMatrix(i, j)
                    .Cell(flexcpData, i + lngMoveRows, lngNewCol) = .Cell(flexcpData, i, j)
                    .Cell(flexcpChecked, i + lngMoveRows, lngNewCol) = .Cell(flexcpChecked, i, j)
                End If
                If i = 0 And j = .Cols - 1 Then blnExitFor = True: Exit For
            Next
            If blnExitFor Then Exit For
        Next
        .Cols = lngNewCols
    End With
End Sub

Private Function GetAllCellWidth(vsfGrid As VSFlexGrid) As Single
    '获取所有单元格总宽度
    Dim i As Long
    Dim sngWith As Single
    
    sngWith = 0
    With vsfGrid
        For i = 0 To .Cols - 1
            sngWith = sngWith + .ColWidth(i) + 10
        Next
    End With
    GetAllCellWidth = sngWith
End Function

Public Function ZL_vsGrid_CurrCellHaveData(ByVal vsGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '检查单元格是否有数据
    On Error GoTo ErrHandler
    With vsGrid
        If lngRow = -1 Then lngRow = .Row
        If lngCol = -1 Then lngCol = .Col
        If lngRow < 0 Or lngCol < 0 Then Exit Function
        If lngRow > .Rows - 1 Or lngCol > .Cols - 1 Then Exit Function
        If Trim(.TextMatrix(lngRow, lngCol)) = "" Then Exit Function
    End With
    ZL_vsGrid_CurrCellHaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZL_vsGrid_RemoveCell(ByVal vsGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '从表格中移除当前选择卡号/卡范围
    '入参：
    '   lngRow - 当前行
    '   lngCol - 当前列
    Dim i As Long, j As Long
    
    On Error GoTo ErrHandler
    With vsGrid
        .Redraw = flexRDNone
        If lngRow = -1 Then lngRow = .Row
        If lngCol = -1 Then lngCol = .Col
        
        For i = lngRow To .Rows - 1
            For j = 0 To .Cols - 1
                If (i = lngRow And j >= lngCol) Or i > lngRow Then
                    If (i < .Rows - 1 And j = .Cols - 1) Then
                        '非最后一行最后一列
                        If Trim(.TextMatrix(i + 1, 0)) = "" Then
                            .Redraw = flexRDBuffered
                            ZL_vsGrid_RemoveCell = True
                            Exit Function
                        Else
                            .TextMatrix(i, j) = .TextMatrix(i + 1, 0)
                            .Cell(flexcpData, i, j) = .Cell(flexcpData, i + 1, 0)
                            .Cell(flexcpChecked, i, j) = .Cell(flexcpChecked, i + 1, 0)
                        End If
                    ElseIf i = .Rows - 1 And j = .Cols - 1 Then
                        '最后一行最后一列
                        .TextMatrix(i, j) = ""
                        .Cell(flexcpData, i, j) = ""
                        .Cell(flexcpChecked, i, j) = 0
                        
                        If i = 0 Then
                            .Cols = .Cols - 1
                        Else
                            .Col = j - 1
                        End If
                        .Redraw = flexRDBuffered
                        ZL_vsGrid_RemoveCell = True
                        Exit Function
                    Else
                        If .TextMatrix(i, j + 1) = "" Then
                            .TextMatrix(i, j) = ""
                            .Cell(flexcpData, i, j) = ""
                            .Cell(flexcpChecked, i, j) = 0
                            
                            If j = 0 Then
                                .Rows = .Rows - 1
                                .Row = .Rows - 1: .Col = .Cols - 1
                            Else
                                If .Col = j Then .Col = j - 1
                            End If
                            .Redraw = flexRDBuffered
                            ZL_vsGrid_RemoveCell = True
                            Exit Function
                        Else
                            .TextMatrix(i, j) = .TextMatrix(i, j + 1)
                            .Cell(flexcpData, i, j) = .Cell(flexcpData, i, j + 1)
                            .Cell(flexcpChecked, i, j) = .Cell(flexcpChecked, i, j + 1)
                        End If
                    End If
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZL_vsGrid_AddCell(ByVal vsGrid As VSFlexGrid, _
    ByVal strText As String, ByVal varData As Variant, Optional ByVal blnCheck As Boolean) As Boolean
    '向表格中添加数据
    '入参：
    '   strText - 表格显示文本
    '   varData - 单元格存储数据
    Dim i As Long, j As Long
    Dim blnNoData As Boolean
    
    On Error GoTo ErrHandler
    With vsGrid
        .Redraw = flexRDNone
        If .Rows = 0 Then .Rows = 1: blnNoData = True
        If .Cols = 0 Then .Cols = 1: blnNoData = True
        If blnNoData Then
            .TextMatrix(0, 0) = strText
            .Cell(flexcpData, 0, 0) = varData
            If blnCheck Then .Cell(flexcpChecked, 0, 0) = 2
            .Row = 0: .Col = 0
            .Redraw = flexRDBuffered
            ZL_vsGrid_AddCell = True
            Exit Function
        End If
        
        For i = .Rows - 1 To 0 Step -1
            For j = .Cols - 1 To 0 Step -1
                If Trim(.TextMatrix(i, j)) <> "" Then
                    If j = .Cols - 1 Then
                        If i = 0 Then
                            .Cols = .Cols + 1
                            .TextMatrix(0, .Cols - 1) = strText
                            .Cell(flexcpData, 0, .Cols - 1) = varData
                            If blnCheck Then .Cell(flexcpChecked, 0, .Cols - 1) = 2
                        Else
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = strText
                            .Cell(flexcpData, .Rows - 1, 0) = varData
                            If blnCheck Then .Cell(flexcpChecked, .Rows - 1, 0) = 2
                        End If
                    Else
                        .TextMatrix(i, j + 1) = strText
                        .Cell(flexcpData, i, j + 1) = varData
                        If blnCheck Then .Cell(flexcpChecked, i, j + 1) = 2
                    End If
                    .Redraw = flexRDBuffered
                    ZL_vsGrid_AddCell = True
                    Exit Function
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


