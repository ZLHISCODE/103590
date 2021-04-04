Attribute VB_Name = "mdlCashBill"
Option Explicit

Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrProductName As String            '产品名称
Public gstrMatchMethod As String
Public gstrDBUser As String                 '当前数据库用户
Private mrsPayMode As ADODB.Recordset
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

'票据控制
Public gobjBillPrint As Object '第三方票据打印部件
Public gblnBillPrint As Boolean '第三方票据打印部件是否可用

Public gstrSQL As String
Public gstr单位名称 As String
Public glngSys  As Long
Public glngModul As Long

Public Enum gBillType '票据种类
    收费收据 = 1
    预交收据 = 2
    结帐收据 = 3
    挂号收据 = 4
    就诊卡 = 5
    消费卡 = 6
    会员卡 = 5
End Enum

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum EM_DrawStyle
    DW_Flat = 0  '= 平面
    Dw_SubKen = -1 '= 凹下
    Dw_Heave = 1  '= 凸起
    Dw_Deepen_Subken = -2 '= 深凹下,
    Dw_Deepen_Heave = 2 ' = 深凸起
End Enum

'控件定位
Public Type ty_ctlObject_Locale
    '控件的位置
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '下拉列表的最小高度和宽度
    minWidth As Single
    minHeight As Single
    
    '下接列表的实际位置
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '屏模相关
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type

Public Enum Em_Appearance
    Show_3D = 1     '3D显示
    Show_Flat = 0   '平面
End Enum
Public Enum Em_BorderStyle
    Show_Fixed_Single = 1
    Show_None = 0   '无边框线
End Enum

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
    g本机公共模块 = 5
    g本机私有模块 = 6
End Enum
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const SPI_GETWORKAREA = 48
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
'系统方案设置----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private mlng部门编码平均长度 As Long
Public gstrLike  As String
Public gstrDec As String

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MshGetColNum(msh As MSHFlexGrid, strColName As String) As Long
'功能:根据列名查找MSHFlexGrid控件中的列序号,没有找到时返回-1
'参数:strColName-列名
    Dim i As Long
    
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strColName Then MshGetColNum = i: Exit Function
    Next
    MshGetColNum = -1
End Function


Public Sub zlRaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '功能：将PictureBox模拟成3D平面按钮
    'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
    Dim PicRect As RECT
    Dim lngTmp As Long
    If picBox Is Nothing Then Exit Sub
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            If intStyle = 2 Then
                    DrawEdge .hDC, PicRect, EDGE_RAISED Or BF_SOFT, BF_RECT
            ElseIf intStyle = -2 Then
                    DrawEdge .hDC, PicRect, EDGE_SUNKEN Or BF_SOFT, BF_RECT
            Else
                DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
            End If
        End If
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) '-10
            End If
            picBox.Print strName
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub

Public Function GetPersonnelDept(ByVal lngID As Long) As ADODB.Recordset
'功能：获取指定人员的所有部门
    Dim strSQL As String
 
    strSQL = "Select B.名称,B.ID From 部门人员 A, 部门表 B Where A.部门id = B.ID And A.人员id = [1] Order by 缺省 Desc"
    On Error GoTo errH
    Set GetPersonnelDept = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function TruncateDate(ByVal datFull As Date) As Date
'去掉日期中的时、分、秒
    TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
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
Public Function ReturnMovedExes(ByVal strNO As String, Optional ByVal bytType As Byte = 2, Optional ByVal strFormCaption As String) As Boolean
'功能:根据用户选择抽选后备数据表中的数据到当前数据表中
'参数:bytType表示单据类型,值::1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐；
'返回:用户选择取消操作,或者抽选数据转出失败,则返回False
    
    MsgBox "当前操作的单据" & strNO & "在后备数据表中!" & vbCrLf _
        & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
    ReturnMovedExes = False
    
'以下是抽选返回数据的过程，暂存，便于将来透明访问时重用
'    If MsgBox("当前操作单据" & strNO & "在后备数据表中,系统需要先把与此单据相关的数据转入到在线数据表才能继续!" & vbCrLf & _
'                             "确定要进行此操作吗?", vbInformation + vbYesNo, gstrSysName) = vbNo Then
'        ReturnMovedExes = False     '此句可省
'        Exit Function
'    End If
'
'    If zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption) Then
'        ReturnMovedExes = True
'    Else
'        '详细错误在之前的执行过程出错时给出
'        MsgBox "因系统错误,与该单据相关的数据未能转入到在线数据表." & vbCrLf & "操作未成功,请与系统管理员联系!", vbInformation, gstrSysName
'        ReturnMovedExes = False
'    End If
End Function

Public Sub zlSetCrlEnbled(ByVal objCrl As Object, blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定控件的Nabled属性,如果为False,同时需要设置相关的背景色
    '入参:objCrl-转入的指定控件
    '     blnEnabled-相关属性
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:44:25
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Select Case UCase(TypeName(objCrl))
    Case UCase("TextBox"), UCase("COMBOBOX")
        objCrl.Enabled = blnEnabled
        zlSetCtrolBackColor objCrl
    Case UCase("dtpicker"), UCase("frame"), UCase("CHECKBOX"), UCase("LABEL"), UCase("COMMANDBUTTON")
        objCrl.Enabled = blnEnabled
    Case Else
       ' objCrl.Enabled = blnEnabled
    End Select
End Sub
Public Sub zlSetCtrolBackColor(ByVal objCtl As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件背景色的颜色
    '入参:objCtl-转入的控件
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:43:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If objCtl.Enabled = False Then
        objCtl.BackColor = &H8000000F
    Else
        objCtl.BackColor = vbWhite
    End If
End Sub

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
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
ErrHand:
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
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
ErrHand:
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
    On Error GoTo ErrHand:
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
ErrHand:
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
    On Error GoTo ErrHand:
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
ErrHand:
End Sub

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
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
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Call OS.TaskbarHeight
End Function


Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:人员选择选择器
    '入参:cboSel-指定的部门选择部件
    '     rsPerson-指定的人员信息(ID,编号,姓名,简码)
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有-所有名称(所有人,所有操作员等)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String, strLike As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编号) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!姓名)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编号)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!姓名))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.TxtSelAll cboSel
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查部门信息是否加载编码
    '返回:显示编码,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 13:11:01
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If mlng部门编码平均长度 = 0 Then
        strSQL = "Select Avg(length(编码)) As 长度 From 部门表"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取部门编码的平均长度")
        mlng部门编码平均长度 = Val(Nvl(rsTemp!长度))
    End If
    '由于编码长度可能过长,无法显示部门的名称,因此自动显示和不显示编码,当大于5时,不显示.小于5时,显示
   zlIsShowDeptCode = mlng部门编码平均长度 <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
 
      
    
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.TxtSelAll cboDept
End Function

Public Function zlGetFeeFields(Optional strTableName As String = "门诊费用记录", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定表的值
    '入参：strTableName:如:门诊费用记录;住院费用记录;....
    '      blnReadDatabase-从数据库中读取
    '出参：
    '返回：字段集
    '编制：刘兴洪
    '日期：2010-03-10 10:41:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo ErrHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "门诊费用记录"
        zlGetFeeFields = "" & _
        "Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, " & _
        "姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, " & _
        "加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, " & _
        "发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, " & _
        "保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
        Exit Function
    Case "住院费用记录"
        zlGetFeeFields = "" & _
         " Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, " & _
         " 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, " & _
         " 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, " & _
         " 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, " & _
         " 结帐id , 结帐金额, 保险大类ID, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
         Exit Function
    Case "病人结帐记录"
        zlGetFeeFields = "Id, No, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期, 备注"
        Exit Function
    Case "病人预交记录"
        zlGetFeeFields = "" & _
        " Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, " & _
        " 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补"
        Exit Function
    Case "人员表"
        zlGetFeeFields = "" & _
        "Id, 编号, 姓名, 简码, 身份证号, 出生日期, 性别, 民族, 工作日期, 办公室电话, 电子邮件, 执业类别, 执业范围, " & _
        "管理职务, 专业技术职务, 聘任技术职务, 学历, 所学专业, 留学时间, 留学渠道, 接受培训, 科研课题, 个人简介, 建档时间, " & _
        "撤档时间, 撤档原因, 别名, 站点"
        Exit Function
    Case "票据领用记录"
        zlGetFeeFields = "ID,票种,使用类别,领用人,前缀文本,开始号码,终止号码,使用方式,登记时间,使用时间," & _
        "登记人,当前号码,剩余数量,批次,核对人,核对时间,核对结果,核对模式,备注,签字人,签字时间"
        Exit Function
    Case "票据使用明细"
        zlGetFeeFields = "ID,票种,号码,性质,原因,领用ID,打印ID,回收次数,使用时间,使用人,核对人,核对时间,核对结果,备注"
        Exit Function
    Case "人员缴款记录"
        zlGetFeeFields = "ID,单据ID,收款员,收款部门ID,结算方式,结算号,金额,摘要,截止时间,登记时间,登记人"
        Exit Function
    Case "消费卡领用记录"
        zlGetFeeFields = "ID,接口编号,领用人,前缀文本,开始卡号,终止卡号,使用方式,登记时间,使用时间," & _
        "登记人,当前卡号,剩余数量,批次,核对人,核对时间,核对结果,核对模式,备注,签字人,签字时间"
        Exit Function
    Case "消费卡使用记录"
        zlGetFeeFields = "ID,接口编号,卡号,性质,原因,领用ID,回收次数,使用时间,使用人,核对人,核对时间,核对结果,备注"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo ErrHand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取列信息", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!column_name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
ErrHand:
  zlGetFeeFields = "*"
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "门诊费用记录", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张数据表中的字段.类似于Select Id,....
    '入参：bytHistory-0-不包含历史数据,1-仅包含历史数据,2-两都都包含( select * from tablename Union select * from Htablename)
    '      strWhere-条件
    '      blnSubTable-是否子表
    '      strAliasName-别名
    '出参：
    '返回：select ID ... From tableName Union ALL
    '编制：刘兴洪
    '日期：2010-03-10 11:19:11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '无
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '仅历史
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '两者都包含
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function



Public Function Select人员选择器(ByVal frmMain As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng部门ID As Long = 0, _
    Optional lng人员ID As Long = 0, _
    Optional bln按部门人员显示 As Boolean = False, _
    Optional strSearchKey As String = "", _
    Optional str人员性质 As String = "", _
    Optional str管理职务 As String = "", _
    Optional str专业技术职务 As String = "", _
    Optional strTittle As String = "人员选择器", _
    Optional strNote As String = "请选择相关的人员", _
    Optional strNotFindMsg As String = "未找到指定的人员,请检查!", _
    Optional strShowField As String = "姓名", _
    Optional strShowSplit As String = "-") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的人员
    '入参:frmMain-调用的父窗口
    '     objCtl-控件(目前只支持文本框)
    '     strKey-输入的建值
    '     lng部门ID-如果不为零,找所有人员,否则, 找指定部门下的人员
    '     str人员性质: 以医生,医生1... 格式
    '     str管理职务及str专业技术职务: 以职务1,职务21... 格式
    '出参:lng人员id-返回人员ID
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, bytType As Byte, str人员性质Table As String, strWhere As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '功能：多功能选择器
    '参数：
    '     frmMain=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    Err = 0: On Error GoTo ErrHand:
    bytType = 0: strWhere = ""
    If str人员性质 <> "" Then
        str人员性质Table = ",人员性质说明 Q1,(Select Column_Value From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) Q2" & vbCrLf
        strWhere = strWhere & " And ( A.ID=Q1.人员ID and Q1.人员性质 = Q2.Column_Value ) " & vbCrLf
    End If
    If str管理职务 <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))  Where a.管理职务=Column_Value) " & vbCrLf
    If str专业技术职务 <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))  Where a.专业技术职务=Column_Value) " & vbCrLf
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
        If lng部门ID = 0 Then
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct A.ID,A.编号,A.姓名,A.别名,A.简码,A.性别,A.民族,A.出生日期,A.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 A " & str人员性质Table & _
                "   Where (A.姓名 like [1] or A.编号 like [1] or A.简码 like Upper([1]) or A.别名 like [1]) " & strWhere & zl_获取站点限制(True, "A") & "" & _
                "       and (A.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                "   order by A.编号"
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 a,部门人员 C " & str人员性质Table & _
                "   Where a.id=c.人员id and c.部门Id=[2]   " & strWhere & zl_获取站点限制(True, "a") & _
                "       and (a.姓名 like [1] or a.编号 like [1] or a.简码 like Upper([1]) or a.别名 like [1]) " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & _
                "   order by 编号"
        End If
     Else
        If lng部门ID = 0 Then
            If bln按部门人员显示 Then
                gstrSQL = "" & _
                "   Select /*+ rule */  id," & IIf(gstrNodeNo <> "-", "1 as 级数ID,-1*NULL as 上级ID", "Level as 级数ID,上级id") & " ,编码,名称,0 末级,'' as 别名,'' as 简码,''as 性别,''as 民族, to_date(Null,'yyyy-mm-dd')  as 出生日期, '' as  办公室电话 ,'' 执业类别, '' 管理职务,'' 专业技术职务" & _
                "   From 部门表 " & _
                "   where 撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd') " & zl_获取站点限制() & _
                    IIf(gstrNodeNo <> "-", "", "   Start with 上级id is null connect by prior id=上级id ") & _
                "   union all " & _
                "   Select  distinct a.ID,999999 AS 级数ID,b.部门id as 上级ID,a.编号,a.姓名,1 as 末级,别名,简码,性别,民族,出生日期,办公室电话,A.执业类别,A.管理职务,A.专业技术职务 " & _
                "   From 人员表 a,部门人员 b  " & str人员性质Table & _
                "   Where a.id=b.人员id and b.缺省=1  " & strWhere & zl_获取站点限制(True, "a") & _
                "         And (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                "   Order by 级数ID,编码"
                bytType = 2
            Else
                gstrSQL = "" & _
                    "   Select  /*+ rule */  distinct A.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                    "   From 人员表 A " & str人员性质Table & _
                    "   Where (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & strWhere & zl_获取站点限制(True, "a") & _
                    "   order by a.编号"
            End If
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 a,部门人员 C " & str人员性质Table & _
                "   Where a.id=c.人员id and c.部门Id=[2] " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & strWhere & zl_获取站点限制(True, "a") & _
                "   order by a.编号"
        End If
    End If
   
   
   '坐标定位
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlcontrol.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, strTittle, bytType = 2, strSearchKey, strNote, bytType = 2, False, Not (bytType = 2), sngX, sngY, lngH, blnCancel, False, False, strKey, lng部门ID, str人员性质, str管理职务, str专业技术职务)
    
    lng人员ID = 0
    If blnCancel = True Then
        Call zlcontrol.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        Call zlcontrol.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    Call zlcontrol.ControlSetFocus(objCtl, True)
    If bytType = 2 Then
        strShowField = "," & strShowField & ",M_刘,"
        strShowField = Replace(strShowField, ",编号,", ",编码,")
        strShowField = Replace(strShowField, ",姓名,", ",名称,")
        strShowField = Mid(strShowField, 2)
        strShowField = Replace(strShowField, ",M_刘,", "")
    End If
    
    '设置相关的值
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!ID)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!ID) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
        objCtl.Tag = Val(rsTemp!ID)
        zlCommFun.PressKey vbKeyTab
    End Select
    lng人员ID = Val(Nvl(rsTemp!ID))
    rsTemp.Close
    Select人员选择器 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function
Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_GetFieldValue(ByVal rsTemp As ADODB.Recordset, _
    Optional ByVal strShowFields As String = "编码,名称", _
    Optional ByVal strShowSplit As String = "-") As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:返回显示字段的相关值
    '入参:rsTemp-记录集
    '     strShowFields-显示的字段
    '     strShowSplit-显示的分离符
    '出参:
    '返回:成功,返回相关的字段值
    '编制:刘兴洪
    '日期:2009-03-06 11:59:19
    '-----------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strValue As String, strLeft As String, strRight As String
    varData = Split(strShowFields, ",")
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.RecordCount = 0 Then Exit Function
    
    Select Case strShowSplit
    Case "[", "[]", "]"
        strLeft = "[": strRight = "]"
    Case "〖〗", "〖", "〗"
        strLeft = "〖": strRight = "〗"
    Case "【】", "【", "】"
        strLeft = "【": strRight = "】"
    Case "（）", "（", "）"
        strLeft = "（": strRight = "）"
    Case "〔〕", "〔", "〕"
        strLeft = "〔": strRight = "〕"
    Case "〈〉", "〈", "〉"
        strLeft = "〈": strRight = "〉"
    Case "［］", "［", "］"
        strLeft = "［": strRight = "］"
    Case "[]", "[", "]"
        strLeft = "[": strRight = "]"
    Case "｛｝", "｛", "｝"
        strLeft = "｛": strRight = "｝"
    Case "{}", "{", "}"
        strLeft = "{": strRight = "}"
    Case "「」", "「", "」"
        strLeft = "「": strRight = "」"
    Case "『』", "『", "』"
        strLeft = "『": strRight = "』"
    Case Else
        strLeft = "": strRight = strShowSplit
    End Select
    
    strValue = ""
    With rsTemp
        For i = 0 To UBound(varData) - 1
            strValue = strValue & strLeft & Nvl(.Fields(varData(i))) & strRight
        Next
        strValue = strValue & Nvl(.Fields(varData(UBound(varData))))
    End With
    zl_GetFieldValue = strValue
End Function

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
    i = cllData.Count + 1
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
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function zlIsOnlyNum(ByVal strAsk As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定字符串是否全部由数字构成
    '入参:strAsk-需要判断的字符
    '出参:
    '返回:如果全用数字构成，返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-17 11:19:15
    '说明:
    '     isnumberic不能检查这些:-099.22,22d2。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            strTemp = Mid(Trim(strAsk), i, 1)
            If InStr("0123456789", strTemp) = 0 Then Exit Function
        Next
        zlIsOnlyNum = True
    End If
End Function

Public Function Get结算方式() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算方式
    '返回:结算方式集
    '编制:刘兴洪
    '日期:2013-09-04 17:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "" & _
    "   Select 编码,名称,性质,nvl(应收款,0) as 应收款,nvl(应付款,0) as 应付款," & _
    "               nvl(缺省标志,0) as 缺省标志  " & _
    "   From 结算方式"
    If mrsPayMode Is Nothing Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "获取结算方式")
    ElseIf mrsPayMode.State <> 1 Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "获取结算方式")
    End If
    Set Get结算方式 = mrsPayMode
End Function
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
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub



Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '功能：由用户输入的部份单号，返回全部的单号。
    '参数：intNum=项目序号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intTYPE As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    Err = 0: On Error GoTo errH:
    
    strSQL = "Select 编号规则,Sysdate as 日期,最大号码 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取号码控制", intNum)
    If rsTmp.EOF Then GetFullNO = strNO: Exit Function
    Select Case Val(Nvl(rsTmp!编号规则))
    Case 0, 1 '0-按年顺序编号,1-按日顺序编号
        If Len(strNO) >= 8 Then
            GetFullNO = Right(strNO, 8)
            Exit Function
        ElseIf Len(strNO) = 7 Then
            GetFullNO = zlStr.PrefixNO & strNO
            Exit Function
        End If
        GetFullNO = strNO
        dtCurDate = Date
        If Not rsTmp.EOF Then
            intTYPE = Val("" & rsTmp!编号规则)
            dtCurDate = rsTmp!日期
            strMaxNo = Nvl(rsTmp!最大号码)
        End If
        strYearStr = zlStr.PrefixNO
        If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
        If intTYPE = 1 Then
            '按日编号
            strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
            GetFullNO = zlStr.PrefixNO & strSQL & Format(Right(strNO, 4), "0000")
            Exit Function
        End If
        '按年编号
        If Len(strNO) = 6 Then
            GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
        End If
        GetFullNO = Left(strMaxNo, 2) & zlStr.Lpad(Right(strNO, 6), 6, "0", True)
    Case 2  '2-按科室分月或分日编号需要读取科室号码表,
    Case 3   '3-按年月日+顺序号(年取两位,顺序号取4位)
        If Len(strNO) <= 6 Then
            GetFullNO = Format(rsTmp!日期, "YYMMDD") & zlStr.Lpad(strNO, 6, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 8 Then
            GetFullNO = Format(rsTmp!日期, "YYMM") & zlStr.Lpad(strNO, 8, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 10 Then
            GetFullNO = Format(rsTmp!日期, "YY") & zlStr.Lpad(strNO, 10, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 12 Then
            GetFullNO = zlStr.Lpad(strNO, 12, "0", True)
            Exit Function
        End If
    Case 4    '4-按执行科室分期间编号(年(期间表中的年)+执行科室编号+月份(期间表中的月)+顺序号)
    Case 5    '5-按年月进行编号(yyyyMM000000)
    Case Else
    End Select
    GetFullNO = strNO
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get轧帐结算性质(ByVal strRollingType As String, _
    ByRef strOut结算性质 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次轧帐的结算性质
    '入参:strRollingType-轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
    '出参:strOut结算性质-返回本次的结算性质,多个用逗号分隔,比如:,2,...
    '     如果是所有类别或预交或消费卡,则返回空
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-03-05 15:04:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strOut结算性质 = ""
    On Error GoTo errHandle

    '预交款填NULL,2-结帐,3-收费,4-挂号,5-就诊卡,6-补充医保结
    'strRollingType:1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
    If InStr("," & strRollingType & ",", ",1,") > 0 Then  '收费
        strOut结算性质 = "3,6,"
    End If
    If InStr("," & strRollingType & ",", ",3,") > 0 Then  '结帐
        strOut结算性质 = strOut结算性质 & "2,"
    End If
    If InStr("," & strRollingType & ",", ",4,") > 0 Then  '挂号
        strOut结算性质 = strOut结算性质 & "4,"
    End If
    If InStr("," & strRollingType & ",", ",5,") > 0 Then  '就诊卡
        strOut结算性质 = strOut结算性质 & "5,"
    End If
    If strOut结算性质 <> "" Then strOut结算性质 = "," & strOut结算性质
    Get轧帐结算性质 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CurrentIsBill(ByVal int票种 As gBillType) As Boolean
    '根据票种判断是否为票据
    '返回：如果是票据，返回TRUE，否则返回False
    Select Case int票种
    Case gBillType.收费收据, gBillType.预交收据, gBillType.结帐收据, gBillType.挂号收据
        CurrentIsBill = True
    Case Else
        CurrentIsBill = False
    End Select
End Function
