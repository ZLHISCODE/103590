Attribute VB_Name = "mdlMspManage"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
'类型定义
Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018

    原始单据 = 0
    冲销记录 = &HFF
    停用项目 = &H8000000C
    启用项目 = 0

    公共模块色 = &HC00000

    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
End Enum

'用户信息
Public Type USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    模块权限 As String
    报告权限 As String
    单位名称 As String
    部门名称 As String
End Type

'系统参数信息
Public Type SYSPARAM_INFO
    项目输入匹配方式 As Integer '0-双向;1-从左
    收费诊疗项目匹配 As Integer
    系统号 As Long
    系统名称 As String
    产品名称 As String
    模块号 As Long
    所有者 As String
End Type

'选择多个文件
Public Type DlgFileInfo
    iCount As Long   '文件数
    sPath As String  '选择路径
    sFile() As String   '文件名
End Type

Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
'----------------------------------------------------------------------------------------------------------------------
'全局变量申明

Public gclsDataOracle As zlDataOracle.clsDataOracle


Public gobjFSO As New Scripting.FileSystemObject        'FSO对象
Public ParamInfo As SYSPARAM_INFO
Public UserInfo As USER_INFO
Public gblnInsure As Boolean
Public gstrSQL As String
Public gblnShowInTaskBar As Boolean
Public gfrmMain As Object
Public glngTXTProc As Long                              '保存默认的消息函数的地址
Public glngShareUseID As Long
Public gblnPersonSet As Boolean
Public gfrmPubResource As frmPubResource
Public gclsMsgBase As New clsBusiness
Public gstrSysName As String
Public glngSys As Long
Public Const M_MAX_HEIGHT = 16150  '卡片的最大高度
Public Const M_MAX_WIDTH = 15005   ''卡片的最大宽度
Public Const M_MIN_HEIGHT = 3900   '卡片的最小高度
Public Const M_MIN_WIDTH = 6000        '卡片的最小宽度

'公共图标定义
Public Const Icon_History = 1000
Public Const Icon_Charge = 1001
Public Const Icon_Item = 1002
Public Const Icon_Report = 1003
Public Const Icon_Archives = 1004
Public Const Icon_Package = 1005
Public Const Icon_WaitPerson = 1006
Public Const Icon_NowPerson = 1007
Public Const Icon_OverPerson = 1008

Public Const Icon_Group = 1009
Public Const Icon_Single = 1010
'Private mclsUnzip As New clsUnZip

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'----------------------------------------------------------------------------------------------------------------------
'模块变量申明
Public zlCommFun As New clsCommFun
Public zlDataBase As New clsDatabase
Public zlComLib As New clsComLib
Public zlControl As New clsControl


'######################################################################################################################
'过程清单

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '功能：去掉TextBox的默认右键菜单
    '参数：
    '返回：
    '说明：如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '******************************************************************************************************************
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '功能：获取登陆用户信息
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    Set rsTmp = zlDataBase.GetUserInfo
    
    UserInfo.姓名 = UserInfo.用户名
    If Not rsTmp.EOF Then
    
        UserInfo.ID = zlCommFun.NVL(rsTmp("ID").Value, 0)
        UserInfo.编号 = zlCommFun.NVL(rsTmp("编号").Value)
        UserInfo.部门ID = zlCommFun.NVL(rsTmp("部门ID").Value, 0)
        UserInfo.简码 = zlCommFun.NVL(rsTmp("简码").Value)
        UserInfo.姓名 = zlCommFun.NVL(rsTmp("姓名").Value)
        UserInfo.部门名称 = zlCommFun.NVL(rsTmp("部门名").Value)
        
        GetUserInfo = True

    End If
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter() = 1 Then
        Resume
    End If
    
    Call zlComLib.SaveErrLog
End Function


Public Function NVL(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '功能:取某字段的值
    '参数:rsObj          被检查的字段
    '     varValue       当rsObj为NULL值时的取新值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
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

Public Function GetImageList(Optional ByVal intIconSize As Integer = 16) As ImageList
    
'    If gfrmPubResource Is Nothing Then
'        Set GetImageList = frmPubResource.GetImageList(intIconSize)
'    Else
'        Set GetImageList = gfrmPubResource.GetImageList(intIconSize)
'    End If
    
End Function


'Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
'    '******************************************************************************************************************
'    '功能:
'    '参数:
'    '返回:
'    '******************************************************************************************************************
'    Dim blnTran As Boolean
'    Dim intLoop As Integer
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If rs.RecordCount > 0 Then
'        If Len(strTitle) = 0 Then strTitle = ParamInfo.系统名称
'        blnTran = True
'
'        If blnHaveTrans Then gcnOracle.BeginTrans
'
'        rs.MoveFirst
'
'        For intLoop = 1 To rs.RecordCount
'
'            strSQL = CStr(rs("SQL").Value)
'            Call zlDataBase.ExecuteProcedure(strSQL, strTitle)
'
'            rs.MoveNext
'        Next
'
'        If blnHaveTrans Then gcnOracle.CommitTrans
'        blnTran = False
'    End If
'
'    SQLRecordExecute = True
'
'    Exit Function
'errHand:
'
'    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
'
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'
'
'End Function



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

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '帮助主题
    
        Call zlComLib.ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.系统号) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlComLib.zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlComLib.zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlComLib.zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '关于
        
        Call zlComLib.ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "0") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.Cols = objVsf.Cols - lngNotPrintCols
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
            lngPrintCol = lngPrintCol + 1
            objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
            objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
            If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                objPrintVsf.ColAlignment(lngPrintCol) = 4
            Else
                objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
            End If
        End If
    Next
    
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                lngPrintCol = lngPrintCol + 1
                
                If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                    objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                Else
                    strFormat = objVsf.ColFormat(lngCol)
                    If strFormat = "" Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                    Else
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                    End If
                End If
            End If
        Next
'        Call gclsBase.SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
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
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
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
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
        
    DockPannelInit = True
    
End Function



Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "【" & strCode & "】" & strName
    Else
        AppendCode = strName
    End If
End Function

Public Function GetCode(ByVal strName As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If strName <> "" And InStr(strName, "】") > InStr(strName, "【") Then
        GetCode = Mid(strName, InStr(strName, "【") + 1, InStr(strName, "】") - InStr(strName, "【") - 1)
    End If
    
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

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function
Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, ParamInfo.系统名称
    
End Sub

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

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

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
'    gblnPersonSet = (Val(zlDataBase.GetPara("使用个性化风格")) = 1)

    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    Dim cbrControl3 As CommandBarControl
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.Visible = cbrControl.Visible
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
            
                Set cbrControl3 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrControl3.Parameter = cbrControl2.Parameter
                cbrControl3.Visible = cbrControl2.Visible
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub CloseRecord(rs As ADODB.Recordset)
    '******************************************************************************************************************
    '功能:关闭连接
    '参数:
    '返回:
    '******************************************************************************************************************
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Function WriteBinByFile(strFile As String, objField As Field) As Boolean
    '******************************************************************************************************************
    '功能:写入文件
    '参数:
    '返回:
    '******************************************************************************************************************
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, intFile, i As Long
    Dim arrBin() As Byte
    
    On Error GoTo errH
    
    intFile = FreeFile
    Open strFile For Binary Access Read As intFile
    lngFileSize = LOF(intFile)
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    objField.Value = Null
    For i = 0 To intBolcks
        If i = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim arrBin(lngCurSize - 1) As Byte
        Get intFile, , arrBin()
        objField.AppendChunk arrBin()
    Next
    Close intFile
    WriteBinByFile = True
    Exit Function
errH:
    Close intFile
End Function

Public Function GetDlgFileInfo(strFileName As String) As DlgFileInfo
    '******************************************************************************************************************
    '功能:返回CommonDialog所选择的文件数量、路径和文件名
    '参数:strFileName为CommonDialog的Filename属性,DlgFileInfo。这是一个自定义类型，其中iCount返回所选择文件的个数，sPath返回所选
    '返回:返回CommonDialog所选择的文件数量、路径和文件名
    '******************************************************************************************************************
    Dim sPath, tmpStr As String
    Dim sFile() As String
    Dim iCount As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    sPath = CurDir()
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '将文件名与路径分离
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        '选择了多个文件(分离后第一个字符为Chr$(0))
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFile(iCount)
            Else
                sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        '只选择了一个文件(注意：根目录下的文件名除去路径后左边没有"\"）
        iCount = 1
        ReDim Preserve sFile(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFile(iCount) = tmpStr
    End If
    
    GetDlgFileInfo.iCount = iCount
    ReDim GetDlgFileInfo.sFile(iCount)
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgFileInfo.sPath = sPath
    
    For i = 1 To iCount
        GetDlgFileInfo.sFile(i) = sFile(i)
    Next
    
    Exit Function

errHandle:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 4000
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function AddPeriodToComboBox(ByRef cbo As Object)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    With cbo
        .Clear
        .AddItem "所  有"
        .AddItem "今  天"
        .AddItem "昨  天"
        .AddItem "本  周"
        .AddItem "本  月"
        .AddItem "本  季"
        .AddItem "本半年"
        .AddItem "本  年"
        .AddItem "前三天"
        .AddItem "前一周"
        .AddItem "前半月"
        .AddItem "前一月"
        .AddItem "前二月"
        .AddItem "前三月"
        .AddItem "前半年"
        .AddItem "前一年"
        .AddItem "前二年"
        .AddItem "自定义"
    End With
    
    AddPeriodToComboBox = True
    
End Function

Public Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "自定义" Then
        '自定义:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), zlDataBase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), zlDataBase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "所  有"
        GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "当  时"      '当时
        GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDataBase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDataBase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function


'Public Function ShowPubSelect(ByVal frmParent As Object, _
'                                ByVal obj As Object, _
'                                ByVal mrsInfoTree As ADODB.Recordset, _
'                                ByVal strDataInfo As String, _
'                                Optional ByVal lngCX As Long = 9000, _
'                                Optional ByVal lngCY As Long = 4500) As String
'    '******************************************************************************************************************
'    '功能：打开树型+列表结构,应用于表格控件
'    '参数：
'    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
'    '返回：0:取消选择;1:选择;2:无数据返回
'    '******************************************************************************************************************
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngObjHeight As Long
'    Dim rs As New ADODB.Recordset
'    Dim objPoint As POINTAPI
'    Dim objSelectDialog As frmEventMsgEditSelect
'
'    On Error GoTo errHand
'
'
'    If obj Is Nothing Then
'        lngX = (Screen.Width - lngCX) / 2
'        lngY = (Screen.Height - lngCY) / 2
'        lngObjHeight = 0
'    Else
'        Call ClientToScreen(obj.hWnd, objPoint)
'
'        lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
'        lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
'        lngObjHeight = obj.CellHeight
'
'    End If
'
'    Set objSelectDialog = New frmEventMsgEditSelect
'
'    ShowPubSelect = objSelectDialog.ShowSelect(frmParent, mrsInfoTree, lngX, lngY, lngCX, lngCY, strDataInfo)
'
'    Exit Function
'
'errHand:
''    If ErrCenter = 1 Then
''        Resume
''    End If
'End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003, _
                                Optional ByVal bytPosition As XTPTabPosition = xtpTabPositionTop) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbc

        With .PaintManager

            .Appearance = bytAppearance
            .ClientFrame = xtpTabFrameSingleLine
            .Color = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
            .BoldSelected = True
            .Position = bytPosition
        End With

'        Set .Icons = frmPubResource.imgPublic.Icons
    End With

    TabControlInit = True

End Function

