Attribute VB_Name = "mdlPubDefine"
Option Explicit

'常量定义
'######################################################################################################################

'公共部份菜单ID定义:*表示有图标
'*********************************************************************
Public Const conMenu_FilePopup = 1 '文件
Public Const conMenu_ManagePopup = 2 '管理
Public Const conMenu_EditPopup = 3 '编辑
Public Const conMenu_ReportPopup = 4 '报表
Public Const conMenu_ViewPopup = 7 '查看
Public Const conMenu_ToolPopup = 8 '工具
Public Const conMenu_HelpPopup = 9 '帮助

''文件菜单
Public Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Public Const conMenu_File_Preview = 102         '*预览(&V)
Public Const conMenu_File_Print = 103           '*打印(&P)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_Exit = 191            '*退出(&X)

'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Page = 745                '查看部门
Public Const conMenu_View_Refresh = 791              '*刷新(&R)

Public Const conMenu_View_Navigatebeginning = 7401           '*第一个(&F)
Public Const conMenu_View_Navigateleft = 7402                '*上一个(&F)
Public Const conMenu_View_Navigateright = 7403               '*下一个(&F)
Public Const conMenu_View_Navigateend = 7404                 '*最后一个(&F)

'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

'其它常量定义
'*********************************************************************
'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'CommandBar虚拟键
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_PAGEUP = &H21
Public Const VK_PAGEDOWN = &H22
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const VsModiBackColor = &HD6FFCA        'vs控件，可编辑单元的背景色
'*********************************************************************

Public Type SYSPARAM_INFO
    系统号 As Long
    系统名称 As String
    产品名称 As String
    模块号 As Long
    所有者 As String
End Type

Public ParamInfo As SYSPARAM_INFO
Public grsData As ADODB.Recordset
Public grsPage As ADODB.Recordset
Public grsList As ADODB.Recordset
Public grsTempFile As ADODB.Recordset
Public gintStartPage As Integer
Public gstrNo As String
Public gobjRect As USERRECT
Public gobjFont As USERFONT
Public gobjPaper As USERPAPER
Public gobjDraw As Object
Public gclsDataSources As clsDataSources
Public glngVirtualPages As Long
Public gdblWaitTime As Double '打印PDF间隔

'
'######################################################################################################################

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '功能：
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, "zl9OpsFormat"
    
End Sub

Public Function CommandBarInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
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

    Set cbsMain.Icons = frmPubResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
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
        objControl.Id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.type = xtpControlButton Or objControl.type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
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
    
    Select Case Control.Id
    Case conMenu_File_PrintSet '打印设置
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_Help_Help              '帮助主题
    
'        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home 'Web上的中联
        
        Call zlHomePage(frmMain.hwnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.hwnd)
        
    Case conMenu_Help_Web_Mail '发送反馈
        
        Call zlMailTo(frmMain.hwnd)
            
    Case conMenu_Help_About '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.Id
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function InsertData(ByVal str类别 As String, _
                                ByVal str对象 As String, _
                                Optional ByVal bytHAlignment As Byte = 1, _
                                Optional ByVal str内容 As String, _
                                Optional ByVal bytVAlignment As Byte = 2, _
                                Optional ByVal blnWordWarp As Boolean, _
                                Optional ByVal intRows As Integer = 1, _
                                Optional ByVal str标志 As String = "", _
                                Optional ByVal blnAutoFit As Boolean = False, _
                                Optional ByVal blnDebug As Boolean = False, _
                                Optional ByVal strPrex As String = "A", _
                                Optional ByVal bytAngle As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    On Error GoTo errHand
    
    grsData.AddNew
    
    gstrNo = Format(Val(gstrNo) + 1, "0000000000")
    
    grsData("序号").value = UCase(strPrex) & gstrNo
    grsData("调试").value = IIf(blnDebug, 1, 0)
    grsData("类别").value = str类别
    grsData("页号").value = gobjRect.Page
    grsData("对象").value = str对象
    grsData("内容").value = str内容
    grsData("X0").value = gobjRect.X0
    grsData("Y0").value = gobjRect.Y0
    grsData("X1").value = gobjRect.X1
    grsData("Y1").value = gobjRect.Y1
    grsData("B0").value = gobjRect.B0
    grsData("R0").value = gobjRect.R0
    grsData("字体").value = gobjFont.Name
    grsData("前景色").value = gobjFont.ForeColor
    grsData("背景色").value = gobjFont.BackColor
    grsData("大小").value = gobjFont.size
    grsData("粗体").value = IIf(gobjFont.Bold, 1, 0)
    grsData("斜体").value = IIf(gobjFont.Italic, 1, 0)
    grsData("下划线").value = IIf(gobjFont.Underline, 1, 0)
    grsData("横向对齐").value = bytHAlignment                                   '1-左;2-中;3-右
    grsData("纵向对齐").value = bytVAlignment                                   '1-左;2-中;3-右
    grsData("自动换行").value = IIf(blnWordWarp, 1, 0)
    grsData("线条宽度").value = IIf(gobjFont.LineWidth = 0, 1, gobjFont.LineWidth)
    grsData("线条类型").value = gobjFont.LineStyle
    grsData("行数").value = intRows
    grsData("自动适应").value = IIf(blnAutoFit, 1, 0)
    grsData("旋转角度").value = bytAngle
    
    InsertData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetLines(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As Integer
    '******************************************************************************************************************
    '功能：获取需要的行数，因为有可能要换行
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim sglSingleChar As Single
    Dim lngChars As Long

    sglSingleChar = objDraw.TextWidth("A")

    lngChars = lngCX \ sglSingleChar

    strSQL = "Select zl_GetTextRows([1],[2]) As 行数 From Dual"
    Set rs = zldatabase.OpenSQLRecord(strSQL, "", strText, lngChars)
    If rs.BOF = False Then
        GetLines = zlStr.NVL(rs("行数").value)
    End If

    If GetLines = 0 Then GetLines = 1
End Function

Public Function GetLineText2(ByVal objDraw As Object, ByVal strText As String, ByVal intRow As Integer, ByVal lngCX As Long) As String
    '******************************************************************************************************************
    '功能：获取指定行的数据，方法是先求出可以最多输出多少个字符，然后调用过程“Get_LineText”求出指定行内容
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim sglSingleChar As Single
    Dim lngChars As Long

    GetLineText2 = strText

    sglSingleChar = objDraw.TextWidth("A")

    lngChars = lngCX \ sglSingleChar

    strSQL = "Select zl_GetText([1],[2],[3]) As 行文本 From Dual"
    Set rs = zldatabase.OpenSQLRecord(strSQL, "", strText, lngChars, intRow)
    If rs.BOF = False Then
        GetLineText2 = zlStr.NVL(rs("行文本").value)
    End If

End Function

Public Function GetLineText(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：获取指定行的数据，方法是先求出可以最多输出多少个字符，然后调用过程“Get_LineText”求出指定行内容
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsLine As ADODB.Recordset
    Dim strLineText  As String
    Dim lngChar As Long
    Dim lngRow As Long
    Dim strChar As String
    Dim strLastChar As String
    Dim strNextChar As String
    Dim blnCrlf As Boolean
    Dim lngTextLength As Long
    
    On Error GoTo errHand
    
    Set rsLine = New ADODB.Recordset
    With rsLine
        .Fields.Append "行号", adBigInt
        .Fields.Append "内容", adVarChar, 1000
        .Open
    End With
    
    lngRow = 0
    strChar = ""
    strLastChar = ""
    strNextChar = ""
    blnCrlf = False
    lngTextLength = Len(strText)
    
    For lngChar = 1 To lngTextLength
        
        blnCrlf = False
        
        strChar = Mid(strText, lngChar, 1)
        strLastChar = strChar
        
        Select Case Asc(strChar)
        Case 13
            '需要再判断下一个字符是否为换行符
            If lngChar + 1 <= lngTextLength Then
                strNextChar = Mid(strText, lngChar + 1, 1)
                        
                If Asc(strNextChar) = 10 Then
                    '是回车换行符
                    
                    strLastChar = vbCrLf
                    
                    '判断此回车换行符所处的位置，主要判断它的前一字符是不是也是回车换行符或首字符，如果是的则为新行
                    If lngChar = 1 Or strLastChar = vbCrLf Then
                        blnCrlf = True
                    End If
                    
                    '循环变量累加1个计数器
                    lngChar = lngChar + 1
                End If
            
            End If
        Case 10
            '直接换行
            blnCrlf = True
        End Select
        
        If objDraw.TextWidth(strLineText & strChar) > lngCX Or blnCrlf Then
            
            lngRow = lngRow + 1
            rsLine.AddNew
            rsLine("行号").value = lngRow
            rsLine("内容").value = strLineText
            
            If blnCrlf = False Then
                strLineText = strChar
            Else
                strLineText = ""
            End If
        Else
            strLineText = strLineText & strChar
        End If
    Next
    
    If strLineText <> "" Then
        lngRow = lngRow + 1
        rsLine.AddNew
        rsLine("行号").value = lngRow
        rsLine("内容").value = strLineText
    End If
    
    Set GetLineText = rsLine
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AppendListData(ByVal strListName As String, ByVal bytList As Byte, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：添加到目录索引
    '参数：
    '返回：
    '******************************************************************************************************************
    
    grsList.AddNew
    
    grsList("目录页号").value = intPage - glngVirtualPages
    grsList("目录名称").value = strListName
    grsList("目录级数").value = bytList
    grsList("目录性质").value = 1
    
    AppendListData = True
    
End Function

Public Function CreateTmpFile(Optional ByVal strFile As String = "zl9PeisGroupRpt.tmp") As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim strFileTemp As String
    Dim strTempPath As String
    
    strFileTemp = OS.TempPath
    
    CreateTmpFile = strTempPath & strFile

End Function


