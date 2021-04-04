VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmLabMainLJ 
   Caption         =   "质控查询"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmLabMainLJ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11760
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo项目 
      Height          =   300
      Left            =   1860
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
   Begin C1Chart2D8.Chart2D chtCopy 
      Height          =   435
      Left            =   1515
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   765
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   1349
      _ExtentY        =   767
      _StockProps     =   0
      ControlProperties=   "frmLabMainLJ.frx":058A
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabMainLJ.frx":0BE9
      Left            =   975
      Top             =   330
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabMainLJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmChartLJ As New frmQCChartLJ                 'LJ控制图窗格
Private mfrmQCTodayReport As New frmQCTodayReport       '填写失控记录
Private mlngSampleID As Long                            '标本ID
Private mstrQCID As String                              '质控品ID
Private mlngMachineID As Long                           '仪器ID
Private mlngResult As Long                              '普通结果ID
Private mEditMode As Integer                            '编辑模式 0=非编辑 1=正在编辑
Private mstrPigeonhole As String                        '归档人
Private mstrReportMan As String                         '报告人
Private mstrStart As String, mstrEnd As String

Private Sub cbo项目_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strStartDate As String, strEndDate As String
    Dim strDateSpace As String
    Dim strNowDate As String
    
    If Me.cbo项目.ListCount = 0 Then Exit Sub
    
    '得当前时间
    gstrSql = "select nvl(核收时间,sysdate) as 核收时间 from 检验标本记录 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
    If rsTmp.EOF = True Then
        MsgBox "没有找到对应的标本!", vbInformation, gstrSysName: Exit Sub
    End If
    strNowDate = Nvl(rsTmp("核收时间"))
    
    strStartDate = Format(getMonthFirst(CDate(strNowDate)), "yyyy-mm-dd"): strEndDate = Format(getMonthLast(CDate(strNowDate)), "yyyy-mm-dd")

    mstrStart = strStartDate
    mstrEnd = strEndDate
    '-----------------------------------------------------------------------------------------------------------------------
    '得到质控品
    mstrQCID = ""
'    gstrSql = "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品, M.水平" & vbNewLine & _
'            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X, ( Select Distinct 仪器id From 检验质控记录 Where 标本id = [1] ) Y " & vbNewLine & _
'            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = Y.仪器id And I.项目id = [2] And" & vbNewLine & _
'            "      X.期间 = [3]" & vbNewLine & _
'            "Order By M.开始日期, M.水平"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, CLng(Me.cbo项目.ItemData(Me.cbo项目.ListIndex)), _
                strDateSpace)
    gstrSql = "Select Distinct M.ID, '' As 选择, M.批号, M.名称 || ', 水平:' || M.水平 As 质控品, M.水平, M.开始日期" & vbNewLine & _
        "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X, (Select Distinct 仪器id From 检验质控记录 Where 标本id = [1]) Y" & vbNewLine & _
        "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = Y.仪器id And I.项目id = [2] And" & vbNewLine & _
        "      To_Date([3], 'YYYY-MM-DD') Between X.开始日期 And Nvl(X.结束日期, M.结束日期) and " & vbNewLine & _
        "      To_Date([4], 'YYYY-MM-DD') Between X.开始日期 And Nvl(X.结束日期, M.结束日期)" & vbNewLine & _
        "Order By M.开始日期, M.水平"
    
    gstrSql = "Select Id,选择,批号,质控品,水平,min(开始日期) As 开始日期,Min(结束日期) As 结束日期" & vbNewLine & _
            "From (" & vbNewLine & _
            "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品, M.水平, to_Char(X.开始日期,'yy-MM-dd') as 开始日期,to_char(Nvl(X.结束日期, M.结束日期),'yy-MM-dd')  as 结束日期" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X, (Select Distinct 仪器id From 检验质控记录 Where 标本id = [1]) Y" & vbNewLine & _
            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = Y.仪器ID And I.项目id = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期))" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品, M.水平, to_Char(X.开始日期,'yy-MM-dd') as 开始日期,to_char(Nvl(X.结束日期, M.结束日期),'yy-MM-dd')  as 结束日期" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X, (Select Distinct 仪器id From 检验质控记录 Where 标本id = [1]) Y" & vbNewLine & _
            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = Y.仪器ID And I.项目id = [2] And" & vbNewLine & _
            "        (  (X.开始日期 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "         Or" & vbNewLine & _
            "          (nvl(X.结束日期,Sysdate) Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')+1-1/24*60*60)" & vbNewLine & _
            "         )" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By      Id,选择,批号,质控品,水平" & vbNewLine & _
            "Order By 质控品,水平"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, CLng(Me.cbo项目.ItemData(Me.cbo项目.ListIndex)), _
                CStr(Format(mstrStart, "yyyy-mm-dd")), CStr(Format(mstrEnd, "yyyy-mm-dd")))
                
    strDateSpace = ""
    Do Until rsTmp.EOF
        mstrQCID = mstrQCID & "," & Val(Nvl(rsTmp("ID")))
        strDateSpace = strDateSpace & ";" & Val("" & rsTmp("ID")) & "=" & Format("" & rsTmp("开始日期"), "yyyy-MM-dd") & "," & Format("" & rsTmp("结束日期"), "yyyy-MM-dd")
        rsTmp.MoveNext
    Loop
    mstrQCID = Mid(mstrQCID, 2)
    If strDateSpace <> "" Then strDateSpace = Mid(strDateSpace, 2)
    '-----------------------------------------------------------------------------------------------------------------------
    mfrmChartLJ.zlRefresh mstrQCID, cbo项目.ItemData(cbo项目.ListIndex), Format(strStartDate, "yyyy-mm-dd"), _
                        Format(strEndDate, "yyyy-mm-dd"), strDateSpace
    gstrSql = "select ID from 检验普通结果 where 检验标本id = [1] and 检验项目id = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, cbo项目.ItemData(cbo项目.ListIndex))
    If rsTmp.EOF = False Then mlngResult = rsTmp("ID")
    mfrmQCTodayReport.zlRefresh mlngResult
    
    gstrSql = "select 报告人, 归档人 from 检验质控报告 where 结果id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResult)
    If rsTmp.EOF = False Then mstrPigeonhole = Trim(Nvl(rsTmp("归档人"))): mstrReportMan = Trim(Nvl(rsTmp("报告人")))
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim lngQC As Long
    Dim rsTmp As New ADODB.Recordset
    
    If Me.Visible = False Then Exit Sub
    
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                  '打印设置
            Call zlPrintSet
        Case conMenu_File_Print                                     '打印控制图
            Call mfrmChartLJ.ChartPrint: Call PrintQC_LJ(True)
        Case conMenu_Edit_Leave_Post                                '另存控制图
            Call mfrmChartLJ.ChartSaveAs
        Case conMenu_Edit_MarkMap                                   '复制控制图
            Call mfrmChartLJ.ChartCopy
        Case conMenu_File_Exit                                      '退出
            Unload Me
        Case conMenu_View_ToolBar_Button                            '标准按钮
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text                              '文本标签
             For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size                              '大图标
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
            
        Case conMenu_Edit_Save                                      '保存
            mlngResult = mfrmQCTodayReport.zlEditSave()
            If mlngResult <> 0 Then
                mfrmQCTodayReport.zlRefresh mlngResult
                mEditMode = 0
            End If
            
        Case conMenu_Edit_Untread                                   '取消
            mfrmQCTodayReport.zlEditCancel
            mEditMode = 0
        
        Case conMenu_Edit_Adjust                                    '报告
            Call mfrmQCTodayReport.ZlEditStart(mlngResult)
            mEditMode = 1
        
        Case conMenu_Edit_Archive                                   '归档
            gstrSql = "select 归档人 from 检验质控报告 where 结果id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResult)
            If rsTmp.EOF = False Then
                If Nvl(rsTmp("归档人")) = "" Then
                    If MsgBox("真的要将当前失控报告归档吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                    gstrSql = "Zl_检验质控报告_Archive(" & mlngResult & ",0)"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    mstrPigeonhole = gstrDBUser
                Else
                    If MsgBox("该失控报告已经归档，真的取消归档吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                    gstrSql = "Zl_检验质控报告_Archive(" & mlngResult & ",1)"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    mstrPigeonhole = ""
                End If
            End If
            Call mfrmQCTodayReport.zlRefresh(mlngResult)
            
        Case conMenu_View_Refresh                                   '刷新
            Call cbo项目_Click
        Case conMenu_Tool_Analyse                                   '失控计算
            If InStr(mstrQCID, ",") > 0 Then
                lngQC = Mid(mstrQCID, 1, InStr(mstrQCID, ",") - 1)
            Else
                lngQC = mstrQCID
            End If
            frmQCCompute.ShowMe Me, mlngMachineID, cbo项目.ItemData(cbo项目.ListIndex), zlDatabase.Currentdate, lngQC
        Case conMenu_Tool_Define                                    '重新定值
            If InStr(mstrQCID, ",") > 0 Then
                lngQC = Mid(mstrQCID, 1, InStr(mstrQCID, ",") - 1)
            Else
                lngQC = mstrQCID
            End If
            frmQCRedefine.ShowMe Me, mlngMachineID, cbo项目.ItemData(cbo项目.ListIndex), zlDatabase.Currentdate, lngQC
        Case conMenu_Help_Web                                       'WEB上的中联
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home                                  '主页
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Mail                                  '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About                                     '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    If Me.Visible = True Then
        Me.dkpMan.RecalcLayout
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_Edit_Adjust
            Control.Enabled = (mEditMode = 0 And mstrPigeonhole = "")
        Case conMenu_Edit_Archive
            Control.Enabled = (mEditMode = 0 And mstrReportMan <> "")
        Case conMenu_Edit_Save, conMenu_Edit_Untread
            Control.Enabled = (mEditMode = 1)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmChartLJ.hwnd
    Case 2
        Item.Handle = mfrmQCTodayReport.hwnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Top = 500
End Sub

Private Sub dkpMan_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarControl, cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印控制图(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "另存控制图(&S)..."): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "复制控制图(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "报告(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档(&T)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值(&N)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "项目")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "项目")
    cbrCustom.Handle = Me.cbo项目.hwnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_Edit_Untread
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
    End With

    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        'conMenu_Edit_Save
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "另存为"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "复制"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置停靠窗格
    Dim panThis As Pane, panChild As Pane

    With Me.dkpMan
        Set panThis = .CreatePane(1, 700, 1000, DockBottomOf, Nothing)
        panThis.Title = "质控图形"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(2, 300, 1000, DockRightOf, panThis)
        panChild.Title = "失控记录"
        panChild.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End With

    Set mfrmChartLJ = New frmQCChartLJ
    Set mfrmQCTodayReport = New frmQCTodayReport
    
    '界面恢复
'    Call RestoreWinState(Me, App.ProductName)

    '得到所做的项目

    gstrSql = "Select Distinct B.ID, B.编码, B.中文名, B.英文名 " & vbNewLine & _
                " From 检验普通结果 A, 诊治所见项目 B, 检验质控品项目 C  " & vbNewLine & _
                " Where A.检验项目id = B.ID And A.检验项目id = C.项目id And a.检验标本ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
    With rsTmp
        Me.cbo项目.Clear
        Do While Not .EOF
            Me.cbo项目.AddItem !编码 & ", " & !中文名 & "/" & !英文名
            Me.cbo项目.ItemData(Me.cbo项目.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo项目.ListCount = 0 Then MsgBox "尚未完成仪器质控品设置！", vbInformation, gstrSysName
        If Me.cbo项目.ListCount > 0 Then
            Me.cbo项目.ListIndex = 0
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub ShowMe(lngSampleID As Long, objfrm As Object, lngMachineID As Long)
    mlngSampleID = lngSampleID
    mlngMachineID = lngMachineID
    Me.Show , objfrm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrQCID = ""
    mlngSampleID = 0
    Unload mfrmChartLJ
    Set mfrmChartLJ = Nothing
End Sub
Private Function getMonthFirst(dtNow As Date) As String
    '功能         得到本月的第一天
    '参数         dtNow 传入日期
    
    getMonthFirst = Format(dtNow, "YYYY-MM")
    getMonthFirst = getMonthFirst & "-01"
    
End Function
Private Function getMonthLast(dtNow As Date) As String
    '功能         得到本月的最后一天
    '参数         dtNow 传入日期
    Dim strYear As String
    Dim strMonth As String
    strYear = Format(dtNow, "YYYY")
    strMonth = Format(dtNow, "MM")
    If CInt(strMonth) = 12 Then strMonth = "00": strYear = CInt(strYear) + 1
    getMonthLast = Format(CDate(strYear & "-" & CInt(strMonth) + 1 & "-01") - 1, "yyyy-mm-dd")
    
End Function

Private Sub PrintQC_LJ(blnPrintMode As Boolean)
    '打印或预览LJ质控图
    '参数           intPrintMode =1 打印 =2 预览
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '对应的单据
    Dim strQCID As String                       '质控品ID可能会是以","分隔的多个ID
    Dim lngQCID As Long                         '单个质控品ID
    Dim lngItemID As String                     '项目ID
    Dim lngMachine As Long                      '仪器ID
    Dim intloop As Integer                      '循环字串
    Dim intReportCount As Integer               '要打印的图像数
    Dim intPrintType As Integer
    
    intReportCount = mfrmChartLJ.ChartPrint
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1209_1"
    
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.报表id And a.编号 = [1] And b.名称 = '质控图'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    '没有找到时退出
    If rsTmp.EOF Then
        MsgBox "在单据定义中没有定义<质控图>,请在单据中定义一个名为<质控图>的图像框!", vbQuestion, Me.Caption
        Exit Sub
    End If
    
    For intloop = 0 To intReportCount - 1
        With Me.chtCopy
            .Load App.path & "\QC_Tmp" & intloop
            Kill App.path & "\QC_Tmp" & intloop
            .Width = Nvl(rsTmp("w"), 1280 * Screen.TwipsPerPixelX)
            .Height = Nvl(rsTmp("h"), 500 * Screen.TwipsPerPixelY)
            .Header.Text = ""
            .ChartLabels.RemoveAll
            .ChartArea.Location.Top = -5
            .ChartArea.Location.Height = .ChartArea.Location.Height + 15
            If intPrintType = 3 Then
                .ChartArea.Location.Left = 30
            End If
            .SaveImageAsJpeg App.path & "\QC" & intloop & ".jpg", 1000, False, False, False
        End With
    Next
    
    '得到质控品ID
    lngQCID = mfrmChartLJ.ZLGetLJ_QCID
    strQCID = mfrmChartLJ.ZLGetLJ_QCIDStr
    
    
    '得到项目ID
    If Me.cbo项目.ListCount = 0 Then Exit Sub
    lngItemID = CLng(Me.cbo项目.ItemData(Me.cbo项目.ListIndex))
    lngMachine = mlngMachineID
    
    If Dir(App.path & "\QC0.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, "质控图=" & App.path & "\QC0.jpg", _
        "质控品ID=" & lngQCID, "项目ID=" & lngItemID, "开始日期=" & mstrStart, "结束日期=" & mstrEnd, _
        "仪器ID=" & lngMachine, "质控品组=" & IIf(strQCID = "", "0", strQCID), _
        "质控图1=" & App.path & "\QC1.jpg", "质控图2=" & App.path & "\QC2.jpg", _
        IIf(blnPrintMode, 2, 1))
    End If
    
    Kill App.path & "\QC*.jpg"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

