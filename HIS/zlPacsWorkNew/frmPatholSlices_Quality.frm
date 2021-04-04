VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmPatholSlices_Quality 
   Caption         =   "玻片质量"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frmPatholSlices_Quality.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11280
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtSlideNum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2670
      TabIndex        =   3
      Top             =   6765
      Width           =   1260
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   420
      ScaleHeight     =   6180
      ScaleWidth      =   9735
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   345
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5145
         Left            =   375
         TabIndex        =   1
         Top             =   825
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   9075
         DefaultCols     =   ""
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontSize    =   10.5
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontSize    =   10.5
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7515
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSlices_Quality.frx":038A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6165
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSlices_Quality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long       '医嘱ID

Private mrecStudyInf As TStudyStateInf
Private mblnCurModifyState As Boolean
Private mblnAllowEditState As Boolean

Private Enum TMenuType
    mtFile = 1          '文件
      mtSave = 2        '保存
      mtCancel = 3      '撤销
      mtQuit = 4        '退出
    
    mtEdit = 5          '编辑
      mtModify = 6      '
      
    mtApplyAll = 7      '应用到所有
    mtClear = 8         '清除
    
    mtFind = 9          '查找
    mtPlace = 10        '占位
End Enum


Public Sub ShowSlideEvaluateWindow(ByVal lngAdviceID As Long, ByVal lngStudyStep As Long, _
    ByVal strPrivs As String, owner As Object)

    '测试方法
'    InitDebugObject 1290, Me, "zlhis", "HIS"
    
        '根据医嘱ID获取病理号等相关状态信息
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
    mblnAllowEditState = CheckPopedom(strPrivs, "质量管理") And lngStudyStep < 6
    
    Call Me.Show(1, owner)
End Sub

Private Sub InitQualityList()
'初始化制片质量列表
   
    ufgData.IsEjectConfig = False
    ufgData.IsShowPopupMenu = False
    ufgData.IsKeepRows = False
    ufgData.IsCopyMode = True
    
    ufgData.RowHeightMin = 315
    ufgData.ColNames = gstrSlicesQualityCols
    
    
    ufgData.ColConvertFormat = gstrSlicesQualityConvertFormat
    ufgData.DataGrid.ExtendLastCol = True
    
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'执行界面功能
Dim strResult As String

On Error GoTo errHandle
    Select Case control.ID
    
        Case TMenuType.mtCancel
            Call CancelModify       '撤销修改
            
'        Case TMenuType.mtModify
'            Call ModifyEvaluate     '修改评价
            
        Case TMenuType.mtClear
            Call ClearEvaluate      '清除评价
            
        Case TMenuType.mtApplyAll
            Call ApplyAll           '应用到所有
            
        Case TMenuType.mtSave
            Call SaveEvaluate       '保存质量评价
            
        Case TMenuType.mtFind
            Call FindData           '查找数据
                        
        
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
            
        Case TMenuType.mtQuit   '退出
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub FindData()
'查找玻片数据
    Dim lngFindRowIndex As Long
    
    If Trim(txtSlideNum.Text) = "" Then Exit Sub
    
    lngFindRowIndex = ufgData.FindRowIndex(Trim(txtSlideNum.Text), "条码号")
    
    If lngFindRowIndex >= 1 Then
        Call ufgData.LocateRow(lngFindRowIndex)
        
        If ufgData.Text(lngFindRowIndex, "玻片质量") = "" And mblnAllowEditState Then
            ufgData.DataGrid.TextMatrix(lngFindRowIndex, ufgData.GetColIndex("玻片质量")) = "甲"
            
            mblnCurModifyState = True
        End If
        
'        Call ufgData.EditNextCellWithCurRow(False)
    End If
    
    Call zlControl.TxtSelAll(txtSlideNum)
End Sub


Private Sub CancelModify()
'撤销修改
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    
    Call ufgData.RestoreList(False)
    Call ufgData.RefreshReadColColor
    
    mblnCurModifyState = False
    'Call ConfigFaceEditState(False)
End Sub

'Private Sub ModifyEvaluate()
''修改评价
'    Call ConfigFaceEditState(True)
'End Sub

Private Sub ClearEvaluate()
'清除评价
    Dim i As Long
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    For i = 1 To ufgData.GridRows - 1
        '这里不能使用ufgdata的text属性赋值，因为该属性会更新行的flexcpData属性，使其取消时不能恢复数据
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("玻片质量")) = ""
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("评审人")) = ""
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("评审日期")) = ""
    Next i
    
    mblnCurModifyState = True
End Sub

Private Sub SaveEvaluate()
'保存评价
    Dim i As Long
    Dim strSql As String
    Dim strQuality As String
    Dim dtCurDate As Date
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    dtCurDate = zlDatabase.Currentdate
    
    '循环更新评审质量数据
    For i = 1 To ufgData.GridRows - 1
        strQuality = Trim(ufgData.Text(i, "玻片质量"))
        
        strSql = "Zl_病理玻片信息_质量评价(" & CLng(Val(ufgData.KeyValue(i))) & _
                                            ",'" & strQuality & _
                                            "','" & UserInfo.姓名 & "')"
                                       
        zlDatabase.ExecuteProcedure strSql, "玻片质量保存"

        
        ufgData.Text(i, "玻片质量") = strQuality '更新flexcpdata数据，以便进行撤销恢复
        ufgData.Text(i, "评审人") = IIf(strQuality = "", "", UserInfo.姓名)
        ufgData.Text(i, "评审日期") = IIf(strQuality = "", "", Format(dtCurDate, "yyyy-mm-dd"))
    Next i
    
    mblnCurModifyState = False
'    Call ConfigFaceEditState(False)
End Sub

Private Sub ApplyAll()
'应用到所有
    Dim strCurValue As String
    Dim i As Long
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    strCurValue = ufgData.CurText("玻片质量")
    
    If strCurValue = "" Then
        Call MsgBoxD(Me, "当前记录未设置玻片质量，不能应用到其他数据中。", vbOKOnly)
        Exit Sub
    End If
    
    For i = 1 To ufgData.GridRows - 1
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("玻片质量")) = strCurValue
    Next i
    
    mblnCurModifyState = True
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'功能：调用帮助主题
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picBack.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top - IIf(stbThis.Visible, stbThis.Height, 0)
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单和按钮显示
On Error Resume Next
    
    Select Case control.ID

        Case TMenuType.mtSave, TMenuType.mtCancel ', TMenuType.mtClear, mtApplyAll
            control.Enabled = mblnCurModifyState And mblnAllowEditState


'        Case TMenuType.mtModify
'            control.Enabled = Not mblnCurModifyState And ufgData.GridRows > 1

        Case TMenuType.mtApplyAll, TMenuType.mtClear
            control.Enabled = ufgData.GridRows > 1 And mblnAllowEditState
    End Select
    
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnCurModifyState = False
    
    Call InitCommandBars
    
    '初始化制片显示列表
    Call InitQualityList
    
    Call LoadSlideData(mrecStudyInf.lngPatholAdviceId)

    '如果当前不允许修改，则将界面修改为只读查看状态
    If Not mblnAllowEditState Then
        Call ConfigFaceEditState(False)
    End If
    
    stbThis.Panels(3).Text = "评审人：" & UserInfo.姓名
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                             '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                     '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "文件(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "撤消(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "编辑(&E)")
    
    'Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtModify, "评价(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtApplyAll, "应用到所有(&A)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtClear, "清除评价(&C)"): cbrControl.IconId = 4008: cbrControl.BeginGroup = True
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------工具栏定义------------------------------------------
        
            
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
        
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存", "保存评价"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "撤消", "撤消修改"): cbrControl.IconId = 3565
    
    'Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtModify, "评价", "玻片评价"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtClear, "清除", "清除评价"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtApplyAll, "应用到所有", "将所有评价内容置为相同"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出", "退出"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '设置右上角定位输入
    Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlLabel, TMenuType.mtPlace, "条码号：")
        cbrControl.ID = TMenuType.mtPlace
        cbrControl.flags = xtpFlagRightAlign
        cbrControl.IconId = 1
        
    Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, TMenuType.mtPlace, "条码号")
        cbrCustom.Handle = txtSlideNum.hWnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Style = xtpButtonIconAndCaption
        
    Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, TMenuType.mtFind, " 定 位(&L) ")
        cbrControl.ID = TMenuType.mtFind
        cbrControl.flags = xtpFlagRightAlign
End Sub


Private Sub ConfigFaceEditState(ByVal blnIsEdit As Boolean)
    Dim lngColIndex As Long
    
    lngColIndex = ufgData.GetColIndex("玻片质量")
    mblnCurModifyState = blnIsEdit
    
    ufgData.ReadOnly = Not blnIsEdit
    
    If ufgData.GridRows <= 1 Then Exit Sub
    
    If blnIsEdit Then
        ufgData.DataGrid.Cell(flexcpBackColor, 1, ufgData.GetColIndex("玻片质量"), ufgData.DataGrid.Rows - 1, ufgData.GetColIndex("玻片质量")) = &H80000005
    Else
        ufgData.DataGrid.Cell(flexcpBackColor, 1, ufgData.GetColIndex("玻片质量"), ufgData.DataGrid.Rows - 1, ufgData.GetColIndex("玻片质量")) = &H8000000F
    End If

End Sub


Private Sub LoadSlideData(ByVal lngPatholAdviceId As Long)
On Error GoTo errHandle
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    '查询来源于制片的已完成玻片信息
    strSql = "select a.ID,a.来源类型, a.来源Id, a.材块ID,a.条码号,a.玻片质量,a.评审人,a.评审日期, d.标本名称,c.取材位置, c.序号 as 材块号," & _
                     "decode(b.制片类型,0,'常规',1,'冰冻',2,'细胞',3,'重切',4,'深切',5,'连切','') as 玻片类型 " & _
                     "from  病理玻片信息 a,病理制片信息 b, 病理取材信息 c, 病理标本信息 d " & _
                     "where a.来源id = b.id and b.材块ID = c.材块id and c.标本Id=d.标本Id and a.来源类型=0 and b.当前状态=2 and a.病理医嘱ID =[1]"
    
    strSql = strSql & vbCrLf & " union all " & vbCrLf
    
    '查询来源于特检的已完成玻片信息
    strSql = "select * from (" & strSql & " select a.ID,a.来源类型, a.来源Id, a.材块ID,a.条码号,a.玻片质量,a.评审人,a.评审日期, d.标本名称,c.取材位置,c.序号 as 材块号," & _
                     "decode(b.特检类型,0,'免疫',1,'特染',2,'分子','') as 玻片类型 " & _
                     "from  病理玻片信息 a,病理特检信息 b, 病理取材信息 c, 病理标本信息 d " & _
                     "where a.来源id = b.id and b.材块ID = c.材块id and c.标本Id=d.标本Id and a.来源类型=1 and b.当前状态=2 and a.病理医嘱ID =[1] ) order by 条码号 "
    
                     
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "获取玻片信息", lngPatholAdviceId)
    
    Call ufgData.ClearListData
    If rsData.RecordCount < 1 Then Exit Sub
    
    Set ufgData.AdoData = rsData
    Call ufgData.RefreshData
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)
errHandle:
End Sub

Private Sub picBack_Resize()
On Error Resume Next
   
    ufgData.Left = 40
    ufgData.Top = 0
    ufgData.Width = picBack.ScaleWidth - 80
    ufgData.Height = picBack.ScaleHeight - 20
End Sub


Private Sub txtSlideNum_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii = 13 Then
        Call FindData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnChangeEdit()
    mblnCurModifyState = True
End Sub

Private Sub ufgData_OnDblClick()
On Error GoTo errHandle
    Call ufgData.EditNextCellWithCurRow(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

