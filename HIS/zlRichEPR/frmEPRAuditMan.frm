VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CO1FBF~1.OCX"
Begin VB.Form frmEPRAuditMan 
   Caption         =   "病历质量审查"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmEPRAuditMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picKind 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   285
      ScaleHeight     =   1020
      ScaleWidth      =   2325
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   765
      Width           =   2325
      Begin VB.OptionButton optKind 
         Caption         =   "门诊病历(&1)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   75
         Width           =   1380
      End
      Begin VB.OptionButton optKind 
         Caption         =   "住院病历(&2)"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optKind 
         Caption         =   "护理病历(&3)"
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   3
         Top             =   720
         Width           =   1380
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   285
      ScaleHeight     =   1680
      ScaleWidth      =   2325
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1935
      Width           =   2325
      Begin VB.CheckBox chkNoData 
         Caption         =   "显示无业务科室(&N)"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1425
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "重新统计(&R)"
         Height          =   350
         Left            =   450
         TabIndex        =   6
         Top             =   900
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   300
         Left            =   450
         TabIndex        =   5
         Top             =   465
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   102891523
         CurrentDate     =   38683
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   300
         Left            =   450
         TabIndex        =   4
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   102891523
         CurrentDate     =   38683
      End
      Begin VB.Label lblDateTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   525
         Width           =   180
      End
      Begin VB.Label lblDateFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   180
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6450
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditMan.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
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
   Begin XtremeSuiteControls.TaskPanel tplThis 
      Height          =   5670
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   690
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   10001
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   5310
      Left            =   3090
      TabIndex        =   0
      Top             =   1050
      Width           =   6405
      _cx             =   11298
      _cy             =   9366
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaper       =   "frmEPRAuditMan.frx":0E1C
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊病历(2005-11-20至2005-11-26)书写情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3090
      TabIndex        =   14
      Top             =   780
      Width           =   4065
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2325
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRAuditMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim lngCount As Long, lngRow As Long, lngCol As Long

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngDeptId As Long, strDeptName As String
    With Me.vfgThis
        lngDeptId = Val(.TextMatrix(.Row, 0))
        strDeptName = .TextMatrix(.Row, 2)
    End With
    
    Select Case Control.ID
    Case conMenu_File_Open:
        Dim cbrPBar As CommandBar
        Dim cbrPItem As CommandBarControl
        
        Set cbrPBar = Me.cbsThis.Add("弹出", xtpBarPopup)
        With Me.vfgThis
            Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10, "病历文件分类审查(&F)")
            If mintKind = 1 Then
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 1, "门诊病人病历审查(&1)")
                cbrPItem.BeginGroup = True
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 5)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 2, "急诊病人病历审查(&2)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 7)) <> 0)
            Else
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 1, "入院病人病历审查(&1)")
                cbrPItem.BeginGroup = True
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 5)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 2, "转入病人病历审查(&2)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 7)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 3, "出院病人病历审查(&3)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 9)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 4, "死亡病人病历审查(&4)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 11)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 5, "转出病人病历审查(&5)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 13)) <> 0)
                If mintKind = 2 Then
                    Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 6, "手术病历审查(&6)")
                    cbrPItem.Enabled = (Val(.TextMatrix(.Row, 15)) <> 0)
                End If
            End If
            Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 9, "全体病人病历审查(&A)")
            cbrPItem.BeginGroup = True
        End With
        cbrPBar.ShowPopup
    Case conMenu_File_Open * 10: Call frmEPRAuditFile.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo)
    Case conMenu_File_Open * 10 + 1: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, IIf(mintKind = 1, "门诊", "入院"))
    Case conMenu_File_Open * 10 + 2: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, IIf(mintKind = 1, "急诊", "转入"))
    Case conMenu_File_Open * 10 + 3: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "出院")
    Case conMenu_File_Open * 10 + 4: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "死亡")
    Case conMenu_File_Open * 10 + 5: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "转出")
    Case conMenu_File_Open * 10 + 6: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "手术")
    Case conMenu_File_Open * 10 + 9: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo)
    
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call RefreshData
    Case conMenu_View_Jump
        If Me.Visible Then Me.vfgThis.SetFocus
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        '执行发布到当前模块的报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                "开始日期=" & Format(dtpDateFrom.Value, "yyyy-MM-dd"), "结束日期=" & Format(dtpDateTo.Value, "yyyy-MM-dd"), _
                "病历种类=" & IIf(optKind(0).Value, "门诊病历", IIf(optKind(1).Value, "住院病历", "护理病历")) & "|" & IIf(optKind(0).Value, 1, IIf(optKind(1).Value, 2, 4)))
        End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop  As Long, lngScaleRight  As Long, lngScaleBottom  As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    With Me.tplThis
        .Left = lngScaleLeft
        .Top = lngScaleTop: .Height = lngScaleBottom - .Top
    End With
    With Me.lblTitle
        .Left = Me.tplThis.Left + Me.tplThis.Width + 30: .Width = lngScaleRight - .Left
        .Top = lngScaleTop + 60
    End With
    With Me.vfgThis
        .Left = Me.tplThis.Left + Me.tplThis.Width: .Width = lngScaleRight - .Left
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 60: .Height = lngScaleBottom - .Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
        With Me.vfgThis
            Control.Enabled = (Val(.TextMatrix(.Row, 1)) <> 0)
            If Control.Enabled = False Then Exit Sub
            For lngCol = 3 To .Cols - 1
                Control.Enabled = (Val(.TextMatrix(.Row, lngCol)) <> 0)
                If Control.Enabled Then Exit Sub
            Next
        End With
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfgThis.Rows > Me.vfgThis.FixedRows + 1)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkNoData_Click()
    Dim blnData As Boolean
    With Me.vfgThis
        If Me.chkNoData.Value = vbChecked Then
            For lngRow = .FixedRows To .Rows - 2
                .ROWHEIGHT(lngRow) = .RowHeightMin
                .RowHidden(lngRow) = False
            Next
        Else
            For lngRow = .FixedRows To .Rows - 2
                blnData = False
                For lngCol = 3 To .Cols - 1
                    If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                Next
                If blnData = False Then
                    .ROWHEIGHT(lngRow) = 0
                    .RowHidden(lngRow) = True
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdSearch_Click()
    If Me.dtpDateTo.Value - Me.dtpDateFrom.Value > 15 Then MsgBox "审查时间范围太长(不能超过15天)！", vbExclamation, gstrSysName: Exit Sub
    
    If Me.optKind(0).Value Then
        mintKind = 1
    ElseIf Me.optKind(1).Value Then
        mintKind = 2
    ElseIf Me.optKind(2).Value Then
        mintKind = 4
    Else
        Me.optKind(1).Value = True: mintKind = 2
    End If
    mstrDateFrom = Format(Me.dtpDateFrom.Value, "yyyy-mm-dd")
    mstrDateTo = Format(Me.dtpDateTo.Value, "yyyy-mm-dd")
    
    Call RefreshData
End Sub

Private Sub dtpDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_Validate(Cancel As Boolean)
    Me.dtpDateFrom.MaxDate = Me.dtpDateTo.Value
    If Me.dtpDateFrom.Value > Me.dtpDateFrom.MaxDate Then Me.dtpDateFrom.Value = Me.dtpDateFrom.MaxDate
End Sub

Private Sub Form_Load()
    Call zlcommfun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    '-----------------------------------------------------
    '初始条件
    Call InitTerm
    Call cmdSearch_Click
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
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
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "展开(&O)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "跳转到表格(&J)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "展开"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示无业务科室", IIf(Me.chkNoData.Value = vbChecked, 1, 0))
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optKind_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_File_Open)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub InitTerm()
    '-------------------------------------------------
    '--功能：初始化条件和布局
    '-------------------------------------------------
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    '-----------------------------------------------------
    '初始数据:
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示无业务科室", 1) = 1 Then
        Me.chkNoData.Value = vbChecked
    Else
        Me.chkNoData.Value = vbUnchecked
    End If
    
    strSQL = "Select Sysdate From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    
    '-----------------------------------------------------
    '显示形态
    Set tplGroup = Me.tplThis.Groups.Add(0, "审查范围:"): tplGroup.Expandable = False
    Set tplItem = tplGroup.Items.Add(0, "病历种类:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picKind
    Me.picKind.BackColor = tplItem.BackColor
    For lngCount = 0 To Me.optKind.Count - 1: Me.optKind(lngCount).BackColor = tplItem.BackColor: Next
    Set tplItem = tplGroup.Items.Add(0, "书写日期范围:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picDate
    Me.picDate.BackColor = tplItem.BackColor
    Me.chkNoData.BackColor = tplItem.BackColor
    
    Set tplGroup = Me.tplThis.Groups.Add(0, "查阅说明:"): tplGroup.Expandable = True
    Set tplItem = tplGroup.Items.Add(0, "  ①、对应各事件的完成病历数，仅包含要求书写一次的病历，不包含要求循环书写的病历；", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "  ②、由于某些事件对应要求书写一次的病历为多种，因此其完成病历数可能超过发生人次数；", xtpTaskItemTypeText)

    '-----------------------------------------------------
    Me.tplThis.Reposition
    Me.BackColor = tplItem.BackColor
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshData()
    '-------------------------------------------------
    '功能:根据审查范围组织显示审查数据
    '-------------------------------------------------
    Dim lngTotal As Long
    
    Select Case mintKind
    Case 1  '门诊病历
        Me.lblTitle.Caption = "门诊病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, P.门诊人次, W.门诊完成, P.急诊人次, W.急诊完成" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select 执行部门id, Sum(Decode(急诊, 1, 0, 1)) As 门诊人次, Sum(Decode(急诊, 1, 1, 0)) As 急诊人次" & vbNewLine & _
                "        From 病人挂号记录" & vbNewLine & _
                "        Where Nvl(执行状态, 0) <> 0 And 登记时间 Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "              To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 执行部门id) P," & vbNewLine & _
                "      (Select W.科室id, Sum(W.已完成) As 已完成, Sum(W.在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.事件, '门诊', W.已完成, Null)) As 门诊完成," & vbNewLine & _
                "               Sum(Decode(F.事件, '急诊', W.已完成, Null)) As 急诊完成" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 1) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 1 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "                     To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '临床' And M.服务对象 In (1, 3) And D.ID = P.执行部门id(+) And" & vbNewLine & _
                "       D.ID = W.科室id(+)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "科室": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "门诊": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "急诊": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
        End With
    
    Case 2  '住院病历
        Me.lblTitle.Caption = "住院病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, I.入院人次, W.入院病历, E.转入人次, W.转入病历, O.出院人次," & vbNewLine & _
                "        W.出院病历, O.死亡人次, W.死亡病历, G.转出人次, W.转出病历, S.手术人次, W.手术病历" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select W.科室id, Sum(已完成) As 已完成, Sum(在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '入院', 1, '首次入院', 1, '再次入院', 1, 0), 0) * 已完成) As 入院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 0, 1), 0), 0) * 已完成) As 转入病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '出院', 1, '24小时出院', 1, 0), 0) * 已完成) As 出院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '死亡', 1, '24小时死亡', 1, 0), 0) * 已完成) As 死亡病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 1, 0), 0), 0) * 已完成) As 转出病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '手术', 1, 0), 0) * 已完成) As 手术病历" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 2) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 2 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W," & vbNewLine
        strSQL = strSQL & "      (Select 入院科室id, Count(*) As 入院人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 入院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 入院科室id) I," & vbNewLine & _
                "      (Select 科室id, Count(*) As 转入人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 开始原因 = 3 And 开始时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 科室id) E," & vbNewLine & _
                "      (Select 出院科室id, Sum(Decode(出院方式, '死亡', 0, 1)) As 出院人次, Sum(Decode(出院方式, '死亡', 1, 0)) As 死亡人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 出院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 出院科室id) O," & vbNewLine & _
                "      (Select 科室id, Count(*) As 转出人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 终止原因 = 3 And 终止时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 科室id) G," & vbNewLine & _
                "      (Select R.病人科室id, Count(*) As 手术人次" & vbNewLine & _
                "        From 病人医嘱记录 R, 病人医嘱发送 S" & vbNewLine & _
                "        Where R.ID = S.医嘱id And R.诊疗类别 = 'F' And S.首次时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By R.病人科室id) S" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '临床' And 服务对象 In (2, 3) And D.ID = W.科室id(+) And D.ID = I.入院科室id(+) And" & vbNewLine & _
                "       D.ID = E.科室id(+) And D.ID = O.出院科室id(+) And D.ID = G.科室id(+) And D.ID = S.病人科室id(+)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "科室": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "入院": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "转入": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "出院": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "死亡": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "转出": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            .TextMatrix(0, 15) = "手术": .TextMatrix(0, 16) = .TextMatrix(0, 15)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
            .TextMatrix(1, 9) = "人次": .TextMatrix(1, 10) = "完成病历"
            .TextMatrix(1, 11) = "人次": .TextMatrix(1, 12) = "完成病历"
            .TextMatrix(1, 13) = "人次": .TextMatrix(1, 14) = "完成病历"
            .TextMatrix(1, 15) = "人次": .TextMatrix(1, 16) = "完成病历"
        End With
    Case 4  '护理病历
        Me.lblTitle.Caption = "护理病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, I.入院人次, W.入院病历, E.转入人次, W.转入病历, O.出院人次," & vbNewLine & _
                "        W.出院病历, O.死亡人次, W.死亡病历, G.转出人次, W.转出病历" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select W.科室id, Sum(已完成) As 已完成, Sum(在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '入院', 1, '首次入院', 1, '再次入院', 1, 0), 0) * 已完成) As 入院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 0, 1), 0), 0) * 已完成) As 转入病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '出院', 1, '24小时出院', 1, 0), 0) * 已完成) As 出院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '死亡', 1, '24小时死亡', 1, 0), 0) * 已完成) As 死亡病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 1, 0), 0), 0) * 已完成) As 转出病历" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 4) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 4 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W," & vbNewLine
        strSQL = strSQL & "      (Select 入院病区id, Count(*) As 入院人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 入院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 入院病区id) I," & vbNewLine & _
                "      (Select 病区id, Count(*) As 转入人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 开始原因 = 3 And 开始时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 病区id) E," & vbNewLine & _
                "      (Select 当前病区id, Sum(Decode(出院方式, '死亡', 0, 1)) As 出院人次, Sum(Decode(出院方式, '死亡', 1, 0)) As 死亡人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 出院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 当前病区id) O," & vbNewLine & _
                "      (Select 病区id, Count(*) As 转出人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 终止原因 = 3 And 终止时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 病区id) G" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '护理' And 服务对象 In (2, 3) And D.ID = W.科室id(+) And D.ID = I.入院病区id(+) And" & vbNewLine & _
                "       D.ID = E.病区id(+) And D.ID = O.当前病区id(+) And D.ID = G.病区id(+)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "病区": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "入院": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "转入": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "出院": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "死亡": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "转出": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
            .TextMatrix(1, 9) = "人次": .TextMatrix(1, 10) = "完成病历"
            .TextMatrix(1, 11) = "人次": .TextMatrix(1, 12) = "完成病历"
            .TextMatrix(1, 13) = "人次": .TextMatrix(1, 14) = "完成病历"
        End With
    End Select
    
    '求合计
    With Me.vfgThis
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "合计"
        For lngCol = 3 To .Cols - 1
            lngTotal = 0
            For lngRow = .FixedRows To .Rows - 2
                lngTotal = lngTotal + Val(.TextMatrix(lngRow, lngCol))
            Next
            .TextMatrix(.Rows - 1, lngCol) = lngTotal
        Next
        .Row = .FixedRows: .Col = 1
        Call .AutoSize(1, .Cols - 1)
    End With
    
    '显示或隐藏空行
    Call chkNoData_Click
    Me.stbThis.Panels(2).Text = "点击“展开(Ctrl+O)”详细审查当前科室病人病历情况或病历分类书写情况…"
    
    If Me.Visible Then Me.vfgThis.SetFocus
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = Me.lblTitle.Caption
    Set objPrint.Title.Font = Me.lblTitle.Font
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

