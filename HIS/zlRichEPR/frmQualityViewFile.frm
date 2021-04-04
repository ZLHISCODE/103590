VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmQualityViewFile 
   Caption         =   "科室病历完成情况"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10815
   Icon            =   "frmQualityViewFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10815
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3780
      Left            =   75
      TabIndex        =   0
      Top             =   660
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   6667
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityViewFile.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   585
      Top             =   4455
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":70E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":7C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":81B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":874C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewFile.frx":8CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1395
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2175
      Left            =   4230
      TabIndex        =   3
      Top             =   675
      Width           =   3900
      _cx             =   6879
      _cy             =   3836
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
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmQualityViewFile.frx":9280
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   8325
      Picture         =   "frmQualityViewFile.frx":9355
      Top             =   3690
      Visible         =   0   'False
      Width           =   2265
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2160
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQualityViewFile.frx":1A41F
      Left            =   900
      Top             =   225
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQualityViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: ID: 种类: 编号: 名称: 完成情况
End Enum

Private Enum mCol2
    ID = 0: 病人ID: 主页ID: 住院号: 床号: 姓名: 性别: 年龄: 住院医师: 费别: 入院日期: 出院日期: 险类: 创建人: 创建时间: 完成时间: 保存人: 保存时间: 最后版本: 签名级别: 归档人: 归档日期
End Enum

Private Enum mViewModeEnum
    所有病历 = 0: 正在书写病历: 已完成病历
End Enum

Private Enum Enum病历种类
    门诊病历 = 1
    住院病历 = 2
    护理病历 = 4
End Enum
Private mvar病历种类 As Enum病历种类

Const conPane_FileTab = 201
Const conPane_FileList = 202
Const conPane_Content = 203
Const conViewAll = 301
Const conViewInEditing = 302
Const conViewFinished = 303

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String         '当前使用者权限串
Private mstrKinds As String         '当前允许定义的病历类型串

Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1

Private mlngCurFileId As Long       '当前文件ID
Private mlngDeptId As Long          '科室ID
Private mstrDeptName As String      '科室名称
Private mstrFrom As String, mstrTo As String
Private mViewMode As mViewModeEnum  '视图模式   0-正在书写病历    1-已完成病历

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long, lngCurRow As Long

Public Sub ShowMe(ByVal 病历种类 As Long, _
    ByRef frmParent As Object, _
    ByVal lngDeptId As Long, _
    ByVal strDeptName As String, _
    ByVal strFrom As String, _
    ByVal strTo As String)
    '显示本窗体
    mvar病历种类 = 病历种类
    mlngDeptId = lngDeptId
    mstrDeptName = strDeptName
    mstrFrom = strFrom
    mstrTo = strTo
    Me.Caption = "科室病历完成情况 - [" & strDeptName & "]"
    Me.Show vbModeless, frmParent
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '       strSubhead，打印的副标题
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = Me.Caption
    objPrint.Title.Font.Name = "黑体"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("科室:" & mstrDeptName)
    Call objAppRow.Add("文件名:" & Me.rptList.FocusedRow.Record.Item(mCol.名称).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFileID As Long
    On Error GoTo LL
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
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
    Case conMenu_View_Refresh
        Call zlRefList(mlngCurFileId)
    Case conViewAll
        mViewMode = 所有病历
        Call rptList_SelectionChanged
    Case conViewInEditing
        mViewMode = 正在书写病历
        Call rptList_SelectionChanged
    Case conViewFinished
        mViewMode = 已完成病历
        Call rptList_SelectionChanged
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
LL:
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_ExportToXML
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conViewAll: Control.Checked = (mViewMode = 所有病历)
    Case conViewInEditing: Control.Checked = (mViewMode = 正在书写病历)
    Case conViewFinished: Control.Checked = (mViewMode = 已完成病历)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    dkpMan.RecalcLayout
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_FileTab
        Item.Handle = Me.rptList.hWnd
    Case conPane_FileList
        Item.Handle = vfgThis.hWnd
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
    mViewMode = 所有病历
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
'    mstrKinds = ""
'    If InStr(1, mstrPrivs, "门诊病历") > 0 Then mstrKinds = mstrKinds & ",1"
'    If InStr(1, mstrPrivs, "住院病历") > 0 Then mstrKinds = mstrKinds & ",2"
'    If InStr(1, mstrPrivs, "护理记录") > 0 Then mstrKinds = mstrKinds & ",3"
'    If InStr(1, mstrPrivs, "护理病历") > 0 Then mstrKinds = mstrKinds & ",4"
'    If InStr(1, mstrPrivs, "疾病证明报告") > 0 Then mstrKinds = mstrKinds & ",5"
'    If InStr(1, mstrPrivs, "知情文件") > 0 Then mstrKinds = mstrKinds & ",6"
'    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    mstrKinds = "1,2,3,4,5,6"
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
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
        Set cbrControl = .Add(xtpControlButton, conViewAll, "所有病历"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conViewInEditing, "正在书写病历")
        Set cbrControl = .Add(xtpControlButton, conViewFinished, "已完成病历")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
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
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_View_Jump
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conViewAll, "所有病历"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conViewInEditing, "正在书写病历")
        Set cbrControl = .Add(xtpControlButton, conViewFinished, "已完成病历")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    
    Dim panFileTab As Pane, panModel As Pane, panCompend As Pane
    Set panFileTab = dkpMan.CreatePane(conPane_FileTab, 180, 400, DockLeftOf, Nothing)
    panFileTab.Title = "病历文件列表"
    panFileTab.Options = PaneNoCaption
    
    Set panModel = dkpMan.CreatePane(conPane_FileList, 400, 200, DockRightOf, Nothing)
    panModel.Title = "病历清单"
    panModel.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, panModel)
    panCompend.Title = "病历内容"
    panCompend.Options = PaneNoCaption
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.种类, "种类", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 150, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.完成情况, "正在书写/已完成", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    If mstrKinds = "" Then
        DoEvents
        Me.stbThis.Panels(2).Text = "你不具备病历文件定义管理权限"
    Else
        lngCount = Me.zlRefList()
        Me.stbThis.Panels(2).Text = "共有" & lngCount & "份病历文件"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim f As New frmEPRView
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mCol2.ID))
    If lngFileID > 0 Then f.ShowMe Me, lngFileID
End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngCurFileId = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngCurFileId = 0
        Else
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value  '获取当前文件ID
        End If
    End With
    FillGrid mstrFrom, mstrTo
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Public Function zlRefList(Optional lngFileID As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim strGroups As String
    
    gstrSQL = "SELECT L.ID, L.种类, L.编号, L.名称, P.正在书写, P.已完成 " & _
        "FROM 病历文件列表 L, " & _
        "     (SELECT M.文件id, SUM(Decode(M.完成时间, NULL, 1, 0)) AS 正在书写, " & _
        "              SUM(Decode(M.完成时间, NULL, Decode(Sign(SYSDATE - M.创建时间 - 1), 1, 1, 0), 0)) AS 书写超时, " & _
        "              SUM(Decode(M.完成时间, NULL, 0, 1)) AS 已完成, " & _
        "              SUM(Decode(M.完成时间, Null, 0, Decode(NVL(M.签名级别, 0), 0, 1, 0))) As 正在修订 " & _
        "       FROM 电子病历记录 M, 病案主页 N " & _
        "       WHERE M.病人id = N.病人id AND Nvl(N.主页id, 0) <> 0 AND Nvl(N.状态, 0) <> 1 AND " & _
        "             N.入院日期 BETWEEN [3] And [4] AND M.科室id = [1] AND M.病历种类 = [2] AND " & _
        "             M.主页ID = N.主页ID " & _
        "       GROUP BY 文件id) P " & _
        "WHERE L.种类 = [2] AND P.文件id(+) = L.ID"
            
    Err = 0: On Error GoTo errHand
    Dim lngNum1 As Long, lngNum2 As Long
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptId, mvar病历种类, CDate(Format(mstrFrom, "YYYY-MM-DD")), CDate(Format(mstrTo, "YYYY-MM-DD") & " 23:59:59"))
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            If InStr(1, strGroups, !种类) = 0 Then strGroups = strGroups & "," & !种类
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!种类)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            Select Case !种类
            Case 1: rptRcd.AddItem CStr("1-门诊病历")
            Case 2: rptRcd.AddItem CStr("2-住院病历")
            Case 3: rptRcd.AddItem CStr("3-护理记录")
            Case 4: rptRcd.AddItem CStr("4-护理病历")
            Case 5: rptRcd.AddItem CStr("5-疾病证明报告")
            Case 6: rptRcd.AddItem CStr("6-知情文件")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem Val(CStr(!编号))
            rptRcd.AddItem CStr(!名称)
            lngNum1 = NVL(!正在书写, 0)
            lngNum2 = NVL(!已完成, 0)
            rptRcd.AddItem IIf(lngNum1 = 0 And lngNum2 = 0, "", lngNum1 & "/" & lngNum2)
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptList
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mCol.种类)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    If lngFileID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngFileID Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        If Me.rptList.FocusedRow.GroupRow Then
            lngFileID = 0
        Else
            lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
    Else
        lngFileID = 0
    End If
    
    zlRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
    lngFileID = 0
End Function

Private Sub InitGrid()
    Dim i As Long
    With Me.vfgThis
        .Clear
        .Rows = 1
        .FixedRows = 1
        .Cols = 22
        .RowHeightMin = 300
        .WallPaper = imgBG.Picture
        .WallPaperAlignment = flexPicAlignRightBottom
        
'        .BackColorAlternate = RGB(240, 240, 255)
        .BackColorSel = RGB(125, 125, 255)
        .ForeColorSel = vbWhite
        .Sort = flexSortCustom
        
        .TextMatrix(0, mCol2.ID) = "ID"
        .TextMatrix(0, mCol2.病人ID) = "病人ID"
        .TextMatrix(0, mCol2.主页ID) = "主页ID"
        .TextMatrix(0, mCol2.住院号) = "住院号"
        .TextMatrix(0, mCol2.床号) = "床号"
        .TextMatrix(0, mCol2.姓名) = "姓名"
        .TextMatrix(0, mCol2.性别) = "性别"
        .TextMatrix(0, mCol2.年龄) = "年龄"
        .TextMatrix(0, mCol2.住院医师) = "住院医师"
        .TextMatrix(0, mCol2.费别) = "费别"
        .TextMatrix(0, mCol2.入院日期) = "入院日期"
        .TextMatrix(0, mCol2.出院日期) = "出院日期"
        .TextMatrix(0, mCol2.险类) = "险类"
        .TextMatrix(0, mCol2.创建人) = "创建人"
        .TextMatrix(0, mCol2.创建时间) = "创建时间"
        .TextMatrix(0, mCol2.完成时间) = "完成时间"
        .TextMatrix(0, mCol2.保存人) = "保存人"
        .TextMatrix(0, mCol2.保存时间) = "保存时间"
        .TextMatrix(0, mCol2.最后版本) = "最后版本"
        .TextMatrix(0, mCol2.签名级别) = "签名级别"
        .TextMatrix(0, mCol2.归档人) = "归档人"
        .TextMatrix(0, mCol2.归档日期) = "归档日期"
        
'        .MergeCol(mCol2.病人ID) = True
'        .MergeCol(mCol2.主页ID) = True
'        .MergeCol(mCol2.住院号) = True
'        .MergeCol(mCol2.床号) = True
'        .MergeCol(mCol2.姓名) = True
'        .MergeCol(mCol2.性别) = True
'        .MergeCol(mCol2.年龄) = True
'        .MergeCol(mCol2.住院医师) = True
'        .MergeCol(mCol2.费别) = True
'        .MergeCol(mCol2.入院日期) = True
'        .MergeCol(mCol2.出院日期) = True
'        .MergeCol(mCol2.险类) = True
'
'        .MergeCells = flexMergeRestrictColumns
        
        For i = 0 To 21
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next
        .ColWidth(mCol2.ID) = 0
        .ColWidth(mCol2.病人ID) = 800
        .ColWidth(mCol2.主页ID) = 600
        .ColWidth(mCol2.住院号) = 1100
        .ColWidth(mCol2.床号) = 800
        .ColWidth(mCol2.姓名) = 800
        .ColWidth(mCol2.性别) = 600
        .ColWidth(mCol2.年龄) = 600
        .ColWidth(mCol2.住院医师) = 800
        .ColWidth(mCol2.费别) = 800
        .ColWidth(mCol2.入院日期) = 1600
        .ColWidth(mCol2.出院日期) = 1600
        .ColWidth(mCol2.险类) = 600
        .ColWidth(mCol2.创建人) = 800
        .ColWidth(mCol2.创建时间) = 1600
        .ColWidth(mCol2.完成时间) = 1600
        .ColWidth(mCol2.保存人) = 800
        .ColWidth(mCol2.保存时间) = 1600
        .ColWidth(mCol2.最后版本) = 800
        .ColWidth(mCol2.签名级别) = 800
        .ColWidth(mCol2.归档人) = 800
        .ColWidth(mCol2.归档日期) = 1600
    End With
End Sub

Private Sub FillGrid(ByVal strFrom As String, ByVal strTo As String)
    With mfrmContent.edtThis
        .ForceEdit = True
        .ReadOnly = False
        .NewDoc
        .ReadOnly = True
        .ForceEdit = False
    End With
    '填充数据
    Dim Rs As ADODB.Recordset, i As Long, lngCount(1 To 10) As Long
    gstrSQL = "select l.ID, c.病人ID, c.主页ID, c.住院号, c.床号, c.姓名, c.性别, c.年龄, " & _
        "       c.住院医师, c.费别, c.入院日期, c.出院日期, c.险类, l.创建人, " & _
        "       l.创建时间, l.完成时间, l.保存人, l.保存时间, l.最后版本, l.签名级别, " & _
        "       l.归档人 , l.归档日期, l.处理状态 " & _
        "  from 电子病历记录 l, 病人信息 b, " & _
        "       (Select A.病人ID, B.主页ID, A.住院号, A.姓名, A.性别, A.年龄, " & _
        "                B.住院医师, B.出院病床 as 床号, B.费别, B.入院日期, B.出院日期, " & _
        "                b.状态 , b.险类 " & _
        "           From 病人信息 A, 病案主页 B " & _
        "          Where a.病人ID = b.病人ID " & _
        "            And Nvl(B.主页ID, 0) <> 0 " & _
        "            And Nvl(B.状态, 0) <> 1 " & _
        "            And " & _
        "                (B.入院日期 Between [1] And [2]) " & _
        "          Order by 住院号 Desc, 主页ID Desc) c " & _
        " Where c.病人ID = B.病人ID " & _
        "   and l.病人id = B.病人id and l.主页ID = c.主页ID " & _
        "   and l.科室id = " & mlngDeptId & " and l.文件id = " & mlngCurFileId & _
        IIf(mViewMode = 所有病历, "", IIf(mViewMode = 正在书写病历, " and l.完成时间 is null ", " and l.完成时间 is not null ")) & _
        " order by 病人ID,主页ID,创建时间 "
    
    Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
    Call InitGrid
    Me.vfgThis.Rows = 1 + Rs.RecordCount
    
    Me.stbThis.Panels(2).Text = IIf(mViewMode = 正在书写病历, "正在书写", IIf(mViewMode = 所有病历, "所有病历", "已经书写")) & Rs.RecordCount & "份病历"
    i = 1
    Do While Not Rs.EOF
        With Me.vfgThis
            .TextMatrix(i, mCol2.ID) = NVL(Rs("ID"), 0)
            .TextMatrix(i, mCol2.病人ID) = NVL(Rs("病人ID"), 0)
            .TextMatrix(i, mCol2.主页ID) = NVL(Rs("主页ID"), 0)
            .TextMatrix(i, mCol2.住院号) = NVL(Rs("住院号"), 0)
            .TextMatrix(i, mCol2.床号) = NVL(Rs("床号"), 0)
            .TextMatrix(i, mCol2.姓名) = NVL(Rs("姓名"))
            .TextMatrix(i, mCol2.性别) = NVL(Rs("性别"))
            .TextMatrix(i, mCol2.年龄) = NVL(Rs("年龄"))
            .TextMatrix(i, mCol2.住院医师) = NVL(Rs("住院医师"))
            .TextMatrix(i, mCol2.费别) = NVL(Rs("费别"))
            .TextMatrix(i, mCol2.入院日期) = Format(NVL(Rs("入院日期")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.出院日期) = Format(NVL(Rs("出院日期")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.险类) = NVL(Rs("险类"))
            .TextMatrix(i, mCol2.创建人) = NVL(Rs("创建人"))
            .TextMatrix(i, mCol2.创建时间) = Format(NVL(Rs("创建时间")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.完成时间) = Format(NVL(Rs("完成时间")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.保存人) = NVL(Rs("保存人"))
            .TextMatrix(i, mCol2.保存时间) = Format(NVL(Rs("保存时间")), "yyyy-MM-DD HH:nn")
            .TextMatrix(i, mCol2.最后版本) = NVL(Rs("最后版本"))
            .TextMatrix(i, mCol2.签名级别) = NVL(Rs("签名级别"))
            .TextMatrix(i, mCol2.归档人) = NVL(Rs("归档人"))
            .TextMatrix(i, mCol2.归档日期) = Format(NVL(Rs("归档日期")), "yyyy-MM-DD HH:nn")
        End With
        Rs.MoveNext
        i = i + 1
    Loop
    Rs.Close
    Set Rs = Nothing
    If vfgThis.Rows > 1 Then vfgThis.Row = 1: Call vfgThis_RowColChange
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mCol2.ID))
    If lngFileID > 0 Then mfrmContent.zlRefresh lngFileID
End Sub

