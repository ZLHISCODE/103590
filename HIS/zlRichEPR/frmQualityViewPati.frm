VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmQualityViewPati 
   Caption         =   "初入院病人病历完成情况"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10875
   Icon            =   "frmQualityViewPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10875
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6555
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityViewPati.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
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
            Picture         =   "frmQualityViewPati.frx":70E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":7C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":81B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":874C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityViewPati.frx":8CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   150
      TabIndex        =   1
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
      Left            =   3240
      TabIndex        =   2
      Top             =   630
      Width           =   2820
      _cx             =   4974
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
      FormatString    =   $"frmQualityViewPati.frx":9280
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2175
      Left            =   225
      TabIndex        =   3
      Top             =   630
      Width           =   2820
      _cx             =   4974
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
      FormatString    =   $"frmQualityViewPati.frx":9355
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2655
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQualityViewPati.frx":942A
      Left            =   900
      Top             =   225
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   8460
      Picture         =   "frmQualityViewPati.frx":943E
      Top             =   4095
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmQualityViewPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColPaiList
    病人ID = 0: 住院号: 姓名: 性别: 年龄: 住院医师: 床号: 费别: 入院日期: 出院日期
End Enum

Private Enum mColDetail
    标志 = 0: 事件及时间: 应写病历: 监测点: 要求时间: 完成时间: 完成记录id: 当前时间: 责任人: 备注说明
End Enum

Private Enum Enum病历种类
    门诊病历 = 1
    住院病历 = 2
    护理病历 = 4
End Enum
Private mvar病历种类 As Enum病历种类

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String         '当前使用者权限串
Private mstrKinds As String         '当前允许定义的病历类型串

Const conPane_PatiList = 201
Const conPane_Detail = 202
Const conPane_Content = 203

Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1

Private mlngCurFileId As Long       '当前文件ID
Private mlngDeptId As Long          '科室ID
Private mstrDeptName As String      '科室名称
Private mstrFrom As String, mstrTo As String

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
    Me.Caption = "初入院病人病历完成情况 - [" & strDeptName & "]"
    Call FillPatiList(mstrFrom, mstrTo)
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
        Control.Enabled = (Me.vfgThis.Records.Count <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    dkpMan.RecalcLayout
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_PatiList
        Item.Handle = vfgList.hWnd
    Case conPane_Detail
        Item.Handle = vfgThis.hWnd
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    If InStr(1, mstrPrivs, "门诊病历") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "住院病历") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "护理记录") > 0 Then mstrKinds = mstrKinds & ",3"
    If InStr(1, mstrPrivs, "护理病历") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "疾病证明报告") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "知情文件") > 0 Then mstrKinds = mstrKinds & ",6"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
'    mstrKinds = "1,2,3,4,5,6"
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    
    Dim panList As Pane, panCompend As Pane, lngCount As Long
    Set panList = dkpMan.CreatePane(conPane_PatiList, 160, 400, DockLeftOf, Nothing)
    panList.Title = "初入院病人列表"
    panList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panList = dkpMan.CreatePane(conPane_Detail, 400, 200, DockRightOf, Nothing)
    panList.Title = "详细情况"
    panList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, panList)
    panCompend.Title = "病历内容"
    panCompend.Options = PaneNoCaption
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
        
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim f As New frmEPRView
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mColDetail.完成记录id))
    If lngFileID > 0 Then f.ShowMe Me, lngFileID
End Sub

Private Sub vfgList_Click()
    Call vfgList_RowColChange
End Sub

Private Sub vfgList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Private Sub vfgList_RowColChange()
    Dim lngPatiId As Long
    lngPatiId = Val(vfgList.TextMatrix(vfgList.Row, mColPaiList.病人ID))
    If lngPatiId > 0 Then FillDetail lngPatiId, 1, mvar病历种类
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dkpMan.RecalcLayout
End Sub

Private Sub FillPatiList(ByVal strFrom As String, ByVal strTo As String)
    '填充病人列表
    Select Case mvar病历种类
    Case 门诊病历
        '应该从挂号记录中提取信息
        gstrSQL = "SELECT A.病人id, A.门诊号, A.姓名, A.性别, A.年龄, B.诊室, B.执行人, " & _
            "       To_Char(B.登记时间, 'yyyy-mm-dd hh24:mi ') AS 登记时间 " & _
            "FROM 病人信息 A, 病人挂号记录 B " & _
            "WHERE A.病人id = B.病人id AND A.门诊号 = B.门诊号 AND Nvl(B.执行状态, 0) <> 0 AND " & _
            "      (A.就诊时间 BETWEEN [2] AND [3]" & _
            " )" & _
            "ORDER BY A.门诊号 DESC"
    Case 住院病历
        gstrSQL = "Select A.病人ID, A.住院号, A.姓名, A.性别, A.年龄, B.住院医师, " & _
            "       B.出院病床 as 床号, B.费别, To_Char(B.入院日期, 'yyyy-mm-dd hh24:mi ') As 入院日期 , To_Char(B.出院日期, 'yyyy-mm-dd hh24:mi ') As 出院日期 " & _
            " From 病人信息 A, 病案主页 B " & _
            " Where a.病人ID = b.病人ID " & _
            "   and A.当前科室ID = [1] " & _
            "   And Nvl(B.主页ID, 0) =1 " & _
            "   And Nvl(B.状态, 0) <> 1 " & _
            "   And (B.入院日期 Between [2] And [3]) " & _
            " Order by B.住院医师 Desc, 住院号 Desc, 主页ID Desc"
    Case 护理病历
        gstrSQL = "Select A.病人ID, A.住院号, A.姓名, A.性别, A.年龄, B.住院医师, " & _
            "       B.出院病床 as 床号, B.费别, To_Char(B.入院日期, 'yyyy-mm-dd hh24:mi ') As 入院日期 , To_Char(B.出院日期, 'yyyy-mm-dd hh24:mi ') As 出院日期 " & _
            " From 病人信息 A, 病案主页 B " & _
            " Where a.病人ID = b.病人ID " & _
            "   and A.当前科室ID = [1] " & _
            "   And Nvl(B.主页ID, 0) =1 " & _
            "   And Nvl(B.状态, 0) <> 1 " & _
            "   And (B.入院日期 Between [2] And [3]) " & _
            " Order by B.住院医师 Desc, 住院号 Desc, 主页ID Desc"
    End Select
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngCount(1 To 6) As Long, strState As String    '用于显示病人分类统计数目

    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptId, CDate(Format(mstrFrom, "YYYY-MM-DD")), CDate(Format(mstrTo, "YYYY-MM-DD") & " 23:59:59"))
    With Me.vfgList
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mColDetail.标志) = 0
    End With
    If vfgList.Rows > 1 Then vfgList.Row = 1: Call vfgList_RowColChange
End Sub

Private Sub FillDetail(ByVal lngPatiId As Long, ByVal lngPageId As Long, ByVal intKind As Integer)
    '填充病历时限检测明细
    With mfrmContent.edtThis
        .ForceEdit = True
        .ReadOnly = False
        .NewDoc
        .ReadOnly = True
        .ForceEdit = False
    End With
    '---------------------------------------------------
    
    '执行时限监测整理
    gstrSQL = "Zl_病历时限监测_Neaten(" & lngPatiId & "," & lngPageId & "," & intKind & ")"
    Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
    
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;
    Dim lngCount As Long, lngBalance As Long, lngDay As Long, lngHour As Long
    gstrSQL = "Select To_Char(B.事件时间, 'yyyy-mm-dd hh24:mi ') || B.变动事件 As 事件及时间, " & _
        "       B.病历编号 || '-' || B.病历名称 As 应写病历, " & _
        "       Decode(B.唯一, 1, '书写', '第' || B.周期号 || '次书写') As 监测点, " & _
        "       B.要求时间, B.完成时间, B.完成记录id, Sysdate As 当前时间, B.责任人, " & _
        "       '' As 备注说明 " & _
        " From 病人信息 A, 病历时限监测 B " & _
        " Where a.病人ID = b.病人ID " & _
        "   and A.病人ID = [1] And B.主页ID=[2] " & _
        "   And (B.病历种类 = [3] Or B.病历种类 in (5, 6) And [3] <> 4) " & _
        "   And B.要求时间 - Sysdate < 2 " & _
        " Order By B.病人ID, B.主页ID, B.病历种类, B.事件时间"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngPatiId, lngPageId, intKind)
    With Me.vfgThis
        .Clear
        .FixedCols = 1
        Set .DataSource = rsTemp
        '标志 = 0: 事件及时间: 应写病历: 监测点: 要求时间: 完成时间: 完成记录ID: 当前时间: 责任人: 备注说明
        
        .MergeCells = flexMergeFree: .MergeCol(mColDetail.事件及时间) = True: .MergeCol(mColDetail.应写病历) = True
        .ColWidth(mColDetail.标志) = 300: .ColWidth(mColDetail.事件及时间) = 2800: .ColWidth(mColDetail.应写病历) = 2000
        .ColWidth(mColDetail.监测点) = 1100: .ColWidth(mColDetail.要求时间) = 1100
        .ColWidth(mColDetail.完成时间) = 0: .ColWidth(mColDetail.完成记录id) = 0: .ColWidth(mColDetail.当前时间) = 0
        .ColWidth(mColDetail.责任人) = 900: .ColWidth(mColDetail.备注说明) = 2200
        
        .FixedAlignment(mColDetail.标志) = flexAlignCenterCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mColDetail.完成时间) = "" Then
                If .TextMatrix(lngCount, mColDetail.完成记录id) = "" Then
                    .TextMatrix(lngCount, mColDetail.备注说明) = "未书写"
                Else
                    .TextMatrix(lngCount, mColDetail.备注说明) = "正在书写"
                End If
                lngBalance = Int((CDate(.TextMatrix(lngCount, mColDetail.当前时间)) - CDate(.TextMatrix(lngCount, mColDetail.要求时间))) * 24)
                .TextMatrix(lngCount, mColDetail.标志) = "！"
                If lngBalance >= 0 Then
                    .Cell(flexcpForeColor, lngCount, mColDetail.标志, lngCount, mColDetail.标志) = RGB(255, 0, 0)
                    If lngBalance > 24 Then
                        '超过24小时，用天算
                        lngDay = lngBalance / 24
                        lngHour = lngBalance Mod 24
                        .TextMatrix(lngCount, mColDetail.备注说明) = .TextMatrix(lngCount, mColDetail.备注说明) & IIf(lngBalance = 0, "", ",已超过" & lngDay & "天" & lngHour & "小时")
                    Else
                        .TextMatrix(lngCount, mColDetail.备注说明) = .TextMatrix(lngCount, mColDetail.备注说明) & IIf(lngBalance = 0, "", ",已超过" & lngBalance & "小时")
                    End If
                    .Cell(flexcpForeColor, lngCount, mColDetail.备注说明, lngCount, mColDetail.备注说明) = RGB(255, 0, 0)
                Else
                    If Abs(lngBalance) < 4 Then
                        .Cell(flexcpForeColor, lngCount, mColDetail.标志, lngCount, mColDetail.标志) = RGB(128, 128, 0)
                        .TextMatrix(lngCount, mColDetail.备注说明) = .TextMatrix(lngCount, mColDetail.备注说明) & ",剩余" & Abs(lngBalance) & "小时,请尽快完成"
                    Else
                        .Cell(flexcpForeColor, lngCount, mColDetail.标志, lngCount, mColDetail.标志) = RGB(0, 0, 255)
                        .TextMatrix(lngCount, mColDetail.备注说明) = .TextMatrix(lngCount, mColDetail.备注说明) & ",剩余" & Abs(lngBalance) & "小时,请按时完成"
                    End If
                End If
            Else
                lngBalance = Int((CDate(.TextMatrix(lngCount, mColDetail.完成时间)) - CDate(.TextMatrix(lngCount, mColDetail.要求时间))) * 24)
                If lngBalance > 0 Then
                    .TextMatrix(lngCount, mColDetail.标志) = "⊕"
                    .Cell(flexcpForeColor, lngCount, mColDetail.标志, lngCount, mColDetail.标志) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mColDetail.备注说明) = "完成,但超过" & lngBalance & "小时"
                    .Cell(flexcpForeColor, lngCount, mColDetail.备注说明, lngCount, mColDetail.备注说明) = RGB(255, 0, 0)
                Else
                    .TextMatrix(lngCount, mColDetail.备注说明) = "正常完成"
                End If
            End If
            .TextMatrix(lngCount, mColDetail.要求时间) = Format(.TextMatrix(lngCount, mColDetail.要求时间), "MM-dd hh:mm")
            .TextMatrix(lngCount, mColDetail.完成时间) = Format(.TextMatrix(lngCount, mColDetail.完成时间), "MM-dd hh:mm")
        Next
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize mColDetail.事件及时间
    End With
    If vfgThis.Rows > 1 Then vfgThis.Row = 1: Call vfgThis_RowColChange
End Sub

Private Sub vfgThis_RowColChange()
    On Error Resume Next
    Dim lngFileID As Long
    lngFileID = Val(vfgThis.TextMatrix(vfgThis.Row, mColDetail.完成记录id))
    If lngFileID > 0 Then
        mfrmContent.zlRefresh lngFileID
    Else
        mfrmContent.edtThis.ForceEdit = True
        mfrmContent.edtThis.ReadOnly = False
        mfrmContent.edtThis.NewDoc
        mfrmContent.edtThis.ReadOnly = True
        mfrmContent.edtThis.ForceEdit = False
    End If
End Sub


