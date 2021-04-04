VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockEPRNew 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "新增病历"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3330
      _Version        =   589884
      _ExtentX        =   5874
      _ExtentY        =   2328
      _StockProps     =   0
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      ShowHeader      =   0   'False
   End
   Begin VB.Frame frmBaby 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
      Begin VB.ComboBox cbo婴儿 
         Height          =   300
         ItemData        =   "frmDockEPRNew.frx":0000
         Left            =   600
         List            =   "frmDockEPRNew.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   100
         Width           =   2475
      End
      Begin VB.Label lbl婴儿 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "母婴"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   360
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
      Begin VB.TextBox txtSearchKey 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "输入简码或名称定位"
         Top             =   105
         Width           =   3030
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   840
      Left            =   135
      TabIndex        =   2
      Top             =   5235
      Width           =   2775
      _cx             =   4895
      _cy             =   1482
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
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   555
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":0015
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":05AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":0B49
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":10E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":167D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":1C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":21B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1470
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":26C3
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":2A15
            Key             =   "屏蔽打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":2FAF
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":3549
            Key             =   ""
            Object.Tag             =   "99"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":3AE3
            Key             =   ""
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":407D
            Key             =   ""
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":4617
            Key             =   ""
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":49B1
            Key             =   ""
            Object.Tag             =   "90004"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":4D4B
            Key             =   "刷新"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockEPRNew.frx":B5AD
            Key             =   "预刷新"
         EndProperty
      EndProperty
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   3225
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label lblNeaten 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病历时限提醒"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   930
      Width           =   1350
      WordWrap        =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Line LinTop 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   3375
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "根据病人的诊疗情况，可新增以下的病历。"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   3555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDockEPRNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conColumn_图标 = 0
Const conColumn_ID = 1
Const conColumn_事件 = 2
Const conColumn_名称 = 3
Private Enum mc
    ID = 1
    事件 = 2
    名称 = 3
    简码 = 4
    说明 = 5
End Enum
'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event NewClick(ByVal FileId As Long, ByVal babyNum As Long) '点击增加新病历按钮

'-----------------------------------------------------
'窗体变量
Private mbytScene As Byte       '使用场合

Private mlngDeptId As Long
Private mlngPatiId As Long
Private mlngVisit As Long
Private mblnShowAll As Boolean
Private mstrPrivs As String
Private mlngAdviceID As Long    '医嘱ID
Private mOldSearchKey As String  '搜索关键字
Private mHeight As Long '婴儿选择框隐藏时缩减高度
Private mlngSRow As Long

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim intNum As Integer
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If rptList.FocusedRow Is Nothing Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cbo婴儿.ItemData(cbo婴儿.ListIndex)
        End If
        RaiseEvent NewClick(rptList.FocusedRow.Record(conColumn_ID).Value, intNum)
    Case conMenu_View_Show
        rptList.PreviewMode = Not rptList.PreviewMode
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "文件说明", IIf(rptList.PreviewMode, 1, 0))
    Case conMenu_View_ShowAll
        mblnShowAll = Not mblnShowAll
        Call zlRefList(mbytScene, mlngPatiId, mlngVisit, mlngDeptId, mstrPrivs, mlngAdviceID)
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", IIf(mblnShowAll, 1, 0))
    End Select
End Sub
Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    If vfgThis.TextMatrix(1, 0) = "" Then
        vfgThis.Height = 0: vfgThis.Visible = False
        lblNeaten.Height = 0: lblNeaten.Visible = False
    Else
        vfgThis.Height = lngScaleBottom / 3: vfgThis.Visible = True
        lblNeaten.Height = 250: lblNeaten.Visible = True
    End If
    With Me.Frame
        .Left = 0: .Width = Me.ScaleWidth: .Top = 350
    End With
    With Me.lblNote
        .Left = 150: .Width = Me.ScaleWidth - .Left * 2
        .Top = 500 + 500
    End With
    With Me.frmBaby
        .Left = 0: .Width = Me.ScaleWidth: .Top = lblNote.Top + 45 + lblNote.Height
    End With
    With Me.LinTop
        .X1 = 0: .Y1 = Me.lblNote.Top + Me.lblNote.Height + 45 + Me.frmBaby.Height + mHeight
        .X2 = Me.ScaleWidth: .Y2 = .Y1
    End With
    With Me.rptList
        .Left = 150: .Width = lngScaleRight - .Left * 2
        .Top = Me.LinTop.Y1 + 90: .Height = lngScaleBottom - .Top - 180 - vfgThis.Height - lblNeaten.Height
    End With
    With LineBottom
        .X1 = 0: .Y1 = rptList.Top + rptList.Height + 45
        .X2 = Me.ScaleWidth: .Y2 = .Y1
    End With
    lblNeaten.Move 150, LineBottom.Y1 + 90
    vfgThis.Move 150, lblNeaten.Top + lblNeaten.Height + 50, lngScaleRight - 300
End Sub
 

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Enabled = (Me.rptList.Rows.Count > 0)
        If Control.Enabled Then Control.Enabled = Not (Me.rptList.FocusedRow Is Nothing)
    Case conMenu_View_Show
        Control.Checked = rptList.PreviewMode
        Control.IconId = IIf(Control.Checked, 90002, 90001)
    Case conMenu_View_ShowAll
        Control.Visible = (InStr(";" & mstrPrivs & ";", ";强制书写;") > 0)
        If Control.Visible Then Control.Checked = mblnShowAll
        Control.IconId = IIf(Control.Checked, 90002, 90001)
    End Select
End Sub
Private Function Search()
 Dim searchKey As String, i As Long
 Dim rptRow As ReportRecord
 
    searchKey = txtSearchKey.Text
    For Each rptRow In Me.rptList.Records
        rptRow.Visible = True
        If InStr(CStr(UCase(rptRow(mc.简码).Value)), UCase(searchKey)) > 0 Or InStr(CStr(UCase(rptRow(mc.名称).Value)), UCase(searchKey)) > 0 Or searchKey = "" Then

        Else
            rptRow.Visible = False
        End If
    Next
    
    rptList.Populate
    If rptList.Rows.Count > 0 Then
        If rptList.Rows(0).GroupRow = False Then
            Set rptList.FocusedRow = rptList.Rows(0)
        ElseIf rptList.Rows(1).GroupRow = False Then
            Set rptList.FocusedRow = rptList.Rows(1)
        End If
    End If
End Function

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl, cbrToolBar As CommandBar
    Dim rptCol As ReportColumn, lngCount As Long

    '-----------------------------------------------------
    '内部菜单工具栏定义
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    With cbsThis
        Set .Icons = zlCommFun.GetPubIcons
        .VisualTheme = xtpThemeOfficeXP
        With Me.cbsThis.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '放在VisualTheme后有效
            .UseDisabledIcons = True
            .LargeIcons = False
            .SetIconSize False, 16, 16
            .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
        End With
        .AddImageList img16
        .EnableCustomization False
        .ActiveMenuBar.Visible = False
    End With
    
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示所有"): cbrControl.STYLE = xtpButtonIconAndCaption: cbrControl.flags = xtpFlagRightAlign
        mblnShowAll = IIf(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", "1")) = 1, True, False)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "显示说明"): cbrControl.STYLE = xtpButtonIconAndCaption: cbrControl.flags = xtpFlagRightAlign
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
    End With
    
    '-----------------------------------------------------
    '列表界面定义
    '-----------------------------------------------------
    Me.BackColor = RGB(240, 240, 240)
    With Me.rptList
        Set rptCol = .Columns.Add(conColumn_图标, "种类", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(conColumn_ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_事件, "事件", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .PreviewMode = IIf(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "文件说明", "1")) = 1, True, False)
        With .PaintManager
            .BackColor = Me.BackColor
            .NoItemsText = "没有可新增的电子病历记录..."
            .HighlightBackColor = RGB(135, 195, 255)
            .HighlightForeColor = RGB(0, 0, 0)
            .HorizontalGridStyle = xtpGridNoLines
            .SetPreviewIndent 18, 0, 8, 6
        End With
    End With
    
    With vfgThis
        .Clear
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "病历"
        .TextMatrix(0, 2) = "到期时间"
        .TextMatrix(0, 3) = "必须"
        .ColWidth(0) = 0
        .ColWidth(1) = 1800
        .ColWidth(2) = 2200
        .ColWidth(3) = 400
        
        For lngCount = 0 To 2
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With
    mblnShowAll = False
End Sub
Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim intNum As Integer
    On Error Resume Next
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.Record(conColumn_ID) Is Nothing Then Exit Sub
        If CLng(.FocusedRow.Record(conColumn_ID).Value) = 0 Then Exit Sub
        If KeyCode <> vbKeyReturn Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cbo婴儿.ItemData(cbo婴儿.ListIndex)
        End If
        RaiseEvent NewClick(CLng(.FocusedRow.Record(conColumn_ID).Value), intNum)
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbrPopupBar As CommandBar, cbrControl As CommandBarControl
    If Button <> vbRightButton Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("弹出", xtpBarPopup)
    With cbrPopupBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "显示文件说明(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示所有文件(&A)")
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim intNum As Integer
    On Error Resume Next
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cbo婴儿.ItemData(cbo婴儿.ListIndex)
        End If
        RaiseEvent NewClick(.FocusedRow.Record(conColumn_ID).Value, intNum)
    End With
End Sub

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------

Public Function zlOpenDefaultEPR(Optional ByVal bytKind As Byte = 1) As Boolean
    
    'bytKind=1初诊病历;=2急诊病历;=3复诊病历
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String
    Dim strSort As String
    Dim rsTmp As New ADODB.Recordset
    
    Select Case bytKind
    Case 1
        strTmp = "初诊"
        strSort = "Decode(b.事件,'初诊',1,'门诊',2,3),a.种类,a.编号"
    Case 2
        strTmp = "急诊"
        strSort = "Decode(b.事件,'急诊',1,'门诊',2,3),a.种类,a.编号"
    Case 3
        strTmp = "复诊"
        strSort = "Decode(b.事件,'复诊',1,'门诊',2,3),a.种类,a.编号"
    Case Else
        strTmp = "门诊"
        strSort = "Decode(b.事件,'门诊',1,2),a.种类,a.编号"
    End Select
    

    strSQL = "Select a.Id, a.种类, a.编号, a.名称, a.说明,b.事件 " & _
            " From 病历文件列表 a,病历时限要求 b " & _
            " Where a.ID=b.文件id(+) And Instr(';' || Zl_Out_Epr_Allowed([1], [2], [3],[4]) || ';', ';' || a.Id || ';') <> 0" & _
            " Order By " & strSort
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiId, mlngVisit, mlngDeptId, mlngAdviceID)
        
    If rsTmp.BOF = False Then

        RaiseEvent NewClick(Val(rsTmp("ID").Value), 0)
        
        zlOpenDefaultEPR = True
        
    End If
        
End Function

Public Function zlRefList(ByVal bytScene As Byte, ByVal lngPatiID As Long, ByVal lngVisit As Long, ByVal lngDeptId As Long, _
                            Optional ByVal strPrivs As String, Optional ByVal lngAdviceID As Long) As Long
    '******************************************************************************************************************
    '功能：显示指定病人可以由指定科室增加的电子病历记录文件清单
    '参数： bytScene，电子病历记录增加场景：1-门诊,2-住院,3-护理
    '       lngPatiId，病人id
    '       lngVisit，病人就诊标志。门诊病人为挂号ID，住院病人为主页id
    '       lngDeptId，科室id
    '返回：可增加的电子病历的文件数目，为0则不能增加
    '******************************************************************************************************************
    Dim strIDs As String, blnShowAll As Boolean, lngCount As Long
    Dim rsTemp As New ADODB.Recordset, rsLimit As New ADODB.Recordset
    Dim rptRcd As ReportRecord, rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim blnDisease As Boolean

    mbytScene = bytScene
    mstrPrivs = strPrivs
    mlngPatiId = lngPatiID
    mlngVisit = lngVisit
    mlngDeptId = lngDeptId
    mlngAdviceID = lngAdviceID
    
    blnDisease = (GetPrivFunc(glngSys, 1249) <> "")
    
    Err = 0: On Error GoTo errHand
    Select Case bytScene
    Case 1                  '门诊
        If mblnShowAll Then
            gstrSQL = "Select a.Id, a.种类, a.编号, a.名称,zlspellcode(a.名称) 简码, a.说明,b.事件 " & _
                    " From 病历文件列表 a,病历时限要求 b,病历应用科室 c " & _
                    " Where a.ID=b.文件id(+) And a.ID=c.文件id(+) And a.种类 In " & IIf(blnDisease, " (1,6) ", " (1,5,6) ") & _
                    " And A.保留<>4 And (a.通用=1 Or a.通用=2 And c.科室id=[2])" & _
                    " Order By a.种类, a.编号"
        Else
            gstrSQL = "Select Zl_Out_Epr_Allowed([1], [2], [3],[4]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, mlngAdviceID)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select a.Id, a.种类, a.编号, a.名称,zlspellcode(a.名称) 简码, a.说明,b.事件 " & _
                    " From 病历文件列表 a,病历时限要求 b " & _
                    " Where a.ID=b.文件id(+) And A.保留<>4 And Instr(';' || [1] || ';', ';' || a.Id || ';') <> 0 " & _
                     IIf(blnDisease, " And a.种类 <> 5 ", "") & _
                    " Order By a.种类, a.编号"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select 文件ID ID,病历编号 || '-' || 病历名称 病历名称, 到期时间, Decode(必须,1,'是','否') 必写" & vbNewLine & _
                    "From 电子病历时机" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 科室id =[3] and 病历种类 in " & IIf(blnDisease, " (1,6) ", " (1,5,6) ") & _
                    "And 病人来源 = 1 And (Nvl(完成记录id, 0) = 0 And 完成时间 Is Null)" & vbNewLine & _
                    "Order By 到期时间"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    Case 2                  '住院
        If mblnShowAll Then
        
            gstrSQL = "Select a.Id, a.种类, a.编号, a.名称,zlspellcode(a.名称) 简码, a.说明,b.事件 " & _
                    " From 病历文件列表 a,病历时限要求 b,病历应用科室 c " & _
                    " Where a.ID=b.文件id(+) And a.ID=c.文件id(+) And a.种类 In " & IIf(blnDisease, " (2,6) ", " (2,5,6) ") & _
                    " And A.保留<>4 And (a.通用=1 Or a.通用=2 And c.科室id=[2])" & _
                    " Order By a.种类, a.编号"
        Else
            gstrSQL = "Select Zl_In_Epr_Allowed([1], [2], [3]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, CLng(lngVisit), lngDeptId)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select Id, 种类, 编号, 名称,zlspellcode(名称) 简码, 说明,'' As 事件" & _
                    " From 病历文件列表" & _
                    " Where 保留<>4 And Instr(';' || [1] || ';', ';' || Id || ';') <> 0" & _
                    IIf(blnDisease, " And 种类 <> 5 ", "") & _
                    " Order By 种类, 编号"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select 文件ID ID,病历编号 || '-' || 病历名称 病历名称, 到期时间, Decode(必须,1,'是','否') 必写" & vbNewLine & _
                    "From 电子病历时机" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 科室id =[3] and 病历种类 in " & IIf(blnDisease, " (2,6) ", " (2,5,6) ") & _
                    "And 病人来源 = 2 And (Nvl(完成记录id, 0) = 0 And 完成时间 Is Null)" & vbNewLine & _
                    "Order By 到期时间"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    Case 3                  '护理
        If mblnShowAll Then
        
            gstrSQL = "Select a.Id, a.种类, a.编号, a.名称,zlspellcode(a.名称) 简码, a.说明,b.事件 " & _
                    " From 病历文件列表 a,病历时限要求 b,病历应用科室 c " & _
                    " Where a.ID=b.文件id(+) And a.ID=c.文件id(+) And a.种类=4 And (a.通用=1 Or a.通用=2 And c.科室id=[2])" & _
                    " Order By a.种类, a.编号"
        Else
            gstrSQL = "Select Zl_Nurse_Epr_Allowed([1], [2], [3]) IDS From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, CLng(lngVisit), lngDeptId)
            If rsTemp.EOF Then
                strIDs = ""
            Else
                strIDs = NVL(rsTemp!IDS)
            End If
            gstrSQL = "Select Id, 种类, 编号, 名称,zlspellcode(名称) 简码, 说明,'' As 事件" & _
                    " From 病历文件列表" & _
                    " Where 保留<>4 And Instr(';' || [1] || ';', ';' || Id || ';') <> 0" & _
                    " Order By 种类, 编号"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs, lngDeptId)
    '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select 文件ID ID,病历编号 || '-' || 病历名称 病历名称, 到期时间, Decode(必须,1,'是','否') 必写" & vbNewLine & _
                    "From 电子病历时机" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 科室id =[3] and 病历种类=4 And 病人来源 = 2 And (Nvl(完成记录id, 0) = 0 And 完成时间 Is Null)" & vbNewLine & _
                    "Order By 到期时间"
        Set rsLimit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit, lngDeptId, gstrUserName)
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Select Case !种类
            Case 1: Set rptItem = rptRcd.AddItem(CStr("1-门诊病历")): rptItem.Icon = !种类
            Case 2: Set rptItem = rptRcd.AddItem(CStr("2-住院病历")): rptItem.Icon = !种类
            Case 3: Set rptItem = rptRcd.AddItem(CStr("3-护理记录")): rptItem.Icon = !种类
            Case 4: Set rptItem = rptRcd.AddItem(CStr("4-护理病历")): rptItem.Icon = !种类
            Case 5: Set rptItem = rptRcd.AddItem(CStr("5-疾病证明报告")): rptItem.Icon = !种类
            Case 6: Set rptItem = rptRcd.AddItem(CStr("6-知情文件")): rptItem.Icon = !种类
            Case Else: rptRcd.AddItem ""
            End Select
            
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr("" & !事件)
            rptRcd.AddItem CStr(!编号 & "-" & !名称)
            rptRcd.AddItem CStr(!简码)
            rptRcd.PreviewText = CStr("" & !说明)
            .MoveNext
        Loop
    End With
    Me.rptList.SortOrder.Add rptList.Columns.Find(conColumn_名称)
    Me.rptList.SortOrder(0).SortAscending = True
    Me.rptList.GroupsOrder.DeleteAll
    Me.rptList.GroupsOrder.Add Me.rptList.Columns.Find(0)
    Me.rptList.GroupsOrder(0).SortAscending = True
    Me.rptList.Populate
    For Each rptRow In Me.rptList.Rows
        If rptRow.GroupRow = True And rptRow.Index > 0 Then
           rptRow.Expanded = False
        End If
    Next
    If Me.rptList.Rows.Count > 0 Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    
    
    With vfgThis
        .Clear
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "病历"
        .TextMatrix(0, 2) = "到期时间"
        .TextMatrix(0, 3) = "必须"
        Set .DataSource = rsLimit
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 1800
        .ColWidth(3) = 500
        
        For lngCount = 0 To 2
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With

    
    '------------------------------------------------------------------------------------------------------------------
    '加载婴儿列表
    gstrSQL = "select 序号,decode(婴儿姓名,null,'婴儿'||序号,婴儿姓名)||' 病历' 名称 from 病人新生儿记录 where 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngVisit)
    If rsTemp.RecordCount = 0 Then
        Me.cbo婴儿.Clear
        Me.frmBaby.Visible = False
        mHeight = -Me.frmBaby.Height
    Else
        With rsTemp
            Me.cbo婴儿.Clear
            Me.cbo婴儿.AddItem ("病人病历")
            Do While Not .EOF
                Me.cbo婴儿.AddItem (!名称)
                Me.cbo婴儿.ItemData(Me.cbo婴儿.NewIndex) = NVL(!序号, 0)
                .MoveNext
            Loop
        End With
        Me.cbo婴儿.ListIndex = 0
        mHeight = 0
        Me.frmBaby.Visible = True
    End If
    zlRefList = rptList.Records.Count
    Call cbsThis_Resize
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = rptList.Records.Count
End Function
Private Sub txtSearchKey_Change()
    Call Search
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call Search
End Sub

Private Sub vfgThis_DblClick()
Dim intNum As Integer
    With Me.vfgThis
        If .TextMatrix(1, 0) = "" Then Exit Sub
        If frmBaby.Visible Then
            intNum = Me.cbo婴儿.ItemData(cbo婴儿.ListIndex)
        End If
        RaiseEvent NewClick(CLng(.TextMatrix(.Row, 0)), intNum)
    End With
End Sub


