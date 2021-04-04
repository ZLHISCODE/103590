VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRBillMan 
   Caption         =   "诊疗单据管理"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8970
   Icon            =   "frmEPRBillMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   75
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   450
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   15
         TabIndex        =   1
         Top             =   30
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   4800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRBillMan.frx":038A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRBillMan.frx":0724
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5625
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRBillMan.frx":0FB6
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRBillMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: ID: 编号: 名称: 说明: 通用: 保留: 格式
End Enum

Const conPane_List = 201
Const conPane_Edit = 202
Const conPane_Apply = 203

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mfrmEdit As frmEPRBillEdit
Private mfrmApply As frmEPRBillApply
Private mObjTabEpr As cTableEPR

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态, 2-适用项目编辑
Private mlngBillID As Long
Private mLngEditWidth As Long, mlngEditHeight As Long      '为适应大字体情况下窗体变大.先读入窗体大小.

Public Function zlRefList(Optional lngItemID As Long) As Long
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ID, 编号, 名称, 说明, 通用,保留,格式 From 病历文件列表 Where 种类 = 7"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr("0")): rptItem.Icon = 0
            rptRcd.AddItem CStr("" & !ID)
            Set rptItem = rptRcd.AddItem(CStr(!编号)): rptItem.SortPriority = Val(("" & !编号))
            rptRcd.AddItem CStr("" & !名称)
            rptRcd.AddItem CStr("" & !说明)
            Select Case Val("" & !通用)
            Case 2: rptRcd.AddItem CStr("有, 自定报表输出")
            Case 1: rptRcd.AddItem CStr("有, 编辑格式输出")
            Case Else: rptRcd.AddItem CStr("")
            End Select
            rptRcd.AddItem NVL(!保留, 0)
			rptRcd.AddItem Decode(NVL(!格式, 0), 0, "普通诊疗单据", 1, "自定义申请单")	
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If lngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If Val(rptRow.Record(mCol.ID).Value) = lngItemID Then
                Set Me.rptList.FocusedRow = rptRow
                Exit For
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "种单据"
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "诊疗单据目录"
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


Private Sub ShowAffixConfig(ByVal lngBillId As Long)
'显示附项模板配置
On Error GoTo ErrHandle
    Dim frmAffix As New frmAffixTemplate
    
    Call frmAffix.ShowAffixConfig(lngBillId, Me)
ErrHandle:
    Set frmAffix = Nothing
    
End Sub


'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long
Dim cbrControl As CommandBarControl
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save
        Select Case mintEditState
        Case 1
            lngRetuId = mfrmEdit.zlEditSave()
            If lngRetuId = 0 Then Exit Sub
            mlngBillID = lngRetuId: Call zlRefList(mlngBillID)
        Case 2
            lngRetuId = mfrmApply.zlEditSave()
            If lngRetuId = 0 Then Exit Sub
        End Select
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_Untread
        Select Case mintEditState
        Case 1
            Call mfrmEdit.zlEditCancel
        Case 2
            Call mfrmApply.zlEditCancel
        End Select
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
	Case conMenu_Edit_ApplyView
        Call ShowApplyCustom
    Case conMenu_Edit_NewItem
        If mfrmEdit.zlEditStart(True, mlngBillID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Modify
        If mlngBillID = 0 Then Exit Sub
        If mfrmEdit.zlEditStart(False, mlngBillID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "真的删除该诊疗单据吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.名称).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_诊疗单据目录_Edit(3," & mlngBillID & ")"
            
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            mlngBillID = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                lngRetuId = lngRetuId + 1
            ElseIf lngRetuId > 0 Then
                lngRetuId = lngRetuId - 1
            End If
            mlngBillID = .Rows(lngRetuId).Record(mCol.ID).Value
            Call Me.zlRefList(mlngBillID)
        End With
    
    Case conMenu_Edit_ApplyTo
        If mfrmApply.zlEditStart() = False Then Exit Sub
        mintEditState = 2: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Apply).Select
        
    Case conMenu_Edit_Compend
        If rptList.FocusedRow.Record.Item(mCol.保留).Value = 0 Then
            Dim Doc As New cEPRDocument
            If mlngBillID = 0 Then Exit Sub
            Doc.InitEPRDoc cprEM_修改, cprET_病历文件定义, mlngBillID
            Doc.ShowEPREditor Me
        Else
            On Error GoTo errHand
            mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_病历文件定义, mlngBillID
        End If
    Case conMenu_Edit_Affix '附项模板设置
        Call ShowAffixConfig(mlngBillID)
        
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(mlngBillID)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        If Control.ID > 401 And Control.ID < 499 Then
            '执行自定义报表
            Call BillPrint_Custom(Control)
        End If
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub ShowApplyCustom()
    Dim objKernel As Object
    
    Err = 0: On Error Resume Next
    Set objKernel = CreateObject("zlPublicAdvice.clsPublicAdvice")
    If Err <> 0 Then
        'Call MsgBox("公共医嘱部件丢失,医嘱信息将显示异常,请与系统管理员联系!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName)
        Err = 0: On Error GoTo 0
        Exit Sub
    End If
    Err = 0: On Error GoTo 0:
    If Not objKernel Is Nothing Then
        Call objKernel.InitCommon(gcnOracle, glngSys)
        Call objKernel.ShowApplyCustom(Me, mlngBillID)
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
        
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
        
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "定义") > 0 And mintEditState = 0)
        
    Case conMenu_Edit_Affix
        Control.Enabled = InStr(1, mstrPrivs, "定义") > 0
        
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_ApplyTo
        Control.Enabled = (InStr(1, mstrPrivs, "定义") > 0 And mintEditState = 0)
        
        If Control.Enabled Then Control.Enabled = mlngBillID <> 0
    Case conMenu_Edit_ApplyView
        Control.Visible = mintEditState = 0 And mlngBillID <> 0
        If Control.Visible Then
            Control.Visible = rptList.FocusedRow.Record.Item(mCol.格式).Value = "自定义申请单"
        End If    
    Case conMenu_Edit_Compend
        Control.Enabled = (InStr(1, mstrPrivs, "样式构造") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = mlngBillID <> 0
        If Control.Enabled Then Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.通用).Value) <> "")
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmEPRBillEdit
        Item.Handle = mfrmEdit.hWnd
    Case conPane_Apply
        If mfrmApply Is Nothing Then Set mfrmApply = New frmEPRBillApply
        Item.Handle = mfrmApply.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmEPRBillEdit.ScaleWidth
    mlngEditHeight = frmEPRBillEdit.ScaleHeight
    
    mintEditState = 0
    mlngBillID = 0
    Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
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
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "适用项目(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "报告构造(&F)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Affix, "附项模板设置(&I)"): cbrControl.BeginGroup = True: cbrControl.IconId = 1
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
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "适用项目"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "报告构造")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "预览申请单"): cbrControl.IconId = 3002
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panSub As Pane
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmEPRBillEdit
    If mfrmApply Is Nothing Then Set mfrmApply = New frmEPRBillApply
    
    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "单据列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_Edit, 550, 350, DockRightOf, Nothing)
    panThis.Title = "单据编辑"
    panThis.Options = PaneNoCaption
    Set panSub = dkpMan.CreatePane(conPane_Apply, 550, 450, DockBottomOf, panThis)
    panSub.Title = "适用该单据的诊疗项目"
    panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 130, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.说明, "说明", 250, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.通用, "报告", 100, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.保留, "保留", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.格式, "格式", 130, False): rptCol.Editable = False: rptCol.Groupable = True
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列性质..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    '数据装入
    Call zlRefList

End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表
    
    '默认参数：单据=病历文件目录id
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
        "单据=" & IIf(mlngBillID = 0, "", mlngBillID))
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panThis = Me.dkpMan.FindPane(conPane_Edit)
    panThis.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, 0
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload mfrmEdit: Set mfrmEdit = Nothing
    Unload mfrmApply: Set mfrmApply = Nothing
    Set mObjTabEpr = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub piclist_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.ID))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mlngBillID = 0 Then Exit Sub
Dim cbrControl As CommandBarControl
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngBillID = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngBillID = 0
        Else
            mlngBillID = Val(.FocusedRow.Record.Item(mCol.ID).Value)
        End If
        Call mfrmEdit.zlRefresh(mlngBillID)
        Call mfrmApply.zlRefresh(mlngBillID)
    End With
End Sub


