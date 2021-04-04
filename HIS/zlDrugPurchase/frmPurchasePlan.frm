VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchasePlan 
   Caption         =   "采购计划导出"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmPurchasePlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10890
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picGetParams 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   3855
      TabIndex        =   3
      Top             =   2880
      Width           =   3855
      Begin VB.Frame fraParams 
         Height          =   3015
         Left            =   180
         TabIndex        =   17
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox chkUpload 
            Caption         =   "包含已经处理的记录"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtParam02 
            Height          =   270
            Left            =   1440
            TabIndex        =   11
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "审核日期(&C)"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   1290
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "采购单号(&R)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.TextBox txtParam01 
            Height          =   270
            Left            =   1440
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdPS 
            Caption         =   "…"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtProvider 
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpParam01 
            Height          =   270
            Left            =   1440
            TabIndex        =   13
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   284491777
            CurrentDate     =   40290
         End
         Begin MSComCtl2.DTPicker dtpParam02 
            Height          =   270
            Left            =   1440
            TabIndex        =   15
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   280428545
            CurrentDate     =   40290
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   1680
            Width           =   180
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   3
            Left            =   1200
            TabIndex        =   14
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblProvider 
            AutoSize        =   -1  'True
            Caption         =   "供应商(&P)"
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "获取数据(&G)"
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
      End
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2760
      ScaleHeight     =   1575
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   1080
      Width           =   3300
      Begin VSFlex8Ctl.VSFlexGrid vsfView 
         Height          =   1000
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         _cx             =   4683
         _cy             =   1764
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483645
         GridColor       =   -2147483633
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
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
   End
   Begin MSComctlLib.TreeView tvwProvider 
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   7320
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      SimpleText      =   $"frmPurchasePlan.frx":1CFA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchasePlan.frx":1D41
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14129
            Text            =   "蓝色字为已处理过的数据； 红色字为不可选择的数据； 黑色体为正常数据。"
            TextSave        =   "蓝色字为已处理过的数据； 红色字为不可选择的数据； 黑色体为正常数据。"
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
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   8520
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPurchasePlan.frx":25D5
      Left            =   8040
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPurchasePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case enm_Pop_File.FilePrintSet
            frmOutsideLinkSet.Show vbModal, Me
        Case enm_Pop_File.EditProcess
            Call ProcProcess
        Case enm_Pop_File.EditCurrChoose
            SignData vsfView, 4, True
        Case enm_Pop_File.EditCurrCancel
            SignData vsfView, 4, False
        Case enm_Pop_File.EditChooChoose
            SignData vsfView, 3, True
        Case enm_Pop_File.EditChooCancel
            SignData vsfView, 3, False
        Case enm_Pop_File.EditAllChoose
            SignData vsfView, 1, True
        Case enm_Pop_File.EditAllCancel
            SignData vsfView, 0, False
        Case enm_Pop_File.ViewRefresh
            Call cmdGetData_Click
        Case enm_Pop_File.ViewFindButton
            Call FindString
        Case enm_Pop_File.ViewToolsButton
            Control.Checked = Not Control.Checked
            cmbMain(2).Visible = Control.Checked
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsLabel
            Dim cbcControl As CommandBarControl
            Control.Checked = Not Control.Checked
            For Each cbcControl In Me.cmbMain(2).Controls
                cbcControl.Style = IIf(cbcControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsIcon
            Control.Checked = Not Control.Checked
            cmbMain.Options.LargeIcons = Not Me.cmbMain.Options.LargeIcons
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewStatebar
            Control.Checked = Not Control.Checked
            stbThis.Visible = Not stbThis.Visible
            cmbMain.RecalcLayout
        Case enm_Pop_File.FileExit
            Unload Me
    End Select
End Sub

Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cmdGetData_Click()
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strSQL As String, strProvider As String
    Dim isConn As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim dtEnd As Date

    If optParams01(0).Value Then
        If Len(Trim(txtParam01.Text)) = 0 Or Len(Trim(txtParam02.Text)) = 0 Then
            MsgBox "请输入要获取[采购单号]开始、结束的信息！", vbInformation, GSTR_MESSAGE
            txtParam01.SetFocus
            Exit Sub
        End If
    Else
        If Len(Trim(dtpParam01.Value)) = 0 Or Len(Trim(dtpParam02.Value)) = 0 Then
            MsgBox "请输入要获取[审核日期]开始、结束的信息！", vbInformation, GSTR_MESSAGE
            dtpParam01.SetFocus
            Exit Sub
        End If
        If IsDate(dtpParam01.Value) = False Or IsDate(dtpParam02.Value) = False Then
            dtpParam01.SetFocus
            MsgBox "请检查输入的[审核日期]！", vbInformation, GSTR_MESSAGE
            Exit Sub
        End If
    End If

'获取外部数据
'step1 连接外部数据库
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    '默认MSSQL方式连接
    isConn = MSSQLServerOpen(strServer, strDB, strUser, strPWD)
    
    If isConn = False Then
        MsgBox "连接服务器失败，请设置中间数据库的连接！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'step2 获取数据集
    Screen.MousePointer = vbHourglass
    
    strProvider = Trim(txtProvider.Text)
    
    On Error GoTo ErrHand
    strSQL = "select a.id,a.no,b.序号,a.审核日期,a.编制日期,null CSTCODE,null CSTNAME,b.药品id,c.名称,null GOODENAME" _
           & "  ,null GOODGNAME,c.规格,c.药库单位,c.药库包装 PACKNUM,e.id 供应商id,b.上次生产商,a.库房id 药库ID,d.名称 药库" _
           & "  ,a.药房ID, f.名称 药房, b.计划数量/c.药库包装 计划数量" _
           & "  ,b.单价*c.药库包装 单价,null MARK,b.上次供应商,null ImportDate,null ImportUserId,null ImportUserFlag,null ToCode, b.是否上传 " _
           & "from 药品采购计划 a, 药品计划内容 b, 药品目录 c, 部门表 d, 供应商 e, 部门表 f " _
           & "where a.id=b.计划id and b.药品id=c.药品id and a.库房id=d.id(+) and a.药房id=f.id(+) and b.上次供应商=e.名称(+) " _
           & "  and Nvl(B.计划数量, 0) > 0 And Nvl(C.药库包装, 0) > 0" _
           & "  and (d.撤档时间 is null or d.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd')) " _
           & "  and (e.撤档时间 is null or e.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd')) "
    '供应商名称
    If strProvider <> "" Then 'And strProvider <> "[全部]" Then
        strSQL = strSQL & " and b.上次供应商 like '%" & strProvider & "%'"
    End If
    '包含已经上传
    If chkUpload.Value = False Then
        strSQL = strSQL & " and nvl(b.是否上传,0)=0 "
    End If
    If optParams01(0).Value Then        '采购单号
        strSQL = strSQL & " and a.no between [1] and [2] order by " & IIf(chkUpload.Value, "b.是否上传,", "") & "a.no,b.序号"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtParam01.Text, txtParam02.Text)
    Else                                '审核日期
        strSQL = strSQL & " and a.审核日期 between [1] and [2] order by " & IIf(chkUpload.Value, "b.是否上传,", "") & "a.no,b.序号"
        dtEnd = dtpParam02.Value & " 23:59:59"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpParam01.Value, dtEnd)
    End If
    
    If rsTmp.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "ZLHIS数据库上暂时无数据可获取！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
'step3 装载数据
    DataLoading vsfView, rsTmp, 0
    RefreshTVWProvider tvwProvider, vsfView
    Screen.MousePointer = vbDefault
    'MsgBox "获取数据完成！", vbInformation, GSTR_MESSAGE
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "程序异常错误！", vbCritical, GSTR_MESSAGE
End Sub

Private Sub cmdPS_Click()
    ProviderSelecter Me, txtProvider, True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1: Item.Handle = picView.hwnd
        Case 2: Item.Handle = tvwProvider.hwnd
        Case 3: Item.Handle = picGetParams.hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    InitCommandBars cmbMain
    Call InitDKPMain
    Call InitToolBar
    Call SetMenu
    InitVSF vsfView, False
    dtpParam01.Value = Date - 7
    dtpParam02.Value = Date
    optParams01_Click 0
End Sub

Private Sub InitDKPMain()
'初始化dkpMain
    Dim pneMain As Pane, pneProvider As Pane, pneGetParams As Pane, pneFind As Pane
    With dkpMain
        Set pneMain = .CreatePane(1, Me.ScaleHeight, 0, DockRightOf)
        pneMain.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        pneMain.Title = "待处理数据"
        
        Set pneProvider = .CreatePane(2, 230, 400, DockLeftOf)
        pneProvider.Options = PaneNoCloseable + PaneNoFloatable '+ PaneNoHideable
        pneProvider.Title = "供应商列表"
        pneProvider.MinTrackSize.Width = 230
        pneProvider.MinTrackSize.Height = 50
        
        Set pneGetParams = .CreatePane(3, 230, 250, DockBottomOf, pneProvider)
        pneGetParams.Options = PaneNoCloseable + PaneNoFloatable
        pneGetParams.Title = "参数设置"
        pneGetParams.MinTrackSize.Height = 250
        pneGetParams.MaxTrackSize.Height = 250
        pneGetParams.MinTrackSize.Width = 230
        
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        If Not cmbMain Is Nothing Then .SetCommandBars cmbMain
    End With
    
End Sub

Private Sub InitToolBar()
    Dim cbcControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set cbrToolBar = cmbMain.Add("工具栏", xtpBarTop)
    'cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagAlignTop
    With cbrToolBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "设置")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "处理")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "刷新")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "退出")
        cbcControl.BeginGroup = True
    End With
    For Each cbcControl In cbrToolBar.Controls
        If cbcControl.Type = xtpControlButton Then
            cbcControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
'    Set cbrToolBar = cmbMain.Add("查找", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagAlignTop
'    With cbrToolBar.Controls
'        Set cbcControl = .Add(xtpControlLabel, enm_Pop_File.ViewFindTitle, "查找(发票号)：")
'        Set cbcControl = .Add(xtpControlEdit, enm_Pop_File.ViewFindEdit, "")
'        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewFindButton, "")
'    End With
    
End Sub


Private Sub SetMenu()
    Dim cbcControl As CommandBarControl, cbcControlParent As CommandBarControl
    Dim cbpMenuBar As CommandBarPopup
    
    cmbMain.ActiveMenuBar.Title = "菜单"
    cmbMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.File, "文件(&F)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.File
    With cbpMenuBar.CommandBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption & arrMenuBars(1).HotKey)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "外联数据库设置(&S)")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "退出(&X)")
        cbcControl.BeginGroup = True        '以上为一组的开始
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Edit, "编辑(&E)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Edit
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "数据处理(&P)")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrChoose, "当前供应商打勾")
        
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrCancel, "当前供应商取消")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooChoose, "选中打勾")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooCancel, "选中取消")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllChoose, "全部打勾")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllCancel, "全部取消")
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.View, "查看(&V)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.View
    With cbpMenuBar.CommandBar.Controls
        Set cbcControlParent = .Add(xtpControlPopup, enm_Pop_File.ViewTools, "工具栏(&T)")
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsButton, "标准按钮(&S)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsLabel, "文本标签(&T)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsIcon, "大图标(&B)", -1, False)
        cbcControl.Checked = True
        
        Set cbcControlParent = .Add(xtpControlButton, enm_Pop_File.ViewStatebar, "状态栏(&S)")
        cbcControlParent.Checked = True
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "刷新(&R)")
        cbcControl.ShortcutText = "F5"
        cbcControl.BeginGroup = True
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Help, "帮助(&H)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Help
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpHelp, "帮助主题(&H)")
        Set cbcControl = .Add(xtpControlPopup, enm_Pop_File.HelpWeb, "&WEB上的中联")
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebhome, "中联主页(&H)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebBBS, "中联论坛(&F)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebFeelback, "发送反馈(&M)", -1, False
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpAbout, "关于(&A)…")
        cbcControl.BeginGroup = True
    End With
    
    '快键绑定
    With cmbMain.KeyBindings
'        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F5, enm_Pop_File.ViewRefresh
        .Add 0, VK_F1, enm_Pop_File.HelpHelp
    End With
    
    For Each cbcControl In cbpMenuBar.Controls
        cbcControl.Style = xtpButtonIconAndCaption
    Next

End Sub

Private Sub optParams01_Click(Index As Integer)
    Dim lngBackColor As Long
    On Error Resume Next
    If Index = 0 Then
        txtParam01.Enabled = True
        txtParam02.Enabled = True
        txtParam01.BackColor = vbWhite
        txtParam02.BackColor = vbWhite
        dtpParam01.Enabled = False
        dtpParam02.Enabled = False
        txtParam01.SetFocus
    Else
        txtParam01.Enabled = False
        txtParam02.Enabled = False
        txtParam01.BackColor = &H80000004
        txtParam02.BackColor = &H80000004
        dtpParam01.Enabled = True
        dtpParam02.Enabled = True
        dtpParam01.SetFocus
    End If
End Sub

Private Sub picGetParams_Resize()
    fraParams.Width = IIf(picGetParams.Width > 300, picGetParams.Width - 300, 0)
    txtProvider.Width = IIf(picGetParams.Width > 700 + cmdPS.Width, picGetParams.Width - 700 - cmdPS.Width, 0)
    cmdPS.Left = IIf(txtProvider.Width > 0, txtProvider.Left + txtProvider.Width + 20, 0)
    'fraParams01.Width = IIf(picGetParams.Width > fraParams01.Left + 500, picGetParams.Width - fraParams01.Left - 500, 0)
End Sub

Private Sub picView_Resize()
    With vsfView
        .Top = 0
        .Left = 0
        .Width = picView.Width
        .Height = picView.Height
    End With
End Sub

Private Sub tvwProvider_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, intCounter As Integer
    Dim bytState As Byte
    'Check状态显示
    vsfView.Redraw = flexRDNone
    If Node.Key = "Root" Then
        For i = 2 To tvwProvider.Nodes.Count
            tvwProvider.Nodes(i).Checked = Node.Checked
        Next
    Else
        For i = 2 To tvwProvider.Nodes.Count
            If i = 2 Then
                If tvwProvider.Nodes(i).Checked Then
                    bytState = 2
                Else
                    bytState = 1
                End If
            Else
                If (bytState = 1 And tvwProvider.Nodes(i).Checked) Or (bytState = 2 And tvwProvider.Nodes(i).Checked = False) Then
                    bytState = 0
                    Exit For
                End If
            End If
        Next
        Select Case bytState
            Case 1: tvwProvider.Nodes(1).Checked = False
            Case 2: tvwProvider.Nodes(1).Checked = True
            Case Else: tvwProvider.Nodes(1).Checked = 0
        End Select
    End If
    '隐藏VSFView不相干的记录
    If Node.Key = "Root" Then
        For i = 1 To vsfView.Rows - 1
            vsfView.RowHidden(i) = Not Node.Checked
        Next
    Else
        For i = 1 To vsfView.Rows - 1
            If Node.Tag = -1 Then
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) = "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            ElseIf Node.Tag = Val(vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))) Then
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) <> "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            End If
            
        Next
    End If
    '重写序号
    intCounter = 1
    For i = 1 To vsfView.Rows - 1
        If vsfView.RowHidden(i) = False Then
            vsfView.TextMatrix(i, 1) = intCounter
            intCounter = intCounter + 1
        End If
    Next
    vsfView.Redraw = flexRDBuffered
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0: txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ProviderSelecter(Me, txtProvider, False)
    End If
End Sub

Private Sub vsfView_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfView
        '选择项可修改
        If Col = .ColIndex("choose") Then
            If Mid(.TextMatrix(Row, .ColIndex("imported")), 3, 1) = "1" Then
                Cancel = False
            Else
                Cancel = True
            End If
        ElseIf Col = .ColIndex("qty") Then Cancel = False
        ElseIf Col = .ColIndex("remark") Then Cancel = False
        Else: Cancel = True
        End If
    End With
End Sub

Private Sub vsfView_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 3 Then Cancel = True
End Sub

Private Sub vsfView_EnterCell()
    With vsfView
        '调整颜色
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 3)
    End With
End Sub

Private Sub vsfView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    
    If vsfView.Rows <= 1 Then Exit Sub
    
    If Button = vbRightButton Then
        Set objPopupMenu = cmbMain.ActiveMenuBar.FindControl(, enm_Pop_File.Edit)
        If Not objPopupMenu Is Nothing Then
            '遍历要隐藏的菜单项
            For Each cbcControl In objPopupMenu.CommandBar.Controls
                If cbcControl.Id = enm_Pop_File.EditProcess Then
                    cbcControl.Visible = False
                    Exit For
                End If
            Next
            objPopupMenu.CommandBar.ShowPopup
            '恢复
            If Not cbcControl Is Nothing Then
                cbcControl.Visible = True
            End If
        End If
    End If
End Sub

Private Sub vsfView_RowColChange()
    '当前记录用箭头指示
    vsfView.Cell(flexcpText, 0, 0, vsfView.Rows - 1, 0) = ""
    If vsfView.Row > 0 Then
        vsfView.Cell(flexcpFontName, , 0) = "Marlett"
        vsfView.TextMatrix(vsfView.Row, 0) = 4
    End If
End Sub

Private Sub FindString()
    Dim cbeFind As CommandBarEdit
    Set cbeFind = cmbMain.FindControl(, enm_Pop_File.ViewFindEdit)
    
    If cbeFind Is Nothing Then Exit Sub
    
    If Trim(cbeFind.Text) <> "" And vsfView.Rows > 1 Then
        '查找发票号
        Dim i As Integer
        With vsfView
            For i = 1 To .Rows - 1
                If UCase(.TextMatrix(i, .ColIndex("invoice"))) = UCase(Trim(cbeFind.Text)) And .RowHidden(i) = False Then
                    .Row = i
                    .TopRow = i
                    .SetFocus
                    Exit Sub
                End If
            Next
        End With
        MsgBox "未找到你录入的发票号！", , GSTR_MESSAGE
    End If
End Sub

Private Sub ProcProcess()
    Dim strTmp As String
    
    If vsfView.Rows <= 1 Or CheckRecord(vsfView) = False Then
        MsgBox "无数据可以处理，请先获取数据！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    '外部数据库是否连接
    On Error GoTo ExitSub
    If gcnOutside.State = adStateClosed Then gcnOutside.Open
    On Error GoTo 0

    '导入数据库
    If MsgBox("你确定要处理吗？", vbInformation Or vbYesNo Or vbDefaultButton2, GSTR_MESSAGE) = vbNo Then Exit Sub
    
    Call ProcExport
    
    Exit Sub
    
ExitSub:
    MsgBox "外部数据库连接失败!", vbCritical
    Exit Sub
End Sub

Private Sub ProcExport()
    '计划单导出数据处理
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMess As String
    Dim i As Long, intReturn As Long
    
    '检查无供应商的数据
    If CheckRowProvider(i) = False Then
        MsgBox "第" & i & "行无供应商信息。", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    With vsfView
        gcnOutside.BeginTrans
        gcnOracle.BeginTrans
        On Error GoTo ErrHand
        For i = 1 To .Rows - 1
            '跳过处理
            If Val(vsfView.ValueMatrix(i, vsfView.ColIndex("choose"))) = 0 Or vsfView.RowHidden(i) = True Then GoTo ProcEnd
            
            strSQL = "declare @i_rtn int, @s_msg varchar(200) " & Chr(13)
            strSQL = strSQL & "execute sj_insertBill_pro " _
                   & " '" & .TextMatrix(i, .ColIndex("planno")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("xh")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("cdate")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("edate")) & "'" _
                   & ", null, null" _
                   & ",'" & .TextMatrix(i, .ColIndex("id")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("name")) & "'" _
                   & ",null, null" _
                   & ",'" & .TextMatrix(i, .ColIndex("spec")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("unit")) & "'" _
                   & ",null" _
                   & ",'" & .TextMatrix(i, .ColIndex("producer")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("wh_id")) _
                   & "|" & IIf(Val(.TextMatrix(i, .ColIndex("dh_id"))) = 0, .TextMatrix(i, .ColIndex("wh_id")), .TextMatrix(i, .ColIndex("dh_id"))) & "'" _
                   & ",'" & IIf(Val(.TextMatrix(i, .ColIndex("dh_id"))) = 0, .TextMatrix(i, .ColIndex("wh")), .TextMatrix(i, .ColIndex("dh"))) & "'" _
                   & "," & .TextMatrix(i, .ColIndex("qty")) _
                   & "," & Round(.TextMatrix(i, .ColIndex("price")), 2) _
                   & ",'" & .TextMatrix(i, .ColIndex("remark")) & "'" _
                   & ",'" & IIf(.TextMatrix(i, .ColIndex("providerid")) = "" _
                     , .TextMatrix(i, .ColIndex("provider")) _
                     , .TextMatrix(i, .ColIndex("providerid"))) & "'" _
                   & ",null,null,null,null,null,@i_rtn output, @s_msg output " & Chr(13)
            strSQL = strSQL & "select @i_rtn i_rtn, @s_msg s_msg "
            rsTmp.Open strSQL, gcnOutside
            If rsTmp.EOF Then
                intReturn = 0
                strMess = ""
            Else
                intReturn = rsTmp!i_rtn
                strMess = rsTmp!s_msg
            End If
            rsTmp.Close
            
            '计划内容标记上传
            strSQL = "zl_药品计划内容_Upload(" & .TextMatrix(i, .ColIndex("planid")) _
                   & "," & .TextMatrix(i, .ColIndex("xh")) & ")"
            gobjComLib.zlDatabase.ExecuteProcedure strSQL, Me.Caption & "-标记计划上传"
            
            .TextMatrix(i, .ColIndex("mess")) = strMess
            If intReturn = 1 Then
                .TextMatrix(i, .ColIndex("mess")) = "OK"
            End If
            
ProcEnd:
        Next

        gcnOracle.CommitTrans
        gcnOutside.CommitTrans
        '清理已经处理的数据
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("mess")) = "OK" Then
                .RemoveItem i
            Else
                If .Cell(flexcpChecked, i, .ColIndex("choose")) = Checked And InStr(.TextMatrix(i, .ColIndex("mess")), "已存在") > 0 Then
                    If MsgBox("第" & i & "行数据已经有导出过。是否继续查看提示？", vbInformation + vbYesNo, GSTR_MESSAGE) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        Next
    
    End With
    
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    gcnOutside.RollbackTrans
    Call gobjComLib.ErrCenter
End Sub

Private Sub SignData(ByVal vsfVal As VSFlexGrid, ByVal bytVal As Byte, ByVal blnVal As Boolean)
'0: 全部取消; 1:全部选中; 2: 选中取消; 3:选中打勾; 4:供应商
    Dim i As Integer
    Dim strTmp As String
    
    If vsfVal.Rows < 2 Then Exit Sub
    
    With vsfVal
        strTmp = .TextMatrix(.Row, .ColIndex("provider"))
        '注意: SelectedRows要生效，SelectMode需要为 flexSelectionListBox
        For i = 1 To .Rows - 1
            Select Case bytVal
                Case 0, 1
                    'vsfView.TextMatrix(i, 2) = IIf(blnVal And Mid(vsfView.TextMatrix(i, vsfView.ColIndex("imported")), 3, 1) = "1", "1", "0")
                    .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                Case 2, 3
                    If .IsSelected(i) = True Then
                        .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                    End If
                Case 4
                    If .TextMatrix(i, .ColIndex("provider")) = strTmp Then
                        .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                    End If
            End Select
        Next
    End With
End Sub

Private Function CheckRowProvider(ByRef lngRow As Long) As Boolean
'-----------------------------------------------
'功能：检查供应商名称
'参数：lngRow没有供应商名称的行号
'返回值：False检查未通过；True检查通过
'-----------------------------------------------
    Dim i As Long

    With vsfView
        For i = 1 To .Rows - 1
            If Val(.ValueMatrix(i, .ColIndex("choose"))) <> 0 And vsfView.RowHidden(i) = False Then
                If Trim(.TextMatrix(i, .ColIndex("provider"))) = "" Then
                    lngRow = i
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRowProvider = True
End Function
