VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrugSendQuery 
   AutoRedraw      =   -1  'True
   Caption         =   "药疗收发查询"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "frmDrugSendQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picQuery 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   1200
      ScaleHeight     =   5535
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   1080
      Width           =   9015
      Begin VSFlex8Ctl.VSFlexGrid vsQuery 
         Height          =   5355
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   9030
         _cx             =   15928
         _cy             =   9446
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
         BackColorSel    =   16764057
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDrugSendQuery.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         AutoSizeMouse   =   0   'False
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
         Begin VB.Frame fraColSel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   195
            Begin VB.Image imgColSel 
               Height          =   195
               Left            =   0
               Picture         =   "frmDrugSendQuery.frx":00A7
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsColumn 
            Height          =   3270
            Left            =   4800
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   1470
            _cx             =   2593
            _cy             =   5768
            Appearance      =   0
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
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorSel    =   14737632
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmDrugSendQuery.frx":05F5
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
            Editable        =   2
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6420
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "通过"
            TextSave        =   "通过"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "疑问"
            TextSave        =   "疑问"
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
   Begin XtremeSuiteControls.TabControl tbcQuery 
      Height          =   5715
      Left            =   825
      TabIndex        =   1
      Top             =   405
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   10081
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   165
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDrugSendQuery.frx":0643
      Left            =   675
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDrugSendQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmCond As frmDrugSendQueryCond
Attribute mfrmCond.VB_VarHelpID = -1

Public mMainPrivs As String 'IN:调用主界面所具有的权限,注意非内部模块权限
Public mlng病区ID As Long 'IN:用于记录主界面的病区及上次查询病区
Public mlng病人ID As Long 'IN
Private mblnOnePati As Boolean 'IN，单病人模式
Private mblnMoved As Boolean
Private mbytSize As Byte '表格字体：0-9号字体（小字体），1-12号字体（大字体）
Private mstrNewHead  As String '发药明细清单在选择列后的列头信息
'发药明细清单原本的列头信息
Private Const mstrOldHead = "期效,850,1;状态,850,1;药品信息,5000,1;付数,850,1;数量,850,1;单价,1000,7;金额,1000,7;单量,850,1;频次,1000,1;用法,1000,1;发送时间,1530,1;发送人,750,1"

'查询条件
Private Type QUERY_COND
    Mode As Byte '0-按医嘱发送时间,1-按药房发药时间
    DateBegin As Date
    DateEnd As Date
    退药DateB As Date
    退药DateE As Date
    给药途径 As String
    NO As String
    发药号 As String
    药房ID As Long
    病人IDs As String
    病区ID As Long
    领药部门ID As Long
    期效 As Integer '2-全部
    状态 As String
End Type
Private mvQuery As QUERY_COND

Public Sub ShowQuery(frmParent As Object, strPriv As String, lng病区ID As Long, lng病人ID As Long, ByVal blnOnePati As Boolean)
    mMainPrivs = strPriv
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mblnOnePati = blnOnePati
    Me.Show , frmParent
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_File_Print
        Call OutputList(1)
    Case conMenu_File_Preview
        Call OutputList(2)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    Case conMenu_FontSet_FontSize_S '小字体
        If mbytSize <> 0 Then
            mbytSize = 0
            Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
            Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
        End If
    Case conMenu_FontSet_FontSize_L '大字体
        If mbytSize <> 1 Then
            mbytSize = 1
            Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
            Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
        End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    With Me.tbcQuery
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_FontSet_FontSize_S '小字体
            Control.Checked = Not (mbytSize = 1)
        Case conMenu_FontSet_FontSize_L '大字体
            Control.Checked = (mbytSize = 1)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Item.Handle = mfrmCond.Hwnd
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objItem As TabControlItem
    Dim objPane As Pane

    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_FontSet, "表格字体")
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_FontSet_FontSize_S, "小字体(&S)", -1, False
            .Add xtpControlButton, conMenu_FontSet_FontSize_L, "大字体(&L)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_FontSet_FontSize_S
        .Add FALT, vbKeyL, conMenu_FontSet_FontSize_L
    End With

    '条件栏----------------------------------------------
    Set mfrmCond = New frmDrugSendQueryCond
    Call mfrmCond.InitParameter(mMainPrivs, mlng病区ID, mlng病人ID, mblnOnePati)
    
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 250, 400, DockLeftOf, Nothing)
    objPane.Title = "查询条件"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    mstrNewHead = mstrOldHead
    '页面数据----------------------------------------------
    
    
    
    With Me.tbcQuery
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
        End With
        Set objItem = .InsertItem(0, "发药明细清单", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        Set objItem = .InsertItem(1, "发药汇总清单", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        Set objItem = .InsertItem(2, "退药明细清单", picQuery.Hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(3, "退药汇总清单", picQuery.Hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(4, "发退汇总清单", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        
        '因为绑定相同,最后要切换回第1个;无数据不影响速度
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    '设置表格字体
    mbytSize = Val(zlDatabase.GetPara("药疗收发查询表格字体", glngSys, p住院医嘱发送, "0"))
    Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
    Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
            
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mvQuery.DateBegin = Empty
    Call zlDatabase.SetPara("药疗收发查询表格字体", mbytSize, glngSys, p住院医嘱发送, InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0)
    If Not mfrmCond Is Nothing Then
        Unload mfrmCond
        Set mfrmCond = Nothing
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitQueryTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    If tbcQuery.Selected.Index = 0 Then
        strHead = mstrNewHead
    ElseIf tbcQuery.Selected.Index = 1 Then
        strHead = "编码,1000,1;药品信息,5000,1;数量,850,1;金额,1000,7"
    ElseIf tbcQuery.Selected.Index = 2 Then
        strHead = "期效,500,1;药品信息,5000,1;付数,850,1;数量,850,1;单价,1000,7;金额,1000,7;单量,850,1;频次,1000,1;用法,1000,1;申请时间,1530,1;申请人,750,1"
    ElseIf tbcQuery.Selected.Index = 3 Then
        strHead = "编码,1000,1;药品信息,5000,1;数量,850,1;金额,1000,7"
    ElseIf tbcQuery.Selected.Index = 4 Then
        strHead = "编码,1000,1;药品信息,5000,1;应发数,850,1;退药数,850,1;实发数,850,1;金额,1000,7"
    End If
    arrHead = Split(strHead, ";")
    With vsQuery
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)
                If .ColWidth(.FixedCols + i) = 0 Then
                    .ColHidden(.FixedCols + i) = True
                End If
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Function LoadQueryData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strKey As String, curTotal As Currency
    Dim strSQL As String, strSQLSend As String, strSQLDel As String
    Dim strCondSend As String, strCondDel As String
    Dim strSub1 As String, strSub2 As String
    Dim i As Long, j As Long
    Dim intBedLen As Integer
    
    If mvQuery.DateBegin = Empty Then
        vsQuery.Rows = vsQuery.FixedRows
        vsQuery.Rows = vsQuery.FixedRows + 1
        LoadQueryData = True: Exit Function
    End If
    
    Me.Refresh
    
    On Error GoTo errH
    
    '发药部份条件及SQL
    '------------------------------------------------------------------------------------------
    '病人无法确定,只能以时间来确定
    mblnMoved = zlDatabase.DateMoved(mvQuery.DateBegin)
    
    If InStr(",0,1,4,", tbcQuery.Selected.Index) > 0 Then
        '药房
        If mvQuery.药房ID <> 0 Then
            strCondSend = strCondSend & " And A.库房ID+0=[1]"
        End If
        
        '时间:要以发送时间,因为退药时的填制日期不一样
        If mvQuery.Mode = 0 Then
            strCondSend = strCondSend & " And (A.NO,A.单据) IN (Select NO,Decode(记录性质,1,8,9) From 病人医嘱发送 Where 发送时间 Between [2] And [3])"
        Else
            strCondSend = strCondSend & " And A.审核日期 Between [2] And [3]"
        End If
        
        'NO
        If mvQuery.NO <> "" Then strCondSend = strCondSend & " And A.NO=[4]"
        '发药号
        If mvQuery.发药号 <> "" Then strCondSend = strCondSend & " And A.汇总发药号=[12] "
        '期效
        If mvQuery.期效 = 0 Or mvQuery.期效 = 1 Then
            strCondSend = strCondSend & " And Nvl(Substr(A.扣率,1,1),0)=[5]"
        End If
        
        '发药状态
        If Val(Mid(mvQuery.状态, 1, 1)) = 1 And Val(Mid(mvQuery.状态, 2, 1)) = 1 Then
            strCondSend = strCondSend & " And (Mod(A.记录状态,3)=1 And A.审核人 is Null Or A.审核人 is Not Null)"
        ElseIf Val(Mid(mvQuery.状态, 1, 1)) = 1 Then
            strCondSend = strCondSend & " And Mod(A.记录状态,3)=1 And A.审核人 is Null"
        ElseIf Val(Mid(mvQuery.状态, 2, 1)) = 1 Then
            strCondSend = strCondSend & " And A.审核人 is Not Null"
        End If
        
        '给药途径
        If mvQuery.给药途径 <> "" Then
            strCondSend = strCondSend & " And A.用法 IN(Select Column_Value From Table(Cast(f_Str2list([8]) As zlTools.t_Strlist)))"
        End If
        
        '公共部份SQL
        strSQLSend = "Select Decode(A.审核人,NULL,0,1) as 状态,A.NO,A.序号," & _
            " Sum(A.填写数量*A.付数) as 数量,Sum(A.零售金额) as 金额" & _
            " From 药品收发记录 A" & _
            " Where A.单据 IN(9,10)" & IIF(mvQuery.领药部门ID = 0, "", " And a.对方部门ID=[13]") & strCondSend & _
            " Group by Decode(A.审核人,NULL,0,1),A.NO,A.序号" & _
            " Having Nvl(Sum(A.填写数量),0)<>0 Or Nvl(Sum(A.零售金额),0)<>0"
        strSQLSend = "Select B.状态,C.NO,D.序号,C.库房ID,C.对方部门ID,D.姓名,D.标识号 as 住院号,D.床号," & _
            " C.药品ID,B.数量,D.标准单价 as 单价,B.金额,Decode(Nvl(Substr(C.扣率,1,1),0),0,'长嘱','临嘱') as 期效," & _
            " C.单量,C.频次,C.用法,A.发送时间 as 时间,A.发送人 as 人员,C.付数,d.病人id,d.主页id" & _
            " From 病人医嘱发送 A,(" & strSQLSend & ") B,药品收发记录 C,住院费用记录 D" & _
            " Where B.NO=C.NO And B.序号=C.序号 And (C.记录状态=1 Or Mod(C.记录状态,3)=0)" & _
             IIF(mvQuery.领药部门ID = 0, "", " And c.对方部门ID=[13]") & _
            " And C.费用ID=D.ID And A.NO=D.NO And A.医嘱ID=D.医嘱序号 And A.记录性质=2" & _
            IIF(mvQuery.病区ID <> 0, " And D.病人病区ID+0=[6]", "") & _
            IIF(mvQuery.病人IDs <> "", " And D.病人ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([7]) As zlTools.t_Numlist)))", "")
        If mblnMoved Then
            strSub1 = strSQLSend
            strSub1 = Replace(strSub1, "住院费用记录", "H住院费用记录")
            strSub1 = Replace(strSub1, "药品收发记录", "H药品收发记录")
            
            strSub2 = strSQLSend
            strSub2 = Replace(strSub2, "病人医嘱发送", "H病人医嘱发送")
            strSub2 = Replace(strSub2, "住院费用记录", "H住院费用记录")
            strSub2 = Replace(strSub2, "药品收发记录", "H药品收发记录")

            strSQLSend = strSQLSend & " Union ALL " & strSub1 & " Union ALL " & strSub2
        End If
        '将分批明细合并
        strSQLSend = "Select 状态,NO,序号,库房ID,对方部门ID,姓名,住院号,床号,药品ID,单价," & _
            " Sum(数量) as 数量,Sum(金额) as 金额,期效,单量,频次,用法,时间,人员,付数,病人id,主页id" & _
            " From (" & strSQLSend & ")" & _
            " Group by 状态,NO,序号,库房ID,对方部门ID,姓名,住院号,床号,药品ID,单价,期效,单量,频次,用法,时间,人员,付数,病人id,主页id"
    End If
    
    '退药部份条件及SQL
    '------------------------------------------------------------------------------------------
    If InStr(",2,3,4,", tbcQuery.Selected.Index) > 0 Then '退药部份
        '药房
        strCondDel = strCondDel & " And A.审核部门ID=[1]"
        
        '退药申请时间
        If mvQuery.退药DateB <> Empty And mvQuery.退药DateE <> Empty Then
            strCondDel = strCondDel & " And A.申请时间 Between [9] And [10]"
        Else
            strCondDel = strCondDel & " And A.申请时间 Between Sysdate-1 And Sysdate"
        End If
        
        'NO
        If mvQuery.NO <> "" Then
            strCondDel = strCondDel & " And B.NO=[4]"
        End If
        '发药号
        If mvQuery.发药号 <> "" Then strCondDel = strCondDel & " And D.汇总发药号=[12] "
        '期效
        If mvQuery.期效 = 0 Or mvQuery.期效 = 1 Then
            strCondDel = strCondDel & " And Nvl(C.医嘱期效,0)=[5]"
        End If
        
        '病区
        strCondDel = strCondDel & " And A.申请部门ID=[6]"
        
        '病人ID
        If mvQuery.病人IDs <> "" Then
            strCondDel = strCondDel & " And B.病人ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([7]) As zlTools.t_Numlist)))"
        End If
        
        '给药途径
        If mvQuery.给药途径 <> "" Then
            strCondDel = strCondDel & " And D.用法 IN(Select Column_Value From Table(Cast(f_Str2list([8]) As zlTools.t_Strlist)))"
        End If
        
        '公共部份SQL：将分批明细合并
        strSQLDel = "Select Distinct -1 as 状态,D.NO,B.序号,D.库房ID,D.对方部门ID,B.姓名,B.标识号 as 住院号,B.床号," & _
            " D.药品ID,A.数量,B.标准单价 as 单价,A.数量*B.标准单价 as 金额,Decode(Nvl(C.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
            " D.单量,D.频次,D.用法,A.申请时间 as 时间,A.申请人 as 人员,D.付数,b.病人id,b.主页id" & _
            " From 病人费用销帐 A,住院费用记录 B,病人医嘱记录 C,药品收发记录 D" & _
            " Where Nvl(A.状态,0)=0 And A.费用ID=B.ID And B.收费类别 IN('5','6','7')" & _
            IIF(mvQuery.领药部门ID = 0, "", " And d.对方部门ID=[13]") & _
            " And B.医嘱序号=C.ID And B.ID=D.费用ID And (D.记录状态=1 Or Mod(D.记录状态,3)=0)" & strCondDel
        If mblnMoved Then
            strSub1 = strSQLDel
            strSub1 = Replace(strSub1, "住院费用记录", "H住院费用记录")
            strSub1 = Replace(strSub1, "药品收发记录", "H药品收发记录")
            
            strSub2 = strSQLDel
            strSub2 = Replace(strSub2, "病人医嘱记录", "H病人医嘱记录")
            strSub2 = Replace(strSub2, "住院费用记录", "H住院费用记录")
            strSub2 = Replace(strSub2, "药品收发记录", "H药品收发记录")

            strSQLDel = strSQLDel & " Union ALL " & strSub1 & " Union ALL " & strSub2
        End If
    End If
    
    '产生不同的查询SQL
    '------------------------------------------------------------------------------------------
    If tbcQuery.Selected.Index = 0 Or tbcQuery.Selected.Index = 2 Then '发药明细、退药明细
        intBedLen = GetMaxBedLen(mvQuery.病区ID, False)
        strSQL = IIF(tbcQuery.Selected.Index = 0, strSQLSend, strSQLDel)
        strSQL = _
            " Select /*+ Rule*/ A.状态,A.NO,A.序号,I.名称 as 药房,H.名称 as 开嘱科室,A.姓名,e.住院号,LPAD(e.出院病床," & intBedLen & ",' ') as 床号," & _
            " Nvl(X.名称,F.名称)||Decode(F.产地,NULL,NULL,'('||F.产地||')')||Decode(F.规格,NULL,NULL,' '||F.规格) as 药品信息," & _
            " A.数量/Nvl(E.住院包装,1) as 数量,E.住院单位,A.单价*Nvl(E.住院包装,1) as 单价," & _
            " A.金额,A.期效,A.单量,G.计算单位 as 剂量单位,A.频次,A.用法,A.时间,A.人员,A.付数,g.类别" & _
            " From (" & strSQL & ") A,药品规格 E,收费项目目录 F,诊疗项目目录 G,部门表 H,部门表 I,收费项目别名 X,病案主页 E" & _
            " Where A.药品ID=E.药品ID And A.药品ID=F.ID And E.药名ID=G.ID" & _
            " And A.对方部门ID=H.ID And A.库房ID=I.ID And a.病人id=e.病人id and a.主页id=e.主页id" & _
            " And F.ID=X.收费细目ID(+) And X.码类(+)=1 And X.性质(+)=[11]" & _
            " Order by 床号,A.NO,A.序号"
    ElseIf tbcQuery.Selected.Index = 1 Or tbcQuery.Selected.Index = 3 Then '发药汇总、退药汇总
        strSQL = IIF(tbcQuery.Selected.Index = 1, strSQLSend, strSQLDel)
        strSQL = "Select B.药品ID,C.编码 as 药品编码,C.名称," & _
            " C.产地,C.规格,B.住院单位,Sum(A.数量/Nvl(B.住院包装,1)) as 数量,Sum(A.金额) as 金额" & _
            " From (" & strSQL & ") A,药品规格 B,收费项目目录 C" & _
            " Where A.药品ID=B.药品ID And A.药品ID=C.ID" & _
            " Group by B.药品ID,C.编码,C.名称,C.产地,C.规格,B.住院单位" & _
            " Having Sum(A.数量/Nvl(B.住院包装,1))<>0 Or Sum(A.金额)<>0"
    
        strSQL = "Select /*+ Rule*/ A.药品编码," & _
            " Nvl(B.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 药品信息," & _
            " A.住院单位,A.数量,A.金额" & _
            " From (" & strSQL & ") A,收费项目别名 B" & _
            " Where A.药品ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[11]" & _
            " Order by A.药品编码"
    ElseIf tbcQuery.Selected.Index = 4 Then '发退汇总
        strSQL = "Select 药品ID,数量 as 应发数,0 as 退药数,金额 From (" & strSQLSend & ")" & _
            " Union ALL Select 药品ID,0 as 应发数,数量 as 退药数,-1*金额 From (" & strSQLDel & ")"
            
        strSQL = "Select B.药品ID,C.编码 as 药品编码,C.名称,C.产地,C.规格,B.住院单位," & _
            " Sum(A.应发数/Nvl(B.住院包装,1)) as 应发数,Sum(A.退药数/Nvl(B.住院包装,1)) as 退药数," & _
            " (Sum(A.应发数)-Sum(A.退药数))/Nvl(B.住院包装,1) as 实发数,Sum(A.金额) as 金额" & _
            " From (" & strSQL & ") A,药品规格 B,收费项目目录 C" & _
            " Where A.药品ID=B.药品ID And A.药品ID=C.ID" & _
            " Group by B.药品ID,C.编码,C.名称,C.产地,C.规格,B.住院单位,Nvl(B.住院包装,1)" & _
            " Having Sum(A.应发数/Nvl(B.住院包装,1))<>0 Or Sum(A.退药数/Nvl(B.住院包装,1))<>0 Or Sum(A.金额)<>0"
    
        strSQL = "Select /*+ Rule*/ A.药品编码," & _
            " Nvl(B.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 药品信息," & _
            " A.住院单位,A.应发数,A.退药数,A.实发数,A.金额" & _
            " From (" & strSQL & ") A,收费项目别名 B" & _
            " Where A.药品ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[11]" & _
            " Order by A.药品编码"
    End If
    
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    With mvQuery
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .药房ID, .DateBegin, .DateEnd, _
            .NO, .期效, .病区ID, .病人IDs, .给药途径, .退药DateB, .退药DateE, IIF(gbyt药品名称显示 = 0, 1, 3), Val(.发药号), Val(.领药部门ID))
    End With
    
    If Not rsTmp.EOF Then
        With vsQuery
            .Redraw = flexRDNone
            .Rows = .FixedRows
            If tbcQuery.Selected.Index = 0 Or tbcQuery.Selected.Index = 2 Then '发药明细、退药明细
                For i = 1 To rsTmp.RecordCount
                    If strKey <> rsTmp!NO Then
                        If strKey <> "" Then
                            j = .FindRow(CStr(strKey))
                            If j <> -1 Then
                                .Cell(flexcpText, j, 0, j, .Cols - 1) = .TextMatrix(j, 0) & Format(curTotal, gstrDec)
                                .Cell(flexcpBackColor, j, 0, j, .Cols - 1) = &HEFF0EF   '单据头
                            End If
                        End If
                        
                        .AddItem ""
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = CStr(rsTmp!NO)
                        .TextMatrix(.Rows - 1, 0) = " 单据号:【" & rsTmp!NO & "】　科室:" & rsTmp!开嘱科室 & _
                            "　病人:【" & rsTmp!姓名 & "】　住院号:" & Nvl(rsTmp!住院号) & _
                            "　床号:" & Nvl(rsTmp!床号) & "　金额:"
                        curTotal = 0
                    End If
                    
                    .AddItem ""
                    
                    If tbcQuery.Selected.Index = 0 Then
                        .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!期效)
                        .TextMatrix(.Rows - 1, 1) = IIF(Nvl(rsTmp!状态, 0) = 0, "未发药", "已发药")
                        .TextMatrix(.Rows - 1, 2) = rsTmp!药品信息
                        If rsTmp!类别 & "" = "7" Then
                            .TextMatrix(.Rows - 1, 3) = rsTmp!付数 & ""
                        End If
                        .TextMatrix(.Rows - 1, 4) = FormatEx(rsTmp!数量, 5) & Nvl(rsTmp!住院单位)
                        .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!单价, 0), gstrDecPrice)
                        .TextMatrix(.Rows - 1, 6) = Format(Nvl(rsTmp!金额, 0), gstrDec)
                        .TextMatrix(.Rows - 1, 7) = IIF(Not IsNull(rsTmp!单量), FormatEx(Nvl(rsTmp!单量, 0), 5) & Nvl(rsTmp!剂量单位), "")
                        .TextMatrix(.Rows - 1, 8) = Nvl(rsTmp!频次)
                        .TextMatrix(.Rows - 1, 9) = Nvl(rsTmp!用法)
                        .TextMatrix(.Rows - 1, 10) = Format(Nvl(rsTmp!时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(.Rows - 1, 11) = Nvl(rsTmp!人员)
                    Else
                        .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!期效)
                        .TextMatrix(.Rows - 1, 1) = rsTmp!药品信息
                        If rsTmp!类别 & "" = "7" Then
                            .TextMatrix(.Rows - 1, 2) = rsTmp!付数 & ""
                        End If
                        .TextMatrix(.Rows - 1, 3) = FormatEx(rsTmp!数量, 5) & Nvl(rsTmp!住院单位)
                        .TextMatrix(.Rows - 1, 4) = Format(Nvl(rsTmp!单价, 0), gstrDecPrice)
                        .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!金额, 0), gstrDec)
                        .TextMatrix(.Rows - 1, 6) = IIF(Not IsNull(rsTmp!单量), FormatEx(Nvl(rsTmp!单量, 0), 5) & Nvl(rsTmp!剂量单位), "")
                        .TextMatrix(.Rows - 1, 7) = Nvl(rsTmp!频次)
                        .TextMatrix(.Rows - 1, 8) = Nvl(rsTmp!用法)
                        .TextMatrix(.Rows - 1, 9) = Format(Nvl(rsTmp!时间), "yyyy-MM-dd HH:mm")
                        .TextMatrix(.Rows - 1, 10) = Nvl(rsTmp!人员)
                    End If
                    
                    If Nvl(rsTmp!状态, 0) = 0 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC00000 '未发药
                    End If
                    curTotal = curTotal + Nvl(rsTmp!金额, 0)
                    
                    strKey = rsTmp!NO
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(2)
            ElseIf tbcQuery.Selected.Index = 1 Or tbcQuery.Selected.Index = 3 Then '发药汇总、退药汇总
                For i = 1 To rsTmp.RecordCount
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!药品编码)
                    .TextMatrix(.Rows - 1, 1) = rsTmp!药品信息
                    .TextMatrix(.Rows - 1, 2) = FormatEx(rsTmp!数量, 5) & Nvl(rsTmp!住院单位)
                    .TextMatrix(.Rows - 1, 3) = Format(Nvl(rsTmp!金额, 0), gstrDec)
                    
                    curTotal = curTotal + Nvl(rsTmp!金额, 0)
                    
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(1)
            ElseIf tbcQuery.Selected.Index = 4 Then  '发退汇总
                For i = 1 To rsTmp.RecordCount
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!药品编码)
                    .TextMatrix(.Rows - 1, 1) = rsTmp!药品信息
                    If Nvl(rsTmp!应发数, 0) <> 0 Then .TextMatrix(.Rows - 1, 2) = FormatEx(rsTmp!应发数, 5) & Nvl(rsTmp!住院单位)
                    If Nvl(rsTmp!退药数, 0) <> 0 Then .TextMatrix(.Rows - 1, 3) = FormatEx(rsTmp!退药数, 5) & Nvl(rsTmp!住院单位)
                    .TextMatrix(.Rows - 1, 4) = FormatEx(rsTmp!实发数, 5) & Nvl(rsTmp!住院单位)
                    .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!金额, 0), gstrDec)
                    
                    If Nvl(rsTmp!退药数, 0) <> 0 Then .Cell(flexcpForeColor, .Rows - 1, 3) = vbRed
                    
                    curTotal = curTotal + Nvl(rsTmp!金额, 0)
                    
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(1)
            End If
            
            '最后一个单头
            If InStr(",1,3,4,", tbcQuery.Selected.Index) > 0 Then
                .AddItem "", .FixedRows
                .MergeRow(.FixedRows) = True
                .RowData(.FixedRows) = 1
                .Cell(flexcpText, .FixedRows, 0, .FixedRows, .Cols - 1) = "金额合计:" & Format(curTotal, gstrDec)
                .Cell(flexcpBackColor, .FixedRows, 0, .FixedRows, .Cols - 1) = &HEFF0EF    '单据头
            ElseIf InStr(",0,2,", tbcQuery.Selected.Index) > 0 Then
                j = .FindRow(CStr(strKey))
                If j <> -1 Then
                    .Cell(flexcpText, j, 0, j, .Cols - 1) = .TextMatrix(j, 0) & Format(curTotal, gstrDec)
                    .Cell(flexcpBackColor, j, 0, j, .Cols - 1) = &HEFF0EF    '单据头
                End If
            End If
            
            Call SetMinRowHeight
            .Row = .FixedRows: .Col = 0
            Call vsQuery_AfterRowColChange(-1, -1, .Row, .Col)
            .Redraw = flexRDDirect
        End With
    End If
    Call zlCommFun.StopFlash
    LoadQueryData = True
    Exit Function
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetMinRowHeight()
'功能：用于AutoSize之后,将假显示的行高调整为最小高度
    Dim i As Long
    
    With vsQuery
        .Redraw = flexRDNone
        For i = 0 To .Rows - 1
            If .RowData(i) <> "" Then
                .RowHeight(i) = 300
            ElseIf .RowHeight(i) < .RowHeightMin Then
                .RowHeight(i) = .RowHeightMin
            End If
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Long
    If Not imgColSel.Visible Then Exit Sub
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                vsQuery.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsQuery.ColHidden(.RowData(i)) Or vsQuery.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub mfrmCond_DoQuery(ByVal 药房ID As Long, ByVal Mode As Byte, ByVal DateBegin As Date, ByVal DateEnd As Date, ByVal 退药DateB As Date, ByVal 退药DateE As Date, ByVal NO As String, ByVal 发药号 As String, ByVal 期效 As Integer, ByVal 状态 As String, ByVal 病区ID As Long, ByVal 病人IDs As String, ByVal 给药途径 As String, ByVal 领药部门ID As Long)
    mvQuery.药房ID = 药房ID
    mvQuery.Mode = Mode
    mvQuery.DateBegin = DateBegin
    mvQuery.DateEnd = DateEnd
    mvQuery.退药DateB = 退药DateB
    mvQuery.退药DateE = 退药DateE
    mvQuery.NO = NO
    mvQuery.发药号 = 发药号
    mvQuery.期效 = 期效
    mvQuery.状态 = 状态
    mvQuery.病区ID = 病区ID
    mvQuery.病人IDs = 病人IDs
    mvQuery.给药途径 = 给药途径
    mvQuery.领药部门ID = 领药部门ID
    
    '查询包含已发药的时，不显示退药数据
    If Mid(mvQuery.状态, 2, 1) = "1" Then
        If tbcQuery.Selected.Index >= 2 Then
            tbcQuery(0).Selected = True
        End If
        tbcQuery(2).Visible = False
        tbcQuery(3).Visible = False
        tbcQuery(4).Visible = False
    Else
        tbcQuery(2).Visible = True
        tbcQuery(3).Visible = True
        tbcQuery(4).Visible = True
    End If
    Call tbcQuery_SelectedChanged(tbcQuery.Selected)
    
    vsQuery.SetFocus
End Sub

Private Sub picQuery_Resize()
    With picQuery
        vsQuery.Top = picQuery.ScaleTop
        vsQuery.Left = picQuery.ScaleLeft
        vsQuery.Height = picQuery.ScaleHeight
        vsQuery.Width = picQuery.ScaleWidth
    End With
    fraColSel.Left = vsQuery.Left
    fraColSel.Top = vsQuery.Top

End Sub

Private Sub tbcQuery_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Visible Then
        Call SaveFlexState(vsQuery, App.ProductName & "\" & Me.Name)
    End If
    vsColumn.Visible = False
    
    vsQuery.Tag = Item.Index
    Call InitQueryTable
    
    If Item.Index = 0 Then
        fraColSel.Visible = True
        imgColSel.Visible = True
        Call InitColumnSelect
    Else
        fraColSel.Visible = False
        imgColSel.Visible = False
    End If
    
    If Visible Then
        Call RestoreFlexState(vsQuery, App.ProductName & "\" & Me.Name)
        Call LoadQueryData
    End If
    
    If Visible Then vsQuery.SetFocus
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long, lnPos As Long
    Dim strOldCOLInfo As String, strNewCOLInfo As String
    strOldCOLInfo = Mid(";" & mstrOldHead & ";", InStr(";" & mstrOldHead & ";", ";" & Trim(vsColumn.TextMatrix(Row, 1))))
    strOldCOLInfo = Mid(strOldCOLInfo, 2, InStr(2, strOldCOLInfo, ";") - 2)
    strNewCOLInfo = Mid(";" & mstrNewHead & ";", InStr(";" & mstrNewHead & ";", ";" & Trim(vsColumn.TextMatrix(Row, 1))))
    strNewCOLInfo = Mid(strNewCOLInfo, 2, InStr(2, strNewCOLInfo, ";") - 2)
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            If vsQuery.ColWidth(lngCol) = 0 Then
                mstrNewHead = Replace(mstrNewHead, Trim(vsColumn.TextMatrix(Row, 1)) & ",0," & Split(strNewCOLInfo, ",")(2), strOldCOLInfo)
                vsQuery.ColWidth(lngCol) = Val(Split(strOldCOLInfo, ",")(1))
            Else
                mstrNewHead = Replace(mstrNewHead, strNewCOLInfo, strOldCOLInfo)
                vsQuery.ColWidth(lngCol) = vsQuery.ColData(lngCol)
            End If
            vsQuery.ColHidden(lngCol) = False
        Else
            vsQuery.ColWidth(lngCol) = 0
            vsQuery.ColHidden(lngCol) = True
            mstrNewHead = Replace(mstrNewHead, strNewCOLInfo, Trim(vsColumn.TextMatrix(Row, 1)) & ",0," & Split(strOldCOLInfo, ",")(2))
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub vsQuery_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsQuery
        If OldRow <> NewRow Then
            If OldRow >= .FixedRows And OldRow <= .Rows - 1 Then
                If .RowData(OldRow) <> "" Then
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEFF0EF
                Else
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = .BackColor
                End If
            End If
            If NewRow >= .FixedRows And NewRow <= .Rows - 1 Then
                If .RowData(NewRow) <> "" Then
                    .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = &HEFF0EF
                Else
                    .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = &HFFCC99
                End If
            End If
        End If
    End With
End Sub

Private Sub vsQuery_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If vsQuery.TextMatrix(0, Col) = "药品信息" Then
        Call vsQuery.AutoSize(Col)
        Call SetMinRowHeight
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsQuery.TextMatrix(vsQuery.FixedRows - 1, Col) & "A")
        If vsQuery.ColWidth(Col) < lngW Then
            vsQuery.ColWidth(Col) = lngW
        ElseIf vsQuery.ColWidth(Col) > vsQuery.Width * 0.5 Then
            vsQuery.ColWidth(Col) = vsQuery.Width * 0.5
        End If
    End If
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '表头
    objOut.Title.Text = tbcQuery.Selected.Caption
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "病区：" & sys.RowValue("部门表", mvQuery.病区ID, "名称") & "/药房：" & sys.RowValue("部门表", mvQuery.药房ID, "名称")
    objRow.Add Format(mvQuery.DateBegin, "yyyy-MM-dd HH:mm") & "/" & Format(mvQuery.DateEnd, "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsQuery
    
    '输出
    vsQuery.Redraw = False
    lngRow = vsQuery.Row: lngCol = vsQuery.Col
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsQuery.Row = lngRow: vsQuery.Col = lngCol
    vsQuery.Redraw = True
End Sub

Private Sub InitColumnSelect()
'功能：根据发药明细清单原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    vsColumn.Rows = vsColumn.FixedRows
    With vsQuery
        For i = .FixedCols To .Cols - 1
            If .TextMatrix(0, i) <> "" Then
                vsColumn.Rows = vsColumn.Rows + 1
                lngRow = vsColumn.Rows - 1
                vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                vsColumn.RowData(lngRow) = i
                If vsQuery.ColHidden(i) Then
                    vsColumn.TextMatrix(lngRow, 0) = 0
                End If
                '固定显示列
                If InStr(",药品信息,数量,,单量,频次,用法,", "," & .TextMatrix(0, i) & ",") > 0 Then
                    vsColumn.TextMatrix(lngRow, 0) = 1
                    vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                End If
            End If
        Next
    End With
    If vsColumn.Rows > 1 Then vsColumn.Row = 1
End Sub




