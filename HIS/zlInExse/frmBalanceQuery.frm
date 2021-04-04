VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceQuery 
   Caption         =   "病人结帐费用查询"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmBalanceQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11715
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   90
      Width           =   510
      Begin VB.Label lblCancel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5355
      ScaleHeight     =   2295
      ScaleWidth      =   3705
      TabIndex        =   4
      Top             =   3255
      Width           =   3705
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1845
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   1320
      Left            =   1740
      TabIndex        =   3
      Top             =   3705
      Width           =   2775
      _Version        =   589884
      _ExtentX        =   4895
      _ExtentY        =   2328
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1425
      ScaleHeight     =   420
      ScaleWidth      =   5550
      TabIndex        =   1
      Top             =   2565
      Width           =   5550
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2460
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名: XXX"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   105
         Width           =   1560
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      SimpleText      =   $"frmBalanceQuery.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalanceQuery.frx":05D1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1260
      Top             =   1230
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBalanceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gViewType
    g_Ed_结帐表 = 0
    g_Ed_明细表 = 1
    g_Ed_项目明细 = 2
    g_Ed_分类表 = 3
    g_Ed_分月表 = 4
    g_Ed_费目表 = 5
    g_Ed_逐日单据 = 6
    g_Ed_逐日费用 = 7
End Enum

Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrPopupMain As CommandBar, mcbrMenuView As CommandBarPopup, mcbrRefresh As CommandBarControl
Private mcbrCmb As CommandBarComboBox
Private mBalanceType As gBalanceBill, mlng结帐ID As Long, mstrPrivs As String, mlngModule As Long
Private mblnDateMoved As Boolean, mViewType As gViewType
Private mlng病人ID As Long
Private mstrTime As String  '病人结帐次数(初始="",可以为"1,2,3...")
Private mdtBeginDate As Date       '病人结帐的开始时间,初始为'1900-01-01'
Private mdtEndDate As Date         '病人结帐的结束时间,初始为'3000-01-01'
Private mstrDeptIDs As String      '病人结帐科室ID串(初始="",可以为"0,1,2,3...",0表示开单部门ID为空)
Private mstrClass As String       '费用类型=""-所有费用(含未设置),"'公费','比例',..."
Private mstrChargeType As String      '收费类别 '34260
Private mstrBaby As String      '是否仅结算婴儿费用(0-所有费用,1-病人费用,2及以上-第mbytbaby-1个婴儿费用)
Private mstrItem As String      '要结的收据费目
Private mbytKind As Byte       '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
Private mblnCurBalanceOwnerFee As Boolean      '当前是否正在结“自费费用”
Private mclsCon As clsBalanceCon


Public Function ShowMe(ByVal frmMain As Object, BalanceType As gBalanceBill, ByRef clsCon As clsBalanceCon, ByVal lng结帐ID As Long, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal ViewType As gViewType = g_Ed_结帐表) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐费用查询的程序入口
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-30 10:38:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mclsCon = clsCon
    If Not mclsCon Is Nothing Then
        With mclsCon
            mstrTime = .strTime
            mlng病人ID = .lng病人ID
            mdtBeginDate = .dtBeginDate
            mdtEndDate = .dtEndDate
            mstrDeptIDs = .strDeptIDs
            mstrClass = .strClass
            mstrChargeType = .strChargeType
            mstrBaby = .strBaby
            mstrItem = .strItem
            mbytKind = .bytKind
            mblnCurBalanceOwnerFee = .blnCurBalanceOwnerFee
        End With
    End If
    mBalanceType = BalanceType
    mViewType = ViewType
    mlng结帐ID = lng结帐ID
    mstrPrivs = strPrivs
    mlngModule = lngModule
    Me.Show vbModal, frmMain
    ShowMe = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim intPara As Integer
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '初始化设置
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
        Set .Font = vsfMain.Font
    End With
    
    cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    
    '-----------------------------------------------------
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ModifyStyle &H400000, 0
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
            mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    zlDefCommandBars = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘尔旋
    '日期:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    Select Case tabMain.Selected.Index
        Case 0
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐表"
        Case 1
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐明细表"
        Case 2
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐项目明细表"
        Case 3
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐分类表"
        Case 4
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐分月表"
        Case 5
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐费目表"
        Case 6
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐逐日单据表"
        Case 7
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐逐日费用表"
    End Select
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    Set objRow = New zlTabAppRow
    objRow.Add lblInfo.Caption
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    If vsBill Is Nothing Then Exit Sub
    '由于打印控件不能识别列隐藏属性
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
    If bytFunc = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '恢复
    With vsBill
        For i = 0 To .Cols - 1
           If .ColHidden(i) = True Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000C
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Public Sub UnloadForm()
    Unload Me
End Sub

Private Sub Form_Load()
    stbThis.Panels(3).Text = UserInfo.姓名
    Call zlDefCommandBars
    Call SetTabControl
    Call InitInfo
    tabMain.Item(mViewType).Selected = True
    Call LoadCardData(tabMain.Selected.Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            '取消按钮
            Unload Me
    End Select
End Sub

Private Sub InitInfo()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If mlng结帐ID = 0 Then
        lblInfo.Caption = ""
        If mclsCon Is Nothing Then Exit Sub
        If mclsCon.lng病人ID <> 0 Then
            strSQL = "Select 姓名,性别,出生日期,年龄,门诊号,住院号 From 病人信息 Where 病人ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mclsCon.lng病人ID)
            If rsTmp.EOF Then Exit Sub
            lblInfo.Caption = "病人姓名:" & NVL(rsTmp!姓名) & "  性别:" & NVL(rsTmp!性别) & "   出生日期:" & NVL(rsTmp!出生日期) & "   年龄:" & NVL(rsTmp!年龄) & "   门诊号:" & NVL(rsTmp!门诊号) & "    住院号:" & NVL(rsTmp!住院号)
        End If
    Else
        lblInfo.Caption = ""
        strSQL = _
            " Select a.姓名, a.性别, a.出生日期, a.年龄, a.门诊号, a.住院号" & vbNewLine & _
            " From 病人信息 A, 病人预交记录 B" & vbNewLine & _
            " Where b.结帐id = [1] And b.病人id = a.病人id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
        If rsTmp.EOF Then Exit Sub
        lblInfo.Caption = "病人姓名:" & NVL(rsTmp!姓名) & "  性别:" & NVL(rsTmp!性别) & "   出生日期:" & NVL(rsTmp!出生日期) & "   年龄:" & NVL(rsTmp!年龄) & "   门诊号:" & NVL(rsTmp!门诊号) & "    住院号:" & NVL(rsTmp!住院号)
    End If
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建TAB控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabMain
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        Set .PaintManager.Font = vsfMain.Font
        .InsertItem 1, "结帐表", picMain.hWnd, 0
        .InsertItem 2, "明细表", picMain.hWnd, 0
        .InsertItem 3, "项目明细", picMain.hWnd, 0
        .InsertItem 4, "分类表", picMain.hWnd, 0
        .InsertItem 5, "分月表", picMain.hWnd, 0
        .InsertItem 6, "费目表", picMain.hWnd, 0
        .InsertItem 7, "逐日单据", picMain.hWnd, 0
        .InsertItem 8, "逐日费用", picMain.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.StaticFrame = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With picInfo
        .Left = Left
        .Top = Top
        .Width = Right - Left
        Line1.X2 = .Left + .Width
    End With
    With tabMain
        .Left = picInfo.Left
        .Top = picInfo.Top + picInfo.Height + 15
        .Width = Right - Left
        .Height = Bottom - .Top
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picCancel.Left = Me.Width - picCancel.Width - 300
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With vsfMain
        .Left = 0
        .Top = 0
        .Height = picMain.Height
        .Width = picMain.Width
    End With
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call LoadCardData(Item.Index)
End Sub

Private Function LoadCardData(ByVal intIndex As Integer) As Boolean
'功能：根据当前选择的病人费用项目卡片，读取并设置费用清单
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim strInfo As String, strPre As String
    Dim strMoney As String, strTmp As String, strTmpSql As String
    Dim arrTotal() As Currency
    Dim intCol As Integer, blnZero As Boolean
    Dim strCond As String, bytType As Byte '0-门诊;1-住院;2-门诊和住院
    Dim DateBegin As Date, DateEnd As Date
    Dim strTable As String, strTimeRange As String
    Dim strOriginal As String, strTmpOriginal As String
    On Error GoTo errH

    strPre = stbThis.Panels(2).Text
    stbThis.Panels(2).Text = "正在读取数据,请稍候 ……"
    Screen.MousePointer = 11
    vsfMain.Redraw = False
    Me.Refresh
    
    If mBalanceType = g_Ed_门诊结帐 Or mBalanceType = g_Ed_住院结帐 Then
        blnZero = zlDatabase.GetPara("处理零费用", glngSys, 1137) = "1"
        strCond = ""
        strCond = strCond & IIf(mstrTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
        If mdtBeginDate <> CDate("0:00:00") Then
            strTimeRange = " And " & IIf(gint费用时间 = 0, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
            DateBegin = CDate(Format(mdtBeginDate, "yyyy-MM-dd 00:00:00"))
            DateEnd = CDate(Format(mdtEndDate, "yyyy-MM-dd 23:59:59"))
        End If
        strCond = strCond & IIf(mstrDeptIDs = "", "", " And Instr([5],','||A.开单部门ID||',')>0")
        strCond = strCond & IIf(mstrBaby = "", "", " And Instr([6],','|| Nvl(A.婴儿费,0) ||',')>0")
        strCond = strCond & IIf(mstrItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
        
        If mbytKind = 1 Then
            strCond = strCond & " And A.门诊标志=4"
        Else
            If InStr(mstrPrivs, ";住院费用结帐;") = 0 Then strCond = strCond & " And A.门诊标志<>2"
            If InStr(mstrPrivs, ";门诊费用结帐;") = 0 Then strCond = strCond & " And A.门诊标志<>1"
            If mbytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
        End If
        
        bytType = IIf(mBalanceType = g_Ed_门诊结帐, 0, 1)
        
        strSQL = _
        " Select NO,Mod(记录性质,10) as 记录性质, Nvl(Sum(实收金额),0) as 实收金额,Nvl(Sum(结帐金额),0) as 结帐金额,序号 " & _
        " From 住院费用记录 A" & _
        " Where 记录状态<>0 And 记帐费用=1 " & strCond & _
        "       And 病人ID=[1]" & _
        " Group by NO,Mod(记录性质,10),序号 " & _
        IIf(blnZero, "", " Having   Nvl(Sum(实收金额),0)-Nvl(Sum(结帐金额),0)<>0 ")
        
        strSQL = _
            " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,Max(A.登记时间) As 登记时间,A.NO,Decode(a.婴儿费, 1, '√', Null) As 婴儿,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
            "        Sum(Decode(Floor(a.记录性质 / 10),0,a.数次,0)) As 数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型,Sum(a.应收金额) As 应收金额" & _
            " From 住院费用记录 A,(" & strSQL & ") B" & _
            " Where A.NO=B.NO And A.结帐ID Is Not Null And Mod(A.记录性质,10)=B.记录性质" & _
            "       And A.记录状态<>0 And A.记帐费用=1 And A.序号=B.序号 " & _
            "       And A.病人ID+0=[1] And Not Exists (Select 1 From 住院费用记录 C, 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond & strTimeRange & _
            "" & _
            " Group by Mod(A.记录性质,10),A.发生时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID," & _
            "       A.收据费目,A.开单部门ID,A.计算单位,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型,a.婴儿费 " & _
            " Having    " & vbNewLine & _
            "        Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)) <> 0 Or (Sum(Nvl(a.实收金额, 0)) = 0 And Sum(Nvl(a.应收金额, 0)) = 0 And Sum(Nvl(a.结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or" & vbNewLine & _
            "             (Sum(Nvl(a.实收金额, 0)) = 0 And Sum(Nvl(a.应收金额, 0)) <> 0 And Sum(Nvl(a.结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or" & vbNewLine & _
            "             Sum(Nvl(a.结帐金额, 0)) = 0 And Sum(Nvl(a.应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0 " & _
            " Union All " & _
            " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,Max(A.登记时间) As 登记时间,A.NO,Decode(a.婴儿费, 1, '√', Null) As 婴儿,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
            "        Sum(Decode(Floor(a.记录性质 / 10),0,a.数次,0)) As 数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型,Sum(a.应收金额) As 应收金额" & _
            " From 住院费用记录 A,(" & strSQL & ") B" & _
            " Where A.NO=B.NO And Mod(A.记录性质,10)=B.记录性质" & _
            "       And A.记录状态<>0 And A.记帐费用=1 And A.序号=B.序号" & _
            "       And A.病人ID+0=[1] And A.结帐ID Is Null " & strCond & strTimeRange & _
            "" & _
            " Group by Mod(A.记录性质,10),A.发生时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID," & _
            "       A.收据费目,A.开单部门ID,A.计算单位,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型,a.婴儿费 "

        If mblnDateMoved Then
            strSQL = strSQL & " Union All " & Replace(strSQL, "住院费用记录", "H住院费用记录")
        End If
        
        Select Case bytType
        Case 0 '门诊
            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
            
            strTmpSql = _
            " Select NO,Mod(记录性质,10) as 记录性质, Nvl(Sum(实收金额),0) as 实收金额,Nvl(Sum(结帐金额),0) as 结帐金额" & _
            " From 住院费用记录 A" & _
            " Where 记录状态<>0 And 记帐费用=1 And Mod(记录性质,10)=5 And 主页ID Is Null " & strCond & strTimeRange & _
            "       And 病人ID=[1]" & _
            " Group by NO,Mod(记录性质,10) " & _
            " "
            
            strTmpSql = _
            " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,Max(A.登记时间) As 登记时间,A.NO,Decode(a.婴儿费, 1, '√', Null) As 婴儿,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
            "        Sum(Decode(Floor(a.记录性质 / 10),0,a.数次,0)) As 数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型,Sum(a.应收金额) As 应收金额" & _
            " From 住院费用记录 A,(" & strTmpSql & ") B" & _
            " Where A.NO=B.NO And Mod(A.记录性质,10)=B.记录性质" & _
            "       And A.记录状态<>0 And A.记帐费用=1 And Mod(A.记录性质,10)=5 And A.主页ID Is Null " & _
            "       And A.病人ID+0=[1] " & strCond & strTimeRange & _
            " " & _
            " Group by Mod(A.记录性质,10),A.发生时间,A.登记时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID," & _
            "       A.收据费目,A.开单部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型,a.婴儿费 "
                If mblnDateMoved Then
                    strTmpSql = strTmpSql & " Union All " & Replace(strTmpSql, "住院费用记录", "H住院费用记录")
                End If
                strTmpSql = Replace(strTmpSql, " And Instr([2],','||Nvl(A.主页ID,0)||',')>0", "")
                strSQL = strSQL & " Union All " & strTmpSql
        Case 1 '住院
        Case Else
            '门诊和住院
             strSQL = strSQL & " Union All " & Replace(Replace(strSQL, "住院费用记录", "门诊费用记录"), " And Instr([2],','||Nvl(A.主页ID,0)||',')>0", "")
        End Select
            
        strTable = "(" & strSQL & ") "
        
        '未结费用清单
        Select Case intIndex
            Case 0
                strSQL = _
                "Select To_Char(A.登记时间,'YYYY-MM-DD') As 时间, a.No, d.名称 As 项目, a.收据费目 As 费目, a.婴儿 As 婴儿, Ltrim(To_Char(Nvl(a.实收金额,0) - Nvl(a.结帐金额,0),'999999999" & gstrDec & "')) As 未结金额" & vbNewLine & _
                "From (" & strTable & vbNewLine & _
                "       ) A, 收费项目目录 D" & vbNewLine & _
                "Where d.Id = a.收费细目id " & _
                       IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(D.费用类型,'未知')||''',')>0") & _
                       IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(D.类别,'无')||''',')>0") & _
                "Order By No,费目 "
                strMoney = "4,4,1,1,1,7"
            Case 1 '明细清单
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')|| To_Char(A.数次,'999999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Nvl(A.应收金额,0),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 未结金额,A.操作员姓名 as 操作员" & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Order by 发生日期,单据号,费目"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 2 '分项目明细
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ') 规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')|| To_Char(A.数次,'999999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Nvl(A.应收金额,0),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 未结金额," & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间" & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1)
                
               strSQL = strSQL & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 开单科室," & _
                "       Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ')||'ZZZZZ' as 规格,NULL,to_char(sum(Nvl(A.数次,1)*Nvl(A.付数,1)), '999999990.9')||' '||A.计算单位 as 数量,NULL as 标准单价," & _
                "       Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额," & _
                "       NULL as 类型,NULL as 操作员,NULL as 登记时间" & _
                " FROM " & strTable & " A,收费项目目录 C,收费项目别名 D" & _
                " Where A.收费细目ID=C.ID And A.收费细目ID=D.收费细目ID(+)" & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                "              And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Group by Nvl(D.名称,C.名称),C.规格,A.计算单位" & _
                " Order by 项目,规格,发生日期,单据号"
                
                strMoney = "4,4,1,1,1,1,1,7,7,7,1,1,1"
            Case 3 '分类明细
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||To_Char(A.数次,'999999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Nvl(A.应收金额,0),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 未结金额,A.操作员姓名 as 操作员 " & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 科室,NULL as 项目,Null as 规格,A.收据费目||'ZZZZZ' as 费目," & _
                "        NULL as 数量,NULL as 标准单价," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额,NULL as 操作员" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.收据费目||'ZZZZZ'" & _
                " Order by 费目,发生日期,单据号"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 4 '分月清单
                strSQL = _
                " SELECT B.期间,A.收据费目 as 费目," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额" & _
                "        FROM " & strTable & " A,期间表 B,收费项目目录 C" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                "       And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by B.期间,A.收据费目" & _
                " Union All" & _
                " SELECT B.期间||'ZZZZZ',NULL as 费目," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,期间表 B,收费项目目录 C" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                "       And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by B.期间||'ZZZZZ'" & _
                " Order by 期间,费目"
                strMoney = "4,4,7,7"
                
            Case 5 '费目
                strSQL = _
                " SELECT A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.收据费目 Order by 费目"
                strMoney = "4,7,7"
            Case 6 '逐日单据
                strSQL = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.收据费目 as 费用项目," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额," & _
                "        A.操作员姓名 as 操作员,A.记录性质" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.记录性质,TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO,A.收据费目,A.操作员姓名"
                strSQL = strSQL & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO||'ZZZZZ' as 单据号,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额,NULL as 操作员,A.记录性质" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " " & _
                " Group by A.记录性质,TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 单据号,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额,NULL as 操作员,-1" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " " & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,记录性质 desc,单据号,费用项目"
                
                strMoney = "4,4,4,7,7,1"
            Case 7 '逐日费用
                strSQL = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.收据费目 as 费用项目," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.收据费目" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Nvl(A.应收金额,0)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " " & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,费用项目"
                strMoney = "4,4,7,7"
        End Select
                
        vsfMain.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, "," & mstrTime & ",", DateBegin, DateEnd, _
                    "," & mstrDeptIDs & ",", "," & mstrBaby & ",", "," & mstrItem & ",", "," & mstrClass & ",", "," & mstrChargeType & ",")
        If rsTmp.RecordCount > 0 Then
            Set vsfMain.DataSource = rsTmp
        Else
            Call Grid.BandRec(vsfMain, rsTmp)
        End If
        
        
        vsfMain.Tag = intIndex
        For i = 0 To vsfMain.Cols - 1
            vsfMain.MergeCol(i) = False
        Next
        
        '求合计(小计)
        Select Case intIndex
            Case 0
                For i = 1 To vsfMain.Rows - 1
                    vsfMain.TextMatrix(i, 5) = Format(vsfMain.TextMatrix(i, 5), gstrDec)
                Next i
            Case 1, 3  '明细清单、分类明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 5)
                            For j = 0 To 7
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小 计:" & Left(strTmp, Len(strTmp) - 5)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 2 '分项目明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 3)
                            For j = 0 To 5
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小 计:" & strTmp
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            vsfMain.TextMatrix(i, 7) = " " '单价列
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 4 '分月清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            For j = 0 To 1
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小计:" & vsfMain.TextMatrix(i - 1, 0)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 2 To vsfMain.Cols - 1
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 5 '分类清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If IsNumeric(vsfMain.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 1))
                        If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 2))
                        vsfMain.MergeRow(i) = False
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.Col = 0: vsfMain.CellAlignment = 4
                    vsfMain.TextMatrix(vsfMain.Row, 0) = "合 计"
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 1) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 6 '逐日单据
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 1) Like "*ZZZZZ") And Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 3))
                            If IsNumeric(vsfMain.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 4))
                            vsfMain.MergeRow(i) = False
                        Else
                            If vsfMain.TextMatrix(i, 1) Like "*ZZZZZ" Then
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "小计:" & vsfMain.TextMatrix(i - 1, 1)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            Else
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 0 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "小计:" & vsfMain.TextMatrix(i - 1, 0)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 2
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 4) = Format(arrTotal(1), " " & gstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                    For i = 0 To vsfMain.Cols - 1
                        If vsfMain.TextMatrix(0, i) = "记录性质" Then vsfMain.ColHidden(i) = True
                    Next i
                End If
            Case 7 '逐日费目
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(1)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.MergeRow(i) = True
                            vsfMain.Row = i
                            vsfMain.Col = 1: vsfMain.CellAlignment = 4
                            vsfMain.TextMatrix(i, 0) = "小计:" & vsfMain.TextMatrix(i - 1, 0)
                            vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            vsfMain.TextMatrix(i, 1) = vsfMain.TextMatrix(i, 0)
                            For j = 2 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                
                    '删除只有一行费用的小计行
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    Else
        strSQL = "Select 发生时间,登记时间,NO,收据费目,费用类型,付数,数次,计算单位,标准单价,结帐金额,操作员姓名,开单部门ID,收费细目ID,结帐ID From 住院费用记录  where 结帐ID= [1]  Union ALL " & _
                 "Select 发生时间,登记时间,NO,收据费目,费用类型,付数,数次,计算单位,标准单价,结帐金额,操作员姓名,开单部门ID,收费细目ID,结帐ID From 门诊费用记录  where 结帐ID= [1]"
        
        If mblnDateMoved Then
            strSQL = Replace(Replace(strSQL, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
        End If
        strSQL = "(" & strSQL & ")"
        
        '读取结帐单时,点结帐分类明细
        Select Case intIndex
            Case 0
                strSQL = _
                " Select Trunc(登记时间) As 日期,A.NO as 单据号,C.名称 as 项目名称," & _
                "       A.收据费目 as 费目," & _
                "       Null as 婴儿费," & _
                "       " & _
                "       " & _
                "       Ltrim(To_Char(A.结帐金额,'999999999" & gstrDec & "')) as 结帐金额" & _
                " From " & strSQL & " A,部门表 B,收费项目目录 C" & _
                " Where A.开单部门ID = B.ID(+) And A.收费细目ID=C.ID" & _
                "       " & _
                " Order by 单据号,费目"
                strMoney = "4,4,1,1,1,7"
            Case 1 '明细
                '发生日期,单据号,科室,项目,费目,数量,单价,应收金额,结帐金额,操作员
                strSQL = _
                " Select To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       Nvl(B.名称,'未知') as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')|| To_Char(A.数次,'999999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(A.标准单价,'99999" & gstrFeePrecisionFmt & "')) as 单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(A.结帐金额,'999999999" & gstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员" & _
                " From " & strSQL & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID = B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by 发生日期,单据号,费目"
                
                '清单格式控制
               strMoney = "4,4,1,1,1,4,1,7,7,7,1"
            Case 2 '分项目明细
                '发生日期,单据号,科室,项目,规格,费目,数量,单价,应收金额,结帐金额,类型,操作员
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ') as 规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||To_Char(A.数次,'99999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 结帐金额," & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间" & _
                " FROM " & strSQL & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ')||'ZZZZZ' as 规格," & _
                "        NULL as 费目,to_char(sum(Nvl(A.数次,1)*Nvl(A.付数,1)),'99999990.9')||' '||A.计算单位 as 数量,NULL as 标准单价," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额,NULL as 类型,NULL as 操作员,NULL as 登记时间" & _
                " FROM " & strSQL & " A,收费项目目录 C,收费项目别名 D" & _
                " Where A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Group by Nvl(D.名称,C.名称),C.规格,A.计算单位" & _
                " Order by 项目,规格,发生日期,单据号"
                strMoney = "4,4,1,1,1,4,1,7,7,7,1,1,1"
            Case 3 '分类明细
                strSQL = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||To_Char(A.数次,'999999990.9')||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Nvl(A.结帐金额,0),'999999999" & gstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员 " & _
                " FROM " & strSQL & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And 码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 科室,NULL as 项目,Null as 规格,A.收据费目||'ZZZZZ' as 费目," & _
                "       NULL as 数量,NULL as 标准单价," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额,NULL as 操作员" & _
                " FROM " & strSQL & " A,部门表 B,收费项目目录 C" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                " Group by A.收据费目||'ZZZZZ' " & _
                " Order by 费目,发生日期,单据号"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 4 '分月清单
                strSQL = _
                " SELECT B.期间,A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额" & _
                " FROM " & strSQL & " A,期间表 B" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                " Group by B.期间,A.收据费目" & _
                " Union All" & _
                " SELECT B.期间||'ZZZZZ',NULL as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额" & _
                " FROM " & strSQL & " A,期间表 B" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                " Group by B.期间||'ZZZZZ'" & _
                " Order by 期间,费目"
                strMoney = "4,4,7,7"
            Case 5 '分类清单
                strSQL = _
                " SELECT A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额" & _
                " FROM " & strSQL & " A" & _
                " Group by A.收据费目 Order by 费目"
                strMoney = "4,7,7"
            Case 6 '逐日单据
                strSQL = _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.收据费目 as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员 " & _
                    " FROM " & strSQL & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO,A.收据费目,A.操作员姓名" & _
                    " Union All" & _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO||'ZZZZZ' as 单据号,NULL as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额, NULL as 操作员  " & _
                    " FROM " & strSQL & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO" & vbCrLf & _
                    " Union All" & _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,'ZZZZZAAAAA' as 单据号,NULL as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额,NULL as 操作员 " & _
                    " FROM  " & strSQL & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                    " Order by 发生日期,单据号,费用项目"
                strMoney = "4,4,4,7,7,1"
            Case 7 '逐日费目
                strSQL = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.收据费目 as 费用项目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额" & _
                " FROM " & strSQL & " A " & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.收据费目" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & gstrDec & "')) as 应收金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & gstrDec & "')) as 结帐金额" & _
                " FROM " & strSQL & " A" & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,费用项目"
                strMoney = "4,4,7,7"
        End Select
        
        vsfMain.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
        If rsTmp.RecordCount > 0 Then
            Set vsfMain.DataSource = rsTmp
        Else
            Call Grid.BandRec(vsfMain, rsTmp)
        End If

        vsfMain.Tag = intIndex
        For i = 0 To vsfMain.Cols - 1
            vsfMain.MergeCol(i) = False
        Next

        Select Case intIndex
            Case 1, 3  '明细清单、分类明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 5)
                            For j = 0 To 7
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小 计:" & Left(strTmp, Len(strTmp) - 5)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 2 '分项目明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 3)
                            For j = 0 To 5
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小 计:" & strTmp
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            vsfMain.TextMatrix(i, 7) = " " '单价列
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
             Case 4 '分月清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            For j = 0 To 1
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "小计:" & vsfMain.TextMatrix(i - 1, 0)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 2 To vsfMain.Cols - 1
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                End If
             Case 5 '分类清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If IsNumeric(vsfMain.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 1))
                        If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 2))
                        vsfMain.MergeRow(i) = False
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.Col = 0: vsfMain.CellAlignment = 4
                    vsfMain.TextMatrix(vsfMain.Row, 0) = "合 计"
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 1) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 6
                For i = 0 To vsfMain.Cols - 1
                    vsfMain.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 1) Like "*ZZZZZ") And Not (vsfMain.TextMatrix(i, 1) Like "*AAAAA") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 3))
                            If IsNumeric(vsfMain.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 4))
                            vsfMain.MergeRow(i) = False
                        Else
                            If Not (vsfMain.TextMatrix(i, 1) Like "*AAAAA") Then
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "单据小计:" & vsfMain.TextMatrix(i - 1, 1)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            Else
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "日小计"
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 2
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 4) = Format(arrTotal(1), " " & gstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 7
                For i = 0 To vsfMain.Cols - 1
                    vsfMain.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not vsfMain.TextMatrix(i, 0) Like "*ZZZZZ" Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = False
                            vsfMain.Col = 0: vsfMain.CellAlignment = 4
                            vsfMain.TextMatrix(i, 0) = Left(vsfMain.TextMatrix(i, 0), Len(vsfMain.TextMatrix(i, 0)) - 5)
                            vsfMain.TextMatrix(i, 1) = "日小计"
                            vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "合 计"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    End If
    
    '总的格式控制
    If vsfMain.Rows = 1 Then vsfMain.Rows = 2
    
    For i = 0 To vsfMain.Cols - 1
        If vsfMain.TextMatrix(0, i) = "结帐金额" Then intCol = i
        vsfMain.FixedAlignment(i) = 4
    Next
    
'    lblCancel.Visible = True
    picCancel.Visible = False
    vsfMain.RowHeight(0) = 350
    For i = 1 To vsfMain.Rows - 1
'        If Val(vsfMain.TextMatrix(i, intCol)) < 0 Then vsfMain.TextMatrix(i, intCol) = Format(-1 * vsfMain.TextMatrix(i, intCol), gstrDec): picCancel.Visible = True
        vsfMain.RowHeight(i) = 300
    Next
    
    '如果取了,由于没有设置初始列宽,打印会异常
'    Call SetGridWidth(vsfMain, Me)
    
    '有个记录性质列
    If intIndex = 6 And mBalanceType = g_Ed_门诊结帐 And mBalanceType = g_Ed_住院结帐 Then
        vsfMain.ColWidth(vsfMain.Cols - 1) = 0
    End If
    
    For i = 0 To UBound(Split(strMoney, ","))
        vsfMain.ColAlignment(i) = Split(strMoney, ",")(i)
    Next
    
'    vsfMain.Row = 1: vsfMain.Col = 0
    
    stbThis.Panels(2).Text = strPre
    
    vsfMain.Redraw = True
    vsfMain.Refresh
    Screen.MousePointer = 0
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    vsfMain.Redraw = True
    If ErrCenter() = 1 Then
        vsfMain.Redraw = False
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    stbThis.Panels(2).Text = strPre
End Function
