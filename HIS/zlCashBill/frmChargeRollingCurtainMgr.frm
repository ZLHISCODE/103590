VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChargeRollingCurtainMgr 
   Caption         =   "收费轧帐管理"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmChargeRollingCurtainMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   15105
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      ScaleHeight     =   255
      ScaleWidth      =   3495
      TabIndex        =   5
      Top             =   1425
      Visible         =   0   'False
      Width           =   3495
      Begin zL9CashBill.ComboxExpend cboType 
         Height          =   255
         Left            =   750
         TabIndex        =   6
         Top             =   0
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   450
         Appearance      =   0
         BorderStyle     =   1
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
         Locked          =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "轧帐类别"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   210
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   1
      Top             =   2790
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   2
         Top             =   -15
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBalanceList 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1005
      ScaleHeight     =   1695
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   705
      Width           =   3210
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1800
         Left            =   270
         TabIndex        =   4
         Top             =   180
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeRollingCurtainMgr.frx":0442
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   10065
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      SimpleText      =   $"frmChargeRollingCurtainMgr.frx":04BC
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeRollingCurtainMgr.frx":0503
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18997
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "刘兴洪"
            TextSave        =   "刘兴洪"
            Object.ToolTipText     =   "当前操作员:刘兴洪"
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
      Left            =   15
      Top             =   -75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmChargeRollingCurtainMgr.frx":0D97
      Left            =   660
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeRollingCurtainMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
'Private mobjComBar As CommandBarComboBox
Private mfrmChargeRollingCurtain As frmChargeRollingCurtain
Private mfrmHistory As frmChargeRollingCurtainHistory
Private mstr轧帐人员性质 As String
Private mstrPrevRollingType As String
Private Enum mPgIndex
    EM_PG_轧帐列表 = 250101
    EM_PG_历史列表 = 250102
End Enum
Private Enum mPaneIndex
    EM_PN_暂存金 = 1
    EM_PN_明细列表 = 2
End Enum
Private mstrPreDate As String   '上次轧帐时间
Private mblnFirst As Boolean, mblnNotice As Boolean
Private mbln预交分别轧帐 As Boolean

Private Function GetPreRollingCurtainTime(ByVal strTYPE As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取上次轧帐时间
    '入参:intType-(0-所有类别(按全额轧帐),1-收费,2-预交(21-门诊预交,22-住院预交),3-结帐,4-挂号,5-就诊卡)
    '返回:返回格式yyyy-mm-dd hh24:mi:ss
    '编制:刘兴洪
    '日期:2015-03-03 10:43:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '规则:
    '    1.如果当前按指定类别轧帐时,则按下规则处理(以终止时间为准):
    '      1) 如果存在有效的轧帐记录时
    '          a)如果当前轧帐人员不存在所有类别的轧帐记录的,则以最后一次轧帐时间为准
    '          b)如果当前轧帐人员存在所有类别的轧帐记录且最后一次轧帐记录的终止时间>当前类别的最后一次轧帐记录的终卡时间的,则以所有类别的最后一次轧帐记录的终止时间为准
    '          c)如果当前轧帐人员存在所有类别的轧帐记录且最后一次轧帐记录的终止时间<当前类别的最后一次轧帐记录的终卡时间的,则以当前类别的最后一次轧帐记录的终卡时间为准
    '      2)如果不存在有效轧帐记录时
    '          a)如果当前轧帐人员存在按所有类别轧帐的有效记录时,则以最后一次轧帐记录的终止时间为准
    '          b)如果不存在按所有类别轧帐的有效记录时,按下面的3以下规则处理
    '    2.如果当前按所有类别轧帐时,则按以下规则处理
    '          a)如果存在按所有类别轧帐的有效记录时,则以最后一次轧帐记录的终止时间为准
    '          b)如果不存在按所有类别轧帐的有效记录时,按下面的3以下规则处理
    '    3.如果存在轧帐规零记录,则以轧帐规零记录的登记时间为准
    '    4.未轧过的,缺省为领取备用金时间
    '    5.如果未领用备用金的，缺省时间为当前时间-1个月的且允许更改上次转帐时间
    
    '记录性质:1-收费员轧账记录(缴款书)；2-财务组收款记录;3-财务组轧账记录(组缴款书); _
    '4-财务收款记录;5-手工缴款(与原来功能保持不变);6-轧帐归零(切换成新模式后，必须先清零结算，这为清零结算记录)
    If strTYPE = "" Then Exit Function
    
    strSQL = "Select  Zl_Rollingcurtain_Lastdate([1],[2]) as 轧帐时间 From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, strTYPE)
    If Not rsTemp.EOF Then
        GetPreRollingCurtainTime = Format(rsTemp!轧帐时间, "yyyy-mm-dd HH:MM:SS")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDefaultTime() As String
    '获取缺省轧账结束时间
    On Error GoTo errH
    Dim strSQL As String, rsTemp As ADODB.Recordset, strValue As String
    Dim datValue As Date, datNow As Date
    
    strValue = zlDatabase.GetPara("缺省轧帐时间", glngSys, mlngModule, "", dtpTime, InStr(1, mstrPrivs, ";参数设置;") > 0)
    
    If strValue = "" Then GetDefaultTime = "": Exit Function
    
    datNow = zlDatabase.Currentdate
    strSQL = "Select 1 From 人员收缴记录 Where 终止时间 >= [1] And 作废时间 Is Null And 收款员 = [2]"
    datValue = CDate(Format(datNow - IIf(Format(datNow, "hh:mm:ss") >= Format(strValue, "hh:mm:ss"), 0, 1), "yyyy-MM-dd") & " " & Format(strValue, "hh:mm:ss"))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datValue, UserInfo.姓名)
    
    If Not rsTemp.EOF Then Exit Function
    
    GetDefaultTime = Format(datNow - IIf(Format(datNow, "hh:mm:ss") >= Format(strValue, "hh:mm:ss"), 0, 1), "yyyy-MM-dd") & " " & Format(strValue, "hh:mm:ss")

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function InitData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 14:20:10
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '获取上次轧帐时间
    mstrPreDate = GetPreRollingCurtainTime(GetRollingType)
    Call LoadCurBalanceData
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadCurBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载当前人员缴款余额数据
    '编制:刘兴洪
    '日期:2015-03-03 12:06:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  decode(nvl(b.性质,0),1,1,2,2,3,10,4,11,4) as 序号, " & _
    "               A.结算方式, A.余额, A.上次轧帐时间 " & _
    "   From 人员缴款余额 A,结算方式  B" & _
    "   Where a.结算方式=b.名称(+)  And A.收款员 = [1] And A.性质 = 1" & _
    "   Order by 上次轧帐时间 Desc,序号,结算方式"
    '--1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,
    '   5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
    Call InitGrid
    With vsBalance
        .Clear 1
        .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .TextMatrix(lngRow, .ColIndex("暂存金额")) = Format(Val(Nvl(rsTemp!余额)), "#,###0.00;-#,###0.00;0.00;-0.00")
            lngRow = lngRow + 1: .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "暂存金列表", False
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlShowChargeRollingCourtain(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
        '-------------------------------------------------------------------------------------------------
        '功能:收费员轧账接口
        '入参:frmMain-调用的主窗体
        '       strOperatorName-收费员,为空时,缺省为当前操作员
        '返回:收费轧帐成功一次以上,返回true,否则返回False
        '编制:刘兴洪
        '日期:2013-08-13 10:31:00
        '说明:
        '-------------------------------------------------------------------------------------------------
        mlngModule = lngModule: mstrPrivs = strPrivs
        mstrPreDate = ""
        If CheckDepend = False Then Exit Function
        If InitData = False Then
            Err = 0: On Error Resume Next
            Unload Me
            Err.Clear: Err = 0
            Exit Function
        End If
        '初始化数据
        Call InitFace
        mblnFirst = True
        If frmMain Is Nothing Then
            Me.Show
        Else
            Me.Show , frmMain
        End If
End Function

Public Sub BHShowList(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngMain As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '编制:刘兴洪
    '日期:2013-10-17 18:17:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mstrPreDate = ""
    If CheckDepend = False Then Exit Sub
    If InitData = False Then
        Err = 0: On Error Resume Next
        Unload Me
        Err.Clear: Err = 0
        Exit Sub
    End If
    '初始化数据
    Call InitFace
    mblnFirst = True
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
    Me.ZOrder 0
End Sub

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '编制:刘兴洪
    '日期:2013-09-03 10:29:48
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    '当前暂存金
    With vsBalance
        .Rows = 3: .Cols = 2
       .TextMatrix(0, 0) = "结算方式"
       .TextMatrix(0, 1) = "暂存金额"
       For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
       Next
       .TextMatrix(1, .ColIndex("结算方式")) = "现金"
       .TextMatrix(1, .ColIndex("暂存金额")) = "100"
       .TextMatrix(2, .ColIndex("结算方式")) = "支票"
       .TextMatrix(2, .ColIndex("暂存金额")) = "100"
       .AutoSizeMode = flexAutoSizeColWidth
       Call .AutoSize(0, .Cols - 1)
    End With
End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2013-09-03 14:43:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitPanel
    Call InitPage
    Set dkpMan.TabPaintManager.Font = vsBalance.Font
    Set dkpMan.PaintManager.CaptionFont = vsBalance.Font
    dkpMan.PanelPaintManager.StaticFrame = True
    stbThis.Panels(3).Text = UserInfo.姓名
    stbThis.Panels(3).ToolTipText = "当前操作员:" & UserInfo.姓名
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 18:21:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
    Dim objCustom As CommandBarControlCustom
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
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
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "轧帐(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "轧账(&Z)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "作废轧账(&D)")
        mcbrControl.IconId = 229
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "现金点钞(&E)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "重打缴款书(&R)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "查看明细数据(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("R"), conMenu_Edit_ChargeBook_Reprint
        .Add FCONTROL, Asc("S"), conMenu_Edit_RollingCurtain
        .Add FCONTROL, Asc("D"), conMenu_Edit_RollingCurtain_Cancel
        .Add FCONTROL, Asc("T"), conMenu_View_Detail
        .Add 0, VK_F6, conMenu_Edit_CheckCash
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        '.AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "轧帐"): mcbrControl.BeginGroup = True
         mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "轧帐作废"): mcbrControl.BeginGroup = True
         mcbrControl.IconId = 229
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "现金点钞")
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "查询明细")
        mcbrControl.IconId = 2322
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        Set objCustom = .Add(xtpControlCustom, conMenu_COMBOX_INTERFACE, "轧帐类别")
        objCustom.Flags = xtpFlagRightAlign
        objCustom.HideFlags = xtpNoHide
        objCustom.Handle = picType.hWnd
        objCustom.BeginGroup = True
        picType.BackColor = CommandBarsGlobalSettings.ColorManager.Color(XPCOLOR_TOOLBAR_FACE)
        
        With cboType
            .Clear
            .AddItem "0", "所有类别", True, True, True
            
            If InStr(1, mstr轧帐人员性质, ",门诊收费员,") > 0 Then
                .AddItem 1, "收费", False, True, True
            End If
            If InStr(1, mstr轧帐人员性质, ",预交收款员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",入院登记员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",发卡登记人,") > 0 Then
                If mbln预交分别轧帐 Then
                    .AddItem 21, "门诊预交", False, True, True
                    .AddItem 22, "住院预交", False, True, True
                Else
                    .AddItem 2, "预交", False, True, True
                End If
            End If
            If InStr(1, mstr轧帐人员性质, ",住院结帐员,") > 0 Then
                .AddItem 3, "结帐", False, True, True
            End If
            If InStr(1, mstr轧帐人员性质, ",门诊挂号员,") > 0 Then
                .AddItem 4, "挂号", False, True, True
            End If
            If InStr(1, mstr轧帐人员性质, ",门诊挂号员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",入院登记员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",发卡登记人,") > 0 Then
                .AddItem 5, "就诊卡", False, True, True
            End If
            If InStr(1, mstr轧帐人员性质, ",门诊挂号员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",入院登记员,") > 0 _
               Or InStr(1, mstr轧帐人员性质, ",发卡登记人,") > 0 Then
                .AddItem 6, "消费卡", False, True, True
            End If
        End With
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        Set objPane = .CreatePane(mPaneIndex.EM_PN_明细列表, 400, 400, DockLeftOf, Nothing)
        objPane.Title = "轧帐信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
      
        Set objPane = .CreatePane(mPaneIndex.EM_PN_暂存金, 100, 100, DockRightOf, objPane)
        objPane.Title = "当前暂存金": objPane.Options = PaneNoCloseable
        objPane.Handle = picBalanceList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Dim strRollingType As String
    Err = 0: On Error GoTo Errhand:
    If mfrmChargeRollingCurtain Is Nothing Then
        Set mfrmChargeRollingCurtain = New frmChargeRollingCurtain
        Load mfrmChargeRollingCurtain
    End If
    strRollingType = GetRollingType
    '初始化变量
    Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.姓名, strRollingType, GetDefaultTime)
    If mfrmHistory Is Nothing Then
        Set mfrmHistory = New frmChargeRollingCurtainHistory
        Load mfrmHistory
    End If
    Call mfrmHistory.zlInitVar(Me, mlngModule, mstrPrivs)
    Set objItem = tbPage.InsertItem(EM_PG_轧帐列表, "轧帐", mfrmChargeRollingCurtain.hWnd, 0)
    objItem.Tag = EM_PG_轧帐列表
    Set objItem = tbPage.InsertItem(EM_PG_历史列表, "历史轧帐信息", mfrmHistory.hWnd, 0)
    objItem.Tag = EM_PG_历史列表
     With tbPage
        Set tbPage.PaintManager.Font = vsBalance.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboType_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    If GetRollingType = "" Then Node.Checked = True: Node.Selected = True
    Call ReloadData(Format(mfrmChargeRollingCurtain.dtpEndDate.Value, "yyyy-mm-dd hh:mm:ss"))
    mstrPrevRollingType = GetRollingType
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
     If Action = PaneActionAttached Then Cancel = True: Exit Sub
     If Action = PaneActionAttaching Then Cancel = True: Exit Sub
     If Action = PaneActionFloated Then Cancel = True: Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    DoEvents
    Call DefaultSetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_轧帐列表
        Call mfrmChargeRollingCurtain.MainKeyDown(KeyCode, Shift)
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mbln预交分别轧帐 = Val(zlDatabase.GetPara("预交轧帐按门诊或住院分别轧帐", glngSys, glngModul, "0")) = 1
    RestoreWinState Me, App.ProductName
    Call zlDefCommandBars
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmChargeRollingCurtain Is Nothing Then Unload mfrmChargeRollingCurtain
    If Not mfrmHistory Is Nothing Then Unload mfrmHistory
    Set mfrmChargeRollingCurtain = Nothing
    Set mfrmHistory = Nothing
End Sub
Private Sub picBalanceList_Resize()
    Err = 0: On Error Resume Next
    With vsBalance
        .Top = picBalanceList.ScaleTop + 20
        .Left = picBalanceList.ScaleLeft + 20
        .Width = picBalanceList.ScaleWidth - .Left * 2
        .Height = picBalanceList.ScaleHeight - .Top * 2 '.RowHeight(0) + .RowHeight(1) + 3 * 50
    End With
End Sub
 Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub ParameterSet()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置
    '编制:刘兴洪
    '日期:2013-09-12 15:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strDefaultTime As String
    If frmChargeRollingCurtainSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    strDefaultTime = GetDefaultTime
    If strDefaultTime <> "" Then
        mfrmChargeRollingCurtain.dtpEndDate.Value = Format(strDefaultTime, "yyyy-MM-dd hh:mm:ss")
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub RollingCurtain()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:轧帐处理
    '编制:刘兴洪
    '日期:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String
    If Val(tbPage.Selected.Tag) = EM_PG_历史列表 Then Exit Sub
    Call mfrmChargeRollingCurtain.SaveDataWithCheck
End Sub
Private Sub RollingCurtainCancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:轧帐作废处理
    '编制:刘兴洪
    '日期:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, blnDel As Boolean
    Dim strRollingType As String
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then Exit Sub
    strRollingType = GetRollingType
    If mfrmHistory.CancelData() Then
        '作废轧帐后,重新读取上次轧帐时间
        Call InitData
        Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.姓名, strRollingType)
        Call mfrmChargeRollingCurtain.RefreshPage
        Exit Sub
    End If
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '退出(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '打印设置
        Case conMenu_File_Preview: Call zlPrintRpt(2)  '预览(&V)
        Case conMenu_File_Print: Call zlPrintRpt(1) '打印(&P)
        Case conMenu_File_Excel: Call zlPrintRpt(3)  '输出到&Excel…
        Case conMenu_File_Parameter: Call ParameterSet '参数设置
        Case conMenu_Edit_RollingCurtain: Call RollingCurtain  '轧账(&Z)
        Case conMenu_Edit_RollingCurtain_Cancel: Call RollingCurtainCancel '作废轧账(&D)
        Case conMenu_Edit_CheckCash: Call CheckCash '现金点钞(&E)
        Case conMenu_Edit_ChargeBook_Reprint:  Call RePrintBill '重打缴款书(&R)
        Case conMenu_View_Detail: Call ShowChargeList '查看明细数据(&V)
        Case conMenu_View_Refresh: zlRefresh '刷新(&R)
        Case conMenu_View_StatusBar '状态栏(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_COMBOX_INTERFACE   '点击选择
            Call ReloadData
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                '执行发布到当前模块的报表
                Call CallCustomRpt(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub ReloadData(Optional ByVal strTime As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '编制:刘兴洪
    '日期:2015-03-03 12:15:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRollingType As String
    
    On Error GoTo errHandle
    strRollingType = GetRollingType
    If strRollingType = "" Then
        MsgBox "请至少选择一项类别进行轧账！", vbInformation, gstrSysName
        Exit Sub
    End If
    mstrPreDate = GetPreRollingCurtainTime(strRollingType)
    Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.姓名, strRollingType, strTime)
    Call mfrmChargeRollingCurtain.RefreshPage
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHavePrivs As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_RollingCurtain ' "轧账(&Z)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "轧帐")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_轧帐列表
    Case conMenu_Edit_RollingCurtain_Cancel ' "作废轧账(&D)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "轧帐作废")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_历史列表
        If Control.Enabled Then
            Control.Enabled = mfrmHistory.GetChargeRollingCurtainID <> 0 _
                And Not mfrmHistory.GetChargeRollingCurtainDel
        End If
    Case conMenu_View_Detail    '查看明细数据(&V)
        If Val(tbPage.Selected.Tag) = EM_PG_历史列表 Then
            With mfrmHistory.vsRollingCurtain
                Control.Enabled = .RowSel >= 1 And .TextMatrix(.RowSel, .ColIndex("轧帐单号")) <> ""
            End With
        Else
            Control.Enabled = True
        End If
    Case conMenu_Edit_CheckCash ' "现金点钞(&E)")
        
    Case conMenu_Edit_ChargeBook_Reprint ' "重打缴款书(&R)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "重打缴款书") And zlStr.IsHavePrivs(mstrPrivs, "缴款书打印")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_历史列表
        If Control.Enabled Then
            Control.Enabled = mfrmHistory.GetChargeRollingCurtainID <> 0 _
                And Not mfrmHistory.GetChargeRollingCurtainDel
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_COMBOX_INTERFACE
        If Not tbPage.Selected Is Nothing Then
            Control.Visible = Val(tbPage.Selected.Tag) <> EM_PG_历史列表
        Else
            Control.Visible = False
        End If

    End Select
End Sub
Private Function CheckDepend() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据依赖性
    '返回:数据合法,返回true，否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 17:10:03
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    CheckDepend = False
    
    On Error GoTo errHandle
     mstr轧帐人员性质 = ""
    gstrSQL = "" & _
    "   Select  B.ID,A.人员性质  " & _
    "   From 人员性质说明 A, 人员表 B " & _
    "   Where A.人员id = B.ID And A.人员性质 In ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人') And B.ID=[1] " & _
    "   Order By ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前操作员是否为相应性质人员", UserInfo.ID)
    If rsTemp.EOF Then
        ShowMsgbox "你不具备“门诊挂号员,门诊收费员,预交收款员,住院结帐员,入院登记员,发卡登记人”的性质，不能使用该模块！"
        rsTemp.Close
        Exit Function
    End If
    Do While Not rsTemp.EOF
        If InStr(mstr轧帐人员性质 & ",", "," & rsTemp!人员性质 & ",") = 0 Then
            mstr轧帐人员性质 = mstr轧帐人员性质 & "," & rsTemp!人员性质
        End If
        rsTemp.MoveNext
    Loop
    If mstr轧帐人员性质 <> "" Then mstr轧帐人员性质 = mstr轧帐人员性质 & ","
    
    Set rsTemp = Get结算方式
    rsTemp.Filter = "性质=1"
    If rsTemp.EOF Then
        rsTemp.Filter = 0
        ShowMsgbox "结算方式中不存在一条件有现金性质的结算方式,请在结算方式管理中设置!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Filter = 0
    rsTemp.Close
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub picType_Resize()
    On Error Resume Next
    With cboType
        .Left = 15
        .Top = 15
        .Width = picType.ScaleWidth - 30
        .Height = picType.ScaleHeight - 30
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call DefaultSetFocus
End Sub

Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "暂存金列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "暂存金列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub zlPrintRpt(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objRow As New zlTabAppRow, bytPrn As Byte
    Dim i As Long, lngRow As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then
        '打印轧帐信息
        Call mfrmChargeRollingCurtain.zlPrint(bytMode)
        Exit Sub
    End If
    Call mfrmHistory.zlPrint(bytMode)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打缴款书
    '编制:刘兴洪
    '日期:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmHistory.RePrintBill
End Sub
Private Sub CheckCash()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:现金点钞
    '编制:刘兴洪
    '日期:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then
        dblMoney = mfrmChargeRollingCurtain.GetCashMoney
    End If
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub
Private Sub zlRefresh()
    '重新进行数据刷新
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then
        Call mfrmChargeRollingCurtain.zlRefresh
    Else
        Call mfrmHistory.zlRefresh
    End If
    Call DefaultSetFocus
End Sub

Public Sub RefreshBasic()
    Call InitData
End Sub
Private Function GetRollingType() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取轧帐类别
    '返回:返回轧帐类别
    '编制:刘兴洪
    '日期:2015-03-06 10:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTYPE As String
    On Error GoTo errHandle
    strTYPE = cboType.GetNodesCheckedDatas(False)
    GetRollingType = strTYPE
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowChargeList()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRollingType As String
    
    On Error GoTo errHandle
    
    strRollingType = GetRollingType
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then
         Call mfrmChargeRollingCurtain.ShowChargeList(Me, strRollingType)
         Exit Sub
    End If
    '历史数据显示
    Call mfrmHistory.ShowChargeList(Me)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CallCustomRpt(ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo errHandle
    If Val(tbPage.Selected.Tag) = EM_PG_轧帐列表 Then
         Call mfrmChargeRollingCurtain.CallCustomRpt(Me, lngSys, strRptCode)
         Exit Sub
    End If
    '历史数据显示
    Call mfrmHistory.CallCustomRpt(Me, lngSys, strRptCode)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub DefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置光标的缺省定位
    '编制:刘兴洪
    '日期:2013-10-16 14:25:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_轧帐列表
        picType.Visible = True
        mfrmChargeRollingCurtain.zlDefaultSetFocus
    Case Else
        picType.Visible = False
        mfrmHistory.zlDefaultSetFocus
    End Select
End Sub
