VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{257A5750-6F4D-4A7A-A149-21D28B3E6EAA}#6.1#0"; "ZLPacsRichPages.ocx"
Begin VB.Form frmReport 
   Caption         =   "报告查阅"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12960
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBarControl 
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12735
      TabIndex        =   2
      Top             =   0
      Width           =   12735
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   2520
         Top             =   240
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   13815
      TabIndex        =   0
      Top             =   960
      Width           =   13815
      Begin VB.PictureBox picVLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5940
         Left            =   5160
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5940
         ScaleWidth      =   30
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   30
      End
      Begin VB.PictureBox picReportContent 
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   8280
         ScaleHeight     =   7095
         ScaleWidth      =   5295
         TabIndex        =   3
         Top             =   120
         Width           =   5295
         Begin ZLPacsRichPageScale.ZLRichPageScaleAct zlDocEditor 
            Bindings        =   "frmReport.frx":6852
            Height          =   1575
            Left            =   1800
            TabIndex        =   4
            Top             =   960
            Width           =   2415
            Object.Visible         =   -1  'True
            AutoScroll      =   0   'False
            AutoSize        =   0   'False
            AxBorderStyle   =   1
            BorderWidth     =   0
            Caption         =   "ZLRichPages"
            Color           =   -16777201
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            KeyPreview      =   0   'False
            PixelsPerInch   =   96
            PrintScale      =   1
            Scaled          =   -1  'True
            DropTarget      =   0   'False
            HelpFile        =   ""
            PopupMode       =   0
            ScreenSnap      =   0   'False
            SnapBuffer      =   10
            DockSite        =   0   'False
            DoubleBuffered  =   0   'False
            ParentDoubleBuffered=   0   'False
            UseDockManager  =   0   'False
            Enabled         =   -1  'True
            AlignWithMargins=   0   'False
            HMenuVisible    =   -1  'True
            VMenuVisible    =   -1  'True
            ReadOnly        =   0   'False
            Orientation     =   0
            BottomMagin     =   2.54
            BoundLeftRight  =   20
            FooterVisible   =   -1  'True
            FooterY         =   10
            MaxPageBreakHeight=   25
            MinPageBreakHeight=   5
            PageBreakHeight =   20
            PageNoFirst     =   1
            PageNoFromNumber=   1
            PageNoHAlign    =   0
            PageNoVAlign    =   0
            PageNoVisible   =   -1  'True
            PageViewMode    =   -1  'True
            RightMargin     =   3.17
            TopMargin       =   2.54
            BackgroundStyle =   3
            CtlColor        =   10070188
            IsShowHint      =   -1  'True
            TabNavigation   =   1
            NoReadOnlyJumps =   0   'False
            NoCaretHighLightJumps=   0   'False
            NoImageResize   =   0   'False
            HideReadOnlyCaret=   -1  'True
            AutoSwitchLang  =   0   'False
            WantTabs        =   -1  'True
            DoNotWantShiftReturns=   0   'False
            DoNotWantReturns=   0   'False
            CtrlJumps       =   -1  'True
            ClearTagOnStyleApp=   0   'False
            IsShowCheckPoints=   0   'False
            IsShowPageBreaks=   0   'False
            IsShowSpecialCharacters=   0   'False
            IsShowHiddenText=   0   'False
            IsShowItemHints =   0   'False
            IsDblClickSelectsWord=   -1  'True
            IsRClickDeselects=   -1  'True
            AlignPageH      =   0
            AlignPageV      =   0
            ViewMode        =   0
            ZoomMode        =   0
            EditZoomMode    =   2
            ZoomPercent     =   68
            ZoomPercentEdit =   100
            ParentCustomHint=   0   'False
            Modified        =   -1  'True
            UndoLimit       =   -1
            IsMarginRectVisible=   -1  'True
            StateView       =   -1  'True
            BackGroundPicture=   "frmReport.frx":6866
            BeginProperty PageNoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderVisible   =   -1  'True
            HeaderY         =   10
            TableAutoAddRow =   -1  'True
            ThumbilsVisible =   -1  'True
            SimpleViewMode  =   0   'False
            SclRVRulerVVisible=   0   'False
            SclRVRulerHVisible=   0   'False
            ScrollBarVVisible=   -1  'True
            ScrollBarHVisible=   -1  'True
            BackGroudVisible=   -1  'True
            BorderPenStyle  =   0
            Ver             =   "2.1"
            StatusBarVisible=   -1  'True
            CanEdit         =   -1  'True
            DisableCopyElement=   0   'False
            PageWidth       =   21
            PageHeight      =   29.7
            CanPopMenu      =   0   'False
            LeftMargin      =   3.17
            CanInput        =   -1  'True
            TableGridVisible=   0   'False
            CanEditHeader   =   -1  'True
            CanEditFooter   =   -1  'True
            IsRevision      =   0   'False
            RevisionTag     =   ""
            RevisionAddColor=   0
            RevisionDelColor=   0
            MaskText        =   ""
            AllowSelection  =   -1  'True
            BeginProperty MaskTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FinalShowMode   =   0   'False
            DocMasterId     =   ""
            PageSetupInPre  =   0   'False
            ServerTime      =   "1899-12-30"
            XMLEncoding     =   ""
            HScrollPos      =   0
            VScrollPos      =   4
            IsShowMargin    =   0   'False
            IsAutoPageWidth =   0   'False
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Bindings        =   "frmReport.frx":6B2E
            Height          =   1095
            Left            =   2640
            TabIndex        =   5
            Top             =   3480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1931
            _Version        =   393217
            ScrollBars      =   3
            TextRTF         =   $"frmReport.frx":6B81
         End
      End
      Begin VB.PictureBox picReportList 
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   120
         ScaleHeight     =   7095
         ScaleWidth      =   5700
         TabIndex        =   1
         Top             =   120
         Width           =   5700
         Begin VB.PictureBox picReportControl 
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   360
            ScaleHeight     =   2055
            ScaleWidth      =   6615
            TabIndex        =   8
            Top             =   600
            Width           =   6615
            Begin VB.OptionButton optTime 
               Caption         =   "3天内"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   510
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.ComboBox cboDoctor 
               Height          =   300
               Left            =   3600
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   840
               Width           =   1500
            End
            Begin VB.OptionButton optTime 
               Caption         =   "7天内"
               Height          =   180
               Index           =   1
               Left            =   1005
               TabIndex        =   19
               Top             =   510
               Width           =   855
            End
            Begin VB.OptionButton optTime 
               Caption         =   "半个月内"
               Height          =   180
               Index           =   2
               Left            =   1875
               TabIndex        =   18
               Top             =   510
               Width           =   1095
            End
            Begin VB.OptionButton optTime 
               Caption         =   "一个月内"
               Height          =   180
               Index           =   3
               Left            =   3000
               TabIndex        =   17
               Top             =   510
               Width           =   1095
            End
            Begin VB.ComboBox cboKind 
               Enabled         =   0   'False
               Height          =   300
               ItemData        =   "frmReport.frx":6C10
               Left            =   2520
               List            =   "frmReport.frx":6C23
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1230
               Width           =   975
            End
            Begin VB.TextBox txtKind 
               Enabled         =   0   'False
               Height          =   270
               Left            =   3600
               TabIndex        =   15
               Top             =   1245
               Width           =   1500
            End
            Begin VB.ComboBox cboDept 
               Height          =   300
               IMEMode         =   1  'ON
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1230
               Width           =   1500
            End
            Begin VB.ComboBox cbo开单科室 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   840
               Width           =   1500
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "门 诊"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Tag             =   "门"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "住 院"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   11
               Tag             =   "住"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "外 来"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   10
               Tag             =   "外"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "体 检"
               Height          =   255
               Index           =   3
               Left            =   3000
               TabIndex        =   9
               Tag             =   "体"
               Top             =   1680
               Width           =   855
            End
            Begin MSComCtl2.DTPicker dtpStart 
               Height          =   255
               Left            =   960
               TabIndex        =   22
               Top             =   83
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   113180675
               CurrentDate     =   42180
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   255
               Left            =   2760
               TabIndex        =   23
               Top             =   83
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   113180675
               CurrentDate     =   42180
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "~"
               Height          =   180
               Left            =   2535
               TabIndex        =   28
               Top             =   120
               Width           =   90
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "检查日期"
               Height          =   180
               Left            =   120
               TabIndex        =   27
               Top             =   120
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "申请医生"
               Height          =   180
               Left            =   2520
               TabIndex        =   26
               Top             =   900
               Width           =   720
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "检查科室"
               Height          =   180
               Left            =   120
               TabIndex        =   25
               Top             =   1290
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "申请科室"
               Height          =   180
               Left            =   120
               TabIndex        =   24
               Top             =   900
               Width           =   720
            End
         End
         Begin TabDlg.SSTab tabReport 
            Height          =   360
            Left            =   480
            TabIndex        =   7
            Top             =   120
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "当前病人报告"
            TabPicture(0)   =   "frmReport.frx":6C4D
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).ControlCount=   0
            TabCaption(1)   =   "全院病人报告"
            TabPicture(1)   =   "frmReport.frx":6C69
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfReportList 
            Height          =   3615
            Left            =   240
            TabIndex        =   6
            Top             =   2640
            Width           =   6735
            _cx             =   11880
            _cy             =   6376
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
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   400
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
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mMenuID
    conMenu_File_PrintSet = 101
    conMenu_File_Preview = 102
    conMenu_File_Print = 103
    conMenu_File_ViewState
    conMenu_File_ViewImage
    conMenu_File_Refresh
    conMenu_File_Exit
End Enum

Private mobjReport As Object
'Private mcnOracle As New ADODB.Connection
Private mrsRecord As ADODB.Recordset

Private mlngAdviceId As Long
Private mstrReportId As String
Private mobjFtp As New clsFtp
Private mobjFtpInfo As tFtpInfo

Public mblnAutoView As Boolean  '是否自动查阅，True--自动查阅，鼠标点击列表后，自动标记“已查阅”
Public mblnRIS As Boolean       '是否RIS中查看，如果是则不显示退出，查阅，打印等按钮，设置查询开始时间为一年
Public mblnShow As Boolean      '是否显示左边的过滤条件

Private mblnIsNewReport As Boolean
Private mblnIsNoAskPrint As Boolean
Private mlngPatFrom As Long
Private mblnIsConfiging As Boolean
Private mblnFirst As Boolean
Private mlngViewReport As Long      '0--审核签名后即可查看报告，1--终审签名后即可查看报告


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = gobjComLib.zlCommFun.GetPubIcons
    
    With cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbrMain.ActiveMenuBar.Visible = False
    
'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    
    If mblnRIS = False Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_PrintSet, "打印设置", "打印设置", 181, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "打印预览", 102, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "打印", 103, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_ViewState, "查阅", "查阅", 2322, True)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_ViewImage, "观片", "观片", 8111, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Refresh, "刷新", "查询刷新数据", 791, True)
    
    If mblnRIS = False Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Exit, "退出", "退出", 2613, True)
    End If
End Sub

''为了查询Clob类型内容，需换成oledb连接方式
'Public Function ConnectOracle(ByVal strUser As String, ByVal strPassW As String, ByVal strServer As String) As Boolean
'On Error GoTo ErrH
'    ConnectOracle = False
'
'    '判断连接状态
'    If mcnOracle.State = adStateOpen Then mcnOracle.Close
'
'    mcnOracle.ConnectionString = "Provider=OraOLEDB.Oracle.1;User ID=" & strUser & ";password=" & strPassW & ";Data Source=" & strServer & ";Persist Security Info=False"
'    mcnOracle.CursorLocation = adUseClient
'
'    mcnOracle.Open
'
'    If mcnOracle.State = adStateOpen Then
'        ConnectOracle = True
'    Else
'        ConnectOracle = False
'    End If
'
'    Exit Function
'ErrH:
'    MsgBox err.Description, vbCritical, "系统信息"
'    err.Clear
'End Function

Public Sub InitEdit()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    '开单科室
    strSql = " Select Distinct A.ID,A.编码,A.名称,b.工作性质" & _
                " From 部门表 A,部门性质说明 B " & _
                " Where B.部门ID = A.ID " & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                " And (B.工作性质 IN('临床','体检','检查'))" & _
                " Order by A.编码"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, Caption)
    
    cbo开单科室.Clear
    cbo开单科室.AddItem ""
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cbo开单科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        If cbo开单科室.ListCount > 0 And cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    End If
    
    cboDept.Clear
    cboDept.AddItem ""
    rsTmp.Filter = "工作性质='检查'"
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    End If
    
    If cboKind.ListCount > 0 And cboKind.ListIndex = -1 Then cboKind.ListIndex = 0
    
    dtpEnd.Value = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    dtpStart = Format(Now - 3, "yyyy-mm-dd HH:MM")
End Sub

Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b, 人员性质说明 c" & vbNewLine & _
                " Where a.人员id = b.Id And b.Id = c.人员id And c.人员性质 = '医生' And" & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, Caption, lng科室ID)
    
    cboDoctor.Clear
    cboDoctor.AddItem ""
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cboDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            If rsTmp!ID = UserInfo.ID Then cboDoctor.ListIndex = cboDoctor.NewIndex
            rsTmp.MoveNext
        Loop
        
        If cboDoctor.ListCount > 0 And cboDoctor.ListIndex = -1 Then cboDoctor.ListIndex = 0
    End If
End Sub

'Private Function InitOracleConn() As Boolean
'    Dim strUser As String, strPwd As String, strServer As String
'
'On Error GoTo Errorhand
'
'    strUser = UCase(gcnOracle.Properties(23))
'    strPwd = gcnOracle.Properties(24)
'    strServer = UCase(Split(gcnOracle.Properties(8), "=")(2))
'    gstrDBUser = strUser
'
'    '为了查询Clob类型内容，需换成oledb连接方式
'    If Not ConnectOracle(strUser, strPwd, strServer) Then Exit Function
'
'    InitOracleConn = True
'
'    Exit Function
'Errorhand:
'    MsgBox err.Description, vbExclamation, gstrSysName
'End Function

Public Sub ShowMe(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional ByVal blnAutoView As Boolean = True, Optional objParent As Object, Optional blnShowModal As Boolean = False)
    
On Error GoTo Errorhand
    
    '直接打开PACS报告窗体时，默认显示左边的列表和报告打印，查阅，退出等按钮
    mblnShow = True
    mblnRIS = False
    
    mblnAutoView = blnAutoView
    
    Call RefreshForm(lngAdviceID, strReportId, objParent)
    
    Call Show(IIf(blnShowModal, 1, 0), objParent)
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function LoadImageFromXml(ByVal strXml As String, ByVal strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'功能：解析xml文档，将图像下载到指定位置，并返回所有图象文件名
'返回：
    Dim objXml As New DOMDocument
    Dim objXmlNodes As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeAttribute As IXMLDOMNode
    Dim strImgSVG As String
    Dim objImgFileName As New Collection
    Dim objSvg As New ZLSvgProcess.zlSvgConvert
    Dim objPic As StdPicture
    
On Error GoTo Errorhand
    
    If objXml.loadXML(strXml) = False Then
        strMsg = "报告内容加载失败！"
        Exit Function
    End If
    
    Set objXmlNodes = objXml.selectNodes("*//image")
    
    If objXmlNodes.length <= 0 Then
        Set LoadImageFromXml = objImgFileName
        strMsg = "此报告没有图像。"
        Exit Function
    End If
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '从FTP上获取图像文件后返回图像串
            strImgSVG = objSvg.DecodeBase64(GetFtpImgSVG(objXmlNodeAttribute.Text, strMsg))
            
            If objSvg.IsSvgContext(strImgSVG) Then
                '将图像放到指定目录中
                Set objPic = objSvg.ContextToPic(strImgSVG)
                Call SavePicture(objPic, strTmpImgDir & objXmlNodeAttribute.Text & ".jpg")
                
                '将图像名称放入图像集合中
                Call objImgFileName.Add(objXmlNodeAttribute.Text)
            End If
        End If
    Next
    
    Set LoadImageFromXml = objImgFileName
    
    Exit Function
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function AddImageInfoToXml(ByVal strXml As String, ByVal strDocId As String) As String
'功能：解析xml文档，将图像信息加入文档中
'返回：
    Dim objXml As New DOMDocument
       
    Dim objXmlNodes As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeAttribute As IXMLDOMNode
    Dim strImgSVG As String
    
On Error GoTo Errorhand
    
    If objXml.loadXML(strXml) = False Then
        MsgBox "报告内容加载失败！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Set objXmlNodes = objXml.selectNodes("*//image")
    
    If objXmlNodes.length <= 0 Then
        AddImageInfoToXml = strXml
        Exit Function
    End If
    
    '初始化FTP相关信息
    Call InitFtpInfo(strDocId)
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '从FTP上获取图像文件后返回图像串
            strImgSVG = GetFtpImgSVG(objXmlNodeAttribute.Text)
            
            Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("img")
            '将图像信息写入xml
            objXmlNodeAttribute.Text = strImgSVG
        End If
    Next
    
    AddImageInfoToXml = objXml.xml
    
    Exit Function
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function InitFtpInfo(ByVal strDocId As String) As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 'ReportImages/' || to_Char(b.创建时间,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
            "a.设备号, a.FTP用户名, a.FTP密码, a.IP地址,'/'||a.Ftp目录||'/' As Root " & _
            "From 影像设备目录 a, 影像报告记录 b where a.设备号 = b.设备号 And b.id = [1]"
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", strDocId)
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInfo.FtpDir = Nvl(rsTmp("Root"))
    mobjFtpInfo.FtpIP = Nvl(rsTmp("IP地址"))
    mobjFtpInfo.FtpPswd = Nvl(rsTmp("FTP密码"))
    mobjFtpInfo.FTPUser = Nvl(rsTmp("FTP用户名"))
    mobjFtpInfo.DiviceId = Trim(Nvl(rsTmp("设备号")))
    
    mobjFtpInfo.SubDir = Nvl(rsTmp("URL"))
    mobjFtpInfo.DestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInfo = True
End Function

Private Function ConnFtp() As Boolean
    If mobjFtp.hConnection = 0 Then
        '连接FTP存储设备
        If mobjFtp.FuncFtpConnect(mobjFtpInfo.FtpIP, mobjFtpInfo.FTPUser, mobjFtpInfo.FtpPswd) = 0 Then
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

'从FTP上获取SVG格式图像
Private Function GetFtpImgSVG(ByVal strKey As String, Optional ByRef strMsg As String = "") As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    
    If strKey = "" Then Exit Function
    
    strLocalFileName = Replace(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir & strKey, "/", "\")
    strVirtualPath = Replace(mobjFtpInfo.FtpDir & mobjFtpInfo.SubDir, "\", "/")
    
    '创建本地路径
    If Not objFSO.FolderExists(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir) Then
        Call MkLocalDir(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir)
    End If
    
    '如果本地存在则删除
    If objFSO.FileExists(strLocalFileName) Then Call objFSO.DeleteFile(strLocalFileName, True)
    
    '连接FTP
    If ConnFtp() = False Then
        strMsg = "FTP不能正常连接，请检查网络设置。"
        Exit Function
    End If
    
    If mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, objFSO.GetFileName(strLocalFileName)) <> 0 Then
        strMsg = "图像内容从FTP服务器上获取失败！"
        Exit Function
    End If
    
    '下载后读取
    GetFtpImgSVG = GetFileContent(strLocalFileName)
End Function

Private Sub InitReportEditor()
    zlDocEditor.FooterVisible = False
    zlDocEditor.HeaderVisible = False
    zlDocEditor.HMenuVisible = False
    zlDocEditor.PageNoVisible = False
    zlDocEditor.ThumbilsVisible = False
    zlDocEditor.VMenuVisible = False
    zlDocEditor.ZoomPercent = 100
    zlDocEditor.CanEdit = False
    zlDocEditor.CanInput = False
    zlDocEditor.TableGridVisible = False
    zlDocEditor.InitOCX hWnd
    rtxtReport.Locked = True
End Sub

Private Sub InitReportList()
'初始化样式配置列表
    With vsfReportList
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(2) = True
        .AutoSize 0, .Cols - 1
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
    End With
End Sub

Private Sub LoadReportContent(ByVal strContent As String, ByVal strDocId As String)
    Dim strXml As String
    
    If strContent = "" Then Exit Sub
    
    '解析xml文档，将图像信息加入文档中
    strXml = AddImageInfoToXml(strContent, strDocId)
    strXml = Replace(strXml, "吠", "名")
    strXml = Replace(strXml, "朠", "服")
    strXml = Replace(strXml, "丠", "不")
    
    zlDocEditor.OpenWithXML strXml
    zlDocEditor.FinalShowMode = True
End Sub

Public Sub RefreshReportContent(ByVal strReportId As String)
    Call ViewReportContent(strReportId)

'    If Trim(strReportId) <> "" Then
'        Call ViewReportContent(strReportId)
'    Else
'        Call LoadReport("")
'    End If
    
End Sub

Private Sub cboDept_Click()
    Call GetFilterData
End Sub

Private Sub cboDept_DropDown()
     On Error GoTo errHandle
    Call SendMessage(cboDept.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cboDoctor_Click()
    Call GetFilterData
End Sub

Private Sub cboDoctor_DropDown()
    On Error GoTo errHandle
    Call SendMessage(cboDoctor.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cboKind_Click()
    txtKind.Text = ""
End Sub

Private Sub cboKind_DropDown()
    On Error GoTo errHandle
    Call SendMessage(cboKind.hWnd, &H160, 100, 0)
errHandle:
End Sub

Private Sub cbo开单科室_Click()
    If cbo开单科室.ListIndex > -1 And cbo开单科室.Text <> "" Then
        InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        cboDoctor.Clear
    End If
    
    Call GetFilterData
End Sub

Private Sub Menu_File_Preview(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal lngAdviceID As Long)
On Error GoTo ErrHand
    If mobjReport Is Nothing Then Set mobjReport = CreateObject("zlRichEPR.cDockReport")        '电子病历报告
    
    If mblnIsNewReport Then
        Call zlDocEditor.PrintPreview(False, False, False, False, True)
    Else
        If Not mobjReport Is Nothing Then
            mobjReport.zlRefresh 0, 0
            mobjReport.zlRefresh lngAdviceID, UserInfo.部门ID
            mobjReport.zlExecuteCommandBars Control
        End If
    End If
    
    Exit Sub
    
ErrHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Menu_File_Print(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objReport As Object
    
    If mblnIsNewReport Then
        Call zlDocEditor.PrintPages
'    Else
'        Set objReport = CreateObject("zlRichEPR.cDockReport")       '电子病历报告
'
'        If Not objReport Is Nothing Then
'            objReport.zlRefresh 0, 0, , , , 1258
'            objReport.zlRefresh mlngAdviceId, UserInfo.部门ID, , , True, 1258
'            objReport.zlExecuteCommandBars Control
'        End If
    End If
End Sub


Private Sub cbo开单科室_DropDown()
     On Error GoTo errHandle
    Call SendMessage(cbo开单科室.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo Errorhand
    Dim lngAdviceID As Long
    
    If Control.ID = conMenu_File_ViewImage Or Control.ID = conMenu_File_Preview Then
        If vsfReportList.RowSel <= 0 Then
            MsgBox "请先选择需要操作的检查", vbExclamation, gstrSysName
            Exit Sub
        End If
        lngAdviceID = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "医嘱ID"))
    End If
    Select Case Control.ID
        Case conMenu_File_PrintSet  '打印设置
            Call zlPrintSet
            
        Case conMenu_File_Preview   '预览
            Call Menu_File_Preview(Control, lngAdviceID)
                
        Case conMenu_File_Print '打印
            Call Menu_File_Print(Control)
                
        Case conMenu_File_ViewState  '查阅
            If vsfReportList.RowSel <= 0 Then Exit Sub
            Call UpdateReportViewState(vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "报告ID")))
                
        Case conMenu_File_ViewImage '观片
            If vsfReportList.RowSel <= 0 Then Exit Sub
            Call mdlPublic.ViewImage(lngAdviceID, Me)
        
        Case conMenu_File_Refresh   '查询，刷新数据
            Call LoadReport(GetFilter)
        
        Case conMenu_File_Exit  '退出
            Unload Me
        
    End Select
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo Errorhand
    Select Case Control.ID
        Case mMenuID.conMenu_File_ViewState
            If vsfReportList.Rows <= 1 Then Exit Sub
            Control.Caption = IIf(vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "查阅")) = "√", "已查阅", "查阅")
            Control.Enabled = Not mblnAutoView And Not vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "查阅")) = "√"
            If Control.Enabled Then Control.Enabled = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "医嘱ID")) = mlngAdviceId
            
        Case mMenuID.conMenu_File_Preview, mMenuID.conMenu_File_Print
            Control.Enabled = vsfReportList.Rows > 1
                    
    End Select
    
    Control.Enabled = Not Control.Enabled
    Control.Enabled = Not Control.Enabled
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function GetFilter() As String
    Dim strPatFrom As String
    Dim strFilter As String
    Dim i As Integer
    Dim intPatFromCount As Integer
    
On Error GoTo Errorhand
    
    If cbo开单科室.Text <> "" Then
        strFilter = strFilter & "开嘱科室id = " & cbo开单科室.ItemData(cbo开单科室.ListIndex)
    End If
    
    If cboDoctor.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "申请人 = '" & Split(cboDoctor.Text, "-")(1) & "'"
    End If
    
    If cboDept.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "执行科室id = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
    End If
    
    If cboKind.Text <> "" And Trim(txtKind.Text) <> "" And tabReport.Tab = 1 Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        
        If cboKind.Text = "姓名" Then
            strFilter = strFilter & cboKind.Text & " like '" & Trim(txtKind.Text) & "%'"
        Else
            strFilter = strFilter & cboKind.Text & " = '" & Trim(txtKind.Text) & "'"
        End If
    End If
    
    For i = 0 To chkPatFrom.Count - 1
        If chkPatFrom(i).Value = 0 Then
            intPatFromCount = intPatFromCount + 1
        End If
    Next
    
    If intPatFromCount <> chkPatFrom.Count Then
        If chkPatFrom(0).Value = 0 Then
            strPatFrom = "来源 <> '门'"
        End If
    
        If chkPatFrom(1).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "来源 <> '住'"
        End If
    
        If chkPatFrom(2).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "来源 <> '外'"
        End If
    
        If chkPatFrom(3).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "来源 <> '体'"
        End If
        
        If strPatFrom <> "" Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & strPatFrom
        End If
    End If
    
    
    GetFilter = strFilter
    
    Exit Function
Errorhand:
    GetFilter = ""
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub LoadReport(strFilter As String, Optional ByVal blnRefreshFormCall As Boolean = False)
'blnRefreshFormCall 是否RefreshForm调用（等价于第一次加载数据）
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
On Error GoTo Errorhand
    If mblnIsConfiging Then Exit Sub
    
    If mlngViewReport = 0 Then
        strTmp = "b.报告状态 in (3,4)"
    Else
        strTmp = "b.报告状态 = 4"
    End If
    
    If tabReport.Tab = 0 Then
        'RIS使用自己的一套查询SQL
        If mblnRIS = True Then
            strSql = "Select * from (Select 2 as 类型,TO_CHAR(RAWTOHEX(b.检查报告id)) 报告ID,a.执行科室id, a.影像类别,a.接收日期 As 检查时间, c.医嘱内容, c.开嘱时间, " & _
                "  f.创建人 as 报告人,f.最后审核人 as 审核人, c.Id As 医嘱id,b.病历id " & _
                " From 影像检查记录 a, 病人医嘱报告 b, 病人医嘱记录 c, 影像检查记录 d, 病人医嘱记录 e,影像报告记录 f" & _
                " Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And a.医嘱id = f.医嘱id And e.Id = " & mlngAdviceId & " And b.医嘱id = c.Id " & _
                " And (c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null " & _
                " And b.检查报告id is not null And b.检查报告id =f.id And f.报告状态 in (2,3,4)" & _
                " Union All " & _
                "Select 1 as 类型,TO_CHAR(b.病历id) 报告ID,a.执行科室id, a.影像类别,a.接收日期 As 检查时间, c.医嘱内容, c.开嘱时间, " & _
                " a.报告人,a.复核人 审核人,c.Id As 医嘱id, b.病历id " & _
                " From 影像检查记录 a, 病人医嘱报告 b, 病人医嘱记录 c, 影像检查记录 d, 病人医嘱记录 e,病人医嘱发送 f" & _
                " Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id = " & mlngAdviceId & " And b.医嘱id = c.Id " & _
                " And a.医嘱ID = f.医嘱ID And (c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null " & _
                " And f.执行过程 > =4 And Nvl(Rawtohex(检查报告id), ' ') = ' ' )  Order By 检查时间 Asc "

        Else
            strSql = "Select* from (Select distinct m.类型, m.报告ID, m.执行科室id, m.查阅, m.打印, m.姓名,m.影像类别, m.医嘱内容,m.申请时间 , m.性别, m.年龄,  " & _
                "Decode(m.病人来源,1,'门',2,'住',3,'外','体') as 来源,   O.住院号, O.出院病床 as 床号, " & _
                "m.申请人, m.报告人, m.终审人, n.门诊号, n.就诊卡号, n.身份证号,m.病人id,m.医嘱id, m.主页ID, m.开嘱科室ID " & _
                "From (Select distinct 2 as 类型,TO_CHAR(RAWTOHEX(b.id)) 报告ID,a.执行科室id,decode(nvl(f.查阅状态,0),0,'','√') 查阅, " & _
                "decode(nvl(b.报告打印,0),0,'','√') 打印,a.姓名,a.性别,a.年龄,a.影像类别,c.医嘱内容,c.病人来源, c.主页ID, " & _
                "c.开嘱时间 as 申请时间,c.开嘱医生 申请人,b.创建人 报告人,b.最后审核人 终审人,c.病人id, c.Id As 医嘱id,c.开嘱科室ID " & _
                "From 影像检查记录 A, 影像报告记录 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E,病人医嘱报告 F " & _
                "Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And b.医嘱id = c.Id And b.医嘱id=f.医嘱id and b.id=f.检查报告id and " & _
                "(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null and " & strTmp & " and e.id = " & mlngAdviceId & " Union All " & _
                "Select distinct 1 as 类型,TO_CHAR(b.病历id) 报告ID,a.执行科室id,decode(nvl(b.查阅状态,0),0,'','√') 查阅, " & _
                "decode(nvl(a.报告打印,0),0,'','√') 打印,a.姓名,a.性别,a.年龄,a.影像类别,c.医嘱内容,c.病人来源,c.主页ID, " & _
                "c.开嘱时间 as 申请时间,c.开嘱医生 申请人,a.报告人,a.复核人 终审人,c.病人id, c.Id As 医嘱id,c.开嘱科室ID " & _
                "From 影像检查记录 A, 病人医嘱报告 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E " & _
                "Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And b.医嘱id = c.Id And b.病历ID Is Not Null And " & _
                "(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null and nvl(a.复核人,' ')<>' ' and e.id = " & mlngAdviceId & ") m,病人信息 n,病案主页 o,病人医嘱发送 p " & _
                "where m.病人id = n.病人id and m.病人id=o.病人id(+) and m.主页ID=o.主页ID(+) and m.医嘱id = p.医嘱id " & _
                IIf(mblnFirst Or blnRefreshFormCall, "", "and p.报到时间 between to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss') and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')") & ") " & _
                IIf(strFilter = "", "", " where " & strFilter)
        End If
    Else
        strSql = "Select * from (select distinct 2 as 类型,TO_CHAR(RAWTOHEX(f.id)) 报告ID, b.执行科室id,decode(nvl(g.查阅状态,0),0,'','√') 查阅," & _
                "decode(nvl(f.报告打印,0),0,'','√') 打印,c.姓名,e.影像类别,b.医嘱内容,b.开嘱时间 as 申请时间,c.性别,c.年龄,Decode(b.病人来源,1,'门',2,'住',3,'外','体') as 来源,d.住院号," & _
                "d.出院病床 as 床号, b.开嘱医生 申请人,f.创建人 as 报告人,f.最后审核人 终审人,c.门诊号, c.就诊卡号, c.身份证号,c.病人id,b.id as 医嘱ID,b.开嘱科室id " & _
                "from 病人医嘱发送 A, 病人医嘱记录 b, 病人信息 c, 病案主页 d, 影像检查记录 e,影像报告记录 f, 病人医嘱报告 g " & _
                "where a.医嘱id=b.id and b.病人id=c.病人id and b.病人id=d.病人id(+) and b.主页id=d.主页id(+) and b.id=e.医嘱id and e.医嘱id=f.医嘱id and f.id=g.检查报告id and b.id=g.医嘱id " & _
                "and a.报到时间 between trunc(to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')) and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss') " & _
                "Union All " & _
                "select distinct 1 as 类型,TO_CHAR(g.病历id) 报告ID, b.执行科室id,decode(nvl(g.查阅状态,0),0,'','√') 查阅," & _
                "decode(nvl(e.报告打印,0),0,'','√') 打印,c.姓名,e.影像类别,b.医嘱内容,b.开嘱时间 as 申请时间,c.性别,c.年龄,Decode(b.病人来源,1,'门',2,'住',3,'外','体') as 来源,d.住院号," & _
                "d.出院病床 as 床号, b.开嘱医生 申请人,e.报告人,e.复核人 终审人,c.门诊号, c.就诊卡号, c.身份证号,c.病人id,b.id as 医嘱ID,b.开嘱科室id " & _
                "from 病人医嘱发送 A, 病人医嘱记录 b, 病人信息 c, 病案主页 d, 影像检查记录 e, 病人医嘱报告 g " & _
                "where a.医嘱id=b.id and b.病人id=c.病人id and b.病人id=d.病人id(+) and b.主页id=d.主页id(+) and b.id=e.医嘱id and e.医嘱id=g.医嘱id and g.病历id is not null and b.id=g.医嘱id " & _
                "and a.报到时间 between trunc(to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')) and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')) " & _
                IIf(strFilter = "", "", " where " & strFilter)
    End If
    
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取报告信息") ' GetRecordset(strSql)
    Debug.Print rsTemp.RecordCount
    If mblnFirst Or blnRefreshFormCall Then rsTemp.Filter = "医嘱id=" & mlngAdviceId
    
    Call SetReportlistDataSource(rsTemp)
    Set mrsRecord = rsTemp
    
    Call InitReportList
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub GetFilterData()
    If mrsRecord Is Nothing Then
        Call LoadReport(GetFilter())
    Else
        mrsRecord.Filter = GetFilter()
        Call SetReportlistDataSource(mrsRecord)
    End If
End Sub

Private Sub chkPatFrom_Click(index As Integer)
    Call GetFilterData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call LoadReport(GetFilter())
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Errorhand
    mblnIsConfiging = True
    
    Call InitReportEditor
    
    If mblnShow Then
        Call InitCommandBars
        Call InitEdit
    End If
    
    '没有权限，或者是RIS调用，不显示全部报告页面
    If (InStr(gstrPrivs, VIEW_ALLREPORT) <= 0 Or mblnRIS = True) Then tabReport.TabVisible(1) = False
    
    picReportControl.Visible = Not mblnRIS
        
    mblnIsConfiging = False
    mblnFirst = True
    
    If Trim(mstrReportId) <> "" Then
        Call RefreshReportContent(mstrReportId)
    Else
        Call LoadReport("")
    End If
    
    mblnFirst = False
    
    Exit Sub
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBarControl.Left = 0
    picBarControl.Top = 0
    picBarControl.Width = ScaleWidth
    
    If Not mblnShow Then
        picVLine.Visible = False
        picReportList.Visible = False
        picBarControl.Visible = False
        picBarControl.Height = 0
    End If
    
    picBack.Left = 0
    picBack.Top = picBarControl.Height
    picBack.Width = ScaleWidth
    picBack.Height = ScaleHeight - picBarControl.Height
    
    picVLine.Top = ScaleTop
    picVLine.Height = picBack.ScaleHeight
    If picVLine.Left < 500 Then picVLine.Left = 500
    If picVLine.Left > picBack.ScaleWidth - 500 Then picVLine.Left = picBack.ScaleWidth - 500
    
    picReportList.Left = ScaleLeft
    picReportList.Width = picVLine.Left - picReportList.Left
    picReportList.Top = picBack.ScaleTop
    picReportList.Height = picBack.ScaleHeight - picReportList.Top
    
    picReportContent.Left = IIf(picReportList.Visible, picVLine.Left + picVLine.Width, 0)
    
    picReportContent.Width = picBack.ScaleWidth - IIf(picReportList.Visible, picReportContent.Left, 0)
    picReportContent.Top = picBack.ScaleTop
    picReportContent.Height = picBack.ScaleHeight - picReportContent.Top
    
    Call picReportList_Resize
End Sub

Private Sub picVLine_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    If Button = 1 Then
        picVLine.Left = picVLine.Left + x
        
        Form_Resize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errorhand
    
    Set mobjReport = Nothing
    
    If Not mobjFtp Is Nothing Then
        If mobjFtp.hConnection <> 0 Then mobjFtp.FuncFtpDisConnect
        Set mobjFtp = Nothing
    End If
    
    Set mrsRecord = Nothing
    
    Exit Sub
    
    zlDocEditor.ClearAll
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub optTime_Click(index As Integer)
On Error Resume Next
    dtpEnd.Value = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    
    If index = 0 Then
        dtpStart = Format(Now - 3, "yyyy-mm-dd HH:MM")
    ElseIf index = 1 Then
        dtpStart = Format(Now - 7, "yyyy-mm-dd HH:MM")
    ElseIf index = 2 Then
        dtpStart = Format(Now - 15, "yyyy-mm-dd HH:MM")
    Else
        dtpStart = Format(Now - 30, "yyyy-mm-dd HH:MM")
    End If
    
    Call LoadReport(GetFilter())
End Sub

Private Sub picReportContent_Resize()
    On Error Resume Next
    
    zlDocEditor.Left = 0
    zlDocEditor.Top = 0
    zlDocEditor.Width = picReportContent.Width
    zlDocEditor.Height = picReportContent.Height
    
    rtxtReport.Left = 0
    rtxtReport.Top = 0
    rtxtReport.Height = picReportContent.Height
    rtxtReport.Width = picReportContent.Width
End Sub

Private Sub picReportList_Resize()
    On Error Resume Next
    
    tabReport.Left = 20
    tabReport.Top = 20
    tabReport.Width = picReportList.ScaleWidth - 40
    
    picReportControl.Left = tabReport.Left
    picReportControl.Top = tabReport.Top + tabReport.Height
    picReportControl.Width = tabReport.Width
    
    vsfReportList.Left = picReportControl.Left
    vsfReportList.Top = picReportControl.Top + picReportControl.Height
    vsfReportList.Width = picReportList.ScaleWidth - 20
    
    If Not mblnShow Then
        picReportControl.Enabled = False
        tabReport.Enabled = False
    End If
	
	vsfReportList.Height = picReportList.ScaleHeight - picReportControl.Top - picReportControl.Height 
End Sub

Private Sub UpdateReportViewState(ByVal strDocId As String)
    Dim strSql As String
    
On Error GoTo Errorhand
    
    strSql = "Zl_影像报告查阅记录_Insert(" & mlngAdviceId & ", '" & strDocId & "')"
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
    
    If vsfReportList.RowSel <= 0 Then Exit Sub
    vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "查阅")) = "√"
    
    Exit Sub
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub tabReport_Click(PreviousTab As Integer)
    '姓名条件，只有全院查询才可用
    cboKind.Enabled = tabReport.Tab = 1
    txtKind.Enabled = cboKind.Enabled
    txtKind.Text = ""
    
    If cbo开单科室.ListCount > 0 Then cbo开单科室.ListIndex = 0
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    
    '切换页面刷新数据
    Call LoadReport(GetFilter)
End Sub

Private Sub txtKind_Change()
    Call GetFilterData
End Sub

Private Sub txtKind_KeyPress(KeyAscii As Integer)
    If cboKind.Text = "门诊号" Or cboKind.Text = "住院号" Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsfReportList_SelChange()
    Dim strCurReportId As String
On Error GoTo ErrorHnad
    If vsfReportList.RowSel <= 0 Then
        If vsfReportList.Rows > 1 Then
            vsfReportList.RowSel = 1
        End If
        Exit Sub
    End If
    
    strCurReportId = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, 1)
    
    Call ViewReportContent(strCurReportId)
     
    If mblnAutoView And vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "查阅")) <> "√" Then
        Call UpdateReportViewState(strCurReportId)
    End If
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub ViewReportContent(ByVal strDocId As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intReportEditor As Integer
    Dim strContent As String
    Dim intLoop As Integer
    
    If strDocId = "" Then
        Call ClearReportContent
        Exit Sub
    End If
    
    If Len(strDocId) = 32 Then
        rtxtReport.Visible = False
        zlDocEditor.Visible = True
        mblnIsNewReport = True
        
        strSql = "Select Length(a.报告内容.GetClobVal()) as ContentLength From 影像报告记录 a Where a.ID = '" & strDocId & "'"
        Set rsTemp = GetRecordset(strSql)
        
        If rsTemp.BOF = False Then
            If rsTemp("ContentLength").Value > 2000 Then
                For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
                    strSql = "select to_char(substr(a.报告内容.GetClobVal()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                             " from 影像报告记录 a where a.ID = '" & strDocId & "'"
                             
                    Set rsTemp = GetRecordset(strSql)
                    
                    If rsTemp.EOF = False Then
                        strContent = strContent & Nvl(rsTemp("send_content").Value)
                    End If
                Next
            Else
                strSql = "Select a.报告内容.GetClobVal() as send_content From 影像报告记录 a Where a.ID = '" & strDocId & "'"
                
                Set rsTemp = GetRecordset(strSql)
                    
                If rsTemp.EOF = False Then
                    strContent = Nvl(rsTemp("send_content").Value)
                End If
            End If
        End If
        
        If strContent = "" Then
            MsgBox "报告内容不存在。"
            Exit Sub
        End If
        
        Call LoadReportContent(strContent, strDocId)
    Else
        rtxtReport.Text = ""
        rtxtReport.TextRTF = ""
        rtxtReport.Visible = True
        zlDocEditor.Visible = False
        mblnIsNewReport = False
        If vsfReportList.RowSel <= 0 Then Exit Sub
        
        '老版报告
        '判断使用电子病历编辑器还是Pscs编辑器
        strSql = "select 参数值 from 病人医嘱报告 a,影像检查记录 b,影像流程参数 c " & _
                 "where a.医嘱id = b.医嘱id and b.执行科室id = c.科室id and a.病历id=" & strDocId & " and c.参数名='查看历史报告'"
        
        Set rsTemp = GetRecordset(strSql)
        
        If rsTemp.RecordCount > 0 Then intReportEditor = Nvl(rsTemp!参数值, "1")
        
        If intReportEditor = 0 Then
            Call LoadRichReportContent(strDocId)
        Else
            Call LoadPacsReportContent(strDocId)
        End If
    End If
End Sub

Private Sub LoadRichReportContent(ByVal strDocId As String)
    Dim tmpPath As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strBlobFile As String
    
    tmpPath = App.Path()
    
    If objFSO.FileExists(tmpPath & "\tmp.zip") Then Call objFSO.DeleteFile(tmpPath & "\tmp.zip")
    If objFSO.FileExists(tmpPath & "\tmp.rtf") Then Call objFSO.DeleteFile(tmpPath & "\tmp.rtf")
        
    strBlobFile = gobjComLib.Sys.ReadLob(100, 5, strDocId, Replace(tmpPath & "\tmp.zip", "\\", "\"))
    If objFSO.FileExists(strBlobFile) = False Then
        strBlobFile = gobjComLib.Sys.ReadLob(100, 5, strDocId, Replace(tmpPath & "\tmp.zip", "\\", "\"), 0, 1)
    End If
     
    
    If objFSO.FileExists(strBlobFile) Then
        '解压并加载显示
        Dim cc As New cUnzip
        
        cc.ZipFile = strBlobFile
        cc.UnzipFolder = tmpPath
        cc.Unzip
        Set cc = Nothing
        
        If objFSO.FileExists(tmpPath & "\tmp.rtf") Then
            Call rtxtReport.LoadFile(tmpPath & "\tmp.rtf")
        End If
    End If
End Sub

Private Sub LoadPacsReportContent(ByVal strDocId As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext As String
    Dim blnShow As Boolean
    Dim strOffset As String
    Dim strTitle As String
    Dim strText As String
    Dim str检查所见 As String
    Dim str诊断意见 As String
    Dim str建议 As String
    
    strOffset = "  "
    
    '读取报告标题定义
    strSql = "Select 参数值,参数名 From 病人医嘱报告 a, 影像检查记录 b, 影像流程参数 c Where a.医嘱id = b.医嘱id " & _
            " And b.执行科室id = c.科室id And a.病历id = " & Val(strDocId)
    Set rsTemp = GetRecordset(strSql)
    If Not rsTemp.EOF Then
        rsTemp.Filter = "参数名='检查所见名称'"
        If Not rsTemp.EOF Then
            str检查所见 = rsTemp!参数值
        Else
            str检查所见 = "检查所见"
        End If
        
        rsTemp.Filter = "参数名='诊断意见名称'"
        If Not rsTemp.EOF Then
            str诊断意见 = rsTemp!参数值
        Else
            str诊断意见 = "诊断意见"
        End If
        
        rsTemp.Filter = "参数名='建议名称'"
        If Not rsTemp.EOF Then
            str建议 = rsTemp!参数值
        Else
            str建议 = "建议"
        End If
    End If
    
    '读取报告的内容
    strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = '" & strDocId & "' And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0  "
    Set rsTemp = GetRecordset(strSql)
    
    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs20 "
                
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!标题
            Case "检查所见"
                strTitle = str检查所见
                strText = vbCrLf & strOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
            Case "诊断意见"
                strTitle = str诊断意见
                strText = vbCrLf & strOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
            Case "建议"
                strTitle = str建议
                strText = vbCrLf & strOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then
            strFormatContext = strFormatContext & "\b\cf2" & strTitle & "\par\b0\cf0" & Replace(strText, vbCrLf, "\par\cf0") & "\par"
        End If
        rsTemp.MoveNext
    Wend
    
    strFormatContext = strFormatContext & "}"
    
    rtxtReport.TextRTF = strFormatContext
End Sub

Public Function zlDocGetList(ByVal lngPatId As Long, Optional ByVal lngPageId As Long, Optional ByVal strRegNo As String) As Recordset
'返回Pacs文档编辑器书写的报告列表
'lngPatID:病人ID
'lngPageID:主页ID
'strRegNo:挂号单
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrH
    
    If Trim(strRegNo) <> "" Then
        strSql = "Select A.ID As 医嘱ID, RawToHex(B.检查报告ID) As 报告ID, C.文档标题, C.报告状态, C.最后编辑时间, C.最后编辑人 " & _
                 "From 病人医嘱记录 A, 病人医嘱报告 B, 影像报告记录 C " & _
                 "Where A.病人ID=" & lngPatId & " And A.挂号单 = '" & strRegNo & "' And A.ID = B.医嘱ID And B.检查报告ID = C.ID"
    Else
        strSql = "Select A.ID As 医嘱ID, RawToHex(B.检查报告ID) As 报告ID, C.文档标题, C.报告状态, C.最后编辑时间, C.最后编辑人 " & _
                 "From 病人医嘱记录 A, 病人医嘱报告 B, 影像报告记录 C " & _
                 "Where A.病人ID=" & lngPatId & " And A.主页ID = " & lngPageId & " And A.ID = B.医嘱ID And B.检查报告ID = C.ID"
    End If
    
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "根据病人获取报告列表") ' GetRecordset(strSql)
    
    Set zlDocGetList = rsTemp
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetList = Nothing
End Function

Public Function zlDocGetListWithAdvice(ByVal strAdviceId As String) As Recordset
'返回Pacs文档编辑器书写的报告列表
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrH
    
    If Trim(strAdviceId) = "" Then Exit Function
         
    strAdviceId = Replace(strAdviceId, "，", ",")
    
    strSql = "Select A.医嘱ID, RawToHex(A.检查报告ID) As 报告ID, B.文档标题, B.报告状态, B.最后编辑时间, B.最后编辑人 " & _
             "From 病人医嘱报告 A, 影像报告记录 B, Table(Cast(f_Str2list('" & strAdviceId & "') As zlTools.t_Strlist)) C " & _
             "Where A.医嘱ID = C.Column_Value And  A.检查报告ID = B.ID "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "根据医嘱获取报告列表") 'GetRecordset(strSql)
    
    Set zlDocGetListWithAdvice = rsTemp
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetListWithAdvice = Nothing
End Function

Public Sub zlDocViewStateUpdate(ByVal blnState As Boolean, ByVal lngAdviceID As Long)
'更新pacs文档编辑器报告的查阅状态,如果blnState=True，表示对应的所有报告为已阅，blnState=False,表示为未阅
'lngPatID:病人ID
'lngPageID:主页ID
'strRegNo:挂号单
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strUserName As String
    
On Error GoTo ErrH
    
    strSql = "Select RawToHex(B.检查报告ID) As 报告ID From 病人医嘱报告 B Where 医嘱ID = " & lngAdviceID
             
    Set rsTemp = GetRecordset(strSql)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If GetUserInfo Then strUserName = UserInfo.姓名
    
    Do While Not rsTemp.EOF
        If Nvl(rsTemp!报告ID) <> "" Then
            If blnState Then
                strSql = "Zl_影像报告查阅记录_Insert(" & lngAdviceID & ", '" & Nvl(rsTemp!报告ID) & "')"
            Else
                strSql = "Zl_影像报告查阅记录_Cancel(" & lngAdviceID & ", '" & Nvl(rsTemp!报告ID) & "','" & strUserName & "')"
            End If
            
            Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
        End If
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Function GetReportImage(ByVal lngAdviceID As Long, ByVal strDocReportId As String, _
    strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'功能:将检查医嘱所对应报告中的所有图像，保存到指定目录,获取PACS报告中的所有图像
'lngAdviceId:医嘱id
'strTmpImgFolder:图像缓存目录
'objImgFileName:报告图像文件名集合

'说明：此过程目前只有体检调用，一个医嘱对应一份报告
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strContent As String
    Dim strDocId As String
    Dim intLoop As Integer
    Dim strFilter As String
    
    If (lngAdviceID <= 0) Then Exit Function
        
    If Trim(strDocReportId) = "" Then
        strFilter = "a.医嘱ID=" & lngAdviceID & ""
    Else
        strFilter = "a.ID=HexToRaw('" & strDocReportId & "')"
    End If
    
    strSql = "Select Length(a.报告内容.GetClobVal()) as ContentLength, A.ID From 影像报告记录 a Where " & strFilter
    Set rsTemp = GetRecordset(strSql)
    
    If rsTemp.RecordCount <= 0 Then Exit Function

    strDocId = Nvl(rsTemp!ID)
    
    If rsTemp("ContentLength").Value > 2000 Then
        For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
            strSql = "select to_char(substr(a.报告内容.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                     " from 影像报告记录 a where " & strFilter
                     
            Set rsTemp = GetRecordset(strSql)
            
            If rsTemp.EOF = False Then
                strContent = strContent & Nvl(rsTemp("send_content").Value)
            End If
        Next
    Else
        strSql = "Select a.报告内容.getclobval() as send_content From 影像报告记录 a Where " & strFilter
        
        Set rsTemp = GetRecordset(strSql)
            
        If rsTemp.EOF = False Then
            strContent = Nvl(rsTemp("send_content").Value)
        End If
    End If
    
    If strContent = "" Then
        strMsg = "报告内容不存在。"
        Exit Function
    End If
    
    '初始化FTP相关信息
    If InitFtpInfo(strDocId) = False Then
        strMsg = "获取图象的FTP信息失败"
        Exit Function
    End If
    
    Set GetReportImage = LoadImageFromXml(strContent, strTmpImgDir, strMsg)
End Function

Public Sub RefreshForm(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional objParent As Object)
    '功能：刷新窗体的参数
    '参数： lngAdviceId -- 医嘱ID
    '       strReportId -- 报告ID
    '       blnAutoView -- 是否自动查阅
    '       objParent -- 父窗口
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mlngAdviceId = lngAdviceID
    mstrReportId = strReportId
    
    '如果是RIS查看，设置开始时间
    If mblnRIS = True Then
        dtpStart.Value = Format(Now - 365, "yyyy-mm-dd HH:MM")
    End If
    
    '如果传入了报告ID，则只显示一份报告，不再显示左边的过滤条件和菜单
    If strReportId <> "" Then mblnShow = False
    
    strSql = "Select 参数值 From 影像流程参数 a, 影像检查记录 b " & _
             "Where a.科室ID = b.执行科室id And a.参数名 = '医生站查看报告' And b.医嘱id = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
    
    If rsTemp.RecordCount > 0 Then
        mlngViewReport = Val(Nvl(rsTemp!参数值))
    Else
        mlngViewReport = 1
    End If
    
    Call LoadReport(GetFilter, True)
    
    Exit Sub
err:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub ClearReportContent()
    '功能：清空报告编辑器
    rtxtReport.Visible = True
    zlDocEditor.Visible = False
    rtxtReport.TextRTF = ""
    rtxtReport.Text = ""
End Sub

Private Sub SetReportlistDataSource(rsData As ADODB.Recordset)
    '功能：设置报告列表的数据源，同时刷新报告编辑器内容
    '参数：rsData -- 数据源
    
    Set vsfReportList.DataSource = rsData
    '如果有查询结果，强制刷新一次内容
    If rsData.EOF = False Then
        Call vsfReportList_SelChange
    Else
        '清空编辑器
        Call ClearReportContent
    End If
End Sub
