VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "报表管理"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRPT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1095
      ScaleWidth      =   1935
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1935
      Begin XtremeSuiteControls.TabControl tbcRPT 
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   1296
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   7080
      ScaleHeight     =   4455
      ScaleWidth      =   3015
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
      Begin VB.PictureBox picGroup_S 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   120
         MousePointer    =   7  'Size N S
         ScaleHeight     =   60
         ScaleWidth      =   2535
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGroup 
         DragIcon        =   "frmMain.frx":058A
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989546359
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
      Begin VSFlex8Ctl.VSFlexGrid vsfGroupDetail 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989543184
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
      Begin VB.Label lblGroupDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报表组成员"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   900
      End
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3840
      ScaleHeight     =   1095
      ScaleWidth      =   3015
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfReport 
         DragIcon        =   "frmMain.frx":0CF4
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989543184
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
   Begin VB.PictureBox picClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   1215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
      Begin XtremeReportControl.ReportControl rptClass 
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
         _Version        =   589884
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   0
         MultipleSelection=   0   'False
         ShowHeader      =   0   'False
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   1905
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1905
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Enter键：查找；F3键：继续查找"
         Top             =   0
         Width           =   1875
      End
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6570
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":145E
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
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
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
   Begin MSComDlg.CommonDialog cdg 
      Left            =   1320
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1920
      Top             =   240
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
            Picture         =   "frmMain.frx":1CF2
            Key             =   "rpt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":208C
            Key             =   "rpt_ena"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2426
            Key             =   "rpt_dis"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27C0
            Key             =   "grp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B5A
            Key             =   "grp_ena"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EF4
            Key             =   "grp_dis"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":328E
      Left            =   480
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enuMenus
    文件 = 1
        参数设置 = 181
        导出报表 = 121
        导入报表 = 122
        导出全部 = 123
        导入全部 = 124
        退出 = 2613
    编辑 = 2
        新增报表类 = 3051
        修改报表类 = 3053
        删除报表类 = 3054
        新增报表组 = 6861
        修改报表组 = 6862
'        撤销报表组 = 6881
        删除报表组 = 6863
        新增报表 = 3001
        修改报表 = 3003
'        移入报表 = 3052
'        移出报表 = 3036
        删除报表 = 3004
        设计报表 = 4113
        报表向导 = 3551
        执行报表 = 3010
        报表启用 = 8106
        报表停用 = 8099
    工具 = 5
        报表发布 = 741
            至导航台菜单 = 100501
            至模块内菜单 = 100502
        取消发布 = 742
            从导航台菜单 = 100511
            从模块内菜单 = 100512
        性能检查 = 100521
        清除历史数据源 = 100522
        报表运行日志 = 100523
    查看 = 7
        工具栏 = 701
            标准按钮 = 702
            文本标签 = 703
            大图标 = 704
        状态栏 = 711
        字体大小 = 721
            小字体 = 722
            大字体 = 723
        查找 = 721
        刷新 = 791
        显示所有分类下级 = 751
        仅显示停用状态 = 7510
        显示独立报表 = 752
        显示子报表 = 753
    帮助 = 9
        帮助主题 = 901
        WEB上的中联 = 911
            中联主页 = 912
            中联论坛 = 913
            发送反馈 = 914
        关于 = 991
    其他 = 10
        选择系统标签 = 1001
        选择系统控件 = 1002
        查找报表标签 = 1003
        查找报表控件 = 1004
        TabRPT_1 = 1011
        TabRPT_2 = 1012
End Enum

Private Const MSTR_REPORT_COLS = _
    "编号,,3,2000|ID,,0,0,n|名称,,3,2500|说明,,3,3000|程序ID,,0,0,n|修改时间,,3,2000,DT|发布时间,,3,2000,DT|系统,,0,0|" & _
    "最后执行时间,,3,2000,DT|最后执行人,,3,1000|种类,,3,1000|类型,,3,1000|报表分类,,3,1500|性能检查结果,,3,2000|" & _
    "所属报表组,,3,2000|其他数据连接,,3,2000|分类ID,,0,0,n|停用,,0,0,n"
Private Const MSTR_GROUP_COLS = _
    "编号,,3,2000|组名,,3,2500|说明,,3,6000|报表分类,,3,1500|ID,,0,0,n|发布时间,,3,2000,DT|程序ID,,0,0,n|分类ID,,0,0,n|" & _
    "停用,,0,0,n"
Private Const MSTR_GROUPDETAIL_COLS = _
    "编号,,3,2000|ID,,0,0,n|名称,,3,2500|说明,,3,3000|程序ID,,0,0,n|修改时间,,3,2000,DT|发布时间,,3,2000,DT|系统,,0,0|" & _
    "最后执行时间,,3,2000,DT|最后执行人,,3,1000|种类,,3,1000|类型,,3,1000|其他数据连接,,3,2000|停用,,0,0,n"

Private WithEvents mobjClass As clsReportControlEx
Attribute mobjClass.VB_VarHelpID = -1
Private WithEvents mobjReport As clsVSFlexGridEx
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mobjGroup As clsVSFlexGridEx
Attribute mobjGroup.VB_VarHelpID = -1
Private WithEvents mobjSub As clsVSFlexGridEx
Attribute mobjSub.VB_VarHelpID = -1

Private mbytFontSize As Byte                                            '1-大字体；0-小字体
Private mbytReportGroup As Byte                                         '1-显示独立报表；0-显示子报表
Private mblnDisplayChild As Boolean                                     'True-显示所有子结点的项目；False-显示当前结点的项目
Private mblnDisable As Boolean                                          'True-报表停用
Private mblnMemory As Boolean                                           '个性化界面
Private mblnAllowPublish As Boolean                                     'True-允许发布到模块；False-不允许
Private mblnReportControlFocus As Boolean                               'ReportControl焦点无响应的替代变量
Private mcbpReport As XtremeCommandBars.CommandBarPopup                 '报表的弹出菜单
Private mobjFindPreviors As Object                                      '上次查找的对象
Private mlngFindRow As Long                                             '上次查找定位的行号
Private mintFind As Integer
Private mblnEnter As Boolean

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As Object
    Dim i As Integer
    Dim lngID As Long
    Dim blnTemp As Boolean
    Dim strPath As String, strTemp As String
    
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.id
    Case enuMenus.执行报表
        Call GetVsfControl(lngID, blnTemp)
        If lngID > 0 Then
            If blnTemp Then
                '报表组
                ''检查所属子报表的权限
                For i = 1 To vsfGroupDetail.Rows - 1
                    If mdlPublic.CheckReportPriv(lngID, True) = False Then
                        MsgBox mdlPublic.FormatString("你没有权限查询报表【[1]】中某些数据源的对象！" _
                                            , Val(vsfGroupDetail.TextMatrix(i, vsfGroupDetail.ColIndex("ID")))) _
                            , vbInformation, App.Title
                        Exit Sub
                    End If
                Next
            Else
                '报表
                If mdlPublic.CheckReportPriv(lngID) = False Then
                    MsgBox "你没有权限查询该报表某些数据源中的对象！", vbInformation, App.Title
                    Exit Sub
                End If
            End If
            
            '执行
            If blnTemp Then
                '报表组
                Set gobjReport = Nothing
                glngGroup = lngID
            Else
                '报表
                If mdlPublic.CheckPass(lngID) = False Then
                    MsgBox "报表数据错误，不能执行该报表！", vbInformation, App.Title
                    Exit Sub
                End If
                
                glngGroup = 0
                Set gobjReport = Nothing
                Set gobjReport = mdlPublic.ReadReport(lngID)
            End If
            
            '使用缺省参数
            garrPars = Array()
            If Not mdlPublic.ShowReport(Me) Then MsgBox "报表打开失败！", vbInformation, App.Title
        End If
    Case enuMenus.参数设置
        If frmReportPara.ShowMe(Me) Then
            '更新参数
            Call mdlPublic.InitPar
        End If
    Case enuMenus.性能检查
        Call CheckSQLPlanEx
    Case enuMenus.导出报表, enuMenus.导出全部
        Call Export(Control.id)
    Case enuMenus.导入报表, enuMenus.导入全部
        Call Import(Control.id)
    Case enuMenus.退出
        Unload Me
    Case enuMenus.新增报表类, enuMenus.新增报表组, enuMenus.新增报表
        mblnReportControlFocus = enuMenus.新增报表类 = Control.id
        Call NewEx
    Case enuMenus.修改报表类, enuMenus.修改报表组, enuMenus.修改报表
        mblnReportControlFocus = enuMenus.修改报表类 = Control.id
        Call Modify
    Case enuMenus.删除报表类, enuMenus.删除报表组, enuMenus.删除报表
        mblnReportControlFocus = enuMenus.删除报表类 = Control.id
        Call Delete(Control.id)
    Case enuMenus.设计报表
        Call Design
    Case enuMenus.报表启用
        Call StateSwitch(Control.id, True)
    Case enuMenus.报表停用
        Call StateSwitch(Control.id)
    Case enuMenus.清除历史数据源
        frmClearHistory.Show vbModal, Me
    Case enuMenus.报表向导
        Call Guide
    Case enuMenus.至导航台菜单
        Call ReportGrantToNavigator
    Case enuMenus.至模块内菜单
        Call ReportGrantToModule
    Case enuMenus.从导航台菜单
        Call ReportRevokeFromNavigator
    Case enuMenus.从模块内菜单
        Call ReportRevokeFromModule
    Case enuMenus.查找
        If txtFind.Visible And txtFind.Enabled Then
            txtFind.SetFocus
        End If
    Case enuMenus.查找报表控件
        Call Find(txtFind.Text)     '查找下一个匹配项
    Case enuMenus.报表运行日志
        Call ShowRunLog
    Case enuMenus.标准按钮
        cbsMain(Val("2-工具栏")).Visible = Not cbsMain(Val("2-工具栏")).Visible
        cbsMain.RecalcLayout
    Case enuMenus.文本标签
        For Each objControl In cbsMain(Val("2-工具栏")).Controls
            If UCase(TypeName(objControl)) = UCase("ICommandBarButton") _
                Or UCase(TypeName(objControl)) = UCase("ICommandBarPopup") Then
                objControl.Style = IIF(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        cbsMain.RecalcLayout
    Case enuMenus.大图标
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    Case enuMenus.小字体
        If mbytFontSize <> 0 Then Call SetControlFontSize(0)
        mbytFontSize = 0
    Case enuMenus.大字体
        If mbytFontSize <> 1 Then Call SetControlFontSize(1)
    Case enuMenus.状态栏
        staMain.Visible = Not Control.Checked
        cbsMain.RecalcLayout
    Case enuMenus.刷新
        rptClass.Tag = ""
        Call RefreshEx
    Case enuMenus.显示所有分类下级
        mblnDisplayChild = Not mblnDisplayChild
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.仅显示停用状态
        mblnDisable = Not mblnDisable
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.显示独立报表
        mbytReportGroup = 0
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.显示子报表
        mbytReportGroup = 1
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.帮助主题
        Call mdlPublic.ShowHelpRpt(Me.hwnd, "main", 0)
    Case enuMenus.中联主页
        Call mdlPublic.zlHomePage(Me.hwnd)
    Case enuMenus.中联论坛
        Call mdlPublic.zlWebForum(Me.hwnd)
    Case enuMenus.发送反馈
        Call mdlPublic.zlMailTo(Me.hwnd)
    Case enuMenus.关于
        Call mdlPublic.ShowAbout(Me)
    Case enuMenus.选择系统控件
        Call SelectedSysComboBox(Control)
    Case enuMenus.TabRPT_1, enuMenus.TabRPT_2
        tbcRPT.Item(Control.id - enuMenus.TabRPT_1).Selected = True
    End Select
    mblnReportControlFocus = False
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If staMain.Visible Then
        Bottom = staMain.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnPublication As Boolean
    
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.id
    Case enuMenus.执行报表
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        Select Case UCase(Me.ActiveControl.name)
        Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
            If tbcRPT.Selected.Index = Val("0-报表页面") Then
                Control.Enabled = vsfReport.Row > 0
            Else
                Control.Enabled = vsfGroup.Row > 0 Or vsfGroupDetail.Row > 0
            End If
        Case Else
            Control.Enabled = False
        End Select
    Case enuMenus.新增报表类
        Control.Enabled = mblnReportControlFocus
    Case enuMenus.修改报表类, enuMenus.删除报表类
        Control.Enabled = mblnReportControlFocus And rptClass.SelectedRows.count > 0
        If Control.Enabled Then
            Control.Enabled = Nvl(rptClass.FocusedRow.Record(mobjClass.GetColIndex("名称")).Value) <> "所有"
        End If
    Case enuMenus.新增报表组
        Control.Enabled = tbcRPT.Selected.Index = Val("1-报表组页面") And glngSys = 0
    Case enuMenus.删除报表组
        If Not Me.ActiveControl Is Nothing Then
            Control.Enabled = UCase(Me.ActiveControl.name) = "VSFGROUP" And glngSys = 0
            If Control.Enabled Then
                Control.Enabled = Me.ActiveControl.Rows > 1
            End If
        End If
    Case enuMenus.修改报表组
        If Not Me.ActiveControl Is Nothing Then
            Control.Enabled = UCase(Me.ActiveControl.name) = "VSFGROUP"
            If Control.Enabled Then
                Control.Enabled = Me.ActiveControl.Rows > 1
            End If
        End If
    Case enuMenus.新增报表
        Control.Enabled = glngSys = 0
    Case enuMenus.修改报表
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0
            ElseIf UCase(Me.ActiveControl.name) = "VSFGROUPDETAIL" Then
                Control.Enabled = vsfGroupDetail.Row > 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.删除报表
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0 And glngSys = 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.设计报表
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0
            ElseIf UCase(Me.ActiveControl.name) = "VSFGROUPDETAIL" Then
                Control.Enabled = vsfGroupDetail.Row > 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.报表启用, enuMenus.报表停用
        If Not Me.ActiveControl Is Nothing Then
            Select Case UCase(ActiveControl.name)
            Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
                blnPublication = ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("发布时间")) <> "" _
                                And glngSys = 0
                If blnPublication Then
                    If Control.id = enuMenus.报表启用 Then
                        blnPublication = Val(ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("停用"))) = 1
                    Else
                        blnPublication = Val(ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("停用"))) <> 1
                    End If
                End If
            Case Else
                blnPublication = False
            End Select
            Control.Enabled = blnPublication
        End If
    Case enuMenus.性能检查
        Control.Enabled = tbcRPT.Selected.Index = Val("0-报表页面")
    Case enuMenus.标准按钮
        Control.Checked = cbsMain(2).Visible
    Case enuMenus.文本标签
        Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption _
                        Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
    Case enuMenus.大图标
        Control.Checked = cbsMain.Options.LargeIcons
    Case enuMenus.小字体
        Control.IconId = IIF(mbytFontSize = 0, 90004, 90003)
    Case enuMenus.大字体
        Control.IconId = IIF(mbytFontSize = 1, 90004, 90003)
    Case enuMenus.状态栏
        Control.Checked = staMain.Visible
    Case enuMenus.显示所有分类下级
        Control.Checked = mblnDisplayChild
    Case enuMenus.仅显示停用状态
        Control.Checked = mblnDisable
    Case enuMenus.显示独立报表
        Control.IconId = IIF(mbytReportGroup = 0, 90004, 90003)
    Case enuMenus.显示子报表
        Control.IconId = IIF(mbytReportGroup = 1, 90004, 90003)
    Case enuMenus.报表运行日志
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        Select Case UCase(Me.ActiveControl.name)
        Case "VSFREPORT", "VSFGROUPDETAIL"
            If tbcRPT.Selected.Index = Val("0-报表页面") Then
                Control.Enabled = vsfReport.Row > 0
            Else
                Control.Enabled = vsfGroupDetail.Row > 0
            End If
        Case Else
            Control.Enabled = False
        End Select
    Case enuMenus.至导航台菜单, enuMenus.至模块内菜单, enuMenus.从导航台菜单, enuMenus.从模块内菜单
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        If glngSys = 0 Then
            Select Case UCase(Me.ActiveControl.name)
            Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
                If tbcRPT.Selected.Index = Val("0-报表页面") Then
                    Control.Enabled = vsfReport.Row > 0
                Else
                    Control.Enabled = vsfGroup.Row > 0 Or vsfGroupDetail.Row > 0
                End If
                If Control.Enabled Then
                    If Control.id = enuMenus.从导航台菜单 Or Control.id = enuMenus.从模块内菜单 Then
                        Control.Enabled = ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("发布时间")) <> ""
                        If Control.id = enuMenus.从模块内菜单 Then
                            '报表组不允许发布到模块，也就没有取消发布
                            Control.Enabled = Control.Enabled And UCase(Me.ActiveControl.name) <> "VSFGROUP"
                        End If
                    ElseIf Control.id = enuMenus.至模块内菜单 Then
                        '报表组不允许发布到模块
                        Control.Enabled = UCase(Me.ActiveControl.name) <> "VSFGROUP"
                    End If
                End If
            Case Else
                Control.Enabled = False
            End Select
        Else
            Control.Enabled = False
        End If
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.handle = picClass.hwnd
    Case 2
        Item.handle = picRPT.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnEnter Then
        Call SetControlFontSize(mbytFontSize)       '字体大小
        mblnEnter = False
    End If
End Sub

Private Sub Form_Load()
    Dim objControl As Object
    Dim strPane As String, strRegPath As String
    Dim lngTmp As Long
    
    mblnEnter = False
    mblnReportControlFocus = False
    strRegPath = mdlPublic.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.name)

    '获取参数值
    mblnMemory = mdlPublic.GetMemoryParam()
    mblnDisplayChild = Val(GetSetting("ZLSOFT", strRegPath, "显示所有分类下级")) = 1
    mblnDisable = Val(GetSetting("ZLSOFT", strRegPath, "仅显示停用状态")) = 1
    mbytReportGroup = Val(GetSetting("ZLSOFT", strRegPath, "显示报表类别"))
    mbytFontSize = Val(GetSetting("ZLSOFT", strRegPath, "字体大小"))
    strPane = GetSetting("ZLSOFT", strRegPath, "布局")
    
    Call InitOther
    Call InitCommandBars
    Call InitDockPane
    Call InitTabControl
    Call InitReportControl
    Call InitVSF
    
    Call FillData(Val("5-cboSystem"))
    Call FillData(Val("1-rptClass"), True)
    If tbcRPT.Selected.Index = Val("0-报表页面") Then
        Call FillData(Val("2-vsfReport"), True)
    Else
        Call FillData(Val("3-vsfGroup"), True)
        Call FillData(Val("4-vsfGroupDetial"), True)
    End If
    
    '恢复上次界面
    If mblnMemory Then
        mdlPublic.RestoreWinState Me, App.ProductName

        'DockingPane
        If strPane <> "" Then
            On Error Resume Next
            dkpMain.LoadStateFromString strPane
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical, App.Title
            End If
            On Error GoTo 0
        End If
    Else
        Me.WindowState = vbMaximized
    End If
    
    Call VisibleToolButton                      '更新Button状态
    
    mblnEnter = True
End Sub

Private Sub InitCommandBars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbmTmp As CommandBarControlCustom
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        Set .Icons = mdlPublic.GetPubIcons
        .EnableCustomization False
        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
    picGroup_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picGroup.BackColor = picGroup_S.BackColor
    lblGroupDetail.BackColor = picGroup_S.BackColor
    
    '文件
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.文件, "文件(&F)", -1, False)
    With cbpTmp
        .id = enuMenus.文件
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.参数设置, "参数设置")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.导出报表, "导出报表"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.导入报表, "导入报表")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.导出全部, "导出全部")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.导入全部, "导入全部")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.退出, "退出"): cbcTmp.BeginGroup = True
    End With
    
    '编辑
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.编辑, "编辑(&E)", -1, False)
    With cbpTmp
        .id = enuMenus.编辑
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.新增报表类, "新增报表分类(&N)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.修改报表类, "修改报表分类(&M)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.删除报表类, "删除报表分类(&D)")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.新增报表组, "新增报表组(&W)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.修改报表组, "修改报表组(&M)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.删除报表组, "删除报表组(&D)")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.新增报表, "新增报表"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.修改报表, "修改报表")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.删除报表, "删除报表")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.设计报表, "设计报表"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.报表向导, "报表向导(&G)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.执行报表, "执行报表")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.报表启用, "启用(&S)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.报表停用, "停用(&T)")
    End With
    
    '工具
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.工具, "工具(&T)", -1, False)
    With cbpTmp
        .id = enuMenus.工具
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.报表发布, "报表发布(&B)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至导航台菜单, "至导航台菜单(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至模块内菜单, "至模块内菜单(&2)")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.取消发布, "取消发布(&U)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从导航台菜单, "从导航台菜单(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从模块内菜单, "从模块内菜单(&2)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.性能检查, "性能检查(&V)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.清除历史数据源, "清除历史数据源(&C)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.报表运行日志, "报表运行日志(&L)")
    End With
    
    '查看
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.查看, "查看(&V)", -1, False)
    With cbpTmp
        .id = enuMenus.查看
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.工具栏, "工具栏(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.标准按钮, "标准按钮(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.文本标签, "文本标签(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大图标, "大图标(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.状态栏, "状态栏(&S)")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.字体大小, "字体大小(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.小字体, "小字体(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大字体, "大字体(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.查找, "查找"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.刷新, "刷新")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.显示所有分类下级, "显示所有分类下级(&A)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.仅显示停用状态, "仅显示停用状态(&P)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.显示独立报表, "只显示独立报表(&R)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.显示子报表, "只显示子报表(&S)")
    End With
    
    '帮助
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.帮助, "帮助(&H)", -1, False)
    With cbpTmp
        .id = enuMenus.帮助
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB上的中联, "&WEB上的中联")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联主页, "中联主页(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联论坛, "中联论坛(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.发送反馈, "发送反馈(&K)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.关于, "关于(&A)"): cbcTmp.BeginGroup = True
    End With
    
    '定义工具栏
    Set cbrTmp = cbsMain.Add("工具栏", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap

        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.新增报表类, "新增类")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.修改报表类, "修改类")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.删除报表类, "删除类")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.新增报表组, "新增组")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.修改报表组, "修改组")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.删除报表组, "删除组")

        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.新增报表, "新增")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.修改报表, "修改")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.删除报表, "删除")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.设计报表, "设计"): cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.报表向导, "向导")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.执行报表, "执行")
        
        Set cbpTmp = .Controls.Add(xtpControlPopup, enuMenus.报表发布, "报表发布"): cbpTmp.BeginGroup = True
        cbpTmp.IconId = enuMenus.报表发布
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至导航台菜单, "至导航台菜单(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至模块内菜单, "至模块内菜单(&2)")
        Set cbpTmp = .Controls.Add(xtpControlPopup, enuMenus.取消发布, "取消发布")
        cbpTmp.IconId = enuMenus.取消发布
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从导航台菜单, "从导航台菜单(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从模块内菜单, "从模块内菜单(&2)")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.刷新, "刷新"): cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.退出, "退出"): cbcTmp.BeginGroup = True
        
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuMenus.选择系统标签, "系统"): cbcTmp.BeginGroup = True
        cbcTmp.Flags = xtpFlagRightAlign
        Set cbcTmp = .Controls.Add(xtpControlComboBox, enuMenus.选择系统控件, "")
        cbcTmp.Flags = xtpFlagRightAlign
        
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuMenus.查找报表标签, "查找")
        cbcTmp.Flags = xtpFlagRightAlign
        Set cbmTmp = .Controls.Add(xtpControlCustom, enuMenus.查找报表控件, "")
        cbmTmp.handle = picFind.hwnd: cbmTmp.Flags = xtpFlagRightAlign
    End With
    
    '菜单项的快键绑定
    With cbsMain.KeyBindings
        'alt
        .Add 16, vbKeyI, enuMenus.导入报表
        .Add 16, vbKeyO, enuMenus.导出报表
        .Add 16, vbKeyF1, enuMenus.导出全部
        .Add 16, vbKeyF2, enuMenus.导入全部
        .Add 16, vbKey1, enuMenus.TabRPT_1
        .Add 16, vbKey2, enuMenus.TabRPT_2
        'ctrl
        .Add 8, vbKeyX, enuMenus.退出
        .Add 8, vbKeyW, enuMenus.新增报表
        .Add 8, vbKeyM, enuMenus.修改报表
        .Add 8, vbKeyF, enuMenus.查找
        .Add 8, vbKeyE, enuMenus.设计报表
        'none
        .Add 0, vbKeyDelete, enuMenus.删除报表
        .Add 0, vbKeyF8, enuMenus.执行报表
        .Add 0, vbKeyF12, enuMenus.参数设置
        .Add 0, vbKeyF1, enuMenus.帮助主题
        .Add 0, vbKeyF3, enuMenus.查找报表控件
        .Add 0, vbKeyF5, enuMenus.刷新
    End With
    
    '有图标，有文本的按钮风格
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub InitDockPane()
    Dim panLeft As Pane, panRight As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeVisio
        
        Set panLeft = .CreatePane(1, 100, 0, DockLeftOf)
        With panLeft
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "报表分类"
            .MaxTrackSize.Width = 400
            .MinTrackSize.Width = 50
        End With
        
        Set panRight = .CreatePane(2, ScaleX(Me.Width, vbTwips, vbPixels) * 0.8, 0, DockRightOf)
        With panRight
            .Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .MinTrackSize.Width = 100
        End With
    End With
End Sub

Private Sub InitTabControl()
    With tbcRPT.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    
    With tbcRPT
        .InsertItem 0, "报表(&1)", picReport.hwnd, 0
        .InsertItem 1, "报表组(&2)", picGroup.hwnd, 0
    End With
End Sub

Private Sub InitOther()
    On Error Resume Next
    
    With txtFind
        .Top = 0
        .Left = 0
        .MaxLength = 20
    End With
    
    With picFind
        .Width = txtFind.Width
        .Height = txtFind.Height
    End With
    
End Sub

Private Sub InitReportControl()
    '初始化rptClass
    
    rptClass.ShowHeader = False
    rptClass.Icons = cbsMain.Icons
        
    If mobjClass Is Nothing Then
        Set mobjClass = New clsReportControlEx
    End If
    
    With mobjClass
        .AppTemplate atTree, rptClass, , "ID|上级ID|说明", "ID|上级ID|名称", Val("100-图标索引")
        .Init Me
    End With
End Sub

Private Sub InitVSF()
    Set mobjReport = New clsVSFlexGridEx
    Set mobjGroup = New clsVSFlexGridEx
    Set mobjSub = New clsVSFlexGridEx
    
    With mobjReport
        .AppTemplate EM_Display, vsfReport, MSTR_REPORT_COLS, "", True
        .Init True
    End With
    
    With mobjGroup
        .AppTemplate EM_Display, vsfGroup, MSTR_GROUP_COLS, "", True
        .Init True
    End With
    
    With mobjSub
        .AppTemplate EM_Display, vsfGroupDetail, MSTR_GROUPDETAIL_COLS, "", True
        .Init True
    End With
End Sub

Private Sub FillData(ByVal bytType As Byte, Optional ByVal blnColumn As Boolean = False)
'功能：为控件加载数据
'参数：
'  blnColumn：True-列头、列体都加载数据；False-只列体加载数据
    
    Dim objCBS_ComBox As CommandBarComboBox
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim lngClassID As Long, lngID As Long
    Dim intTab As Integer
    
    Set objCBS_ComBox = cbsMain.FindControl(, enuMenus.选择系统控件, , True)
    
    Select Case bytType
    Case Val("1-rptClass")
        strSQL = _
            "Select * " & vbCr & _
            "From (" & vbCr & _
            "    Select ID, Nvl(上级id, 0) 上级id, 名称, 说明" & vbCr & _
            "    From zlRPTClasses" & vbCr & _
            "    Union All " & vbCr & _
            "    Select 0, Null, '所有', null From Dual" & vbCr & _
            ")" & vbCr & _
            "Start With 上级ID Is Null Connect By Prior ID  = 上级ID"
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "获取报表类信息")
        If blnColumn Then
            Call mobjClass.Repaint(rsData, True)
        Else
            Call mobjClass.Repaint(rsData)
        End If
        If rptClass.Rows.count > 0 Then
            rptClass.FocusedRow = rptClass.Rows(0)
        End If
        rsData.Close
        
    Case Val("2-vsfReport")
        '系统号
        lngID = objCBS_ComBox.ItemData(objCBS_ComBox.ListIndex)
        '报表分类
        lngClassID = Val(rptClass.FocusedRow.Record.Item(mobjClass.GetColIndex("ID")).Value)
        
        strSQL = _
            "Select A.ID,A.编号,A.名称,A.说明,A.程序ID,A.修改时间,A.发布时间,A.系统,A.最后执行时间, a.分类ID, " & vbNewLine & _
            "    Decode(Nvl(A.票据, 0), 1, '票据', '报表') 种类, " & vbNewLine & _
            "    Decode(Nvl(A.系统, 0), 0, '自制', '系统') 类型, " & vbNewLine & _
            "    A.执行人员 最后执行人, zlSpellCode(A.名称) 简码, b.名称 报表分类, c.所属报表组, d.其他数据连接, " & vbNewLine & _
            "    A.是否停用 停用 " & vbNewLine & _
            "From zlReports A, zlRPTClasses B," & vbNewLine & _
            "   (Select c1.报表id, f_List2Str(Cast(Collect(c2.名称) as t_StrList)) 所属报表组" & vbNewLine & _
            "    From zlRPTSubs C1, ZlRPTGroups C2" & vbNewLine & _
            "    Where c1.组id = c2.ID And c2.系统 Is Null" & vbNewLine & _
            "    Group By c1.报表id" & vbNewLine & _
            "    ) C," & vbNewLine & _
            "   (Select d1.报表id, f_list2str(Cast(Collect(d2.名称) As t_Strlist)) 其他数据连接" & vbNewLine & _
            "    From zlRPTDatas D1, zlConnections D2" & vbNewLine & _
            "    Where d1.数据连接编号 = d2.编号" & vbNewLine & _
            "    Group By d1.报表id) D" & vbNewLine
        
        strSQL = strSQL & _
            "Where a.分类id = b.id(+) And a.id = c.报表id(+) And a.id = d.报表id(+)" & vbNewLine & _
            IIF(lngID <= 0 _
                    , "    And a.系统 Is Null " _
                    , "    And a.系统 = [1] ") & vbNewLine & _
            IIF(mbytReportGroup = 1 _
                    , "    And Exists(Select 1 From zlRPTSubs Where 报表id = a.Id) " _
                    , "    And Not Exists(Select 1 From zlRPTSubs Where 报表id = a.Id) ") & vbNewLine & _
            IIF(mblnDisplayChild _
                    , IIF(lngClassID > 0 _
                            , " And b.Id In (Select ID From ZLRPTClasses Start With Id = [2] Connect By Prior ID = 上级id) " _
                            , "") _
                    , IIF(lngClassID > 0 _
                            , " And b.Id = [2] " _
                            , " And Nvl(a.分类Id, 0) = 0 ")) & _
            IIF(mblnDisable, " And a.是否停用 = 1 ", " ") & vbNewLine & _
            "Order by A.编号"
        
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "获取报表信息" _
                    , lngID, lngClassID)
                    
        mobjReport.Recordset = rsData
        If blnColumn Then
            Call mobjReport.Repaint(RT_ColsAndRows)
        Else
            Call mobjReport.Repaint(RT_Rows)
        End If
        rsData.Close
        
        If mbytReportGroup = Val("0-显示独立报表") Then
            mobjReport.ColsHide = "性能检查结果|所属报表组"
        Else
            mobjReport.ColsHide = "性能检查结果"
        End If
        If mblnDisplayChild = False Or lngID > 0 Then
            mobjReport.ColsHide = mobjReport.ColsHide & "|报表分类"
        End If
        mobjReport.SetColsHide
        
    Case Val("3-vsfGroup")
        '系统号
        lngID = objCBS_ComBox.ItemData(objCBS_ComBox.ListIndex)
        '报表分类
        lngClassID = rptClass.FocusedRow.Record.Item(mobjClass.GetColIndex("ID")).Value
        '当前页面
        intTab = tbcRPT.Selected.Index
        
        strSQL = _
            "Select a.编号, a.名称 组名, a.说明, a.发布时间, a.ID, a.程序id, a.分类id, b.名称 报表分类, a.是否停用 停用 " & vbNewLine & _
            "From zlRPTGroups A, zlRPTClasses B " & vbNewLine & _
            "Where a.分类id = b.Id(+) " & _
            IIF(lngID <= 0, " And a.系统 Is Null", " And a.系统 = [1]") & vbNewLine & _
            IIF(mblnDisplayChild = True And intTab = 1 _
                    , IIF(lngClassID > 0 _
                            , "    And a.分类id in (Select Id From ZLRPTClasses Start With Id = [2] Connect By Prior ID = 上级id)" _
                            , "") _
                    , IIF(lngClassID > 0, "    And a.分类id = [2] ", " And Nvl(a.分类id, 0) = 0 ")) & vbNewLine & _
            IIF(mblnDisable, " And a.是否停用 = 1 ", " ") & vbNewLine & _
            "Order By a.编号 "
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "获取报表组信息" _
                   , lngID, lngClassID)
        mobjGroup.Recordset = rsData
        If blnColumn Then
            Call mobjGroup.Repaint(RT_ColsAndRows)
        Else
            Call mobjGroup.Repaint(RT_Rows)
        End If
        rsData.Close
        
        If mblnDisplayChild And lngID <= 0 Then
            mobjGroup.ColsHide = ""
        Else
            mobjGroup.ColsHide = "报表分类"
        End If
        mobjGroup.SetColsHide
        
    Case Val("4-vsfGroupDetail")
        '报表组ID
        If vsfGroup.Row >= 1 Then
            lngID = Val(vsfGroup.TextMatrix(vsfGroup.Row, vsfGroup.ColIndex("ID")))
        End If
        
        strSQL = _
            "Select a.Id, b.组Id, a.编号, a.名称, a.说明, a.程序id, a.修改时间, a.发布时间, a.系统, a.最后执行时间," & vbNewLine & _
            "    Decode(Nvl(A.票据, 0), 1, '票据', '报表') 种类, " & vbNewLine & _
            "    Decode(Nvl(A.系统, 0), 0, '自制', '系统') 类型, " & vbNewLine & _
            "    a.执行人员 最后执行人, zlSpellCode(a.名称) 简码, d.其他数据连接, a.是否停用 停用 " & vbNewLine & _
            "From ZLReports A, ZLRPTSubs B," & vbNewLine
'            "    (Select C1.报表id, f_List2str(Cast(Collect(C2.名称) As t_Strlist)) 所属报表组" & vbNewLine & _
'            "     From zlRPTSubs C1, zlRPTGroups C2" & vbNewLine & _
'            "     Where C1.组id = C2.Id And C2.系统 Is Null" & vbNewLine & _
'            "     Group By C1.报表id) C," & vbNewLine &
        strSQL = strSQL & _
            "    (Select D1.报表id, f_List2str(Cast(Collect(D2.名称) As t_Strlist)) 其他数据连接" & vbNewLine & _
            "     From zlRPTDatas D1, Zlconnections D2" & vbNewLine & _
            "     Where D1.数据连接编号 = D2.编号" & vbNewLine & _
            "     Group By D1.报表id) D" & vbNewLine & _
            "Where a.Id = b.报表id And a.Id = d.报表id(+)" & vbNewLine & _
            IIF(mblnDisable, " And a.是否停用 = 1 ", " ") & vbNewLine & _
            "    And b.组id = [1] " & vbNewLine & _
            "Order By a.编号 "
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "获取报表组的子表信息" _
                   , lngID)
        mobjSub.Recordset = rsData
        If blnColumn Then
            Call mobjSub.Repaint(RT_ColsAndRows)
        Else
            Call mobjSub.Repaint(RT_Rows)
        End If
        rsData.Close
        
    Case Val("5-cboSystem")
        If Not objCBS_ComBox Is Nothing Then
            objCBS_ComBox.Clear
            
            strSQL = _
                "Select 0 编号, '所有系统共享' 名称 From Dual Union All " & _
                "Select 编号, 名称||'【'||编号||'】' From zlSystems Order By 编号"
            Set rsData = mdlPublic.OpenSQLRecord(strSQL, "获取安装系统信息")
            With rsData
                Do While .EOF = False
                    objCBS_ComBox.AddItem rsData!名称
                    objCBS_ComBox.ItemData(objCBS_ComBox.ListCount) = rsData!编号
                    .MoveNext
                Loop
                .Close
            End With
        
            If objCBS_ComBox.ListCount > 0 Then
                objCBS_ComBox.ListIndex = 1
                glngSys = objCBS_ComBox.ItemData(1)
            End If
            objCBS_ComBox.Width = 160
        End If
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String, strPane As String
    
    mdlPublic.SaveWinState Me, App.ProductName
    
    strRegPath = mdlPublic.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.name)
    If glngSys <= 0 Then
        strPane = dkpMain.SaveStateToString
        Call SaveSetting("ZLSOFT", strRegPath, "布局", strPane)
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "显示所有分类下级", IIF(mblnDisplayChild, "1", "0"))
    Call SaveSetting("ZLSOFT", strRegPath, "显示报表类别", mbytReportGroup)
    Call SaveSetting("ZLSOFT", strRegPath, "字体大小", mbytFontSize)
    Call SaveSetting("ZLSOFT", strRegPath, "仅显示停用状态", IIF(mblnDisable, "1", "0"))
End Sub

Private Sub mobjGroup_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("发布时间")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("停用")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("发布时间") > intCol Then
        intCol = vsfVar.ColIndex("发布时间")
    End If
    
    If Col = intCol Then
        lngIcon = Val("4-报表组")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("发布时间")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("停用"))) = 1 Then
                lngIcon = Val("6-报表停用")
            Else
                lngIcon = Val("5-报表启用")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub mobjReport_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("发布时间")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("停用")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("发布时间") > intCol Then
        intCol = vsfVar.ColIndex("发布时间")
    End If

    If Col = intCol Then
        lngIcon = Val("1-报表")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("发布时间")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("停用"))) = 1 Then
                lngIcon = Val("3-报表停用")
            Else
                lngIcon = Val("2-报表启用")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub mobjSub_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("发布时间")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("停用")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("发布时间") > intCol Then
        intCol = vsfVar.ColIndex("发布时间")
    End If

    If Col >= intCol Then
        lngIcon = Val("1-报表")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("发布时间")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("停用"))) = 1 Then
                lngIcon = Val("3-报表停用")
            Else
                lngIcon = Val("2-报表启用")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("编号")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub picClass_Resize()
    On Error Resume Next
    
    With rptClass
        .Left = 0
        .Top = 0
        .Width = picClass.ScaleWidth
        .Height = picClass.ScaleHeight
    End With
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    
    With txtFind
        .Left = 0
        .Top = 0
        .Width = picFind.ScaleWidth
        If .Height > picFind.Height Then
            picFind.Height = .Height
        End If
    End With
End Sub

Private Sub picGroup_Resize()
    On Error Resume Next
    
    With picGroup_S
        .Left = 0
        .Width = picGroup.ScaleWidth
        If .Top > picGroup.ScaleHeight Then
            .Top = picGroup.ScaleHeight - 1500
        End If
    End With
    
    With vsfReport
        .Left = 0
        .Top = 0
        .Width = picReport.ScaleWidth
        .Height = picReport.ScaleHeight
    End With
    
    With vsfGroup
        .Left = 0
        .Top = 0
        .Width = picGroup.ScaleWidth
        .Height = picGroup_S.Top
    End With
    
    With lblGroupDetail
        .Top = picGroup_S.Top + picGroup_S.Height + 60
        .Left = 60
    End With
    
    With vsfGroupDetail
        .Left = 0
        .Top = lblGroupDetail.Top + lblGroupDetail.Height + 60
        .Width = picGroup.ScaleWidth
        .Height = picGroup.ScaleHeight - vsfGroup.Height - lblGroupDetail.Height - 60 * 2
    End With
End Sub

Private Sub picGroup_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '拖动时改变颜色
    If Button = vbLeftButton Then picGroup_S.BackColor = &H80000010
End Sub

Private Sub picGroup_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        With picGroup_S
            If .Top + Y < picGroup.ScaleHeight * 0.3 Then
                .Top = picGroup.ScaleHeight * 0.3
                Exit Sub
            End If
            If .Top + Y > picGroup.ScaleHeight * 0.8 Then
                .Top = picGroup.ScaleHeight * 0.8
                Exit Sub
            End If
            .Move .Left, .Top + Y
        End With
    End If
End Sub

Private Sub picGroup_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picGroup_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call picGroup_Resize
End Sub

Private Sub picReport_Resize()
    On Error Resume Next
    
    With vsfReport
        .Left = 0
        .Top = 0
        .Width = picReport.ScaleWidth
        .Height = picReport.ScaleHeight
    End With
End Sub

Private Sub picRPT_Resize()
    On Error Resume Next
    
    With tbcRPT
        .Left = 0
        .Top = 0
        .Width = picRPT.ScaleWidth
        .Height = picRPT.ScaleHeight
    End With
End Sub

Private Sub rptClass_DragDrop(Source As Control, X As Single, Y As Single)
    Dim lngSelRow As Long, l As Long, lngClassID As Long, lngID As Long, lngTemp As Long
    Dim strSQL As String
    Dim objInfo As XtremeReportControl.ReportHitTestInfo
    
    On Error GoTo hErr
    
    Set objInfo = rptClass.HitTest(Me.ScaleX(X, vbTwips, vbPixels) _
                                 , Me.ScaleY(Y, vbTwips, vbPixels))
    If objInfo Is Nothing Then Exit Sub
    If objInfo.Row Is Nothing Then Exit Sub
    
    Select Case UCase(Source.name)
    Case "VSFREPORT", "VSFGROUP"
        lngSelRow = 0
        For l = 1 To Source.Rows - 1
            If Source.SelectedRow(lngSelRow) = l Then
                '添加报表至分类
                
                '检查分类ID
                lngID = Val(Source.TextMatrix(l, Source.ColIndex("ID")))
                lngClassID = Val(objInfo.Row.Record(mobjClass.GetColIndex("ID")).Value)
                lngTemp = Val(Source.TextMatrix(l, Source.ColIndex("分类ID")))
                If lngTemp <> 0 And lngTemp = lngClassID Then
                    MsgBox "拒绝同一分类的拖动！", vbInformation, App.Title
                    Exit Sub
                End If
            
                '修改
                If UCase(Source.name) = "VSFREPORT" Then
                    strSQL = _
                        "Update zlReports " & vbCrLf & _
                        "Set 分类ID = " & IIF(lngClassID <= 0, "Null", lngClassID) & vbCrLf & _
                        "Where ID = " & lngID
                Else
                    strSQL = _
                        "Update zlRPTGroups " & vbCrLf & _
                        "Set 分类ID = " & IIF(lngClassID <= 0, "Null", lngClassID) & vbCrLf & _
                        "Where ID = " & lngID
                End If
                gcnOracle.Execute strSQL
                
                lngSelRow = lngSelRow + 1
            End If
        Next
        
        rptClass.Tag = ""
        Call RefreshEx
    End Select
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub rptClass_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mblnReportControlFocus = True
    Call Modify
End Sub

Private Sub rptClass_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If rptClass.Visible Then rptClass.SetFocus
    Call PopupMenuEx(Val("3-报表类菜单"))
End Sub

Private Sub RefreshEx(Optional ByVal bytType As Byte = 0)
'功能：
'参数：
'  bytType：0-刷新按钮触发；1-点击结点触发

    Dim lngID As Long
    
    If Me.Visible = False Then Exit Sub
    
    If bytType = 1 Then
        mblnReportControlFocus = glngSys <= 0
    Else
        mblnReportControlFocus = False
    End If
    
    lngID = mobjClass.GetColIndex("ID")
    If rptClass.Tag <> rptClass.FocusedRow.Record.Item(lngID).Value Then
        If tbcRPT.Selected.Index = Val("0-报表页面") Then
            Call FillData(Val("2-vsfReport"), False)
        Else
            Call FillData(Val("3-vsfGroup"), False)
            Call FillData(Val("4-vsfGroupDetail"), False)
        End If
    End If
    rptClass.Tag = rptClass.FocusedRow.Record.Item(lngID).Value
    Call UpdateStatusBar(rptClass)
    
    If mblnReportControlFocus Then
        Call VisibleToolButton(2)
    Else
        If tbcRPT.Selected.Index = 0 Then
            Call VisibleToolButton(0)
            vsfReport.SetFocus
        Else
            Call VisibleToolButton(1)
            vsfGroup.SetFocus
        End If
    End If
End Sub

Private Sub rptClass_SelectionChanged()
    Call RefreshEx(1)
End Sub

Private Sub SetControlFontSize(ByVal bytSize As Byte)
'功能：设置窗体控件的字体大小
'参数：
'  bytSize：0-小字体；1-大字体

    mbytFontSize = bytSize
    Call mdlPublic.SetPublicFontSize(Me, bytSize)
    picFind.Height = txtFind.Height
    
    If bytSize = 1 Then
        mobjReport.HeightColumn = 450
        mobjReport.HeightRow = 350
        mobjGroup.HeightColumn = 450
        mobjGroup.HeightRow = 350
        mobjSub.HeightColumn = 450
        mobjSub.HeightRow = 350
    Else
        mobjReport.HeightColumn = 350
        mobjReport.HeightRow = 250
        mobjGroup.HeightColumn = 350
        mobjGroup.HeightRow = 250
        mobjSub.HeightColumn = 350
        mobjSub.HeightRow = 250
    End If
    '重绘高度
    mobjReport.RepaintRowHeight
    mobjGroup.RepaintRowHeight
    mobjSub.RepaintRowHeight
End Sub

Private Sub tbcRPT_GotFocus()
    mblnReportControlFocus = False
End Sub

Private Sub tbcRPT_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible = False Then Exit Sub
    
    Set mobjFindPreviors = Nothing
    rptClass.Tag = ""
    Call rptClass_SelectionChanged
    
    mblnReportControlFocus = False
    If Item.Index = Val("0-报表页面") Then
        vsfReport.SetFocus
        Call VisibleToolButton
    Else
        vsfGroup.SetFocus
        Call VisibleToolButton(1)
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
    mblnReportControlFocus = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        '查找
        Call Find(txtFind.Text, True)
    End Select
End Sub

Private Sub vsfGroup_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngID As Long
    
    If Me.Visible = False Then Exit Sub
    
    If OldRow <> NewRow Then
        lngID = Val(vsfGroup.TextMatrix(NewRow, vsfGroup.ColIndex("ID")))
        Call FillData(Val("4-vsfGroupDetail"), False)
        Call UpdateStatusBar(vsfGroup)
    End If
End Sub

Private Sub CheckSQLPlanEx()
'功能：检查当前列表中的报表执行计划是否存在性能问题
    Dim i As Long
    Dim objReport As Report, objData As RPTData
    Dim strSQLCheck As String, strErr As String, strFields As String
    Dim strMsg As String, objPar As RPTPar, strSQL As String
    Dim lngCount As Long

    If MsgBox("当前目录一共" & vsfReport.Rows - 1 & "张报表，即将对这些报表(及参数)数据源中的SQL解析执行计划，" & _
              "然后检查执行计划是否存在以下情况：" & vbCrLf & _
              "    1.大表或中型表的全表扫描;" & vbCrLf & _
              "    2.大表或中型表的索引全扫描或跳跃式索引扫描;" & vbCrLf & _
              "    3.大表上引用基础表（非大表）的外键索引（例：病人医嘱记录_IX_诊疗项目ID）;" & vbCrLf & _
              "    其中大表是指zlBakTables ZlBigTables中定义的表;" & vbCrLf & _
              "    中型表是指收集统计信息后记录行数缺省在3千到1百万之间的表 (在设计界面的执行计划查看中可重新定义);" & vbCrLf & vbCrLf & _
              "此过程可能会花费几分钟的时间，你确定要继续吗？" _
        , vbQuestion + vbOKCancel + vbDefaultButton1, "性能检查") = vbCancel Then
        Exit Sub
    End If
    
    For i = 1 To vsfReport.Rows - 1
        Set objReport = ReadReport(Val(vsfReport.TextMatrix(i, vsfReport.ColIndex("ID"))), , True)
        strMsg = ""
        For Each objData In objReport.Datas
            With objData
                If .数据连接编号 > 0 Then GoTo makContinue
                
                '先检查数据源的SQL
                strSQLCheck = ""
                strFields = ""
                strSQL = RemoveNote(.SQL)
                strSQL = TrimChar(strSQL)
                strSQL = Replace(strSQL, "[系统]", glngSys)
                If GetParCount(strSQL) = 0 Then
                    strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .数据连接编号)
                Else
                    strFields = mdlPublic.CheckSQL(strSQL, strErr, ReplaceParSysNo(.Pars, glngSys) _
                        , strSQLCheck, , objReport.Datas, .数据连接编号)
                End If
                If strFields <> "" Then
                    If strSQLCheck <> "" Then
                        If mdlPublic.CheckSQLPlan(strSQLCheck, , .数据连接编号) = True Then
                            strMsg = strMsg & "," & .名称
                        End If
                    End If
                End If
                '再检查参数明细和分类SQL
                For Each objPar In .Pars
                    '排除已经检查过的
                    If objPar.分类SQL <> "" And InStr(strMsg, "(" & objPar.名称 & ")[分类]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.分类SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[系统]", glngSys)
                        Call mdlPublic.CheckParsRela(strSQL, objReport.Datas, objPar.名称, True)
                        strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .数据连接编号)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If mdlPublic.CheckSQLPlan(strSQLCheck, , .数据连接编号) = True Then
                                    strMsg = strMsg & "," & .名称 & "(" & objPar.名称 & ")[分类]"
                                End If
                            End If
                        End If
                    End If
                    
                    If objPar.明细SQL <> "" And InStr(strMsg, "(" & objPar.名称 & ")[明细]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.明细SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[系统]", glngSys)
                        Call mdlPublic.CheckParsRela(strSQL, objReport.Datas, objPar.名称, True)
                        strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , , objData.数据连接编号)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If mdlPublic.CheckSQLPlan(strSQLCheck, , objData.数据连接编号) = True Then
                                    strMsg = strMsg & "," & .名称 & "(" & objPar.名称 & ")[明细]"
                                End If
                            End If
                        End If
                    End If
                Next
            End With
makContinue:
        Next
        
        strMsg = Mid(strMsg, 2)
        If strMsg <> "" Then
            vsfReport.TextMatrix(i, vsfReport.ColIndex("性能检查结果")) = strMsg
            lngCount = lngCount + 1
        End If
        
        ShowFlash "正在检查报表数据源SQL存在的性能问题,请稍候 ...", i / (vsfReport.Rows - 1)
    Next
    
    vsfReport.ColHidden(vsfReport.ColIndex("性能检查结果")) = False
    ShowFlash
    
    If lngCount > 0 Then
        MsgBox "其中" & lngCount & "张报表(及参数)的数据源可能存在性能问题，详见""性能问题数据源""列的信息。" & vbCrLf & _
               "请在报表设计界面查看详细的执行计划，并进行SQL性能优化。" _
            , vbInformation, "性能检查结果"
    End If
End Sub

Private Function GetVsfControl(ByRef lngID As Long, ByRef blnIsGroup As Boolean _
    , Optional ByRef vsfActive As VSFlexGrid _
    , Optional ByRef strIDs As String) As Boolean
    
    Dim l As Long, lngSelRow As Long
    
    If Me.ActiveControl Is Nothing Then Exit Function
    If Me.ActiveControl.name = "" Then Exit Function
    
    lngID = 0
    blnIsGroup = False
    Set vsfActive = Nothing
    
    Select Case UCase(Me.ActiveControl.name)
    Case "VSFREPORT"
        Set vsfActive = vsfReport
    Case "VSFGROUP"
        Set vsfActive = vsfGroup: blnIsGroup = True
    Case "VSFGROUPDETAIL"
        Set vsfActive = vsfGroupDetail
    Case Else
        Exit Function
    End Select
    
    If Not vsfActive Is Nothing Then
        If vsfActive.Row > 0 Then
            lngID = Val(vsfActive.TextMatrix(vsfActive.Row, vsfActive.ColIndex("ID")))
        End If
    End If
    
    '多选
    lngSelRow = 0: strIDs = ""
    If vsfActive.SelectedRows > 0 Then
        For l = 1 To vsfActive.Rows
            If vsfActive.SelectedRow(lngSelRow) = l Then
                strIDs = strIDs & "," & vsfActive.TextMatrix(l, vsfActive.ColIndex("ID"))
                lngSelRow = lngSelRow + 1
            End If
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End If
    
    GetVsfControl = True
End Function

Private Sub Import(ByVal lngMenuID As Long)
'功能：导入报表

    Dim arrFile As Variant
    Dim i As Long, lngCurGroup As Long, lngGroupID As Long, lngID As Long
    Dim rsFiles As ADODB.Recordset, rsGroups As ADODB.Recordset
    Dim strRegPath As String, strPath As String, strFile As String, strSQL As String
    Dim strName As String, strCode As String
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim objFSO As New FileSystemObject, objFile As File, objFolder As Folder
    Dim arrTmp As Variant
    
    On Error GoTo hErr
    
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        '默认VSF控件
        If tbcRPT.Selected.Index = Val("0-报表页面") Then
            Set vsfTemp = vsfReport
            blnGroup = False
        Else
            Set vsfTemp = vsfGroupDetail
            blnGroup = True
        End If
    End If
    
    If UCase(vsfTemp.name) = "VSFGROUPDETAIL" Then
        '子报表
        Set vsfTemp = vsfGroup
        blnGroup = True
        lngGroupID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        '组报表
        lngID = 0
        lngGroupID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    Else
        '报表
        lngGroupID = 0
        lngID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    End If
    
    strRegPath = "公共模块\" & App.ProductName & "\Path"
    
    If lngMenuID = enuMenus.导入报表 Then
        '导入报表
        cdg.DialogTitle = "选择导入报表"
        cdg.Filter = "自定义报表文件|*.ZLR"
        cdg.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        cdg.InitDir = GetSetting("ZLSOFT", strRegPath, "Import", GetSetting("ZLSOFT", strRegPath, "Export", App.Path))
        cdg.FileName = ""
        cdg.MaxFileSize = 32767
        cdg.CancelError = True
        
        On Error Resume Next
        cdg.ShowOpen
        If Err.Number = 0 Then
            On Error GoTo hErr
            
            If cdg.FileTitle = "" Then
                '选择多个文件导入
                Call SaveSetting("ZLSOFT", strRegPath, "Import", Left(cdg.FileName, InStr(cdg.FileName, Chr(0)) - 1))
                arrFile = Split(cdg.FileName, Chr(0))
                For i = 1 To UBound(arrFile)
                    strFile = strFile & "|" & arrFile(0) & "\" & arrFile(i)
                Next
                strFile = Mid(strFile, 2)
            Else
                '选择单个文件导入
                Call SaveSetting("ZLSOFT", strRegPath, "Import", Left(cdg.FileName, InStrRev(cdg.FileName, "\")))
                strFile = cdg.FileName
            End If
            If strFile = "" Then Exit Sub
            
            arrFile = Split(strFile, "|")
            
            Set rsFiles = CopyNewRec(Nothing, , True _
                            , Array("FilePath", adVarChar, 1000, Empty _
                                  , "FileName", adVarChar, 200, Empty _
                                  , "组ID", adBigInt, Empty, Empty _
                                  , "同名ID", adBigInt, Empty, Empty _
                                  , "导入类型", adInteger, Empty, Empty _
                                  , "覆盖类型", adInteger, Empty, Empty _
                                  , "ErrType", adInteger, Empty, Empty _
                                  , "ImportResult", adInteger, Empty, Empty _
                                  , "ImportInfo", adVarChar, 200, Empty) _
                            )
            For i = LBound(arrFile) To UBound(arrFile)
                rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型", "覆盖类型" _
                                   , "ErrType", "ImportResult", "ImportInfo") _
                             , Array(arrFile(i), gobjFile.GetFileName(arrFile(i)), 0, 0, 0, 0, 0, 0, "")
            Next
            
            '导入
            Call ImportReportBeach(glngSys, lngGroupID, lngID, rsFiles, vsfTemp)
        End If
        Err.Clear: On Error GoTo hErr
    Else
        '导入全部
        strPath = BrowseForFolder(Me.hwnd, "选择需要导入报表所在目录", strPath)
        If strPath <> "" Then
            If MsgBox("是否导入“" & strPath & "”文件夹及子文件夹下的所有报表？", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Exit Sub
            End If
            
            lngCurGroup = lngGroupID
            
            'FilePath=报表全路径；FileName=报表文件名；组ID=报表要导入的报表组ID
            '同名ID=与将要导入的报表同名的报表的报表ID，固定报表通过编码匹配，非固定通过名称匹配
            '导入类型=0-不导入，1-新增导入,2-覆盖导入;覆盖类型=0-整体覆盖，1-仅数据源覆盖
            'ErrType=0-无错误,1-多个相同报表一起新增，2-多个相同报表一起覆盖，3-系统报表只能覆盖，但是无同名报表。
            '                            4-内容存在问题,5-版本存在问题,6-名称编号存在问题
            'ImportResult=-1-已经成功导入但是报表对象检查未通过，0-不导入,1-导入成功,2-导入失败
            'ImportInfo=报表成功导入后返回的报表信息
            Set rsFiles = CopyNewRec(Nothing, , True _
                                , Array("FilePath", adVarChar, 1000, Empty _
                                      , "FileName", adVarChar, 200, Empty _
                                      , "组ID", adBigInt, Empty, Empty _
                                      , "同名ID", adBigInt, Empty, Empty _
                                      , "导入类型", adInteger, Empty, Empty _
                                      , "覆盖类型", adInteger, Empty, Empty _
                                      , "ErrType", adInteger, Empty, Empty _
                                      , "ImportResult", adInteger, Empty, Empty _
                                      , "ImportInfo", adVarChar, 200, Empty) _
                                )
            
            With rsFiles
                '搜集导入到所有报表中的的报表,即当前文件夹下的报表
                For Each objFile In objFSO.GetFolder(strPath).Files
                    If UCase(objFile.name) Like "*.ZLR" Then
                        rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型" _
                                           , "覆盖类型", "ErrType", "ImportResult", "ImportInfo") _
                            , Array(objFile.Path, objFile.name, 0, 0, 0, 0, 0, 0, "")
                    End If
                Next
                '仅需要查找自定义报表的分组
                '固定报表由于编码唯一性，已经确定分组
                If glngSys = 0 Then
                    strSQL = "Select ID,编号,名称 From zlRPTGroups Where 系统 Is Null"
                    Set rsGroups = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption))
                End If
                
                '搜集当前文件下的子级文件夹
                For Each objFolder In objFSO.GetFolder(strPath).SubFolders
                    strFile = ""
                    For Each objFile In objFolder.Files
                        If UCase(objFile.name) Like "*.ZLR" Then
                            strFile = strFile & "|" & objFile.name
                        End If
                    Next
                    
                    If strFile <> "" Then
                        arrTmp = Split(Mid(strFile, 2), "|")
                        lngGroupID = 0
                        '仅自定报表需要查找分组，固定报表会有系统号编码确定分组
                        If glngSys = 0 Then
                            Call SplitNameCode(objFolder.name, strName, strCode)
                            rsGroups.Filter = "编号='" & strCode & "'"                          '编号唯一性
                            If rsGroups.EOF Then rsGroups.Filter = "名称='" & strName & "'"     '可能子分类没有编码
                            If Not rsGroups.EOF Then
                                lngGroupID = Nvl(rsGroups!id, 0)
                            Else
                                '生成常规性的报表组
                                '将编码名称规范化，并生成新的编码名称
                                lngGroupID = GetNextID("zlRPTGroups")
                                If TLen(strName) > 30 Then strName = ConvertSBC(MidB(strName, 1, 30))
                                If strCode <> "" Then
                                    If TLen(strCode) > 20 Then strCode = ConvertSBC(MidB(strCode, 1, 20))
                                    If CheckExist("zlRPTGroups", "编号", strCode) Then
                                        strCode = GetNextNO(True)
                                    End If
                                Else
                                    strCode = GetNextNO(True)
                                End If
                                strSQL = "Insert Into zlRPTGroups(ID,编号,名称,说明) Values(" & _
                                                lngGroupID & "," & _
                                                "'" & strCode & "'," & _
                                                "'" & strName & "',Null)"
                                On Error Resume Next
                                gcnOracle.Execute strSQL
                                If Err.Number <> 0 Then
                                    lngGroupID = 0  '生成报表组失败，则自动将该分组下的报表导入到所遇分类
                                Else '生成分组成功，加入到组信息缓存中
                                    rsGroups.AddNew Array("ID", "编号", "名称"), Array(lngGroupID, strCode, strName)
                                End If
                                On Error GoTo hErr
                            End If
                        End If
                        
                        For i = LBound(arrTmp) To UBound(arrTmp)
                            rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型" _
                                               , "覆盖类型", "ErrType", "ImportResult", "ImportInfo") _
                                    , Array(objFolder.Path & "\" & arrTmp(i), arrTmp(i), lngGroupID, 0, 0, 0, 0, 0, "")
                        
                        Next
                    End If
                Next
                
                .Filter = "": .Sort = "组ID"
                If .RecordCount = 0 Then
                    MsgBox "当前路径下未找到任何可导入的报表", vbInformation, App.Title
                    Exit Sub
                End If
                
                Call ImportReportBeach(glngSys, lngCurGroup, lngID, rsFiles, vsfTemp, True)
            End With
        End If
    End If
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    Call mdlPublic.ErrCenter
End Sub

Private Sub Export(ByVal lngMenuID As Long)
'功能：导出报表

    Dim strPath As String, strRegPath As String, strChoose As String
    Dim strCode As String, strName As String, strFile As String, strPathTmp As String
    Dim strSQL As String
    Dim blnGroup As Boolean, blnDo As Boolean
    Dim lngID As Long, lngCount As Long, l As Long, lngSelRow As Long, lngExp As Long
    Dim vsfTemp As VSFlexGrid
    Dim rsReports As ADODB.Recordset
    Dim objFile As New FileSystemObject
    
    On Error GoTo hErr
    
    strRegPath = mdlPublic.FormatString("公共模块\[1]\Path", App.ProductName)
    strPath = GetSetting("ZLSOFT" _
            , strRegPath _
            , "Export" _
            , GetSetting("ZLSOFT", strRegPath, "Import", App.Path))

    If lngMenuID = enuMenus.导出报表 Then
        '检查
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            MsgBox "请选中待导出的独立报表或子报表！", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "请选中待导出的独立报表或子报表！", vbInformation, App.Title
            Exit Sub
        End If
        If UCase(vsfTemp.name) = "VSFGROUP" Then
            Set vsfTemp = vsfGroupDetail
        End If
        
        If vsfTemp.SelectedRows > 1 Then
            strChoose = frmMsgBox.ShowMsgBox(App.Title _
                        , "请选择报表导出方式。" & _
                          "^导出当前清单中的所有报表时，文件自动按“[编号]名称”命名；" & _
                          "^如果导出目录中存在相同名称的报表文件，文件内容将被覆盖。" _
                        , "所有报表(&Y),!选中报表(&N),?取消(&C)" _
                        , Me)
        Else
            strChoose = frmMsgBox.ShowMsgBox(App.Title _
                        , "请选择报表导出方式。" & _
                          "^导出当前清单中的所有报表时，文件自动按“[编号]名称”命名；" & _
                          "^如果导出目录中存在相同名称的报表文件，文件内容将被覆盖。" _
                        , "所有报表(&Y),!当前报表(&N),?取消(&C)" _
                        , Me)
        End If
        If strChoose = "" Or strChoose = "取消" Then Exit Sub
        
        If strChoose = "当前报表" Then
            '缺省以报表名称作文件名
            strCode = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("编号"))
            strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("名称"))
            
            strFile = "[" & strCode & "]" & strName & ".ZLR"
            strFile = Replace(strFile, "\", "﹨")
            strFile = Replace(strFile, "/", "∕")
            strFile = Replace(strFile, ":", "：")
            strFile = Replace(strFile, "*", "﹡")
            strFile = Replace(strFile, "?", "？")
            strFile = Replace(strFile, """", "")
            strFile = Replace(strFile, "<", "〈")
            strFile = Replace(strFile, ">", "〉")
            strFile = Replace(strFile, "|", "∣")

            cdg.DialogTitle = "导出报表文件"
            cdg.Filter = "自定义报表文件|*.ZLR"
            cdg.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
            cdg.InitDir = strPath
            cdg.FileName = strFile
            cdg.CancelError = True

            On Error Resume Next
            Call cdg.ShowSave
            If Err.Number = 0 Then
                Call SaveSetting("ZLSOFT" _
                        , strRegPath _
                        , "Export" _
                        , Left(cdg.FileName, InStrRev(cdg.FileName, "\")))
                Call mdlPublic.ExportReport(CLng(lngID), cdg.FileName)
            End If
            On Error GoTo 0
        Else
            strFile = BrowseForFolder(Me.hwnd, "选择报表导出目录", strPath)
            If strFile <> "" Then
                strPath = strFile
                Call SaveSetting("ZLSOFT", strRegPath, "Export", strPath)
                
                lngCount = IIF(strChoose = "选中报表", vsfTemp.SelectedRows, vsfTemp.Rows - 1)
                If MsgBox("本次共导出 " & lngCount & " 张报表到 " & strPath & "，要继续吗？" _
                        , vbQuestion + vbYesNo + vbDefaultButton2 _
                        , App.Title) = vbNo Then
                    Exit Sub
                End If
                
                lngSelRow = 0
                For l = 1 To vsfTemp.Rows - 1
                    lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
                    strCode = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("编号"))
                    strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
                    strFile = "[" & strCode & "]" & strName & ".ZLR"
                    
                    blnDo = False
                    If strChoose = "选中报表" Then
                        If vsfTemp.SelectedRow(lngSelRow) = l Then
                            blnDo = True
                            lngSelRow = lngSelRow + 1
                        End If
                    Else
                        blnDo = True
                    End If
                    
                    If blnDo And lngID > 0 Then
                        Call ShowFlash("正在导出:" & strFile & ".ZLR", l / lngCount, Me, True)
                        If mdlPublic.ExportReport(lngID, strPath & "\" & strFile) = False Then
                            Call ShowFlash
                            If MsgBox("导出报表时出现错误，要继续导出下一张报表吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                        End If
                    End If
                Next
                Call ShowFlash
            End If
        End If
    Else
        '当前系统全部导出
        strPath = BrowseForFolder(Me.hwnd, "选择报表导出目录", strPath)
        If strPath <> "" Then
            Call SaveSetting("ZLSOFT", strRegPath, "Export", strPath)
            strSQL = _
                "Select A.Id, A.编号, A.名称, C.Id 组id, C.编号 组编号, C.名称 组名 " & vbNewLine & _
                "From zlReports A, zlRPTSubs B, zlRPTGroups C " & vbNewLine & _
                "Where A.Id = B.报表id(+) And B.组id = C.Id(+) And " & vbNewLine & _
                IIF(glngSys = 0, " A.系统 Is Null ", " A.系统=[1] ") & vbNewLine & _
                "Order By C.编号,A.编号 "
            Set rsReports = OpenSQLRecord(strSQL, Me.Caption, glngSys)
            lngCount = rsReports.RecordCount
            
            If lngCount = 0 Then
                MsgBox "目前无报表可导出！", vbInformation, App.Title
                Exit Sub
            End If
            
            If MsgBox("本次共导出 " & lngCount & " 张报表到 " & strPath & "，要继续吗？" _
                , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            
            lngExp = 0
            rsReports.MoveFirst
            For l = 1 To rsReports.RecordCount
                lngExp = lngExp + 1
                Call ShowFlash("正在导出：" & rsReports!名称 & ".ZLR", lngExp / lngCount, Me, True)
                
                If Nvl(rsReports!组ID, 0) = 0 Then
                    strPathTmp = strPath
                Else
                    strPathTmp = strPath & "\[" & rsReports!组编号 & "]" & rsReports!组名
                    If Not objFile.FolderExists(strPathTmp) Then
                        Call objFile.CreateFolder(strPathTmp)
                    End If
                End If
                strFile = "[" & rsReports!编号 & "]" & rsReports!名称 & ".ZLR"
                
                If Not ExportReport(rsReports!id, strPathTmp & "\" & strFile) Then
                    Call ShowFlash
                    If MsgBox("导出报表时出现错误，要继续导出下一张报表吗？" _
                        , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                End If
                rsReports.MoveNext
            Next
            rsReports.Close

            Call ShowFlash
        End If
    End If
    
    Exit Sub
    
hErr:
    Call mdlPublic.ErrCenter
End Sub

Private Function ImportReportBeach(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal lngCurPRTID As Long _
    , ByVal rsFiles As ADODB.Recordset, vsfImp As VSFlexGrid _
    , Optional ByVal blnALLImp As Boolean) As Boolean
'功能：批量导入报表，可以导入1个至多个
'参数：
'    lngSys = 当前选择的系统
'    lngGroup = 当前选择的记录集
'    rsFiles = 需要导入的报表文件
'    lngCurPRTID = 当前选择的报表ID
'    blnALLImp=是否是全部倒入，非固定报表全部导入时，也需要读取所有报表
'    vsfImp=
'返回：是否成功导入

    Dim rsReports As New ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long, lngGroupID As Long, lngClassID As Long
    Dim blnSingle  As Boolean, strFileName As String
    Dim strCurRPT As String, strSameRPT As String
    
    On Error GoTo hErr
    
    '固定报表，以及非显示独立项下的非固定报表的所有报表分组时，需要读取所有报表
    If lngSys <> 0 Or mbytReportGroup <> 0 And lngGroup = 0 And lngSys = 0 Or blnALLImp Then
        '查询所有的报表
        strSQL = _
            "Select A.ID,A.编号,A.名称,A.说明,Nvl(B.组id,0) 组id " & vbNewLine & _
            "From zlReports A, zlRPTSubs B " & vbNewLine & _
            "Where " & IIF(lngSys = 0, " A.系统 Is Null ", " A.系统=[1] ") & vbNewLine & _
            "    And A.ID=B.报表ID(+)" & vbNewLine & _
            "Order by A.编号"
    Else
        '非固定报表读取
        If lngGroup <> 0 Then
            strSQL = _
                "Select Id,编号,名称,[2] 组id " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where Id In (Select 报表id From Zlrptsubs Where 组id = [2]) " & vbNewLine & _
                "Order By 编号"
        Else
            strSQL = _
                "Select ID,编号,名称,0 组id " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where " & IIF(lngSys = 0, " 系统 Is Null ", " 系统=[1] ") & vbNewLine & _
                "    And ID Not In (Select 报表ID From zlRPTSubs) " & vbNewLine & _
                "Order by 编号"
        End If
    End If
    Set rsReports = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption, lngSys, lngGroup))
    
    If lngCurPRTID <> 0 Then
        rsReports.Filter = "ID=" & lngCurPRTID
        If rsReports.EOF Then
            MsgBox "当前选中报表已经不存在，请刷新后继续！", vbInformation, App.Title
            Exit Function
        Else
            strCurRPT = "[" & rsReports!编号 & "]" & rsReports!名称
        End If
    End If
    
    '获取当前报表分类ID
    lngClassID = 0
    If Not rptClass.FocusedRow Is Nothing Then
        lngClassID = Val(rptClass.FocusedRow.Record(mobjClass.GetColIndex("ID")).Value)
    End If
    If lngClassID < 0 Then lngClassID = 0
    
    With rsFiles
        '不同子文件导入到同一分组时的同名文件检查
        '具体情况如下：[GROUP_001]住院工作报表ASD，住院工作报表，[GROUP_001]住院工作报表
        '                        这三个子文件的报表可以导入到[GROUP_001]住院工作报分组中
        '不同文件名的报表，可能是同一个报表。
        '检查导入文件，以及确定导入类型，倒入分组以及覆盖的报表ID等
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = rsFiles.RecordCount = 1 '是否单个报表导入
        If blnSingle Then strFileName = rsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '获取文件信息
            If UBound(arrTmp) <> 2 Then
                intErrType = 4 '文件检查
            ElseIf Val(arrTmp(2)) <> 9 Then
                intErrType = 5  '版本检查
                If blnSingle Then strFileName = strFileName & "(原始名称：[" & arrTmp(0) & "]" & arrTmp(1) & ")"
            Else
                If blnSingle Then strFileName = strFileName & "(原始名称：[" & arrTmp(0) & "]" & arrTmp(1) & ")"
                If lngSys = 0 Then '非系统报表要求分组的报表中不能存在相同报表
                    '非固定报表全部导入已经确定报表要导入的分组
                    rsReports.Filter = "名称='" & arrTmp(1) & "' And 编号='" & arrTmp(0) & "' And ID>0 " & IIF(blnALLImp, " And 组ID=" & !组ID, "")
                    If rsReports.EOF Then rsReports.Filter = "名称='" & arrTmp(1) & "'  And ID>0 " & IIF(blnALLImp, " And 组ID=" & !组ID, "")
                Else '系统报表通过编号直接查找
                    rsReports.Filter = "名称='" & arrTmp(1) & "' And 编号='" & arrTmp(0) & "' And ID>0"
                    If rsReports.EOF Then rsReports.Filter = "编号='" & arrTmp(0) & "' And ID>0"
                End If
                '确定报表导入的分组，如果存在的同名的，优先查找没有分组的报表
                rsReports.Sort = "ID Desc,组ID"
                If Not rsReports.EOF Then
                    lngRPTID = rsReports!id: lngImpGroup = rsReports!组ID
                    If lngRPTID = 0 Then
                        intErrType = 1 '该报表已经被标记新增
                    ElseIf lngRPTID < 0 Then
                        intErrType = 2 '该报表已经被标记覆盖
                    Else
                        intImpType = 2
                        '编号名称不匹配
                        If (CStr(arrTmp(0)) <> rsReports!编号 & "" Or CStr(arrTmp(1)) <> rsReports!名称) Then intErrType = 6
                        rsReports.Update "Id", lngRPTID * -1 '标记已经覆盖
                        If blnSingle Then strSameRPT = "[" & rsReports!编号 & "]" & rsReports!名称
                    End If
                Else
                    If lngSys <> 0 Then
                        intErrType = 3  '系统固定报表必须覆盖同名报表
                    Else
                        intImpType = 1  '非系统报表没有同名，则新增报表
                        If lngSys = 0 And blnALLImp Then
                            lngImpGroup = !组ID         '导入取原来的分组
                        Else
                            lngImpGroup = lngGroup      '导入到界面指定的分组
                        End If
                        '该报表是新增报表，则加入缓存，防止多次增加
                        If !组ID = 0 Then
                            rsReports.AddNew Array("Id", "编号", "名称", "组iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), lngImpGroup)
                        Else
                            rsReports.AddNew Array("Id", "编号", "名称", "组iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), !组ID)
                        End If
                    End If
                End If
            End If
            If lngSys = 0 And blnALLImp Then lngImpGroup = !组ID '非固定报表导入取原来的分组
            .Update Array("组ID", "同名ID", "导入类型", "ErrType") _
                  , Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            .MoveNext
        Loop
        
        If blnSingle Then
            '单个报表文件
            .Filter = ""
            Select Case !ErrType
            Case 4
                MsgBox "报表“" & strFileName & "”由于内容存在问题而无法导入！", vbInformation, App.Title
                Exit Function
            Case 5
                MsgBox "报表“" & strFileName & "”由于版本不对而无法导入！", vbInformation, App.Title
                Exit Function
            Case 3
                If lngCurPRTID <> 0 Then '更新状态，默认覆盖当前的报表
                    .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 6)
                Else
                    MsgBox "请选择你要覆盖的报表后继续！", vbInformation, App.Title
                    Exit Function
                End If
            End Select
            
            Select Case !导入类型
            Case 1
                strReturn = frmMsgBox.ShowMsgBox(App.Title, "是否新增导入报表""" & strFileName & """！", "新增导入(&N),!?取消(&C)", Me)
            Case 2
                If lngSys = 0 And lngGroup = 0 Then '所有系统共享的为分组的报表,此时可以存在新增报表选项
                    If lngCurPRTID = !同名ID Then
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """编号或名称" & vbNewLine & "与要覆盖的当前选择报表""" & strCurRPT & """不相符，请选择确认！", _
                                    "报表""" & strFileName & """编号和名称" & vbNewLine & "与当前选择报表""" & strCurRPT & """都相符，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),新增导入(&N),!?取消(&C)", Me)
                    ElseIf lngCurPRTID = 0 Then
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """存在部分匹配的报表""" & strSameRPT & """," & vbNewLine & "但是二者编号或名称不相符，请选择确认！", _
                                    "报表""" & strFileName & """存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖匹配(&O),新增导入(&N),!?取消(&C)", Me)
                    Else
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """的编号或名称" & vbNewLine & "与部分匹配报表""" & strSameRPT & """" & vbNewLine & "以及当前选择报表""" & strCurRPT & """均不相符，请选择确认！", _
                                    "报表""" & strFileName & """编号或名称" & vbNewLine & "与当前选择报""" & strCurRPT & """不相符，" & vbNewLine & "但是存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),覆盖匹配(&O),新增导入(&N),!?取消(&C)", Me)
                    End If
                Else
                   If lngCurPRTID = !同名ID Then
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """编号或名称" & vbNewLine & "与要覆盖的当前选择报表""" & strCurRPT & """不相符，请选择确认！", _
                                    "报表""" & strFileName & """编号和名称" & vbNewLine & "与当前选择报表""" & strCurRPT & """都相符，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),!?取消(&C)", Me)
                    ElseIf lngCurPRTID = 0 Then
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """存在部分匹配的报表""" & strSameRPT & """," & vbNewLine & "但是二者编号或名称不相符，请选择确认！", _
                                    "报表""" & strFileName & """存在" & vbNewLine & "编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖匹配(&O),!?取消(&C)", Me)
                    Else
                        strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """的编号或名称" & vbNewLine & "与部分匹配报表""" & strSameRPT & """" & vbNewLine & " 以及当前选择报表""" & strCurRPT & """均不相符，请选择确认！", _
                                    "报表""" & strFileName & """编号或名称" & vbNewLine & "与当前选择报""" & strCurRPT & """不相符，" & vbNewLine & "但是存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),覆盖匹配(&O),!?取消(&C)", Me)
                    End If
                End If
            End Select
            
            If strReturn = "" Then
                Exit Function
            ElseIf strReturn = "新增导入" Then
                .Update Array("组ID", "同名ID", "导入类型", "ErrType") _
                      , Array(lngGroup, 0, 1, 0)
            Else
                If strReturn = "覆盖当前" Then
                    .Update Array("组ID", "同名ID", "导入类型", "ErrType") _
                          , Array(lngGroup, lngCurPRTID, 2, 0)
                Else
                    .Update Array("导入类型", "ErrType") _
                          , Array(2, 0)
                End If
                strMsg = frmMsgBox.ShowMsgBox(App.Title _
                            , "是否只导入数据源？" & vbNewLine & _
                              "只导入数据源可以保持现有报表的格式，更详细的情况请咨询系统管理员！" _
                            , "仅数据源(&D),!?整体导入(&F)" _
                            , Me)
                If strMsg = "仅数据源" Then
                    .Update "覆盖类型", 1
                End If
            End If
        Else
            '多个报表文件
            If MsgBox("当前导入多张报表，系统将自动寻找编码或名称匹配的报表进行覆盖。请确认是否继续！", vbInformation + vbYesNo, App.Title) = vbNo Then
                Exit Function
            End If
            
            '不能导入的类型信息生成
            .Filter = "ErrType>0 And ErrType<6": .Sort = "ErrType": intImpType = 0
            Do While Not .EOF
                If intImpType <> Val(!ErrType & "") Then
                    If intImpType <> 0 Then
                        strMsg = strMsg & vbNewLine
                    End If
                    intImpType = Val(!ErrType & ""): lngCount = 0
                    Select Case intImpType
                    Case 1
                        strMsg = strMsg & vbNewLine & "以下报表由于存在相同内容的报表而无法新增导入："
                    Case 2
                        strMsg = strMsg & vbNewLine & "以下报表由于存在相同内容的报表而无法覆盖导入："
                    Case 3
                        strMsg = strMsg & vbNewLine & "以下报表由于没有可以覆盖的报表而无法导入："
                    Case 4
                        strMsg = strMsg & vbNewLine & "以下报表由于内容存在问题而无法导入："
                    Case 5
                        strMsg = strMsg & vbNewLine & "以下报表由于版本不对而无法导入："
                    End Select
                End If
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            
            .Filter = "导入类型<>0"
            If .RecordCount = 0 Then '没有导入报表
                MsgBox "没有可以导入的报表！" & Mid(strMsg, 1, Len(strMsg) - 2) & "。", vbInformation, App.Title
                Exit Function
            End If
            
            '文件名以及编码不匹配提示
            .Filter = "ErrType=6"
            If Not .EOF Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "编号或名称与覆盖的报表不相符，请选择确认："
                Do While Not .EOF
                    If lngCount < 5 Then
                        strMsg = strMsg & vbNewLine & CStr(lngCount + 1) & "." & !FileName
                    ElseIf lngCount = 5 Then
                        strMsg = strMsg & vbNewLine & "..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
                .Filter = "ErrType=0" '不存在可以直接导入的，则提示是否继续
                If .RecordCount = 0 Then
                    strReturn = frmMsgBox.ShowMsgBox(App.Title _
                                    , Mid(strMsg, 1, Len(strMsg) - Len(vbNewLine)) _
                                    , "整体覆盖(&A),数据源覆盖(&D),!?取消(&C)" _
                                    , Me)
                    If strReturn = "" Then Exit Function
                End If
            End If
            
            .Filter = "导入类型=2 And ErrType=0": .Sort = "ErrType" '存在覆盖报表，则提示选择整体覆盖，还是数据源覆盖
            If Not .EOF Then
                strMsg = strMsg & vbNewLine & "以下报表将会覆盖原有报表，请选择确认："
                strOption = "整体覆盖(&A),数据源覆盖(&D),!?取消(&C)"
                lngCount = 0
            End If

            Do While Not .EOF
                If lngCount < 5 Then
                    strMsg = strMsg & vbNewLine & CStr(lngCount + 1) & "." & !FileName
                ElseIf lngCount = 5 Then
                    strMsg = strMsg & vbNewLine & "..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            
            .Filter = "导入类型=1" '新增导入
            If .RecordCount <> 0 And strReturn = "" And strOption = "" Then '所有报表新增
                strReturn = frmMsgBox.ShowMsgBox(App.Title _
                                , Mid(strMsg, Len(vbNewLine) + 1) & "请确认是否导入？" _
                                , "导入(&N),!?取消(&C)" _
                                , Me)
                If strReturn = "" Then Exit Function
            End If
            
            '选择覆盖类型
            If strReturn = "" And strOption <> "" Then '存在覆盖,且不存在ErrType=6的类型
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1, Len(strMsg) - Len(vbNewLine) * 2), strOption, Me)
                If strReturn = "" Then Exit Function
            End If
        End If
        
        If strReturn = "数据源覆盖" Then
            .Filter = "导入类型=2"
            Do While Not .EOF
                .Update "覆盖类型", 1
                .MoveNext
            Loop
        End If
        
        Screen.MousePointer = vbHourglass
        
        .Filter = "导入类型<>0"
        .Sort = "导入类型"
        lngCount = .RecordCount
        Do While Not .EOF
            If Not blnSingle Then
                Call ShowFlash("正在导入:" & !FileName, i / lngCount, Me, True)
            Else
                Call ShowFlash("正在导入:" & !FileName, , Me, True)
            End If
            Me.Refresh
            DoEvents
            
            '正式导入文件
            strInfo = ImportReport(!FilePath & "", Val(!同名ID & ""), Val(!覆盖类型 & "") = 1 _
                                    , Val(!组ID & ""), lngClassID)
            .Update Array("ImportResult", "ImportInfo"), Array(IIF(strInfo <> "", 1, 2), strInfo)
            
            '报表对象权限检查
            If strInfo <> "" Then
                arrTmp = Split(strInfo, "|")
                If Not mdlPublic.CheckReportPriv(CLng(arrTmp(0))) Then
                    .Update Array("ImportResult", "同名ID"), Array(-1, Val(arrTmp(0)))
                Else
                    .Update "同名ID", Val(arrTmp(0))
                End If
            End If
            
            i = i + 1
            .MoveNext
        Loop
        Call ShowFlash
        
        '导入情况提示
        strMsg = ""
        If Not blnSingle Then
            .Filter = "ImportResult=1 Or ImportResult=-1"
            If .RecordCount = 0 Then
                strMsg = "所有报表均为导入成功。"
            Else
                strMsg = "成功导入了 " & .RecordCount & " 张报表。"
            End If
            
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表的报表文件内容可能已被非法修改："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=-1 And 导入类型=1"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "你没有权限查询以下导入报表中全部或部份数据对象："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=-1 And 导入类型=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "你没有权限查询以下导入报表中全部或部份数据对象,在使用该报表之前,请手工对报表内容进行调整："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=1 And 导入类型=2"
            If .RecordCount <> 0 And lngSys <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表成功覆盖相应报表,你可能需要重新授权才能正常使用这些报表："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表导入失败："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
        Else
            .Filter = ""
            Select Case !ImportResult
            Case -1
                strMsg = "你没有权限查询报表“" & strFileName & "”中全部或部份数据对象" & _
                         IIF(!导入类型 = 2, "。你可能需要手工对报表内容进行调整并重新授权才能正常使用该报表！", "！")
            Case 1
                strMsg = "报表“" & strFileName & "”导入成功" & _
                         IIF(!导入类型 = 2, "。你可能需要重新授权才能正常使用该报表！", "！")
            Case 2
                strMsg = "报表“" & strFileName & "”" & _
                         IIF(!导入类型 = 2, "覆盖失败。报表文件内容可能已被非法修改！", "新增导入失败！")
            End Select
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strMsg, vbInformation, App.Title
    End With
    
    Exit Function
    
hErr:
    Call ShowFlash
End Function

Private Sub Delete(ByVal lngMenuID As Long)
    Dim rsCheck As New ADODB.Recordset, rsGetGroups As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, lngID As Long, lngSelRow As Long, lngCount As Long
    Dim strSQL As String, strIDs As String, strRec As String
    Dim vsfTemp As VSFlexGrid
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim arrItem As Variant
    Dim colSQL As New Collection
    
    If mblnReportControlFocus = False Then
        If GetVsfControl(lngID, blnGroup, vsfTemp, strIDs) = False Then
            MsgBox "请选中待导出的独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "请选中待导出的独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    Select Case lngMenuID
    Case enuMenus.删除报表类
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "请选中一个报表类！", vbInformation, App.Title
            Exit Sub
        End If
        
        strRec = rptClass.FocusedRow.Record(mobjClass.GetColIndex("名称")).Value
        
        If MsgBox(mdlPublic.FormatString("你确定删除【[1]】报表分类？" & vbCrLf & _
                                         "注意：独立报表、报表组将无分类，但报表、报表组仍然存在。" _
                                , strRec) _
            , vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        
        '删除
        With rptClass
            lngID = Val(.FocusedRow.Record(mobjClass.GetColIndex("ID")).Value)
            
            On Error GoTo hErr
            
            strSQL = _
                "Update zlReports Set 分类id = Null " & vbNewLine & _
                "Where 分类id In (Select ID From zlRPTClasses Start With ID = " & lngID & " Connect By Prior ID = 上级id)"
            Call AddArray(colSQL, strSQL)
            
            strSQL = _
                "Update zlRPTGroups Set 分类id = Null " & vbNewLine & _
                "Where 分类id In (Select ID From zlRPTClasses Start With ID = " & lngID & " Connect By Prior ID = 上级id)"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete zlRPTClasses Where ID = " & lngID
            Call AddArray(colSQL, strSQL)
            
            '执行DML
            gcnOracle.BeginTrans: blnTrans = True
            For lngRow = 1 To colSQL.count
                gcnOracle.Execute colSQL(lngRow)
            Next
            gcnOracle.CommitTrans: blnTrans = False
        End With
        
        '刷新
        Call FillData(Val("1-报表类"))
        
    Case enuMenus.删除报表组
        If blnGroup = False Then
            MsgBox "请选择报表组！", vbInformation, App.Title
            Exit Sub
        End If
        
        '检查是否已发布
        strRec = "": lngSelRow = 0: lngCount = 0
        For lngRow = 1 To vsfTemp.Rows - 1
            If lngCount <= 4 Then
                If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                    If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("发布时间")) <> "" Then
                        strRec = strRec & vbCrLf & CStr(lngCount + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("组名"))
                        lngCount = lngCount + 1
                    End If
                    lngSelRow = lngSelRow + 1
                End If
            Else
                strRec = strRec & vbCrLf & "..."
                Exit For
            End If
        Next
        If strRec <> "" Then
            MsgBox "下列报表已经发布，请先取消发布后再删除！" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        strRec = GetSelectedReport(vsfTemp, "组名")
        If MsgBox("你确定要删除下列报表组吗？" & strRec _
            , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        On Error GoTo hErr: blnTrans = True
        gcnOracle.BeginTrans
        gcnOracle.Execute "Delete zlRPTSubs Where 组ID=" & lngID
        gcnOracle.Execute "Delete zlRPTGroups Where ID=" & lngID
        gcnOracle.CommitTrans: blnTrans = False
        
    Case enuMenus.删除报表
        '检查是否为报表组成员
        lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.名称 " & vbNewLine & _
            "From zlReports A, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = d.Column_Value " & vbNewLine & _
            "    And Exists(Select 1 From zlRPTSubs Where 报表id = a.Id) " & vbNewLine & _
            "Order By a.名称 "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!名称
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If strRec <> "" Then
            MsgBox "请先把下列报表从报表组中移除后再删除！" & strRec _
                , vbInformation, App.Title
            Exit Sub
        End If
        
        '检查是否已发布
        strRec = "": lngSelRow = 0: lngCount = 0
        For lngRow = 1 To vsfTemp.Rows - 1
            If lngCount <= 4 Then
                If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                    If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("发布时间")) <> "" Then
                        strRec = strRec & vbCrLf & CStr(lngCount + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("名称"))
                        lngCount = lngCount + 1
                    End If
                    lngSelRow = lngSelRow + 1
                End If
            Else
                If lngCount > 4 Then
                    strRec = strRec & vbCrLf & "..."
                End If
                Exit For
            End If
        Next
        If strRec <> "" Then
            MsgBox "下列报表已经发布，请先取消发布后再删除！" & strRec, vbInformation, App.Title
            Exit Sub
        End If

        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.名称 " & vbNewLine & _
            "From zlReports A, zlRPTPuts B, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = b.报表Id And a.Id = d.Column_Value " & vbNewLine & _
            "Order By a.名称 "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!名称
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If strRec <> "" Then
            MsgBox "下列报表已经发布，请先取消发布后再删除！" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        '检查是否与其他报表有关联
        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(A, 10)*/ a.Id 报表ID, a.名称 " & vbNewLine & _
            "From zlReports A, Zlrptrelation B, Table(Cast(f_Str2List([1]) as t_StrList)) C " & vbNewLine & _
            "Where a.id = b.报表id and a.id = c.Column_Value " & vbNewLine & _
            "Union all " & vbNewLine & _
            "Select /*+ cardinality(A, 10)*/ a.Id 报表ID, a.名称 " & vbNewLine & _
            "From zlReports A, Zlrptrelation B, Table(Cast(f_Str2List([1]) as t_StrList)) C " & vbNewLine & _
            "Where a.id = b.关联报表id and a.id = c.Column_Value "
        strSQL = "Select Distinct 报表ID, 名称 From (" & strSQL & ")"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!名称
                strRec = strRec & GetRelationList(rsCheck!报表ID)
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        If strRec <> "" Then
            MsgBox "下列报表存在关联，请先取消关联后再删除！" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        '获取待删除报表名称
        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.名称 " & vbNewLine & _
            "From zlReports A, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = d.Column_Value " & vbNewLine & _
            "Order By a.名称 "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!名称
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If MsgBox("确定要删除下列报表吗？" & strRec _
                , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Exit Sub
        End If
        
        On Error GoTo hErr
        blnTrans = True
        gcnOracle.BeginTrans
        
        arrItem = Split(strIDs, ",")
        For lngRow = LBound(arrItem) To UBound(arrItem)
            lngID = arrItem(lngRow)
            If lngID <> 0 Then
                gcnOracle.Execute "Delete From zlReports Where ID=" & CStr(lngID)
            End If
        Next
        
        gcnOracle.CommitTrans
        blnTrans = False
        On Error GoTo 0
    End Select
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Call ErrCenter
End Sub

Private Sub SplitNameCode(ByVal strInput As String, ByRef strName As String, ByRef strCode As String)
'功能:分割编码名称
'参数：strInput=输入的字符串，如果格式为[编码]名称,则自动分割，否则默认为只获取到名称
'返回：strName=名称
'           strCode=编码
    Dim arrTmp As Variant
    Dim strTmp As Variant
    If InStr(strInput, "\") > 0 Then
        strTmp = StrReverse(strInput)
        strInput = StrReverse(Mid(strTmp, 1, InStr(strTmp, "\") - 1))
    End If
    
    If strInput Like "[[]?*[]]?*" Then '符合规范的文件名
        arrTmp = Split(strInput, "]")
        strName = arrTmp(1)
        strCode = Mid(arrTmp(0), 2)
    Else
        strName = strInput
        strCode = ""
    End If
End Sub

Private Sub Modify()
    Dim lngID As Long, lngProgID As Long, lngGroupID As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim strCode As String, strName As String, strDescription As String
    Dim bytMode As Byte
    
    '检查
    If mblnReportControlFocus Then
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "请选中一个报表类！", vbInformation, App.Title
            Exit Sub
        End If
    Else
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
        
        lngProgID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("程序ID")))
        strCode = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("编号"))
        strDescription = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("说明"))
    End If
        
    If mblnReportControlFocus Then
        '报表类
        bytMode = Val("2-报表类")
        lngProgID = 0
        strCode = ""
        With rptClass.FocusedRow
            lngGroupID = Val(Nvl(.Record(mobjClass.GetColIndex("上级ID")).Value, 0))
            lngID = Val(Nvl(.Record(mobjClass.GetColIndex("ID")).Value, 0))
            strName = .Record(mobjClass.GetColIndex("名称")).Value
            strDescription = Nvl(.Record(mobjClass.GetColIndex("说明")).Value)
        End With
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("组名"))
        bytMode = Val("1-报表组")
    Else
        If UCase(vsfTemp.name) = "VSFGROUPDETAIL" Or mbytReportGroup = 1 Then
            bytMode = Val("3-子报表")
        Else
            bytMode = 0
        End If
        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("名称"))
    End If
    
    If mblnReportControlFocus Then
    Else
        If blnGroup Then
            lngGroupID = lngID
            lngID = 0
        Else
            lngGroupID = 0
        End If
    End If
    
    '修改报表
    If frmReportEdit.ShowMe(Me, glngSys, bytMode, lngProgID, lngGroupID, lngID, strName, strCode, strDescription) Then
        If mblnReportControlFocus Then
            '刷新分类控件
            Call FillData(1, False)
        End If
        
        '刷新
        rptClass.Tag = ""
        Call RefreshEx
    End If
    Unload frmReportEdit
    Exit Sub
    
hErr:
    Call ErrCenter
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub Design()
    Dim lngID As Long, lngindex As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup = True Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If

    If CheckPass(lngID) = False Then
        MsgBox "报表数据错误，不能设计该报表！", vbInformation, App.Title
        Exit Sub
    End If
    If CheckReportPriv(lngID) = False Then
        MsgBox "你没有权限查询该报表某些数据源中的对象，请在设计环境下修正！", vbInformation, App.Title
    End If
    
    frmDesign.lngRPTID = lngID
    
    On Error Resume Next
    frmDesign.Show vbModal, Me
    On Error GoTo hErr
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub vsfGroup_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton(1)
End Sub

Private Sub vsfGroup_DblClick()
    mblnReportControlFocus = False
    Call Modify
End Sub

Private Sub vsfGroup_GotFocus()
    mblnReportControlFocus = False
    Call VisibleToolButton(1)
End Sub

Private Sub vsfGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And vsfGroup.Rows > 1 Then
        Call vsfGroup.Drag
    End If
End Sub

Private Sub vsfGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfGroup.Visible And vsfGroup.Enabled Then vsfGroup.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("2-报表组菜单"))
    End If
End Sub

Private Sub vsfGroupDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    Call UpdateStatusBar(vsfGroupDetail)
End Sub

Private Sub vsfGroupDetail_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfGroupDetail_DblClick()
    Call Design
End Sub

Private Sub vsfGroupDetail_GotFocus()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfGroupDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfGroupDetail.Visible And vsfGroupDetail.Enabled Then vsfGroupDetail.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("1-报表菜单"))
    End If
End Sub

Private Sub vsfReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    If NewRow <> OldRow Then Call UpdateStatusBar(vsfReport)
End Sub

Private Sub vsfReport_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfReport_DblClick()
    Dim cbcTemp As CommandBarControl
    
    Set cbcTemp = cbsMain.FindControl(, enuMenus.设计报表, , True)
    If Not cbcTemp Is Nothing Then
        cbcTemp.Execute
    End If
End Sub

Private Sub SetToolBarButtonIcon(ByVal lngMenuID As Long, ByVal lngIcon As Long)
    Dim cbcControl As CommandBarControl
    
    Set cbcControl = cbsMain.FindControl(, lngMenuID, , True)
    If Not cbcControl Is Nothing Then
        If cbcControl.IconId <> lngIcon Then cbcControl.IconId = lngIcon
    End If
End Sub

Private Sub vsfReport_GotFocus()
    mblnReportControlFocus = False
End Sub

Private Sub vsfReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mbytReportGroup = 0 And vsfReport.Rows > 1 Then
            Call vsfReport.Drag
        End If
    End If
End Sub

Private Sub vsfReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfReport.Visible And vsfReport.Enabled Then vsfReport.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("1-报表菜单"))
    End If
End Sub

Private Sub PopupMenuEx(ByVal bytType As Byte)
    Dim cbrTmp As XtremeCommandBars.CommandBar
    Dim cbbTmp As XtremeCommandBars.CommandBarButton
    Dim cbpTmp As XtremeCommandBars.CommandBarPopup
    
    Select Case bytType
    Case Val("1-报表菜单或子报表菜单")
        Set cbrTmp = cbsMain.Add("报表", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.新增报表, "新增报表")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.修改报表, "修改报表")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.删除报表, "删除报表")
            
            Set cbbTmp = .Add(xtpControlButton, enuMenus.设计报表, "设计报表"): cbbTmp.BeginGroup = True
            Set cbbTmp = .Add(xtpControlButton, enuMenus.执行报表, "执行报表")
            
            If glngSys = 0 Then
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.报表发布, "报表发布"): cbpTmp.BeginGroup = True
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至导航台菜单, "至导航台菜单(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至模块内菜单, "至模块内菜单(&2)")
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.取消发布, "取消发布")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从导航台菜单, "从导航台菜单(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从模块内菜单, "从模块内菜单(&2)")
                
                Set cbbTmp = .Add(xtpControlButton, enuMenus.报表启用, "启用(&S)"): cbbTmp.BeginGroup = True
                Set cbbTmp = .Add(xtpControlButton, enuMenus.报表停用, "停用(&T)")
            End If
        End With
    Case Val("2-报表组菜单")
        Set cbrTmp = cbsMain.Add("报表组", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.新增报表组, "新增报表组(&N)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.修改报表组, "修改报表组(&M)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.删除报表组, "删除报表组(&D)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.执行报表, "执行报表组"): cbbTmp.BeginGroup = True
            
            If glngSys = 0 Then
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.报表发布, "报表发布"): cbpTmp.BeginGroup = True
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至导航台菜单, "至导航台菜单(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.至模块内菜单, "至模块内菜单(&2)")
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.取消发布, "取消发布")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从导航台菜单, "从导航台菜单(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.从模块内菜单, "从模块内菜单(&2)")
                    
                Set cbbTmp = .Add(xtpControlButton, enuMenus.报表启用, "启用(&S)"): cbbTmp.BeginGroup = True
                Set cbbTmp = .Add(xtpControlButton, enuMenus.报表停用, "停用(&T)")
            End If
        End With
    Case Val("3-报表类菜单")
        Set cbrTmp = cbsMain.Add("报表类", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.新增报表类, "新增报表分类(&N)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.修改报表类, "修改报表分类(&M)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.删除报表类, "删除报表分类(&D)")
        End With
    End Select
    
    If Not cbrTmp Is Nothing Then
        Call cbrTmp.ShowPopup
    End If
End Sub

Private Sub NewEx()
    Dim lngProgID As Long, lngGroupID As Long, lngID As Long, l As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim strCode As String
    Dim bytMode As Byte
    
    '检查
    If mblnReportControlFocus Then
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "请选中一个报表类！", vbInformation, App.Title
            Exit Sub
        End If
    Else
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            '缺省控件
            On Error Resume Next
            vsfReport.SetFocus
            If Err.Number = 0 Then
                Set vsfTemp = Me.vsfReport
            Else
                MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
                Exit Sub
            End If
            On Error GoTo 0
        End If
    End If

    If mblnReportControlFocus Then
        '报表类
        bytMode = Val("2-报表类")
        With rptClass.FocusedRow
            lngGroupID = Val(Nvl(.Record(mobjClass.GetColIndex("上级ID")).Value, 0))
        End With
    ElseIf UCase(vsfTemp.name) = "VSFGROUPDETAIL" Then
        bytMode = Val("0-报表")
        lngProgID = Val(vsfGroup.TextMatrix(vsfGroup.Row, vsfGroup.ColIndex("程序ID")))
        If lngProgID <= 0 Then
            MsgBox "报表组“程序ID”数据异常，请联系管理！", vbInformation, App.Title
            Exit Sub
        End If
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        bytMode = Val("1-报表组")
    Else
        bytMode = Val("0-报表")
    End If
    
    If frmReportEdit.ShowMe(Me, glngSys, bytMode, lngProgID, lngGroupID, , , strCode) Then
        If mblnReportControlFocus Then
            '刷新分类控件
            Call FillData(1, False)
        Else
            If (UCase(vsfTemp.name) = "VSFREPORT" Or UCase(vsfTemp.name) = "VSFGROUPDETAIL") Then
                '刷新
                rptClass.Tag = ""
                Call RefreshEx
                
                '定位
                For l = 1 To vsfTemp.Rows - 1
                    If UCase(strCode) = UCase(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("编号"))) Then
                        '设计
                        vsfTemp.Row = l
                        
                        If MsgBox("需要立即设计报表吗？", vbQuestion + vbDefaultButton1 + vbYesNo) = vbYes Then
                            Call Design
                        End If
                        
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Remove()
    Dim lngID As Long, lngGroupID As Long, lngSelRow As Long, lngRow As Long
    Dim rsGroup As ADODB.Recordset
    Dim colSQL As New Collection
    Dim strRec As String, strSQL As String
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "请选中子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    strRec = "": lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        If lngSelRow <= 4 Then
            If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                strRec = strRec & vbCrLf & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("名称"))
                lngSelRow = lngSelRow + 1
            End If
        Else
            strRec = strRec & "|..."
            Exit For
        End If
    Next
    
    If MsgBox("你确定要从报表组中移除以下报表吗？" & strRec _
        , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        '获取子表的组表
        If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
            lngID = Val(vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("ID")))
            Set rsGroup = GetReportGroups(lngID)
            Do While rsGroup.EOF = False
                lngGroupID = Nvl(rsGroup!id, 0)
                
                strSQL = "Update zlRPTSubs Set 序号=序号-1 " & _
                         "Where 序号>(Select 序号 From zlRPTSubs Where 报表ID=" & lngID & " And 组ID=" & lngGroupID & ")" & _
                         "    And 组ID=" & lngGroupID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Delete From zlRPTSubs Where 报表ID=" & lngID & " And 组ID=" & lngGroupID
                Call AddArray(colSQL, strSQL)
                
                rsGroup.MoveNext
            Loop
            rsGroup.Close
            
            lngSelRow = lngSelRow + 1
        End If
    Next
    
    '执行DML
    gcnOracle.BeginTrans: blnTrans = True
    For lngRow = 1 To colSQL.count
        gcnOracle.Execute colSQL(lngRow)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
End Sub

Private Function GetReportGroups(ByVal lngID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = _
        "Select a.Id, a.编号, a.名称 " & vbNewLine & _
        "From zlRPTGroups A, zlRPTSubs B " & vbNewLine & _
        "Where a.Id = b.组id And 系统 Is Null And b.报表id = [1] " & vbNewLine & _
        "Order By a.名称"
    Set GetReportGroups = mdlPublic.OpenSQLRecord(strSQL, "获取报表组信息", lngID)
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Guide()
    Dim objReport As Report, objItem As Object
    Dim objCBS_ComBox As CommandBarComboBox
    Dim objControl As Object
    Dim lngNextID As Long
    Dim strSQL As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    Set objReport = New Report
    With objReport
        '进纸；15-缺省为自动选择
        .进纸 = 15
        '缺省使用当前打印机
        If Printers.count > 0 Then .打印机 = Printer.DeviceName
        '缺省为A4幅面,为纵向
        .Fmts.Add 1, "格式1", INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    End With
    
    frmGuide.blnNew = True
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts(1)
    frmGuide.Show vbModal, Me
    
    If gblnOK Then
        Set objControl = cbsMain.FindControl(, enuMenus.选择系统控件, , True)
        If Not objControl Is Nothing Then
            '恢复至系统共享选项
            objControl.ListIndex = 1
            
            '刷新界面
            Call SelectedSysComboBox(objControl)
        End If
        
        '生成报表
        With frmGuide
            Set objReport.Items = .objGuide.Items       '报表元素对象集合
            Set objReport.Datas = .objGuide.Datas       '报表数据源对象集合
            Set objReport.Fmts = .objGuide.Fmts         '报表格式对象集合
            
            lngNextID = GetNextID("zlReports")
            strSQL = "Insert Into zlReports(ID,编号,名称,说明,系统,密码) " & vbCrLf & _
                     "Values (" & _
                        lngNextID & _
                        ",'" & .txtNO.Text & "'" & _
                        ",'" & .txtTitle.Text & "'" & _
                        ",'" & .txtNote.Text & "'" & _
                        "," & IIF(glngSys = 0, "NULL", glngSys) & _
                        "," & AdjustStr(GetPass(.txtNO, .txtTitle)) & ")"
                        
            On Error GoTo hErr
            
            gcnOracle.BeginTrans: blnTrans = True
            gcnOracle.Execute strSQL
            gcnOracle.CommitTrans: blnTrans = False
            
            '报表内容
            If Not SaveReport(lngNextID, objReport, staMain.Panels(2)) Then
                gcnOracle.BeginTrans: blnTrans = True
                gcnOracle.Execute "Delete From zlReports Where ID=" & lngNextID
                gcnOracle.CommitTrans: blnTrans = False
                
                MsgBox "在生成报表时遇到意外错误,请重试该操作！", vbInformation, App.Title
                Unload frmGuide
                Exit Sub
            End If

        End With
        Unload frmGuide
        
        '刷新
        rptClass.Tag = ""
        Call RefreshEx
    End If
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Unload frmGuide
End Sub

Private Sub SelectedSysComboBox(ByVal objControl As XtremeCommandBars.CommandBarComboBox)
    Dim lngID As Long

    If objControl Is Nothing Then
        glngSys = 0
        GoSub proPub
        Exit Sub
    End If
    
    '更新界面
    If objControl.ListIndex > Val("1-系统共享") Then
        If dkpMain.Panes(Val("1-报表类")).Closed = False Then dkpMain.Panes(Val("1-报表类")).Close
        rptClass.FocusedRow = rptClass.Rows(0)
    Else
        dkpMain.ShowPane Val("1-报表类")
    End If
    
    '更新变量
    glngSys = objControl.ItemData(objControl.ListIndex)
    GoSub proPub
    
    '更新界面
    rptClass.Tag = ""
    Call rptClass_SelectionChanged
    Exit Sub
    
proPub:
    mblnAllowPublish = False
    Select Case glngSys
    Case Val("0-共享"), Val("200-人事"), Val("500-成本"), Val("700-账务")
        mblnAllowPublish = True
    Case Else
        If Val(GetSysVersion(lngID)) >= 10 Then
            mblnAllowPublish = True
        End If
    End Select
    Return
End Sub

Private Sub ShowRunLog()
    Dim lngID As Long
    Dim strName As String
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中要查看日志的独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "请选中要查看日志的独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    strName = Trim(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("名称")))
    
    '查看报表运行日志记录
    If lngID > 0 Then
        Call frmReportRunLog.ShowMe(Me, lngID, "报表“" & strName & "”的运行日志")
    End If
End Sub

Private Sub VisibleToolButton(Optional ByVal bytMode As Byte = 0)
'功能：更新工具栏“新增、修改、删除”按钮显示
'功能：
'  bytMode：0-报表；1-报表组；2-报表类

    Dim objTemp As Object
    
    Select Case bytMode
    Case 1
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.新增报表类, enuMenus.修改报表类, enuMenus.删除报表类 _
                , enuMenus.新增报表, enuMenus.修改报表, enuMenus.删除报表
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    Case 2
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.新增报表组, enuMenus.修改报表组, enuMenus.删除报表组 _
                , enuMenus.新增报表, enuMenus.修改报表, enuMenus.删除报表
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    Case Else
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.新增报表组, enuMenus.修改报表组, enuMenus.删除报表组 _
                , enuMenus.新增报表类, enuMenus.修改报表类, enuMenus.删除报表类
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    End Select
End Sub

Private Function GetReportObjects(ByVal lngID As Long, Optional ByVal AllConnect As Boolean = False) As String
'功能：获取指定报表的数据库访问对象
'参数：
'  lngID：报表ID
'  AllConnect：True-所有数据连接；False-当前数据连接
'返回：对象名称
    
    Dim strObjects As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim l As Long, k As Long
    Dim arrTmp As Variant, arrItem As Variant
    
    On Error GoTo hErr

    '分析报表的数据源访问对象
    strObjects = ""
    strSQL = "Select 对象 From zlRPTDatas " & _
             "Where 对象 is Not NULL And 报表ID=[1] " & _
             IIF(AllConnect, "", "And nvl(数据连接编号,0) <= 0")
    Set rsTmp = OpenSQLRecord(strSQL, "获取指定报表的数据库访问对象名称", lngID)
    Do While rsTmp.EOF = False
        arrTmp = Split(rsTmp!对象, ",")
        For l = LBound(arrTmp) To UBound(arrTmp)
            If InStr(strObjects & ",", "," & arrTmp(l) & ",") = 0 Then
                strObjects = strObjects & "," & arrTmp(l)
            End If
        Next
        rsTmp.MoveNext
    Loop
    rsTmp.Close

    '分析该报表的参数数据源访问对象
    strSQL = "Select B.对象 From zlRPTDatas A,zlRPTPars B " & _
             "Where A.ID=B.源ID And B.对象 is Not NULL And A.报表ID=[1] " & _
             IIF(AllConnect, "", "And nvl(数据连接编号,0) <= 0")
    Set rsTmp = OpenSQLRecord(strSQL, "获取指定报表所有参数的数据库访问对象名称", lngID)
    Do While rsTmp.EOF = False
        arrTmp = Split(rsTmp!对象, "|")
        For l = LBound(arrTmp) To UBound(arrTmp)
            arrItem = Split(arrTmp(l), ",")
            For k = LBound(arrItem) To UBound(arrItem)
                If InStr(strObjects & ",", "," & arrItem(k) & ",") = 0 Then
                    strObjects = strObjects & "," & arrItem(k)
                End If
            Next
        Next
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    GetReportObjects = strObjects
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ReportGrantToNavigator()
'功能：发布当前报表(组)到导航台,可能不是第一次
'说明：不支持发布所有报表；支持发布多选或全选报表；

    Dim lngID  As Long, lngReportID As Long, lngGroupID As Long, lngProgID As Long, lngTemp As Long
    Dim l As Long, k As Long, lngSelRow As Long, lngMenuID As Long, lngMenuPID As Long
    Dim blnGroup As Boolean, blnResult As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim rsReport As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, strName As String, strObject As String, strTmp As String
    Dim strDesc As String, strOwner As String
    Dim objNode As Node
    Dim arrTmp As Variant
    Dim colSQL As Collection
    
    mblnReportControlFocus = False
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '选择一个菜单位置
    Set rsTmp = GetMainTreeMenu
    If rsTmp Is Nothing Then
        MsgBox "读取菜单体系时遇到异常，“报表发布”中断！", vbInformation, App.Title
        Exit Sub
    End If
    
    Load frmSelTree
    If blnGroup Then
        frmSelTree.Caption = "发布报表组到导航台 - 菜单位置选择"
    Else
        frmSelTree.Caption = "发布报表到导航台 - 菜单位置选择"
    End If
    
    With frmSelTree.tvw
        .Nodes.Clear
        Do While rsTmp.EOF = False
            If rsTmp!Flag = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!标题, "Root")
                objNode.Tag = "请选择本系统下一个具体的菜单位置！"
            Else
                If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                    Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, rsTmp!标题, "Path")
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, rsTmp!标题 _
                                        , IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                    objNode.ForeColor = vbBlue
                    objNode.Tag = "这是已发布的报表,选择一个菜单位置！"

                    '不能发布到相同位置
                    If blnGroup Then
                        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("组名"))
                        If objNode.Text = strName Then
                            objNode.Parent.Tag = "同一个报表或组不能发布到相同的位置,请选择其他菜单位置！"
                        End If
                    Else
                        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("名称"))
                        If objNode.Text = strName Then
                            objNode.Parent.Tag = "同一个报表或组不能发布到相同的位置,请选择其他菜单位置！"
                        End If
                    End If
                End If
            End If
            objNode.Expanded = True
            
            rsTmp.MoveNext
        Loop
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
        rsTmp.Close
    End With
    
    frmSelTree.Show vbModal, Me
    If Not gblnOK Then Exit Sub
    lngMenuPID = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2))      '菜单的上级ID
    Unload frmSelTree
    
    '选择的1..N报表或报表组
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        Set colSQL = New Collection
        lngMenuID = GetNextID("zlMenus")    '菜单ID
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        If blnGroup Then
            '报表组
            
            '检查
            lngGroupID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
            strSQL = _
                "Select a.Id, a.名称 " & vbCrLf & _
                "From zlReports A, zlRPTSubs B " & vbCrLf & _
                "Where a.Id = b.报表id And b.组id = [1] "
            Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "获取报表组的报表ID", lngGroupID)
            Do While rsTmp.EOF = False
                lngReportID = Nvl(rsTmp!id, 0)
                strName = Nvl(rsTmp!名称)
                GoSub proVerify
                If blnResult = False Then
                    Exit Sub
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
            
            '数据组织
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("组名"))
            strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("说明"))
            lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("程序ID")))
            If lngProgID = 0 Then
                '首次发布
                lngProgID = GetNewProgID()          '程序ID
                
                strSQL = _
                    "Update zlRPTSubs A Set 功能=(Select 名称 From zlReports Where ID=A.报表ID) Where 组ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Update zlRPTGroups Set 程序ID=" & lngProgID & ",发布时间=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlPrograms(序号,标题,说明,系统,部件) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    "(" & lngProgID & _
                    ",'" & strName & "'" & _
                    "," & IIF(strDesc = "", "Null", "'" & strDesc & "'") & _
                    "," & IIF(glngSys <= 0, "Null", glngSys) & _
                    ",'zl9Report' ) "
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlProgFuncs(系统,序号,功能,说明)" & vbCrLf & _
                    "Select " & _
                    IIF(glngSys <= 0, "Null", glngSys) & _
                    "," & lngProgID & ",名称,说明 From zlReports " & vbCrLf & _
                    "Where ID In (Select 报表ID From zlRPTSubs Where 组ID=" & lngID & ")"
                Call AddArray(colSQL, strSQL)
                
                '循环读取各子报表的权限
                strSQL = "Select A.报表ID,B.名称 From zlRPTSubs A,zlReports B Where A.组ID=[1] And A.报表ID=B.ID"
                Set rsReport = OpenSQLRecord(strSQL, "获取报表组子报表信息", lngID)
                Do While rsReport.EOF = False
                    strObject = GetReportObjects(rsReport!报表ID)
                    If strObject <> "" Then
                        '表格有可能不访问数据库
                        strObject = Mid(strObject, 2)
                        arrTmp = Split(strObject, ",")
                        For k = LBound(arrTmp) To UBound(arrTmp)
                            strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                            If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                                strTmp = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                                strSQL = GetInsertProgPrivs(glngSys, lngProgID, rsReport!名称, strTmp, strOwner, "SELECT")
                                Call AddArray(colSQL, strSQL)
                            End If
                        Next
                    End If
    
                    rsReport.MoveNext
                Loop
                rsReport.Close
            Else
                '重新发布
                strSQL = "Update zlRPTGroups Set 发布时间 = Sysdate Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
        Else
            '独立报表、子报表
            
            '检查
            lngReportID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
            GoSub proVerify
            If blnResult = False Then
                Exit Sub
            End If
            
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
            strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("说明"))
            lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("程序ID")))
            If lngProgID = 0 Then
                '首次发布
                
                '新生成程序ID
                strObject = GetReportObjects(lngID)
                If strObject <> "" Then strObject = Mid(strObject, 2)
                lngProgID = GetNewProgID()
                
                strSQL = _
                    "Update zlReports " & vbCrLf & _
                    "Set 功能 = '基本', 程序ID = " & lngProgID & ", 发布时间 = Sysdate " & vbCrLf & _
                    "Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlPrograms(序号,标题,说明,系统,部件) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    " (" & lngProgID & _
                    ",'" & strName & "'" & _
                    ",'" & strDesc & "'" & _
                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                    ",'zl9Report' )"
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlProgFuncs(系统,序号,功能) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    "(" & IIF(glngSys <= 0, "Null", glngSys) & _
                    "," & lngProgID & _
                    ",'基本')"
                Call AddArray(colSQL, strSQL)
                
                If strObject <> "" Then
                    '该表格有可能不访问数据库
                    arrTmp = Split(strObject, ",")
                    For k = LBound(arrTmp) To UBound(arrTmp)
                        strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                        If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                            strTmp = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                            strSQL = GetInsertProgPrivs(glngSys, lngProgID, "基本", strTmp, strOwner, "SELECT")
                            Call AddArray(colSQL, strSQL)
                        End If
                    Next
                End If
               
            Else
                '重复发布
                strSQL = "Update zlReports Set 发布时间 = Sysdate Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
            End If
        End If
        
        '生成菜单数据
        strSQL = _
            "Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) " & vbCrLf & _
            "Values('缺省'," & lngMenuID & "," & lngMenuPID & ",'" & strName & "',Null" & _
            "," & IIF(strDesc = "", "Null", "'" & strDesc & "'") & _
            "," & IIF(glngSys <= 0, "Null", glngSys) & _
            "," & lngProgID & ",'" & strName & "',105)"
        Call AddArray(colSQL, strSQL)
        
        '执行DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            'Debug.Print colSQL(k)
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1

makContinue:
    Next
    
    Set grsReport = Nothing
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx

    Exit Sub
    
proVerify:
    blnResult = False
    
    '验证密码
    If CheckPass(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("【[1]】报表验证不能通过，拒绝发布！", strName) _
            , vbInformation, App.Title
        Return
    End If
    
    '权限
    If CheckReportPriv(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("你没有【[1]】报表中数据源涉及数据库对象的查询权限，请检查！", strName) _
            , vbInformation, App.Title
        Return
    End If
    
    blnResult = True
    Return
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then Resume
    End If
    Call SaveErrLog
    
    Set grsReport = Nothing
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
End Sub

Private Function GetMainTreeMenu(Optional ByVal lngProgID As Long) As ADODB.Recordset
'功能：获取发布到导航台报表树形菜单体系
'参数：lngProgID=是否只显示指定程序ID的报表
'说明：菜单体系中包含自定义报表发布的菜单项(如果有),标志为"FLAG=999"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngSys As Long
    
    On Error GoTo errH
    
    lngSys = glngSys
    If lngSys = 0 Then
        '只显示用户发布部份报表
        strSQL = _
            "Select Distinct * From (" & _
            " Select 编号 as SCOL,0 as Flag,-编号 as ID,-NULL as 上级ID,'['||编号||']'||名称 as 标题,-NULL as 模块 " & _
            " From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(上级ID,-系统) as 上级ID,标题,模块 " & _
            " From zlMenus Where 组别='缺省' And 模块 is NULL" & _
            " Start With 上级ID is NULL And 组别='缺省' Connect by Prior ID=上级ID And 组别='缺省'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            "     And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            "     And A.系统 is NULL And B.系统 is Null And C.系统 is Null" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlReports C" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            "     And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            "     And A.系统 is NULL And B.系统 is Null And C.系统 is Null" & _
            ") Order by SCOL,FLAG,ID"
    Else
        '只显示固定部份报表(已授权部份)
        strSQL = _
            "Select Distinct * From (" & _
            " Select 编号 as SCOL,0 as Flag,-编号 as ID,-NULL as 上级ID,'['||编号||']'||名称 as 标题,-NULL as 模块 " & _
            " From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(上级ID,-系统) as 上级ID,标题,模块 " & _
            " From zlMenus Where 组别='缺省' And 模块 is NULL" & _
            " Start With 上级ID is NULL And 组别='缺省' Connect by Prior ID=上级ID And 组别='缺省'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            "     And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            "     And A.系统=B.系统 And A.系统=C.系统 And Trunc(B.系统/100)=D.系统 And B.序号=D.序号" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlReports C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            "     And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            "     And A.系统=B.系统 And A.系统=C.系统 And Trunc(B.系统/100)=D.系统 And B.序号=D.序号" & _
            ") Order by SCOL,FLAG,ID"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngProgID)
    Set GetMainTreeMenu = rsTmp
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewProgID() As Long
'功能：获取下一个可用的自定义报表程序号,用于发布
'说明：程序号从100000开始,并自动补缺
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Decode(Sign(Max(序号)-99999),1,Max(序号),99999) as ID From zlPrograms"
    Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "获取新模块ID（自定义报表）")
    GetNewProgID = IIF(IsNull(rsTmp!id), 100000, rsTmp!id + 1)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReportRevokeFromNavigator()
'功能：取消当前报表或报表组在导航台上的发布
'说明：
'  1.如果发布处于多个位置，则让使用者选择取消发布的一个位置，删除zlMenus对应位置内容
'  2.如果只有一个发布位置，则将zlReport中的程序ID=NULL，删除zlPrograms中的发布模块
'参数：
''''''  blnRevokeByProgram：True-取消模块；False-取消导航台

    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node
    Dim lngID As Long, lngProgID As Long, lngMenuID As Long, lngSelRow As Long
    Dim l As Long, k As Long
    Dim blnTrans As Boolean, blnGroup As Boolean
    Dim colSQL As Collection
    Dim vsfTemp As VSFlexGrid
    Dim strSQL As String, strName As String, strMsg As String, strTemp As String
    
    mblnReportControlFocus = False
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "请选中一个独立报表、报表组、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    If UCase(vsfTemp.name) = UCase("vsfGroup") Then
        strMsg = GetSelectedReport(vsfTemp, "组名")
    Else
        strMsg = GetSelectedReport(vsfTemp, "名称")
    End If
    If MsgBox("如果把以下报表从导航台菜单中取消发布，其他用户不能再使用该报表。要继续吗？" & _
        strMsg, vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '选择的1..N报表或报表组
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        '分析发布位置
        Set colSQL = New Collection
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("程序ID")))
        If lngProgID <= 0 Then
            If blnGroup Then
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("组名"))
                strTemp = mdlPublic.FormatString("【[1]】报表组没有发布到导航台菜单！", strName)
            Else
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
                strTemp = mdlPublic.FormatString("【[1]】报表没有发布到导航台菜单！", strName)
            End If
            If lngSelRow < vsfTemp.SelectedRows - 1 Then
                strTemp = strTemp & "要继续“取消发布”剩下的报表吗？"
                If MsgBox(strTemp, vbInformation + vbDefaultButton1 + vbYesNo) = vbYes Then
                    lngSelRow = lngSelRow + 1
                    GoTo makContinue
                Else
                    Exit Sub
                End If
            Else
                MsgBox strTemp, vbInformation, App.Title
                Exit Sub
            End If
        Else
            If blnGroup Then
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("组名"))
            Else
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
            End If
        End If
        
        '获取模块信息
        Set rsTmp = GetMainTreeMenu(lngProgID)
        rsTmp.Filter = "模块=" & lngProgID
        If rsTmp Is Nothing Then
            MsgBox "读取菜单体系时遇到异常，“取消发布”中断！", vbInformation, App.Title
            Exit Sub
        ElseIf rsTmp.EOF Or rsTmp.RecordCount = 1 Then
            If rsTmp.EOF Then
                '发布位置异常
                strTemp = mdlPublic.FormatString("【[1]】报表的发布处于不正常状态，可能是数据不正确引起！", strName)
                MsgBox strTemp, vbInformation, App.Title
            End If
            
            If blnGroup Then
                '报表组
                strSQL = "Update zlRPTGroups Set 程序ID=NULL,发布时间=NULL,是否停用=NULL Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Update zlRPTSubs A Set 功能=Null Where 组ID=" & lngID
                Call AddArray(colSQL, strSQL)
            Else
                '报表
                strSQL = _
                    "Update zlReports Set 功能=NULL,程序ID=NULL,是否停用=NULL,发布时间=NULL " & vbCrLf & _
                    "Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
            strSQL = "Delete From zlMenus Where 模块=" & lngProgID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlPrograms Where 序号=" & lngProgID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
        
            strSQL = "Delete From zlRoleGrant Where 序号=" & lngProgID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
        Else
            '多个发布位置
            rsTmp.Filter = ""
            
            '创建菜单树结构
            Load frmSelTree
            frmSelTree.Caption = "取消发布 - 导航台菜单位置"
            
            With frmSelTree.tvw
                .Nodes.Clear
                Do While rsTmp.EOF = False
                    If rsTmp!Flag = 0 Then
                        Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!标题, "Root")
                        objNode.Tag = "请在本系统下选择一个要取消发布的报表或组！"
                    Else
                        If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                            Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, rsTmp!标题, "Path")
                            objNode.Tag = "请在菜单上选择一个要取消发布的报表或组！"
                        Else
                            Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, rsTmp!标题, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                            objNode.ForeColor = vbBlue
                            If .SelectedItem Is Nothing Then
                                objNode.Selected = True
                            ElseIf .SelectedItem.Index = 1 Then
                                objNode.Selected = True
                            End If
                        End If
                    End If
                    objNode.Expanded = True

                    '标记有报表(组)的路径
                    If rsTmp!Flag = 999 Or rsTmp!Flag = 888 Then
                        objNode.SelectedImage = objNode.Image
                        Do While Not objNode.Parent Is Nothing
                            Set objNode = objNode.Parent
                            objNode.SelectedImage = objNode.Image
                        Loop
                    End If
    
                    rsTmp.MoveNext
                Loop
    
                '删除没有报表(组)的路径
                For k = .Nodes.count To 1 Step -1
                    If .Nodes(k).SelectedImage = "" Then
                        .Nodes.Remove k
                    End If
                Next
            End With
            
            frmSelTree.Show vbModal, Me
            If Not gblnOK Then Exit Sub
            lngMenuID = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2))     '报表菜单ID
            Unload frmSelTree
            
            If blnGroup Then
                strSQL = "Update zlRPTGroups Set 发布时间=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Update zlRPTSubs A Set 功能=Null Where 组ID=" & lngID
                Call AddArray(colSQL, strSQL)
            Else
                strSQL = "Update zlReports Set 发布时间=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
            strSQL = "Delete From zlMenus Where ID=" & lngMenuID & " And Nvl(系统,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
        End If
        
        '执行DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
'            Debug.Print colSQL(k)
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1
        
makContinue:
    Next
    
    Set grsReport = Nothing
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub

hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReportGrantToModule()
'功能：发布当前报表到模块,可能不是第一次
'说明：不支持报表组发布到模块

    Dim l As Long, j As Long, k As Long
    Dim lngSelRow As Long, lngID As Long, lngSys As Long, lngProgID As Long
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim rsModule As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strFunc As String, strIDs As String, strSQL As String, strObject As String
    Dim strDesc As String, strOwner As String, strName As String
    Dim objNode As Node
    Dim colSQL As Collection
    Dim arrTmp As Variant, arrItems As Variant
    
    mblnReportControlFocus = False
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    lngSelRow = 0
    strIDs = ""
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue1
        
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        If CheckPass(lngID) = False Then
            MsgBox "报表数据错误，不能发布该报表！", vbInformation, App.Title
            Exit Sub
        End If
        If Not CheckReportPriv(lngID) Then
            MsgBox "你没有权限查询该报表某些数据源中的对象，操作不能继续！", vbInformation, App.Title
            Exit Sub
        End If
        
        strIDs = strIDs & "," & CStr(lngID)
        lngSelRow = lngSelRow + 1
        
makContinue1:
    Next
    If strIDs = "" Then Exit Sub
    
    strIDs = Mid(strIDs, 2)
    Set rsModule = GetModuleTreeMenu(strIDs)
    If rsModule Is Nothing Then
        MsgBox "读取模块菜单体系时遇到意外错误，报表发布中断！", vbInformation, App.Title
        Exit Sub
    End If
    
    '1.选择一个菜单模块位置
    Load frmSelTree
    frmSelTree.Caption = "发布报表到模块 - 模块位置选择"
    With frmSelTree.tvw
        .Nodes.Clear
        Do While rsModule.EOF = False
            If IsNull(rsModule!上级ID) Then
                Set objNode = .Nodes.Add(, , "_" & rsModule!id, rsModule!标题)
            Else
                Set objNode = .Nodes.Add("_" & rsModule!上级ID, 4, "_" & rsModule!id, rsModule!标题)
            End If
            
            If Left(rsModule!id, 1) = "S" Then     'System
                objNode.Image = "Root"
                objNode.Tag = "请选择本系统中菜单下的模块位置。"
            ElseIf Left(rsModule!id, 1) = "T" Then 'MenuTree
                objNode.Image = "Path"
                objNode.Tag = "请选择本系统中菜单下的模块位置。"
            ElseIf Left(rsModule!id, 1) = "M" Then 'Module
                objNode.Image = "App"
            ElseIf Left(rsModule!id, 1) = "R" Then 'Report
                objNode.Image = "ReportNode"
                objNode.ForeColor = vbBlue
                objNode.Tag = "这是已发布的报表,选择其他菜单下的模块位置。"
                objNode.Parent.Tag = "报表不能重复发布到同一个模块,请选择其他模块。"
            End If
            objNode.Expanded = True
            
            '标记有下级模块的菜单(用SQL较慢)
            If Left(rsModule!id, 1) = "M" Then
                If objNode.Parent.SelectedImage = "" Then
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
            End If
            
            rsModule.MoveNext
        Loop
        
        '删除无下级模块的空菜单
        For k = .Nodes.count To 1 Step -1
            If .Nodes(k).SelectedImage = "" And Mid(.Nodes(k).Key, 2, 1) = "T" Then
                .Nodes.Remove k
            End If
        Next
        
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    
    frmSelTree.Show vbModal, Me
    If Not gblnOK Then
        Exit Sub
    End If
    rsModule.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
    If rsModule.EOF Then Exit Sub
    
    '数据重复检查
    lngSys = rsModule!系统
    lngProgID = rsModule!程序id
    Unload frmSelTree
    rsModule.Close
    
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue2

        strFunc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strSQL = _
            "Select 功能 From zlRPTPuts Where 报表ID=[1] And 系统=[2] And 程序ID=[3] " & vbCrLf & _
            "Union ALL " & vbCrLf & _
            "Select 功能 From zlProgFuncs Where 系统=[2] And 序号=[3] And 功能=[4] "
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngID, lngSys, lngProgID, strFunc)
        If Not rsTmp.EOF Then
            MsgBox "报表发布位置或发布功能重复，数据库中的数据可能不正确。", vbInformation, App.Title
            Exit Sub
        End If
        rsTmp.Close

        lngSelRow = lngSelRow + 1

makContinue2:
    Next

    '2.授权权限分析

    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue3

        strObject = ""
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strFunc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
        strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("说明"))

        strObject = GetReportObjects(lngID)
        If strObject <> "" Then strObject = Mid(strObject, 2)
        
        '3.填写程序、权限
        Set colSQL = New Collection
        
        strSQL = "Update zlReports Set 发布时间=Sysdate Where ID=" & lngID
        Call AddArray(colSQL, strSQL)
        
        strSQL = _
            "Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) " & vbCrLf & _
            "Values(" & lngID & _
            "," & lngSys & _
            "," & lngProgID & _
            ",'" & strFunc & "')"
        Call AddArray(colSQL, strSQL)
        
        strSQL = _
            "Insert Into zlProgFuncs(系统,序号,功能,说明) " & vbCrLf & _
            "Values(" & lngSys & _
            "," & lngProgID & _
            ",'" & strFunc & "'" & _
            ",'" & strDesc & "')"
        Call AddArray(colSQL, strSQL)
        
        If strObject <> "" Then
            '该表格有可能不访问数据库
            arrTmp = Split(strObject, ",")
            For k = 0 To UBound(arrTmp)
                strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                    strName = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                    strSQL = GetInsertProgPrivs(lngSys, lngProgID, strFunc, strName, strOwner, "SELECT")
                    Call AddArray(colSQL, strSQL)
                End If
            Next
        End If
        
        '执行DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False

        lngSelRow = lngSelRow + 1

makContinue3:
    Next
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetModuleTreeMenu(ByVal strIDs As String) As ADODB.Recordset
'功能：获取发布到模块的报表树形菜单体系
'参数：
'  strIDs：要发布或取消发布的报表ID。格式：id1[,id2[...]]

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '按菜单显示模块的方式
    '-------------------------------------------------------------------------------------------------------------
    '1.系统 + 2.中间菜单 + 3.模块菜单(授权模块) + 4.发布报表(发布到授权模块下)
    '注意同一模块可能重复位于不同菜单,都显示(包括下面的报表)
    '排开独立的自定义报表模块(部件='zl9Report')
    
    '只显示模块的方式
    strSQL = _
        " Select '1' as Sort1,To_Char(编号) as Sort2, 'S'||编号 as ID,Null as 上级ID,编号 as 系统" & vbCrLf & _
        "     ,-Null as 程序ID,Null as 功能,'['||编号||']'||名称 as 标题" & vbCrLf & _
        " From zlSystems" & vbCrLf & _
        " Union ALL " & vbCrLf & _
        " Select '3' as Sort1,To_Char(B.序号) as Sort2, 'M'||B.序号||'_'||B.系统 as ID,'S'||B.系统 as 上级ID,B.系统" & vbCrLf & _
        "     ,B.序号 as 程序ID,Null as 功能,'['||B.序号||']'||B.标题" & vbCrLf & _
        " From zlPrograms B,(Select 系统,序号 From zlRegFunc Group By 系统,序号) C" & vbCrLf & _
        " Where Upper(B.部件)<>Upper('zl9Report') And Trunc(B.系统/100)=C.系统 And B.序号=C.序号" & vbCrLf & _
        " Union All " & vbCrLf & _
        " Select '4' as Sort1,C.编号 as Sort2," & vbCrLf & _
        "   'R'||Rownum as ID,'M'||B.程序ID||'_'||B.系统 as 上级ID,B.系统,B.程序ID,B.功能,'['||C.编号||']'||C.名称 as 标题" & vbCrLf & _
        " From zlPrograms A,zlRPTPuts B,zlReports C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & vbCrLf & _
        "    , Table(Cast(f_Num2List([1]) as T_NumList)) E" & vbCrLf & _
        " Where A.系统=B.系统 And A.序号=B.程序ID And Upper(A.部件)<>Upper('zl9Report')" & vbCrLf & _
        "   And Trunc(A.系统/100)=D.系统 And A.序号=D.序号" & vbCrLf & _
        "   And B.报表ID=C.ID And c.Id = e.Column_Value " & vbCrLf & _
        " Order by Sort1,Sort2"
    
    '固定允许发布到人事、成本、帐务系统的模块，其它系统仅10版本允许
    strSQL = "Select A.* From (" & strSQL & ") A,zlSystems B" & _
        " Where A.系统=B.编号 And (To_Number(Substr(B.版本号,1,Instr(B.版本号,'.')-1))>=10 Or Trunc(编号/100) IN(2,5,7))" & _
        " Order by Sort1,Sort2"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strIDs)
    Set GetModuleTreeMenu = rsTmp
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetSelectedReport(ByVal vsfVar As VSFlexGrid, ByVal strColName As String) As String
    Dim strResult As String
    Dim l As Long, lngSelRow As Long
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For l = 1 To vsfVar.Rows - 1
        If vsfVar.SelectedRow(lngSelRow) = l Then
            strResult = strResult & vbCrLf & CStr(lngSelRow + 1) & "." & vsfVar.TextMatrix(l, vsfVar.ColIndex(strColName))
            lngSelRow = lngSelRow + 1
        End If
        If lngSelRow >= 5 Then
            '最多显示5个信息
            strResult = strResult & vbCrLf & "..."
            Exit For
        End If
    Next
    GetSelectedReport = strResult
    Exit Function
    
hErr:
    Call ErrCenter
End Function

Private Sub ReportRevokeFromModule()
'功能：取消当前报表在模块上的一个发布
'说明：
'  1.如果发布位置大于1，则让使用者选择取消发布的一个位置
'  2.如果只有一个发布位置，则直接提示处理

    Dim rsTmp As ADODB.Recordset
    Dim strFunc As String, strSQL As String, strMsg As String, strName As String
    Dim strTmp As String
    Dim objNode As Node
    Dim lngID As Long, lngSys As Long, lngProgID As Long, lngSelRow As Long
    Dim l As Long, k As Long
    Dim blnTrans As Boolean, blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim colSQL As Collection
    
    mblnReportControlFocus = False
    
    '检查
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup Then
        MsgBox "请选中一个独立报表、子报表！", vbInformation, App.Title
        Exit Sub
    End If
    
    If UCase(vsfTemp.name) = UCase("vsfGroup") Then
        strMsg = GetSelectedReport(vsfTemp, "组名")
    Else
        strMsg = GetSelectedReport(vsfTemp, "名称")
    End If
    If MsgBox("如果把以下报表从该模块菜单中取消发布，其他用户不能再使用该报表。要继续吗？" & _
        strMsg, vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("名称"))
        strFunc = strName
        Set colSQL = New Collection
        
        '1.分析当前发布位置
        strSQL = "Select 系统,程序ID,功能 From zlRPTPuts Where 报表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngID)
        If rsTmp.EOF Then
            rsTmp.Close
            strTmp = mdlPublic.FormatString("当前报表【[1]】没有发布到模块中", strName)
            If vsfTemp.SelectedRows - 1 <= lngSelRow Then
                strTmp = strTmp & "！"
                MsgBox strTmp, vbInformation, App.Title
                rptClass.Tag = ""
                Call RefreshEx
                Exit Sub
            Else
                strTmp = strTmp & "，要继续“取消发布”剩下的报表吗？"
                If MsgBox(strTmp, vbInformation + vbDefaultButton1 + vbYesNo, App.Title) = vbYes Then
                    lngSelRow = lngSelRow + 1
                    GoTo makContinue
                Else
                    rptClass.Tag = ""
                    Call RefreshEx
                    Exit Sub
                End If
            End If
        ElseIf rsTmp.RecordCount = 1 Then
            '只有一个发布位置
            lngSys = rsTmp!系统
            lngProgID = rsTmp!程序id
            rsTmp.Close
            
            strSQL = "Update zlReports Set 发布时间=NULL, 是否停用=NULL Where 程序ID Is Null And ID=" & lngID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRPTPuts Where 报表ID=" & lngID & " And 系统=" & lngSys & " And 程序ID=" & lngProgID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRoleGrant Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
        
        Else
            '还有多个发布位置，选择性取消
            Set rsTmp = GetModuleTreeMenu(lngID)
            If rsTmp Is Nothing Then
                MsgBox "读取模块菜单体系时遇到意外错误，报表“取消发布”中断！", vbInformation, App.Title
                Exit Sub
            End If
            
            Load frmSelTree
            frmSelTree.Caption = "取消发布 - 模块菜单位置"
            With frmSelTree.tvw
                .Nodes.Clear
                Do While rsTmp.EOF = False
                    If IsNull(rsTmp!上级ID) Then
                        Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!标题)
                    Else
                        Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, rsTmp!标题)
                    End If
                    If Left(rsTmp!id, 1) = "S" Then     'System
                        objNode.Image = "Root"
                        objNode.Tag = "请选择要取消发布的报表。"
                    ElseIf Left(rsTmp!id, 1) = "T" Then 'MenuTree
                        objNode.Image = "Path"
                        objNode.Tag = "请选择要取消发布的报表。"
                    ElseIf Left(rsTmp!id, 1) = "M" Then 'Module
                        objNode.Image = "App"
                        objNode.Tag = "请选择要取消发布的报表。"
                    ElseIf Left(rsTmp!id, 1) = "R" Then 'Report
                        objNode.Image = "ReportNode"
                        objNode.ForeColor = vbBlue
                    End If
                    objNode.Expanded = True
                    
                    '标记有发布报表的上级
                    If Left(rsTmp!id, 1) = "R" Then
                        objNode.SelectedImage = objNode.Image
                        If objNode.Parent.SelectedImage = "" Then
                            Do While Not objNode.Parent Is Nothing
                                Set objNode = objNode.Parent
                                objNode.SelectedImage = objNode.Image
                            Loop
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                
                '删除无发布报表的路径
                For k = .Nodes.count To 1 Step -1
                    If .Nodes(k).SelectedImage = "" Then
                        .Nodes.Remove k
                    End If
                Next
                
                If .Nodes.count > 0 Then .Nodes(1).Selected = True
            End With
            
            frmSelTree.Show vbModal, Me
            If Not gblnOK Then
                Exit Sub
            End If
            rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
            If rsTmp.EOF Then
                MsgBox "获取模块菜单数据异常，报表“取消发布”中断！", vbInformation, App.Title
                Exit Sub
            End If
            
            lngSys = rsTmp!系统
            lngProgID = rsTmp!程序id
            Unload frmSelTree
            
            strSQL = "Delete From zlRPTPuts Where 报表ID=" & lngID & " And 系统=" & lngSys & " And 程序ID=" & lngProgID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRoleGrant Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
        End If
        
        '执行DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1
        
makContinue:
    Next
    
    Set grsReport = Nothing '清除缓存
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Find(ByVal strText As String, Optional ByVal blnFirst As Boolean = False)
'功能：查找匹配的项，并定位
'参数：
'  strText：待查找匹配的文本
'  blnFirst：True首次查找；False再次查找

    Dim lngRow As Long, l As Long
    Dim arrOrder As Variant
    Dim i As Integer, j As Integer, intIndex As Integer
    Dim blnStart As Boolean
    Dim colOrder As Collection
    Dim strName As String
    
    If Trim(strText) = "" Then Exit Sub
    
    '准备待查找的对象和行号
    If blnFirst Or mobjFindPreviors Is Nothing Then
        '首次
        lngRow = 0
        blnFirst = True
    Else
        '再次
        lngRow = mlngFindRow
    End If
    
    Set colOrder = New Collection
    If tbcRPT.Selected.Index = Val("0-报表页面") Then
        If blnFirst Then
            strName = "VSFREPORT"
        Else
            strName = UCase(mobjFindPreviors.name)
        End If
        '载入查找的对象
        colOrder.Add Me.vsfReport
        colOrder.Add Me.rptClass
    Else
        If blnFirst Then
            strName = "VSFGROUP"
        Else
            strName = UCase(mobjFindPreviors.name)
        End If
        '载入查找的对象
        colOrder.Add Me.vsfGroup
        colOrder.Add Me.vsfGroupDetail
        colOrder.Add Me.rptClass
    End If
    
    '查找
    blnStart = False
    For i = 1 To colOrder.count
        With colOrder(i)
            If .Visible = False Then GoTo makContinue
            
            If UCase(.name) = strName Or blnStart = True Then
                '开始查找
                If UCase(.name) = "RPTCLASS" Then
                    '报表分类
                    For l = lngRow + IIF(blnStart, 0, 1) To .Rows.count - 1
                        For j = 0 To .Columns.count - 1
                            If UCase(.Rows(l).Record(j).Value) Like "*" & UCase(strText) & "*" _
                                And .Columns(j).Visible Then
                                '找到
                                .Rows(l).Selected = True
                                .SetFocus
                                lngRow = l
                                GoTo makFinish
                            End If
                        Next
                    Next
                Else
                    '独立报表、报表组、子报表
                    For l = lngRow + 1 To .Rows - 1
                        For j = 0 To .Cols - 1
                            If UCase(.TextMatrix(l, j)) Like "*" & UCase(strText) & "*" _
                                And .ColWidth(j) > 0 Then
                                '找到
                                .TopRow = l
                                .LeftCol = j
                                .Row = l
                                .Col = j
                                .SetFocus
                                lngRow = l
                                GoTo makFinish
                            End If
                        Next
                    Next
                End If
            
                blnStart = True
                lngRow = 0
            End If
        End With
        
makContinue:
    Next
    
    Set mobjFindPreviors = Nothing
    mlngFindRow = 0
    
    If mintFind <= 0 Then
        If MsgBox("未查找到匹配的内容，是否从头开始查找？", vbInformation + vbDefaultButton1 + vbYesNo, App.Title) = vbYes Then
            mintFind = mintFind + 1
            Call Find(strText, True)
        Else
            mintFind = 0
        End If
    Else
        mintFind = 0
    End If
    
    Exit Sub
    
makFinish:
    '缓存信息下次使用
    Set mobjFindPreviors = colOrder(i)
    mlngFindRow = lngRow
    mintFind = 0
End Sub

Private Sub UpdateStatusBar(ByVal objFocus As Object)
'功能：更新状态栏的显示信息
'参数：
'  objFocus：焦点对象

    Dim strMsg As String
    Dim lngID As Long

    With objFocus
        Select Case UCase(objFocus.name)
        Case "VSFGROUP"
            If mblnReportControlFocus Then Exit Sub
            
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strMsg = mdlPublic.FormatString("【[1]】[2]：包含共 [3] 张报表" _
                        , .TextMatrix(.Row, .ColIndex("编号")) _
                        , .TextMatrix(.Row, .ColIndex("组名")) _
                        , vsfGroupDetail.Rows - 1)
            If .TextMatrix(.Row, .ColIndex("发布时间")) <> "" Then
                strMsg = strMsg & "； 发布位置：" & GetMenuPath(lngID, True)
            End If
        Case "RPTCLASS"
            If tbcRPT.Selected.Index = Val("0-报表页面") Then
                strMsg = mdlPublic.FormatString("【[1]】分类下有 [2] 张报表" _
                            , .FocusedRow.Record(mobjClass.GetColIndex("名称")).Value _
                            , vsfReport.Rows - 1)
            Else
                strMsg = mdlPublic.FormatString("【[1]】分类下有 [2] 份报表组" _
                            , .FocusedRow.Record(mobjClass.GetColIndex("名称")).Value _
                            , vsfGroup.Rows - 1)
            End If
        Case Else
            If mblnReportControlFocus Then Exit Sub
            
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strMsg = mdlPublic.FormatString("【[1]】[2]" _
                        , .TextMatrix(.Row, .ColIndex("编号")) _
                        , .TextMatrix(.Row, .ColIndex("名称")))
            If .TextMatrix(.Row, .ColIndex("发布时间")) <> "" Then
                strMsg = strMsg & "； 发布位置：" & GetMenuPath(lngID, False)
            End If
            If .TextMatrix(.Row, .ColIndex("说明")) <> "" Then
                strMsg = strMsg & "； 说明：" & .TextMatrix(.Row, .ColIndex("说明"))
            End If
        End Select
    End With
    
    Me.staMain.Panels(2).Text = strMsg
End Sub

Private Sub StateSwitch(ByVal lngID As Long, Optional ByVal blnEnabled As Boolean = False)
'功能：报表启用、停用的切换
'参数：
'  lngID：菜单ID
'  blnEnabled：True启用；False停用

    Dim lngRow As Long, lngSelRow As Long, lngReportID As Long
    Dim vsfTemp As VSFlexGrid
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim strIDs As String, strRec As String, strNonRec As String, strName As String
    Dim strSQL  As String, strTmp As String
    Dim colSQL As New Collection
 
    If mblnReportControlFocus = False Then
        If GetVsfControl(lngID, blnGroup, vsfTemp, strIDs) = False Then
            MsgBox "请选中独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "请选中独立报表、报表组、子报表！", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    '检查
    strName = IIF(blnGroup, "组名", "名称")
    For lngRow = 1 To vsfTemp.Rows - 1
        If lngSelRow <= 5 Then
            If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("发布时间")) = "" Then
                    If lngSelRow >= 5 Then
                        strNonRec = strNonRec & vbCrLf & "..."
                    Else
                        strNonRec = strNonRec & vbCrLf & CStr(lngSelRow + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex(strName))
                    End If
                Else
                    If lngSelRow >= 5 Then
                        strRec = strRec & vbCrLf & "..."
                    Else
                        strRec = strRec & vbCrLf & CStr(lngSelRow + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex(strName))
                    End If
                End If
                lngSelRow = lngSelRow + 1
            End If
        Else
            Exit For
        End If
    Next
    If strNonRec <> "" Then
        MsgBox "请确保以下报表" & IIF(blnGroup, "组", "") & "已发布！" & strNonRec, vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '处理
    strTmp = IIF(blnEnabled, "启用", "停用")
    strNonRec = IIF(blnGroup, "组", "")
    If MsgBox("你确定要“" & strTmp & "”下列报表" & strNonRec & "吗？" & strRec, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
            lngReportID = Val(vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("ID")))
            If blnGroup Then
                '报表组
                strSQL = "Update zlRPTGroups " & vbCrLf & _
                         "Set 是否停用 = " & IIF(blnEnabled, "Null", "1") & vbCrLf & _
                         "Where Not 发布时间 Is Null And ID = " & lngReportID & " "
            Else
                '报表
                strSQL = "Update zlReports " & vbCrLf & _
                         "Set 是否停用 = " & IIF(blnEnabled, "Null", "1") & vbCrLf & _
                         "Where Not 发布时间 Is Null And ID = " & lngReportID & " "
            End If
            Call AddArray(colSQL, strSQL)
            
            lngSelRow = lngSelRow + 1
        End If
    Next
    
    '执行DML
    gcnOracle.BeginTrans: blnTrans = True
    For lngRow = 1 To colSQL.count
        gcnOracle.Execute colSQL(lngRow)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Screen.MousePointer = vbDefault
    
    '刷新
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
    Call ErrCenter
End Sub

Private Function GetRelationList(ByVal lngReportID As Long) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo hErr
    
    strSQL = _
        "Select a.报表id, b.名称, a.关联报表id, c.名称 关联名称 " & vbNewLine & _
        "From Zlrptrelation A, zlReports B, zlReports C " & vbNewLine & _
        "Where a.报表id = b.Id(+) And a.关联报表id = c.Id(+) And a.报表id = [1] " & vbNewLine & _
        "Union All " & vbNewLine & _
        "Select a.报表id, b.名称, a.关联报表id, c.名称 关联名称 " & vbNewLine & _
        "From Zlrptrelation A, zlReports B, zlReports C " & vbNewLine & _
        "Where a.报表id = b.Id(+) And a.关联报表id = c.Id(+) And a.关联报表id = [1] "
    strSQL = "Select Distinct 报表id, 名称, 关联报表id, 关联名称 From (" & strSQL & ")"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "", lngReportID)
    Do While rsTemp.EOF = False
        If i <= 4 Then
            If rsTemp!报表ID = lngReportID Then
                GetRelationList = GetRelationList & vbCrLf & String(4, " ") & Chr(97 + i) & ") " & rsTemp!关联名称 & "（主动）"
            Else
                GetRelationList = GetRelationList & vbCrLf & String(4, " ") & Chr(97 + i) & ") " & rsTemp!名称 & "（被动）"
            End If
        Else
            GetRelationList = GetRelationList & "..."
            Exit Do
        End If
        
        i = i + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

