VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.1#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiFeeQuery 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "病人费用查询"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   -780
   ClientWidth     =   15045
   Icon            =   "frmPatiFeeQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptPati 
      Height          =   2265
      Left            =   120
      TabIndex        =   16
      Top             =   4980
      Width           =   3015
      _Version        =   589884
      _ExtentX        =   5318
      _ExtentY        =   3995
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7365
      ScaleHeight     =   345
      ScaleWidth      =   3045
      TabIndex        =   26
      Top             =   210
      Width           =   3045
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   270
         Left            =   15
         TabIndex        =   28
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   476
         ShowSortName    =   0   'False
         IDKindStr       =   "床|床号|1|1|0|0|0|0;住|住院号|0|2|0|0|0|0;姓|姓名|1|3|0|0|0|0;医|医保号|0|4|0|0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;CTRL++;CTRL+-;CTRL+P;CTRL+F;CTRL+F3;CTRL+F5;CTRL+A;CTRL+9;CTRL+U"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         TabIndex        =   27
         ToolTipText     =   "查找病人(Ctrl+F)"
         Top             =   30
         Width           =   1155
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrint 
      Height          =   780
      Left            =   5925
      TabIndex        =   22
      Top             =   420
      Visible         =   0   'False
      Width           =   2760
      _cx             =   4868
      _cy             =   1376
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin VB.PictureBox picCondition 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox PicRptPati 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame fraCondition 
      Height          =   4170
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   3270
      Begin zlIDKind.IDKindNew IDKindPati 
         Height          =   255
         Left            =   135
         TabIndex        =   29
         Top             =   2340
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         ShowSortName    =   0   'False
         IDKindStr       =   "姓|姓名|0|0|0|0|0|0;床|床号|0|0|0|0|0|0;住|住院号|1|0|0|0|0|0;医|医保号|1|0|0|0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;CTRL++;CTRL+-;CTRL+P;CTRL+F;CTRL+F3;CTRL+F5;CTRL+A;CTRL+9;CTRL+U"
         MustSelectItems =   "就诊卡"
         BackColor       =   -2147483633
      End
      Begin VB.Frame fra站点 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   90
         TabIndex        =   23
         Top             =   195
         Width           =   3120
         Begin VB.ComboBox cboNode 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Width           =   2085
         End
         Begin VB.Label lbl站点 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "院区(&0)"
            Height          =   180
            Left            =   345
            TabIndex        =   25
            Top             =   60
            Width           =   630
         End
      End
      Begin VB.TextBox txt预交 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1455
         TabIndex        =   12
         Top             =   2745
         Width           =   570
      End
      Begin VB.CheckBox chk仅显未审核病人 
         Caption         =   "仅显未审核病人(&U)"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   3405
         Width           =   1890
      End
      Begin VB.CheckBox chk仅显未结清病人 
         Caption         =   "仅显未结清病人(&M)"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   3105
         Width           =   1890
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "过滤(&S)"
         Height          =   350
         Left            =   1905
         TabIndex        =   15
         Top             =   3660
         Width           =   1100
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2355
         Width           =   2085
      End
      Begin VB.ComboBox cboState 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1110
         TabIndex        =   7
         Top             =   1605
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   255328259
         CurrentDate     =   36257.9999884259
         MinDate         =   30682
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1110
         TabIndex        =   5
         Top             =   1260
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   255328259
         CurrentDate     =   36257
         MinDate         =   30682
      End
      Begin VB.CheckBox chk预交款 
         Caption         =   "预交余额小于       元的病人"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   2790
         Width           =   2745
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2010
         Width           =   2085
      End
      Begin VB.ComboBox cboUnit 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1110
         TabIndex        =   1
         Text            =   "cboUnit"
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&3)"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         Caption         =   "病人状态(&2)"
         Height          =   180
         Left            =   90
         TabIndex        =   2
         Top             =   945
         Width           =   990
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "病人病区(&1)"
         Height          =   180
         Left            =   90
         TabIndex        =   0
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&4)"
         Height          =   180
         Left            =   90
         TabIndex        =   6
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&5)"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   2070
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   1440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeQuery.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeQuery.frx":0464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   9255
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiFeeQuery.frx":0D3E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17348
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "审核"
            Object.ToolTipText     =   "病人是否已审核"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   6180
      Left            =   3360
      TabIndex        =   20
      Top             =   780
      Width           =   9330
      _Version        =   589884
      _ExtentX        =   16457
      _ExtentY        =   10901
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiFeeQuery.frx":15D2
      Left            =   840
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopAudit 
         Caption         =   "审核(&A)"
      End
      Begin VB.Menu mnuPopUnAudit 
         Caption         =   "取消审核(&U)"
      End
      Begin VB.Menu mnuPopBilling 
         Caption         =   "记帐(&B)"
      End
      Begin VB.Menu mnuPopAudit_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDisp 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPopCard 
         Caption         =   "病人信息卡片(&K)"
      End
      Begin VB.Menu mnuPopNotify 
         Caption         =   "打印多张催款单(&N)"
      End
      Begin VB.Menu mnuPopCurr 
         Caption         =   "打印单张催款单(&C)"
      End
   End
End
Attribute VB_Name = "frmPatiFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModul As Long
Private mblnUnload As Boolean
Private mblnHavePara As Boolean
Private WithEvents mclsFeeQuery As clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private mfrmPatiFeeVerfy As frmPatiFeeVerfy

Private mclsAdvices As Object
Private mcolSubForm As Collection
Private mfrmActive As Form

Private mintFindType As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstrPrePati As String
Private mrsDept As ADODB.Recordset
Private Enum mPan
    Condition = 1
    Pati = 2
End Enum
Private mblnSelPatiList As Boolean '是否选中了病人列表
Private Type t_ViewState
    OnePati As Boolean
End Type
Private mlngUnitID As Long
Private mvs As t_ViewState
Private mbln预交 As Boolean '操作员是否有预交模块功能(有预交,则允许在费用查询中缴预交)
Private mstr截止日期 As String

'字段名,宽度,是否允许分组;(宽度,是否允许分组不写表示隐藏数据列)
Private mstrPatiHead As String
'Private Const mstrPatiHead = "病人ID;主页ID;登记时间;状态;病人性质;数据转出;当前科室ID;险类;密码;当前病区ID;就诊卡号;" & _
                           "类别,30,0;姓名,60,0;住院号,60,0;床号,50,0;费别,60,1;性别,40,1;年龄,60,1;入院时间,100,0;出院时间,100,0;" & _
                           "当前科室,80,1;次数,40,1;结清,40,1;医保号,120,0;联系电话,80,0;医疗付款方式,120,1;" & _
                           "审核,45,1;病人类型,100,1;当前病区,80,1"
Private mobjPatient As Object
Private mobjPlugIn As Object
Private mblnNotClick As Boolean
Private mbln缺省读卡 As Boolean
Private mbln启用站点 As Boolean
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrCaption As String
Private mbytFontSize As Byte
Private mintInsure As Integer '险类:31883

Private Sub ReMoveCtrol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移动控件位置
    '编制:刘兴洪
    '日期:2012-06-19 11:29:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    cboUnit.Left = lblUnit.Left + lblUnit.Width + 20
    
    lbl站点.Top = cboNode.Top + (cboNode.Height - lbl站点.Height) \ 2
    cboNode.Left = cboUnit.Left
    cboNode.Width = dtpBegin.Width
    lbl站点.Left = cboNode.Left - 20 - lbl站点.Width
    fra站点.Width = fraCondition.Width - 60
    fra站点.Left = 15
    
    If mbln启用站点 Then
        cboUnit.Top = fra站点.Top + fra站点.Height + IIf(mbytFontSize = 9, 15, 60)
    Else
        cboUnit.Top = fra站点.Top
    End If
    cboUnit.Width = dtpBegin.Width
    lblUnit.Top = cboUnit.Top + (cboUnit.Height - lblUnit.Height) \ 2
    cboUnit.Left = cboNode.Left
    
    cboState.Top = cboUnit.Top + cboUnit.Height + 50
    cboState.Left = cboNode.Left
    cboState.Width = dtpBegin.Width
    lblState.Top = cboState.Top + (cboState.Height - lblState.Height) \ 2
    
    dtpBegin.Top = cboState.Top + cboState.Height + 50
    dtpBegin.Height = cboState.Height: dtpEnd.Height = dtpBegin.Height
    dtpBegin.Left = cboNode.Left
    lblStartDate.Top = dtpBegin.Top + (dtpBegin.Height - lblStartDate.Height) \ 2
    
    dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 50
    dtpEnd.Left = cboNode.Left
    lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
    
    dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 50
    dtpEnd.Left = cboNode.Left
    lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
        
    txt住院号.Top = dtpEnd.Top + dtpEnd.Height + 50
    txt住院号.Height = cboNode.Height
    txt住院号.Left = cboNode.Left
    txt住院号.Width = dtpBegin.Width
    lbl住院号.Top = txt住院号.Top + (txt住院号.Height - lbl住院号.Height) \ 2
    
    txt姓名.Top = txt住院号.Top + txt住院号.Height + 50
    txt姓名.Height = cboNode.Height
    txt姓名.Left = cboNode.Left
    IDKindPati.Top = txt姓名.Top + (txt姓名.Height - IDKindPati.Height) \ 2
    txt姓名.Width = cboState.Width
    IDKindPati.Left = lblUnit.Left + lblUnit.Width - IDKindPati.Width
    txt预交.Top = txt姓名.Top + txt姓名.Height + 50: txt预交.Height = cboNode.Height
    chk预交款.Top = txt预交.Top + (txt预交.Height - chk预交款.Height) \ 2
    txt预交.Left = chk预交款.Left + TextWidth("<预交余额小于 >")
    
    chk仅显未结清病人.Top = txt预交.Top + txt预交.Height + 50
    chk仅显未审核病人.Top = chk仅显未结清病人.Top + chk仅显未结清病人.Height + 50
    cmdSearch.Top = chk仅显未审核病人.Top + chk仅显未审核病人.Height + 50
    
    dkpMain.Panes(1).MinTrackSize.Height = IIf(mbytFontSize = 9, 270, 300)
    dkpMain.Panes(1).MinTrackSize.Width = IIf(mbytFontSize = 9, 225, 295)
    dkpMain.RedrawPanes
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
 
 
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    Call ReMoveCtrol
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置字体大小
    '编制:刘兴洪
    '日期:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Call mclsFeeQuery.SetFontSize(bytSize)
    Call mfrmPatiFeeVerfy.SetFontSize(bytSize)
    Call mfrmActive.SetFontSize(bytSize)
    If Not mclsAdvices Is Nothing Then Call mclsAdvices.SetFontSize(bytSize)
    
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("刘") + 20
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("刘") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    IDKindPati.FontSize = mbytFontSize
    IDKindPati.Refrash
    Call Form_Resize
End Sub

Private Sub cboNode_Click()
    If mblnNotClick Then Exit Sub
    Call LoadUnits
End Sub

Private Sub cboState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cboUnit_Click()
    mlngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    Dim lngID As Long
    
    If cboUnit.ListIndex >= 0 Then Exit Sub
    lngID = mlngUnitID
   zlControl.CboLocate cboUnit, lngID, True
   If cboUnit.ListIndex < 0 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0
End Sub

Private Sub chk仅显未结清病人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub chk仅显未审核病人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
   Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    Dim strKind As String
    '问题:42946
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        'If mobjICCard Is Nothing Then Exit Sub
        'txtFind.Text = mobjICCard.Read_Card()
        txtFind.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtFind.IMEMode = 0
        
        If txtFind.Text = "" Then Exit Sub
        ExecFindPati objCard, True
        Exit Sub
    End If
   txtFind.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtFind.IMEMode = 0
    If lng卡类别ID <= 0 Then
         txtFind.PasswordChar = ""
         '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
         txtFind.IMEMode = 0
         Exit Sub
    End If
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtFind.Text = strOutCardNO
    If txtFind.Text = "" Then
        txtFind.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtFind.IMEMode = 0
        Exit Sub
    End If
    
    ExecFindPati objCard, True
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtFind.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtFind.Locked Then Exit Sub
    txtFind.Text = objPatiInfor.卡号
    If txtFind.Text = "" Then
        txtFind.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtFind.IMEMode = 0
        Exit Sub
    End If
    ExecFindPati objCard, True
End Sub

Private Sub IDKindPati_Click(objCard As zlIDKind.Card)
   Dim strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    Dim strKind As String, intFindType As Integer
    
    '问题:42946
    strKind = objCard.名称
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        txt姓名.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txt姓名.IMEMode = 0
        If txt姓名.Text = "" Then Exit Sub
        ExecFindPati objCard, , True
        Exit Sub
    End If
    txt姓名.PasswordChar = IIf(IDKindPati.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txt姓名.IMEMode = 0
    
    If objCard.接口序号 <= 0 Then
        txt姓名.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txt姓名.IMEMode = 0
        Exit Sub
    End If
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, objCard.接口序号, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt姓名.Text = strOutCardNO
    If txt姓名.Text = "" Then
        txt姓名.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txt姓名.IMEMode = 0
        Exit Sub
    End If
    
    Call LoadPatients(objCard, True)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
End Sub

Private Sub IDKindPati_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txt姓名.IMEMode = 0
End Sub
Private Sub IDKindPati_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txt姓名.Locked Then Exit Sub
    txt姓名.Text = objPatiInfor.卡号
    If txt姓名.Text = "" Then
        txt姓名.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txt姓名.IMEMode = 0
        Exit Sub
    End If
    
    Call LoadPatients(objCard, True)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
End Sub

Private Sub mclsFeeQuery_RequestRefresh()
'功能：子窗体要求刷新
    Call LoadPatients(IDKindPati.GetCurCard)
End Sub

Private Sub mclsFeeQuery_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.sta.Panels(2).Text = Text
End Sub

Private Sub cboState_Click()
    Dim objControl As CommandBarButton
    
    dtpBegin.Enabled = cboState.Text = "出院病人" Or cboState.Text = "所有病人"
    dtpEnd.Enabled = dtpBegin.Enabled
        
    rptPati.Columns(GetRptColumn(rptPati, "出院时间")).Visible = cboState.Text <> "在院病人"
    
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboUnit.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    Dim strRootCaption As String
    strRootCaption = ""
    If InStr(mstrPrivs, ";所有病区;") > 0 Then strRootCaption = "所有病区"
    If cboNode.ListCount > 0 Then
        mrsDept.Filter = "站点=" & cboNode.ItemData(cboNode.ListIndex)
    End If
    
    If zlSelectDept(Me, mlngModul, cboUnit, mrsDept, cboUnit.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
     
End Sub


Private Sub ExecPrintMultiBill()
    Dim i As Long
    Dim rptr As ReportRecord
    '--27894
    If cboUnit.ItemData(cboUnit.ListIndex) = 0 Then
        MsgBox "不允许一次打印所有病区病人的通知单，" & vbCr & "请选择具体的病区后再试！", vbInformation, gstrSysName
        Exit Sub
    End If
    '问题:调整如下:34770
    If frmPatiPressMoney.zlPatiPressMoney(Me, mlngModul, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex), cboUnit.Text) = False Then Exit Sub
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objRow As ReportRow, objControl As CommandBarControl
    Dim i As Long, blnSelect As Boolean
     
    Select Case Control.ID
        Case conMenu_File_PrintMultiBill
            ExecPrintMultiBill
        Case conMenu_File_Preview_Pati  '打印预览病人列表
            Call zlRptPrint(2)
        Case conMenu_File_Print_Pati   '打印病人列表
            Call zlRptPrint(1)
        Case conMenu_File_Excel_Pati   '病人列表输出到Excel
            Call zlRptPrint(3)
        Case conMenu_Edit_PreBalanceAll
            Call ExecPreBalanceAll
        Case conMenu_Edit_Balance   '结帐
            Call ExecBalance
        Case conMenu_Edit_PrePayMoney '缴预交
            Call ExecPrePayMoney
        Case conMenu_Manage_Change_InsureSel
            Call ModeInsurePatiDisease  '31883
        Case conMenu_Edit_FeeAudit  '审核
            Call ExecAuditingAndCancelAudit(1)
        Case conMenu_Edit_OverFeeAudit '完成审核
            Call ExecAuditingAndCancelAudit(2)
        Case conMenu_Edit_FeeUnAudit   '取消审核
            Call ExecAuditingAndCancelAudit(0)
            
        Case conMenu_View_OnePati   '多次住院只显示一个病人
            Control.Checked = Not Control.Checked: mvs.OnePati = Control.Checked
            Call LoadPatients(IDKindPati.GetCurCard)
        Case conMenu_View_GroupCol * 10 + 1 To conMenu_View_GroupCol * 10 + UBound(Split(mstrPatiHead, ";"))
            Control.Checked = Not Control.Checked
            If Control.Checked Then
                i = GetRptColumn(rptPati, Control.Caption)
                rptPati.GroupsOrder.Add rptPati.Columns(i)
                rptPati.Columns(i).Visible = False
                rptPati.Populate
            Else
                Call rptRemoveGroupsItem(rptPati, Control.Caption)
            End If
'
'        Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + mcllBrushCard.Count  '查找方式
'            mintFindType = Val(Control.Parameter)
'            cbsMain.RecalcLayout
'            txtFind.Text = ""
'            txtFind.SetFocus
'            Call InitCardType
'        Case conMenu_View_Filter * 100# + 1 To conMenu_View_Filter * 100# + mcllBrushCard.Count
'            '弹出菜单的显示
'            '刘兴洪:24913
'            IDKindPati.Tag = Val(Control.Parameter)
'            IDKindPati.Caption = Replace(Split(Control.Caption, "(")(0) & "↓(&6)", " ", "")
'            If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
'            zlDatabase.setPara "病人过滤类别", Val(IDKindPati.Tag), glngSys, mlngModul, mblnHavePara
'            Call InitSearchType
        Case conMenu_View_Find '查找
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus '有时需要定位一下
                If txtFind.Text <> "" Then
                    Call ExecFindPati(IDKIND.GetCurCard)
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext '查找下一个
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call ExecFindPati(IDKIND.GetCurCard, True)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                cbsMain(i).Visible = Not cbsMain(i).Visible
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            sta.Visible = Not sta.Visible
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_FontSize_S    '小字体
            Call SetFontSize(0)
        Case conMenu_View_FontSize_L    '大字体
            Call SetFontSize(1)
        Case conMenu_View_Expend_CurCollapse '折叠当前组
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    rptPati.SelectedRows(0).Expanded = False
                ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                        rptPati.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPati_SelectionChanged
        Case conMenu_View_Expend_CurExpend '展开当前组
            If rptPati.SelectedRows.Count > 0 Then
                rptPati.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse '折叠所有组
            For Each objRow In rptPati.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPati_SelectionChanged
        Case conMenu_View_Expend_AllExpend '展开所有组
            For Each objRow In rptPati.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(hWnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(hWnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(hWnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, hWnd, Name, Int((glngSys) / 100))
        Case conMenu_File_SchemeSet '报警方案设置
             Call zlSchemeSet
        Case conMenu_File_Exit '退出
            Unload Me
        Case Else
            Select Case Me.tbcSub.Selected.Tag
            Case "费用", "医嘱"
                 Call mclsFeeQuery.zlExecuteCommandBars(Control)
            End Select
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strPrivsTemp As String
    
    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub
    blnVisible = True
    
    Select Case Control.ID
        '76511:刘尔旋,2014-08-12,预结所有病人的权限判断缺失
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(";" & mstrPrivs, ";预结所有病人;") > 0
            Control.Category = "已判断"
        Case conMenu_Edit_Balance
            strPrivsTemp = GetInsidePrivs(Enum_Inside_Program.p病人结帐)
            blnVisible = InStr(strPrivsTemp, "住院费用结帐") > 0
            Control.Category = "已判断"
        Case conMenu_Edit_PrePayMoney '预交
            strPrivsTemp = ";" & GetInsidePrivs(Enum_Inside_Program.p预交款) & ";"
            blnVisible = InStr(strPrivsTemp, ";预交收款;") > 0
            Control.Category = "已判断"
        Case conMenu_Manage_Change_InsureSel
            blnVisible = InStr(";" & mstrPrivs, ";病种选择;") > 0
            Control.Category = "已判断"
        Case conMenu_Edit_FeeAudit
            blnVisible = InStr(";" & mstrPrivs, ";审核病人;") > 0
            Control.Category = "已判断"
        Case conMenu_Edit_OverFeeAudit  '完成审核
            blnVisible = InStr(";" & mstrPrivs, ";审核病人;") > 0
            Control.Category = "已判断"
        Case conMenu_Edit_FeeUnAudit
            blnVisible = InStr(";" & mstrPrivs, ";取消审核病人;") > 0
            Control.Category = "已判断"
    End Select
    Control.Visible = blnVisible
    Control.Enabled = blnVisible    '51135 :必须设置Enabled属性,不然快键处理不了
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub


Private Sub ExecPreBalanceAll()
    Dim rptr As ReportRecord, rsTmp As ADODB.Recordset
    Dim arrInfo() As Variant, str结算费用 As String, i As Integer
    Dim lng病人ID As Long, int险类 As Integer, blnDateMoved As Boolean
    Dim str医保号 As String, str密码 As String, str姓名 As String, dat登记时间 As Date
            
    arrInfo = Array()
    For Each rptr In rptPati.Records
        If Trim(rptr(GetRptRsColumn("医保号")).Value) <> "" And Trim(rptr(GetRptRsColumn("出院时间")).Value) = "" Then
            ReDim Preserve arrInfo(UBound(arrInfo) + 1)
             '姓名,病人ID,险类,医保号,密码
            arrInfo(UBound(arrInfo)) = rptr(GetRptRsColumn("姓名")).Value & "|" & Val(rptr(GetRptRsColumn("病人ID")).Value) & "|" & _
                rptr(GetRptRsColumn("险类")).Value & "|" & rptr(GetRptRsColumn("医保号")).Value & "|" & rptr(GetRptRsColumn("密码")).Value & "|" & rptr(GetRptRsColumn("登记时间")).Value
        End If
    Next
    
    If UBound(arrInfo) = -1 Then
        MsgBox "当前病人清单中没有发现当前险类的在院医保病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("该操作将对当前病人清单中的所有在院医保病人进行预结算," & _
        vbCrLf & "这可能会花费较长的时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    For i = 0 To UBound(arrInfo)
        str姓名 = Split(arrInfo(i), "|")(0)
        lng病人ID = Val(Split(arrInfo(i), "|")(1))
        int险类 = Val(Split(arrInfo(i), "|")(2))
        str医保号 = Split(arrInfo(i), "|")(3)
        str密码 = Split(arrInfo(i), "|")(4)
        dat登记时间 = CDate(Split(arrInfo(i), "|")(5))
        
        If Not gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, int险类) Then
            blnDateMoved = zlDatabase.DateMoved(dat登记时间, , , Caption)
            
            Call zlCommFun.ShowFlash("正在处理医保病人""" & str姓名 & """ ...", Me)
            Refresh
            
            Set rsTmp = GetVBalance(1, "住院费用结帐", int险类, lng病人ID, , , , , blnDateMoved)
            If Not rsTmp Is Nothing Then
                If Not rsTmp.RecordCount = 0 Then
                    str结算费用 = gclsInsure.WipeoffMoney(rsTmp, lng病人ID, str医保号, "0", int险类, "|0") '当成中途结算
                End If
            End If
        End If
    Next
    sta.Panels(3).Text = "预结算成功!"
    Call zlCommFun.StopFlash
    
    mstrPrePati = ""
    Call rptPati_SelectionChanged
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnSelect As Boolean, lngColTmp As Long, blnEnabled As Boolean, blnQueryFee As Boolean
    Dim strTemp As String
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    If rptPati.SelectedRows.Count > 0 Then blnSelect = Not rptPati.SelectedRows(0).GroupRow
    
    Select Case Control.ID
        '文件
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintDayDetail
            Control.Enabled = rptPati.Records.Count > 0
    
        '编辑
        Case conMenu_Edit_PreBalanceAll
            Control.Enabled = rptPati.Records.Count > 0 And cboState.Text <> "出院病人"
        Case conMenu_Edit_Balance
            Control.Enabled = blnSelect
        Case conMenu_Edit_PrePayMoney '预交款
            Control.Enabled = blnSelect
        Case conMenu_Manage_Change_InsureSel
            '31883
             Control.Enabled = blnSelect And mintInsure <> 0
        Case conMenu_Edit_FeeAudit
            '病人审核
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("审核状态")).Value)
                Control.Enabled = (strTemp = "")
            End If
        Case conMenu_Edit_OverFeeAudit '完成审核
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("审核状态")).Value)
                Control.Enabled = strTemp = "开始"
            End If
        Case conMenu_Edit_FeeUnAudit
            '取消审核
            Control.Enabled = blnSelect
            If blnSelect Then
                strTemp = Trim(rptPati.SelectedRows(0).Record(GetRptRsColumn("审核状态")).Value)
                Control.Enabled = strTemp <> ""
            End If
       '查看
        Case conMenu_View_OnePati
            Control.Enabled = cboState.Text = "出院病人" Or cboState.Text = "所有病人"
            Control.Checked = mvs.OnePati
        
        Case conMenu_View_GroupCol * 10 + 1 To conMenu_View_GroupCol * 10 + UBound(Split(mstrPatiHead, ";")) '分组列
            Control.Enabled = rptPati.Rows.Count > 1
            
            Control.Checked = False
            For lngColTmp = 0 To rptPati.GroupsOrder.Count - 1
                If Control.Caption = rptPati.GroupsOrder(lngColTmp).Caption Then
                    Control.Checked = True
                    Exit Sub
                End If
            Next
        
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.sta.Visible
        Case conMenu_View_Expend_CurExpend '展开当前组
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptPati.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse '折叠当前组
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.SelectedRows(0).GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).Expanded
                ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend '折叠/展开组
            Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
        Case conMenu_View_FindType '查找方式
        Case conMenu_View_FindNext
            Control.Enabled = rptPati.Records.Count > 1
        Case conMenu_File_SchemeSet  '报警方案设置:35386
             Control.Visible = InStr(1, mstrPrivs, ";报警方案设置;") > 0
        Case conMenu_View_FontSize_S         '小字体
             Control.Checked = mbytFontSize = 9
        Case conMenu_View_FontSize_L    '大字体
             Control.Checked = mbytFontSize <> 9
        Case Else
            Select Case tbcSub.Selected.Tag
            Case "费用", "医嘱"
                Call mclsFeeQuery.zlUpdateCommandBars(Control)
            End Select
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim i As Long, strKey As String, objControl As CommandBarControl
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "费用"
           Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
       End Select
    End Select
End Sub

Private Sub cmdSearch_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间不能大于结束时间!", vbInformation, gstrSysName
        If dtpBegin.Visible And dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Sub
    End If
    mlng病人ID = 0
    Call LoadPatients(IDKindPati.GetCurCard)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
 End Sub


Private Sub ExecBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结帐操作
    '编制:刘兴洪
    '日期:2015-02-05 12:00:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    Dim bln门诊留观病人 As Boolean
    
    bln门诊留观病人 = ZlIsOutpatientObserve(mlng病人ID, mlng主页ID)
    strPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p病人结帐) & ";"
    If Val(zlDatabase.GetPara("结帐界面风格", glngSys, 1137, "1")) = 0 Then
        If frmPatiBalanceTraditional.ShowMe(Me, _
            IIf(bln门诊留观病人, g_Ed_门诊结帐, g_Ed_住院结帐), strPrivs, mlng病人ID, CStr(mlng主页ID)) = False Then Exit Sub
    Else
        If frmPatiBalanceSplit.ShowMe(Me, _
            IIf(bln门诊留观病人, g_Ed_门诊结帐, g_Ed_住院结帐), strPrivs, mlng病人ID, CStr(mlng主页ID)) = False Then Exit Sub
    End If
    If MsgBox("当前内容可能已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        mstrPrePati = ""
        Call RefreshData
    End If
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPan.Pati
        Item.Handle = PicRptPati.hWnd
    Case mPan.Condition
        Item.Handle = picCondition.hWnd
    End Select
End Sub


'111515:李南春，2017/8/14，界面大小调整就忽略错误
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Call cbsMain_Resize
    txt住院号.Width = dtpBegin.Width
    txt姓名.Width = dtpBegin.Width
    cboNode.Width = dtpBegin.Width
    cboUnit.Width = dtpBegin.Width
    cboState.Width = dtpBegin.Width
End Sub

Private Sub picCondition_Resize()
    Err = 0: On Error Resume Next
    With fraCondition
        .Top = 0
        .Left = 0
        .Width = picCondition.ScaleWidth
        .Height = picCondition.ScaleHeight
        fra站点.Width = .Width - 60
    End With
End Sub

 

Private Sub picFind_Resize()
    Err = 0: On Error Resume Next
    With picFind
        IDKIND.Left = .ScaleLeft
        txtFind.Left = IDKIND.Left + IDKIND.Width
        txtFind.Width = .ScaleWidth - txtFind.Left
        'txtFind.Width = .ScaleWidth - txtFind.Left - IIf(cmdReadCard.Visible, cmdReadCard.Width + 50, 0)
    End With
End Sub

Private Sub PicRptPati_Resize()
    Err = 0: On Error Resume Next
    With rptPati
        .Top = PicRptPati.ScaleTop
        .Left = PicRptPati.ScaleLeft
        .Width = PicRptPati.ScaleWidth
        .Height = PicRptPati.ScaleHeight
    End With
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Dim i As Long, j As Long, arrTmp As Variant
    If Button = 2 Then
        Set objHitTest = rptPati.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestHeader Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_GroupCol, True, True)
            Set objCommandBar = objPopup.CommandBar
        ElseIf objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.ActiveMenuBar.FindControl(xtpControlButtonPopup, conMenu_View_Expend, True, True)
                Set objCommandBar = objPopup.CommandBar
            Else
                Set objCommandBar = cbsMain.Add("PopupPati", xtpBarPopup)
                With objCommandBar.Controls
                    Set objControl = .Add(xtpControlButton, conMenu_File_Preview_Pati, "打印预览病人列表(&T)")
                    objControl.BeginGroup = True:        objControl.IconId = conMenu_File_Preview
                    Set objControl = .Add(xtpControlButton, conMenu_File_Print_Pati, "打印病人列表(&O)")
                    objControl.IconId = conMenu_File_Print
                    Set objControl = .Add(xtpControlButton, conMenu_File_Excel_Pati, "病人列表输出到Excel(&E)")
                    objControl.IconId = conMenu_File_Excel
                    
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "病种选择(&Z)")
                    objControl.BeginGroup = True
                    
                    .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结当前病人(&W)").BeginGroup = True
                    
                    .Add xtpControlButton, conMenu_Edit_Balance, "结帐(&B)"
                    .Add xtpControlButton, conMenu_Edit_Billing, "记帐(&C)"
                    
                    .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt病人审核方式 = 1, "开始审核(&A)", "审核(&A)")).BeginGroup = True
                    .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "完成审核(&O)").IconId = 252
                    .Add xtpControlButton, conMenu_Edit_FeeUnAudit, "取消审核(&U)"
                    .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "缴预交(&P)").IconId = 3816
                
                    '加入病人信息,一日清单,催款单
                    .Add(xtpControlButton, conMenu_View_PatInfor, "病人详细信息(&K)").BeginGroup = True
                    .Add xtpControlButton, conMenu_File_PrintSingleBill, "打印单张催款单(&C)…"
                    .Add xtpControlButton, conMenu_File_PrintDayDetail, "打印一日清单(&D)…"
                End With
            End If
        End If
        If Not objCommandBar Is Nothing Then objCommandBar.ShowPopup
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not rptPati.SelectedRows(0).GroupRow Then Call ShowPatiCard
End Sub


Private Sub ShowPatiCard()
    If mlng病人ID <> 0 Then
        frmDegreeCard.mlng病人ID = mlng病人ID
        frmDegreeCard.mlng主页ID = mlng主页ID
        frmDegreeCard.Show 1, Me
    End If
End Sub

Private Sub rptPati_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mblnSelPatiList = True
End Sub

Private Sub rptPati_SelectionChanged()
    Dim strTmp As String, lng住院天数 As Long
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then
            mlng病人ID = Val(.Record(GetRptRsColumn("病人ID")).Value)
            mlng主页ID = Val(.Record(GetRptRsColumn("主页ID")).Value)
            mintInsure = Val(.Record(GetRptRsColumn("险类")).Value)
            If .Record(GetRptRsColumn("出院时间")).Value = "" Then
                lng住院天数 = DateDiff("d", CDate(Format(.Record(GetRptRsColumn("入院时间")).Value, "yyyy-mm-dd")), CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd")))
            Else
                lng住院天数 = DateDiff("d", CDate(Format(.Record(GetRptRsColumn("入院时间")).Value, "yyyy-mm-dd")), CDate(Format(.Record(GetRptRsColumn("出院时间")).Value, "yyyy-mm-dd")))
            End If
            If lng住院天数 = 0 Then lng住院天数 = 1
            strTmp = "病人ID:" & mlng病人ID & ",姓名:" & .Record(GetRptRsColumn("姓名")).Value & _
                ",第 " & mlng主页ID & " 次住院,住院天数:" & lng住院天数 & "天,入院时间:" & .Record(GetRptRsColumn("入院时间")).Value & ",出院时间:" & _
                .Record(GetRptRsColumn("出院时间")).Value
        Else
            mlng病人ID = 0
            mlng主页ID = 0
            mintInsure = 0
            strTmp = ""
        End If
    End With
    
    If mstrPrePati = mlng病人ID & ":" & mlng主页ID Then Exit Sub
    mstrPrePati = mlng病人ID & ":" & mlng主页ID
        
    sta.Panels(2).Text = strTmp
    Call tbcSub_SelectedChanged(tbcSub.Selected)
    If rptPati.Visible Then rptPati.SetFocus
End Sub


Private Sub RefreshData()
    If rptPati.Records.Count = 0 Then
        Call LoadPatients(IDKindPati.GetCurCard)
    Else
        Call rptPati_SelectionChanged
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "病人颜色" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Visible Then Exit Sub
    
    Call SubWinDefCommandBar(Item)
    Call SubWinRefreshData(Item)
End Sub

Private Sub txtFind_Change()
    txtFind.Tag = ""
End Sub
Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, intLen As Integer
    Select Case IDKIND.GetCurCard.名称
    Case "姓名"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        blnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, IDKIND.ShowPassText)
        intLen = IDKIND.GetCardNoLen
    Case "床号"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case "住院号"
        '63494:刘尔旋,2013-10-25 ,住院号不能定位病人列表的问题
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "医保号"
    Case Else
            If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
           If IDKIND.GetCurCard.接口序号 > 0 Then
                blnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, IDKIND.ShowPassText)
                intLen = IDKIND.GetCardNoLen
            End If
     End Select
     
    '刷卡完毕或输入号码后回车
    If (blnCard And Len(txtFind.Text) = intLen - 1 Or KeyAscii = 13) And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtFind.Text = txtFind.Text & Chr(KeyAscii)
            txtFind.SelStart = Len(txtFind.Text)
        End If
        KeyAscii = 0:
        Call ExecFindPati(IDKIND.GetCurCard, , blnCard)
        zlControl.TxtSelAll txtFind
   End If
End Sub

Private Sub txt姓名_Change()
    txt姓名.Tag = ""
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
    Call OpenIme(gstrIme)
End Sub
 
Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, intTYPE As Integer
    Dim strKind As String, intLen As Integer
    'Switch(intTmp = 0, "姓名↓(&6)", intTmp = 1, "就诊卡↓(&6)", intTmp = 2, "床号↓(&6)", intTmp = 3, "医保号↓(&6)", True, "姓名↓(&6)")
    intTYPE = Val(IDKindPati.Tag)
    
    txt姓名.PasswordChar = IIf(IDKindPati.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txt姓名.IMEMode = 0
    strKind = IDKindPati.GetCurCard.名称    '56866
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        KeyAscii = Asc(Chr(KeyAscii))
        blnCard = zlCommFun.InputIsCard(txt姓名, KeyAscii, IDKindPati.ShowPassText)
        intLen = IDKindPati.GetCardNoLen
    Case "床号"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case "住院号"
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "医保号"
    Case Else
            If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
           If IDKindPati.GetCurCard.接口序号 > 0 Then
                blnCard = zlCommFun.InputIsCard(txt姓名, KeyAscii, IDKindPati.ShowPassText)
                intLen = IDKindPati.GetCardNoLen
            End If
     End Select
     
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txt姓名.Text) = intLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txt姓名.Text = txt姓名.Text & Chr(KeyAscii)
            txt姓名.SelStart = Len(txt姓名.Text)
        End If
        KeyAscii = 0
        If Trim(txt姓名.Text) <> "" Then
            If blnCard Then
                Call LoadPatients(IDKindPati.GetCurCard, blnCard)
                Call tbcSub_SelectedChanged(tbcSub.Selected)
            Else
                Call cmdSearch_Click
                zlControl.TxtSelAll txt姓名
            End If
        Else
            zlCommFun.PressKey vbKeyTab
        End If
   End If
     
End Sub

Private Sub txt姓名_Validate(Cancel As Boolean)
    txt姓名.Text = Trim(txt姓名.Text)
    Call OpenIme
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    '24547
    If Trim(txt住院号.Text) <> "" Then
        Call cmdSearch_Click
        zlControl.TxtSelAll txt住院号
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt住院号_Validate(Cancel As Boolean)
    txt住院号.Text = Trim(txt住院号.Text)
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2012-05-21 14:32:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '字段名,宽度,是否允许分组;(宽度,是否允许分组不写表示隐藏数据列)
      mstrPatiHead = "" & _
    "病人ID;主页ID;登记时间;状态;病人性质;数据转出;当前科室ID;险类;密码;当前病区ID;就诊卡号;" & _
    "类别,30,0;姓名,60,0;住院号,60,0;床号,50,0;费别,60,1;性别,40,1;年龄,60,1;入院时间,100,0;出院时间,100,0;" & _
    "当前科室,80,1;次数,40,1;结清,40,1;医保号,120,0;联系电话,80,0;医疗付款方式,120,1"
    If gTy_System_Para.byt病人审核方式 = 0 Then
        mstrPatiHead = mstrPatiHead & ";审核状态;审核人,45,1;病人类型,100,1;当前病区,80,1"
    Else
        mstrPatiHead = mstrPatiHead & ";审核状态,100,1;审核人,45,1;病人类型,100,1;当前病区,80,1"
    End If
End Sub
Private Sub Form_Load()
    Dim lngTmp As Long, strTmp As String, DatTmp As Date, blnAdviceQuery As Boolean
    Dim objPan As Pane, strValue As String, objCondition As Pane
    mbytFontSize = IIf(Val(zlDatabase.GetPara("显示字体大小", glngSys, glngModul)) = 0, 9, 12)
    Call InitData    ' 初始化必要数据
    mblnSelPatiList = False
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    mblnHavePara = InStr(1, mstrPrivs, ";参数设置;") > 0
    Call InitMenus
    
    mstr截止日期 = ""
    blnAdviceQuery = GetInsidePrivs(p住院医嘱下达) <> ""
    If InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "住院记帐") > 0 Then Call InitLocPar(Enum_Inside_Program.p住院记帐)
    If InStr(GetInsidePrivs(Enum_Inside_Program.p病人结帐), "住院费用结帐") > 0 Then Call InitLocPar(Enum_Inside_Program.p病人结帐)
    If Val(zlDatabase.GetPara("使用个性化风格", , , True)) = 1 Then
        IDKindPati.IDKIND = IIf(Val(zlDatabase.GetPara("病人过滤类别", glngSys, mlngModul, "1")) = 0, 1, Val(zlDatabase.GetPara("病人过滤类别", glngSys, mlngModul, "1")))
        GetRegInFor g私有模块, Me.Name, "IDKind", strValue
        IDKIND.IDKIND = IIf(Val(strValue) = 0, 1, Val(strValue))
    End If
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Set fraCondition.Container = picCondition
    fraCondition.Top = 0
    fraCondition.Left = 0
    '菜单初始
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    '窗体布局面板初始
    '-----------------------------------------------------
    dkpMain.SetCommandBars Me.cbsMain
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = False
    
    Set objCondition = dkpMain.CreatePane(mPan.Condition, 220, 200, DockLeftOf, Nothing)
    objCondition.Title = "查询条件": objCondition.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPan = dkpMain.CreatePane(mPan.Pati, 200, 500, DockBottomOf, objCondition)
    objPan.Title = "病人列表": objPan.Options = PaneNoCloseable Or PaneNoFloatable
        
     'TabControl
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    Set mclsFeeQuery = New clsFeeQuery
    If blnAdviceQuery Then Set mclsAdvices = CreateObject("zlCISKernel.clsDockInAdvices")
    
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_费用"
    If blnAdviceQuery Then mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    
    Set mfrmPatiFeeVerfy = New frmPatiFeeVerfy
    Load mfrmPatiFeeVerfy
    mcolSubForm.Add mfrmPatiFeeVerfy, "_医嘱与费用"
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "费用查询", mcolSubForm("_费用").hWnd, 0).Tag = "费用"
        If blnAdviceQuery Then .InsertItem(1, "医嘱查询", mcolSubForm("_医嘱").hWnd, 0).Tag = "医嘱"
        .InsertItem(2, "医嘱与费用", mcolSubForm("_医嘱与费用").hWnd, 0).Tag = "医嘱与费用"
        
        Call SubWinDefCommandBar(.Selected)   '初始刷新定义一次菜单及按钮
        Call SubWinRefreshData(.Selected)
    End With
                    
    Call InitRPTPati
    
    PicRptPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '病人病区
    If Not InitUnits Then mblnUnload = True: Exit Sub

    If InStr(";" & mstrPrivs, ";出院病人查询;") = 0 Then
        strTmp = "在院病人,预出院病人"
    Else
        strTmp = "在院病人,预出院病人,出院病人,所有病人"
    End If
    Call CboAddByStrings(cboState, strTmp, False)
    
    strTmp = zlDatabase.GetPara("病人状态", glngSys, mlngModul, "在院病人")
    Call zlControl.CboLocate(cboState, strTmp)
    
    
    lngTmp = Val(zlDatabase.GetPara("间隔天数", glngSys, mlngModul, -1))
    If lngTmp > 100 Then lngTmp = 7 '只间隔7天
    DatTmp = zlDatabase.Currentdate()
    '42849
    dtpEnd.Value = Format(DatTmp, "yyyy-mm-dd 23:59:59")
    If lngTmp = -1 Then
        dtpBegin.Value = CDate(Format(DateAdd("m", -1, DatTmp), "yyyy-mm-dd") & " 00:00:00")
    Else
        dtpBegin.Value = CDate(Format(DateAdd("d", -lngTmp, DatTmp), "yyyy-mm-dd") & " 00:00:00")
    End If
    
    chk仅显未结清病人.Value = IIf(zlDatabase.GetPara("仅显未结清病人", glngSys, mlngModul, "0") = "1", 1, 0)
    chk仅显未审核病人.Value = IIf(zlDatabase.GetPara("仅显未审核病人", glngSys, mlngModul, "0") = "1", 1, 0)
    
    mvs.OnePati = zlDatabase.GetPara("单次显示", glngSys, mlngModul, "0") = "1"
        
         
    If Val(zlDatabase.GetPara("使用个性化风格", , , True)) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If

    Call RestoreWinState(Me, App.ProductName)
    Call tbcSub_SelectedChanged(tbcSub.Selected)
    '50793
    Call SetFontSize(mbytFontSize)
    Call picFind_Resize
    dkpMain.Panes(1).MinTrackSize.Height = IIf(mbytFontSize = 9, 270, 300)
    dkpMain.Panes(1).MinTrackSize.Width = IIf(mbytFontSize = 9, 225, 295)
    dkpMain.RedrawPanes
End Sub

Private Sub InitRPTPati()
    Dim arrTmp As Variant, arrItem As Variant, i As Long
    Dim rptCol As ReportColumn
    
    With rptPati
        Set .Container = PicRptPati
        arrTmp = Split(mstrPatiHead, ";")
        For i = 0 To UBound(arrTmp)
            arrItem = Split(arrTmp(i), ",")
            If UBound(arrItem) > 0 Then
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), Val(arrItem(1)), True)
                rptCol.Visible = True
                rptCol.Alignment = xtpAlignmentCenter
                
                rptCol.Groupable = Val(arrItem(2)) = 1
            Else
                Set rptCol = .Columns.Add(i, CStr(arrItem(0)), 0, False)
                rptCol.Visible = False
            End If
        Next
        
        .SetImageList imgPati
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .AutoColumnSizing = False
        .ShowGroupBox = True
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有找到符合条件的病人..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

End Sub


Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long
    
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).Style
    End If
    
    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hWnd)
        
    Me.Caption = objItem.Caption
        
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '主窗口重新加入
    Call MainDefCommandBar
    
    '子窗口重新加入
    Select Case objItem.Tag
        Case "费用"
            Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 0)
        Case "医嘱"
            Call mclsAdvices.zlDefCommandBars(Me, Nothing, 1)
    End Select
    
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = bytStyle
        Next
        cbsMain(lngCount).Visible = blnShowBar
        
    Next
    
    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
    Call IDKIND.Refrash
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新子窗体数据及状态
    Dim blnDateMoved As Boolean
    Dim lngDeptID As Long
    
    Select Case objItem.Tag
        Case "费用"
            If mlng病人ID = 0 Then
                '问题:25850
                If cboUnit.ListIndex >= 0 Then lngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
                Call mclsFeeQuery.zlRefresh(0, 0, 0, lngDeptID, 0, False, False, False)
            Else
                With rptPati.SelectedRows(0)
                    If Val(.Record(GetRptRsColumn("数据转出")).Value) = 1 Then
                        blnDateMoved = True
                    Else
                        blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("入院时间")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                    End If
                    '问题:25850
                    lngDeptID = Val(.Record(GetRptRsColumn("当前病区ID")).Value)
                    If cboUnit.ListIndex >= 0 And lngDeptID = 0 Then lngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
                     
                    Call mclsFeeQuery.zlRefresh(mlng病人ID, mlng主页ID, Val(.Record(GetRptRsColumn("住院号")).Value), Val(.Record(GetRptRsColumn("当前病区ID")).Value), _
                        Val(.Record(GetRptRsColumn("险类")).Value), blnDateMoved, Trim(.Record(GetRptRsColumn("出院时间")).Value) <> "", Trim(.Record(GetRptRsColumn("结清")).Value) <> "", , Trim(.Record(GetRptRsColumn("出院时间")).Value) <> "")
                End With
            End If
        Case "医嘱与费用"
                If mlng病人ID <> 0 Then
                    With rptPati.SelectedRows(0)
                        If Val(.Record(GetRptRsColumn("数据转出")).Value) = 1 Then
                            blnDateMoved = True
                        Else
                            blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("入院时间")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                        End If
                    End With
                End If
                Call mfrmPatiFeeVerfy.ShowData(mlng病人ID, mlng主页ID, blnDateMoved)
        Case "医嘱"
            If mlng病人ID = 0 Then
                Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
            Else
                With rptPati.SelectedRows(0)
                    '先更新病人相关的变量，再调医嘱的刷新
                    If Val(.Record(GetRptRsColumn("数据转出")).Value) = 1 Then
                        blnDateMoved = True
                    Else
                        blnDateMoved = zlDatabase.DateMoved(Format(.Record(GetRptRsColumn("入院时间")).Value, "yyyy-MM-dd 00:00:00"), , , Caption)
                    End If
                    
                    Call mclsFeeQuery.zlRefresh(mlng病人ID, mlng主页ID, Val(.Record(GetRptRsColumn("住院号")).Value), Val(.Record(GetRptRsColumn("当前病区ID")).Value), _
                        Val(.Record(GetRptRsColumn("险类")).Value), blnDateMoved, Trim(.Record(GetRptRsColumn("出院时间")).Value) <> "", Trim(.Record(GetRptRsColumn("结清")).Value) <> "", True)

                    Call mclsAdvices.zlRefresh(mlng病人ID, mlng主页ID, Val(.Record(GetRptRsColumn("当前病区ID")).Value), Val(.Record(GetRptRsColumn("当前科室ID")).Value), _
                        IIf(.Record(GetRptRsColumn("出院时间")).Value = "", IIf(Val(.Record(GetRptRsColumn("状态")).Value) = 3, 1, 0), 2), Val(.Record(GetRptRsColumn("数据转出")).Value) = 1)
                End With
            End If
    End Select
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim arrTmp As Variant, i As Long, j As Long
        
    '-----------------------------------------------------
    '菜单定义
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview_Pati, "打印预览病人列表(&T)")
        objControl.BeginGroup = True:        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Pati, "打印病人列表(&O)")
         objControl.IconId = conMenu_File_Print
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel_Pati, "病人列表输出到Excel(&E)")
         objControl.IconId = conMenu_File_Excel
        
        .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)").BeginGroup = True
        .Add xtpControlButton, conMenu_File_Print, "打印(&P)"
        .Add xtpControlButton, conMenu_File_Excel, "输出到Excel(&L)"
        .Add xtpControlButton, conMenu_File_PrintMultiBill, "打印多张催款单(&N)…"
        .Add(xtpControlButton, conMenu_File_SchemeSet, "报警方案设置(&F)…").BeginGroup = True
        .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)").BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_Edit_PreBalanceAll, "预结所有病人(&I)"
        .Add xtpControlButton, conMenu_Edit_PreBalance, "预结当前病人(&W)"
        .Add xtpControlButton, conMenu_Edit_Balance, "结帐(&B)"
        .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt病人审核方式 = 1, "开始审核(&A)", "审核(&A)")).BeginGroup = True
        .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "完成审核(&O)").IconId = 252
        .Add xtpControlButton, conMenu_Edit_FeeUnAudit, "取消审核(&U)"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "病种选择(&Z)")
        objControl.BeginGroup = True
       Set objControl = .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "缴预交(&P)")
       objControl.IconId = 3816: objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
       Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
  
        Set objControl = .Add(xtpControlButton, conMenu_View_FontSize_S, "小字体(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FontSize_L, "大字体(&U)")

        Set objControl = .Add(xtpControlButton, conMenu_View_OnePati, "多次住院只显一次病人(&O)")
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_GroupCol, "病人分组依据(&G)")
        arrTmp = Split(mstrPatiHead, ";")
        For i = 0 To UBound(arrTmp)
            If UBound(Split(arrTmp(i), ",")) > 1 Then
                If Val(Split(arrTmp(i), ",")(2)) = 1 Then
                    j = j + 1
                    objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_View_GroupCol * 10 + j, Split(arrTmp(i), ",")(0)
                End If
            End If
        Next
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With


    '查找项特殊处理
    '-----------------------------------------------------
    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
'        Set objPopup = .Add(xtpControlPopup, conMenu_View_FindType, "查找")
'        objPopup.ID = conMenu_View_FindType
'        objPopup.flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = picFind.hWnd
        objCustom.flags = xtpFlagRightAlign
        IDKIND.BackColor = picFind.BackColor
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop) '固有
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FeeAudit, IIf(gTy_System_Para.byt病人审核方式 = 1, "开始审核", "审核"))
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OverFeeAudit, "完成审核")
        objControl.IconId = 252
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Balance, "结帐")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PrePayMoney, "缴预交")
        objControl.IconId = 3816
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    
    
    
    '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, VK_F3, conMenu_View_FindNext
        .Add 0, VK_F5, conMenu_View_Refresh
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_FeeAudit
        If gTy_System_Para.byt病人审核方式 = 1 Then
            .Add FCONTROL, Asc("O"), conMenu_Edit_OverFeeAudit
        End If
        .Add FCONTROL, Asc("U"), conMenu_Edit_FeeUnAudit
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsMain.Options     '如果隐藏了，控件在菜单第一次显示时没有调用update事件
'        .AddHiddenCommand conMenu_View_Owe
'        .AddHiddenCommand conMenu_View_UnAudit
        '.AddHiddenCommand conMenu_View_OnePati
    End With
        
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1139_3")   '打印催款单
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, lngTmp As Long
    
    SaveWinState Me, App.ProductName
    lngTmp = Val(dtpEnd.Value - dtpBegin.Value)
    If lngTmp > 100 Then lngTmp = 7
    zlDatabase.SetPara "显示字体大小", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
      
    zlDatabase.SetPara "病人状态", cboState.Text, glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "间隔天数", lngTmp, glngSys, mlngModul, mblnHavePara
     zlDatabase.SetPara "仅显未结清病人", IIf(chk仅显未结清病人.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "仅显未审核病人", IIf(chk仅显未审核病人.Value = 1, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "单次显示", IIf(mvs.OnePati, 1, 0), glngSys, mlngModul, mblnHavePara
    zlDatabase.SetPara "病人过滤类别", IDKindPati.IDKIND, glngSys, mlngModul, mblnHavePara
    SaveRegInFor g私有模块, Me.Name, "IDKind", IDKIND.IDKIND
    If Val(zlDatabase.GetPara("使用个性化风格", , , True)) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    Call SaveWinState(Me, App.ProductName)
    mlng病人ID = 0
    mlng主页ID = 0
    mstrPrePati = ""
    mblnUnload = False
    
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mfrmActive = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsAdvices = Nothing
    Set mobjPatient = Nothing
     
End Sub
Private Function CheckIsAllowAudit(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许审核
    '编制:刘兴洪
    '返回:允许审核,返回true,否则返回False
    '日期:2012-06-19 14:04:21
    '问题:50778
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select NO,记录性质 From 住院费用记录 Where 病人ID=[1] and 主页ID=[2] and 记帐费用=1 And 记录状态=0 and rownum<=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
      MsgBox "该病人还有未生效的费用,不能进行审核!", vbInformation + vbOKOnly, gstrSysName
      Exit Function
    End If
    CheckIsAllowAudit = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ExecAuditingAndCancelAudit(ByVal bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行审核或取消审核操作
    '参数:bytType-0-取消审核;1-开始审核或审核;2-完成审核;
    '编制:刘兴洪
    '日期:2012-05-21 11:57:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytAudit As Byte, strSQL As String, strTemp As String, strExpend As String
    Dim blnCheck As Boolean
    If gTy_System_Para.byt病人审核方式 = 1 Then
        If CheckIsAllowAudit(mlng病人ID, mlng主页ID) = False Then Exit Sub
    End If
    
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear
        On Error GoTo errH:
        If Not mobjPlugIn Is Nothing Then
            Call mobjPlugIn.Initialize(gcnOracle, glngSys, mlngModul)
        End If
    End If
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        blnCheck = mobjPlugIn.PatiFeeAuditingAndCancelCheck(mlngModul, mlng病人ID, mlng主页ID, bytType = 0, strExpend)
        If Err = 0 Then
            '存在检查接口
            If blnCheck = False Then
                MsgBox "无法对病人进行" & IIf(bytType = 0, "取消审核!", "审核!"), vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            '不存在检查接口的不检查
            Err.Clear
        End If
        On Error GoTo errH:
    End If
    
    With rptPati.SelectedRows(0)
        'bytAudit-0-未审核;1-开始审核或已审核;2-完成审核
        strTemp = Trim(.Record(GetRptRsColumn("审核状态")).Value)
        bytAudit = Switch(strTemp = "开始" Or strTemp = "已审", 1, strTemp = "完成", 2, True, 0)
        If bytType = 0 Then
            bytAudit = Switch(bytAudit = 2, 1, bytAudit = 1, 0, True, 0)
        Else
            bytAudit = bytType
        End If
        
        On Error GoTo errH
        ' Zl_病人审核_Execute
        strSQL = "Zl_病人审核_Execute("
        '  病人id_In   病案主页.病人id%Type,
        strSQL = strSQL & "" & mlng病人ID & ","
        '  主页id_In   病案主页.主页id%Type,
        strSQL = strSQL & "" & mlng主页ID & ","
        '  审核标志_In 病案主页.审核标志%Type,
        strSQL = strSQL & "" & bytAudit & ","
        '  审核人_In   病案主页.审核人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作方式_In Integer:=0 --操作方式_In:0-审核;1-取消审核
        strSQL = strSQL & IIf(bytType = 0, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        '为速度优化,不重新Call LoadPatients,直接改写状态
        Select Case bytAudit
        Case 1
            .Record(GetRptRsColumn("审核状态")).Value = IIf(gTy_System_Para.byt病人审核方式 = 1, "开始", "已审")
            .Record(GetRptRsColumn("审核人")).Value = UserInfo.姓名
        Case 2
            .Record(GetRptRsColumn("审核状态")).Value = "完成"
            .Record(GetRptRsColumn("审核人")).Value = UserInfo.姓名
        Case Else
            .Record(GetRptRsColumn("审核状态")).Value = ""
            .Record(GetRptRsColumn("审核人")).Value = ""
        End Select
        If bytAudit = "1" Or bytAudit = "2" Then
            .Record(GetRptRsColumn("姓名")).ForeColor = &H33AA22
        Else
            .Record(GetRptRsColumn("姓名")).ForeColor = .Record(GetRptRsColumn("住院号")).ForeColor   '还原病人类型颜色
        End If
    End With
    rptPati.Populate
    Select Case bytType
    Case 0
        sta.Panels(3).Text = "取消" & IIf(gTy_System_Para.byt病人审核方式 = 1, IIf(bytAudit = 0, "开始", "完成"), "") & "审核成功!"
    Case 1
        sta.Panels(3).Text = IIf(gTy_System_Para.byt病人审核方式 = 1, "开始", "") & "审核成功!"
    Case 2
        sta.Panels(3).Text = "完成审核成功!"
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'功能：初始化病区数据
    Dim i As Long, strNodes As String, strDefaultNode As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";所有病区;") = 0, "1,2,3", "护理", True, True, True)
    strNodes = ""
    mblnNotClick = True
    With mrsDept
        Do While Not .EOF
            If Nvl(!站点) <> "" And InStr(1, "," & strNodes & ",", "," & Nvl(!站点) & ",") = 0 Then
                strNodes = strNodes & "," & Nvl(!站点)
            End If
            If mrsDept!ID = UserInfo.部门ID Then strDefaultNode = Nvl(!站点)
             .MoveNext
        Loop
    End With
     
    cboNode.Clear
    If strNodes <> "" Then
        strNodes = Mid(strNodes, 2)
        gstrSQL = "" & _
        "   Select /*+ RULE */A.编号,A.名称 " & _
        "   From zlNodeList A,Table(Cast(f_num2list([1]) As Zltools.t_numlist)) J" & _
        "   where A.编号=j.Column_Value " & _
        "   Order by 编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNodes)
        With cboNode
             Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!编号))
                If .ItemData(.NewIndex) = Val(strDefaultNode) Then
                    .ListIndex = .NewIndex
                End If
                rsTemp.MoveNext
             Loop
             If .ListIndex < 0 And .ListCount >= 1 Then .ListIndex = 0
        End With
    End If
    fra站点.Visible = cboNode.ListCount > 0
    mbln启用站点 = cboNode.ListCount > 0
    If mrsDept.RecordCount <> 0 Then mrsDept.MoveFirst
        
    '问题:50743
    If mrsDept.EOF Then
        MsgBox "没有发现护理病区信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not mrsDept.EOF Then
        If LoadUnits = False Then Exit Function
    ElseIf InStr(";" & mstrPrivs, ";所有病区;") > 0 Then
        MsgBox "没有发现护理病区信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    Call SizeWinCons
    mblnNotClick = False
    InitUnits = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SizeWinCons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整窗体大小
    '编制:刘兴洪
    '日期:2011-02-28 18:06:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim lngTop As Long
     lngTop = IIf(cboNode.ListCount > 0, fra站点.Top + fra站点.Height, fra站点.Top + 10)
     cboUnit.Top = lngTop
     lblUnit.Top = cboUnit.Top + (cboUnit.Height - lblUnit.Height) \ 2
     lngTop = cboUnit.Top + cboUnit.Height + 30
     cboState.Top = lngTop
     lblState.Top = cboState.Top + (cboState.Height - lblState.Height) \ 2
     dtpBegin.Top = cboState.Top + cboState.Height + 30
     lblStartDate.Top = dtpBegin.Top + (dtpBegin.Height - lblStartDate.Height) \ 2
     dtpEnd.Top = dtpBegin.Top + dtpBegin.Height + 30
     lblEndDate.Top = dtpEnd.Top + (dtpEnd.Height - lblEndDate.Height) \ 2
     txt住院号.Top = dtpEnd.Top + dtpEnd.Height + 30
     lbl住院号.Top = txt住院号.Top + (txt住院号.Height - lbl住院号.Height) \ 2
     txt姓名.Top = txt住院号.Top + txt住院号.Height + 30
     IDKindPati.Top = txt姓名.Top + (txt姓名.Height - IDKindPati.Height) \ 2
     txt预交.Top = txt姓名.Top + txt姓名.Height + 30
     chk预交款.Top = txt预交.Top + (txt预交.Height - chk预交款.Height) \ 2
     chk仅显未结清病人.Top = txt预交.Top + txt预交.Height + 30
     chk仅显未审核病人.Top = chk仅显未结清病人.Top + chk仅显未结清病人.Height + 30
     cmdSearch.Top = chk仅显未审核病人.Top + chk仅显未审核病人.Height + 30
End Sub
 
 
Private Function LoadUnits() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:str站点-指定站点,str站点="",表示不区分站点
    '出参:
    '返回:加载成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-02-28 17:51:27
    '问题:36048
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    mrsDept.Filter = 0
    If cboNode.ListIndex >= 0 Then
         mrsDept.Filter = "站点=" & cboNode.ItemData(cboNode.ListIndex) & " or 站点=NULL"
    End If
    cboUnit.Clear
    If InStr(";" & mstrPrivs, ";所有病区;") > 0 Then cboUnit.AddItem "所有病区"
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            cboUnit.AddItem !编码 & "-" & !名称
            cboUnit.ItemData(cboUnit.ListCount - 1) = !ID
            If mrsDept!ID = UserInfo.部门ID Then cboUnit.ListIndex = cboUnit.NewIndex
            .MoveNext
        Loop
    End With
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    LoadUnits = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadPatients(ByVal objCard As Card, Optional blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定范围内的病人列表
    '入参:blnCard-是否刷卡
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-12-01 14:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intBedLen As Integer
    Dim rsPati As ADODB.Recordset, dbl预交余额 As Double
    Dim i As Long, j As Long, blnUnIndex As Boolean, strCount As String, str住院号 As String, str姓名 As String
    Dim DateBegin As Date, DateEnd As Date, lngUnitID As Long, lngPatiRow As Long, strPatiRow As String
    Dim objRecord As ReportRecord, objRow As ReportRow
    Dim objItem As ReportRecordItem, strKind As String, intTYPE As Integer
    Dim strWhere As String, strNodeNo As String   '站点
    Dim lng卡类别ID As Long, lng病人ID As Long
    
    Dim strPassWord As String, strErrMsg As String
    On Error GoTo errH
    Screen.MousePointer = 11
    
    strNodeNo = gstrNodeNo  '当前站点号
    
    Call zlCommFun.ShowFlash("正在统计数据,请稍候 ...", Me)
    DoEvents
    Refresh
    mstrPrePati = ""
    DateBegin = CDate(Format(dtpBegin.Value, "yyyy-MM-dd HH:MM:SS"))
    DateEnd = CDate(Format(dtpEnd.Value, "yyyy-MM-dd HH:MM:SS"))
    '问题:50743
    If cboUnit.ListIndex < 0 Then
        lngUnitID = -1
    Else
        lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    End If
    
    
    str住院号 = Trim(txt住院号.Text)
    str姓名 = Trim(txt姓名.Text)
    
    If cboState.Text = "在院病人" Then
        '当前在院的病人
        strSQL = " And A.在院=1 And Nvl(B.状态,0)<>3 And A.主页ID=B.主页ID "
    ElseIf cboState.Text = "出院病人" Then
        '该期间内出院
        strSQL = " And B.入院日期<=[2] And B.出院日期 Between [1] And [2]"
        If mvs.OnePati Then strSQL = strSQL & " And A.主页ID=B.主页ID"
    ElseIf cboState.Text = "预出院病人" Then
        '预出院病人
        strSQL = " And A.在院=1 And B.状态=3 "
    Else '所有病人
        If (str住院号 = "" And str姓名 = "") Or (str住院号 = "" And Len(str姓名) = 1 And IDKindPati.IDKIND = 1) Then
            strSQL = " And ((A.在院=1  And A.主页ID=B.主页ID) Or (B.出院日期 Between [1] And [2]))"
        Else
            strSQL = ""
        End If
        If mvs.OnePati Then strSQL = strSQL & " And A.主页ID=B.主页ID"
    End If
    
    If str住院号 <> "" Then
        strSQL = strSQL & " And A.病人ID = (Select distinct 病人ID From 病案主页 Where 住院号=[4])": blnUnIndex = True
    End If
    
    If str姓名 <> "" Then
         strKind = objCard.名称
        Select Case strKind
            Case "姓名"
                lng病人ID = Val(txt姓名.Tag)
                If blnCard Then
                    If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
                        lng卡类别ID = IDKIND.GetfaultCard.接口序号
                    Else
                        lng卡类别ID = "-1"
                    End If
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, str姓名, True, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then lng病人ID = 0
                    txt姓名.Tag = lng病人ID
                    lng病人ID = Val(txt姓名.Tag)
                    strSQL = strSQL & " And A.病人ID=[10]"
                ElseIf lng病人ID <> 0 Then
                    strSQL = strSQL & " And A.病人ID=[10]"
                Else
                    strSQL = strSQL & " And A.姓名 like [5]": If Len(str姓名) > 1 Then blnUnIndex = True    '如果是只有一个字母,可能影响性能,所以不用索引
                End If
            Case "床号"
                strSQL = strSQL & " And B.出院病床=[6]":   blnUnIndex = True
            Case "医保号"
                strSQL = strSQL & " And (F.医保号=[6] or F.医保号 IS NULL and  D.信息值=[6])":   blnUnIndex = True
            Case "住院号"
                If gbln每次住院新住院号 Or True Then
                    strSQL = strSQL & " And A.病人ID = (Select distinct 病人ID From 病案主页 Where 住院号=[4])": blnUnIndex = True
                Else
                    strSQL = strSQL & " And A.住院号=[4]": blnUnIndex = True
                End If
                '问题:50788
                str住院号 = str姓名
            Case Else
               lng卡类别ID = objCard.接口序号
                If lng卡类别ID > 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, str姓名, True, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, str姓名, True, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strSQL = strSQL & " And A.病人ID=[10]"
        End Select
    End If
    
    If cboState.Text = "出院病人" Then blnUnIndex = True
    If lngUnitID > 0 Then strSQL = strSQL & " And B.当前病区ID" & IIf(blnUnIndex, "+0", "") & "=[3]"
     If cboNode.ListCount > 0 And cboNode.ListIndex >= 0 Then
        '只能查该站点的所有病区的病人
        strNodeNo = cboNode.ItemData(cboNode.ListIndex)
    End If
    
    If chk仅显未结清病人.Value = 1 Then strSQL = strSQL & " And Nvl(X.费用余额,0) <> 0 And Exists (Select 1 From 病人未结费用 Where 病人id = a.病人id And Nvl(主页id,0) = Nvl(b.主页id,0) And Rownum < 2)"
    If chk仅显未审核病人.Value = 1 Then strSQL = strSQL & " And 审核人 is Null"
     
    
    dbl预交余额 = Val(txt预交.Text)
    intBedLen = GetMaxBedLen(lngUnitID)
    
    If chk预交款.Value = 1 Then
        '问题:27546
        '显示病人预交款余额小于？？少的病人
        '因为预出院病人的出院时间在变动记录中,所以要取病人变动记录这张表
        If cboState.Text = "在院病人" Or cboState.Text = "出院病人" Then
            strSQL = "Select A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出, B.出院科室id As 当前科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID, Nvl(b.姓名, a.姓名) As 姓名, B.住院号, B.出院病床 As 床号," & vbNewLine & _
                "       B.费别, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) as 年龄, B.入院日期, B.出院日期, C.名称 As 当前科室, Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清," & vbNewLine & _
                "       Nvl(E.医保号, D.信息值) 医保号, A.家庭电话, B.医疗付款方式,B.审核标志, B.审核人, B.病人类型, H.名称 当前病区" & vbNewLine & _
                "From 病人信息 A, 病案主页 B, 病案主页从表 D, 医保病人档案 E, 医保病人关联表 F, 病人余额 X,保险模拟结算 X1, 部门表 C, 部门表 H" & vbNewLine & _
                "Where A.病人ID = B.病人ID And B.出院科室ID = C.ID And Nvl(B.主页ID, 0) <> 0 And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And" & vbNewLine & _
                "      D.信息名(+) = '医保号' And A.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2 And B.病人ID=X1.病人ID(+) and B.主页id=X1.主页ID(+) And A.病人ID = F.病人ID(+) And A.险类 = F.险类(+) And F.标志(+) = 1 And" & vbNewLine & _
                "      F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+) And B.当前病区ID + 0 = H.ID" & vbNewLine & _
                "    And (H.站点=[8] Or H.站点 is Null)" & vbNewLine & strSQL & vbNewLine & _
                "Group by A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出, B.出院科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID,  Nvl(b.姓名, a.姓名), B.住院号, B.出院病床,B.费别, Nvl(b.性别, a.性别), Nvl(b.年龄, a.年龄), B.入院日期, B.出院日期, C.名称, Decode(Nvl(X.费用余额, 0), 0, '√', ''), " & vbNewLine & _
                "          Nvl(E.医保号, D.信息值), A.家庭电话, B.医疗付款方式,B.审核标志, B.审核人, B.病人类型, H.名称 " & _
                "having (Max(nvl(X.预交余额,0))-Max(nvl(x.费用余额,0))+Sum(nvl(X1.金额,0)))<[7] " & vbNewLine & _
                IIf(lngUnitID = 0, " Order by 住院号 Desc", " Order by LPAD(床号,10,' ')")
                
        Else
            strSQL = "Select A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出,B.出院科室id As 当前科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID, Nvl(b.姓名, a.姓名) As 姓名, B.住院号, B.出院病床 As 床号," & vbNewLine & _
                "       B.费别, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) as 年龄, B.入院日期, Decode(B.出院日期, Null, Z.开始时间, B.出院日期) 出院日期, C.名称 As 当前科室," & vbNewLine & _
                "       Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清, Nvl(E.医保号, D.信息值) 医保号, A.家庭电话, B.医疗付款方式,B.审核标志," & vbNewLine & _
                "       B.审核人, B.病人类型, H.名称 当前病区" & vbNewLine & _
                "From 病人信息 A, 病案主页 B, 病案主页从表 D, 病人变动记录 Z, 医保病人档案 E, 医保病人关联表 F, 病人余额 X,保险模拟结算 X1, 部门表 C, 部门表 H" & vbNewLine & _
                "Where A.病人ID = B.病人ID And B.出院科室ID = C.ID And Nvl(B.主页ID, 0) <> 0 And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And" & vbNewLine & _
                "      D.信息名(+) = '医保号' And A.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2 And A.病人ID = F.病人ID(+) And A.险类 = F.险类(+) And F.标志(+) = 1 And" & vbNewLine & _
                "      F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+) And B.当前病区ID + 0 = H.ID" & vbNewLine & _
                "    And (H.站点=[8] Or H.站点 is Null)" & vbNewLine & _
                "   And B.病人ID = Z.病人ID(+) And B.主页ID = Z.主页ID(+) And Z.开始原因(+) = 10 And Z.附加床位(+) = 0" & vbNewLine & _
                "   And B.病人ID=X1.病人ID(+) and B.主页id=X1.主页ID(+)  " & vbNewLine & strSQL & vbNewLine & _
                "Group by A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出,B.出院科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID, Nvl(b.姓名, a.姓名), B.住院号, B.出院病床," & _
                "         B.费别, Nvl(b.性别, a.性别) , Nvl(b.年龄, a.年龄), B.入院日期, Decode(B.出院日期, Null, Z.开始时间, B.出院日期) , C.名称,Decode(Nvl(X.费用余额, 0), 0, '√', '')  , Nvl(E.医保号, D.信息值) , A.家庭电话, B.医疗付款方式," & _
                "         B.审核标志,B.审核人, B.病人类型, H.名称 " & vbNewLine & _
                "having (Max(nvl(X.预交余额,0))-Max(nvl(x.费用余额,0))+Sum(nvl(X1.金额,0)))<[7] " & vbNewLine & _
                 IIf(lngUnitID = 0, " Order by 住院号 Desc", " Order by LPAD(床号,10,' ')")
        End If
        
    Else
            '因为预出院病人的出院时间在变动记录中,所以要取病人变动记录这张表
            If cboState.Text = "在院病人" Or cboState.Text = "出院病人" Then
                strSQL = "Select A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出, B.出院科室id As 当前科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID, Nvl(b.姓名, a.姓名) As 姓名, B.住院号, B.出院病床 As 床号," & vbNewLine & _
                    "       B.费别, Nvl(b.性别, a.性别) As 性别 , Nvl(b.年龄, a.年龄) as 年龄, B.入院日期, B.出院日期, C.名称 As 当前科室, Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清," & vbNewLine & _
                    "       Nvl(E.医保号, D.信息值) 医保号, A.家庭电话, B.医疗付款方式,B.审核标志, B.审核人, B.病人类型, H.名称 当前病区" & vbNewLine & _
                    "From 病人信息 A, 病案主页 B, 病案主页从表 D, 医保病人档案 E, 医保病人关联表 F, 病人余额 X, 部门表 C, 部门表 H" & vbNewLine & _
                    "Where A.病人ID = B.病人ID And B.出院科室ID = C.ID And Nvl(B.主页ID, 0) <> 0 And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And" & vbNewLine & _
                    "      D.信息名(+) = '医保号' And A.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2 And A.病人ID = F.病人ID(+) And A.险类 = F.险类(+) And F.标志(+) = 1 And" & vbNewLine & _
                    "      F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+) And B.当前病区ID + 0 = H.ID" & vbNewLine & _
                    "    And (H.站点=[8] Or H.站点 is Null)" & vbNewLine & _
                    strSQL & IIf(lngUnitID = 0, " Order by 住院号 Desc", " Order by LPAD(床号,10,' ')")
            Else
                strSQL = "Select A.病人ID, B.主页ID, A.登记时间, B.状态, B.病人性质, B.数据转出,B.出院科室id As 当前科室id, A.就诊卡号, B.险类, E.密码, B.当前病区ID,Nvl(b.姓名, a.姓名) As 姓名, B.住院号, B.出院病床 As 床号," & vbNewLine & _
                    "       B.费别, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) as 年龄, B.入院日期, Decode(B.出院日期, Null, Z.开始时间, B.出院日期) 出院日期, C.名称 As 当前科室," & vbNewLine & _
                    "       Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清, Nvl(E.医保号, D.信息值) 医保号, A.家庭电话, B.医疗付款方式," & vbNewLine & _
                    "       B.审核标志,B.审核人, B.病人类型, H.名称 当前病区" & vbNewLine & _
                    "From 病人信息 A, 病案主页 B, 病案主页从表 D, 病人变动记录 Z, 医保病人档案 E, 医保病人关联表 F, 病人余额 X, 部门表 C, 部门表 H" & vbNewLine & _
                    "Where A.病人ID = B.病人ID And B.出院科室ID = C.ID And Nvl(B.主页ID, 0) <> 0 And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And" & vbNewLine & _
                    "      D.信息名(+) = '医保号' And A.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2 And A.病人ID = F.病人ID(+) And A.险类 = F.险类(+) And F.标志(+) = 1 And" & vbNewLine & _
                    "      F.医保号 = E.医保号(+) And F.险类 = E.险类(+) And F.中心 = E.中心(+) And B.当前病区ID + 0 = H.ID" & vbNewLine & _
                    "    And (H.站点=[8] Or H.站点 is Null)" & vbNewLine & _
                    " And B.病人ID = Z.病人ID(+) And B.主页ID = Z.主页ID(+) And Z.开始原因(+) = 10 And Z.附加床位(+) = 0" & vbNewLine & _
                    strSQL & IIf(lngUnitID = 0, " Order by 住院号 Desc", " Order by LPAD(床号,10,' ')")
            End If
    End If
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Caption, DateBegin, DateEnd, lngUnitID, Val(str住院号), str姓名 & "%", str姓名, dbl预交余额, strNodeNo, str住院号, lng病人ID)
    '记录现在选中的病人
    If rptPati.SelectedRows.Count > 0 And mlng病人ID <> 0 Then
        If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
            lngPatiRow = rptPati.SelectedRows(0).Index '用于快速重新定位
            strPatiRow = rptPati.SelectedRows(0).Record.Tag
        End If
    End If
    rptPati.Records.DeleteAll
    
    If rsPati.RecordCount > 0 Then
        With rsPati
            
            For i = 1 To .RecordCount
                
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = !病人ID & "," & Val("" & !主页ID)
                '隐藏列
                objRecord.AddItem Val(!病人ID)
                objRecord.AddItem Val("" & !主页ID)
                objRecord.AddItem Format(!登记时间, "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem Val("" & !状态)
                objRecord.AddItem Val("" & !病人性质)
                objRecord.AddItem Val("" & !数据转出)
                objRecord.AddItem Val("" & !当前科室id)
                objRecord.AddItem CStr("" & !险类)
                objRecord.AddItem CStr("" & !密码)
                objRecord.AddItem Val("" & !当前病区ID)
                objRecord.AddItem ("" & !就诊卡号)
                
                '显示列
                Set objItem = objRecord.AddItem("")
                objItem.Icon = IIf(Val("" & !病人性质) = 0, 0, 1)
                
                Set objItem = objRecord.AddItem(CStr("" & !姓名))
                If IsNull(!审核人) = False Then objItem.ForeColor = &H33AA22
                
                objRecord.AddItem "" & !住院号
                
                If intBedLen > Len("" & !床号) Then
                    objRecord.AddItem Space(intBedLen - Len("" & !床号)) & !床号
                Else
                    objRecord.AddItem "" & !床号
                End If
                objRecord.AddItem CStr("" & !费别)
                objRecord.AddItem CStr("" & !性别)
                objRecord.AddItem CStr("" & !年龄)
                objRecord.AddItem Format(Nvl(!入院日期, ""), "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem Format(Nvl(!出院日期, ""), "yyyy-MM-dd HH:mm:ss")
                objRecord.AddItem CStr("" & !当前科室)
                objRecord.AddItem Val("" & !主页ID)
                objRecord.AddItem CStr("" & !结清)
                objRecord.AddItem CStr("" & !医保号)
                objRecord.AddItem CStr("" & !家庭电话)
                objRecord.AddItem CStr("" & !医疗付款方式)
                If Val(Nvl(!审核标志)) = 1 Then
                    objRecord.AddItem IIf(gTy_System_Para.byt病人审核方式 = 1, "开始", "已审")
                ElseIf Val(Nvl(!审核标志)) = 2 Then
                    objRecord.AddItem CStr("完成")
                Else
                    objRecord.AddItem " "
                End If
                objRecord.AddItem CStr("" & !审核人)
                objRecord.AddItem CStr("" & !病人类型)
                objRecord.AddItem CStr("" & !当前病区)
                
                For j = 0 To rptPati.Columns.Count - 1
                    If Not (IsNull(!审核人) = False And rptPati.Columns(j).Caption = "姓名") Then
                        objRecord.Item(j).ForeColor = zlDatabase.GetPatiColor(Nvl(!病人类型))
                    End If
                Next
                                
                If Not mvs.OnePati Then
                    If InStr(strCount & ",", "," & rsPati!病人ID & ",") = 0 Then strCount = strCount & "," & rsPati!病人ID    '可能多次住院显示了多条记录，所以不能直接求记录数
                End If
                rsPati.MoveNext
            Next
        End With
        rptPati.Populate
        
        '取指定病人行
        If strPatiRow <> "" Then
            '先快速定位
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        '取第一个非分组行
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        
        '查询结果不唯一时,缺省不显示当前病人费用相关信息
        If Not (rsPati.RecordCount > 1 And lngPatiRow = 0) Then
            Set rptPati.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        End If
                
        If mvs.OnePati Then
            i = rptPati.Records.Count
        Else
            i = UBound(Split(Mid(strCount, 2), ",")) + 1
        End If
                
        If cboState.Text = "在院病人" Then
            sta.Panels(3).Text = " 该病区在院病人人数:" & i
        ElseIf cboState.Text = "出院病人" Then
            sta.Panels(3).Text = " 时间: " & Format(dtpBegin.Value, "yyyy-MM-dd") & " 至 " & Format(dtpEnd.Value, "yyyy-MM-dd") & ",人数:" & i
        ElseIf cboState.Text = "预院病人" Then
            sta.Panels(3).Text = " 预出院病人人数:" & i
        Else
            sta.Panels(3).Text = " 所有病人,人数:" & i
        End If
    Else
        rptPati.Populate
        mlng病人ID = 0: mlng主页ID = 0
        Call tbcSub_SelectedChanged(tbcSub.Selected)
        sta.Panels(2).Text = "指定的条件没有筛选出任何病人."
        sta.Panels(3).Text = ""
                  
    End If
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Refresh
    Exit Function
errH:
    Call zlCommFun.StopFlash
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetRptRsColumn(ByVal strColumn As String) As Long
'功能：根据列名返回记录集的列序号(界面改变列顺序后，它不变)，没有找到时返回-1
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    arrTmp = Split(mstrPatiHead, ";")
    
    GetRptRsColumn = -1
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If InStr(1, strTmp, ",") > 0 Then strTmp = Mid(strTmp, 1, InStr(1, strTmp, ",") - 1)
        
        If strTmp = strColumn Then GetRptRsColumn = i: Exit For
    Next
End Function

Private Sub ExecFindPati(ByVal objCard As Card, Optional blnNext As Boolean, Optional blnCard As Boolean = True)
    Dim strKind As String, strValue As String, i As Long, lngPoint As Long
    Dim lng住院号 As Long, lng姓名 As Long, lng就诊卡 As Long, lng床号 As Long, lng医保号 As Long
    Dim lng卡类别ID As Long, lng病人ID As Long, lng病人IDCol As Long
    Dim strErrMsg  As String
    
    If rptPati.Records.Count = 0 Then Exit Sub
    strValue = Trim(txtFind.Text)
    If strValue = "" Then Call txtFind.SetFocus: Exit Sub
    strKind = objCard.名称
    If Not IsNumeric(strValue) And strKind = "住院号" Then
        MsgBox "住院号要求输入数字值!", vbInformation, gstrSysName
        Call txtFind.SetFocus: Call zlControl.TxtSelAll(txtFind)
        Exit Sub
    End If
    
    If Not blnNext Then
        lngPoint = 0
    Else
        lngPoint = rptPati.SelectedRows(0).Index + 1
    End If
    
    Select Case strKind
        Case "床号"
            lng床号 = GetRptRsColumn("床号")
        Case "住院号"
            lng住院号 = GetRptRsColumn("住院号")
        Case "姓名"
            lng姓名 = GetRptRsColumn("姓名")
            '读卡或刷卡
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
            strErrMsg = ""
            If blnNext = False Then
                If blnCard Then
                    If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
                        lng卡类别ID = IDKIND.GetfaultCard.接口序号
                    Else
                        lng卡类别ID = "-1"
                    End If
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strValue, True, lng病人ID, "", strErrMsg, lng卡类别ID) = False Then lng病人ID = 0
                    txtFind.Tag = lng病人ID
                    lng病人ID = Val(txtFind.Tag)
                    strKind = "病人ID"
                    lng病人IDCol = GetRptRsColumn("病人ID")
                End If
            Else
                 lng病人ID = Val(txtFind.Tag)
                 If lng病人ID <> 0 Then strKind = "病人ID": lng病人IDCol = GetRptRsColumn("病人ID")
           End If
        Case "医保号"
            lng医保号 = GetRptRsColumn("医保号")
        Case Else ' "就诊卡"
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
            strErrMsg = ""
            If blnNext = False Then
                    lng卡类别ID = objCard.接口序号
                    If lng卡类别ID > 0 Then
                        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strValue, True, lng病人ID, "", strErrMsg) = False Then GoTo NotFoundPati:
                        If lng病人ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(strKind, strValue, True, lng病人ID, _
                            "", strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                    If lng病人ID <= 0 Then GoTo NotFoundPati:
                    txtFind.Tag = lng病人ID
            End If
            lng病人ID = Val(txtFind.Tag)
            strKind = "病人ID"
            lng病人IDCol = GetRptRsColumn("病人ID")
    End Select
    '查找病人
    For i = lngPoint To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            Select Case strKind
                Case "床号"
                    If Trim(rptPati.Rows(i).Record(lng床号).Value) = strValue Then Exit For
                Case "住院号"
                    If rptPati.Rows(i).Record(lng住院号).Value = strValue Then Exit For
                Case "姓名"
                    If rptPati.Rows(i).Record(lng姓名).Value Like IIf(gstrLike = "%", "*", "") & strValue & "*" Then Exit For
                Case "医保号"
                    If rptPati.Rows(i).Record(lng医保号).Value = strValue Then Exit For
                Case Else
                    If rptPati.Rows(i).Record(lng病人IDCol).Value = lng病人ID Then Exit For
            End Select
        End If
    Next
    If i = rptPati.Rows.Count Then GoTo NotFoundPati:
    Set rptPati.FocusedRow = rptPati.Rows(i)    '引发SelectionChanged事件
    Exit Sub
NotFoundPati:
    
    If blnNext Then
        MsgBox "已经没有符合输入条件的病人！", vbInformation, gstrSysName
    Else
        If strErrMsg <> "" Then
            MsgBox strErrMsg, vbInformation, gstrSysName
        Else
            MsgBox "没有找到符合输入条件的病人！", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub chk预交款_Click()
    txt预交.Enabled = chk预交款.Value = 1
End Sub

Private Sub chk预交款_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt预交_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt预交_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt预交, KeyAscii, m负金额式
End Sub
Private Sub zlSchemeSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:报警方案设置
    '编制:刘兴洪
    '日期:2011-01-20 09:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmPatiPressMoneySet.zlShowMe Me, mlngModul, mstrPrivs
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2011-01-31 14:22:25
    '问题:35550
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    '装载数据
    Call zlRptControlToVsGrid(rptPati, vsPrint)
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "病人清单"
    objRow.Add "病人病区：" & cboUnit.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "病人状态：" & cboState.Text
    objRow.Add "开始日期：" & Format(dtpBegin.Value, "yyyy-mm-dd")
    objRow.Add "结束日期：" & Format(dtpEnd.Value, "yyyy-mm-dd")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    bytPrn = bytFunc
    Err = 0: On Error GoTo ErrHand:
    
    Set objPrint.Body = vsPrint
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
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
 Private Function ExecPrePayMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行缴预交款
    '返回:执行成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-02-17 15:16:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dbl缴款额 As Double, bln门诊留观病人 As Boolean
    
    On Error GoTo errHandle
    If mobjPatient Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjPatient = CreateObject("zl9Patient.clsPatient")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   病人管理部件不存在,不能缴预交,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Err = 0
            Exit Function
        End If
    End If
    
    bln门诊留观病人 = ZlIsOutpatientObserve(mlng病人ID, mlng主页ID)
    strSQL = "Select Zl1_Getdef_Prepaymoney([1],[2],[3]) as 缴款额 from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, IIf(bln门诊留观病人, 1, 2))
    dbl缴款额 = Nvl(rsTemp!缴款额, 0)
    
    '门诊留观病人缴门诊预交
    'PlusDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    '    ByVal strDBUser As String, Optional bytCallObject As Byte = 0, _
    '    Optional lng病人ID As Long, Optional lng主页ID As Long, Optional dblDefPrePayMoney As Double = 0) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能： 调用预交款收款窗口
    '    '参数：
    '    '   lngModul:需要执行的功能序号
    '    '   cnMain:主程序的数据库连接
    '    '   frmMain:主窗体
    '    '   strDBUser:当前数据库登录用户名
    '    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用)
    '    '  lng病人ID-缺省的病人ID
    '    '  lng主页ID-缺省的主页ID
    '    '  dblDefPrePayMoney-缺省的预付金额
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    ExecPrePayMoney = mobjPatient.PlusDeposit(glngSys, gcnOracle, Me, _
        gstrDBUser, 1, mlng病人ID, IIf(bln门诊留观病人, 0, mlng主页ID), dbl缴款额)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Sub InitMenus()
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtFind)
    Call IDKindPati.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0|0|0|0|0|0;床|床号|0|0|0|0|0|0;医|医保号|1|0|0|0|0|0", txt姓名)
    
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKIND.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKIND.Cards.按缺省卡查找
End Sub

Private Sub ModeInsurePatiDisease()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病人病种
    '编制:刘兴洪
    '日期:2013-02-20 11:41:07
    '问题:31883
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String
    'ChooseDisease(ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal intInsure As Integer = 0, _
    Optional ByRef strAdvance As String = "")
    If mlng病人ID = 0 Or mintInsure = 0 Then Exit Sub
    Call gclsInsure.ChooseDisease(mlng病人ID, mlng主页ID, mintInsure, strAdvance)
End Sub




