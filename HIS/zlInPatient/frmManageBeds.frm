VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBeds 
   Caption         =   "病区床位管理"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManageBeds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   2760
      ScaleHeight     =   3135
      ScaleWidth      =   2505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   2505
      Begin XtremeReportControl.ReportControl rptUnit 
         Height          =   2130
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   3757
         _StockProps     =   0
         MultipleSelection=   0   'False
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   2505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2505
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2130
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   3757
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
         MultipleSelection=   0   'False
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7860
      Width           =   11535
      _ExtentX        =   20346
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
            Object.Width           =   15266
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
            AutoSize        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1260
      Left            =   6720
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   2222
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":0CCA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":0EE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":10FE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1318
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1532
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":174C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1966
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1B80
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1D9A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1FB4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":21CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   1440
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":2AA8
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":2DC2
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":30DC
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":33F6
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3710
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3A2A
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":3D44
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":405E
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4378
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4692
            Key             =   "MASK_共用_非编"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2025
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":49AC
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4CC6
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":4FE0
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":52FA
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5614
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":592E
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5A88
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5BE2
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5D3C
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5E96
            Key             =   "MASK_共用_非编"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":5FF0
            Key             =   "加床_Empty"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":6CCA
            Key             =   "非编_Empty"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":79A4
            Key             =   "共用_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":867E
            Key             =   "共用_加床_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":9358
            Key             =   "共用_非编_Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":A032
            Key             =   "加床_M_Empty"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":AD0C
            Key             =   "非编_M_Empty"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":B9E6
            Key             =   "共用_M_Empty"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":C6C0
            Key             =   "共用_加床_M_Empty"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":D39A
            Key             =   "共用_非编_M_Empty"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":E074
            Key             =   "加床_F_Empty"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":ED4E
            Key             =   "非编_F_Empty"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":FA28
            Key             =   "共用_F_Empty"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":10702
            Key             =   "共用_加床_F_Empty"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":113DC
            Key             =   "共用_非编_F_Empty"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":120B6
            Key             =   "加床_Holding"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":12D90
            Key             =   "非编_Holding"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":13A6A
            Key             =   "共用_Holding"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":14744
            Key             =   "共用_加床_Holding"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":1541E
            Key             =   "共用_非编_Holding"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":160F8
            Key             =   "加床_Remedy"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":169D2
            Key             =   "非编_Remedy"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":176AC
            Key             =   "共用_Remedy"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":18386
            Key             =   "共用_加床_Remedy"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":19060
            Key             =   "共用_非编_Remedy"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBeds.frx":19D3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmManageBeds.frx":1A614
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmManageBeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnUnload As Boolean
Private mintEmpty As Integer, intHolding, intRemedy As Integer
Private Const STR_HEAD = "床号,600,0,1;科室,1200,0,2;房间号,800,0,2;状态,600,0,2;性别分类,1000,0,2;等级,1000,0,2;床位编制,1000,0,2;姓名,1000,0,0;性别,600,0,0;年龄,600,0,0"
Private mstrPrivs As String

Const conPane_Type = 201
Const conPane_List = 202
Const conPane_Edit = 203

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mfrmEditBed As frmBedEdit

Private mlngUnit As Long
Private mLngEditWidth As Long       '编辑区域宽度
Private mstrGroupBy As String       '记录分组的列
'床号,科室,房间号,状态,性别分类,等级,床位编制,姓名,性别,年龄,病区ID,等级id,病人ID,共用,科室ID
Public Enum mCol
    图标 = 0: 床号: 科室: 房间号: 顺序号: 状态: 性别分类: 等级: 床位编制: 单价: 姓名: 性别: 年龄: 住院状态: 病区ID: 科室ID: 病人ID: 等级ID: 共用
End Enum

Private Enum mIcon
    iEmpty = 0: iM_Empty: iF_Empty: i_To_Empty: i_To_Repair
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Dim objControl As CommandBarControl
    'Dim objCombo As CommandBarComboBox
    Dim objRow As ReportRow, i As Long
    
    Dim strBedNO As String
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    '------------------------------------
    
    'Set objCombo = cbsMain(cbsMain.Count).FindControl(, conMenu_Edit_SelUnit, True)
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_Save:                                                                 '保存
        strBedNO = mfrmEditBed.zlEditSave
        If strBedNO <> "" Then
            Call zlRefList(strBedNO)
            If mfrmEditBed.chkContAdd.Value Then
                If mfrmEditBed.zlEditStart(True, mlngUnit) = False Then
                    MsgBox "数据初始化错误！", vbExclamation, gstrSysName
                    Exit Sub
                    
                End If
            Else
                ShowEdit False
                mintEditState = 0: Me.picUnit.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
            End If
        End If

    Case conMenu_Edit_Untread:                                                              '取消
        Call ShowEdit(False)
        Call mfrmEditBed.zlEditCancel
        mintEditState = 0: Me.picUnit.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem                                                               '增加

        If mlngUnit <= 0 Then
            MsgBox "请选择病区！", vbExclamation, gstrSysName
            rptUnit.SetFocus
            Exit Sub
        End If
        If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit
        
        Call ShowEdit(True)
        If mfrmEditBed.zlEditStart(True, mlngUnit) = False Then Call ShowEdit(False): Exit Sub
        mintEditState = 1: Me.picUnit.Enabled = False: Me.picList.Enabled = False
    Case conMenu_Edit_Modify                                                            '调整
        If mlngUnit <= 0 Then
            MsgBox "请选择病区！", vbExclamation, gstrSysName: Exit Sub
        End If
        
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择要调整的病床！", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.状态).Value = "占用" Then
                MsgBox "该病床已被病人占用,现在不能进行调整！", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.状态).Value = "占用" = "修缮" Then
                MsgBox "该病床正在修缮,现在不能进行调整！", vbExclamation, gstrSysName: Exit Sub
            End If
        End With
        
'        On Error Resume Next
'        Err.Clear
        Call ShowEdit(True)
        If mfrmEditBed.zlEditStart(False, mlngUnit, rptList.FocusedRow.Record) = False Then Call ShowEdit(False):  Exit Sub
        
        mintEditState = 1: Me.picUnit.Enabled = False: Me.picList.Enabled = False
    Case conMenu_Edit_Delete                                                            '删除

        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择要撤消的病床！", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.状态).Value = "占用" Then
                MsgBox "该病床已被病人占用,现在不能撤消！", vbExclamation, gstrSysName: Exit Sub
            End If
            If MsgBox("确实要撤消病床" & .FocusedRow.Record(mCol.床号).Value & " 吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            intIndex = .FocusedRow.Index

            gstrSQL = "zl_床位状况记录_Delete('" & Trim(.FocusedRow.Record(mCol.床号).Value) & "'," & mlngUnit & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            If .Rows.Count > intIndex + 1 Then
                If .Rows(intIndex + 1).GroupRow = False Then strBedNO = .Rows(intIndex + 1).Record(mCol.床号).Value
            ElseIf intIndex > 0 Then
                If .Rows(intIndex - 1).GroupRow = False Then strBedNO = .Rows(intIndex - 1).Record(mCol.床号).Value
            End If
            Call Me.zlRefList(strBedNO)
        End With
    Case conMenu_Edit_Bed_ToRepair                                                          '转修缮
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择要修缮的病床！", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.状态).Value <> "空床" Then
                MsgBox "该病床不是空床,不能执行该操作！", vbExclamation, gstrSysName: Exit Sub
            End If
            gstrSQL = "zl_床位状况记录_STOP('" & Trim(.FocusedRow.Record(mCol.床号).Value) & "'," & .FocusedRow.Record(mCol.病区ID).Value & ")"
            
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            strBedNO = Trim(.FocusedRow.Record(mCol.床号).Value)
            Call zlRefList(strBedNO)

        End With
    Case conMenu_Edit_Bed_ToEmpty                                                           '转空床
        With rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择已经修缮好的病床！", vbExclamation, gstrSysName: Exit Sub
            End If
            If .FocusedRow.Record(mCol.状态).Value <> "修缮" Then
                MsgBox "该病床没有进行修缮,不能执行该操作！", vbExclamation, gstrSysName: Exit Sub
            End If
            
            gstrSQL = "zl_床位状况记录_REUSE('" & Trim(.FocusedRow.Record(mCol.床号).Value) & "'," & .FocusedRow.Record(mCol.病区ID).Value & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            strBedNO = Trim(.FocusedRow.Record(mCol.床号).Value)
            Call zlRefList(strBedNO)
        End With
        
    Case conMenu_View_ToolBar_Button                                                        '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text                                                          '显示文本
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If Not (objControl.Type = xtpControlLabel Or objControl.Type = xtpControlComboBox) Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size                                                          '大图标/小图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar                                                             '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Column                                                                '选择列
        
    Case conMenu_View_Refresh                                                               '刷新
        '59753:刘鹏飞,2013-4-18
        If Me.rptList.FocusedRow Is Nothing Then
            zlRefList
        Else
            If Me.rptList.FocusedRow.GroupRow Then
                zlRefList
            Else
                zlRefList Trim(Me.rptList.FocusedRow.Record(mCol.床号).Value)
            End If
        End If
    Case conMenu_Help_Help:     Call ShowHelp("", Me.hWnd, Me.Name, Int((glngSys) / 100))   '帮助
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        '--执行自定义报表
        If Control.ID > 401 And Control.ID < 499 Then
            If Me.rptList.FocusedRow Is Nothing Then
                strBedNO = ""
            Else
                If Me.rptList.FocusedRow.GroupRow Then
                    strBedNO = ""
                Else
                    strBedNO = Trim(Me.rptList.FocusedRow.Record(mCol.床号).Value)
                End If
            End If
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                "病区=" & mlngUnit, "床号=" & strBedNO)
        End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
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
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        Control.Enabled = (mintEditState <> 0)

    Case conMenu_Edit_NewItem
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "床位编辑") > 0 And mintEditState = 0)
        
    Case conMenu_Edit_Modify
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "床位编辑") > 0 And mintEditState = 0 And Me.rptList.Rows.Count)
        'If Control.Enabled Then Control.Enabled = mstr编码 <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_Edit_Delete
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        Control.Enabled = (InStr(1, mstrPrivs, "床位编辑") > 0 And mintEditState = 0 And Me.rptList.Rows.Count And rptList.FocusedRow.Record(mCol.状态).Value <> "修缮")
        'If Control.Enabled Then Control.Enabled = mstr编码 <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_Edit_Bed_ToRepair
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        If Me.rptList.Records.Count > 0 Then
            Control.Enabled = (mintEditState = 0 And rptList.FocusedRow.Record(mCol.状态).Value = "空床")
        Else
            Control.Enabled = False
        End If
        
    Case conMenu_Edit_Bed_ToEmpty
        Control.Visible = InStr(1, mstrPrivs, "床位编辑") > 0
        If Me.rptList.Records.Count > 0 Then
            Control.Enabled = (mintEditState = 0 And rptList.FocusedRow.Record(mCol.状态).Value = "修缮")
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsMain(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = 1
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit

    Select Case Item.ID
    Case conPane_Type
        Item.Handle = Me.picUnit.hWnd
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        Item.Handle = mfrmEditBed.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me
End Sub

Private Sub Form_Load()
    
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mLngEditWidth = frmBedEdit.ScaleWidth
    
    
    Call InitCommandBar
    Call InitDockPannel
    Call InitReportColumn
    Call RestoreWinState(Me, App.ProductName)
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, False)

    Call MakeBedIcon

    '读取病区
    If Not InitUnits Then mblnUnload = True: Exit Sub
    If rptUnit.Records.Count = 0 Then
        MsgBox "你不具有所有病区的权限,并且不能确定你所属病区,不能使用床位管理！", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If

'    If Not ReadBeds(mlngUnit) Then
'        mblnUnload = True: Exit Sub
'    End If
    
    mstrPrivs = gstrPrivs
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "调整(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "撤销(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToRepair, "转修缮(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToEmpty, "转空床(&T)")

    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            '.Add xtpControlButton, conMenu_View_Append, "病区选择(&U)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        'Set objControl = .Add(xtpControlButton, conMenu_View_Column, "选择列(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "调整")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "撤销")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToRepair, "修缮"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_ToEmpty, "空床")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, ""): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
'    With objBar.Controls
'        Set objControl = .Add(xtpControlLabel, 0, "病区 "): objControl.BeginGroup = True
'        Set objCombo = .Add(xtpControlComboBox, conMenu_Edit_SelUnit, "")  '无法显示图标
'            objCombo.DropDownListStyle = True
'            objCombo.Width = 150
'            objCombo.DefaultItem = True
'            objCombo.flags = xtpFlagControlStretched
'    End With
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save '保存
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '撤销
        .Add FCONTROL, vbKeyR, conMenu_Edit_Bed_ToRepair '修缮
        .Add FCONTROL, vbKeyE, conMenu_Edit_Bed_ToEmpty '空床
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'        .AddHiddenCommand conMenu_File_Excel '输出到Excel
'    End With
    '添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub InitDockPannel()
    Dim objPaneType As Pane, objPaneList As Pane, objPaneEdit As Pane
    
    If mfrmEditBed Is Nothing Then Set mfrmEditBed = New frmBedEdit
    
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    'Me.dkpMain.Options.AlphaDockingContext = True
    Me.dkpMain.Options.HideClient = True
    
    
    Set objPaneType = Me.dkpMain.CreatePane(conPane_Type, 200, 600, DockLeftOf)
    objPaneType.Title = "病区列表"
    objPaneType.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
    
    Set objPaneList = Me.dkpMain.CreatePane(conPane_List, 750, 600, DockRightOf, objPaneType)
    objPaneList.Title = "病区床位列表"
    objPaneList.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
    
    Set objPaneEdit = Me.dkpMain.CreatePane(conPane_Edit, 250, 600, DockRightOf)
    objPaneEdit.Title = "病区床位编辑"
    objPaneEdit.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPaneEdit.MaxTrackSize.SetSize 380, 600
    
    objPaneEdit.Close

End Sub

Private Sub InitReportColumn()
'功能：初始化病人列表表格
    Dim objCol As ReportColumn
        
    With rptUnit
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        .AutoColumnSizing = False  '必须在列设置之前设置，才能生效
        Set objCol = .Columns.Add(mCol.图标, "", 18, False): objCol.Editable = False: objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(1, "病区ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(2, "病区", 1200, True): objCol.Editable = False
        .ShowHeader = True
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "没有可显示的病区信息..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With

    With rptList
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        .AutoColumnSizing = False   '必须在列设置之前设置，才能生效
        Set objCol = .Columns.Add(mCol.图标, "", 18, False):  objCol.Groupable = False: objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(mCol.床号, "床号", 50, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.科室, "科室", 100, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.房间号, "房间号", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.顺序号, "顺序号", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.状态, "状态", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.性别分类, "性别分类", 60, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.等级, "等级", 140, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.床位编制, "床位编制", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.单价, "单价", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.姓名, "姓名", 50, True):  objCol.Groupable = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.性别, "性别", 50, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.年龄, "年龄", 30, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.住院状态, "住院状态", 100, True):  objCol.Groupable = True: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.病区ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.科室ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.病人ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.等级ID, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
        Set objCol = .Columns.Add(mCol.共用, "", 0, False):  objCol.Groupable = False: objCol.Visible = False: objCol.AutoSize = False
    
        .ShowHeader = True
        .ShowGroupBox = True
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
    
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的床位信息..."
        End With
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    Dim objItem As TaskPanelGroupItem
    
    If Me.rptList.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    '表头
    objPrint.Title.Text = "病区床位清单"
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人：" & UserInfo.姓名)
    Call objAppRow.Add("打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow

            Call objAppRow.Add("病区:" & ZLCommFun.GetNeedName(rptUnit.FocusedRow.Record(2).Value))  'cboUnit.Text

    Call objPrint.UnderAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
   
End Sub

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand
    For Each rptRow In rptList.Rows
        If rptRow.Childs.Count > 0 Then rptRow.Expanded = True
    Next
    If rptList.Rows.Count < 1 Then zlReportToVSFlexGrid = False: Exit Function
        
    With vfgList
        .Clear
        .Rows = 1: .FixedRows = 1: .RowHeight(.Rows - 1) = 280
        .Cols = 0
        .MergeCells = flexMergeFree
        
        '标题行复制
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = rptCol.Caption
                .ColData(.Cols - 1) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .colAlignment(.Cols - 1) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .colAlignment(.Cols - 1) = flexAlignCenterCenter
                Case xtpAlignmentRight: .colAlignment(.Cols - 1) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, .Cols - 1, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.width < 20 * IIf(rptList.GroupsOrder.Count = 0, 1, rptList.GroupsOrder.Count) Then
                    .ColWidth(.Cols - 1) = 0
                Else
                    .ColWidth(.Cols - 1) = rptCol.width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '数据行复制
        Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: .RowHeight(.Rows - 1) = 280
            If rptRow.GroupRow Then
                intTiers = 0
                Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1
                    Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                .MergeRow(.Rows - 1) = True
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).Value
                Next
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0
                        Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1
                            Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptRow.Record(.ColData(lngCol)).Value
                    Else
                        .TextMatrix(.Rows - 1, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, lngCol, .Rows - 1) = .colAlignment(lngCol)
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Private Sub Form_Resize()
    Dim panType As Pane
    Dim panEdit As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panType = Me.dkpMain.FindPane(conPane_Type)
    Set panEdit = Me.dkpMain.FindPane(conPane_Edit)
    panType.MinTrackSize.SetSize 150, 600
    panType.MaxTrackSize.SetSize 200, 375
    panEdit.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265
    panEdit.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
'
'    panEdit.MinTrackSize.SetSize 0, 0
'    panEdit.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 375
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmEditBed Is Nothing Then
        Unload mfrmEditBed
        If mfrmEditBed.mintCancle = 1 Then Cancel = 1: Exit Sub
        Set mfrmEditBed = Nothing
    End If
    mblnUnload = False
    mintEditState = 0
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
'    mfrmEditBed.fraEdit.Width = mfrmEditBed.Width
'    mfrmEditBed.fraEdit.Height = Me.ScaleHeight - Me.picList.ScaleHeight
End Sub

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 1 To img32.ListImages.Count
        If Not img32.ListImages(i).Key Like "MASK_*" Then
            img32.ListImages.Add , "加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_加床", i)
            img32.ListImages.Add , "非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_非编", i)
            img32.ListImages.Add , "共用_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用", i)
            img32.ListImages.Add , "共用_加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_加床", i)
            img32.ListImages.Add , "共用_非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_非编", i)
        End If
    Next
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院科室
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim i As Integer, lngUnitID As Long, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    Dim intCurrIndes As Integer
    
    On Error GoTo errH
    
    'Set objCombo = cbsMain(cbsMain.Count).FindControl(, conMenu_Edit_SelUnit, True)
    '包含门诊观察室
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    '问题30922 by lesfeng 2010-06-18 b
    If blnLimitUnit Then strUnitIDs = UserInfo.ID
    'by lesfeng 2010-1-8 性能优化
    gstrSQL = _
        " Select A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & IIf(blnLimitUnit, ",部门人员 C ", "") & _
        " Where B.部门ID = A.ID" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
        IIf(blnLimitUnit, " And A.ID = C.部门ID And C.人员ID In ([1])", "") & _
        " And (A.站点=[2] Or A.站点 is Null)" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    
    '问题30922 by lesfeng 2010-06-18 e
    If Not rsTmp.EOF Then
        Me.rptUnit.Records.DeleteAll

        Do While Not rsTmp.EOF
            '填入数据
            Set objRecord = rptUnit.Records.Add()
            Set objItem = objRecord.AddItem(""): objItem.Icon = 35
            objRecord.AddItem Val("" & rsTmp!ID)
            objRecord.AddItem Nvl(rsTmp!名称)    'Nvl(rsTmp!编码) & "-" &
            rsTmp.MoveNext
        Loop
        With Me.rptUnit
            .Populate
        End With
        If Me.rptUnit.Records.Count - 1 > 0 Then Me.rptUnit.FocusedRow = rptUnit.Rows(IIf(rptUnit.Rows(1).GroupRow, 1, 0))
    ElseIf InStr(";" & mstrPrivs, "所有病区") > 0 Then
        MsgBox "没有设置病区,请你先到部门管理中设置工作性质为护理的部门！", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "你没有 [所有病区] 的权限,并且你所在部门不是病区！", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBeds(lngUnitID As Long) As Boolean
    '功能：读取指定病区的床位列表
    Dim i As Integer, j As Integer
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim intBedLen As Integer
    Dim mrsBeds As ADODB.Recordset
    Dim str价格等级 As String
    
    intHolding = 0: intRemedy = 0: mintEmpty = 0
    
    '清空统计数据
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID)
    
    gstrSQL = "Select a.价格等级" & vbNewLine & _
            "  From 收费价格等级应用 A,收费价格等级 B" & vbNewLine & _
            "  where A.价格等级=b.名称  And a.性质=0 And b.是否适用普通项目=1 and a.站点=[1]" & vbNewLine & _
            "        and nvl(b.撤档时间,sysdate+1)>sysdate"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    If mrsBeds.RecordCount > 0 Then
        str价格等级 = mrsBeds!价格等级 & ""
    End If
    gstrSQL = "Select a.*" & vbNewLine & _
            "From (Select LPad(a.床号, [1], ' ') 床号, a.病区id, a.房间号,a.顺序号,a.性别分类, a.床位编制, a.等级id, a.状态, a.病人id, a.共用, e.现价 As 价格," & vbNewLine & _
            "              Nvl(b.名称, Decode(a.共用, 1, '<共用病床>', Null)) As 科室, a.科室id, c.名称 As 等级, d.姓名, d.性别, d.年龄," & vbNewLine & _
            "              Decode(d.状态, 0, '正常住院', 2, '准备转出', 3, '准备出院' || '(' || To_Char(f.开始时间, 'YYYY-MM-DD HH24:MI:SS') || ')') As 住院状态," & vbNewLine & _
            "              Row_Number() Over(Partition By a.床号, e.收费细目id Order By Decode(e.价格等级, [3], 1, Null, 2, 3)) As Top" & vbNewLine & _
            "       From 床位状况记录 A, 部门表 B, 收费项目目录 C," & vbNewLine & _
            "            (Select m.病人id, Nvl(n.姓名, m.姓名) 姓名, Nvl(n.性别, m.性别) 性别, Nvl(n.年龄, m.年龄) 年龄, n.状态" & vbNewLine & _
            "              From 病人信息 M, 病案主页 N" & vbNewLine & _
            "              Where m.病人id = n.病人id And m.主页id = n.主页id And n.当前病区id = [2]) D, 收费价目 E," & vbNewLine & _
            "            (Select q.病人id, 开始时间" & vbNewLine & _
            "              From 病案主页 P, 病人变动记录 Q" & vbNewLine & _
            "              Where p.病人id = q.病人id And p.主页id = q.主页id And 开始原因 = 10 And p.当前病区id = [2] And p.状态 = 3) F" & vbNewLine & _
            "       Where a.科室id = b.Id(+) And a.等级id = c.Id(+) And a.病人id = d.病人id(+) And e.收费细目id = c.Id And a.病人id = f.病人id(+) And" & vbNewLine & _
            "             a.病区id = [2] And Sysdate Between e.执行日期 And Nvl(e.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             Nvl(e.价格等级, [3]) = [3]" & vbNewLine & _
            "       Order By a.顺序号,LPad(a.床号, [1], ' ')) A" & vbNewLine & _
            "Where Top = 1"

    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intBedLen, lngUnitID, IIf(str价格等级 = "", "空", str价格等级))
    
    Me.rptList.Records.DeleteAll
    Me.rptList.SortOrder.DeleteAll  '每次清空排序默认按床号升序排列
    Do While Not mrsBeds.EOF
        '填入数据
        Set objRecord = rptList.Records.Add()
        Select Case mrsBeds!状态
            Case "空床"
                If mrsBeds!性别分类 = "男床" Then
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
                ElseIf mrsBeds!性别分类 = "女床" Then
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 2
                Else
                    Set objItem = objRecord.AddItem(""): objItem.Icon = 0
                End If
                mintEmpty = mintEmpty + 1
            Case "占用"
                Set objItem = objRecord.AddItem(""): objItem.Icon = 3
                intHolding = intHolding + 1
            Case "修缮"
                Set objItem = objRecord.AddItem(""): objItem.Icon = 4
                intRemedy = intRemedy + 1
            Case Else   '当作修缮
                Set objItem = objRecord.AddItem(""): objItem.Icon = 4
                intRemedy = intRemedy + 1
        End Select
        
        objRecord.AddItem (Trim(mrsBeds!床号))
        objRecord.AddItem (Nvl(mrsBeds!科室))
        objRecord.AddItem (Nvl(mrsBeds!房间号))
        objRecord.AddItem (Nvl(mrsBeds!顺序号))
        objRecord.AddItem (Nvl(mrsBeds!状态))
        objRecord.AddItem (Nvl(mrsBeds!性别分类))
        objRecord.AddItem (Nvl(mrsBeds!等级))
        objRecord.AddItem (Nvl(mrsBeds!床位编制))
        objRecord.AddItem Format((Nvl(mrsBeds!价格)), "0.00")
        objRecord.AddItem (Nvl(mrsBeds!姓名))
        objRecord.AddItem (Nvl(mrsBeds!性别))
        objRecord.AddItem (Nvl(mrsBeds!年龄))
        objRecord.AddItem (Nvl(mrsBeds!住院状态))
        objRecord.AddItem (Nvl(mrsBeds!病区ID))
        objRecord.AddItem (Nvl(mrsBeds!科室ID))
        objRecord.AddItem (Nvl(mrsBeds!病人ID))
        objRecord.AddItem (Nvl(mrsBeds!等级ID))
        objRecord.AddItem (Nvl(mrsBeds!共用))

        If objRecord.Item(mCol.床位编制).Value = "加床" Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "加床_" & img16.ListImages(objRecord.Item(mCol.图标).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.图标).Icon = i - 1
        ElseIf objRecord.Item(mCol.床位编制).Value = "非编" Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "非编_" & img16.ListImages(objRecord.Item(mCol.图标).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.图标).Icon = i - 1
        End If
        If Val(objRecord(mCol.共用).Value) <> 0 Then
            For i = 1 To img16.ListImages.Count
                If img16.ListImages(i).Key = "共用_" & img16.ListImages(objRecord.Item(mCol.图标).Icon + 1).Key Then
                    Exit For
                End If
            Next
            objRecord.Item(mCol.图标).Icon = i - 1
        End If
        mrsBeds.MoveNext
    Loop
    With Me.rptList
        .Populate
    End With
    If Me.rptList.Records.Count - 1 > 0 Then Me.rptList.FocusedRow = rptList.Rows(IIf(rptList.Rows(1).GroupRow, 1, 0))
    Call SetBedNOLen(lngUnitID)
    ReadBeds = True
    stbThis.Panels(2) = "当前病区共 " & rptList.Records.Count & " 张病床,其中病人占用 " & intHolding & " 张,空床 " & mintEmpty & " 张,正在修缮 " & intRemedy & " 张！"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowEdit(blnShow As Boolean)
    '功能       是否显示编辑窗体
    Dim objPane As Pane
    Set objPane = dkpMain.FindPane(conPane_Edit)
    If blnShow = True Then
        objPane.Select
    Else
        objPane.Close
    End If
    dkpMain.RecalcLayout
End Sub

Public Function zlRefList(Optional strBedNO As String) As Long
    '功能：刷新装入床位清单，并定位到指定的床位上
    Dim objCombo As CommandBarComboBox
    Dim rptRow As ReportRow
    Dim rptParent As ReportRow
    

    If ReadBeds(mlngUnit) Then
        If strBedNO <> "" Then
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow = False Then
                    If Trim(rptRow.Record(mCol.床号).Value) = strBedNO Then
                        Set rptParent = rptRow.ParentRow
                        Set Me.rptList.FocusedRow = rptRow
                        Exit For
                    End If
                End If
            Next
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow Then
                    If Not (rptRow Is rptParent) Then rptRow.Expanded = False
                End If
            Next
            Set Me.rptList.FocusedRow = Me.rptList.FocusedRow
        Else
            For Each rptRow In Me.rptList.Rows
                If rptRow.GroupRow Then rptRow.Expanded = True
            Next
        End If
        If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Public Sub SetBedNOLen(ByVal lngUnitID As Long)
    Dim bytLen As Byte, i As Integer
    If rptList.Records.Count = 0 Then Exit Sub
    
    bytLen = GetMaxBedLen(lngUnitID)
    For i = 0 To rptList.Records.Count - 1
        rptList.Records(i).Item(mCol.床号).Value = Space(bytLen - Len(CStr(rptList.Records(i).Item(mCol.床号).Value))) & Trim(rptList.Records(i).Item(mCol.床号).Value)
    Next
End Sub

Private Sub picUnit_Resize()
    Err = 0: On Error Resume Next
    With Me.rptUnit
        .Left = Me.picUnit.ScaleLeft: .width = Me.picUnit.ScaleWidth - .Left
        .Top = Me.picUnit.ScaleTop: .Height = Me.picUnit.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    If Me.cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls(2)
    Set objBar = Me.cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each objControl In objPopup.CommandBar.Controls
        Set objControl = objBar.Controls.Add(xtpControlButton, objControl.ID, objControl.Caption)
        objControl.BeginGroup = objControl.BeginGroup
    Next
    objBar.ShowPopup
End Sub

Private Sub rptUnit_SelectionChanged()
    If rptUnit.FocusedRow Is Nothing Then Exit Sub
    mlngUnit = rptUnit.FocusedRow.Record(1).Value
    
    If Not ReadBeds(mlngUnit) Then
        mblnUnload = True: Exit Sub
    End If
    Me.Refresh
End Sub

