VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisStationAuditing 
   Caption         =   "报告批量审核"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmLisStationAuditing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8670
   Begin MSComctlLib.ListView lvwApparatus 
      Height          =   285
      Left            =   1650
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "仪器"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbSave 
      Height          =   165
      Left            =   1440
      TabIndex        =   27
      Top             =   5625
      Visible         =   0   'False
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6570
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":28E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":2B00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7245
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":3C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":438E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAuditing.frx":45AE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   2070
      Left            =   3645
      TabIndex        =   25
      Top             =   2100
      Width           =   3450
      _cx             =   6085
      _cy             =   3651
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
      BackColorSel    =   16768667
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
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
      Editable        =   2
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8670
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选"
               Object.Tag             =   "&A.全选"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清"
               Object.Tag             =   "&R.全清"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&D.审核"
               Key             =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "&D.审核"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&E.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "&E.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   5520
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLisStationAuditing.frx":47CE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10213
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1350
      Left            =   3615
      TabIndex        =   15
      Top             =   720
      Width           =   2880
      _cx             =   5080
      _cy             =   2381
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
      BackColorSel    =   16768667
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
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
      Editable        =   2
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fra 
      Height          =   4695
      Left            =   90
      TabIndex        =   19
      Top             =   750
      Width           =   3120
      Begin VB.CommandButton CmdApparatus 
         Caption         =   "&P"
         Height          =   255
         Left            =   2730
         TabIndex        =   29
         Top             =   2850
         Width           =   255
      End
      Begin VB.TextBox TxtApparatus 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2820
         Width           =   2715
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "重置条件(&J)"
         Height          =   350
         Left            =   1620
         TabIndex        =   14
         Top             =   3795
         Width           =   1185
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3420
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2220
         Width           =   2715
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "立即搜索(&S)"
         Height          =   350
         Left            =   300
         TabIndex        =   13
         Top             =   3795
         Width           =   1185
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2820
         Width           =   2715
      End
      Begin VB.TextBox Txt 
         Height          =   300
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   1635
         Width           =   2715
      End
      Begin VB.TextBox Txt 
         Height          =   300
         Index           =   0
         Left            =   300
         TabIndex        =   4
         ToolTipText     =   "标本号以,分隔、以～指定范围"
         Top             =   1020
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   25493507
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1785
         TabIndex        =   2
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   25493507
         CurrentDate     =   38229
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   1590
         TabIndex        =   26
         Top             =   480
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.病人科室"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   3195
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.检验科室"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.检 验 人"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.标本号码"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   795
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.标本时间"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.检验仪器"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2595
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&E"
      Height          =   350
      Index           =   4
      Left            =   405
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&H"
      Height          =   350
      Index           =   3
      Left            =   405
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&D"
      Height          =   350
      Index           =   2
      Left            =   405
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&R"
      Height          =   350
      Index           =   1
      Left            =   405
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&A"
      Height          =   350
      Index           =   0
      Left            =   405
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1100
   End
   Begin VB.Image imgX 
      Height          =   45
      Left            =   3240
      MousePointer    =   7  'Size N S
      Top             =   2220
      Width           =   2595
   End
End
Attribute VB_Name = "frmLisStationAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
Private mblnChangeEdit As Boolean
Private mstrPrivs As String
Private mlngDeptID As Long
Private mstrAuditingMan As String                        '权限人
Private mintAuditing As String                           '时间限制
Private mDataAuditing As Date                            '限制开始时间

Private Enum mCol
    选择 = 0
    急诊
    标本号
    标本类型
    核收时间
    核收人
    检验人
    申请时间
    申请人
    申请科室
    检验仪器
    执行科室
    医嘱id
    发送号
    病人ID
    标本ID
End Enum

Private Function RefreshData(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vsfDetail.Rows = 2
    vsfDetail.Cell(flexcpText, 1, 0, 1, vsfDetail.Cols - 1) = ""
    
    mstrSQL = "SELECT D.ID,E.病人ID,D.医嘱id,F.发送号,D.报告结果,ROWNUM AS 序号,B.中文名 AS 检验项目,C.缩写 AS 英文缩写,A.检验结果,A.结果标志,A.结果参考 " & _
                "FROM 检验普通结果 A,诊治所见项目 B,检验项目 C,检验标本记录 D,病人医嘱记录 E,病人医嘱发送 F,检验报告项目 G " & _
                "WHERE A.检验项目id = B.ID " & _
                    "AND B.ID = C.诊治项目ID " & _
                    "AND A.记录类型 =D.报告结果 " & _
                    "AND D.ID=A.检验标本ID " & _
                    "AND D.医嘱ID=E.相关ID " & _
                    "AND G.诊疗项目ID=E.诊疗项目ID+0 AND B.ID=G.报告项目ID(+) " & _
                    "AND E.相关ID=F.医嘱ID AND F.执行状态=3 " & _
                    "AND D.ID=[1] Order By Nvl(G.诊疗项目ID,0),Nvl(G.排列序号,99)"
                        
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)

    If rs.BOF = False Then
        vsfDetail.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsfDetail, rs)
        vsfDetail.TextMatrix(0, 0) = ""
    End If
    
    RefreshData = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '功能:根据修改状态设置按钮、菜单等的可用状态
    '-----------------------------------------------------------------------------------------
    cmd(2).Enabled = True
        
    If mblnChangeEdit = False Then cmd(2).Enabled = False
        
    tbrThis.Buttons("审核").Enabled = cmd(2).Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    If vsf.Rows = 2 And Trim(vsf.TextMatrix(1, 1)) = "" Then
        stbThis.Panels(2).Text = "没有标本信息。"
    Else
        stbThis.Panels(2).Text = "共找到 " & vsf.Rows - 1 & " 个标本信息。"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, Optional ByVal lngDeptID As Long = 0, Optional ByVal strPrivs As String, _
                         Optional ByVal strAuditingMan As String, Optional ByVal intAuditing As Integer, _
                         Optional ByVal DataAuditing As Date) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示本编辑窗体
    '参数：             lngdeptid = 科室ID strprivs = 权限 strAuditingMan = 申请人 intAuditing = 时间限制
    '                   DataAuditing = 开始时间
    '返回：
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
'    mstrPrivs = strPrivs
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    mstrAuditingMan = strAuditingMan
    mintAuditing = intAuditing
    mDataAuditing = DataAuditing
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vsf.Cols = 0
    Call NewColumn(vsf, "选择", 510, 4)
    Call NewColumn(vsf, "急诊", 500, 1)
    Call NewColumn(vsf, "标本号", 750, 1)
    Call NewColumn(vsf, "标本类型", 900, 1)
    Call NewColumn(vsf, "核收时间", 1080, 1)
    Call NewColumn(vsf, "检验人", 750, 1)
    Call NewColumn(vsf, "申请人", 750, 1)
    Call NewColumn(vsf, "申请时间", 1080, 1)
    Call NewColumn(vsf, "申请科室", 1200, 1)
    Call NewColumn(vsf, "核收人", 750, 1)
    Call NewColumn(vsf, "检验仪器", 1200, 1)
    Call NewColumn(vsf, "执行科室", 1200, 1)
    Call NewColumn(vsf, "医嘱id", 0, 1)
    Call NewColumn(vsf, "发送号", 0, 1)
    Call NewColumn(vsf, "病人ID", 0, 1)
    Call NewColumn(vsf, "标本ID", 0, 1)
    vsf.ColDataType(mCol.选择) = flexDTBoolean
    
    vsfDetail.Cols = 0
    Call NewColumn(vsfDetail, "", 240, 4)
    Call NewColumn(vsfDetail, "检验项目", 2100, 1)
    Call NewColumn(vsfDetail, "英文缩写", 900, 1)
    Call NewColumn(vsfDetail, "检验结果", 1200, 1)
    Call NewColumn(vsfDetail, "结果标志", 900, 1)
    Call NewColumn(vsfDetail, "结果参考", 1800, 1)
    vsfDetail.FixedCols = 1
    
        
    InitData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngLoop As Long
    Dim varItem As Variant                          '分解","号
    Dim varBetween As Variant                       '分解"~"
    
    
    On Error GoTo ErrHand
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        
    strWhere = " AND A.核收时间 BETWEEN TO_DATE('" & Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')" ' AND B.执行科室ID + 0 = " & cbo(1).ItemData(cbo(1).ListIndex)
    
    If Trim(Txt(2).Text) <> "" Then strWhere = strWhere & " AND A.检验人 = '" & Trim(Txt(2).Text) & "'"
    
'    If cbo(0).ListIndex > 0 Then strWhere = strWhere & _
'        IIf(cbo(0).ListIndex = 1, " AND A.仪器id IS Null", " AND A.仪器id=" & cbo(0).ItemData(cbo(0).ListIndex))
        
    If InStr(1, "," & Me.TxtApparatus.Tag & ",", ",A,") <= 0 Then
        If InStr(1, "," & Me.TxtApparatus.Tag & ",", ",B,") > 0 Then
            strWhere = strWhere & " AND A.仪器id IS Null"
        Else
            strWhere = strWhere & " AND A.仪器id in (" & Me.TxtApparatus.Tag & ")"
        End If
    End If
    If cbo(1).ListIndex > 0 Then strWhere = strWhere & " AND A.执行科室ID + 0=" & cbo(1).ItemData(cbo(1).ListIndex)
    
    If Trim(Txt(0).Text) <> "" Then
'        varTmp2 = Split(Trim(txt(0).Text), ",")
'        strTmp = ""
'        For mlngLoop = 0 To UBound(varTmp2)
'            If InStr(varTmp2(mlngLoop), "-") = 0 Then
'                strTmp = strTmp & "  OR A.标本序号=" & TransSampleNO(varTmp2(mlngLoop))
'            Else
'                strTmp = strTmp & "  OR A.标本序号 BETWEEN " & TransSampleNO(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "~") - 1)) & " AND " & TransSampleNO(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "~") + 1))
'            End If
'        Next
'        If strTmp <> "" Then strWhere = strWhere & " AND (1=2 " & strTmp & ")"
        'strWhere = strWhere & " AND A.标本序号 BETWEEN '" & txt(0).Text & "' AND '" & txt(0).Text & "'"
        varItem = Split(Trim(Txt(0).Text), ",")
        For mlngLoop = 0 To UBound(varItem)
            varBetween = Split(varItem(mlngLoop), "~")
            If UBound(varBetween) > 0 Then
                strTmp = strTmp & "  OR A.标本序号 BETWEEN " & TransSampleNO(varBetween(0)) & " AND " & TransSampleNO(varBetween(1))
            Else
                strTmp = strTmp & " OR A.标本序号=" & TransSampleNO(varItem(mlngLoop))
            End If
        Next
        If strTmp <> "" Then strWhere = strWhere & " AND (1=2 " & strTmp & ")"
    End If
        
    mstrSQL = "select DISTINCT A.ID,A.医嘱id,F.发送号,0 AS 选择," & _
                      " Decode(A.仪器id, Null, " & vbCrLf & _
                        " to_Char(Trunc(A.标本序号/10000)+1,'0000')|| '-'||to_Char(MOD(A.标本序号,10000),'0000'), A.标本序号) As 标本号, " & _
                      "A.标本类型," & _
                      "TO_CHAR(A.核收时间,'MM-DD HH24:MI') AS 核收时间," & _
                      "A.核收人," & _
                      "A.检验人," & _
                      "TO_CHAR(B.开嘱时间,'MM-DD HH24:MI') AS 申请时间," & _
                      "B.开嘱医生 AS 申请人," & _
                      "C.名称 AS 申请科室," & _
                      "E.名称 AS 执行科室," & _
                      "B.病人ID," & _
                      "A.ID as 标本ID, " & _
                      "D.名称 AS 检验仪器,Decode(A.标本类别,1,'√','') As 急诊 " & _
                 "from 检验标本记录 A, 病人医嘱记录 B, 部门表 C, 检验仪器 D,部门表 E,病人医嘱发送 F " & _
                "WHERE A.医嘱ID = B.相关ID AND C.ID = B.开嘱科室ID AND B.ID=F.医嘱id AND F.执行状态=3 AND " & _
                      "A.仪器ID = D.ID(+) AND E.ID=B.执行科室id AND A.样本状态 = 1 " & strWhere & _
                " ORDER BY 标本号 "

    Call OpenRecord(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
    End If
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    
    '检验是否允许审核,即是否满足审核的条件
    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 Then
            If CheckIsAllowAuditing(Me, Val(vsf.RowData(mlngLoop))) = False Then
                vsf.TextMatrix(mlngLoop, mCol.选择) = 0
            End If
            
            If InStr(1, mstrPrivs, "审核限制") > 0 And vsf.TextMatrix(mlngLoop, mCol.检验人) = UserInfo.姓名 Then
                '没有登陆审核人
                If mintAuditing = 0 Then
                    '同一病人被权限控制不能进行审核
                    vsf.TextMatrix(mlngLoop, mCol.选择) = 0
                End If
                '审核时间是否过期
                If mintAuditing < 0 Then
                    If DateDiff("h", mDataAuditing, Now) > Abs(mintAuditing) Then
                        vsf.TextMatrix(mlngLoop, mCol.选择) = 0
                    End If
                End If
                
                '判断登陆时后的审核人是否为同一人.
                If vsf.TextMatrix(mlngLoop, mCol.检验人) = mstrAuditingMan Then
                    '登陆进入的审核人和当前用户为同一个人
                    vsf.TextMatrix(mlngLoop, mCol.选择) = 0
                End If
                If vsf.TextMatrix(mlngLoop, mCol.选择) = 0 Then
                    MsgBox "没有登陆审核人，请登陆审核人后再试！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    ValidData = True
    
    Exit Function
ErrHand:
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blnAutoPrint As Boolean
'    Dim strSQL() As String
    Dim strsql As String
    
    On Error GoTo ErrHand
    
'    ReDim strSQL(1 To 1)
    
    With pgbSave
        .Visible = True
        .Min = 0: .Max = vsf.Rows - 1
        .Value = 0
    End With
    
    zlcommfun.ShowFlash "正在审核请稍候..."
    
    blnAutoPrint = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "审核打印", 0))
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    For mlngLoop = 1 To vsf.Rows - 1
        pgbSave.Value = mlngLoop
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 And Val(vsf.RowData(mlngLoop)) > 0 Then
            
            If CheckChargeState(Val(vsf.RowData(mlngLoop)), False) = False Then
                '未收费
                If InStr(mstrPrivs, "未收费审核") > 0 Then
'                    strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_报告审核(" & Val(Vsf.RowData(mlngLoop)) & ",'" & _
'                                                 IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "')"
                    strsql = "ZL_检验标本记录_报告审核(" & Val(vsf.RowData(mlngLoop)) & ",'" & _
                                                 IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "')"
                End If
            Else
'                strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_报告审核(" & Val(Vsf.RowData(mlngLoop)) & ",'" & _
'                                                 IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "')"
                strsql = "ZL_检验标本记录_报告审核(" & Val(vsf.RowData(mlngLoop)) & ",'" & _
                                                 IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "')"
            End If
            
            zlDatabase.ExecuteProcedure strsql, gstrSysName
            
            If blnAutoPrint Then
                If GetReportCode(Val(vsf.TextMatrix(mlngLoop, mCol.医嘱id)), Val(vsf.TextMatrix(mlngLoop, mCol.发送号)), strReportCode, strReportParaNo, bytReportParaMode, _
                     False) Then
                    Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, _
                        "医嘱ID=" & Val(vsf.TextMatrix(mlngLoop, mCol.医嘱id)), _
                        "病人ID=" & Val(vsf.TextMatrix(mlngLoop, mCol.病人ID)), _
                        "标本ID=" & Val(vsf.TextMatrix(mlngLoop, mCol.标本ID)), 2)
                End If
            End If
        End If
        DoEvents
    Next
    pgbSave.Visible = False
    blnTran = True
    
    zlcommfun.StopFlash
'    gcnOracle.BeginTrans
'    For mlngLoop = 1 To UBound(strSQL)
'        If strSQL(mlngLoop) <> "" Then
'            Call ExecuteProc(strSQL(mlngLoop), Me.Caption)
'        End If
'    Next
'    gcnOracle.CommitTrans
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    zlcommfun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
'    If blnTran Then gcnOracle.RollbackTrans
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case 0
        For mlngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(mlngLoop)) > 0 Then
                vsf.TextMatrix(mlngLoop, mCol.选择) = 1
            End If
        Next
        
        mblnChangeEdit = True
        Call AdjustEnableState
    Case 1
        For mlngLoop = 1 To vsf.Rows - 1
            vsf.TextMatrix(mlngLoop, mCol.选择) = 0
        Next
        mblnChangeEdit = False
        Call AdjustEnableState
    Case 2
        If mblnChangeEdit Then
        
            If ValidData = False Then Exit Sub
            If SaveData() = False Then Exit Sub
            
            mblnOK = True
            
            mblnChangeEdit = False
            Call AdjustEnableState

            Unload Me
            Exit Sub
        End If
        
    Case 3
        ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
    Case 4
        Unload Me
    End Select
End Sub

Private Sub CmdApparatus_Click()
    With Me.lvwApparatus
        .Top = Me.TxtApparatus.Top + Me.TxtApparatus.Height + 600
        .Left = Me.TxtApparatus.Left
        .Height = fra.Height - (Me.TxtApparatus.Top + Me.TxtApparatus.Height)
        .Width = Me.TxtApparatus.Width
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub cmdRefresh_Click()
    
    Call ReadData
    
    Call RefreshData(vsf.RowData(vsf.Row))
    
    mblnChangeEdit = False
    Call AdjustEnableState
    Call RefreshStatus
    
    vsf.Col = 1
    vsf.SetFocus
    vsf.Col = 0
End Sub

Private Sub cmdReset_Click()
    Dim ControlcboDept As CommandBarComboBox
    
    dtp(0).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtp(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    
'    cbo(0).ListIndex = 0
'    cbo(2).ListIndex = 0
    
'    cbo(1).ListIndex = mfrmMain.cboDept.ListIndex - 1
    If cbo(1).ListIndex = -1 Then
        zlControl.CboLocate cbo(1), UserInfo.部门ID, True
        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    End If
    
    Txt(0).Text = ""
    Txt(2).Text = ""
    
'    dtp(0).SetFocus
    Me.TxtApparatus.SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim lngDefaultDev As Long, mlngLoop As Long
    Dim ItmX As ListItem
    Dim ControlcboDept As CommandBarComboBox
    Dim strsql As String
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    '检验部门
    cbo(1).Clear
    strsql = "select A.编码||'-'||A.名称,a.ID from 部门表 a where a.id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngDeptID)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs, False)
    
    If cbo(1).ListIndex = -1 Then
        zlControl.CboLocate cbo(1), UserInfo.部门ID, True
        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    End If
    
    '病人科室
    mstrSQL = "SELECT A.编码||'-'||A.名称,ID FROM 部门表 A,部门性质说明 B WHERE A.ID=B.部门id AND B.工作性质='临床' ORDER BY A.编码||'-'||A.名称"
    Call OpenRecord(rs, mstrSQL, Me.Caption)
    cbo(2).AddItem "所有科室"
    If rs.BOF = False Then Call AddComboData(cbo(2), rs, False)
    zlControl.CboLocate cbo(2), UserInfo.部门ID, True
    If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    
    '检验仪器
    mstrSQL = "SELECT A.编码||'-'||A.名称,ID FROM 检验仪器 A where 使用小组id = [1] ORDER BY A.编码||'-'||A.名称"
    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngDeptID)
    cbo(0).AddItem "手工"
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    lngDefaultDev = Val(Split(GetConnectDevs & ";1", ";")(0))
    cbo(0).ListIndex = FindComboItem(cbo(0), lngDefaultDev)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    Me.TxtApparatus.Tag = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "仪器选择_ID", "A")
    Me.TxtApparatus.Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "仪器选择_Name", "所有仪器")
    
    On Error GoTo errH
    mstrSQL = "SELECT A.编码,A.名称,ID FROM 检验仪器 A where 使用小组id = [1] ORDER BY A.编码||'-'||A.名称"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngDeptID)
    With Me.lvwApparatus
        Set ItmX = .ListItems.Add(, "B", "")
        ItmX.SubItems(1) = "手工"
        Do Until rs.EOF
            Set ItmX = .ListItems.Add(, "A" & rs("ID"), rs("编码"))
            ItmX.SubItems(1) = rs("名称")
            If InStr(1, "," & Me.TxtApparatus.Tag & ",", "," & rs("ID") & ",") > 0 Then
                ItmX.Checked = True
            End If
            rs.MoveNext
        Loop
    End With
    
    
    
    Call cmdReset_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra
        .Left = 0
        .Top = cbrThis.Height - 90
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    With vsf
        .Left = fra.Left + fra.Width
        .Top = cbrThis.Height
        .Width = Me.ScaleWidth - .Left
        .Height = imgX.Top - .Top
    End With
    
    With imgX
        .Left = vsf.Left
        .Width = vsf.Width
    End With
    
    With vsfDetail
        .Left = vsf.Left
        .Top = imgX.Top + imgX.Height
        .Width = vsf.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    If mblnChangeEdit Then
'        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
'        If Cancel Then Exit Sub
'    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "仪器选择_ID", Me.TxtApparatus.Tag)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "仪器选择_Name", Me.TxtApparatus.Text)
    
    Call SaveWinState(Me, App.ProductName)
    
    
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    imgX.Top = imgX.Top + y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000

    Form_Resize
End Sub

Private Sub lvwApparatus_DblClick()
    lvwApparatus.Visible = False
    GetSelcetlvw
    Me.TxtApparatus.SetFocus
End Sub

Private Sub lvwApparatus_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    If Item.Key = "A" Or Item.Key = "B" Then
        For i = 1 To Me.lvwApparatus.ListItems.Count
            Me.lvwApparatus.ListItems(i).Checked = False
        Next
        Item.Checked = True
    Else
'        Me.lvwApparatus.ListItems("A").Checked = False
        Me.lvwApparatus.ListItems("B").Checked = False
    End If
End Sub

Private Sub lvwApparatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lvwApparatus.Visible = False
        GetSelcetlvw
        Me.TxtApparatus.SetFocus
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "全选"
        Call cmd_Click(0)
    Case "全清"
        Call cmd_Click(1)
    Case "审核"
        Call cmd_Click(2)
    Case "帮助"
        Call cmd_Click(3)
    Case "退出"
        Call cmd_Click(4)
    End Select
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index = 2 Then zlcommfun.OpenIme True
    
    zlControl.TxtSelAll Txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    Else
        Select Case Index
        Case 0
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789,-~")
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 2 Then zlcommfun.OpenIme False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(Txt(Index).Text, Txt(Index).MaxLength)
End Sub

Private Sub TxtApparatus_GotFocus()
    Me.TxtApparatus.SelStart = 0
    Me.TxtApparatus.SelLength = Len(Me.TxtApparatus)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
    Call AdjustEnableState
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If NewRow + 1 > vsf.FixedRows And OldRow + 1 > vsf.FixedRows Then
        vsf.Cell(flexcpBackColor, OldRow, 0, OldRow, vsf.Cols - 1) = vsf.BackColor
        vsf.Cell(flexcpBackColor, NewRow, 0, NewRow, vsf.Cols - 1) = vsf.BackColorSel
    End If
    
    If NewRow <> OldRow Then
        Call RefreshData(vsf.RowData(NewRow))
    End If
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If NewRow + 1 > vsfDetail.FixedRows And OldRow + 1 > vsfDetail.FixedRows Then
        vsfDetail.Cell(flexcpBackColor, OldRow, 1, OldRow, vsfDetail.Cols - 1) = vsfDetail.BackColor
        vsfDetail.Cell(flexcpBackColor, NewRow, 1, NewRow, vsfDetail.Cols - 1) = vsfDetail.BackColorSel
    End If
End Sub

Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub GetSelcetlvw()
    Dim i As Integer
    With Me.lvwApparatus
        Me.TxtApparatus.Tag = ""
        Me.TxtApparatus.Text = ""
        For i = 1 To .ListItems.Count - 1
            If .ListItems(i).Checked = True Then
                If Me.TxtApparatus.Tag = "" Then
                    Me.TxtApparatus.Tag = IIf(.ListItems(i).Key = "A" Or .ListItems(i).Key = "B", .ListItems(i).Key, Mid(.ListItems(i).Key, 2))
                Else
                    Me.TxtApparatus.Tag = Me.TxtApparatus.Tag & "," & IIf(.ListItems(i).Key = "A" Or .ListItems(i).Key = "B", .ListItems(i).Key, Mid(.ListItems(i).Key, 2))
                End If
                
                If Me.TxtApparatus.Text = "" Then
                    Me.TxtApparatus.Text = .ListItems(i).SubItems(1)
                Else
                    Me.TxtApparatus.Text = Me.TxtApparatus.Text & " " & .ListItems(i).SubItems(1)
                End If
            End If
        Next
        If Me.TxtApparatus.Text = "" And Me.TxtApparatus.Tag = "" Then
            If .SelectedItem.Key = "A" Or .SelectedItem.Key = "B" Then
                Me.TxtApparatus.Tag = .SelectedItem.Key
            Else
                Me.TxtApparatus.Tag = Mid(.SelectedItem.Key, 2)
            End If
            Me.TxtApparatus.Text = .SelectedItem.SubItems(1)
        End If
    End With
End Sub
Private Function AuditionCheck() As Boolean
    Dim strVerifyMan As String
    
'    If Not rptList.FocusedRow Is Nothing Then
'        With Me.rptList.FocusedRow
'            strVerifyMan = .Record(mCol.检验人).Value
'        End With
'    End If

    If InStr(1, mstrPrivs, "审核标本") <= 0 Then
        '没有权限和其他用户登陆时退出
        MsgBox "你没有权限进行审核,请重新登陆具有审核人员进行审核!", vbInformation, gstrSysName
        Exit Function
    End If

'    If InStr(1, mstrPrivs, "审核限制") > 0 And strVerifyMan = UserInfo.姓名 Then
    If strVerifyMan = UserInfo.姓名 Then
        '没有登陆审核人
        If mintAuditing = 0 Then
            '同一病人被权限控制不能进行审核
'            MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
            Exit Function
        End If
        '审核时间是否过期
        If mintAuditing < 0 Then
            If DateDiff("h", mDataAuditing, Now) > Abs(mintAuditing) Then
'                MsgBox "审核有效时间已过,请重新登陆审核人!", vbInformation, gstrSysName
                '在有效时间段内可以进行审核
                Exit Function
            End If
        End If
        
        '判断登陆时后的审核人是否为同一人.
        If strVerifyMan = mstrAuditingMan Then
            '登陆进入的审核人和当前用户为同一个人
'            MsgBox "登陆进入的审核人和当前用户为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    AuditionCheck = True
    
End Function
