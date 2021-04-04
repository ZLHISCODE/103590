VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeBat 
   Caption         =   "批量记帐"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargeBat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15105
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   10590
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmChargeBat.frx":0442
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18150
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   2
            Key             =   "MedicareType"
            Object.ToolTipText     =   "保险大类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   951
            MinWidth        =   951
            Picture         =   "frmChargeBat.frx":0CD6
            Key             =   "Drugstore"
            Object.Tag             =   "Drugstore"
            Object.ToolTipText     =   "药房设置"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmChargeBat.frx":0FF0
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmChargeBat.frx":162A
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBillList 
      BorderStyle     =   0  'None
      Height          =   9180
      Left            =   3045
      ScaleHeight     =   9180
      ScaleWidth      =   11985
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   915
      Width           =   11985
      Begin VB.CheckBox chk加班 
         Caption         =   "加班(&A)"
         Height          =   270
         Left            =   7125
         TabIndex        =   7
         Top             =   75
         Width           =   1170
      End
      Begin VB.CheckBox chk急诊 
         Caption         =   "急诊费用"
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   8475
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cbo开单人 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4455
         TabIndex        =   6
         Top             =   30
         Width           =   2205
      End
      Begin VB.PictureBox picBillBottom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   75
         ScaleHeight     =   2805
         ScaleWidth      =   11835
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   6390
         Width           =   11835
         Begin VSFlex8Ctl.VSFlexGrid vsMoney 
            Height          =   1665
            Left            =   15
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1110
            Width           =   3420
            _cx             =   6032
            _cy             =   2937
            Appearance      =   2
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
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmChargeBat.frx":1C64
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
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
            OwnerDraw       =   1
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
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "取消(&C)"
            Height          =   420
            Left            =   3945
            TabIndex        =   22
            ToolTipText     =   "热键:Esc"
            Top             =   1710
            Width           =   1560
         End
         Begin VB.Frame fraDrawDept 
            Height          =   1155
            Left            =   0
            TabIndex        =   32
            Top             =   -105
            Width           =   13575
            Begin VB.TextBox txt应收 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   9585
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   165
               Width           =   2175
            End
            Begin VB.ComboBox cboDrawDept 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   4140
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   720
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox txtMemo 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   210
               Width           =   7485
            End
            Begin VB.ComboBox cbo执行性质 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1155
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   720
               Width           =   1725
            End
            Begin MSMask.MaskEdBox txtDate 
               Height          =   360
               Left            =   9480
               TabIndex        =   19
               Top             =   720
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   635
               _Version        =   393216
               AutoTab         =   -1  'True
               HideSelection   =   0   'False
               MaxLength       =   19
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "yyyy-MM-dd hh:mm:ss"
               Mask            =   "####-##-## ##:##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label lbl应收 
               AutoSize        =   -1  'True
               Caption         =   "应收"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   9015
               TabIndex        =   12
               Top             =   240
               Width           =   510
            End
            Begin VB.Line lnSplitH 
               BorderColor     =   &H80000010&
               X1              =   30
               X2              =   18000
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Line lnSplitB 
               BorderColor     =   &H80000014&
               X1              =   0
               X2              =   18000
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Label lblDate 
               AutoSize        =   -1  'True
               Caption         =   "时间"
               Height          =   240
               Left            =   8955
               TabIndex        =   18
               Top             =   780
               Width           =   480
            End
            Begin VB.Label lblDrawDrugDept 
               AutoSize        =   -1  'True
               Caption         =   "领药部门"
               Height          =   255
               Left            =   3090
               TabIndex        =   16
               Top             =   773
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl病人备注 
               AutoSize        =   -1  'True
               Caption         =   "备注"
               Height          =   240
               Left            =   585
               TabIndex        =   10
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lbl执行性质 
               AutoSize        =   -1  'True
               Caption         =   "执行性质"
               Height          =   240
               Left            =   120
               TabIndex        =   14
               Top             =   780
               Width           =   960
            End
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00C0C0C0&
            Caption         =   "确定(&O)"
            Height          =   420
            Left            =   3960
            TabIndex        =   21
            ToolTipText     =   "热键：F2"
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   1110
         TabIndex        =   4
         Text            =   "cbo开单科室"
         Top             =   30
         Width           =   2160
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5760
         Left            =   60
         TabIndex        =   9
         Top             =   585
         Width           =   13065
         _ExtentX        =   23045
         _ExtentY        =   10160
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         TxtCheck        =   -1  'True
         TxtCheck        =   -1  'True
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   360
         RowHeightMin    =   360
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lbl开单人 
         AutoSize        =   -1  'True
         Caption         =   "开单人"
         Height          =   240
         Left            =   3690
         TabIndex        =   5
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lbl开单科室 
         Caption         =   "开单科室"
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   90
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1605
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":1CB5
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":224F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":27E9
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":2B3B
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":939D
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":FBFF
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":10199
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHead 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      ScaleHeight     =   780
      ScaleWidth      =   15300
      TabIndex        =   23
      Top             =   30
      Width           =   15300
      Begin VB.Frame fraTitle 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         TabIndex        =   24
         ToolTipText     =   "清除:F6"
         Top             =   0
         Width           =   15015
         Begin VB.CommandButton cmdSaveWholeSet 
            Caption         =   "保存为成套收费项目(&W)"
            Height          =   375
            Left            =   3300
            TabIndex        =   30
            Top             =   180
            Width           =   2790
         End
         Begin VB.ComboBox cboNO 
            ForeColor       =   &H00C00000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   13485
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   195
            Width           =   1425
         End
         Begin VB.CheckBox chkIn 
            Caption         =   "导"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "导入记帐单:F3"
            Top             =   180
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtIn 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   390
            Left            =   585
            MaxLength       =   8
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CommandButton cmdSelWholeSet 
            Caption         =   "成套(&T)"
            Height          =   375
            Left            =   2190
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "单据号"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   12705
            TabIndex        =   29
            Top             =   255
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picPatiList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9750
      Left            =   15
      ScaleHeight     =   9750
      ScaleWidth      =   2955
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1005
      Width           =   2955
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   7455
         Left            =   105
         TabIndex        =   0
         Top             =   360
         Width           =   2655
         _Version        =   589884
         _ExtentX        =   4683
         _ExtentY        =   13150
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cbo开单科室选择 
         Height          =   360
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   15
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.ComboBox cbo开单人选择 
         Height          =   360
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   2160
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   30
      Top             =   795
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'以下为主程序传入
Private mlng病区ID As Long
Private mlngDeptID  As Long
Private mlng病人ID As Long
Private mlngModule As Long
Private mstrPrivs As String '模块权限串
Private mbln补费 As Boolean '33744
Private mblnNurseStation As Boolean
Private mbytUseType As Byte '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐

'----------------------------------------------------------------------
Private mstrPrivsOpt As String '记帐操作权限
Private mstr成套项目 As String '成套项目
'数量合计,0,1;应收合计,0,7;实收合计,0,7 暂时不处理,主要原因是数量合计,应收合计等需要根据病人选择数来,可能会影响性能.
Private Const STR_HEAD = "" & _
"行,450,4;类别,750,1;项目,2175,1;商品名,1800,1;规格,1105,1;单位,520,4;付数,520,1;数次,570,1;单价,1055,7;" & _
"应收金额,1030,7;执行科室,1255,1;标志,520,4;类型,520,4"

'内部控制
Private mblnOK As Boolean    '数据保存是否存功
'数据对象
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsPati As New ADODB.Recordset '病人信息
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsWork As New ADODB.Recordset '当天上班的药房
Private mrsWarn As ADODB.Recordset  '病区报警线
Private mrsMedPayMode As ADODB.Recordset '所有可用的医疗付款方式
Private mrs费用类型 As ADODB.Recordset '费用类型
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '可选医生和护士
Private mrs领药部门 As ADODB.Recordset
Private mobjItem As XtremeReportControl.IReportRecordItem
Private mobjName As XtremeReportControl.IReportRecordItem
Private mobjBaseItem As Object    '成套项目设置部件

'程序对象
Private mobjBill As ExpenseBill '★★★费用单据对象★★★
Private mcolBillDetails As BillDetails '单据的收费细目集
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mcolBillInComes As BillInComes '收费细目的收入项目集
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes  '★★收入项目汇总集合(显示及打印时使用)★★
'相关枚举类型
Private Enum mPatiCol
    COL_病人ID = 0
    COL_主页ID = 1
    COL_选择 = 2
    COL_床号 = 3
    COL_姓名 = 4
    COL_性别 = 5
    COL_年龄 = 6
    COL_住院号 = 7
    COL_费别 = 8
    COL_险类 = 9
    COL_保险类别 = 10
    COL_婴儿 = 11
    COL_剩余款 = 12
    COL_预交余额 = 13
    COL_费用余额 = 14
    COL_担保额 = 15
    COL_当日额 = 16
    COL_适用病人 = 17
    COL_医疗付款方式 = 18
    COL_开单人 = 19
    COL_开单科室ID = 20
    COL_开单科室 = 21
End Enum
Private Enum mPanceIdx
    EM_HeadList = 1
    EM_PatiList = 2
    EM_BILLList = 3
End Enum
Private Enum BillColType       '单据控件的列类型
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum

Private Enum BillCol
    行 = 0
    类别 = 1
    项目 = 2
    商品名 = 3
    规格 = 4
    单位 = 5
    付数 = 6
    数次 = 7
    单价 = 8
    应收金额 = 9
    执行科室 = 10
    标志 = 11
    类型 = 12
End Enum
'程序变量
Private mblncboEnterCell As Boolean '避免循环调用
Private mblncboClick  As Boolean    '避免循环调用
Private mlngPreRow As Long '当前行号,用于列改变时判断
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库的出库检查方式
Private mbln处方职务检查 As Boolean     '是否进行处方职务检查
Private mbln处方限量检查 As Boolean     '是否进行处方限量检查
Private mblnOne As Boolean '是否只有一个可用收费类别
Private mblnWork As Boolean '当前是否有正在上班的药房
Private mlng药品类别ID As Long '当前单据操作的药品入出类别ID
Private mlng卫材类别ID As Long '当前单据操作的卫材入出类别ID
Private mstrUnitIDs As String   '当前操作员的所有病区ID
Private mstrWarn As String '已经报过警并选择继续的类别
Private mblnSendMateria As Boolean  '记帐后自动发药
Private mblnFirst As Boolean
Private mlngX As Long, mlngY As Long
Private mblnDrop As Boolean '在KeyDown中判断cbo开单人当前是否弹出
Private mblnValid As Boolean
Private mblnNewRow As Boolean
Private mdblItemNum As Double '数据库中当前输入费目的数次
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器
Private mblnNotClick As Boolean
Private mstr病人IDs As String   '当前选中的病人IDs
Private mlngSelPatiCount As Long  '当前选中的病人人数
Private mstrInsures As String   '当前选中的医保险类,多个用逗号分离
Private mblnPrintDrugList As Boolean '是否打印发药清单
Private mblnKeyReturn As Boolean '是否处理了回车

'医保相关参数
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
    实时监控 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private Type TY_PATIINFOR
    病人ID As Long
    主页ID As Long
    婴儿 As Integer
    险类 As Integer
    姓名 As String
    年龄 As String
    性别 As String
    适用病人 As String
    住院号 As String
    床号 As String
    费别 As String
    医疗付款方式 As String
    入院日期 As String
    出院日期  As String
    病人性质 As Integer
    状态 As Integer
    保险类别 As String
    剩余款 As Currency
    预交余额 As Currency
    费用余额 As Currency
    担保额 As Currency
    当日额 As Currency
    开单人 As String
    开单科室ID As Long
End Type
Private mstr最后转科时间 As String

Private Enum Pan
    C2提示信息 = 2
End Enum
Private mstr药品价格等级 As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Public Function ShowMe(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal bytUseType As Byte, ByVal lng病区ID As Long, ByVal lngDeptID As Long, ByVal lng病人ID As Long, _
    ByVal bln补费 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗口
    '     lng病人ID-病人ID
    '     lng病区ID-指定的病区
    '     bln补费-是否补费
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-08 17:46:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    mlng病区ID = lng病区ID: mlng病人ID = lng病人ID: mlngDeptID = lngDeptID
    mbln补费 = False: mbytUseType = bytUseType
    
    If gblnNurseStation Then
        mblnNurseStation = True
    Else
        mblnNurseStation = False
    End If
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    
    ShowMe = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    mlngX = X
    mlngY = Y
End Sub

Private Function GetRptPositionX(intTYPE As Integer) As Long
    On Error GoTo errH
    Dim i As Long
    If intTYPE = 1 Then
        For i = 0 To mlngX
            If rptPati.HitTest(i, 0).Column.Caption = "开单人" Then
                GetRptPositionX = i
                Exit For
            End If
        Next i
    Else
        For i = 0 To mlngX
            If rptPati.HitTest(i, 0).Column.Caption = "开单科室" Then
                GetRptPositionX = i
                Exit For
            End If
        Next i
    End If
    Exit Function
errH:
    Err.Clear
    GetRptPositionX = mlngX
End Function

Private Function GetRptPositionY() As Long
    On Error GoTo errH
    Dim i As Long
    For i = 0 To mlngY
        If rptPati.HitTest(mlngX, i).Row Is rptPati.SelectedRows(0) Then
            GetRptPositionY = i
            Exit For
        End If
    Next i
    Exit Function
errH:
    Err.Clear
    GetRptPositionY = mlngY
End Function

Private Sub LockedScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:锁定屏幕,以便在保存时,操作相关的按钮
    '入参:blnLocked-true:锁定屏幕;False-不锁定
    '编制:刘兴洪
    '日期:2015-07-13 10:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    cmdOK.Enabled = Not blnLocked
    cmdCancel.Enabled = Not blnLocked
    picHead.Enabled = Not blnLocked
    picBillBottom.Enabled = Not blnLocked
    picPatiList.Enabled = Not blnLocked
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
Private Sub cbo开单人_GotFocus()
    mblnKeyReturn = False
End Sub

Private Sub cmdOK_Click()
    
    If isValied() = False Then Exit Sub
    
    '数据保存
    Call LockedScreen(True)
    If SaveData = False Then
        Call LockedScreen(False)
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        Exit Sub
    End If
    Call LockedScreen(False)

    '恢复当前站点价格等级
    mstr药品价格等级 = gstr药品价格等级
    mstr卫材价格等级 = gstr卫材价格等级
    mstr普通价格等级 = gstr普通价格等级
    
    Call ClearRows: Call Bill.ClearBill: Call SetColNum
    Call ClearMoney
    Call SetMoneyList
    Call NewBill
    Call SetDrawDrugDeptEnabled
    If rptPati.Visible Then rptPati.SetFocus
    mblnOK = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call LockedScreen(False)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count = 0 Or Not Bill.Active Then Unload Me: Exit Sub
    
    If MsgBox("确实要清除当前单据中的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '急诊费用
    chk急诊.Value = 0: chk急诊.Visible = False
    txt应收.Text = gstrDec
    'txt实收.Text = gstrDec:
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    Call NewBill
    If Bill.Enabled And Bill.Visible Then Bill.SetFocus
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub cbo开单科室选择_Click()
    If mblnNotClick Then Exit Sub
    If Not mobjItem Is Nothing Then
        mobjItem.Value = cbo开单科室选择.ItemData(cbo开单科室选择.ListIndex)
        mobjName.Value = zlStr.NeedName(cbo开单科室选择.Text)
    End If
    cbo开单科室选择.Visible = False
    rptPati.Populate
End Sub

Private Sub cbo开单科室选择_LostFocus()
    cbo开单科室选择.Visible = False
End Sub

Private Sub cbo开单人选择_Click()
    If mblnNotClick Then Exit Sub
    If Not mobjItem Is Nothing Then
        mobjItem.Value = zlStr.NeedName(cbo开单人选择.Text)
    End If
    cbo开单人选择.Visible = False
    rptPati.Populate
End Sub

Private Sub cbo开单人选择_LostFocus()
    cbo开单人选择.Visible = False
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill
    Dim i As Long, lngPre As Long, strPre As String, strTmp As String, str药房IDs As String
    glngFormW = 15345: glngFormH = 11520
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call InitPanel
    
    RestoreWinState Me, App.ProductName, Me.Name
    sta.Visible = True
    
    mblnValid = False: mblnFirst = True: gbln处方限量 = False
    
    chkIn.Visible = True: txtIn.Visible = True
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    
    Call zlLoadDrawDeptData(mbytUseType, mlngDeptID)
    
    '初始化单据数据
    Set mobjBill = New ExpenseBill

    mstrUnitIDs = GetUserUnits
    
    
    '加载病人信息
    If Not InitData Then Unload Me: Exit Sub
    
    mstr药品价格等级 = gstr药品价格等级
    mstr卫材价格等级 = gstr卫材价格等级
    mstr普通价格等级 = gstr普通价格等级
    
    Call InitFace: Call NewBill
    Call LoadPatiInfo
    cbo开单科室.SelStart = 0
    cbo开单科室.SelLength = 0
    Call Auto开单科室
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    '调整发药部件
    Call SetDrawDrugDeptVisible
    
    mblnFirst = False
    On Error Resume Next
    'If Bill.Visible Then Bill.SetFocus

    Call SetDrawDrugDeptEnabled
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF3    '导入单据
            If chkIn.Visible And chkIn.Enabled Then chkIn.Value = IIf(chkIn.Value = 1, 0, 1)
        Case vbKeyF4
        Case vbKeyF6    '定位到病人选择框
            If rptPati.Visible Then rptPati.SetFocus
        Case vbKeyF7    '切换输入法
            If Not gbln简码切换 Then Exit Sub
            If Not (sta.Panels("WB").Visible And sta.Panels("PY").Visible) Then Exit Sub
            
            If sta.Panels("WB").Bevel = sbrRaised Then
                Call sta_PanelClick(sta.Panels("WB"))
            Else
                Call sta_PanelClick(sta.Panels("PY"))
            End If
        Case vbKeyF9 '定位到单据号输入框
            cboNO.SetFocus
            Call zlControl.TxtSelAll(cboNO)
        Case vbKeyF11
            'If cmd配方.Enabled And cmd配方.Visible Then Call cmd配方_Click
        Case vbKeyF12
            If Shift <> vbAltMask Then Exit Sub
            
            Call sta_PanelClick(sta.Panels("Drugstore"))
        Case vbKeyA, vbKeyR             '全选，全清
        Case vbKeyQ
            If Shift <> vbCtrlMask Then Exit Sub
            Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If Me.ActiveControl Is Bill Or Me.ActiveControl Is txtMemo Then Exit Sub
    '问题:29464
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub  '可能存在类似的刷卡:   ;1088029?
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub


Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域设置
    '编制:刘尔旋
    '日期:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngWidth As Long, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    Dim lngHeight As Long
    Dim strName As String
    If mlng病区ID <> 0 Then
       strName = "当前病区:" & GetDeptName(mlng病区ID)
    ElseIf mlngDeptID = 0 Then
       strName = "当前科室:" & GetDeptName(mlngDeptID)
    Else
        strName = "病人信息"
    End If
    
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_HeadList, 250, 580, DockTopOf, Nothing)
    lngHeight = picHead.Height / Screen.TwipsPerPixelY
    panThis.Title = ""
    panThis.Tag = mPanceIdx.EM_HeadList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picHead.hWnd
    panThis.MaxTrackSize.Height = lngHeight
    panThis.MinTrackSize.Height = lngHeight
    
    lngWidth = 2955 / Screen.TwipsPerPixelX
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_PatiList, lngWidth, 300, DockBottomOf, panThis)
    
    panThis.Title = strName
    panThis.Tag = mPanceIdx.EM_PatiList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picPatiList.hWnd
    
    lngWidth = (Me.ScaleWidth - 2955) / Screen.TwipsPerPixelX
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_BILLList, lngWidth, 580, DockRightOf, panThis)
    panThis.Title = ""
    panThis.Tag = mPanceIdx.EM_BILLList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picBillList.hWnd
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    
    Set dkpMain.PaintManager.CaptionFont = lbl开单科室.Font
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub
 
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPanceIdx.EM_HeadList
        Item.Handle = picHead.hWnd
    Case mPanceIdx.EM_PatiList
        Item.Handle = picPatiList.hWnd
    Case mPanceIdx.EM_BILLList
        Item.Handle = picBillList.hWnd
    End Select
End Sub


Private Sub InitReportColumn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Report控件列
    '编制:刘兴洪
    '日期:2015-07-09 11:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    
    On Error GoTo errHandle
    
    With rptPati
        
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False)
        Set objCol = .Columns.Add(COL_主页ID, "主页ID", 0, False)
        Set objCol = .Columns.Add(COL_选择, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_床号, "床号", 45, True)
        Set objCol = .Columns.Add(COL_姓名, "姓名", 120, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 30, True)
        Set objCol = .Columns.Add(COL_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(COL_住院号, "住院号", 60, True)
        Set objCol = .Columns.Add(COL_费别, "费别", 60, True)
        
        Set objCol = .Columns.Add(COL_险类, "险类", 0, False)
        Set objCol = .Columns.Add(COL_保险类别, "保险类别", 0, False)
        Set objCol = .Columns.Add(COL_婴儿, "婴儿", 0, False)
        Set objCol = .Columns.Add(COL_剩余款, "剩余款", 80, True)
        Set objCol = .Columns.Add(COL_预交余额, "预交余额", 80, True)
        Set objCol = .Columns.Add(COL_费用余额, "费用余额", 80, True)
        Set objCol = .Columns.Add(COL_担保额, "担保额", 0, False)
        Set objCol = .Columns.Add(COL_当日额, "当日额", 0, False)
        Set objCol = .Columns.Add(COL_适用病人, "适用病人", 0, False)
        Set objCol = .Columns.Add(COL_医疗付款方式, "医疗付款方式", 0, False)
        If mblnNurseStation Then
            Set objCol = .Columns.Add(COL_开单人, "开单人", 90, True)
            Set objCol = .Columns.Add(COL_开单科室ID, "开单科室ID", 0, False)
            Set objCol = .Columns.Add(COL_开单科室, "开单科室", 150, True)
        End If

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadPatiInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病区的相关病人信息
    '编制:刘兴洪
    '日期:2015-07-08 10:35:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, strSqlBaby As String
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim lngSelectRow As Long, lng病区ID As Long, rsBaby As ADODB.Recordset
    Dim strSQL As String, objChild As ReportRecord
    Dim lngColor As Long
    
     On Error GoTo errH
 
    Set mrsPati = GetPatiRsByUnit(mlng病区ID, mlng病人ID, True, True, False)
    lngSelectRow = -1
    With rptPati
        .Columns.Column(3).TreeColumn = True
        .PaintManager.TreeStructureStyle = xtpTreeStructureNone
        For i = 1 To mrsPati.RecordCount
            If Val(mrsPati!审核标志 & "") < 1 Or gTy_System_Para.byt病人审核方式 <> 1 Then
                If Val(mrsPati!婴儿序号 & "") = 0 Then
                    Set objRecord = .Records.Add()
                    objRecord.Tag = "0"
                    Set objItem = objRecord.AddItem(mrsPati!病人ID & "")
                    Set objItem = objRecord.AddItem(mrsPati!主页ID & "")
                    Set objItem = objRecord.AddItem("")
                    Set objItem = objRecord.AddItem(mrsPati!床号 & "")
                    Set objItem = objRecord.AddItem(mrsPati!姓名 & "")
                        objItem.Icon = img16.ListImages.Item(IIf(mrsPati!性别 & "" = "男", "Man", "Woman")).Index - 1
                    Set objItem = objRecord.AddItem(mrsPati!性别 & "")
                    Set objItem = objRecord.AddItem(mrsPati!年龄 & "")
                    Set objItem = objRecord.AddItem(mrsPati!住院号 & "")
                    Set objItem = objRecord.AddItem(mrsPati!费别 & "")
                    Set objItem = objRecord.AddItem(mrsPati!险类 & "")
                    Set objItem = objRecord.AddItem(mrsPati!保险类别 & "")
                    Set objItem = objRecord.AddItem(Val(mrsPati!婴儿序号 & ""))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!剩余款 & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!预交余额 & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!费用余额 & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!担保额 & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!当日额 & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Trim(mrsPati!适用病人 & ""))
                    Set objItem = objRecord.AddItem(Trim(mrsPati!医疗付款方式 & ""))
                    If mblnNurseStation Then
                        Set objItem = objRecord.AddItem(Trim(mrsPati!开单人 & ""))
                        Set objItem = objRecord.AddItem(Val(Nvl(mrsPati!开单科室ID)))
                        Set objItem = objRecord.AddItem(Nvl(mrsPati!开单科室名称))
                    End If
                Else
                    strSqlBaby = "Select 婴儿姓名, 婴儿性别, Zl_Age_Calc(0, 出生时间, Sysdate) As 年龄 From 病人新生儿记录 Where 病人id = [1] And 主页id = [2] And 序号 = [3]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSqlBaby, Me.Caption, Val(mrsPati!病人ID & ""), Val(mrsPati!主页ID & ""), Val(mrsPati!婴儿序号 & ""))
                    Set objChild = objRecord.Childs.Add
                    Set objItem = objChild.AddItem(mrsPati!病人ID & "")
                    Set objItem = objChild.AddItem(mrsPati!主页ID & "")
                    Set objItem = objChild.AddItem("")
                    Set objItem = objChild.AddItem(mrsPati!床号 & "")
                    If Not rsBaby.EOF Then
                        Set objItem = objChild.AddItem(Nvl(rsBaby!婴儿姓名))
                            objItem.Icon = img16.ListImages.Item(IIf(InStr(rsBaby!婴儿性别, "男") > 0, "Man", "Woman")).Index - 1
                        Set objItem = objChild.AddItem(IIf(InStr(rsBaby!婴儿性别, "男") > 0, "男", "女"))
                        Set objItem = objChild.AddItem(rsBaby!年龄 & "")
                    Else
                        Set objItem = objChild.AddItem(mrsPati!姓名 & "")
                            objItem.Icon = img16.ListImages.Item(IIf(mrsPati!性别 & "" = "男", "Man", "Woman")).Index - 1
                        Set objItem = objChild.AddItem(mrsPati!性别 & "")
                        Set objItem = objChild.AddItem(mrsPati!年龄 & "")
                    End If
                    Set objItem = objChild.AddItem(mrsPati!住院号 & "")
                    Set objItem = objChild.AddItem(mrsPati!费别 & "")
                    Set objItem = objChild.AddItem(mrsPati!险类 & "")
                    Set objItem = objChild.AddItem(mrsPati!保险类别 & "")
                    Set objItem = objChild.AddItem(Val(mrsPati!婴儿序号 & ""))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!剩余款 & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!预交余额 & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!费用余额 & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!担保额 & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!当日额 & ""), "0.00"))
                    Set objItem = objChild.AddItem(Trim(mrsPati!适用病人 & ""))
                    Set objItem = objChild.AddItem(Trim(mrsPati!医疗付款方式 & ""))
                    If mblnNurseStation Then
                        Set objItem = objChild.AddItem(Trim(mrsPati!开单人 & ""))
                        Set objItem = objChild.AddItem(Val(Nvl(mrsPati!开单科室ID)))
                        Set objItem = objChild.AddItem(Nvl(mrsPati!开单科室名称))
                    End If
                End If
                
                '病人颜色
                If Not IsNull(mrsPati!病人类型) Then
                    '保险病人用指定色显示
                    lngColor = zlDatabase.GetPatiColor(mrsPati!病人类型)
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = lngColor
                    Next
                ElseIf Not IsNull(mrsPati!险类) Then
                    '未指定病人类型的保险病人用红色显示
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = vbRed
                    Next
                End If
                '上次是否选择
                If mrsPati!病人ID = mlng病人ID Then
                    objRecord.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
                    objRecord.Tag = "1"
                    lngSelectRow = objRecord.Index
                    mlngSelPatiCount = mlngSelPatiCount + 1
                End If
            End If
            mrsPati.MoveNext
        Next
        .Populate
        If .Records.Count <> 0 Then Set .FocusedRow = .Rows(0)
        If lngSelectRow <> -1 Then Set .FocusedRow = .Rows(lngSelectRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
   SaveWinState Me, App.ProductName, Me.Name

    mlng药品类别ID = 0: mlng卫材类别ID = 0
    Set mrs费用类型 = Nothing
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrsWarn = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    Set mobjBaseItem = Nothing
    If Not OS.IsDesinMode Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    mlngSelPatiCount = 0
    mblnNurseStation = False
    mstrInsures = ""
    mstr病人IDs = ""
End Sub

Private Sub picBillBottom_Resize()
    Err = 0: On Error Resume Next
    If Not lbl执行性质.Visible Then
        lblDrawDrugDept.Left = lbl执行性质.Left
        cboDrawDept.Left = cbo执行性质.Left
    Else
        lblDrawDrugDept.Left = cbo执行性质.Left + cbo执行性质.Width + 100
        cboDrawDept.Left = lblDrawDrugDept.Left + lblDrawDrugDept.Width + 10
    End If
    txt应收.Left = picBillBottom.ScaleWidth - txt应收.Width - 100
    lbl应收.Left = txt应收.Left - lbl应收.Width - 10
    
    txtMemo.Width = lbl应收.Left - txtMemo.Left - 100
    fraDrawDept.Width = picBillBottom.ScaleWidth - fraDrawDept.Left + 10
    lnSplitB.X2 = fraDrawDept.Width
    lnSplitH.X2 = fraDrawDept.Width
    txtDate.Left = picBillBottom.ScaleWidth - txtDate.Width - 100
    lblDate.Left = txtDate.Left - lblDate.Width - 10
End Sub

Private Sub picBillList_Resize()
    Err = 0: On Error Resume Next
    With picBillList
        
        picBillBottom.Left = .ScaleLeft
        picBillBottom.Top = .ScaleHeight - picBillBottom.Height - 100
        picBillBottom.Width = .ScaleWidth
        
        Bill.Top = cbo开单科室.Top + cbo开单科室.Height + 50
        Bill.Left = 50
        Bill.Width = .ScaleWidth - Bill.Left * 2
        Bill.Height = picBillBottom.Top - Bill.Top - 50
    End With
End Sub
Private Sub picHead_Resize()
    Err = 0: On Error Resume Next
    With picHead
        fraTitle.Width = .ScaleWidth - fraTitle.Left
        cboNO.Left = fraTitle.Width - cboNO.Width - 100
        lblNO.Left = cboNO.Left - lblNO.Width
    End With
End Sub

Private Sub picPatiList_Resize()
    Err = 0: On Error Resume Next
    With picPatiList
        rptPati.Top = 50
        rptPati.Left = 50
        rptPati.Width = .ScaleWidth - rptPati.Left * 2
        rptPati.Height = .ScaleHeight - rptPati.Top * 2
    End With
End Sub
 
Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cbo开单科室.Enabled And cbo开单科室.Visible Then
            Debug.Print "cbo开单科室.SetFocus"
            cbo开单科室.SetFocus
            Debug.Print TypeName(Me.ActiveControl)
            Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode <> vbKeySpace Then Exit Sub
    If rptPati.SelectedRows.Count <= 0 Then Exit Sub
    Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_选择))
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn, i As Long
    Dim j As Long
    '如果点击表头的图片，就选中全部
    If Button <> 1 Then Exit Sub
    If rptPati.HitTest(X, Y).ht <> xtpHitTestHeader Then Exit Sub

    Set objColumn = rptPati.HitTest(X, Y).Column
    If objColumn Is Nothing Then Exit Sub
    If objColumn.Index <> COL_选择 Then Exit Sub

    If objColumn.Caption = "" Then
        objColumn.Caption = "1"
        rptPati.Columns(COL_选择).Icon = img16.ListImages("AllCheck").Index - 1
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_选择).Icon = img16.ListImages("Check").Index - 1
            For j = 0 To rptPati.Records(i).Childs.Count - 1
                rptPati.Records(i).Childs.Record(j).Item(COL_选择).Icon = img16.ListImages("Check").Index - 1
            Next j
            rptPati.Rows(i).Record.Tag = "1"
        Next
    Else
        objColumn.Caption = ""
        rptPati.Columns(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_选择).Icon = -1
            For j = 0 To rptPati.Records(i).Childs.Count - 1
                rptPati.Records(i).Childs.Record(j).Item(COL_选择).Icon = -1
            Next j
            rptPati.Rows(i).Record.Tag = "0"
        Next
    End If
    mstr病人IDs = GetPatiIDsBySel(mlngSelPatiCount)
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Select Case Item.Index
    Case 19
        '开单人
        cbo开单人选择.Top = rptPati.Top + GetRptPositionY * 15
        cbo开单人选择.Left = rptPati.Left + GetRptPositionX(1) * 15
        mblnNotClick = True
        cbo开单人选择.ListIndex = 0
        zlControl.CboLocate cbo开单人选择, Item.Value
        mblnNotClick = False
        Set mobjItem = Item
        cbo开单人选择.ZOrder: cbo开单人选择.Visible = True
        cbo开单人选择.SetFocus
    Case 21
        '开单科室
        cbo开单科室选择.Top = rptPati.Top + GetRptPositionY * 15
        cbo开单科室选择.Left = rptPati.Left + GetRptPositionX(2) * 15
        mblnNotClick = True
        cbo开单科室选择.ListIndex = 0
        zlControl.CboLocate cbo开单科室选择, Item.Value
        mblnNotClick = False
        Set mobjItem = Row.Record.Item(COL_开单科室ID)
        Set mobjName = Item
        cbo开单科室选择.ZOrder: cbo开单科室选择.Visible = True
        cbo开单科室选择.SetFocus
    Case Else
        If Row.Record.Tag = "1" Then
            Row.Record.Item(COL_选择).Icon = -1
            Row.Record.Tag = "0"
        Else
            Row.Record.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
            Row.Record.Tag = "1"
        End If
        rptPati.Populate
        mstr病人IDs = GetPatiIDsBySel(mlngSelPatiCount)
        Call Auto开单科室
    End Select
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存单据数据
    '编制:刘兴洪
    '日期:2015-07-08 14:17:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str病人IDs As String '记帐成功的病人ID
    Dim strNotPatiIDs As String '记帐不成功的病人ID
    Dim tyPati As TY_PATIINFOR, dtdtCurDate As Date
    Dim blnSavePrice As Boolean '是否保存为划价单
    Dim i As Long, strMessage As String
    Dim lngBringCount As Long, lngNotBringCount As Long
    
    On Error GoTo errHandle
    dtdtCurDate = zlDatabase.Currentdate
 
    '刘兴洪:重新获取领药部门
    Call zlReSetDrawDrugDept
    str病人IDs = "": strNotPatiIDs = ""
    lngBringCount = 0: lngNotBringCount = 0
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows.Count > 1 Then
            zlControl.StaShowPercent i / (rptPati.Rows.Count - 1), sta.Panels(2), Me
        End If
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i, mblnNurseStation)
            blnSavePrice = False
            Call zlCommFun.ShowFlash("正在对病人:" & tyPati.姓名 & "进行记帐,请稍后...")
            
            '初始化医保参数
            If tyPati.险类 <> 0 Then Call InitInsurePara(tyPati.病人ID, tyPati.险类)
            
            '重新给单据附值
            Call reSetBillObject(tyPati, mobjBill)
            
            '补录医保摘要
            Call InputItemMemo(tyPati)
            
            '重新取价格等级
            If gintPriceGradeStartType >= 2 Then
                Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, tyPati.病人ID, tyPati.主页ID, tyPati.医疗付款方式, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            End If

            '重新计算实收金额
            Call CalcMoneys(tyPati)
            
            '报警检查
            If CheckPatiChargeWrang(tyPati, blnSavePrice) = False Then Exit Function
            mobjBill.登记时间 = dtdtCurDate      '注意:打印发药单时要用到这个时间
            
            '重新计算医保统筹
            Call ReCalcInsure(tyPati)
            
            '按病人保存单据
    
            If SaveBill(tyPati, blnSavePrice, strMessage) Then
                '打印票据
                Call BillPrint(blnSavePrice)
                str病人IDs = str病人IDs & "," & tyPati.病人ID
                lngBringCount = lngBringCount + 1
            Else
                strNotPatiIDs = strNotPatiIDs & "," & tyPati.病人ID & "|" & strMessage
                lngNotBringCount = lngNotBringCount + 1
            End If
  
        End If
    Next
    Call zlCommFun.StopFlash
       
    str病人IDs = Mid(str病人IDs, 2)
    strNotPatiIDs = Mid(strNotPatiIDs, 2)
    If lngNotBringCount = 0 Then
        MsgBox "你共选择了" & mlngSelPatiCount & "个病人,成功记帐:" & lngBringCount & "个病人!", vbInformation + vbOKOnly, gstrSysName
    Else
        If MsgBox("你共选择了" & mlngSelPatiCount & "个病人,但成功记帐:" & lngBringCount & "个病人,未成功记帐:" & lngNotBringCount & "个病人!" & vbCrLf & "是否查看未记帐成功病人明细?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            frmChargeBatFailNote.ShowMe Me, strNotPatiIDs
        End If
    End If
    sta.Panels(2).Text = ""
    SaveData = str病人IDs <> ""
    Exit Function
errHandle:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPatiIDsBySel(ByRef lngCount As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取需要记帐的病人IDs
    '出参:lngCount-当前选中的人数
    '返回:返回病人ID, 多个用逗号分隔
    '编制:刘兴洪
    '日期:2015-07-09 10:13:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str病人IDs As String, lng病人ID As Long, i As Long
    Dim lng险类 As Long
    
    On Error GoTo errHandle
    lngCount = 0
    str病人IDs = "": mstrInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng病人ID = Val(rptPati.Rows(i).Record(COL_病人ID).Value)
            lng险类 = Val(rptPati.Rows(i).Record(COL_险类).Value)
            If lng病人ID <> 0 Then
                str病人IDs = str病人IDs & "," & lng病人ID
                lngCount = lngCount + 1
            End If
            If lng险类 <> 0 And InStr(mstrInsures & ",", "," & lng险类 & ",") = 0 Then
                mstrInsures = mstrInsures & "," & lng险类
            End If

        End If
    Next
    If str病人IDs <> "" Then str病人IDs = Mid(str病人IDs, 2)
    If mstrInsures <> "" Then mstrInsures = Mid(mstrInsures, 2)
    
    GetPatiIDsBySel = str病人IDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function SaveBill(tyPati As TY_PATIINFOR, _
    Optional blnSavePrice As Boolean, Optional ByRef strMessage As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前输入的记帐单据
    '入参:tyPati-当前病人信息
    '     blnSavePrice-当前保存为划价单
    '返回:单据保存返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 14:27:09
    '说明:mobjBill=单据对象
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, arrSQL As Variant, arrSMSQL As Variant
    Dim int序号 As Integer, int行号 As Integer, strNo As String, strTmp As String, str汇总号 As String
    Dim intParent As Integer, intParentNO As Integer
    Dim str消息 As String, intInsure As Integer
    Dim dbl数次 As Double, dbl单价 As Double
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, strStuffDept As String '记录卫料发料部门
    Dim strAddDate As String '记帐发生,自动发药,发料的时间
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str中药形态 As String
    Dim cllSMSQL As Collection, cllExctPro As Collection
    Dim varData As Variant, varTemp As Variant
    Dim rsItems As ADODB.Recordset
    Dim lng开单部门ID As Long
    Err = 0: On Error GoTo ErrHand:
    mobjBill.NO = zlDatabase.GetNextNo(14)
    mobjBill.病区ID = mlng病区ID
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate      '注意:打印发药单时要用到这个时间
    strAddDate = "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    Set cllSMSQL = New Collection
    Set cllExctPro = New Collection
    
    If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then strMessage = "记帐合法性检查失败!": Exit Function
    If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then strMessage = "记帐合法性检查失败!": Exit Function
    
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.数次 <> 0 Then
            intParent = 0: intParentNO = int序号
            For Each mobjBillIncome In mobjBillDetail.InComes
                int序号 = int序号 + 1 '当前记录序号
                '单据主体
                With mobjBill
                    If tyPati.病人性质 <> 1 Then
                        gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & IIf(.主页ID = 0, "NULL", .主页ID) & "," & _
                            IIf(Val(.标识号) = 0, "NULL", .标识号) & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "',"
                        If mblnNurseStation Then
                            gstrSQL = gstrSQL & IIf(.病区ID = 0, tyPati.开单科室ID, .病区ID) & "," & IIf(.科室ID = 0, tyPati.开单科室ID, .科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & tyPati.开单科室ID & ",'" & tyPati.开单人 & "',"
                            lng开单部门ID = tyPati.开单科室ID
                        Else
                            gstrSQL = gstrSQL & IIf(.病区ID = 0, .开单部门ID, .病区ID) & "," & IIf(.科室ID = 0, .开单部门ID, .科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                            lng开单部门ID = .开单部门ID
                        End If
                    Else
                        '门诊留观病人记门诊费用
                        gstrSQL = "Zl_门诊记帐记录_Insert('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.标识号) & "," & _
                            "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .费别 & "'," & .加班标志 & "," & .婴儿费 & ","
                        If mblnNurseStation Then
                            gstrSQL = gstrSQL & IIf(.科室ID = 0, tyPati.开单科室ID, .科室ID) & "," & tyPati.开单科室ID & ",'" & tyPati.开单人 & "',"
                            lng开单部门ID = tyPati.开单科室ID
                        Else
                            gstrSQL = gstrSQL & IIf(.科室ID = 0, .开单部门ID, .科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                            lng开单部门ID = .开单部门ID
                        End If
                    End If
                End With
                
                '收费细目部份
                With mobjBillDetail
                    '处理从属父号
                    If .序号 <> int行号 Then
                        int行号 = .序号
                        
                        '重新处理从属父号
                        If mobjBill.Details(.序号).从属父号 = 0 Then    '只有存在父项时,才会更新从属项
                            For i = .序号 + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).从属父号 = .序号 Then
                                    mobjBill.Details(i).从属父号 = int序号 '当父项目有多个收入项目(多个序号)时,取第一个序号
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                    
                    If tyPati.病人性质 <> 1 Then
                        gstrSQL = gstrSQL & IIf(.保险项目否, 1, 0) & "," & IIf(.保险大类ID = 0, "NULL", .保险大类ID) & ",'" & .保险编码 & "',"
                    End If
                    
                    dbl数次 = .数次
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                        dbl数次 = Format(.数次 * .Detail.住院包装, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & "," & IIf(.执行部门ID = 0, "NULL", .执行部门ID) & ","
                    
                    '收集卫料发料部门,以便自动发料
                    If Not (gbytBilling = 1 Or blnSavePrice) And gint卫材发料控制 <> 0 Then
                        'gint卫材发料控制:0-不自动发料，1-自动发料，2-本科室开单时自动发料
                        If .执行部门ID <> 0 And .收费类别 = "4" And .Detail.跟踪在用 _
                            And ((gint卫材发料控制 = 2 And .执行部门ID = lng开单部门ID) Or gint卫材发料控制 = 1) Then
                            If InStr("," & strStuffDept, "," & .执行部门ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .执行部门ID
                            End If
                        End If
                    End If
                End With
                
                '收入项目部份
                With mobjBillIncome
                    intParent = intParent + 1
                    dbl单价 = .标准单价
                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And gbln住院单位 Then
                        dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.住院包装, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .收入项目ID & "," & _
                        "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & ","
                    If tyPati.病人性质 <> 1 Then gstrSQL = gstrSQL & ZVal(.统筹金额) & ","
                End With
                
                If cbo执行性质.ListIndex < 0 Or cbo执行性质.Enabled = False Then
                    strTmp = "NULL,NULL"
                ElseIf cbo执行性质.ItemData(cbo执行性质.ListIndex) = 0 Then
                    strTmp = "NULL,NULL"
                Else
                    strTmp = "1," & cbo执行性质.ItemData(cbo执行性质.ListIndex)
                End If
               
                If mobjBillDetail.收费类别 = "7" Then
                    str中药形态 = "'" & mobjBillDetail.Detail.中药形态 & "'"
                Else
                    str中药形态 = "NULL"
                End If
                
                '其它部分
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    strAddDate & ",NULL," & IIf(gbytBilling = 1 Or blnSavePrice, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',"
                If tyPati.病人性质 <> 1 Then
                    gstrSQL = gstrSQL & "0," & IIf(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                        "NULL,'" & mobjBillDetail.摘要 & "'," & chk急诊.Value & "," & ZVal(lng医嘱ID) & "," & _
                        "Null,Null,'|" & mobjBill.煎法 & "', " & strTmp & ",NULL,'" & mobjBillDetail.Detail.类型 & "',0," & _
                        mobjBill.领药部门ID & "," & str中药形态 & ")"
                Else
                    gstrSQL = gstrSQL & "NULL,'" & mobjBillDetail.摘要 & "'," & ZVal(lng医嘱ID) & ",Null,Null,'|" & mobjBill.煎法 & "'," & _
                        strTmp & ",1," & str中药形态 & ",0,NULL," & ZVal(mobjBill.主页ID) & ","
                    If mblnNurseStation Then
                        gstrSQL = gstrSQL & IIf(mobjBill.病区ID = 0, tyPati.开单科室ID, mobjBill.病区ID) & ")"
                    Else
                        gstrSQL = gstrSQL & IIf(mobjBill.病区ID = 0, mobjBill.开单部门ID, mobjBill.病区ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
            Next
        End If
    Next

    If UBound(arrSQL) >= 0 Then
        '对SQL序列按收费细目ID排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        For i = 0 To UBound(arrSQL)
            varData = Split(arrSQL(i), ";")
            zlAddArray cllExctPro, varData(1)
        Next
        
        '执行自动发料
        If strStuffDept <> "" Then
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))
                strSQL = "zl_材料收发记录_处方发料(" & Split(strStuffDept, ",")(i) & ",25,'" & mobjBill.NO & "','" & _
                    UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1," & strAddDate & ")"
                zlAddArray cllExctPro, strSQL
            Next
        End If
                    
                    
        '执行SQL语句
        On Error GoTo errH
        blnTrans = True
        Call zlExecuteProcedureArrAy(cllExctPro, Me.Caption, True)
        
        '准备自动发药(仅普通记帐),必须在事务中才能读到数据
        If mblnSendMateria Then
            Set rsTmp = Get待发药清单(mobjBill.NO, Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), False)
            If rsTmp.RecordCount > 0 Then
                str汇总号 = zlDatabase.GetNextNo(20)
                For i = 0 To rsTmp.RecordCount - 1
                   strSQL = "ZL_药品收发记录_部门发药(" & rsTmp!库房ID & "," & rsTmp!ID & ",'" & UserInfo.姓名 & "'," & strAddDate & ",Null,Null,Null," & str汇总号 & ")"
                    zlAddArray cllSMSQL, strSQL
                    rsTmp.MoveNext
                Next
            End If
            rsTmp.Close
        End If
        
        '执行自动发药
        zlExecuteProcedureArrAy cllSMSQL, Me.Caption, True, True
        
        '记帐实时上传
        If gbytBilling = 0 And Not blnSavePrice And tyPati.险类 <> 0 Then
            '医保传输费用明细
            If MCPAR.记帐上传 And Not MCPAR.记帐完成后上传 Then
                str消息 = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , tyPati.险类) Then
                    gcnOracle.RollbackTrans
                    If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
                    strMessage = str消息
                    Exit Function
                End If
            End If
        End If
        gcnOracle.CommitTrans
        blnTrans = False
        '2.记帐后实时上传
        If gbytBilling = 0 And Not blnSavePrice And tyPati.险类 <> 0 Then
            '医保传输费用明细
            If MCPAR.记帐上传 And MCPAR.记帐完成后上传 Then
                str消息 = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , tyPati.险类) Then
                    If str消息 <> "" Then
                        MsgBox str消息, vbInformation, gstrSysName
                    Else
                        MsgBox "单据""" & mobjBill.NO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '加入单据历史记录(所有类型单据)
        For i = 0 To cboNO.ListCount - 1
            strNo = strNo & "," & cboNO.List(i)
        Next
        strNo = mobjBill.NO & strNo
        cboNO.Clear
        For i = 0 To UBound(Split(strNo, ","))
            cboNO.AddItem Split(strNo, ",")(i)
            If i = 9 Then Exit For '只显示10个
        Next
        '医保接口
        If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
    End If
    Call zlChargeSaveAfter_Plugin(mlngModule, mobjBill.病人ID, mobjBill.主页ID, False, 2, mobjBill.NO)
    SaveBill = True
    Exit Function
ErrHand:
    strMessage = Err.Description
    If ErrCenter = 1 Then Resume
    Exit Function
errH:
    If Err.Description Like "*当前计算单价不一致*" Then
       If blnTrans Then gcnOracle.RollbackTrans
       If MsgBox("某些分批药品价格已发生变化，要自动重算价格吗？", vbYesNo + vbQuestion + vbDefaultButton1, App.ProductName) = vbYes Then
           Call CalcMoneys(tyPati)
           Call ShowDetails
           Call ShowMoney
           If InStr(Err.Description, "[ZLSOFT]") > 0 Then
                strMessage = Split(Err.Description, "[ZLSOFT]")(1)
           Else
                strMessage = Err.Description
           End If
           Exit Function
       End If
    Else
        If blnTrans Then gcnOracle.RollbackTrans
        If InStr(Err.Description, "[ZLSOFT]") > 0 Then
             strMessage = Split(Err.Description, "[ZLSOFT]")(1)
        Else
             strMessage = Err.Description
        End If
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

Private Sub Auto开单科室()

    Dim i As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnFind As Boolean
    
    If mstr病人IDs <> "" Then
        strSQL = "Select 当前科室ID From 病人信息 Where 病人ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Split(mstr病人IDs, ",")(0)))
    Else
        strSQL = "Select 当前科室ID From 病人信息 Where 病人ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    End If
    
    blnFind = False
    If Not rsTmp.EOF Then
        For i = 1 To cbo开单科室.ListCount
            If Val(cbo开单科室.ItemData(i - 1)) = Val(Nvl(rsTmp!当前科室id)) Then
                blnFind = True
                cbo开单科室.ListIndex = i - 1
                Exit For
            End If
        Next i
    End If
    If blnFind = False Then cbo开单科室.ListIndex = 0
End Sub


Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-08 17:33:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dtdtCurDate As Date     '服务器当前时间
    On Error GoTo errH
    
    '读取中药输入快捷
    Call ReadABCNum(mstrPrivsOpt)
    
    '不同药房药品出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    

    '------------------批量读取------------------
    strSQL = " Select '处方职务' As 分类,count(药名ID) As num From 药品特性 Where 处方职务<>'00' Union All " & _
             " Select '处方限量' As 分类,count(药名ID) As num From 药品特性 Where 处方限量>0    "
    
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    rsTmp.Filter = "分类='处方职务'"
    If Not rsTmp.EOF Then mbln处方职务检查 = (rsTmp!Num > 0)
    
    rsTmp.Filter = "分类='处方限量'"
    If Not rsTmp.EOF Then mbln处方限量检查 = (rsTmp!Num > 0)

    
    '------------------批量读取------------------
            
    If Init开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mstrPrivs, 0, mlngDeptID) = False Then Exit Function
    
    If mblnNurseStation Then
        If Init开单人开单科室(cbo开单人选择, cbo开单科室选择, mrs开单人, mrs开单科室, mstrPrivs, 0, mlngDeptID) = False Then Exit Function
    End If
    
    If gstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.编码,A.名称 as 类别 " & _
        "   From 收费项目类别 A," & _
        "          (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) J " & _
        "   Where A.编码=J. Column_Value " & _
        "   Order by 序号"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr收费类别, "'", ""))
    
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别,请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '当只有一种可选收费类别时,不用用户选择
    mblnOne = (mrsClass.RecordCount = 1)
    
    If InStr(gstr收费类别, "'5'") > 0 Or InStr(gstr收费类别, "'6'") > 0 _
        Or InStr(gstr收费类别, "'7'") > 0 Or gstr收费类别 = "" Then
        mlng药品类别ID = ExistIOClass(9)
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr收费类别, "'4'") > 0 Or gstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(41)
        If mlng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '执行部门
    strSQL = _
    "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
    " From 部门表 A,部门性质说明 B " & _
    " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
    " And B.部门ID=A.ID and B.服务对象 IN(2,3) " & _
    " Order by B.服务对象,A.编码"
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    dtdtCurDate = zlDatabase.Currentdate
    txtDate.Text = Format(dtdtCurDate, "yyyy-MM-dd HH:mm:ss")
    
    '自动识别加班
    If OverTime(dtdtCurDate) Then chk加班.Value = Checked
    Set mrsWarn = GetUnitWarn
 
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetDeptName(ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取部门名称
    '入参:lngDeptID-部门ID(或病区ID)
    '返回:返回取部门名称
    '编制:刘兴洪
    '日期:2015-07-15 17:52:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 名称 From 部门表 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
    If rsTemp.EOF = False Then GetDeptName = Nvl(rsTemp!名称)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据表单要完成的功能设置界面布局
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-08 16:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead() As String, i As Long
    With cbo执行性质
        .Clear
        .AddItem "正常"
        .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "离院带药"
        .ItemData(.NewIndex) = 3
        .AddItem "自取药"
        .ItemData(.NewIndex) = 4
    End With
                
    Call InitReportColumn
            
    '初始化表格
    arrHead = Split(STR_HEAD, ";")
    With Bill
        .Font.Size = 10.5
        .cboObj.Font.Size = 10.5
        
        .Cols = UBound(arrHead) + 1
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.项目
        .PrimaryCol = BillCol.项目
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.行) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        .ColData(BillCol.行) = BillColType.UnFocus
        
        .ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
        If mblnOne Then .ColData(BillCol.类别) = BillColType.UnFocus
        If .ColData(BillCol.类别) <> BillColType.UnFocus Then
            .LocateCol = BillCol.类别
        End If
        .ColData(BillCol.项目) = BillColType.CommandButton  '项目输入,按扭可选
        .ColData(BillCol.数次) = BillColType.Text '数/次输入
        .ColData(BillCol.商品名) = BillColType.UnFocus    '商品名跳过
        .ColData(BillCol.规格) = BillColType.UnFocus    '规格跳过
        .ColData(BillCol.单位) = BillColType.UnFocus  '单位跳过
        .ColData(BillCol.付数) = BillColType.UnFocus  '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
        .ColData(BillCol.单价) = BillColType.UnFocus '单价缺省跳过,当项目变价时,设为输入(4)
        .ColData(BillCol.应收金额) = BillColType.UnFocus  '应收金额跳过
'        .ColData(BillCol.数量合计) = BillColType.UnFocus   '数量合计跳过
'        .ColData(BillCol.实收合计) = BillColType.UnFocus   '实收合计跳过
'        .ColData(BillCol.应收合计) = BillColType.UnFocus   '应收合计跳过
         .ColData(BillCol.执行科室) = BillColType.ComboBox '默认取开单科室或上一科室
        .ColData(BillCol.标志) = BillColType.UnFocus '标志缺省跳过,当为手术时,设为复选(-1)
        .ColData(BillCol.类型) = BillColType.UnFocus  '类型缺省跳过
          
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
        .MsfObj.ScrollBars = 3
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
    End With
    
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name)
    If gTy_System_Para.byt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
    
    Call SetMoneyList '初始化汇总列表
     
    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = True
    sta.Panels("PY").Visible = gbln简码切换 '35242
    sta.Panels("WB").Visible = gbln简码切换
    '简码匹配方式：0-拼音,1-五笔,2-两者
    If gbytCode = 0 Then
        sta.Panels("PY").Bevel = sbrInset
        sta.Panels("WB").Bevel = sbrRaised
    ElseIf gbytCode = 1 Then
        sta.Panels("PY").Bevel = sbrRaised
        sta.Panels("WB").Bevel = sbrInset
    Else
        sta.Panels("PY").Bevel = sbrInset
        sta.Panels("WB").Bevel = sbrInset
    End If
    txt应收.Text = gstrDec ': txt实收.Text = gstrDec
 
    Call SetShowCol ' 设置付数列
    
    '普通记帐和科室分散记帐或划价时,新增或修改操作中允许输入中药配方
    cbo执行性质.Visible = True: lbl执行性质.Visible = True
    cmdSelWholeSet.Visible = True
    cmdSaveWholeSet.Visible = True
    
    '交换开单科室与开单人位置
    If gblnFromDr Then
        Call ExChangeLocate(cbo开单科室, cbo开单人)
        Call ExChangeLocate(lbl开单科室, lbl开单人)
        cbo开单科室.TabStop = False
    End If
    
    If mblnNurseStation Then
        lbl开单科室.Visible = False
        cbo开单科室.Visible = False
        lbl开单人.Visible = False
        cbo开单人.Visible = False
        lbl执行性质.Visible = False
        cbo执行性质.Visible = False
        lblDrawDrugDept.Visible = False
        cboDrawDept.Visible = False
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    If cbo开单科室选择.Visible = True Then cbo开单科室选择.Visible = False
    If cbo开单人选择.Visible Then cbo开单人选择.Visible = False
End Sub

 Private Sub SetMoneyList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前收入项目行数调整各列宽
    '编制:刘兴洪
    '日期:2015-07-08 17:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngW As Long, i As Long
    
    With vsMoney
        .Clear
        .Cols = 2: .Rows = 2
        .TextMatrix(0, 0) = "项目"
        .TextMatrix(0, 1) = "金额"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, 0)
            .FixedAlignment(i) = 4
        Next
        lngW = .Width - 60
        If .Rows > .Height / .RowHeight(0) Then lngW = lngW - 250
        .ColWidth(0) = lngW * 0.5: .ColWidth(1) = lngW * 0.5
        .ColAlignment(0) = 1: .ColAlignment(1) = 7
        .Row = 1
    End With
End Sub
Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
    Case "PY", "WB"
        If Panel.Bevel = sbrRaised And gbln简码切换 Then
            '切换并保存简码匹配方式
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "简码方式", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("简码方式", , , 0))
        End If
    Case "Drugstore"
        With frmSetExpence
            .mlngModul = mlngModule
            .mstrPrivs = mstrPrivs
            '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
            '           0:普通记帐,1-科室分散记帐,2-医技科室记帐
            .mbytInFun = 0
            .mbytUseType = 0
            .mblnOnlyDrugStock = True
            .Show 1, Me
        End With
    End Select
End Sub
 
Private Sub SetShowCol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:付数列的控制(浏览时展开)
    '编制:刘兴洪
    '日期:2015-07-08 18:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mrsClass.Filter = "编码='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.付数) = 0
    ElseIf Bill.ColWidth(BillCol.付数) = 0 Then
        Bill.ColWidth(BillCol.付数) = 520
    End If
End Sub
Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
    '功能：获取指定列的原始列宽
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub cboDrawDept_Click()
    Dim lng领药部门ID As Long
    If cboDrawDept.ListIndex <> -1 Then lng领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
    If Not mobjBill Is Nothing Then
        If mobjBill.领药部门ID = lng领药部门ID Then Exit Sub
        mobjBill.领药部门ID = lng领药部门ID
    End If
End Sub

Private Sub cboDrawDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 And Not cboDrawDept.Locked Then
        lngIdx = zlControl.CboMatchIndex(cboDrawDept.hWnd, KeyAscii)
        If lngIdx = -1 And cboDrawDept.ListCount > 0 Then lngIdx = 0
        cboDrawDept.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        If cboDrawDept.ListIndex = -1 Then Beep: Exit Sub
        mobjBill.领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo开单科室_GotFocus()
    zlControl.TxtSelAll cbo开单科室
End Sub

Private Sub cbo开单科室_LostFocus()
    cbo开单科室.SelLength = 0
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then cbo开单科室.Text = ""
End Sub

Private Sub cbo执行性质_Click()
    If mobjBill Is Nothing Then Exit Sub
    mobjBill.执行性质 = cbo执行性质.ItemData(cbo执行性质.ListIndex)
End Sub

Private Sub cbo执行性质_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lng执行科室ID As Long
    Dim rsTemp As ADODB.Recordset, dbl数次 As Double, dbl价格 As Double
    Dim strSQL As String
    
    '保存为存套收费项目
    '问题:27327
    
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:序号,父号,收费细目ID,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    Err = 0: On Error GoTo ErrHand:

    With mobjBill
        strItems = ""
        For i = 1 To .Details.Count
             '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
            If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.类别) > 0 Then
                lng执行科室ID = 0
            ElseIf InStr(1, ",0,4", .Details(i).Detail.执行科室) > 0 Then
                lng执行科室ID = .Details(i).执行部门ID
            Else
                lng执行科室ID = 0
            End If
            '问题:52349
            dbl数次 = .Details(i).数次
            dbl价格 = IIf(.Details(i).Detail.变价, .Details(i).InComes(1).标准单价, 0)
            If InStr(",5,6,7,", .Details(i).收费类别) > 0 And gbln住院单位 Then
                dbl数次 = Format(dbl数次 * .Details(i).Detail.住院包装, gstrFeePrecisionFmt)
                dbl价格 = Format(dbl价格 / .Details(i).Detail.住院包装, gstrFeePrecisionFmt)
            End If
            strItems = strItems & "|" & .Details(i).序号 & "," & .Details(i).从属父号 & "," & .Details(i).收费细目ID & "," & .Details(i).付数 & "," & dbl数次 & "," & dbl价格 & "," & lng执行科室ID
         Next
         If strItems = "" Then
            MsgBox "单据未输入任何信息,不能保存为成套收费项目,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
    End With
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, 1150, mstrPrivsOpt, strItems)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelWholeSet_Click()
    '选成套项目
    Dim rsSel As ADODB.Recordset, lng病人ID As Long, lng开单部门ID As Long
    Dim tmpBill As New ExpenseBill, byt婴儿费 As Byte, dtCurdate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim intInsure As Integer
    Dim bln中药 As Boolean
    
    intInsure = 0
    If mlngSelPatiCount = 0 Then
        MsgBox "请先选择病人,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
            
    If mobjBill Is Nothing Then
        lng病人ID = 0
        If cbo开单科室.ListIndex < 0 Then
            lng开单部门ID = 0
        Else
            lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
    Else
        lng病人ID = mobjBill.病人ID: lng开单部门ID = mobjBill.开单部门ID
    End If
    
    If mlngSelPatiCount = 0 Then
        MsgBox "请先选择病人,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If zlSelectWholeItems(Me, mlngModule, mstrPrivsOpt, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo ErrHand:
    Screen.MousePointer = 11
    
    Set tmpBill = ImportWholeSet(Me, intInsure, rsSel, lng病人ID, gbln住院单位, lng开单部门ID, byt婴儿费, 2, chk加班.Value = 1, _
        0, 2, UserInfo.姓名, zlStr.NeedName(cbo开单人.Text), , mblnNurseStation, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
        
    '处理数据
    '清除导入的病人信息
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    bln中药 = False
    With mobjBill
        For i = 1 To .Details.Count - 1
            If .Details(i).收费类别 = "7" Then bln中药 = True: Exit For
        Next
    End With
    
    dtCurdate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.登记时间 = dtCurdate
    mobjBill.操作员编号 = UserInfo.编号
    mobjBill.操作员姓名 = UserInfo.姓名
    mobjBill.加班标志 = chk加班.Value
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
    
    
    Bill.Redraw = False
    Bill.ClearBill
    '问题号:116774,焦博,2017/12/28,批量记帐调用全是药品的成套项目后,点击单据会报错
    Bill.Rows = IIf(mobjBill.Details.Count = 0, 2, mobjBill.Details.Count + 1)
    
    Call InitBillColumnColor
    '记帐分类报警
    mstrWarn = ""
        
    Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
        
    '等上面的读病人后确定费别后,再计算价格
    Dim tyPati As TY_PATIINFOR
    Call CalcMoneys(tyPati)   '在CalcMoneys中不计算实收金额(所以要传入空的tyPati)
    
    Call ShowDetails
    Call ShowMoney
    With Bill
        For i = 1 To .Rows - 1
            .TextMatrix(i, BillCol.行) = i
        Next
    End With
    Bill.Redraw = True
    Call SetDrawDrugDeptEnabled
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ReSetDefault执行科室(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置缺省的执行科室
    '编制:刘兴洪
    '日期:2015-07-09 10:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人科室ID As Long, lngDoUnit As Long, str药房IDs As String
    
    Dim dblStock As Double
    Err = 0: On Error GoTo ErrHand:
    With mobjBill.Details(lngRow)
         '卫材和药品部分
        '病人科室ID
        lng病人科室ID = mobjBill.科室ID
        If cbo开单科室.Visible Then
            If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        Else
            If lng病人科室ID = 0 Then lng病人科室ID = mlngDeptID
            If lng病人科室ID = 0 Then lng病人科室ID = GetNurseStationFirstPatiDeptID '护士工作站,取第一个病人科室ID
            If lng病人科室ID = 0 Then lng病人科室ID = mlng病区ID
        End If
        
         '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
        If .Detail.类别 = "4" Then
            lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, mobjBill.病区ID)
            If lngDoUnit = 0 Then lngDoUnit = IIf(cbo开单科室.Visible, Get开单科室ID, lng病人科室ID)
        End If
        
        lngDoUnit = Get收费执行科室ID(.Detail.类别, .Detail.ID, _
             .Detail.执行科室, lng病人科室ID, IIf(cbo开单科室.Visible, Get开单科室ID, lng病人科室ID), 2, lngDoUnit, mobjBill.病区ID, .执行部门ID)
       .执行部门ID = lngDoUnit
        
        If InStr(",5,6,7,", .Detail.类别) > 0 Then
            '当前行药品库存
            If Not gbln分离发药 Then
                dblStock = GetStock(.Detail.ID, lngDoUnit)
                If gbln住院单位 Then
                    dblStock = dblStock / .Detail.住院包装
                End If
                  .Detail.库存 = dblStock
                Call ShowStock(.Detail.名称, .Detail.库存)
            Else
                str药房IDs = Decode(.Detail.类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                If str药房IDs <> "" Then
                    dblStock = GetMultiStock(.Detail.ID, str药房IDs)
                    If gbln住院单位 Then
                        dblStock = dblStock / .Detail.住院包装
                    End If
                    .Detail.库存 = dblStock
                    Call ShowStock(.Detail.名称, .Detail.库存)
                End If
            End If
        ElseIf .Detail.类别 = "4" And .Detail.跟踪在用 Then
            dblStock = GetStock(.Detail.ID, lngDoUnit)
            .Detail.库存 = dblStock
            Call ShowStock(.Detail.名称, .Detail.库存)
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
 End Sub
 
Private Sub ShowStock(str药品 As String, dbl库存 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示药品或卫材的库存
    '编制:刘兴洪
    '日期:2015-07-09 10:51:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
     On Error GoTo errHandle
    If InStr(1, mstrPrivsOpt, ";显示库存;") > 0 Then
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & dbl库存
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]" & IIf(dbl库存 > 0, "有", "无") & "库存."
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, i As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
 '   If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    
    '类别
    If Bill.TextMatrix(0, Bill.Col) = "类别" Then
        If Bill.ListIndex > 0 Then Exit Sub
        For i = 0 To Bill.cboObj.ListCount - 1
            If IsNumeric(Bill.CboText) Then
                If Split(Bill.cboObj.List(i), "-")(0) = Val(Bill.CboText) Then
                    Bill.ListIndex = i
                    Exit Sub
                     
                End If
            ElseIf zlCommFun.IsCharAlpha(Bill.CboText) Then
                If zlCommFun.SpellCode(Split(Bill.List(i) & "-", "-")(1)) Like UCase(Bill.CboText) & "*" Then
                    Bill.ListIndex = i
                    Exit Sub
                End If
            ElseIf Split(Bill.ItemData(i) & "-", "-")(1) Like "*" & UCase(Bill.CboText) & "*" Then
                Bill.ListIndex = i
                Exit Sub
            End If
        Next
        Exit Sub
    End If
    
    If Bill.TextMatrix(0, Bill.Col) <> "执行科室" Then Exit Sub
    If Bill.ListIndex <> -1 Then Exit Sub
    
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .收费类别) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub
Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mobjBill.Details.Count >= Row Then
        '带从属项目的项删除确认
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("该项目带有 " & bytsubs & " 个从属项目,删除该项目也将删除它的从属项目,继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).从属父号 <> 0 Then '从属项目删除确认
            If MsgBox("该项目是[" & mobjBill.Details(mobjBill.Details(Row).从属父号).Detail.名称 & "]的从属项目,确定要删除它吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln从项汇总折扣 = gbln从项汇总折扣
            End If
        ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        If bln从项汇总折扣 Then lngMainRow = mobjBill.Details(Bill.Row).从属父号 '如果是从项,删除之前记下从项的从属父号,如果是主项,则级联删除,不用重算
        

        
        '删除处理
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).从属父号 = Row Then
                Call DeleteDetail(i) '反顺序删除其从属行
            End If
        Next
        Call DeleteDetail(Row) '删除该行
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理删除
        
        mlngPreRow = 0    '表示行改变了
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        Call SetDrawDrugDeptEnabled
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, tyPati As TY_PATIINFOR
    Dim lng执行科室 As Long, str执行科室 As String
    If mblncboClick Then Exit Sub  '避免同一过程中因设置bill的值循环调用,注意在任何exit sub 之前设置mblncboClick = False
    '药品库存检查
    If Not (ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "执行科室") Then Exit Sub
    If mobjBill.Details.Count < Bill.Row Then Exit Sub
    
    mblncboClick = True
    
    With mobjBill.Details(Bill.Row)
        If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
            lng执行科室 = .执行部门ID: str执行科室 = Bill.TextMatrix(Bill.Row, Bill.Col)
            .执行部门ID = Bill.ItemData(Bill.ListIndex)
            
            Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
            
            If InStr(",5,6,7,", .收费类别) > 0 Then
                '取库存
                dblStock = GetStock(.收费细目ID, .执行部门ID)
                If gbln住院单位 Then
                    dblStock = dblStock / .Detail.住院包装
                End If
                .Detail.库存 = dblStock  '记录当前行药品库存
                Call ShowStock(.Detail.名称, .Detail.库存)
                
                '药房改变,实价药品重新计算价格
                Call CalcMoneys(tyPati, Bill.Row)  '实收金额不参与计算,所以typati传入为空
                Call ShowDetails(Bill.Row)
                Call ShowMoney
                
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                '取库存
                dblStock = GetStock(.收费细目ID, .执行部门ID)
                .Detail.库存 = dblStock
                Call ShowStock(.Detail.名称, .Detail.库存)
                
                '发料部门改变,时价卫材重新计算价格
                If .Detail.变价 Then
                    Call CalcMoneys(tyPati, Bill.Row)
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                End If
            ElseIf InStr(",4,5,6,7,", .收费类别) = 0 Then
                If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
            End If
            If mobjBill.Details(Bill.Row).数次 <> 0 Then
                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                    Bill.Text = "": Bill.TxtVisible = False
                    Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                    mblncboClick = False: Exit Sub
                End If
            End If
        End If
    End With
    mblncboClick = False
End Sub


Private Sub Bill_CellCheck(Row As Long, Col As Long)
    '说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean, tyPati As TY_PATIINFOR
    
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    
    '新增的未处理行无效
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "F" And mobjBill.Details(i).附加标志 = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '可能只存在附加手术后又改成了主手术,需要重新计处理:25495
        blnReSet = (strCheck = "" And mobjBill.Details(Row).收费类别 = "F" And mobjBill.Details(Row).附加标志 = 1)
    End If
    
    If blnReSet Then
        With mobjBill.Details(Row)
            .附加标志 = IIf(strCheck = "", 0, 1)
            Call CalcMoneys(tyPati, Row)   '实收金额不参与计算,所以typati传入为空
            
            Call ShowDetails(Row)
        End With
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub
Private Sub Bill_CommandClick()
    Dim lng项目id As Long, blnCancel As Boolean, bln护士 As Boolean
    Dim str类别 As String, str特准项目 As String
    Dim int病人来源 As Integer, int险类 As Integer
    Dim str排除类别 As String
    
    Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
    If gbln收费类别 Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
        End If
    Else
        str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
    End If
    int病人来源 = 2
    
    If zlCheckBill存在非散装草药() = True Then mblnSelect = False: Exit Sub
    
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, gbln住院单位, str类别, , , str特准项目, _
        0, str排除类别, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    If lng项目id = 0 Then mblnSelect = False: Exit Sub
    
    Bill.Text = lng项目id
    mblnSelect = True
    Call Bill_KeyDown(13, 0, blnCancel)
    Bill.SetFocus
    If Not blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False
        Call zlCommFun.PressKey(13)
    End If
End Sub



Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '功能：处理单据输入
    Dim rs药品信息 As ADODB.Recordset
    Dim lng项目id As Long, str类别 As String, bln护士 As Boolean
    Dim str特准项目 As String, int病人来源 As Integer, lng病人科室ID As Long, int险类 As Integer
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double, dblNum As Double, lngOld付数 As Long
    Dim blnSkip As Boolean, curTotal As Currency
    Dim lngDoUnit As Long, str摘要 As String, blnInput As Boolean
    Dim str药房IDs As String, bln负数记帐 As Boolean, cur余额 As Currency
    Dim curItemMoney As Currency
    Dim colStock As Collection, str排除类别 As String
    Dim cllData As Collection, tyPati As TY_PATIINFOR
    
    On Error GoTo errH
    '问题号:110693,焦博,2017/08/07,读取执行科室不正确
    mobjBill.病区ID = mlng病区ID
    If Not (KeyCode = 13 And Bill.Active) Then Exit Sub
    
    If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
                        
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "类别"
            If Bill.ListIndex = -1 Then Exit Sub
            '不输入类别时不会定位到类别列
            If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                '一旦改更收费类别,则清除(如有)原有该项目内容
                For i = 2 To Bill.Cols - 1
                    Bill.TextMatrix(Bill.Row, i) = ""
                Next
                If mobjBill.Details.Count >= Bill.Row Then
                    Set mobjBill.Details(Bill.Row).Detail = New Detail
                    Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                    With mobjBill.Details(Bill.Row)
                        .收费细目ID = 0: .收费类别 = ""
                    End With
                    Call CalcMoneys(tyPati) 'tyPati传入空时,实收金额不参与计算
                    Call ShowMoney
                End If
            End If
            Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '暂时用RowData记录所选择的收费类别
        Case "项目"
            '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
            If Bill.Text <> "" Then
                '如果在已输入的项目上按回车,或选择器选择
                If mobjBill.Details.Count >= Bill.Row Then
                    '通过按钮选择是返回的ID,而输入则是文本,如果是一样的,则不改变
                    If Bill.TextMatrix(Bill.Row, BillCol.项目) = Bill.Text Then
                        Bill.TxtVisible = False: Bill.CmdVisible = False: Exit Sub
                    End If
                End If
            
                sta.Panels(2).Text = "": sta.Panels("MedicareType").Text = ""
                blnInput = True
                If mblnSelect Then
                    mblnSelect = False '立即清除该标志
                    Set mobjDetail = GetInputDetail(Val(Bill.Text))
                Else
                    If gbln收费类别 Then
                        If Bill.RowData(Bill.Row) = 0 Then
                            sta.Panels(2) = "没有确定费用类别,请先输入类别！"
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                        str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                    Else
                        Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
                        str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
                    End If
                    
                    int病人来源 = 2
                    If zlCheckBill存在非散装草药 Then
                        '存在非散装的,界面中就不能进行录入
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.SetFocus: Cancel = True: Exit Sub
                    End If
                    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, gbln住院单位, str类别, Bill.Text, _
                        Bill.TxtHwnd, str特准项目, 0, , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                    If lng项目id <> 0 Then
                        Set mobjDetail = GetInputDetail(lng项目id)
                    Else
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.SetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                Bill.TxtVisible = False '(不加不行)
                
                '主项适用病人病区科室
                If InStr(",5,6,7,", mobjDetail.类别) = 0 Then
                    If Not CheckFeeItemLimitDept(mobjDetail.ID, IIf(mbytUseType = 2, UserInfo.部门ID, mobjBill.病区ID), IIf(mbytUseType = 2, UserInfo.部门ID, mobjBill.科室ID)) Then
                        If mbytUseType = 2 Then
                            MsgBox "该收费项目对当前病人病区和科室不适用！", vbInformation, gstrSysName
                        Else
                            MsgBox "该收费项目对当前病人病区和科室不适用！", vbInformation, gstrSysName
                        End If
                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                If InStr(",5,6,7,", mobjDetail.类别) > 0 And mblnNurseStation Then
                    MsgBox "护士站批量记帐不能录入药品项目！", vbInformation, gstrSysName
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
                

                '检查毒理分类和价值分类权限
                If CheckDrugType(mobjDetail) = False Then
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
 
                
                '检查药品输入是否重复:分批及时价同一药房不允许重复(这里只提醒)
                If InStr(",5,6,7,", mobjDetail.类别) > 0 Or _
                    (mobjDetail.类别 = "4" And mobjDetail.跟踪在用) Then
                    If PhysicExist(mobjDetail, Bill.Row) Then
                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                '检查处方职务
                If InStr(",5,6,7,", mobjDetail.类别) > 0 And mbln处方职务检查 Then
                    mobjDetail.处方职务 = Get处方职务(mobjDetail.ID)
                    '所有病人
                    If CheckDuty(mobjDetail, True) > 0 Then
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                '读取药品相关信息
                
                '病人科室ID
                lng病人科室ID = mobjBill.科室ID
                If cbo开单科室.Visible Then
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                Else
                    If lng病人科室ID = 0 Then lng病人科室ID = mlngDeptID
                    If lng病人科室ID = 0 Then lng病人科室ID = GetNurseStationFirstPatiDeptID '护士工作站,取第一个病人科室ID
                    If lng病人科室ID = 0 Then lng病人科室ID = mlng病区ID
                End If
                
                '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
                If mobjDetail.类别 = "4" Then
                    lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, mobjBill.病区ID)
                    If lngDoUnit = 0 Then lngDoUnit = IIf(cbo开单科室.Visible, Get开单科室ID, lng病人科室ID)
                End If
                
                lngDoUnit = Get收费执行科室ID(mobjDetail.类别, mobjDetail.ID, _
                    mobjDetail.执行科室, lng病人科室ID, IIf(cbo开单科室.Visible, Get开单科室ID, lng病人科室ID), 2, lngDoUnit, mobjBill.病区ID)
                
                If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                    '当前行药品库存
                    dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                    If gbln住院单位 Then
                        dblStock = dblStock / mobjDetail.住院包装
                    End If
                    mobjDetail.库存 = dblStock
                    Call ShowStock(mobjDetail.名称, mobjDetail.库存)
          
                ElseIf mobjDetail.类别 = "4" And mobjDetail.跟踪在用 Then
                    dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                    mobjDetail.库存 = dblStock
                    Call ShowStock(mobjDetail.名称, mobjDetail.库存)
                End If
                
                 '处方限量
                If InStr(",5,6,7,", mobjDetail.类别) > 0 And mbln处方限量检查 Then
                    mobjDetail.处方限量 = Get处方限量(mobjDetail.ID)
                End If
                
                '保险项目对应检查
                If CheckInsureTheCode(mobjDetail) = False Then
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
                
                '输入摘要(取已有的行以便修改)
                If mobjBill.Details.Count >= Bill.Row Then
                    If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                        str摘要 = mobjBill.Details(Bill.Row).摘要
                    End If
                End If
                
                '加入或修改该收费细目行
                Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                '59051:先调用GetItemInfor
                '输入摘要(根据新输入的行更改摘要)
                
                mobjBill.Details(Bill.Row).Tag = ""
                If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                    If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                        mobjBill.Details(Bill.Row).摘要 = str摘要
                        mobjBill.Details(Bill.Row).Tag = str摘要
                    End If
                End If
                
                Call CalcMoney(tyPati, Bill.Row)                         '此时,即使是主从项的主项,从项还没有生成
                
                '记帐分类报警(在已经算出该行费用但未显示前)
                If mobjBill.Details.Count = Bill.Row Then
                    If CheckAllPatiChargeWrang(Bill.Row) = False Then
                         Bill.Text = "": Cancel = True: Exit Sub
                    End If
                End If
                
                If mobjBill.Details(Bill.Row).数次 <> 0 Then
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                        mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                        Bill.Text = "": Cancel = True: Exit Sub
                    End If
                End If
                
                Call ShowDetails(Bill.Row)
                Call ShowMoney
                
                '费用类型检查
                Call Check费用类型(Bill.Row)
                Call SetDrawDrugDeptEnabled
                Bill.Text = "": Bill.SetFocus
            End If
            
            If mobjBill.Details.Count >= Bill.Row Then
                mlngPreRow = 0  '修改已有列时,恢复此值,以便显示库存
                With mobjBill.Details(Bill.Row)
                    '下一列的性质确定
                    If .收费类别 = "7" And gblnPay Then Bill.ColData(BillCol.付数) = BillColType.Text  '付数
                    If .收费类别 = "F" Then Bill.ColData(BillCol.标志) = BillColType.CheckBox '附加标志
                    
                    '变价允许输入数次
                    If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                        And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                        Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '数次
                        Bill.ColData(BillCol.单价) = BillColType.Text '单价
                    Else
                        Bill.ColData(BillCol.数次) = BillColType.Text '数次
                        Bill.ColData(BillCol.单价) = BillColType.UnFocus '单价
                    End If
                    
                    '执行科室
                    '在FillBillComboBox中设置ListIndex时调用CboClick事件
                    mblncboEnterCell = True: Bill.Col = BillCol.执行科室: mblncboEnterCell = False
                    Call FillBillComboBox(Bill.Row, BillCol.执行科室, Not blnInput)  '直接回车时保持执行科室
                    mblncboEnterCell = True: Bill.Col = BillCol.项目: mblncboEnterCell = False
                    
                    blnSkip = Bill.ListCount = 1
                    If Not blnSkip And InStr(",4,5,6,7,", .收费类别) > 0 Then
                        '指定了固定药房时,不允许再选择
                        Select Case .收费类别
                            Case "4"
                                blnSkip = glng发料部门 > 0 And .执行部门ID = glng发料部门
                            Case "5"
                                blnSkip = glng西药房 > 0 And .执行部门ID = glng西药房
                            Case "6"
                                blnSkip = glng成药房 > 0 And .执行部门ID = glng成药房
                            Case "7"
                                blnSkip = glng中药房 > 0 And .执行部门ID = glng中药房
                        End Select
                    End If
                    If blnSkip Then
                        Bill.ColData(BillCol.执行科室) = BillColType.UnFocus: .Key = 1
                    Else
                        Bill.ColData(BillCol.执行科室) = BillColType.ComboBox: .Key = Bill.ListCount
                    End If
                    
                    '检查卫生材料的灭菌效期,在确定执行科室之后
                    If .收费类别 = "4" And .Detail.跟踪在用 Then
                        Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                    End If
                    
                     '从属项目处理,仅该行收费项目有从属项目及尚未取才取,药品无需判断,药品不能设置主从项
                    If Bill.TextMatrix(0, Bill.Col) = "项目" And InStr(",5,6,7,", .收费类别) = 0 Then
                        If (gbln从项汇总折扣 And mobjBill.Details(Bill.Row).从属父号 = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                            If ShouldDO(Bill.Row) Then
                               Call SetSubItem
                               mlngPreRow = 0 '通过行变化标志来重新确定列性质
                            End If
                        End If
                    End If
                    
                End With
            End If
            
            '只输入一次付数
            If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                If mobjBill.Details(Bill.Row).收费类别 = "7" Then
                    For i = 1 To Bill.Row - 1
                        If mobjBill.Details(i).收费类别 = "7" Then
                            '正常执行该过程：本身会定位下一个单元,先定位到付数,则下一个单元是数次
                            '选择调用该过程：调用后会送个回车，这里不能再回车，否则是三个回车的效果(控件原因)。
                            Bill.Col = BillCol.付数: Exit For
                        End If
                    Next
                End If
            End If
            
        Case "付数"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                '数字合法性
                If Not IsNumeric(Bill.Text) Then
                    MsgBox "非法数值！", vbInformation, gstrSysName
                    Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                End If
                If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                    MsgBox "付数应该为正的整数！", vbInformation, gstrSysName
                    Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                End If
                
                '最大金额检查
                If gcurMaxMoney > 0 Then
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                        If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                        End If
                    End If
                End If
                
            
                '仅中药及非从属项目才可更改付数(主项付数改变,从属也变)
                If mobjBill.Details(Bill.Row).收费类别 = "7" Then
                    '分批或时价药品不足禁止输入(没有分批的时价药品可以修改付数、数次)
                    If mobjBill.Details(Bill.Row).Detail.分批 Or mobjBill.Details(Bill.Row).Detail.变价 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount) > mobjBill.Details(Bill.Row).Detail.库存 Then
                            MsgBox """" & mobjBill.Details(Bill.Row).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                        End If
                    End If
                          
                    '检查其它时价或分批中药更改付数后库存是否足够
                    For i = 1 To mobjBill.Details.Count
                        If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" _
                            And (mobjBill.Details(i).Detail.变价 Or mobjBill.Details(i).Detail.分批) Then
                            If Val(Bill.Text) * mobjBill.Details(i).数次 * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount) > mobjBill.Details(i).Detail.库存 Then
                                MsgBox "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
                    Next
                                            
                    lngOld付数 = mobjBill.Details(Bill.Row).付数
                    '计算并刷新该行
                    mobjBill.Details(Bill.Row).付数 = Bill.Text
                    
                    If mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            mobjBill.Details(Bill.Row).付数 = lngOld付数
                            Call CalcMoneys(tyPati, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call CalcMoneys(tyPati, Bill.Row)
                    Call ShowDetails(Bill.Row)
                    
                    '处理其它中药付数,如果是独立项,则修改其它非从项的,如果是从项,则修改同一主项的从项的.因为限定为中草药,不可能有主项
                    For i = 1 To mobjBill.Details.Count
                        If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" And mobjBill.Details(i).从属父号 = mobjBill.Details(Bill.Row).从属父号 Then
                            If mobjBill.Details(i).从属父号 = 0 Or (mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).Detail.固有从属 = 0) Then     '1和2固定和按比例的不改
                                mobjBill.Details(i).付数 = Bill.Text
                                Call CalcMoneys(tyPati, i)
                                Call ShowDetails(i)
                            End If
                        End If
                    Next
                    Call ShowMoney
                Else
                    sta.Panels(2) = "从属项目的付数不能更改！"
                    Bill.Text = mobjBill.Details(Bill.Row).付数: Beep '恢复原有付数值
                End If
            End If
        Case "数次"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                 With mobjBill.Details(Bill.Row)
                     '中药快捷输入转换
                    If .收费类别 = "7" Then Bill.Text = ConvertABCtoNUM(Bill.Text)
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = .数次: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = .数次: Cancel = True: Exit Sub
                        End If
                    End If
                    '药品输入小数
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivsOpt, ";药品输入小数;") = 0 Then
                            MsgBox "你没有权限输入小数！", vbInformation, gstrSysName
                            Bill.Text = .数次: Cancel = True: Exit Sub
                        End If
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .付数 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = .数次: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                        dblNum = Val(Bill.Text) * .付数 * .Detail.住院包装
                    Else
                        dblNum = Val(Bill.Text) * .付数
                    End If
                        
                    '负数合法性检查
                    If Val(Bill.Text) * .付数 < 0 Then
                        MsgBox "批量记帐时不允许负数记帐！", vbInformation, gstrSysName
                        Bill.Text = .数次: Cancel = True: Exit Sub
                    End If
                    
                    '药品库存检查
                    If Not CheckDrugStoreIsEnough(FormatEx(.付数 * Val(Bill.Text), 6), mobjBill.Details(Bill.Row)) Then
                        Bill.Text = .数次: Cancel = True: Exit Sub
                    End If
                    
                    dblPreTime = .数次
                    .数次 = Bill.Text
                    
                    '处方限量检查
                    If mbln处方限量检查 And Not gbln处方限量 Then
                        If Not CheckLimit(mobjBill, Bill.Row, gbln住院单位) Then
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If .Detail.录入限量 > 0 And dblNum > .Detail.录入限量 Then
                        If MsgBox("输入的数次超过了录入限量" & FormatEx(.Detail.录入限量 / IIf(gbln住院单位, .Detail.住院包装, 1), 5) & ",是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
          
                    '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                    If .从属父号 <> 0 And .Detail.固有从属 <> 0 Then
                        sta.Panels(2) = "该项目是固有从属项目,其数次不能够更改。"
                        .数次 = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                                        
                    Call CalcMoneys(tyPati, Bill.Row)
                    
                    '数据溢出检查(在已经算出该行费用但未显示前)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                        .数次 = dblPreTime
                        Call CalcMoneys(tyPati, Bill.Row)
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    If .数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            .数次 = dblPreTime
                            Call CalcMoneys(tyPati, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    '记帐分类报警(在已经算出该行费用但未显示前):在最后保存时报警
                End With
                    
                Call ShowDetails(Bill.Row)
                '更改其固有从属的数次
                For i = Bill.Row + 1 To mobjBill.Details.Count
                    If mobjBill.Details(i).从属父号 = Bill.Row Then
                        '28136
                        '如果是输入的负数,需要将下级中的负数集中更新成负数
                        With mobjBill.Details(i)
                            If .Detail.固有从属 = 0 Then  '非固有从属
                                If Abs(.数次) <> Abs(.Detail.从项数次) Then GoTo NotCalc:
                                .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.从项数次
                            ElseIf .Detail.固有从属 = 1 Then '固定的固有从属
                                .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.从项数次 = 0, 1, .Detail.从项数次)
                            ElseIf .Detail.固有从属 = 2 Then   '按比例的固有从属
                                .数次 = Val(Bill.Text) * .Detail.从项数次
                            Else
                                 GoTo NotCalc:
                            End If
                        End With
                        
                        Call CalcMoneys(tyPati, i)
                        Call ShowDetails(i)
NotCalc:
                    End If
                Next

                
                Call ShowMoney
            ElseIf mobjBill.Details.Count >= Bill.Row Then
                If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                    If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: Exit Sub
                    End If
                End If
            End If
                
            If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                If CheckItemHaveSub(Bill.Row) Then
                    KeyCode = 0
                    Call LocateMainItemNextRow(Bill.Row)
                End If
            End If
            
        Case "单价"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                '数字合法性
                If Not IsNumeric(Bill.Text) Then
                    MsgBox "非法数值！", vbInformation, gstrSysName
                    Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                End If
                If Val(Bill.Text) < 0 Then
                    MsgBox "项目价格不应该为负数！", vbInformation, gstrSysName
                    Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                End If
                
                '最大金额检查
                If gcurMaxMoney > 0 Then
                    If Val(Bill.Text) * mobjBill.Details(Bill.Row).付数 * mobjBill.Details(Bill.Row).数次 > gcurMaxMoney Then
                        If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                End If

                Bill.Text = FormatEx(Bill.Text, 5)
                
                '如果没有对应的收入项目,则无法计算
                If mobjBill.Details(Bill.Row).Detail.变价 And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                    If Not (mobjBill.Details(Bill.Row).InComes(1).现价 = 0 And mobjBill.Details(Bill.Row).InComes(1).原价 = 0) Then
                        strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).原价, mobjBill.Details(Bill.Row).InComes(1).现价, CCur(Bill.Text))
                        If strScope <> "" Then
                            sta.Panels(2) = strScope
                            If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).标准单价
                            If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                            Cancel = True: Beep: Exit Sub
                        End If
                    End If
                    
                    dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).标准单价
                    
                    mobjBill.Details(Bill.Row).InComes(1).标准单价 = Bill.Text '这种收费细目只能对应一个收入项目
                    Call CalcMoneys(tyPati, Bill.Row)
                    '记帐分类报警(在已经算出该行费用但未显示前),在最后保存时报警
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                Else
                    Bill.Text = "0"
                    sta.Panels(2) = "该项目设有设置对应的费目，所以无法计算费用！"
                    Beep
                End If
            End If
        Case "执行科室"
            If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                With mobjBill.Details(Bill.Row)
                    If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                        .执行部门ID = Bill.ItemData(Bill.ListIndex)
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                    End If
                    
                    '药品库存检查:动态药房,分批或时价药品也要检查了
                    If Not CheckDrugStoreIsEnough(FormatEx(.付数 * .数次, 6), mobjBill.Details(Bill.Row), True) Then
                        Cancel = True
                    End If
            
                    '检查卫生材料的灭菌效期,在确定执行科室之后
                    If .收费类别 = "4" And .Detail.跟踪在用 Then
                        Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                    End If
                    
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                    
                    Call CalcMoneys(tyPati, Bill.Row, True)
                    Call ShowDetails(Bill.Row)
                    If .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                        Call ShowStock(.Detail.名称, .Detail.库存)
                    End If
                    If mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                End With
            End If
    Case "标志"
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    '注意:在任何exit sub 之前设置mblncboClick = False,否则,无法进入行
    Dim strStock As String, i As Long
    Dim str药房IDs As String
        
    If Not Bill.Active Then Exit Sub
    
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    
    If mblncboEnterCell Then Exit Sub  '避免同一过程中因设置bill的值而循环调用,注意在任何exit sub 之前设置mblncboClick = False
    mblncboEnterCell = True
        
    '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置     mlngPreRow    当前行是否改变
    If zlCheckBill存在非散装草药 = True Then
        '如果单据中存在非散装的,则不能输入
        Call SetBill中草药EditEnabled
        mblncboEnterCell = False
         Exit Sub
    End If
   
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '显示库存
            If InStr(",5,6,7,", .收费类别) > 0 And .收费细目ID <> 0 Then
                If gbln其它药房 Or gbln其它药库 Then
                    strStock = GetStockInfo(.收费细目ID, gbln其它药房, gbln其它药库, gbln住院单位)
                    If strStock <> "" Then
                        If InStr(1, mstrPrivsOpt, ";显示库存;") > 0 Then
                            sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行库存:" & strStock
                        Else
                            sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行有库存."
                        End If
                    End If
                End If
                If strStock = "" Then
                    '更新库存显示
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If gbln住院单位 Then
                        .Detail.库存 = .Detail.库存 / .Detail.住院包装
                    End If
                    Call ShowStock(.Detail.名称, .Detail.库存)
                End If
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                Call ShowStock(.Detail.名称, .Detail.库存)
            Else
                sta.Panels(2) = ""
            End If
                     
            Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.项目) = BillColType.CommandButton
            
             '如果是从属项目的主项目或从项,则不允许更改类别和项目
            If CheckItemHaveSub(Row) Or .从属父号 > 0 Then
                Bill.ColData(BillCol.类别) = BillColType.Text_UnModify
                Bill.ColData(BillCol.项目) = BillColType.Text_UnModify
            End If

            If .收费类别 = "7" And gblnPay Then
                Bill.ColData(BillCol.付数) = BillColType.Text
            Else
                Bill.ColData(BillCol.付数) = BillColType.UnFocus
            End If
            
            '变价允许输入数次
            If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '数次
                Bill.ColData(BillCol.单价) = BillColType.Text '金额
            Else
                Bill.ColData(BillCol.数次) = BillColType.Text
                Bill.ColData(BillCol.单价) = BillColType.UnFocus
            End If
            
            If .Key = "1" Then    '指定了固定药房时,不允许再选择执行科室
                Bill.ColData(BillCol.执行科室) = BillColType.UnFocus
            Else
                Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
            End If
                
            If .收费类别 = "F" Then
                Bill.ColData(BillCol.标志) = BillColType.CheckBox
            Else
                Bill.ColData(BillCol.标志) = BillColType.UnFocus
            End If
            
            '只允许一个类别,不允许选择类别
            If mblnOne Then Bill.ColData(BillCol.类别) = BillColType.UnFocus
        End With
    
        '显示摘要
        If mobjBill.Details(Bill.Row).摘要 <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  摘要:" & mobjBill.Details(Bill.Row).摘要
        End If
    End If
    
    '如果点击未保存的行,则恢复列的性质
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)  '类别列,当主从项时会被改变
        Bill.ColData(BillCol.项目) = BillColType.CommandButton   '项目列,当主从项时会被改变
    End If
    
    
    '-----------------------------------------------------------------
    '2.列改变相关数据处理和显示设置
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then  '加载当前列的下拉项数据
        Call FillBillComboBox(Bill.Row, Bill.Col, True)
    End If
    
    If gbln收费类别 And Bill.TextMatrix(Row, BillCol.类别) = "" And mblnOne Then
        mrsClass.Filter = "编码=" & gstr收费类别
        Bill.TextMatrix(Row, BillCol.类别) = mrsClass!类别
        Bill.RowData(Row) = Asc(mrsClass!编码)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "类别" '不输入类别时不会定位到类别列
            SetWidth Bill.cboHwnd, 70
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "编码=" & gstr收费类别
                    Bill.TextMatrix(Row, Col) = mrsClass!类别
                    Bill.RowData(Row) = Asc(mrsClass!编码)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "执行科室"
            SetWidth Bill.cboHwnd, 130
        Case "付数"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789." & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                    If InStr(mstrPrivsOpt, ";药品输入小数;") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                End If
                '中药快捷输入
                If mobjBill.Details(Bill.Row).收费类别 = "7" Then
                        Bill.TextMask = Bill.TextMask & gstrABC & LCase(gstrABC)
                End If
            End If
        Case "单价"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
    
    '新行,或更改已有行的类别时,看作换行还没有开始
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
    
    mblncboEnterCell = False
End Sub
Private Sub Bill_LostFocus()
    Bill.TxtVisible = False
    Bill.CmdVisible = False
    Bill.CboVisible = False
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        .ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)  '类别列,当主从项时会被改变
        .ColData(BillCol.项目) = BillColType.CommandButton    '项目列,当主从项时会被改变
        .ColData(BillCol.付数) = BillColType.UnFocus   '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
        .ColData(BillCol.单价) = BillColType.UnFocus  '单价缺省跳过,当项目变价时,设为输入(4)
        .ColData(BillCol.标志) = BillColType.UnFocus  '标志缺省跳过,当为手术时,设为复选(-1)
        '针对列编辑性质设置颜色
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
        
        .TextMatrix(Row, BillCol.行) = Row
        '特殊地方手动调用不执行
        If Visible And Bill.Active And Row > 0 And .ColData(BillCol.类别) <> BillColType.UnFocus And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub SetDefaultDoctor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省开单人
    '编制:刘兴洪
    '日期:2015-07-10 15:20:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If cbo开单人.ListCount = 0 Then Exit Sub
    If cbo开单人.ListCount = 1 Then cbo开单人.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.开单部门ID = lng开单部门ID Then Exit Sub
    If mrs领药部门.RecordCount <> 0 Then
        For i = 0 To cboDrawDept.ListCount - 1
             If cboDrawDept.ItemData(i) = lng开单部门ID Then
                mobjBill.领药部门ID = lng开单部门ID
                cboDrawDept.ListIndex = i: Exit For
             End If
        Next
    End If
    
    mobjBill.开单部门ID = lng开单部门ID
        
    '开单科室确定医生
    If Not gblnFromDr Then
        If cbo开单科室.ListIndex <> -1 Then
            If gbln它科人 Then
                Call FillDoctor(cbo开单人, mrs开单人)
            Else
                Call FillDoctor(cbo开单人, mrs开单人, lng开单部门ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo开单人.Clear
        End If
        Call cbo开单人_Click
    End If
    
    
    '重新设置相关项目的执行科室
    If cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '仅处理收费项目
                If InStr(",4,5,6,7,", .Detail.类别) = 0 And .Detail.执行科室 = 6 Then '6-开单人科室
                    .执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    '刷新显示从项执行科室
                    If i <= Bill.Rows - 1 And .执行部门ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                End If
            End With
        Next
    End If
End Sub

     
Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo开单人.ListIndex >= 0 Then lng医生ID = cbo开单人.ItemData(cbo开单人.ListIndex)
    If mrs开单科室 Is Nothing Then Call FillDept(cbo开单科室, mrs开单科室, mrs开单人, mstrPrivs, mbytUseType, mlngDeptID, lng医生ID)
    
    If zlSelectDept(Me, mlngModule, cbo开单科室, mrs开单科室, cbo开单科室.Text) = False Then
        Call Beep: mobjBill.开单部门ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo开单人_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Locked Then Exit Sub
    
    If cbo开单人.Text <> "" Then
        If cbo.FindIndex(cbo开单人, zlStr.NeedName(cbo开单人.Text), True) = -1 Then cbo开单人.ListIndex = -1: cbo开单人.Text = ""
    End If
'    If cbo开单人.Text = "" And mblnKeyReturn Then
'        Call cbo开单人_KeyPress(vbKeyReturn)
'    End If
    mblnKeyReturn = False
    '当开单科室确定开单人时,可能此时不选开单人,先去调整开单科室后再来选
    If gblnFromDr And gbln开单人 And cbo开单人.ListIndex = -1 And mlngSelPatiCount <> 0 Then Cancel = True
End Sub

Private Sub cbo开单人_Click()
    Dim lng开单人ID As Long
 
    If mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text)) Then Exit Sub
    
    mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    If gblnFromDr Then
        If cbo开单人.ListIndex <> -1 Then
            lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
            
            Call FillDept(cbo开单科室, mrs开单科室, mrs开单人, mstrPrivs, mbytUseType, mlngDeptID, lng开单人ID)
            Call SetDefaultDept(cbo开单科室, mrs开单科室, mrs开单人, lng开单人ID)
        Else
            cbo开单科室.Clear
        End If
        Call cbo开单科室_Click
    End If
                        
    '护士类别
    If Bill.Active Then
        If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
            And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '清除无效输入
            Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.类别 Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) '刷新
        End If
    End If
    
    '护士类别:判断非法输入
    If CheckInhibitiveByNurse(mobjBill, mrs开单人) Then
        MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
    End If
End Sub
Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo开单人.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    Dim strAdded As String
    If KeyAscii = vbKeyTab Then
        mblnKeyReturn = True
    End If
    
    If Not KeyAscii = 13 Then Exit Sub
    
    If cbo开单人.Locked Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    
    
    strText = UCase(cbo开单人.Text)
    If cbo开单人.ListIndex <> -1 Then
        '弹出列表时,又在文本框输入了内容
        If strText <> cbo开单人.List(cbo开单人.ListIndex) Then Call zlControl.CboSetIndex(cbo开单人.hWnd, -1)
    End If
    If strText = "" Then
        cbo开单人.ListIndex = -1
    ElseIf cbo开单人.ListIndex = -1 Then
        intIdx = -1
        strFilter = IIf(gbln护士, "人员性质<>''", "人员性质<>'护士'")
        '刘兴洪:22383
        '先复制记录集
        Set rsTemp = zlDatabase.zlCopyDataStructure(mrs开单人)
        Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
        Dim strCompents As String '匹配串
        
        strCompents = Replace(gstrLike, "%", "*") & strText & "*"
        
        If IsNumeric(strText) Then
            intInputType = 0
        ElseIf zlCommFun.IsCharAlpha(strText) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        
        mrs开单人.Filter = strFilter: iCount = 0
        With mrs开单人
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not mrs开单人.EOF
                Select Case intInputType
                Case 0  '输入的是全数字
                    '如果输入的数字,需要检查:
                    '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                    '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                    
                    
                    '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                    If Nvl(!编号) = strText Then strResult = Nvl(!姓名): iCount = 0: Exit Do
                    
                    '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                    If Val(Nvl(!编号)) = Val(strText) Then
                        If iCount = 0 Then strResult = Nvl(!姓名)
                        iCount = iCount + 1
                    End If
                    
                    '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                     If Val(Nvl(!编号)) Like strText & "*" Then
                        If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                            strAdded = strAdded & "," & Nvl(!编号) & ","
                        End If
                     End If
                Case 1  '输入的是全字母
                    '规则:
                    ' 1.输入的简码相等,则直接定位
                    ' 2.根据参数来匹配相同数据
                    
                    '1.输入的简码相等,则直接定位
                    If Trim(Nvl(!简码)) = strText Then
                        If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                        iCount = iCount + 1
                    End If
                    
                    '2.根据参数来匹配相同数据
                    If Trim(Nvl(!简码)) Like strCompents Then
                        If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                            strAdded = strAdded & "," & Nvl(!编号) & ","
                        End If
                    End If
                Case Else  ' 2-其他
                    '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                    '1.编码\简码相等,直接定位
                    '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                    
                    '1.编码\简码相等,直接定位
                    If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                        If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                        iCount = iCount + 1
                    End If
                    
                    '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                    If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                        If isCheck开单人Exists(Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                            strAdded = strAdded & "," & Nvl(!编号) & ","
                        End If
                    End If
                End Select
                mrs开单人.MoveNext
            Loop
        End With
         If iCount > 1 Then strResult = ""
        If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
        '刘兴洪:直接定位
        If strResult <> "" Then
            rsTemp.Close: Set rsTemp = Nothing
            If isCheck开单人Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab: mblnKeyReturn = True
            Exit Sub
        End If
        
        '需要检查是否有多条满足条件的记录
        If rsTemp.RecordCount <> 0 Then
            '先按某种方式进行排序
            Select Case intInputType
            Case 0 '输入全数字
                rsTemp.Sort = "编号"
            Case 1 '输入全拼音
                rsTemp.Sort = "简码"
            Case Else
                '根据选择来定
                If gbyt开单人显示 = 1 Then '简码
                    rsTemp.Sort = "简码"
                Else
                    rsTemp.Sort = "编号"
                End If
            End Select
            '弹出选择器
            Dim rsReturn As ADODB.Recordset
            If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo开单人, rsTemp, True, "", "缺省,职务,优先级别", rsReturn) Then
                If Not rsReturn Is Nothing Then
                    If rsReturn.RecordCount <> 0 Then
                        '进行定位
                        If isCheck开单人Exists(Nvl(rsReturn!姓名), True) Then
                            'zlCommFun.PressKey vbKeyTab
                            mblnKeyReturn = True
                        End If
                    End If
                End If
            End If
        Else
            '未找到
            rsTemp.Close: Set rsTemp = Nothing
            KeyAscii = 0: zlControl.TxtSelAll cbo开单人: Exit Sub
        End If
        rsTemp.Close: Set rsTemp = Nothing
         
    ElseIf Not mblnDrop Then
        '回车光标经过
        Call cbo开单人_Click
        Call zlCommFun.PressKey(vbKeyTab)
        mblnKeyReturn = True
        Exit Sub
    End If
    If cbo开单人.ListIndex = -1 Then
        cbo开单人.Text = ""
        mobjBill.开单人 = ""
        If gblnFromDr Then Exit Sub
    Else
        mobjBill.开单人 = zlStr.NeedName(cbo开单人.Text)
        If intIdx <> -1 And mblnDrop Then
            '弹出回车-强行激活Click
            Call cbo开单人_Click
        ElseIf intIdx <> cbo开单人.ListIndex And intIdx <> -1 Then
            '弹出让选择-自动激活Click
            cbo开单人.SetFocus
            Call zlCommFun.PressKey(vbKeyF4)
            Exit Sub
        ElseIf intIdx <> -1 Then
            '一次性输中-强行激活Click
            Call cbo开单人_Click
        End If
    End If
    Call zlCommFun.PressKey(vbKeyTab)
    mblnKeyReturn = True
End Sub
  
Private Function isCheck开单人Exists(ByVal str姓名 As String, _
    Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '返回:存在返回gtrue,否则返回False
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo开单人.ListCount - 1
        If zlStr.NeedName(cbo开单人.List(i)) = str姓名 Then
            If blnLocateItem Then cbo开单人.ListIndex = i
            isCheck开单人Exists = True
            Exit Function
        End If
    Next
End Function
Private Sub chk加班_Click()
    Dim blnAdd As Boolean
    Dim tyPati As TY_PATIINFOR
    
    If Not chk加班.Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk加班.Value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Checked
        End If
    End If
    
    If chk加班.Value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Unchecked
        End If
    End If
    
    mobjBill.加班标志 = IIf(chk加班.Value = Checked, 1, 0)
    
    '重新计算价格
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys(tyPati)
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    cboNO.Locked = True
End Sub


 
Private Sub SetSubItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入收费项目后,加载当前收费项目的从属项目到费用集对象,并显示在单据控件中
    '编制:刘兴洪
    '日期:2015-07-10 11:48:01
    '调用者:Bill_KeyDown中输入项目后
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, lngMainRow As Long
    Dim lngDoUnit As Long, lng病人科室ID As Long
    Dim bln从项汇总折扣 As Boolean
    Dim str摘要 As String, tyPati As TY_PATIINFOR
    Dim dblStock As Double
    Dim cllData As Collection
    
    lngMainRow = Bill.Row               '主项的行
    If gbln从项汇总折扣 Then            '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
        bln从项汇总折扣 = Not mobjBill.Details(lngMainRow).Detail.屏蔽费别
    End If
    
    lng病人科室ID = mobjBill.科室ID
    If cbo开单科室.Visible Then
        If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        If lng病人科室ID = 0 Then lng病人科室ID = mobjBill.病区ID
    End If
    
    With mobjBill.Details(lngMainRow)
        Set mcolDetails = New Details
        Set mcolDetails = GetSubDetails(.收费细目ID)
        For i = 1 To mcolDetails.Count
            If mobjBill.Details.Count >= Bill.Rows - 1 Then
                Bill.Rows = Bill.Rows + 1
                mblnNewRow = True
                Call bill_AfterAddRow(Bill.Rows - 1)
                mblnNewRow = False
            End If
            Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = "" '有必要加上
            
            'a.从属项目为非药品项目的执行科室
            lngDoUnit = 0
            If InStr(",4,5,6,7,", mcolDetails(i).类别) = 0 Then
                 If mcolDetails(i).类别 = .收费类别 Or mcolDetails(i).执行科室 = 0 Then
                    '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                    '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                    lngDoUnit = .执行部门ID
                 Else
                    '3.其它非药项目的执行科室
                    lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, _
                        mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, 2, , mobjBill.病区ID)
                 End If
            'b.从属项目为药品,卫材的执行科室
            Else
                lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, _
                    mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, 2, .执行部门ID, mobjBill.病区ID)  '卫材从项缺省与主项执行科室相同
            End If
            
            '重新获取库存
            Call SetDetailtStock(lngDoUnit, mcolDetails(i))
     
                       
            '保险项目对应检查
            If CheckInsureTheCode(mcolDetails(i)) = False Then
                Exit Sub
            End If
             
            Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
            
            Call CalcMoney(tyPati, Bill.Rows - 1, bln从项汇总折扣)
            Call ShowDetails(Bill.Rows - 1)
            'CalcMoney中先调用GetuItemInsure可能返回摘要
             str摘要 = mobjBill.Details(Bill.Rows - 1).摘要
        Next
        
        If bln从项汇总折扣 Then
            Call CalcMoney(tyPati, lngMainRow, bln从项汇总折扣) '先重算主项的应收与实收,因为在没有加入从项前可能是按单独打折算的.
            
            Call Calc重算主项实收(lngMainRow)
        End If
        
        Call ShowMoney
    End With
End Sub



Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义主项目的下一行(即从属项）
    '编制:刘兴洪
    '日期:2015-07-10 11:44:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 = lngRow Then
            If mobjBill.Details(i).Detail.固有从属 = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.数次
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位到新行
    '编制:刘兴洪
    '日期:2015-07-10 11:46:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.类别
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.类别
    End If
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetDetailtStock(ByVal lng执行科室ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置明细的库存数据
    '编制：刘兴洪
    '日期：2010-07-12 14:27:51
    '说明：
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim str药房IDs As String, dblStock As Double
    '获取库存
    '不处理非药品或卫材
    If InStr(1, "5,6,7,4", objDetail.类别) = 0 Then Exit Sub
    If objDetail.类别 = "4" And objDetail.跟踪在用 = False Then Exit Sub
    If objDetail.类别 = "4" Then
        '卫材
        dblStock = GetStock(objDetail.ID, lng执行科室ID)
        objDetail.库存 = dblStock
        Exit Sub
    End If
    dblStock = GetStock(objDetail.ID, lng执行科室ID)
    If gbln住院单位 Then
        dblStock = dblStock / objDetail.住院包装
    End If
    objDetail.库存 = dblStock  '记录当前行药品库存
    Exit Sub
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据主项执行科室的变化,刷新非药从项的执行科室
    '编制:刘兴洪
    '日期:2015-07-10 14:52:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lng病人科室ID As Long
    
    With mobjBill
        '获取所有从项及其执行科室类型,必须现取(因为界面上的从项信息可能是修改过的)
        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID)
        
        lng病人科室ID = .科室ID
        If cbo开单科室.Visible Then
            If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        Else
            If lng病人科室ID = 0 Then lng病人科室ID = .病区ID
        End If

        For i = lngRow + 1 To .Details.Count
            If .Details(i).从属父号 = lngRow Then
                '从属项为药品和卫材的项目的执行科室不随主项变动
                If InStr(",4,5,6,7,", .Details(i).收费类别) = 0 Then
                    If .Details(i).收费类别 = .Details(lngRow).收费类别 Then
                        '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                        .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).执行科室 = 0 Then
                                '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                                 .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                            Else
                                '3.其它非药项目的执行科室
                                .Details(i).执行部门ID = Get收费执行科室ID(mcolDetails(j).类别, mcolDetails(j).ID, _
                                    mcolDetails(j).执行科室, lng病人科室ID, Get开单科室ID, 2, , mobjBill.病区ID)
                            End If
                        End If
                    End If
                    
                    '刷新显示从项执行科室
                    If .Details(i).执行部门ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                    
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前行的项目是否具有从属项目
    '入参:lngRow- 指定行
    '编制:刘兴洪
    '日期:2015-07-10 14:53:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function



Private Function CheckInsureVerfyItem(objDetail As Detail, rsVerfyItem As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医保病人费用审批
    '入参:objDetail-当前明细信息
    '     rsVerfyItem-需要审批的项目
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 11:08:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, int险类 As Integer
    Dim str姓名 As String, i As Long, lng主页ID As Long
    Dim str医疗付款方式 As String
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    On Error GoTo errHandle
    
    If rsVerfyItem Is Nothing Then CheckInsureVerfyItem = True: Exit Function
    If rsVerfyItem.State <> 1 Then CheckInsureVerfyItem = True: Exit Function
    rsVerfyItem.Filter = 0
    If rsVerfyItem.RecordCount = 0 Then CheckInsureVerfyItem = True: Exit Function
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng病人ID = Val(rptPati.Rows(i).Record(COL_病人ID).Value)
            lng主页ID = Val(rptPati.Rows(i).Record(COL_主页ID).Value)
            int险类 = Val(rptPati.Rows(i).Record(COL_险类).Value)
            str姓名 = rptPati.Rows(i).Record(COL_姓名).Value
            str医疗付款方式 = rptPati.Rows(i).Record(COL_医疗付款方式).Value
            
            If lng病人ID <> 0 And int险类 <> 0 Then
                bln医保 = False: bln公费 = False
                Call zlIsCheckMedicinePayMode(str医疗付款方式, bln医保, bln公费)
                If bln医保 Then
                    Set mrsMedAudit = GetAuditRecord(lng病人ID, lng主页ID)
                Else
                    Set mrsMedAudit = Nothing
                End If
                If Not mrsMedAudit Is Nothing Then
                    rsVerfyItem.Filter = "收费细目ID=" & mobjDetail.ID & " and 险类=" & int险类
                    If rsVerfyItem.RecordCount = 0 Then CheckInsureVerfyItem = True: Exit Function
                    
                    mrsMedAudit.Filter = "项目ID=" & mobjDetail.ID
                    If mrsMedAudit.RecordCount = 0 Then
                        MsgBox "病人:" & str姓名 & " 未被批准使用[" & mobjDetail.名称 & "]！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If Not IsNull(mrsMedAudit!可用数量) Then
                        If mrsMedAudit!可用数量 <= 0 Then
                            MsgBox "病人:" & str姓名 & "　使用[" & mobjDetail.名称 & "]已达到批准的使用限量" & FormatEx(mrsMedAudit!使用限量 / IIf(gbln住院单位, mobjDetail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    CheckInsureVerfyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDrugType(objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查毒理分类和价值分类权限
    '入参:objDetail-当前明细信息
    '返回:合法成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 11:28:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs药品信息 As ADODB.Recordset
    On Error GoTo errHandle
    If InStr(",5,6,7,", objDetail.类别) = 0 Then CheckDrugType = True: Exit Function
    
    Set rs药品信息 = Read药品信息(objDetail.ID)
    If Not rs药品信息 Is Nothing Then
        If IIf(IsNull(rs药品信息!毒理分类), "", rs药品信息!毒理分类) = "麻醉药" _
            And InStr(mstrPrivsOpt, ";麻醉药品记帐;") = 0 Then
            MsgBox """" & mobjDetail.名称 & """为麻醉药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf IIf(IsNull(rs药品信息!毒理分类), "", rs药品信息!毒理分类) = "毒性药" _
            And InStr(mstrPrivsOpt, ";毒性药品记帐;") = 0 Then
            MsgBox """" & mobjDetail.名称 & """为毒性药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf (IIf(IsNull(rs药品信息!价值分类), "", rs药品信息!价值分类) = "贵重" _
            Or IIf(IsNull(rs药品信息!价值分类), "", rs药品信息!价值分类) = "昂贵") _
            And InStr(mstrPrivsOpt, ";贵重药品记帐;") = 0 Then
            MsgBox """" & mobjDetail.名称 & """为贵重或昂贵药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDrugType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function CheckInsureTheCode(objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保险对码检查
    '入参:objDetail-当前明细信息
    '返回:存在对码返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 11:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng病人ID As Long, intInsure As Integer, str姓名 As String
    Dim strInsures As String, strPriceGrade As String
    On Error GoTo errHandle
    
    strInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng病人ID = Val(rptPati.Rows(i).Record(COL_病人ID).Value)
            intInsure = Val(rptPati.Rows(i).Record(COL_险类).Value)
            str姓名 = rptPati.Rows(i).Record(COL_姓名).Value
            If lng病人ID <> 0 And intInsure <> 0 Then
                If InStr(strInsures & ",", "," & intInsure & ",") = 0 Then
                    If InStr(",5,6,7,", objDetail.类别) > 0 Then
                        strPriceGrade = mstr药品价格等级
                    ElseIf objDetail.类别 = "4" Then
                        strPriceGrade = mstr卫材价格等级
                    Else
                        strPriceGrade = mstr普通价格等级
                    End If
                    If Not CheckMediCareItem(objDetail.ID, intInsure, objDetail.名称, objDetail.变价 = False, True, strPriceGrade) Then Exit Function
                    strInsures = strInsures & "," & intInsure
                End If
            End If
        End If
    Next
    CheckInsureTheCode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function InputItemMemo(tyPati As TY_PATIINFOR) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入摘要
    '入参:tyPati-病人信息
    '返回:输入成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 11:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str摘要 As String
    On Error GoTo errHandle
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            str摘要 = ""
'            If tyPati.险类 <> 0 And .Detail.补充摘要 = False Then '90304
            If .Detail.补充摘要 = False Then
                str摘要 = gclsInsure.GetItemInfo(tyPati.险类, tyPati.病人ID, .Detail.ID, str摘要, 2)
            End If
            .摘要 = str摘要
        End With
    Next
    InputItemMemo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function CheckAllPatiChargeWrang(Optional ByVal lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查所有的病人的报警
    '入参:lngRow:当前行,-1表示所有
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str病人IDs As String
    Dim tyPati As TY_PATIINFOR
    
    mrsWarn.Filter = ""
    If mrsWarn.RecordCount = 0 Then CheckAllPatiChargeWrang = True: Exit Function
    If mlngSelPatiCount = 0 Then CheckAllPatiChargeWrang = True: Exit Function
    On Error GoTo errHandle
    
    str病人IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            If tyPati.病人ID <> 0 Then
                If InStr(str病人IDs & ",", "," & tyPati.病人ID & ",") = 0 Then
                    If CheckPatiChargeWrang(tyPati) = False Then Exit Function
                    str病人IDs = str病人IDs & "," & tyPati.病人ID
                End If
            End If
        End If
    Next
    CheckAllPatiChargeWrang = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPatiChargeWrang(tyPati As TY_PATIINFOR, _
      Optional blnSavePriceBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定病人的报警
    '入参:tyPati-病人信息
    '出参:blnSavePriceBill-是否保存为划价单
    '返回:合法返回true(包含提示选择继续),否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '记帐分类报警(在已经算出该行费用但未显示前)
    Dim curTotal As Currency, curItemMoney As Currency
    Dim cur余额 As Currency, str类别 As String, str类别名称 As String
    Dim rsTmp As ADODB.Recordset, cur当日额 As Currency
    Dim i As Long
    
    
    blnSavePriceBill = False
    mrsWarn.Filter = ""
    If mrsWarn.RecordCount = 0 Then CheckPatiChargeWrang = True: Exit Function
    On Error GoTo errHandle
    
    curTotal = CalcGridToTal(False) ' GetAllPatiTotal(tyPati.病人ID, -1, -1, mobjBill)
    If curTotal <= 0 Then CheckPatiChargeWrang = True: Exit Function
    
    '刷新病人预交款信息
    Set rsTmp = GetMoneyInfo(tyPati.病人ID, 0, True, 2)
    
    If Not rsTmp Is Nothing Then
        cur余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
    End If
    
    '重新读取当日额
    cur当日额 = GetPatiDayMoney(tyPati.病人ID)
    
    If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, tyPati.病人ID)
    
    '在已确认是记帐保存时,以正常的方式报警。
    '如果是划价模式,因为无按钮设置,则可以以新的方式报警。
    For i = 1 To mobjBill.Details.Count
        
        gbytWarn = BillingWarn(mstrPrivsOpt, tyPati.姓名 & IIf(tyPati.住院号 = "", "", "(住院号:" & tyPati.住院号 & " 床号:" & tyPati.床号 & ")"), mlng病区ID, tyPati.适用病人, mrsWarn, cur余额, cur当日额, curTotal, _
                     tyPati.担保额, mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn, , gblnPrice And gbytBilling = 1)

        
        '返回:0;没有报警,继续
        '     1:报警提示后用户选择继续
        '     2:报警提示后用户选择中断
        '     3:报警提示必须中断
        '     4:强制记帐报警,继续
        '     5.报警提示后用户选择继续,但只允许保存存为划价单
        '     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
        
        Select Case gbytWarn
        Case 2, 3 '报警提示后用户选择中断和报警提示必须中断
            Exit Function
        Case 1, 4   '报警提示后用户选择继续,强制记帐报警,继续
            CheckPatiChargeWrang = True: Exit Function
        Case 5 '报警提示后用户选择继续,但只允许保存存为划价单
            blnSavePriceBill = True:    CheckPatiChargeWrang = True
            Exit Function
        Case Else
        End Select
    Next
    CheckPatiChargeWrang = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
'
'Private Function GetAllPatiTotal(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng婴儿 As Long, _
'    objBill As ExpenseBill) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:获取指定病人单据费目合计金额
'    '入参:lng病人ID-病人ID
'    '     lng主页ID-主页ID(-1 时，代表所有)
'    '     lng婴儿-第几个婴儿的费用(-1时,表示所有(含婴儿),0时,代表病人本人 >0时表示第几个婴儿)
'    '     objBill-单据对象
'    '出参:
'    '返回:返回指定病人的合计金额
'    '编制:刘兴洪
'    '日期:2015-07-09 14:50:57
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim objBillDetail As New BillDetail
'    Dim curMoney As Currency
'
'    On Error GoTo errHandle
'    For Each objBillDetail In objBill.Details
'        curMoney = GetPatiBillRowTotal(lng病人ID, objBillDetail.InComes, lng主页ID, lng婴儿)
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
'
'Private Function GetPatiBillRowTotal(ByVal lng病人ID As Long, objBillInComes As BillInComes, _
'    Optional ByVal lng主页ID As Long = 1, Optional ByVal lng婴儿 As Long = -1) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:获取指定行的汇总金额
'    '入参:lng病人ID-病人ID
'    '     lng主页ID-主页ID(-1 时，代表所有)
'    '     lng婴儿-第几个婴儿的费用(-1时,表示所有(含婴儿),0时,代表病人本人 >0时表示第几个婴儿)
'    '     objBillInComes-单据对象行对象
'    '返回:返回单据指定行的合计金额
'    '编制:刘兴洪
'    '日期:2015-07-09 15:37:00
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim objBillIncome As New BillInCome
'    Dim curMoney As Currency
'    On Error GoTo errHandle
'
'    For Each objBillIncome In objBillInComes
'        curMoney = GetPati实收金额(lng病人ID, objBillIncome, lng主页ID, lng婴儿)
'        GetPatiBillRowTotal = GetPatiBillRowTotal + curMoney
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
'
'
'Private Function GetPati实收金额(ByVal lng病人ID As Long, _
'    objBillIncome As BillInCome, _
'    Optional ByVal lng主页ID As Long = -1, _
'    Optional ByVal lng婴儿 As Long = -1) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:获取指定行单据金额
'    '入参:lng病人ID-病人ID
'    '     lng主页ID-主页ID(-1代表所有)
'    '     lng婴儿-第几个婴儿的费用(-1时,表示所有(含婴儿),0时,代表病人本人 >0时表示第几个婴儿)
'    '     objBill-单据对象
'    '出参:
'    '返回:返回指定病人的合计金额
'    '编制:刘兴洪
'    '日期:2015-07-09 14:54:42
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllData As Collection, i As Long
'    On Error GoTo errHandle
'    If objBillIncome Is Nothing Then Exit Function
'    If UCase(TypeName(objBillIncome.Tag)) = UCase("Empty") Then Exit Function
'    If UCase(TypeName(objBillIncome.Tag)) <> UCase("Collection") Then Exit Function
'
'    '病人ID,主页ID,第几个婴儿(0时代表病人本人),实收金额
'    Set cllData = objBillIncome.Tag
'    For i = 1 To cllData.Count
'       If cllData(i)(0) = lng病人ID _
'            And (cllData(i)(1) = lng主页ID Or lng主页ID = -1) _
'            And (cllData(i)(2) = lng婴儿 Or lng婴儿 = -1) Then
'            GetPati实收金额 = GetPati实收金额 + Val(cllData(i)(3))
'       End If
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据输入的有效性
    '返回: 数据有效返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 10:58:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl数次 As Double, strTmp As String
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim colStock As Collection, dblTotal As Double
    Dim str收费细目IDs As String, rsVerfyItem As ADODB.Recordset
    Dim tyPati  As TY_PATIINFOR
    
    On Error GoTo errHandle
    
    mstr病人IDs = GetPatiIDsBySel(mlngSelPatiCount)
    '检查是否选择了病人
    If mlngSelPatiCount = 0 Then
        MsgBox "没有选择要批量记帐的病人,请选择后再点［确定］！", vbInformation, gstrSysName
        If rptPati.Visible Then rptPati.SetFocus
        Exit Function
    End If
    
    If mobjBill.Details.Count = 0 Then
        MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
            

    i = Check执行科室
    If i <> 0 Then
        MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
    
    If mblnNurseStation Then
        For i = 0 To rptPati.Rows.Count - 1
            If rptPati.Rows(i).Record.Tag = "1" Then
                tyPati = GetPatiInforByReport(i, mblnNurseStation)
                If tyPati.开单人 = "" And gbln开单人 Then
                    MsgBox "请确定病人" & tyPati.姓名 & "的开单人！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(tyPati.开单科室ID) = 0 Then
                    MsgBox "请确定病人" & tyPati.姓名 & "的开单科室！", vbInformation, gstrSysName
                    Exit Function
                End If
                If mbln补费 Then
                    If mlngDeptID <> Val(tyPati.开单科室ID) Then
                        MsgBox "注意:" & vbCrLf & "    开单科室不是病人" & tyPati.姓名 & "转科的科室,不能进行补费操作!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    If mobjBill.开单部门ID = 0 And cbo开单科室.Visible Then
        MsgBox "请确定开单科室！", vbInformation, gstrSysName
        If cbo开单科室.Enabled And cbo开单科室.Visible Then cbo开单科室.SetFocus
        Exit Function
    End If
    If mobjBill.开单人 = "" And gbln开单人 And cbo开单人.Visible Then
        MsgBox "请输入开单人！", vbInformation, gstrSysName
        If cbo开单人.Enabled And cbo开单人.Visible Then cbo开单人.SetFocus
        Exit Function
    End If
    
    '护士类别:判断非法输入
    If CheckInhibitiveByNurse(mobjBill, mrs开单人) Then
        MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
        
    '发生时间检查
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    
    '按病人检查相关数据
    If CheckAllPatiIsValied = False Then Exit Function
    
    '补费检查
    If mbln补费 And cbo开单科室.Visible Then
        If cbo开单科室.ItemData(cbo开单科室.ListIndex) <> mlngDeptID Then
            MsgBox "注意:" & vbCrLf & "    开单科室不是病人转科的科室,不能进行补费操作!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
            
    
    '输入的负数检查
    If CheckIsExistsNegativeNums = False Then Exit Function

    mblnSendMateria = False
    dbl数次 = 0: strTmp = ""
    For i = 1 To mobjBill.Details.Count
        If InStr(1, str收费细目IDs & ",", "," & mobjBill.Details(i).收费细目ID & ",") = 0 Then
            str收费细目IDs = str收费细目IDs & "," & mobjBill.Details(i).收费细目ID
        End If
        If mobjBill.Details(i).数次 <> 0 And dbl数次 = 0 Then
            dbl数次 = mobjBill.Details(i).数次
        End If
        If mobjBill.Details(i).收费细目ID = 0 Then
            MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Function
        ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
            '收集药品的发药药房
            strTmp = strTmp & "," & mobjBill.Details(i).收费细目ID
        End If
        
        
        '主项适用病人病区科室
        If InStr(",5,6,7,", mobjBill.Details(i).收费类别) = 0 Then
            If CheckItemHaveSub(i) Then
                If Not CheckFeeItemLimitDept(mobjBill.Details(i).收费细目ID, IIf(mbytUseType = 2, UserInfo.部门ID, mobjBill.病区ID), IIf(mbytUseType = 2, UserInfo.部门ID, mobjBill.科室ID)) Then
                    If mbytUseType = 2 Then
                        MsgBox "第" & i & "行的收费项目对你所在的科室不适用！", vbInformation, gstrSysName
                    Else
                        MsgBox "第" & i & "行的收费项目对当前病人病区和科室不适用！", vbInformation, gstrSysName
                    End If
                    Bill.Row = i: Bill.MsfObj.TopRow = i
                    Bill.Col = BillCol.项目: Bill.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        '检查分批或时价药品同一药房是否有重复输入
        With mobjBill.Details(i)
            If (.Detail.分批 Or .Detail.变价) _
                And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                For j = 1 To mobjBill.Details.Count
                    If i <> j And .收费细目ID = mobjBill.Details(j).收费细目ID And .执行部门ID = mobjBill.Details(j).执行部门ID Then
                        If .收费类别 = "4" Then
                            MsgBox "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入，请合并！", vbInformation, gstrSysName
                        Else
                            MsgBox "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入，请合并！", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                Next
            End If
        End With
        
        '检查自动发药
        If CheckAutoSendDrugAndStuff(mobjBill.Details(i), False, mblnSendMateria) = False Then Exit Function
    Next
    If InStr(mstrPrivsOpt, ";药品发药;") = 0 Then mblnSendMateria = False
    
    '27467,52828
    If FormatEx(dbl数次, 7) = 0 Then
        MsgBox "单据中至少要有一条不为零的数次,请检查！", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    
    '检查药品的发药药房对应的服务科室(存储库房)
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
        Set rsTmp = GetServiceDept(strTmp)
        
        If Not rsTmp Is Nothing Then
            strTmp = ""
            For i = 1 To mobjBill.Details.Count
            
                
                If InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                    strInfo = mobjBill.Details(i).收费细目ID
                    '先检查是否是允许的存储库房
                    rsTmp.Filter = "收费细目ID=" & strInfo & " And 执行科室id=" & mobjBill.Details(i).执行部门ID
                    If rsTmp.RecordCount = 0 Then
                        strTmp = strTmp & "," & i
                    Else
                        '再检查是否是允许的服务科室(没有设置服务科室的,开单科室ID为零)
                        rsTmp.Filter = "(" & rsTmp.Filter & " And 开单科室ID=" & mobjBill.科室ID & ") Or (" & rsTmp.Filter & " And 开单科室ID=0)"
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        End If
                    End If
                End If

 
            Next
            If strTmp <> "" Then
                strTmp = Mid(strTmp, 2)
                MsgBox "请检查,第" & strTmp & "行药品是否违反以下规则:" & vbCrLf & vbCrLf & _
                    "A.选择的执行科室不是药品的存储库房" & vbCrLf & _
                    "B.病人科室[" & GET部门名称(mobjBill.科室ID, mrs开单科室) & "]不属于药品在此存储库房的服务科室.", _
                    vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    '所有病人项目
    i = CheckDuty(, True)
    If i > 0 Then
        Bill.Row = i: Bill.MsfObj.TopRow = i
        Bill.Col = BillCol.项目: Bill.SetFocus
        Exit Function
    End If
 

    
    '药品禁忌检查
    strInfo = CheckDisable(mobjBill)
    If strInfo <> "" Then
        If strInfo Like "*(互相禁用)*" Then
            MsgBox strInfo, vbInformation, gstrSysName
            Exit Function
        End If
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
                    
    '处方限量检查
    If Not gbln处方限量 And mbln处方限量检查 Then
        If Not CheckLimit(mobjBill, , gbln住院单位) Then Exit Function
    End If
    
    '获取医保病人需要审批的费用
    If str收费细目IDs <> "" Then
        str收费细目IDs = Mid(str收费细目IDs, 2)
        Call Get要求审批(str收费细目IDs, rsVerfyItem)
    End If
                
    
    '药品库存检查,71188:刘尔旋,2014-04-03,对不足提醒的也要进行检查
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
            If InStr(",5,6,7,", .收费类别) > 0 Then
                If .Detail.分批 Or .Detail.变价 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
             
                ElseIf colStock("_" & .执行部门ID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                    If dblTotal > .Detail.库存 Then
                        MsgBox "第 " & i & " 行药品""" & .Detail.名称 & _
                            """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .执行部门ID) = 1 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                    
                    If dblTotal > .Detail.库存 Then
                        If MsgBox("第 " & i & " 行药品""" & .Detail.名称 & _
                            """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,要继续吗?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                If .Detail.分批 Or .Detail.变价 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If dblTotal > .Detail.库存 Then
                        MsgBox "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                            """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .执行部门ID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    
                    If dblTotal > .Detail.库存 Then
                        MsgBox "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                            """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .执行部门ID) = 1 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If dblTotal > .Detail.库存 Then
                        If MsgBox("第 " & i & " 行卫生材料""" & .Detail.名称 & _
                            """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,要继续吗?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            End If
            
            '检查医保审批
            If Not CheckInsureVerfyItem(.Detail, rsVerfyItem) Then Exit Function
        End With
    Next
    
    '零差价检查,105875
    If Not gobjPublicDrug Is Nothing Then
        'Private Function zlCheckPriceAdjustBySell(ByVal lng药品id As Long, ByVal lng药房id As Long) As Boolean
        '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
        '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
        '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
        '销售出库时只判断药房
        '返回：True-正常进行销售出库；false-不能进行销售出库
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If gobjPublicDrug.zlCheckPriceAdjustBySell(.收费细目ID, .执行部门ID) = False Then
                        Exit Function
                    End If
                End If
            End With
        Next
    End If
         
    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling)) = False Then
        Exit Function
    End If
        
    '刘兴洪:22441,检查主手术和附加手术情况
    If CheckMainOperation = False Then Exit Function
    If mblnSendMateria And gbytSendMateria = 2 Then
        If MsgBox("记帐完成后自动执行发药吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mblnSendMateria = False
        End If
    End If
    
    mblnPrintDrugList = False
    If mblnSendMateria Then
        mblnPrintDrugList = MsgBox("单据发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckPatiIsValied(tyPati As TY_PATIINFOR, _
    objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人单据的有效性
    '入参:typati-病人信息
    '     objBill-单据信息
    '返回:有效返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 16:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim lngInsure As Long
    
    
    '1.费用时间相关检查
    If CheckPatiFeeDateIsValied(tyPati) = False Then Exit Function
    
    '2.补费检查是否超过时限
    If mbln补费 Then
        If zlCheckPatiFeeRenewValied(tyPati.病人ID, tyPati.主页ID, mobjBill.病区ID, mobjBill.科室ID, mstr最后转科时间) = False Then Exit Function
        
        If txtDate.Text > mstr最后转科时间 And mstr最后转科时间 <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    病人:" & tyPati.姓名 & " 补录的费用时间超过了最后转出的时间(" & mstr最后转科时间 & "),不能进行补费操作!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
    
    '3.医保检查
    If InsureItemCheck(tyPati, mobjBill) = False Then Exit Function
    
    '4.检查病人是否能进行记帐:出院强制记帐权限检查
    If Not PatiCanBilling(tyPati.病人ID, tyPati.主页ID, mstrPrivsOpt) Then Exit Function

    '5.检查病人是否允许变动
    If zlIsAllowFeeChange(tyPati.病人ID, tyPati.主页ID, , tyPati.姓名) = False Then Exit Function
    '6.检查病案是否已经编目
    If zlPatiIS病案已编目(tyPati.病人ID, tyPati.主页ID) = True Then Exit Function
    
    '7.检查是否审批
    '   要求审批,医保费用项目是否审批检查,输入时已检查，保存时再检查是因为：
    '   1).先输单据再确定医保身份；2).主从项批量添加时只检查了主项；3).导入单据时未检查,
    '   4).通过配方输入时未检查
    If tyPati.险类 <> 0 And Not mrsMedAudit Is Nothing Then
        lngInsure = tyPati.险类
        If Not CheckExamine(mobjBill.Details, mrsMedAudit, lngInsure, tyPati.姓名) Then Exit Function
    End If
    
    
    CheckPatiIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub reSetBillObject(tyPatiInfor As TY_PATIINFOR, _
    ByRef objBill As ExpenseBill)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置单据对象
    '入参:tyPatiInfor-病人信息
    '编制:刘兴洪
    '日期:2015-07-09 17:17:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With tyPatiInfor
        objBill.标识号 = tyPatiInfor.住院号
        objBill.病人ID = tyPatiInfor.病人ID
        objBill.主页ID = tyPatiInfor.主页ID
        objBill.床号 = tyPatiInfor.床号
        objBill.费别 = tyPatiInfor.费别
        objBill.年龄 = tyPatiInfor.年龄
        objBill.性别 = tyPatiInfor.性别
        objBill.姓名 = tyPatiInfor.姓名
        objBill.婴儿费 = tyPatiInfor.婴儿
        '重新计算实收金额
        
    End With
End Sub

Private Function GetPatiInforByReport(ByVal lngRow As Long, Optional blnNurseStation As Boolean = False) As TY_PATIINFOR
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据报表指定的行获取病人信息
    '入参:lngRow-指定的行
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 17:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyPati As TY_PATIINFOR
    
    On Error GoTo errHandle
    With rptPati.Rows(lngRow)
        tyPati.病人ID = Val(.Record(COL_病人ID).Value)
        tyPati.主页ID = Val(.Record(COL_主页ID).Value)
        tyPati.险类 = Val(.Record(COL_险类).Value)
        tyPati.婴儿 = Val(.Record(COL_婴儿).Value)
        tyPati.姓名 = .Record(COL_姓名).Value
        tyPati.性别 = .Record(COL_性别).Value
        tyPati.年龄 = .Record(COL_年龄).Value
        tyPati.费别 = .Record(COL_费别).Value
        tyPati.当日额 = Val(.Record(COL_当日额).Value)
        tyPati.剩余款 = Val(.Record(COL_剩余款).Value)
        tyPati.担保额 = Val(.Record(COL_担保额).Value)
        tyPati.适用病人 = .Record(COL_适用病人).Value
        tyPati.住院号 = .Record(COL_住院号).Value
        tyPati.床号 = .Record(COL_床号).Value
        tyPati.保险类别 = .Record(COL_保险类别).Value
        tyPati.医疗付款方式 = .Record(COL_医疗付款方式).Value
        tyPati.入院日期 = ""
        tyPati.出院日期 = ""
        tyPati.状态 = 0
        If Not mrsPati Is Nothing Then
            mrsPati.Filter = "病人ID=" & tyPati.病人ID
            If Not mrsPati.EOF Then
                tyPati.入院日期 = Format(mrsPati!入院日期, "yyyy-MM-DD HH:MM:SS")
                tyPati.出院日期 = Format(mrsPati!出院日期, "yyyy-MM-DD HH:MM:SS")
                tyPati.状态 = Val(Nvl(mrsPati!状态))
                tyPati.病人性质 = Val(Nvl(mrsPati!病人性质))
            End If
        End If
        If blnNurseStation = True Then
            tyPati.开单人 = .Record(COL_开单人).Value
            tyPati.开单科室ID = .Record(COL_开单科室ID).Value
        End If
    End With
    GetPatiInforByReport = tyPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureItemCheck(tyPati As TY_PATIINFOR, _
    ByVal objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保相关检查
    '入参:objBill-票据对象
    '编制:刘兴洪
    '日期:2015-07-09 16:48:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If tyPati.险类 = 0 Then InsureItemCheck = True: Exit Function
    If Not gclsInsure.GetCapability(support实时监控, tyPati.病人ID, tyPati.险类) Then InsureItemCheck = True: Exit Function

    On Error GoTo errHandle
    If gclsInsure.CheckItem(tyPati.险类, 1, 0, MakeDetailRecord(objBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling)) = False Then Exit Function
    InsureItemCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 Private Function CheckIsExistsNegativeNums() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查允许负数记帐
    '返回:允许返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 10:36:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsNegative As Boolean '存在负数
    Dim dblNums As Double, lngTempRow As Long
    Dim bln负数记帐 As Boolean
    Dim i As Long
    
    On Error GoTo errHandle

    lngTempRow = 0
    With mobjBill
        For i = 1 To .Details.Count
            blnIsNegative = .Details(i).付数 * .Details(i).数次 < 0
            lngTempRow = i: If blnIsNegative Then Exit For
        Next
    End With
    If blnIsNegative Then
        MsgBox "批量记帐不允许进行负数记帐(在" & lngTempRow & "行)！", vbInformation, gstrSysName
        If Bill.Rows - 1 >= Bill.Row Then Bill.Row = lngTempRow
    End If
    CheckIsExistsNegativeNums = True: Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDrugStoreIsEnough(ByVal dblNums As Double, _
    objBillDetail As BillDetail, Optional ByVal bln重读库存 As Boolean = False, _
    Optional lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否充足
    '入参:dblNums-药品数量
    '    objBillDetail-药品明细
    '返回:充足返回返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 11:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTempNums As Double
    Dim colStock As Collection
    
    On Error GoTo errHandle
    
    dblTempNums = dblNums * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount)
    
    
    With objBillDetail
        If Not (.收费类别 = "4" And .Detail.跟踪在用) Or (InStr(",5,6,7,", .收费类别) > 0) Then CheckDrugStoreIsEnough = True: Exit Function
        
        If dblNums = 0 Then
            dblNums = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
            dblTempNums = dblNums * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount)
        End If
        
        If .Detail.分批 Or .Detail.变价 Then
            '分批或时价药品不足禁止输入
            If bln重读库存 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
            End If
            If dblTempNums <= .Detail.库存 Then CheckDrugStoreIsEnough = True: Exit Function
            
            If .收费类别 = "4" Then
                If lngRow > 0 Then
                    MsgBox "第 " & lngRow & " 行时价或分批卫生材料""" & .Detail.名称 & _
                        """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTempNums & """。", vbInformation, gstrSysName
                    Exit Function
                Else
                    MsgBox """" & .Detail.名称 & """为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                End If
            Else
                 If lngRow > 0 Then
                    MsgBox "第 " & lngRow & " 行时价或分批药品""" & .Detail.名称 & _
                        """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTempNums & """。", vbInformation, gstrSysName
                    Exit Function
                 Else
                    MsgBox """" & .Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                 End If
            End If
            Exit Function
        End If
    
        Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
        If colStock("_" & .执行部门ID) <> 0 _
            And Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
            
            If dblTempNums <= .Detail.库存 Then CheckDrugStoreIsEnough = True: Exit Function
        
            If colStock("_" & .执行部门ID) = 1 Then
                If MsgBox("""" & .Detail.名称 & """的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf colStock("_" & .执行部门ID) = 2 Then
                MsgBox """" & .Detail.名称 & """的当前可用库存不足输入数量！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    CheckDrugStoreIsEnough = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckAllPati审批量(ByVal dblNum As Double, objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查审批量
    '入参:dblNum-输入量
    '     objBillDetail-单据明细数据
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInsures As String, tyPati As TY_PATIINFOR
    Dim i As Long
    
    On Error GoTo errHandle
    strInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            If tyPati.险类 <> 0 And InStr(strInsures & ",", "," & tyPati.险类 & ",") = 0 Then
              
              If CheckPati审批量(dblNum, tyPati, objBillDetail) = False Then Exit Function
              strInsures = strInsures & "," & tyPati.险类
    
            End If
        End If
    Next
    CheckAllPati审批量 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPati审批量(ByVal dblNum As Double, tyPati As TY_PATIINFOR, objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定病人的审批量
    '入参:dblNum-输入量
    '     objBillDetail-单据明细数据
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 11:23:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If tyPati.险类 = 0 Then CheckPati审批量 = True: Exit Function
    If mrsMedAudit Is Nothing Then CheckPati审批量 = True: Exit Function
    With objBillDetail
        If Not .Detail.要求审批 Then CheckPati审批量 = True: Exit Function
        mrsMedAudit.Filter = "项目ID=" & .收费细目ID
        If mrsMedAudit.RecordCount = 0 Then CheckPati审批量 = True: Exit Function
        If IsNull(mrsMedAudit!可用数量) Then CheckPati审批量 = True: Exit Function
        If dblNum > mrsMedAudit!可用数量 Then
            MsgBox "输入的数次超过了批准的可用数量" & FormatEx(mrsMedAudit!可用数量 / IIf(gbln住院单位, .Detail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckPati审批量 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是手术输入情况(如果不存在主要手术,但存在附加手术,则禁止
    '入参:
    '出参:lngRow-返回附加手术的行
    '返回:存在主手术或没有输入附加手术,返回true,否则返回False
    '编制:
    '修改:刘兴洪(退号时,增加定位功能),增加参数;strBackNo
    '日期:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   '指定行
    Dim i As Long
    
    lngCount = 0
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "F" Then
           If mobjBill.Details(i).附加标志 = 0 Then CheckMainOperation = True: Exit Function     '存在主要手术,则不检查,直接返回true
           lngCount = lngCount + 1  '表示附加手术
           If lngRow <= 0 Then lngRow = i
        End If
    Next
    If lngCount <> 0 Then
          MsgBox "单据中不存主要手术,但存在附加手术,请检查！", vbInformation, gstrSysName
          If Bill.Rows > lngRow Then Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
End Function





  
 

Private Sub Calc重算主项实收(ByVal lngMainRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当从项汇总折扣时,根据指定的主项的行ID的第一个收入项目重算主项的实收金额
    '入参:lngMainRow-主项行ID
    '编制:刘兴洪
    '日期:2015-07-10 12:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim cur打折前应收合计 As Currency     '记录所有主从项的应收合计
    Dim cur打折后实收 As Currency
    
    With mobjBill
        For i = lngMainRow To .Details.Count
            If i = lngMainRow Or .Details(i).从属父号 = lngMainRow Then
                For j = 1 To .Details(i).InComes.Count
                    cur打折前应收合计 = cur打折前应收合计 + .Details(i).InComes(j).应收金额
                Next
            End If
        Next
        
        cur打折后实收 = CCur(Format(ActualMoney(.费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计, 0, 0, 0, 0), gstrDec))
        cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
        .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
        
        Call ShowDetails(lngMainRow)
    End With
End Sub
Private Function CheckPatiFeeDateIsValied(tyPati As TY_PATIINFOR) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人的费用日期是否有效
    '入参:tyPati-病人信息
    '返回:有效返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-17 14:45:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String
    On Error GoTo errHandle
    strDate = Format(CDate(txtDate.Text), "yyyy-MM-dd HH:mm:ss")
    '检查发生时间不能早于病人的入院时间
    If strDate < tyPati.入院日期 And tyPati.入院日期 <> "" Then
        MsgBox "费用的发生时间不能小于病人""" & tyPati.姓名 & """的入院时间:" & tyPati.入院日期 & "！", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    '发生时间检查
    If strDate > tyPati.出院日期 And tyPati.出院日期 <> "" Then
        MsgBox "强制对出院病人(" & tyPati.姓名 & ")记帐时，费用时间不能大于病人出院时间:" & tyPati.出院日期, vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    
    If tyPati.险类 <> 0 And strDate < tyPati.入院日期 And tyPati.入院日期 <> "" Then
        MsgBox "费用的发生时间不能小于医保病人的入院时间(" & tyPati.姓名 & "):" & tyPati.入院日期, vbInformation, gstrSysName
        Exit Function
    End If
    CheckPatiFeeDateIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckAllPatiIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：检查发生时间是否合法
    '返回：数据合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 15:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, varData As Variant
    Dim dtDate As Date, strMsg As String
    Dim strYBItemIDs As String, strGFItemIDs As String
    Dim str医疗付款方式 As String, tyPati As TY_PATIINFOR
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    On Error GoTo errH
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            '初始化医保参数
            If tyPati.险类 <> 0 Then Call InitInsurePara(tyPati.病人ID, tyPati.险类)
            '重新给单据附值
            Call reSetBillObject(tyPati, mobjBill)
            
            '病人相关的检查
            bln医保 = False: bln公费 = False
            Call zlIsCheckMedicinePayMode(tyPati.医疗付款方式, bln医保, bln公费)
            If bln医保 Then
                Set mrsMedAudit = GetAuditRecord(tyPati.病人ID, tyPati.主页ID)
            Else
                Set mrsMedAudit = Nothing
            End If
                
            If CheckPatiIsValied(tyPati, mobjBill) = False Then Exit Function
                          
            If InStr(str医疗付款方式 & "','", "','" & tyPati.医疗付款方式 & "','") = 0 And tyPati.医疗付款方式 <> "" Then
                str医疗付款方式 = str医疗付款方式 & "','" & tyPati.医疗付款方式
                '医保或公费病人问题:45605
                If zlIsCheckMedicinePayMode(tyPati.医疗付款方式) Then
                    '处方职务检查
                    i = CheckDuty(, False, tyPati.姓名)
                    If i > 0 Then
                        Bill.Row = i: Bill.MsfObj.TopRow = i
                        Bill.Col = BillCol.项目: Bill.SetFocus
                        Exit Function
                    End If
                End If
                If Check费用类型(tyPati.医疗付款方式, , strYBItemIDs, strGFItemIDs, tyPati.姓名) = False Then Exit Function
            End If
            If Check服务对象(tyPati.病人性质, tyPati.姓名) > 0 Then Exit Function
        End If
    Next
    CheckAllPatiIsValied = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function Check费用类型(ByVal str医疗付款方式 As String, _
    Optional intRow As Integer, _
    Optional strYBItemIDs As String, _
    Optional strGFItemIDs As String, _
    Optional str姓名 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    '入参:intRow-指定行
    '出参: strYBItemIDs-已经检查的医保部分项目,多个用逗号分离
    '      strGFItemIDs-已经检查的公费部分项目,多个用逗号分离
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-10 17:24:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, bytType As Byte
    Dim i As Integer
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    Check费用类型 = True
    
    On Error GoTo errHandle
    

    '无法检查
    If str医疗付款方式 = "" Then Exit Function
    
    '医保或公费病人
    '问题:45605
    '只检查医保病人和公费病人
    If zlIsCheckMedicinePayMode(str医疗付款方式, bln医保, bln公费) = False Then Exit Function
    '确定病人类型
    bytType = IIf(bln医保, 1, 2)
    
    '读取检查数据
    If mrs费用类型 Is Nothing Then
        strSQL = " Select '医保' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Union All " & _
                 " Select '公费' As 类别,编码,名称 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") "
        Set mrs费用类型 = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs费用类型, strSQL, Me.Caption)
    End If
    mrs费用类型.Filter = ""
    If mrs费用类型.RecordCount = 0 Then Exit Function
        
    If bytType = 1 Then
        strSQL = " And 类别='医保'"
    Else
        strSQL = " And 类别='公费'"
    End If
    
    If intRow > 0 Then
        If bytType = 1 Then '医保
            If InStr("," & strYBItemIDs & ",", "," & mobjBill.Details(intRow).收费细目ID & ",") > 0 Then Exit Function
            strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(intRow)
        Else
            If InStr("," & strGFItemIDs & ",", "," & mobjBill.Details(intRow).收费细目ID & ",") > 0 Then Exit Function
            strGFItemIDs = strGFItemIDs & "," & mobjBill.Details(intRow).收费细目ID
        End If
        
        If mobjBill.Details(intRow).Detail.类型 = "" Then
            If InStr("," & strYBItemIDs & "," & strGFItemIDs & ",", "," & mobjBill.Details(intRow).收费细目ID & ",") > 0 Then Exit Function
            MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            Check费用类型 = False
        Else
            mrs费用类型.Filter = "名称='" & mobjBill.Details(intRow).Detail.类型 & "'" & strSQL
            If mrs费用类型.EOF Then
                
                MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型为""" & _
                    mobjBill.Details(intRow).Detail.类型 & """,不是" & _
                    IIf(bytType = 1, "医保", "公费") & "费用类型" & IIf(str姓名 <> "", "(" & str姓名 & ")", "") & "！", vbInformation, gstrSysName
                Check费用类型 = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.类型 = "" Then
                If InStr("," & strYBItemIDs & "," & strGFItemIDs & ",", "," & mobjBill.Details(i).收费细目ID & ",") > 0 Then Exit Function
                strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(i).收费细目ID
                
                If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型未设置！" & vbCrLf & "确实要保存单据吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check费用类型 = False: Exit For
                End If
            Else
                If bytType = 1 Then '医保
                    If InStr("," & strYBItemIDs & ",", "," & mobjBill.Details(i).收费细目ID & ",") > 0 Then Exit Function
                    strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(i).收费细目ID
                Else
                    If InStr("," & strGFItemIDs & ",", "," & mobjBill.Details(i).收费细目ID & ",") > 0 Then Exit Function
                    strGFItemIDs = strGFItemIDs & "," & mobjBill.Details(i).收费细目ID
                End If
                
                mrs费用类型.Filter = "名称='" & mobjBill.Details(i).Detail.类型 & "'" & strSQL
                If mrs费用类型.EOF Then
                    If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型为""" & _
                        mobjBill.Details(i).Detail.类型 & """,不是" & _
                        IIf(bytType = 1, "医保", "公费") & "费用类型" & IIf(str姓名 <> "", "(" & str姓名 & ")", "") & "！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check费用类型 = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Check服务对象(ByVal int病人性质 As Integer, _
    Optional str姓名 As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前病人的记帐费用项目的服务对象是否一致
    '入参:int病人性质-病人性质
    '返回：不一致的费用行,为0时正常
    '编制:刘兴洪
    '日期:2015-07-13 10:29:11
    '说明：因为加入了门诊留观病人,所以有此检查
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
        
    On Error GoTo errHandle
    For i = 1 To mobjBill.Details.Count
        If int病人性质 = 0 Or int病人性质 = 2 Then
            '住院病人或住院留观病人,不能用只服务于门诊的项目
            If mobjBill.Details(i).Detail.服务对象 = 1 Then
                If str姓名 = "" Then
                    MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于门诊,该病人不能使用.", vbInformation, gstrSysName
                Else
                    MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于门诊,病人:" & str姓名 & "不能使用该项目.", vbInformation, gstrSysName
                End If
                Check服务对象 = i: Exit Function
            End If
        ElseIf int病人性质 = 1 Or int病人性质 = -1 Then
            '门诊或出院病人(医技记帐)或门诊留观病人,不能用只服务于住院的项目
            If mobjBill.Details(i).Detail.服务对象 = 2 Then
                 If str姓名 = "" Then
                    MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于住院,该病人不能使用.", vbInformation, gstrSysName
                Else
                    MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于住院,,病人:" & str姓名 & "不能使用该项目.", vbInformation, gstrSysName
                End If
                Check服务对象 = i: Exit Function
            End If
        End If
    Next


    Check服务对象 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckAutoSendDrugAndStuff(ByVal objDetail As BillDetail, _
    ByVal blnSavePrice As Boolean, ByRef blnSendMaterial As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否自动发药
    '入参:objDetail-单据明细
    '     blnSavePrice-是否保存为划价单
    '出参:blnSendMaterial-是否自动发药到病除
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 10:37:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double
    
    On Error GoTo errHandle
    With objDetail
        If .收费类别 = "4" And .Detail.跟踪在用 Then
            dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
            If Not CheckValidity(.收费细目ID, .执行部门ID, dblTotal) Then Exit Function
        End If
        If InStr(1, ",5,6,7,", .收费类别) > 0 Then
            '打印发药单,仅普通记帐,且划价单除外
            If gbytSendMateria <> 0 And mbytUseType = 0 And gbytBilling = 0 And Not blnSavePrice Then
                '全部药品都确定了药房的才自动发药(分离发药时,没有确定药房)
                blnSendMaterial = .执行部门ID <> 0
            End If
        End If
    End With
    CheckAutoSendDrugAndStuff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub BillPrint(ByVal blnSavePrice As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印票据
    '入参:blnSavePrice-是否保存的划价单
    '编制:刘兴洪
    '日期:2015-07-13 10:46:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If gbytBilling = 0 And Not blnSavePrice And gbln记帐打印 Then
         Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
     ElseIf (gbytBilling = 1 Or blnSavePrice) And gbln划价打印 Then
         Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
     End If
    
    '打印发药单
    If mblnPrintDrugList Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), 2)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置领药部门的Enabled属性
    '编制:刘兴洪
    '日期:2015-07-13 10:57:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaveDrug As Boolean '存在药品
    
    '如果没有领用部门的选择,则直接退出
    If cboDrawDept.Visible = False Then cboDrawDept.Enabled = False: lblDrawDrugDept.Enabled = False: Exit Sub
    blnHaveDrug = False
    For i = 1 To mobjBill.Details.Count
        If InStr(1, ",5,6,7,", "," & mobjBill.Details(i).收费类别 & ",") > 0 Then
            blnHaveDrug = True
            Exit For
        End If
    Next
    cboDrawDept.Enabled = blnHaveDrug: lblDrawDrugDept.Enabled = blnHaveDrug
End Sub

Private Sub SetBill中草药EditEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置中草药的编辑状态
    '编制:刘兴洪
    '日期:2015-07-13 11:02:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With Bill
        For i = 0 To .Cols - 1
            .ColData(i) = IIf(.TextMatrix(0, i) = "项目", 0, 5)
        Next
    End With
End Sub
 

Private Sub zlReSetDrawDrugDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据相应的规则,重新获取领药部门
    '编制:刘兴洪
    '日期:2015-07-13 11:39:26
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    '4)  住院记帐、科室分散记帐，可能由病区使用，也可能由医技科室使用。
    '    a)  判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
    '    b)  如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytUseType = 2 Then
        '医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
        mobjBill.领药部门ID = mlngDeptID: Exit Sub
    End If
    
    If mrs领药部门.RecordCount = 0 Then
        '判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
        mobjBill.领药部门ID = mobjBill.病区ID: Exit Sub
    End If
    '如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    If mrs领药部门.RecordCount = 1 Then
        '只有一个部分,肯定是他
        If mrs领药部门.EOF Then mrs领药部门.MoveFirst
         mobjBill.领药部门ID = Val(Nvl(mrs领药部门!ID)): Exit Sub
    End If
    '选择的科室是哪个就是哪个
    With cboDrawDept
        If .ListIndex < 0 Then Exit Sub
        If mobjBill.领药部门ID <> .ItemData(.ListIndex) Then mobjBill.领药部门ID = .ItemData(.ListIndex): Exit Sub
    End With
End Sub

Private Sub zlLoadDrawDeptData(ByVal bytUseType As Byte, Optional ByVal lngDeptID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发药部门
    '入参:bytUseType:记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
    '问题:24729,24731
    '编制:刘兴洪
    '日期:2009-07-29 15:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    '4)  住院记帐、科室分散记帐，可能由病区使用，也可能由医技科室使用。
    '    a)  判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
    '    b)  如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    
    On Error GoTo errHandle
    
    '医技科室
    If bytUseType = 2 Then
        '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
        strSQL = "Select ID,编码,名称 From 部门表 where id=[2]"
    Else
        strSQL = _
            " Select distinct  A.ID, A.编码,A.名称   " & vbNewLine & _
            " From 部门表 A, 部门性质说明 B,部门人员 C" & vbNewLine & _
            " Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)  " & _
            "       And A.ID = B.部门id and a.id=C.部门ID and C.人员id=[1] " & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "       AND B.工作性质 IN('检查','检验','手术','治疗','营养') " & _
            " Order by 编码"
    End If
    Set mrs领药部门 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, lngDeptID)
    With mrs领药部门
        cboDrawDept.Clear
        Do While Not .EOF
            cboDrawDept.AddItem IIf(zlIsShowDeptCode, Nvl(!编码) & "-", "") & Nvl(!名称)
            cboDrawDept.ItemData(cboDrawDept.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = UserInfo.部门ID Then cboDrawDept.ListIndex = cboDrawDept.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 And cboDrawDept.ListIndex < 0 Then cboDrawDept.ListIndex = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置领药部门的visibled属性
    '编制:刘兴洪
    '日期:2009-07-29 19:07:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ' mbytUseType As Byte '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    If mblnNurseStation Then Exit Sub
    If mbytUseType = 2 Then
        cboDrawDept.Visible = False
    Else
        cboDrawDept.Visible = mrs领药部门.RecordCount > 1 And gbytBilling <> 2         '
    End If
    lblDrawDrugDept.Visible = cboDrawDept.Visible
End Sub


Private Function GetLastDeptID(ByVal str类别 As String, ByVal lngRow As Long, _
    ByVal strDeptIDs As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最近输入的相同类别项目的执行科室ID
    '入参:str类别-收费类别
    '     lngRow-指定行
    '     strDeptIDs-执行部门ID,多个用逗号分离
    '返回:成功,返回最后一个执行部门ID,否则返回0
    '编制:刘兴洪
    '日期:2015-07-13 11:52:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).收费类别 = str类别 _
            And mobjBill.Details(i).执行部门ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).执行部门ID
                Exit Function
            End If
        End If
    Next
    
    '如果是卫生材料,再取与最近其它类别相匹配的执行科室
    If str类别 = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).执行部门ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).执行部门ID
                    Exit Function
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillBillComboBox(lngRow As Long, lngCol As Long, Optional blnEnter As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据列设置下拉列表框内容
    '入参:blnEnter=是否按进入该列处理,比如执行科室保持不变
    '编制:刘兴洪
    '日期:2015-07-13 11:53:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, bln护士 As Boolean
    Dim strSQL As String, strIDs As String, i As Long
    Dim lng病区ID As Long, lng科室ID As Long, j As Long
    Dim bln草药类别 As Boolean '是否允许输入草药类别
    
    Bill.Clear
    
    On Error GoTo errHandle
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "类别"
            Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                j = 1
                For i = 1 To mrsClass.RecordCount
                    '护士类别:限制
                    If Not (bln护士 And InStr(",E,M,4,", mrsClass!编码) = 0) Then
                        Bill.AddItem j & "-" & mrsClass!类别
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!编码)  '存放类别编码的ASCII码
                        j = j + 1
                    End If
                    mrsClass.MoveNext
                Next
            End If
            Bill.cboStyle = DropDownAndEdit  ' DropOlnyDown
        Case "执行科室"
            Bill.cboStyle = DropDownAndEdit
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .收费类别) > 0 Then
                        Call GetWorkUnit(.收费细目ID, .收费类别)
                        If mrsWork.RecordCount > 0 Then
                            '取上一个药的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '进入该列时保持已确定值不变
                                lng科室ID = GetLastDeptID(.收费类别, lngRow, Mid(strIDs, 2))
                            End If
                            If lng科室ID = 0 Then lng科室ID = .执行部门ID
                            
                            '确定当前行的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!编码 & "-", "") & mrsWork!名称
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        lng科室ID = mobjBill.科室ID
                        If lng科室ID = 0 Then lng科室ID = Get开单科室ID
                        
                        lng病区ID = mobjBill.病区ID
                        If lng病区ID = 0 Then lng病区ID = Get病区ID(lng科室ID)
                        If lng病区ID = 0 Then lng病区ID = lng科室ID
                        
                        '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                        Select Case .Detail.执行科室
                            Case 0 '不明确
                                mrsUnit.Filter = 0
                            Case 1 '病人科室
                                mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & .执行部门ID
                            Case 2 '病人病区
                                mrsUnit.Filter = "ID=" & lng病区ID & " Or ID=" & .执行部门ID
                            Case 3 '操作员科室
                                mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                            Case 4 '指定科室
                                strSQL = "" & _
                                "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                                "   From 收费执行科室 A,部门表 C" & _
                                "   Where A.收费细目ID=[1]　And A.执行科室ID+0=C.ID " & _
                                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                                " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .收费细目ID, 2, lng科室ID)
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!执行科室ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .执行部门ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                                End If
                            Case 5 '院外执行(预留,程序暂未用)
                            Case 6 '开单人科室
                               mrsUnit.Filter = "ID=" & Get开单科室ID & " Or ID=" & .执行部门ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                                '刘兴洪:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                    
                                    '设置缺省执行科室
                                    If Not blnEnter Then '进入该列时保持已确定值不变
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '与上一行非药品相同
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).执行部门ID And mobjBill.Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).收费类别) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                            
                            If Not blnEnter And .Detail.执行科室 = 4 Then    '执行科室为指定科室的,缺省为操作员所在科室
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = UserInfo.部门ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                            
                            If Bill.ListIndex = -1 Then '如果没有则取现有的执行科室
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = .执行部门ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                            
                            If mblnNurseStation Then    '护士站缺省按第一个病人的执行科室来缺省.
                                Dim tyPati As TY_PATIINFOR
                                For i = 0 To rptPati.Rows.Count - 1
                                    If rptPati.Rows(i).Record.Tag = "1" Then
                                        tyPati = GetPatiInforByReport(i)
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = tyPati.开单科室ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .执行部门ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .执行部门ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:界面设置为不可修改状态
    '编制:刘兴洪
    '日期:2015-07-13 11:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
        
    cboNO.Locked = Not bln
    cbo开单科室.Locked = Not bln
    cbo开单人.Locked = Not bln
    chk加班.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function CalcMoneys(tyPati As TY_PATIINFOR, _
    Optional lngRow As Long = 0, Optional ByVal blnNoPrompt As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算或重新计算指定行或所有行的金额
    '入参:tyPati-当前病人信息
    '     lngRow=指定行,为0表示计算所有行
    '返回:计算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 12:00:02
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strMainRows As String
    Dim bln从项汇总折扣 As Boolean
    
    On Error GoTo errHandle
    
    If mobjBill.Details.Count = 0 Then CalcMoneys = True: Exit Function
    If lngRow > mobjBill.Details.Count Then CalcMoneys = True: Exit Function
    
    Call reSetBillObject(tyPati, mobjBill)   '重新设置对象
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        bln从项汇总折扣 = False
        If gbln从项汇总折扣 Then                    '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
            If mobjBill.Details(i).从属父号 > 0 Then    '从项
                bln从项汇总折扣 = Not mobjBill.Details(mobjBill.Details(i).从属父号).Detail.屏蔽费别
                If bln从项汇总折扣 And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).从属父号      '单独计算一行的时候
            Else
                If CheckItemHaveSub(i) Then                            '主项或独立项
                     bln从项汇总折扣 = Not mobjBill.Details(i).Detail.屏蔽费别
                     If bln从项汇总折扣 Then strMainRows = strMainRows & "," & i  '一页可能有多个主从项,先记录主项行号,后面再重算主项折扣
                End If
            End If
        End If
        Call CalcMoney(tyPati, i, bln从项汇总折扣, blnNoPrompt)
    Next
    
    '重算所有主项,不能用bln从项汇总折扣变量,因为可能在遇到不是从项的行时已改变
    If gbln从项汇总折扣 Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc重算主项实收(Split(strMainRows, ",")(i))
        Next
    End If
    CalcMoneys = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CalcMoney(ByRef tyPati As TY_PATIINFOR, _
    lngRow As Long, Optional bln从项汇总折扣 As Boolean, Optional ByVal blnNoPrompt As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算或重新计算指定行的金额
    '入参:tyPati-当前计算的病人信息
    '     lngRow=指定行
    '返回:计算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 13:58:14
    '说明：1.ExpenseBill集合的索引对应单据的行号
    '      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
    '      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
    '      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, i As Long
    Dim dblMoney As Double '用户输入的变价金额

    Dim dblAllTime As Double, dbl加班加价率 As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        strPriceGrade = mstr药品价格等级
    ElseIf mobjBill.Details(lngRow).收费类别 = "4" Then
        strPriceGrade = mstr卫材价格等级
    Else
        strPriceGrade = mstr普通价格等级
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).收费细目ID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.价格等级 = [2]" & vbNewLine & _
            "            Or (b.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From 收费价目" & vbNewLine & _
            "                               Where b.收费细目Id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                     And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    strSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID, strPriceGrade)
    If rsTmp.EOF Then
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        CalcMoney = True
        Exit Function
    End If
    
    '先获取操作员以前输入的变价金额
    With mobjBill.Details(lngRow)
        If InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用) Then
            '计算药品时价(分批或不分批)
            '必然有记录(输入该项目时已判断)
            dblAllTime = .付数 * .数次
            If gbln住院单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                dblAllTime = dblAllTime * .Detail.住院包装 '库存时价按售价数量进行计算
            End If
            If dblAllTime <> 0 Or Not .Detail.变价 Then
                Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                            Me.Caption, .收费细目ID, .执行部门ID, dblAllTime)
                If rsPrice.EOF Then
                    '获取价格失败
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        MsgBox "第 " & lngRow & " 行药品""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                    Else
                        MsgBox "第 " & lngRow & " 行卫生材料""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                    End If
                Else
                    strPrice = Nvl(rsPrice!Price) & "|||"
                    varPrice = Split(strPrice, "|")
                    dblMoney = Val(varPrice(0))
                    dbl剩余数量 = Val(varPrice(2))
                    
                    If dbl剩余数量 <> 0 And .Detail.变价 Then
                        '数量未分解完毕
                        If Not blnNoPrompt Then
                            If InStr(",5,6,7,", .收费类别) > 0 Then
                                MsgBox "第 " & lngRow & " 行时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            End If
                        End If
                        dblMoney = 0
                    End If
                End If
            Else
                dblMoney = 0
            End If
        Else
            If .Detail.变价 Then
                If .InComes.Count = 0 Then '第一次计算金额取缺省值
                    dblMoney = Val(Nvl(rsTmp!缺省价格))
                Else                        '获取操作员以前输入的变价金额
                    dblMoney = .InComes(1).标准单价
                    '如果用户输入的变价不满足变价范围，则取缺省值
                    If CheckScope(Val(Nvl(rsTmp!原价)), Val(Nvl(rsTmp!现价)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!缺省价格))
                    End If
                End If
            End If
        End If
    End With
    
    '再清除原有记录
    Set mobjBill.Details(lngRow).InComes = New BillInComes
    
    '填写现有费用记录
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .收入项目ID = rsTmp!收入项目ID
            .收入项目 = rsTmp!名称
            .收据费目 = Nvl(rsTmp!收据费目)
            .原价 = Val(Nvl(rsTmp!原价))
            .现价 = Val(Nvl(rsTmp!现价))
            
            If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
                If gbln住院单位 Then
                    .标准单价 = Format(dblMoney * mobjBill.Details(lngRow).Detail.住院包装, gstrFeePrecisionFmt)
                Else
                    .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                End If
            Else
                If mobjBill.Details(lngRow).Detail.变价 Then
                    .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .标准单价 = Format(Nvl(rsTmp!现价, 0), gstrFeePrecisionFmt)
                End If
            End If
            
            '应收金额=单价 * 付数 * 数次
            .应收金额 = .标准单价 * IIf(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数) * mobjBill.Details(lngRow).数次
            
            '附加手术费率用计算(所有收入项目)
            If mobjBill.Details(lngRow).附加标志 = 1 And mobjBill.Details(lngRow).收费类别 = "F" Then
                .应收金额 = .应收金额 * IIf(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
            End If
            
            '加班费用率计算
            dbl加班加价率 = 0
            If mobjBill.加班标志 = 1 And mobjBill.Details(lngRow).Detail.加班加价 Then
                dbl加班加价率 = IIf(IsNull(rsTmp!加班加价率), 0, rsTmp!加班加价率 / 100)
                .应收金额 = .应收金额 + .应收金额 * dbl加班加价率
            End If
            
            .应收金额 = CCur(Format(.应收金额, gstrDec))
            dblAllTime = mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
            If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
                If gbln住院单位 Then dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.住院包装
            End If
            
            If mobjBill.Details(lngRow).Detail.屏蔽费别 _
                Or bln从项汇总折扣 Or .应收金额 = 0 Or tyPati.病人ID = 0 Then
                .实收金额 = .应收金额
            Else
                If .应收金额 = 0 Then
                    .实收金额 = 0
                    mobjBill.Details(lngRow).费别 = mobjBill.费别
                Else
                     '药品按成本价加收,传入数量
                    .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额, _
                         mobjBill.Details(lngRow).收费细目ID, mobjBill.Details(lngRow).执行部门ID, dblAllTime, dbl加班加价率), gstrDec))
                End If
            End If
            
            '获取项目保险信息,医保病人才处理,不需要连接医保
            If tyPati.病人ID <> 0 And tyPati.险类 <> 0 Then
                strInfo = gclsInsure.GetItemInsure(tyPati.病人ID, _
                    mobjBill.Details(lngRow).收费细目ID, .实收金额, False, tyPati.险类, _
                     mobjBill.Details(lngRow).摘要 & "||" & dblAllTime)
                     
                If strInfo <> "" Then
                    mobjBill.Details(lngRow).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(lngRow).保险大类ID = Val(Split(strInfo, ";")(1))
                    .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                    mobjBill.Details(lngRow).保险编码 = CStr(Split(strInfo, ";")(3))
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).摘要 = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.类型 = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            End If
            '实收金额存入Key中,以处理分币问题(即Key中存放原始实收金额,不变)
            mobjBill.Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
        End With
        rsTmp.MoveNext
    Next
    CalcMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub ShowDetails(Optional lngRow As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新显示指定行或所有行的内容
    '入参:lngRow=指定行,为0表示显示所有行
    '编制:刘兴洪
    '日期:2015-07-13 14:11:09
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, curTotal As Currency
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
End Sub
Private Sub ShowDetail(lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新显示指定行的内容
    '入参:lngRow=指定行
    '编制:刘兴洪
    '日期:2015-07-13 14:12:47
    '说明：ExpenseBill集合的索引对应单据的行号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    If lngRow > mobjBill.Details.Count Then Exit Sub
    
    '清除单据行
    For i = 1 To Bill.Cols - 1
        '输入时收费类别不清除
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).收费类别 <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).收费类别)
    End If
    
    '刷新单据行
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "类别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类别名称
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.规格
            Case "商品名"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.商品名
            Case "单位"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And gbln住院单位 Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.住院单位
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.计算单位
                End If
            Case "付数"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数)
            Case "数次"
                '数次在第一次显示时已默认设置为1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).数次, 5)
            Case "单价"
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                dbl单价 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl单价 = dbl单价 + mobjBill.Details(lngRow).InComes(j).标准单价
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, gstrFeePrecisionFmt)
            Case "应收金额"
                '应收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "执行科室"
                If mobjBill.Details(lngRow).执行部门ID <> 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).执行部门ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = mrsUnit!编码 & "-" & mrsUnit!名称
                    Else
                        Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "标志"
                If mobjBill.Details(lngRow).收费类别 = "F" And mobjBill.Details(lngRow).附加标志 = 1 Then
                    Bill.TextMatrix(lngRow, i) = "√"
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新显示收入项目费用区
    '编制:刘兴洪
    '日期:2015-07-13 14:14:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean, curTotal As Currency, cur应收Total As Currency
    
    vsMoney.Redraw = False
    
    '产生汇总费目
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '查找是否已经加入此类收入项目,如是则合计,否则新入
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).收入项目ID = mobjBill.Details(i).InComes(j).收入项目ID Then
                    blnExist = True: Exit For
                End If
            Next
            If blnExist Then
                mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + mobjBill.Details(i).InComes(j).实收金额
                mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + mobjBill.Details(i).InComes(j).应收金额
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额
                End With
            End If
        Next
    Next
    
    '刷新显示
    If mcolMoneys.Count > 0 Then
        vsMoney.Rows = mcolMoneys.Count + 1
    End If
    If vsMoney.Rows < 5 Then vsMoney.Rows = 5

    Call SetMoneyList
    
    '刷新显示
    If mcolMoneys.Count > 0 Then
        vsMoney.Rows = mcolMoneys.Count + 1
    End If
    If vsMoney.Rows < 5 Then vsMoney.Rows = 5
    
    
    For i = 1 To mcolMoneys.Count
        vsMoney.TextMatrix(i, 0) = mcolMoneys(i).收入项目
        vsMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).实收金额, gstrDec)
        curTotal = curTotal + mcolMoneys(i).实收金额
        cur应收Total = cur应收Total + mcolMoneys(i).应收金额
    Next
    txt应收.Text = Format(cur应收Total, gstrDec)
    'txt实收.Text = Format(curTotal, gstrDec)
    For i = 1 To vsMoney.Rows - 1
        vsMoney.TopRow = i
    Next
    vsMoney.Redraw = True
End Sub

Private Function GetCur应收() As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人当前单据合计金额(收费病人累加单据时用)
    '编制:刘兴洪
    '日期:2015-07-13 14:15:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCur应收 = GetCur应收 + mcolMoneys(i).应收金额
    Next
End Function

Private Function GetInputDetail(ByVal lng项目id As Long) As Detail
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取收费项目信息
    '编制:刘兴洪
    '日期:2015-07-13 14:15:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mstrInsures <> "" Then
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,M.要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C.住院包装) as 住院包装," & _
        "       Decode(A.类别,'4',A.计算单位,C.住院单位) as 住院单位,D.跟踪在用,A.录入限量,C.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,C.剂量系数" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,诊疗项目目录 M1," & _
        "       (   Select A1.收费细目ID,max(A1.要求审批) as 要求审批  " & _
        "           From 保险支付项目 A1,Table(f_Num2List([2])) B1 " & _
        "           Where A1.收费细目ID=[1] and a1.险类=b1.Column_value " & _
        "           Group by A1.收费细目ID) M" & _
        " Where A.类别=B.编码 And A.ID=C.药品ID(+) And C.药名ID=M1.id(+) And A.ID=D.材料ID(+)" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.ID=M.收费细目ID(+)  " & vbNewLine & _
        "       And A.ID=[1]"
    Else
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,0 as 要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C.住院包装) as 住院包装," & _
        "       Decode(A.类别,'4',A.计算单位,C.住院单位) as 住院单位,D.跟踪在用,A.录入限量,C.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,C.剂量系数" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,诊疗项目目录 M1" & _
        " Where A.类别=B.编码 And A.ID=C.药品ID(+) And C.药名ID=M1.id(+) And A.ID=D.材料ID(+)" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, mstrInsures)
    With objDetail
        .ID = rsTmp!ID
        .药名ID = Nvl(rsTmp!药名ID, 0) '用于判断输入重复
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .规格 = Nvl(rsTmp!规格)
        .计算单位 = Nvl(rsTmp!计算单位)
        .住院单位 = Nvl(rsTmp!住院单位)
        .住院包装 = Nvl(rsTmp!住院包装, 1)
        .分批 = Nvl(rsTmp!分批, 0) = 1 '是否药房分批
        .变价 = Nvl(rsTmp!是否变价, 0) = 1 '对药品表明是否时价
        .类型 = Nvl(rsTmp!费用类型)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .服务对象 = Nvl(rsTmp!服务对象, 0)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
        .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
        .中药形态 = Val(Nvl(rsTmp!中药形态))
        .商品名 = Nvl(rsTmp!商品名)
        .诊疗名称 = Nvl(rsTmp!诊疗名称)
        .剂量单位 = Nvl(rsTmp!剂量单位)
        .剂量系数 = Val(Nvl(rsTmp!剂量系数))
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, _
    Optional bytParent As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的收费细目对象设定单据指点定行的收费细目：(新增的或修改)
    '编制:刘兴洪
    '日期:2015-07-13 14:18:31
    '说明：
    '      1.用于新输入或更改收费细目行！！！
    '      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    '取其它中药的付数
    intPay = GetPay(lngRow)
    If Detail.类别 <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '如果该行对应的程序对象尚未初始,则加入
        With Detail
            '序号=行号,父号=0
            '次数=1,从属项目的次数由主项计算确定
            '执行部门ID:根据细目执行科室标志取
            '附加标志:以第一行为假,其它为真优先权
            '收入集=空
            If bytParent <> 0 Then
                '设置该行RowData
                Bill.RowData(lngRow) = Asc(Detail.类别)
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = IIf(Detail.从项数次 = 0, 1, Detail.从项数次)
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Details(bytParent).数次
                End If
            Else
                
                If InStr(",5,6,7,", Detail.类别) > 0 Then
                    dblTime = 0
                                     
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add Detail, .ID, CByte(lngRow), CInt(bytParent), 0, 0, 0, 0, "", "", "", _
            0, 0, mobjBill.费别, 0, .类别, .计算单位, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else '如果该行已经存在,则修改
        
        If InStr(",5,6,7,", Detail.类别) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .费别 = mobjBill.费别
            .付数 = intPay
            .附加标志 = 0
            .计算单位 = Detail.计算单位
            .收费类别 = Detail.类别
            .收费细目ID = Detail.ID
            .数次 = dblTime
            .序号 = lngRow
            .从属父号 = 0
            .执行部门ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断该行是否应该取从属项目
    '返回:有从属项目返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 14:19:48
    '说明：仅该行收费项目有从属项目及尚未取才取。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(从项ID) as NUM From 收费从属项目 Where 主项ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).从属父号 = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng项目id As Long) As Details
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回一个收费细目的从属项目集
    '入参:lng项目id-收费细目ID
    '返回:返回Details对象
    '编制:刘兴洪
    '日期:2015-07-13 14:20:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
    Dim dblStock As Double
    
    Set GetSubDetails = New Details
    If mstrInsures <> "" Then
        strSQL = _
        " Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别,G.要求审批," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D.住院包装) as 住院包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D.住院单位) as 住院单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,D.剂量系数" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1, 诊疗项目目录 M1," & _
        "       (   Select A1.收费细目ID,max(A1.要求审批) as 要求审批  " & _
        "           From 保险支付项目 A1,Table(f_Num2List([2])) B1 " & _
        "           Where A1.收费细目ID=[1] and a1.险类=b1.Column_value " & _
        "           Group by A1.收费细目ID) G" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And D.药名ID=M1.id(+) And A.ID=E.材料ID(+)" & _
        "       And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "       And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And C.主项ID=[1] And A.ID=G.收费细目ID(+)   " & _
        " Order by 编码"
    Else
        strSQL = _
        "Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别,0 as 要求审批," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D.住院包装) as 住院包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D.住院单位) as 住院单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,D.剂量系数" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1,诊疗项目目录 M1" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And D.药名ID=M1.id(+)  And A.ID=E.材料ID(+)" & _
        "   And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "   And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "   And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "   And C.主项ID=[1] " & _
        " Order by 编码"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        If mblnNurseStation And InStr(",5,6,7,", rsTmp!类别) > 0 Then
            rsTmp.MoveNext
        Else
            Set objDetail = New Detail
            With objDetail
                .ID = rsTmp!ID
                .药名ID = Nvl(rsTmp!药名ID, 0)
                .编码 = rsTmp!编码
                .变价 = Nvl(rsTmp!是否变价, 0) = 1
                .规格 = Nvl(rsTmp!规格)
                .住院包装 = Nvl(rsTmp!住院包装, 1)
                .住院单位 = Nvl(rsTmp!住院单位)
                .计算单位 = Nvl(rsTmp!计算单位)
                .分批 = Nvl(rsTmp!分批, 0) = 1
                .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
                .类别 = rsTmp!类别
                .类别名称 = rsTmp!类别名称
                .名称 = rsTmp!名称
                .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
                .执行科室 = Nvl(rsTmp!执行科室, 0)
                .服务对象 = Nvl(rsTmp!服务对象, 0)
                .固有从属 = Nvl(rsTmp!固有从属, 0)
                .从项数次 = Nvl(rsTmp!从项数次, 1)
                .类型 = Nvl(rsTmp!费用类型)
                .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
                .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
                .中药形态 = Val(Nvl(rsTmp!中药形态))
                .商品名 = Nvl(rsTmp!商品名)
                .诊疗名称 = Nvl(rsTmp!诊疗名称)
                .剂量单位 = Nvl(rsTmp!剂量单位)
                .剂量系数 = Val(Nvl(rsTmp!剂量系数))
                GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, _
                    .住院包装, .住院单位, .分批, .变价, .加班加价, .执行科室, .服务对象, .类型, .补充摘要, .固有从属, .从项数次, .跟踪在用, , , , , , .要求审批, , .中药形态, .商品名, .诊疗名称, .剂量单位, .剂量系数
            End With
            rsTmp.MoveNext
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除指定收费项目行
    '编制:刘兴洪
    '日期:2015-07-13 14:22:03
    '说明：这时不处理从属行的删除,但要对其它单据行从属关系作相应的调整
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).从属父号 > lngRow Then
            mobjBill.Details(i).从属父号 = mobjBill.Details(i).从属父号 - 1
        End If
        mobjBill.Details(i).序号 = mobjBill.Details(i).序号 - 1 '序号与行号对应
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化一张新的单据(程序对象)
    '编制:刘兴洪
    '日期:2015-07-13 14:22:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnKeepDate As Boolean
    Dim dtdtCurDate As Date     '服务器当前时间
    
    mlngPreRow = 0
        
    sta.Panels(3).Text = ""
    Set mrsMedAudit = Nothing
    mstrWarn = ""
    cboNO.Text = ""
    Set mobjBill = New ExpenseBill
    Bill.ColData(BillCol.类别) = IIf(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
    dtdtCurDate = zlDatabase.Currentdate
    chk加班.Value = IIf(OverTime(dtdtCurDate), 1, 0)
    txtDate.Text = Format(dtdtCurDate, "yyyy-MM-dd HH:mm:ss")
    
    Call cbo开单科室_Click
    
    cmdOK.Visible = True
    
    
    With mobjBill
        .门诊标志 = 2
        .划价人 = UserInfo.姓名
        .开单人 = zlStr.NeedName(cbo开单人.Text)
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
        .发生时间 = CDate(txtDate.Text)
        .加班标志 = chk加班.Value
        .婴儿费 = 0
        
        If cbo开单科室.ListIndex = -1 Then
            .开单部门ID = 0
        Else
            .开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        If cboDrawDept.ListIndex = -1 Then
            .领药部门ID = 0
        Else
            .领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        End If
    End With
End Sub
Private Sub ClearMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除费用显示区
    '编制:刘兴洪
    '日期:2015-07-13 14:25:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    On Error GoTo errHandle
    
    vsMoney.Redraw = flexRDNone
    For i = 1 To vsMoney.Rows - 1
        For j = 0 To vsMoney.Cols - 1
            vsMoney.TextMatrix(i, j) = ""
        Next
    Next
    vsMoney.Rows = 5
    vsMoney.Redraw = flexRDBuffered

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitInsurePara(ByVal lng病人ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2015-07-13 15:36:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    MCPAR.记帐上传 = gclsInsure.GetCapability(support记帐上传, , intInsure)
    MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , intInsure)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chkIn_Click()
    sta.Panels(2) = ""
    If chkIn.Value = Checked Then
        txtIn.Enabled = True
        txtIn.BackColor = &H80000005
        sta.Panels(2) = "请输入要导入的记帐单单据号码"
        txtIn.SetFocus
    Else
        txtIn.Text = ""
        txtIn.Enabled = False
        txtIn.BackColor = &HE0E0E0
        Bill.SetFocus
    End If
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim tmpBill As New ExpenseBill
    Dim tyPati As TY_PATIINFOR
    Dim lng病人ID As Long, i As Long
    Dim lngPre As Long, strPre As String
    Dim dtCurdate As Date     '服务器当前时间
 
    
    On Error GoTo errH
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
        Exit Sub
    End If
    
    txtIn.Text = GetFullNO(txtIn.Text, 14)
   
    Set tmpBill = ImportBill(txtIn.Text, False, Me, False, gbln住院单位, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    If tmpBill.NO = "" Then
        MsgBox "读取单据失败。", vbExclamation, gstrSysName
        txtIn.Text = "": txtIn.SetFocus: Exit Sub
    End If

    '单据修改及显示
    Screen.MousePointer = 11
                
    lng病人ID = tmpBill.病人ID
    lngPre = tmpBill.开单部门ID
    strPre = tmpBill.开单人
    If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then strPre = ""
    
    '清除导入的病人信息
    tmpBill.病人ID = 0
    tmpBill.主页ID = 0
    tmpBill.床号 = ""
    tmpBill.标识号 = 0
    tmpBill.姓名 = ""
    tmpBill.性别 = ""
    tmpBill.年龄 = ""
    tmpBill.费别 = ""
    tmpBill.病区ID = 0
    tmpBill.科室ID = 0
    
    '刘兴洪:25882
    For i = 1 To tmpBill.Details.Count
        tmpBill.Details(i).病人ID = 0
        tmpBill.Details(i).主页ID = 0
        tmpBill.Details(i).姓名 = ""
        tmpBill.Details(i).性别 = ""
        tmpBill.Details(i).年龄 = ""
        tmpBill.Details(i).费别 = ""
        tmpBill.Details(i).病区ID = 0
        tmpBill.Details(i).科室ID = 0
    Next
    
    '保留现有病人信息
    If Not mobjBill Is Nothing Then
        If mobjBill.病人ID > 0 Then
            lng病人ID = mobjBill.病人ID
            lngPre = mobjBill.开单部门ID
            strPre = mobjBill.开单人
        End If
    End If
    
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    
    dtCurdate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.登记时间 = dtCurdate
    mobjBill.操作员编号 = UserInfo.编号
    mobjBill.操作员姓名 = UserInfo.姓名
    mobjBill.加班标志 = chk加班.Value
    mobjBill.婴儿费 = 0
    
    '取当前时间
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
 
    Bill.Redraw = False
    Bill.ClearBill
    Bill.Rows = mobjBill.Details.Count + 1
    
    Call InitBillColumnColor
    
    '记帐分类报警
    mstrWarn = ""
    
    mobjBill.开单部门ID = lngPre
    mobjBill.开单人 = strPre
    Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
    '等上面的读病人后确定费别后,再计算价格
    Call CalcMoneys(tyPati)
    Call ShowDetails
    Call ShowMoney
    
    Bill.Redraw = True
    chkIn.Value = 0
    Call SetDrawDrugDeptEnabled
    Call SetColNum
    
    Screen.MousePointer = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReCalcInsure(tyPati As TY_PATIINFOR)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改单据时,重新计算统筹金额及更新相关信息
    '入参:tyPati-病人信息
    '编制:刘兴洪
    '日期:2015-07-13 15:44:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String, varTemp As Variant
    If tyPati.病人ID = 0 Or tyPati.险类 = 0 Then Exit Sub
    Err = 0: On Error GoTo ErrHand:
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            For j = 1 To .InComes.Count
                dblAllTime = .付数 * .数次
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If gbln住院单位 Then dblAllTime = dblAllTime * .Detail.住院包装
                End If
                
                strInfo = gclsInsure.GetItemInsure(tyPati.病人ID, .收费细目ID, .InComes(j).实收金额, False, tyPati.险类, _
                     .摘要 & "||" & dblAllTime)
                If strInfo <> "" Then
                    varTemp = Split(strInfo & ";;;;", ";")
                    
                    .保险项目否 = Val(varTemp(0)) <> 0
                    .保险大类ID = Val(varTemp(1))
                    .InComes(j).统筹金额 = Val(varTemp(2))
                    .保险编码 = CStr(varTemp(3))
                    
                    If UBound(varTemp) >= 4 Then
                        If CStr(varTemp(4)) <> "" Then .摘要 = CStr(varTemp(4))
                        If UBound(varTemp) >= 5 Then
                            If varTemp(5) <> "" Then .Detail.类型 = varTemp(5)
                        End If
                    End If
                End If
            Next
        End With
    Next
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Function GetAdviceIDs(ByVal lng医嘱ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取一组医嘱包含的医嘱记录ID串
    '入参:lng医嘱ID=一组医嘱记录的组ID:Nvl(相关ID,ID)
    '返回: 返回一组医嘱ID,用逗号分隔
    '编制:刘兴洪
    '日期:2015-07-13 15:52:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    GetAdviceIDs = Mid(strSQL, 2)
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClearRows()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:表除行内部标志
    '编制:刘兴洪
    '日期:2015-07-13 15:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取其它中药的付数
    '编制:刘兴洪
    '日期:2015-07-13 16:00:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).付数
            Exit For
        End If
    Next
End Function

Private Sub InitBillColumnColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化单据列颜色
    '编制:刘兴洪
    '日期:2015-07-13 15:59:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Bill.SetColColor BillCol.类别, &HE7CFBA
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.执行科室, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    Bill.SetColColor BillCol.标志, &HE0E0E0
End Sub

Private Function GetDetailNum(tyPati As TY_PATIINFOR, _
    ByVal lngRow As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人指定细目的总记帐数据(含本单据中)
    '入参:lngRow=当前单据行
    '返回:返回总数量
    '编制:刘兴洪
    '日期:2015-07-13 16:00:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long, i As Long, lng收费细目ID As Long
    Dim strSQL As String
    If tyPati.病人ID = 0 Then Exit Function
    If lngRow > mobjBill.Details.Count Then Exit Function
    
    lng收费细目ID = mobjBill.Details(lngRow).收费细目ID
    
    '当前单据中的数量
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If i <> lngRow And .收费细目ID = lng收费细目ID Then
                lngNum = lngNum + .数次 * IIf(.付数 = 0, 1, .付数)
            End If
        End With
    Next
    
    '数据库中的数量
    strSQL = _
    " Select Sum(A.数次*Nvl(A.付数,1)" & IIf(gbln住院单位, "/Nvl(B.住院包装,1)", "") & ") as Num" & _
    " From 住院费用记录 A,药品规格 B" & _
    " Where A.价格父号 is Null And A.记帐费用=1" & _
            IIf(gbytBilling = 0, " And A.记录状态<>0", "") & _
    " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2] And A.收费细目ID=B.药品ID(+) And A.收费细目ID+0=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, tyPati.病人ID, tyPati.主页ID, lng收费细目ID)
    If Not rsTmp.EOF Then
        lngNum = lngNum + Nvl(rsTmp!Num, 0)
    End If
    GetDetailNum = lngNum
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lng药品ID As Long, ByVal str类别 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取所有可供选择的药房
    '入参:lng药品ID-药品ID
    '     str类别-类别
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 16:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, str药房 As String, bytDay As Byte
    Dim int服务对象 As Integer, str服务对象 As String
    Dim int病人来源 As Integer, lng开单科室ID As Long
    
    '根据项目及权限确定药房的服务对象
    int服务对象 = Get服务对象(lng药品ID)
    
    If int服务对象 = 1 Then
        str服务对象 = "1,3"
    ElseIf int服务对象 = 2 Then
        str服务对象 = "2,3"
    ElseIf int服务对象 = 3 Then
        If InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
            str服务对象 = "1,2,3"
        Else
            str服务对象 = "2,3"
        End If
    Else
            str服务对象 = "2,3"
    End If
    
    '确定病人来源
    int病人来源 = 2
    
    lng开单科室ID = mobjBill.科室ID
    If cbo开单科室.Visible Then
        If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        If lng开单科室ID = 0 Then lng开单科室ID = mlngDeptID
        If lng开单科室ID = 0 Then lng开单科室ID = GetNurseStationFirstPatiDeptID '护士工作站,取第一个病人科室ID
        If lng开单科室ID = 0 Then lng开单科室ID = mlng病区ID
    End If
       
    If str类别 = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.编码, c.简码, c.名称, b.工作性质, b.服务对象" & vbNewLine & _
        "From 收费执行科室 A, 部门性质说明 B, 部门表 C" & vbNewLine & _
        "Where a.执行科室id + 0 = b.部门id And b.工作性质 = '发料部门' And b.服务对象 IN(" & str服务对象 & ") And b.部门id = c.Id And" & vbNewLine & _
        "      (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null) And" & vbNewLine & _
        "      (a.病人来源 Is Null Or a.病人来源 = [1]) And" & vbNewLine & _
        "      (a.开单科室id Is Null Or a.开单科室id = [2] Or Exists (Select 1 From 病区科室对应 Where 科室id = [2] And a.开单科室id = 病区id)) And a.收费细目id = [3]" & vbNewLine & _
        "Order By b.服务对象, c.编码"
    Else
        '由药品材质确定药房性质
        Select Case str类别
            Case "5"
                str药房 = "西药房"
            Case "6"
                str药房 = "成药房"
            Case "7"
                str药房 = "中药房"
        End Select
        
        '药品从系统指定的储备药房中找
        If Not gbln药房上班安排 Then
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
            "       And B.服务对象 IN(" & str服务对象 & ") And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            "       And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
            "       And B.服务对象 IN(" & str服务对象 & ") And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And D.部门ID=C.ID And D.星期=[5]" & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            "       And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        End If
    End If
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int病人来源, lng开单科室ID, lng药品ID, str药房, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CalcGridToTal(Optional bln应收 As Boolean) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算应收或收实总金额
    '入参:bln应收-是否取应收金额True-应收,False-实收
    '返回:返回总金额
    '编制:刘兴洪
    '日期:2015-07-13 16:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    On Error GoTo errHandle
    
    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If bln应收 Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.应收金额
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.实收金额
                End If
            Next
        Next
        Exit Function
    End If

    For i = 1 To Bill.Cols - 1
        If bln应收 Then
            If Bill.TextMatrix(0, i) = "应收金额" Then intCol = i: Exit For
        Else
            If Bill.TextMatrix(0, i) = "实收金额" Then intCol = i: Exit For
        End If
    Next

    For i = 1 To Bill.Rows - 1
        CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
    Next
    
    

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub SetColNum(Optional intRow As Long = 1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新显示各行的行号
    '入参:intRow=从该行开始
    '编制:刘兴洪
    '日期:2015-07-13 16:11:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.行) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, _
    Optional blnCommon As Boolean = True, Optional strName As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定药品行的职务是否与当前医生的职务相匹配
    '入参:tmpDetail=输入的项目,不传为所有行
    '     blnCommon=是否正常的判断,否则为医保或公费病人的判断
    '返回：不匹配的行,0为正确
    '编制:刘兴洪
    '日期:2015-07-13 16:12:25
    '说明：职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, int职务A As Integer, int职务B As Integer
    Dim strTmp As String, strAllDuty As String
    
    If cbo开单人.ListIndex = -1 Then Exit Function
    strAllDuty = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Call GetOperatorInfo(mrs开单人, mobjBill.开单人, , int职务A)
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                If Not blnCommon Then
                    int职务B = Val(Right(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            If strName = "" Then
                                strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            Else
                                strTmp = "病人:" & strName & "是医保或公费病人,但第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            End If
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            If strName = "" Then
                                strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                            Else
                                strTmp = "病人:" & strName & "是医保或公费病人,但第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                            End If
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    int职务B = Val(Left(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.类别) = 0 Then Exit Function
        If Not blnCommon Then
            int职务B = Val(Right(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    If strName = "" Then
                        strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    Else
                        strTmp = "病人:" & strName & "是医保或公费病人,但药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    End If
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    If strName = "" Then
                        strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                    Else
                        strTmp = "病人:" & strName & "是医保或公费病人,但药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                    End If
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定药品在单据中是否已经存在
    '入参:objDetail=项目
    '     intRow=要判断的行
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-13 16:13:59
    '说明：时价或分批药品在同一药房禁止重复输入(这里仅提示,保存时禁止)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.分批 Or mobjBill.Details(i).Detail.变价) _
                    And (objDetail.分批 Or objDetail.变价) Then
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价材料,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function
Private Function Check执行科室() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .执行部门ID = 0 Or Bill.TextMatrix(i, BillCol.执行科室) = "" Then
                If InStr(",5,6,7,", .收费类别) = 0 Then
                    Check执行科室 = i: Exit Function
                End If
            End If
        End With
    Next
End Function
Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = UserInfo.部门ID
    End If
End Function

Public Function zl获取中药形态(Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据是否录入了中草药的
    '入参:blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中成药在配方中已经存在,就不需要检查
    '     lngRow-当前操作的行
    '出参:
    '返回:录入了中草药的,则返回中药形态属性(0-散装,1-饮片,2-免煎剂),否则返回-1 表示还没有录入中药形态项目
    '编制:刘兴洪
    '日期:2010-02-02 11:44:17
    '问题:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl获取中药形态 = -1
    '如果未指定页,则用当前页
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly中成药, ",6,", ",6,7,")
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).收费类别 & ",") > 0 And .Item(i).收费细目ID <> 0 And i <> lngRow Then
                zl获取中药形态 = .Item(i).Detail.中药形态
                Exit Function
            End If
        Next
    End With
End Function
Private Function zlCheckBill存在非散装草药() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中存在非散装草药形态
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-26 10:19:46
    '问题:38328
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mobjBill Is Nothing Then Exit Function
    If mobjBill.Details.Count = 0 Then Exit Function
    With mobjBill
        For i = 1 To mobjBill.Details.Count
            If .Details(i).收费类别 = "7" Then
                If .Details(i).Detail.中药形态 <> 0 Then    '0-散装;1-中药饮片;2-免煎剂
                    zlCheckBill存在非散装草药 = True: Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function Get要求审批(ByVal str收费细目ID As String, _
    ByRef rsItem As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取要求审批的医保项目
    '出参:返回需要审批的收费细目记录集(险类,收费细目ID,要求审批)
    '返回:要求审批返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-14 17:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If str收费细目ID > 4000 Then
        strSQL = "" & _
        "   Select A1.收费细目ID,A1.险类,A1.要求审批  " & _
        "   From 保险支付项目 A1,Table(f_Num2List([1])) B1  " & _
        "   Where  a1.险类=b1.Column_value And A1.收费细目ID in (" & str收费细目ID & ")" & _
        "           And nvl(A1.要求审批,0)=1 "
    Else
        strSQL = "" & _
        "   Select A1.收费细目ID,A1.险类,A1.要求审批  " & _
        "   From 保险支付项目 A1,Table(f_Num2List([1])) B1,Table(f_Num2List([2])) B2 " & _
        "   Where  a1.险类=b1.Column_value And A1.收费细目ID=B2.Column_value " & _
        "           And nvl(A1.要求审批,0)=1 "
    End If
    Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInsures, str收费细目ID)
    Get要求审批 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetNurseStationFirstPatiDeptID() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对护士工作产，获取第一个病人的病人部门ID
    '返回:针对护士工作产,返回第一个部门ID,否则返回0
    '编制:刘兴洪
    '日期:2017-11-14 17:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim lng病人ID As Long, i As Long
    Dim lngFirstDeptID As Long, lngFirstSelDeptID As Long
    Dim lngFirtPatiID As Long
    
    On Error GoTo errHandle
     
    If mblnNurseStation = False Then GetNurseStationFirstPatiDeptID = 0: Exit Function
    lngFirtPatiID = 0
    For i = 0 To rptPati.Rows.Count - 1
        If lngFirtPatiID = 0 Then
            lngFirtPatiID = Val(rptPati.Rows(i).Record(COL_病人ID).Value)
            lngFirstDeptID = Val(rptPati.Rows(i).Record(COL_开单科室ID).Value)
        End If
        If rptPati.Rows(i).Record.Tag = "1" Then
            GetNurseStationFirstPatiDeptID = Val(rptPati.Rows(i).Record(COL_开单科室ID).Value)
            Exit Function
        End If
    Next
    GetNurseStationFirstPatiDeptID = lngFirstDeptID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
