VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmNewCheckCard 
   Caption         =   "药品盘点表"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15015
   Icon            =   "frmNewCheckCard.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   15015
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6360
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   24
      Top             =   6600
      Width           =   3855
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "粗体-停用药品"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2520
         TabIndex        =   31
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "盘亏"
         Height          =   180
         Left            =   2040
         TabIndex        =   30
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "盘盈"
         Height          =   180
         Left            =   360
         TabIndex        =   29
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "盘平"
         Height          =   180
         Left            =   1200
         TabIndex        =   28
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   10920
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   14895
      TabIndex        =   0
      Top             =   600
      Width           =   14955
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   2
         Top             =   4080
         Width           =   11355
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   2805
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   12060
         _cx             =   21272
         _cy             =   4948
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmNewCheckCard.frx":06EA
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
         ExplorerBar     =   5
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
      Begin VB.Label lbl修改日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改日期"
         Height          =   180
         Left            =   7140
         TabIndex        =   35
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl修改人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改人"
         Height          =   180
         Left            =   5280
         TabIndex        =   34
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Txt修改人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5880
         TabIndex        =   33
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt修改日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7920
         TabIndex        =   32
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   11760
         TabIndex        =   21
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   10245
         TabIndex        =   20
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   19
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   18
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点库房"
         Height          =   180
         Left            =   270
         TabIndex        =   17
         Top             =   660
         Width           =   720
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品盘点表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10320
         TabIndex        =   14
         Top             =   195
         Width           =   480
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   10800
         TabIndex        =   13
         Top             =   165
         Width           =   1425
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   12
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   11
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12930
         TabIndex        =   10
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10830
         TabIndex        =   9
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额差合计："
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10350
         TabIndex        =   6
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   9480
         TabIndex        =   5
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "盘点金额合计："
         Height          =   180
         Left            =   1920
         TabIndex        =   4
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblCostPrice 
         AutoSize        =   -1  'True
         Caption         =   "盘点成本金额合计："
         Height          =   180
         Left            =   4080
         TabIndex        =   3
         Top             =   3840
         Width           =   1620
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":075F
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":0979
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":0B93
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":0DAD
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":0FC7
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":11E1
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":13FB
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":1615
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":182F
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":1A49
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":1C63
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":1E7D
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":2097
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":22B1
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":24CB
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":26E5
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   7995
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmNewCheckCard.frx":28FF
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13229
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6879
            MinWidth        =   6879
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmNewCheckCard.frx":3193
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmNewCheckCard.frx":3695
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imglvw 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":3B97
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":58A1
            Key             =   "child"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":75AB
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewCheckCard.frx":DE0D
            Key             =   "lapse"
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
   Begin XtremeCommandBars.ImageManager imgTool 
      Bindings        =   "frmNewCheckCard.frx":1466F
      Left            =   1320
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNewCheckCard.frx":14683
   End
End
Attribute VB_Name = "frmNewCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mcon按库存提取 As Integer = 101
Private Const mcon重置 As Integer = 102
Private Const mcon确定 As Integer = 103
Private Const mcon帮助 As Integer = 104
Private Const mcon取消 As Integer = 105
Private Const mcon查找 As Integer = 106
Private Const mcon盘点到最后批次 As Integer = 107
Private Const mcon固定列 As Integer = 108
Private Const mcon实盘数清零 As Integer = 109
Private Const mconFind As Integer = 110
Private Const mcon保存 As Integer = 111
Private Const mcon保存退出 As Integer = 112
Private Const mcon打印 As Integer = 117

Private Const mnuFirst As Integer = 201
Private Const mnuSecond As Integer = 202
Private Const mnuDefault As Integer = 203

Private Const mcon编码和名称 As Integer = 301
Private Const mcon编码 As Integer = 302
Private Const mcon名称 As Integer = 303

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl
Private mcbrToolBar As CommandBar
Private mcbrMenuBar As CommandBarPopup

Private mintSelectStock As Integer           '是否可选库房
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、汇总盘点记录单,产生盘点表;6、全部盘为零;7、库房全部药品盘点;8、特殊药品盘点;9、自动生成有账面数量未盘点的药品
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean                '第一次显示
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Dim mstrPrivs As String                     '权限
Private mblnNoStock As Boolean              '本地参数：是否允许盘点没有设置存储库房的药品
Private mblnLoadData As Boolean             '用于检查是否已装入数据（对于已存在单据）
Private mstr分类ID As String
Private mbln盘停用药品 As Boolean
Private mbln忽略盘点时间 As Boolean         '为真时始终以当前库存作为帐面数量
Private mbln忽略服务对象 As Boolean         '为真时忽略药品的服务对象
Private mbln忽略药品盘点属性 As Boolean     '为真时忽略药品的盘点属性
Private mbln盘亏减可用数量检查 As Boolean     '为真时检查盘亏的药品会不会将可用数量减为负数（mint库存检查<>0时有效）

Private mstrLast盘点时间 As String      '记录上次盘点时间，判断是否需要重新加载记录集

Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private Const MStrCaption As String = "药品盘点表"
Private mbln检查变动 As Boolean         '检查库存是否变动   true-已经检查，false-未检查，只有审核业务才有用
Private mbln即时保存 As Boolean         '编辑后保存不退出编辑界面，还可以继续编辑保存

Private mint盘点类型 As Integer '0.近效期药品，1.毒麻精神药品,2.停用药品,3.无数量但有库存金额的药品,5.基本药物
Private mstr近效期 As String '格式C1:C2;C1表示近效期的天数，为0时C2为1则表示只选择了已失效的
Private mstr毒理 As String 'C1:C2:C3:C4(麻醉药，毒性药，精神I类药，精神II类药),例：1:0:0:0,表示只选择了麻醉药
Private mint盘点模式 As Integer '记录生成盘点的模式（自动、汇总、全部盘为0等与编辑状态有所区别）

Private mlng药品ID As Long '用于定位在该单据的位置
Private mlng批次 As Long '用于定位在该单据的位置
Private mstr药品信息 As String '用于自动生成有账面数量未盘点的药品

Private mstr盘点单号 As String              '盘点单号(记录汇总生成盘点表的盘点单号)
Private mbln删除盘点单 As Boolean           '汇总生成盘点表后是否删除对应的盘点单
Private mblnLoad As Boolean

Private mlngFindFirst As Long
Private mlngFind As Long                             '用于查找
Private mrsFindName As ADODB.Recordset              '用于查找

Private mblnNotTrigger As Boolean
Private mblnKeyPressReturn As Boolean

Private Const mlngColor_盘盈 As Long = vbRed
Private Const mlngColor_盘亏 As Long = vbBlue
Private Const mlngColor_盘平 As Long = vbBlack
Private mlngCurrColor As Long
Private mlngNextColor As Long
'Private blnColorRefresh As Boolean
Private mstrMsg As String
Private mlongCurrRow As Long                '当前选中行
Private mlngFindCurrRow As Long             '查询到的当前行
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mlng库房 As Long

Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库（说明，等于0时有大小包装区分，大于0时为默认包装）
Private mint大单位 As Integer
Private mint小单位 As Integer

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称
Private mbln检查可用数量 As Boolean         '盘亏时检查可用数量：0－不检查；1－检查

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数

Private mintMoneyDigit As Integer           '金额小数位数

Private mintCostDigit0 As Integer            '小单位成本价小数位数
Private mintPriceDigit0 As Integer           '小单位售价小数位数
Private mintNumberDigit0 As Integer          '小单位数量小数位数

Private mintCostDigit1 As Integer            '大单位成本价小数位数
Private mintPriceDigit1 As Integer           '大单位售价小数位数
Private mintNumberDigit1 As Integer          '大单位数量小数位数


Private mintMaxMoneyBit As Integer          '药品库存表中金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

Private mlngSum As Long '记录库存不足药品数量
Private mlng生产商长度 As Long                 '生产商字段长度
Private mlng原产地长度 As Long                 '原产地字段长度
'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol来源 As Integer = 4
Private Const mconIntCol基本药物 As Integer = 5
Private Const mconIntCol序号 As Integer = 6
Private Const mconIntCol规格 As Integer = 7
Private Const mconIntCol批次 As Integer = 8
Private Const mconIntCol可用数量 As Integer = 9
Private Const mconIntCol比例系数 As Integer = 10
Private Const mconIntCol比例系数大 As Integer = 11
Private Const mconIntCol比例系数小 As Integer = 12
Private Const mconIntcol加成率 As Integer = 13
Private Const mconIntCol实际差价 As Integer = 14
Private Const mconIntCol实际金额 As Integer = 15
Private Const mconIntCol产地 As Integer = 16
Private Const mconIntCol原产地 As Integer = 17
Private Const mconIntCol库房货位 As Integer = 18
Private Const mconIntCol单位 As Integer = 19

Private Const mconIntCol批号 As Integer = 20
Private Const mconIntCol效期 As Integer = 21
Private Const mconIntCol批准文号 As Integer = 22

Private Const mconintCol帐面数量 As Integer = 23

Private Const mconintCol大包装帐面数量 As Integer = 24
Private Const mconIntCol帐面数量单位大 As Integer = 25

Private Const mconintCol小包装帐面数量 As Integer = 26
Private Const mconIntCol帐面数量单位小 As Integer = 27

Private Const mconintCol实盘数量 As Integer = 28

Private Const mconintCol大包装实盘数量 As Integer = 29
Private Const mconIntCol实盘数量单位大 As Integer = 30

Private Const mconintCol小包装实盘数量 As Integer = 31
Private Const mconIntCol实盘数量单位小 As Integer = 32

Private Const mconintCol合计 As Integer = 33
Private Const mconintCol当前库存 As Integer = 34
Private Const mconintCol可用数量占用 As Integer = 35
Private Const mconintCol标志 As Integer = 36
Private Const mconintCol数量差 As Integer = 37
Private Const mconintCol成本价 As Integer = 38
Private Const mconIntCol售价 As Integer = 39
Private Const mconintCol金额差 As Integer = 40
Private Const mconintCol差价差 As Integer = 41
Private Const mconintCol盘点金额 As Integer = 42
Private Const mconintCol盘点成本金额 As Integer = 43
Private Const mconintCol盘点成本金额差 As Integer = 44
Private Const mconintCol库存数量 As Integer = 45      '取库存原始数量
Private Const mconIntCol药品编码和名称 As Integer = 46
Private Const mconIntCol药品编码 As Integer = 47
Private Const mconIntCol药品名称 As Integer = 48
Private Const mconIntCol新批次 As Integer = 49
Private Const mconIntCol排序编码 As Integer = 50
Private Const mconIntCol分批属性 As Integer = 51
Private Const mconIntColS  As Integer = 52              '总列数
'=========================================================================================


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.id
        '文件
        Case mcon编码和名称, mcon编码, mcon名称
            If mcon编码和名称 = Control.id Then cbsColDrug 0
            If mcon编码 = Control.id Then cbsColDrug 1
            If mcon名称 = Control.id Then cbsColDrug 2
        Case mcon帮助
            cbsHelp
        Case mcon按库存提取
            cbsBatch
        Case mnuFirst
            cbsFirst
        Case mnuSecond
            cbsSecond
        Case mnuDefault
            cbsDefault
        Case mcon重置
            cbsReset
        Case mcon盘点到最后批次
            cbsSet
        Case mcon实盘数清零
            cbsZero
        Case mcon确定, mcon保存, mcon保存退出
            cbsSave Control.id
        Case mcon打印
            cbsPrintBill
        Case mcon取消
            cbsCancel
    End Select
    
End Sub


Private Sub cbsColDrug(Index As Integer)
    Dim n As Integer
    
    With mobjPopup
        For n = 1 To .Controls.count
            .Controls.Item(n).Checked = False
        Next
        
        .Controls.Item(Index + 1).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    If mblnLoad = False Then Exit Sub
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Me.Pic单据.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - staThis.Height
    
    initControl
    
End Sub

Private Sub Form_Load()
    
    mblnLoad = False
    
    InitComandBars
    
    Call GetDefineSize
    
    mlngFindCurrRow = 1
    mbln检查可用数量 = (Val(zlDataBase.GetPara("盘亏时检查可用数量", glngSys, 模块号.药品盘点)) = 1)
    mblnNoStock = (Val(zlDataBase.GetPara("存储库房", glngSys, 模块号.药品盘点)) = 1)
    mintMaxMoneyBit = gtype_UserDrugDigits.Digit_金额
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    mbln忽略服务对象 = (Val(zlDataBase.GetPara("忽略药品服务对象", glngSys, 模块号.药品盘点)) = 1)
    
    mbln盘亏减可用数量检查 = (Val(zlDataBase.GetPara("盘亏减可用数量检查", glngSys, 模块号.药品盘点)) = 1)
    
    txtStock = mfrmMain.cboStock.Text
    txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    mlng库房 = txtStock.Tag
    Call Get大小单位
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品盘点管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mobjPopup.Controls.Item(mintDrugNameShow + 1).Checked = True
    
    mblnLoadData = False
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    initCard
    
    mstrTime_Start = GetBillInfo(12, mstr单据号)
    
    staThis.Panels(3).Picture = picColor
    
    mblnLoad = True
End Sub


Private Sub InitComandBars()
    '初始化工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With

    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons
    
    
    '工具栏定义
    Set mcbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny Or xtpFlagHideWrap
    
    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mcon按库存提取, "按库存提取")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon重置, "重置")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字

        Set cbrControlMain = .Add(xtpControlButton, mcon盘点到最后批次, "盘点到最后批次")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon实盘数清零, "实盘数清零")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        

        
        Set cbrControlMain = .Add(xtpControlPopup, mcon固定列, "固定列"): cbrControlMain.BeginGroup = True
        cbrControlMain.id = mcon固定列
        cbrControlMain.IconId = mcon固定列
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mnuFirst, "从药名到单位列(&1)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mnuSecond, "从药名到效期列(&2)", -1, False
        cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mnuDefault, "恢复(&D)", -1, False).BeginGroup = True
        cbrControlMain.Visible = False
        
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon打印, "打印")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon保存, "保存")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon确定, "保存新增")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon保存退出, "保存退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon取消, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        Set cbrControlMain = .Add(xtpControlButton, mcon帮助, "帮助")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        
        Set cbrControlMain = .Add(xtpControlLabel, mcon查找, "查找")
        cbrControlMain.Flags = xtpFlagRightAlign    '靠右对齐

        Set ctrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, mconFind, "查询")
        ctrCustom.Handle = txtCode.hWnd
        ctrCustom.Flags = xtpFlagRightAlign
    End With

    cbsMain.Item(1).Delete
    
    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_ESCAPE, mcon取消
    End With
    
    '右键菜单
    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopup.Controls
        Set mobjControl = .Add(xtpControlButton, mcon编码和名称, "药名(编码和名称)")
        Set mobjControl = .Add(xtpControlButton, mcon编码, "药名(仅编码)")
        Set mobjControl = .Add(xtpControlButton, mcon名称, "药名(仅名称)")
    End With

End Sub



Private Sub Form_Resize()
    If mblnLoad = False Then Exit Sub
    initControl
End Sub

Private Sub initControl()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Width <= 15255 Then Me.Width = 15255
    If Me.Height <= 7800 Then Me.Height = 7800
    
    '重新布局控件位置
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With
    
    With vsfBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    
    With txtNo
        .Left = vsfBill.Left + vsfBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    txtCheckDate.Left = vsfBill.Left + vsfBill.Width - txtCheckDate.Width
    lblCheckDate.Left = txtCheckDate.Left - lblCheckDate.Width - 100
    
    LblStock.Left = vsfBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100

    With Lbl填制人
        .Top = Pic单据.Height - 200 - .Height
        .Left = vsfBill.Left + 100
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With
    
    With Lbl填制日期
        .Top = Lbl填制人.Top
        .Left = Txt填制人.Left + Txt填制人.Width + 250
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With lbl修改人
        .Top = Lbl填制人.Top
        .Left = Pic单据.Width / 2 - (450 + Txt修改人.Width + lbl修改人.Width + Txt修改日期.Width + lbl修改日期.Width) / 2
    End With
    
    With Txt修改人
        .Top = Lbl填制人.Top - 80
        .Left = lbl修改人.Left + lbl修改人.Width + 100
    End With
    
    With lbl修改日期
        .Top = Lbl填制人.Top
        .Left = Txt修改人.Left + Txt修改人.Width + 250
    End With
    
    With Txt修改日期
        .Top = Lbl填制人.Top - 80
        .Left = lbl修改日期.Left + lbl修改日期.Width + 100
    End With
    
    With Txt审核日期
        .Top = Lbl填制人.Top - 80
        .Left = vsfBill.Left + vsfBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制人.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl审核日期.Left - 200 - .Width
    End With
    
    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Txt审核人.Left - 100 - .Width
    End With
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = vsfBill.Left + vsfBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = vsfBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = Pic单据.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic单据.TextWidth(lblCheckSum.Caption) + 200
    End With
    
    With lblCostPrice
        .Top = lblCheckSum.Top
        .Left = lblCheckSum.Left + lblCheckSum.Width + 200
    End With
    
    With vsfBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(7).Width - IIf(staThis.Panels(4).Visible, staThis.Panels(4).Width, 0) - IIf(staThis.Panels(5).Visible, staThis.Panels(5).Width, 0) - staThis.Panels(6).Width - .Width - 300
    End With
    
    
    If mint编辑状态 = 1 Then
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon按库存提取, , True).Visible = True
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon重置, , True).Visible = True
'        cmdBatch.Visible = True
'        cmdReset.Visible = True
    ElseIf mint编辑状态 = 5 Then
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon按库存提取, , True).Visible = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon重置, , True).Visible = True
'        cmdBatch.Visible = False
'        cmdReset.Visible = True
    Else
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon按库存提取, , True).Visible = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon重置, , True).Visible = False
'        cmdBatch.Visible = False
'        cmdReset.Visible = False
    End If
    
    Me.cbsMain(1).Controls.Find(xtpControlButton, mcon盘点到最后批次, , True).Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8)
    Me.cbsMain(1).Controls.Find(xtpControlButton, mcon实盘数清零, , True).Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8)
    Me.cbsMain(1).Controls.Find(xtpControlButton, mcon保存退出, , True).Visible = mint编辑状态 = 1 Or mint编辑状态 = 7 Or mint编辑状态 = 8
    Me.cbsMain(1).Controls.Find(xtpControlButton, mcon保存, , True).Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8)
    If mint编辑状态 = 7 Or mint编辑状态 = 8 Then Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Visible = False
    If mint编辑状态 = 8 Then Me.cbsMain(1).Controls.Find(xtpControlButton, mcon重置, , True).Visible = True
    If mint编辑状态 <> 4 Then Me.cbsMain(1).Controls.Find(xtpControlButton, mcon打印, , True).Visible = False '非查看不可见
'    cmdSet.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
'    cmdZero.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
    
End Sub


Private Function CheckUnVerify(ByVal lng库房ID As Long) As Boolean
    '检查未审核单据：返回真表示通过检查
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select id From 药品收发记录" & _
            " Where 审核人 Is NULL And 库房ID=[1] AND Rownum<2 "
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "检查未审核单据", lng库房ID)
    If rsData.EOF Then
        CheckUnVerify = True
    Else
        CheckUnVerify = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Get大小单位()
    Dim intUnit As Integer, strUnit As String, strDefault As String
    Dim strCompare As String
    Dim str大小单位 As String
    Dim int性质 As Integer
    
    Const conInt计算精度 As Integer = 0
    
    Const conInt药品 As Integer = 1
    
    Const conint售价单位 As Integer = 1
    Const conint门诊单位 As Integer = 2
    Const conint住院单位 As Integer = 3
    Const conint药库单位 As Integer = 4
    
    Const conInt成本价 As Integer = 1
    Const conInt售价 As Integer = 2
    Const conInt数量 As Integer = 3
    Const conInt金额 As Integer = 4
    
    int性质 = conInt计算精度
        
    strCompare = "药库单位;门诊单位;住院单位;售价单位"
    
    '取得大包装单位
    strDefault = GetDrugUnit(Val(txtStock.Tag), "药品盘点管理")
    
    '取得小包装单位
    intUnit = Val(zlDataBase.GetPara("小包装单位", glngSys, 模块号.药品盘点))
    
    If intUnit = 0 Then
        strUnit = strDefault
    Else
        strUnit = Split(strCompare, ";")(intUnit - 1)
    End If

    '将指定单位与缺省单位按大单位、小单位的顺序排列
    mintUnit = 0
    If strUnit <> strDefault Then
        If InStr(1, strCompare, strUnit) < InStr(1, strCompare, strDefault) Then
            str大小单位 = strUnit & "|" & strDefault
        Else
            mintUnit = 0
            str大小单位 = strDefault & "|" & strUnit
        End If
        
        mintMoneyDigit = GetDigit(int性质, conInt药品, conInt金额)
        
        Call GetDrugDigit(mlng库房, "药品盘点管理", 0, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    Else
        Call GetDrugDigit(mlng库房, "药品盘点管理", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If
    
    If str大小单位 = "" Then Exit Sub
    
    '取大单位的精度（售价、数量、金额）
    Select Case Split(str大小单位, "|")(0)
        Case "售价单位"
            mint大单位 = conint售价单位
        Case "门诊单位"
            mint大单位 = conint门诊单位
        Case "住院单位"
            mint大单位 = conint住院单位
        Case "药库单位"
            mint大单位 = conint药库单位
    End Select
    
    mintCostDigit1 = GetDigit(int性质, conInt药品, conInt成本价, mint大单位)
    mintPriceDigit1 = GetDigit(int性质, conInt药品, conInt售价, mint大单位)
    mintNumberDigit1 = GetDigit(int性质, conInt药品, conInt数量, mint大单位)

    '取小单位的精度（数量）
    Select Case Split(str大小单位, "|")(1)
        Case "售价单位"
            mint小单位 = conint售价单位
        Case "门诊单位"
            mint小单位 = conint门诊单位
        Case "住院单位"
            mint小单位 = conint住院单位
        Case "药库单位"
            mint小单位 = conint药库单位
    End Select
    
    mintCostDigit0 = GetDigit(int性质, conInt药品, conInt成本价, mint小单位)
    mintPriceDigit0 = GetDigit(int性质, conInt药品, conInt售价, mint小单位)
    mintNumberDigit0 = GetDigit(int性质, conInt药品, conInt数量, mint小单位)
    
'    '数量小数按最大精度取值，否则可能盘不干净
'    mintNumberDigit = gtype_UserDrugDigits.Digit_数量
'    mintNumberDigit0 = gtype_UserDrugDigits.Digit_数量
End Sub

Private Sub RefreshListSN()
    '用于排序后更新序号
    Dim lngRow As Long
    
    With vsfBill
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                .TextMatrix(lngRow, mconIntCol行号) = lngRow
            End If
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If vsfBill.rows < 2 Then Exit Sub
    If vsfBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfBill.rows - 1
            If vsfBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(vsfBill.TextMatrix(n, mconIntCol序号)) = 0, n, Val(vsfBill.TextMatrix(n, mconIntCol序号)))
                !药品id = Val(vsfBill.TextMatrix(n, 0))
                !批次 = Val(vsfBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
           & "FROM 药品单据性质 A, 药品入出类别 B " _
           & "Where A.类别id = B.ID AND A.单据 = 12  and b.系数=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品盘点表的入库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    strSQL = "SELECT B.Id " _
           & "FROM 药品单据性质 A, 药品入出类别 B " _
           & "Where A.类别id = B.ID AND A.单据 = 12  and b.系数=-1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品盘点表的出库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetStocktakingColor(ByVal vsfObj As VSFlexGrid, ByVal Row As Long)
    '盘亏盘盈行用颜色区分：蓝色字体-盘盈；红色字体-盘亏；黑色字体-盘平
    With vsfObj
        .Row = Row
        mlngCurrColor = .CellForeColor
        If .TextMatrix(Row, mconintCol标志) = "盈" Then
            mlngNextColor = mlngColor_盘盈
        ElseIf .TextMatrix(Row, mconintCol标志) = "亏" Then
            mlngNextColor = mlngColor_盘亏
        Else
            mlngNextColor = mlngColor_盘平
        End If
        
        If mlngNextColor <> mlngCurrColor Then
            .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = mlngNextColor
        End If
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional blnSuccess As Boolean = False, Optional lng药品ID As Long = 0, Optional lng批次 As Long = 0, Optional str药品信息 As String = "")
    Dim strTitle As String
    
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1307)
    
    mlng药品ID = lng药品ID
    mlng批次 = lng批次
    mstr药品信息 = str药品信息
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Or mint编辑状态 = 9 Then
        mblnEdit = True
        If mint编辑状态 = 5 Or mint编辑状态 = 6 Or mint编辑状态 = 9 Then Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "确定"
    ElseIf mint编辑状态 = 2 Then
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "保存退出"
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "审核"
'        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Visible = False
'        CmdSave.Caption = "打印(&P)"
        If mint编辑状态 = 4 Then
            If InStr(mstrPrivs, "单据打印") = 0 Then
                Me.cbsMain(1).Controls.Find(xtpControlButton, mcon打印, , True).Visible = False
    '            CmdSave.Visible = False
            Else
                Me.cbsMain(1).Controls.Find(xtpControlButton, mcon打印, , True).Visible = True
    '            CmdSave.Visible = True
            End If
        End If
    End If
    
    '1.新增；2、修改；3、验收；4、查看；5、汇总盘点记录单,产生盘点表;6、全部盘为零;7、库房全部药品盘点
    If mint编辑状态 = 1 Then
        strTitle = "(自动生成盘点表)"
    ElseIf mint编辑状态 = 5 Then
        strTitle = "(汇总记录单产生盘点表)"
    ElseIf mint编辑状态 = 6 Then
        strTitle = "(全部盘为零)"
    ElseIf mint编辑状态 = 7 Then
        strTitle = "(库房全部药品盘点)"
    ElseIf mint编辑状态 = 8 Then
        strTitle = "(特殊药品盘点)"
    End If
    
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption & strTitle
    
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub vsfBill_MoveNextCell(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntCol药名
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                .Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
            Case mconIntCol批号
                If (Val(.TextMatrix(Row, mconIntCol批次)) = -1 Or Val(.TextMatrix(Row, mconIntCol新批次)) = 1) And .TextMatrix(Row, mconIntCol效期) = "" Then
                    .Col = mconIntCol效期
                Else
                    .Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
                End If
            Case mconIntCol效期
                .Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
            Case mconintCol实盘数量
                If Row < .rows - 1 Then
                    .Row = Row + 1
                    If .TextMatrix(.Row, mconIntCol药名) = "" Then
                        .Col = mconIntCol药名
                    Else
                        .Col = mconintCol实盘数量
                    End If
                Else
                    If Val(.TextMatrix(Row, 0)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                        .Col = mconIntCol药名
                    End If
                End If
            Case mconintCol大包装实盘数量, mconintCol小包装实盘数量
                If Col = mconintCol大包装实盘数量 Then
                    If .ColWidth(mconintCol小包装实盘数量) > 0 Then
                        .Col = mconintCol小包装实盘数量
                    Else
                        '如果下一行为空或者药名列为空则返回到药名列，否则返回到实盘数量列
                        If Row < .rows - 1 Then
                            .Row = Row + 1
                            If .TextMatrix(Row, mconIntCol药名) <> "" Then
                                .Col = mconintCol大包装实盘数量
                            Else
                                .Col = mconIntCol药名
                            End If
                        Else
                            If Val(.TextMatrix(Row, 0)) <> 0 Then
                                .rows = .rows + 1
                                .Row = .rows - 1
                                .Col = mconIntCol药名
                            End If
                        End If
                    End If
                Else
                    If Row < .rows - 1 Then
                        .Row = Row + 1
                        If .TextMatrix(Row, mconIntCol药名) <> "" Then
                            .Col = mconintCol大包装实盘数量
                        Else
                            .Col = mconIntCol药名
                        End If
                    Else
                        If Val(.TextMatrix(Row, 0)) <> 0 Then
                            .rows = .rows + 1
                            .Row = .rows - 1
                            .Col = mconIntCol药名
                        End If
                    End If
                End If
        End Select
        
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub cbsBatch()
    '保证库存中有的记录都提取出来
    Dim rsPhysic As ADODB.Recordset '药品库存记录集
    Dim rsDetail As ADODB.Recordset
    Dim str盘点属性 As String
    Dim dbl成本价 As Double, dbl零售价 As Double, dbl加成率 As Double
    Dim bln库房 As Boolean
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim intCol As Integer
    Dim rs时价分批 As ADODB.Recordset
    Dim str药名 As String
    Dim strOrder As String, strCompare As String
    Dim str盘点时间 As String
    Dim dbl盘点金额 As Double

    str盘点时间 = txtCheckDate.Caption
    
    If MsgBox("重置条件，界面中已有数据将清除，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    Else
        With vsfBill
            .rows = 2
            For intCol = 0 To .Cols - 1
                .TextMatrix(1, intCol) = ""
            Next
        End With
    End If
    
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
    strCompare = Mid(strOrder, 1, 1)
    
    gstrSQL = "Select  Distinct a.药品id, b.编码, b.名称, c.库房货位 " & _
        " From 药品库存 A, 收费项目目录 B, 药品储备限额 C " & _
        " where　a.性质 = 1 And a.药品id = b.Id And a.库房id = c.库房id(+) And a.药品id = c.药品id(+) And a.库房id = [1]" & _
        " And (Nvl(A.实际数量,0)<>0 Or Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0 )"

    
    If mbln忽略服务对象 = False Then
        gstrSQL = gstrSQL & _
            " and (Decode(B.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3)) " & _
                " or Decode(B.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]))"
    End If
    
    gstrSQL = gstrSQL & " Order by " & _
              IIf(strCompare = "0", "B.编码", IIf(strCompare = "1", "B.编码", IIf(strCompare = "2", "B.名称", "C.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",B.编码"
    
    Set rsPhysic = zlDataBase.OpenSQLRecord(gstrSQL, "查询库存药品", Val(txtStock.Tag))
    With vsfBill
        Do While Not rsPhysic.EOF
            '取该药品的详细信息（可能分多个批次）
            Set rsDetail = GetPhysicDetail(Val(txtStock.Tag), rsPhysic!药品id, False, False, False)
            Do While Not rsDetail.EOF
                If (rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1) And .TextMatrix(.rows - 1, 0) <> "" Then .rows = .rows + 1
                '时价药品重算售价
                dbl成本价 = Nvl(rsDetail!平均成本价, 0)
                dbl零售价 = Nvl(rsDetail!售价, 0)
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(CLng(rsPhysic!药品id), Val(txtStock.Tag), CLng(rsDetail!批次), 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = rsPhysic!药品id
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = zlStr.Nvl(rsDetail!产地, zlStr.Nvl(rsDetail!缺省产地))
                .TextMatrix(.rows - 1, mconIntCol原产地) = zlStr.Nvl(rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                .TextMatrix(.rows - 1, mconintCol标志) = "平"
                .TextMatrix(.rows - 1, mconintCol数量差) = "0"
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol帐面数量), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = .TextMatrix(.rows - 1, mconintCol大包装帐面数量)
                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & .TextMatrix(.rows - 1, mconIntCol实盘数量单位小)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If
                
                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                '如果是分批药品，将批次改填为-1，表示新增批次
                .TextMatrix(.rows - 1, mconIntCol批次) = zlStr.Nvl(rsDetail!批次, 0)
                If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(.rows - 1, mconIntCol批次)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol批次) = -1
'                    '调试用，自动为新增批次设置批号与效期
'                    .TextMatrix(.Rows - 1, mconIntCol批号) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntCol效期) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                If mbln盘停用药品 = True Then
                    '如果是停用药品，该行粗体显示
                    If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                    End If
                End If
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + Val(.TextMatrix(.rows - 1, mconintCol金额差))) - (zlStr.Nvl(rsDetail!实际差价, 0) + Val(.TextMatrix(.rows - 1, mconintCol差价差))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmNewCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
    Else
        vsfBill.Col = mconIntCol药名
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
'        vsfBill.EditCell
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsCancel()
    Unload Me
End Sub

Private Sub cbsPrintBill()
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub cbsHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cbsReset()
    Dim str用途ID As String, str库房货位 As String, str剂型编码 As String, strALL剂型编码 As String
    Dim str材质分类 As String, lng库房ID As Long, int盘点方式 As Integer, str盘点时间 As String
    Dim int盘无库存药品 As Integer, bln盘点单 As Boolean   '是否只针对盘点单中的药品进行盘点，FALSE-表示对所有药品进行盘点，盘点单中不存在的药品自动盘为零
    Dim bln盘无库存有金额药品 As Boolean
    Dim intCol As Integer
    Dim cbrMenuControl As CommandBarControl
    
'    If mblnFirst = False Then Exit Sub
    
    With vsfBill
        If MsgBox("重置条件，界面中已有数据将清除，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End With
    
    mblnLoadData = False
    If mintParallelRecord <> 1 Then mblnChange = False
    
    '初始化变量
    str用途ID = "": str剂型编码 = ""
    
    If mint编辑状态 = 1 Then
        '自动搜索或手工输入盘点表
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str剂型编码, lng库房ID, int盘点方式, str盘点时间, int盘无库存药品, str库房货位, bln盘无库存有金额药品, mstr分类ID, mbln忽略盘点时间) = True Then
            If mlng库房 = 0 Then
                mlng库房 = lng库房ID
            End If
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
'            Call Get大小单位
            Call SearchData(str剂型编码, lng库房ID, int盘点方式, str盘点时间, (int盘无库存药品 = 1), str库房货位, bln盘无库存有金额药品)
        Else
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
            Exit Sub
        End If
        
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon取消, , True)
        
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    ElseIf mint编辑状态 = 5 Then
        '产生盘点表（汇总指定时刻的盘点记录单与指定时刻的库存）
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng库房ID, mstr盘点单号, bln盘点单, mbln删除盘点单) = True Then
            If mlng库房 = 0 Then
                mlng库房 = lng库房ID
            End If
            vsfBill.rows = 2
            Call Get大小单位
            Call SearchTableData(lng库房ID, bln盘点单)
        Else
            Exit Sub
        End If
        
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon取消, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
        
        
    ElseIf mint编辑状态 = 8 Then
        '产生盘点表（特殊药品盘点）
         Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
         
        If Not frmCheckDate.GetCondition(Me, str盘点时间, 8, mint盘点类型, mstr近效期, mstr毒理) Then
            Unload Me
            Exit Sub
        End If
        
        vsfBill.rows = 2
        For intCol = 0 To vsfBill.Cols - 1
            vsfBill.TextMatrix(1, intCol) = ""
            vsfBill.Cell(flexcpPicture, 1, mconIntCol效期, 1, mconIntCol效期) = Nothing
        Next
        
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房ID
        mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng库房 = 0 Then
            mlng库房 = lng库房ID
        End If
        Call Get大小单位
        
        SearchSpecialData lng库房ID, mint盘点类型, mstr近效期, mstr毒理
    
    End If
    
    mblnLoadData = True
End Sub

Private Sub cbsSet()
    Dim lngRow As Long, n As Long
    Dim rsDetail As ADODB.Recordset
    Dim lng药品ID As Long, lng批次 As Long, dbl实盘数量 As Double
    Dim dlbSum As Double
    Dim intMoneyBit As Integer
    Dim dbl金额差 As Double, dbl差价差 As Double
    Dim dbl盘点金额 As Double
    
    On Error GoTo ErrHand
    
    If MsgBox("该操作将药品的实盘数量汇总到最后批次上，是否进行该操作？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    '考虑表格排序，可能相同药品不是连续的，先把界面数据装入数据集处理
    Set rsDetail = New ADODB.Recordset
    With rsDetail
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "实盘数量", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfBill.rows - 1
            If vsfBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !药品id = Val(vsfBill.TextMatrix(n, 0))
                !批次 = Val(vsfBill.TextMatrix(n, mconIntCol批次))
                !实盘数量 = Val(vsfBill.TextMatrix(n, mconintCol实盘数量))
                
                .Update
            End If
        Next
        
        .Sort = "药品id,批次"
        
        Do While Not .EOF
            If lng药品ID <> !药品id Then
                dlbSum = !实盘数量
                lng药品ID = !药品id
            Else
                dlbSum = dlbSum + !实盘数量
            End If
            
            !实盘数量 = 0
            .Update
            
            .MoveNext
            
            '如果后面已经没有数据了或者后面不是同一个药品时，将实盘数量汇总到最后一个批次上
            If .EOF Then
                .MovePrevious
                !实盘数量 = dlbSum
                .Update
                
                .MoveNext
            Else
                If lng药品ID <> !药品id Then
                    .MovePrevious
                    !实盘数量 = dlbSum
                    .Update
                    
                    .MoveNext
                End If
            End If
        Loop
    End With
    
    
    
    With vsfBill
        .Redraw = flexRDNone
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                lng药品ID = Val(vsfBill.TextMatrix(lngRow, 0))
                lng批次 = Val(vsfBill.TextMatrix(lngRow, mconIntCol批次))
                
                rsDetail.Filter = "药品id=" & lng药品ID & " And 批次=" & lng批次
                If Not rsDetail.EOF Then
                    '按数据集的实盘数量更新盘点数据
                    dbl实盘数量 = rsDetail!实盘数量
                    
                    '换算成大小包装单位
                    If mintUnit = 0 Then
                        .TextMatrix(lngRow, mconintCol大包装实盘数量) = zlStr.FormatEx(Int(dbl实盘数量 / Val(.TextMatrix(lngRow, mconIntCol比例系数大))), mintNumberDigit0, , True)
                        .TextMatrix(lngRow, mconintCol小包装实盘数量) = zlStr.FormatEx((dbl实盘数量 - Val(.TextMatrix(lngRow, mconintCol大包装实盘数量)) * Val(.TextMatrix(lngRow, mconIntCol比例系数大))) / Val(.TextMatrix(lngRow, mconIntCol比例系数小)), mintNumberDigit0, , True)
                        .TextMatrix(lngRow, mconintCol合计) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit, , True) & .TextMatrix(lngRow, mconIntCol帐面数量单位小)
                    End If
                    
                    .TextMatrix(lngRow, mconintCol实盘数量) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit, , True)
                    .TextMatrix(lngRow, mconintCol数量差) = zlStr.FormatEx(Abs(dbl实盘数量 - Val(.TextMatrix(lngRow, mconintCol帐面数量))), mintNumberDigit, , True)
                    If dbl实盘数量 > Val(.TextMatrix(lngRow, mconintCol帐面数量)) Then
                        .TextMatrix(lngRow, mconintCol标志) = "盈"
                    ElseIf dbl实盘数量 < Val(.TextMatrix(lngRow, mconintCol帐面数量)) Then
                        .TextMatrix(lngRow, mconintCol标志) = "亏"
                    Else
                        .TextMatrix(lngRow, mconintCol标志) = "平"
                    End If
                    
                    '实盘数量为0与库存数量比较判断盈亏
                    If Val(.TextMatrix(lngRow, mconintCol实盘数量)) = 0 And Val(.TextMatrix(lngRow, mconintCol库存数量)) > 0 Then
                        .TextMatrix(lngRow, mconintCol标志) = "亏"
                    End If
                
                    '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
                    '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
                    '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
                    If Val(.TextMatrix(lngRow, mconIntCol新批次)) = 1 Then
                        intMoneyBit = mintMoneyDigit
                    ElseIf dbl实盘数量 = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(lngRow, 0))) = True And Val(.TextMatrix(lngRow, mconIntCol售价)) = Val(.TextMatrix(lngRow, mconintCol成本价))) Then
                        '盘0或者零差价药品盘点时
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                
                    '金额差=当前售价*实盘数量-实际金额
                    '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                    dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol售价)) * dbl实盘数量, mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                    .TextMatrix(lngRow, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(lngRow, mconIntCol实际金额)), intMoneyBit, , True)
                    .TextMatrix(lngRow, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol售价)) - Val(.TextMatrix(lngRow, mconintCol成本价))) * dbl实盘数量 - Val(.TextMatrix(lngRow, mconIntCol实际差价)), intMoneyBit, , True)
                    dbl金额差 = Val(.TextMatrix(lngRow, mconintCol金额差))
                    dbl差价差 = Val(.TextMatrix(lngRow, mconintCol差价差))
                    If .TextMatrix(lngRow, mconintCol标志) = "亏" Then
                        .TextMatrix(lngRow, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol金额差)), intMoneyBit, , True)
                        .TextMatrix(lngRow, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol差价差)), intMoneyBit, , True)
                    End If
                
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(lngRow, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol售价)) * dbl实盘数量, mintMoneyDigit, , True)
                    
                    '.TextMatrix(lngRow, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol成本价)) * Val(.TextMatrix(lngRow, mconintCol实盘数量)), mintMoneyDigit)
                    '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                    .TextMatrix(lngRow, mconintCol盘点成本金额) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol实际金额)) + dbl金额差) - (Val(.TextMatrix(lngRow, mconIntCol实际差价)) + dbl差价差), mintMoneyDigit, , True)
                    .TextMatrix(lngRow, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol金额差)) - Val(.TextMatrix(lngRow, mconintCol差价差)), intMoneyBit, , True)
                
                    '盘亏盘盈行用颜色区分
                    Call SetStocktakingColor(vsfBill, lngRow)
                End If
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
    
    Me.MousePointer = vbDefault
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsZero()
    Dim lngRow As Integer
    Dim dbl实盘数量 As Double
    Dim dbl金额差 As Double, dbl差价差 As Double
    Dim intMoneyBit As Integer
    Dim dbl盘点金额 As Double
    
    If MsgBox("是否把实盘数清零？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    dbl实盘数量 = 0
    
    With vsfBill
        .Redraw = flexRDNone
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
    
                '换算成大小包装单位
                If mintUnit = 0 Then
                      .TextMatrix(lngRow, mconintCol大包装实盘数量) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit0, , True)
                      .TextMatrix(lngRow, mconintCol小包装实盘数量) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit0, , True)
                      .TextMatrix(lngRow, mconintCol合计) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit, , True) & .TextMatrix(lngRow, mconIntCol帐面数量单位小)
                End If
              
                .TextMatrix(lngRow, mconintCol实盘数量) = zlStr.FormatEx(dbl实盘数量, mintNumberDigit, , True)
                .TextMatrix(lngRow, mconintCol数量差) = zlStr.FormatEx(Abs(dbl实盘数量 - Val(.TextMatrix(lngRow, mconintCol帐面数量))), mintNumberDigit, , True)
                If dbl实盘数量 > Val(.TextMatrix(lngRow, mconintCol帐面数量)) Then
                    .TextMatrix(lngRow, mconintCol标志) = "盈"
                ElseIf dbl实盘数量 < Val(.TextMatrix(lngRow, mconintCol帐面数量)) Then
                    .TextMatrix(lngRow, mconintCol标志) = "亏"
                Else
                    .TextMatrix(lngRow, mconintCol标志) = "平"
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(lngRow, mconintCol实盘数量)) = 0 And Val(.TextMatrix(lngRow, mconintCol库存数量)) > 0 Then
                    .TextMatrix(lngRow, mconintCol标志) = "亏"
                End If
                
                  intMoneyBit = mintMaxMoneyBit
        
                  '金额差=当前售价*实盘数量-实际金额
                  '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                  dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol售价)) * dbl实盘数量, mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                  .TextMatrix(lngRow, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(lngRow, mconIntCol实际金额)), intMoneyBit, , True)
                  .TextMatrix(lngRow, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol售价)) - Val(.TextMatrix(lngRow, mconintCol成本价))) * dbl实盘数量 - Val(.TextMatrix(lngRow, mconIntCol实际差价)), intMoneyBit, , True)
                  dbl金额差 = Val(.TextMatrix(lngRow, mconintCol金额差))
                  dbl差价差 = Val(.TextMatrix(lngRow, mconintCol差价差))
                  If .TextMatrix(lngRow, mconintCol标志) = "亏" Then
                      .TextMatrix(lngRow, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol金额差)), intMoneyBit, , True)
                      .TextMatrix(lngRow, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol差价差)), intMoneyBit, , True)
                  End If
          
                  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  .TextMatrix(lngRow, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol售价)) * dbl实盘数量, mintMoneyDigit, , True)
        
                  '.TextMatrix(lngRow, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol成本价)) * Val(.TextMatrix(lngRow, mconintCol实盘数量)), mintMoneyDigit)
                  '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                  .TextMatrix(lngRow, mconintCol盘点成本金额) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol实际金额)) + dbl金额差) - (Val(.TextMatrix(lngRow, mconIntCol实际差价)) + dbl差价差), mintMoneyDigit, , True)
                  .TextMatrix(lngRow, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol金额差)) - Val(.TextMatrix(lngRow, mconintCol差价差)), intMoneyBit, , True)
              
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, lngRow)
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If Trim(txtCode.Text) = "" Then
            txtCode.SetFocus
        Else
            Call FindGridRow(txtCode.Text)
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CheckDataUpdate()
    '检查数据是否发生变化，如果变化则提示用户并自动更新界面数据
    '只有审核时才调用此过程
    Dim intRow As Integer
    Dim lng药品ID As Long
    Dim lng库房ID As Long
    Dim lng批次 As Long
    Dim dat盘点时间 As Date
    Dim dbl原账面数量 As Double
    Dim dbl现账面数量 As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim intMoneyBit As Integer
    Dim rsTemp As ADODB.Recordset
    Dim bln变动 As Boolean
    Dim dbl盘点金额 As Double
    
    On Error GoTo ErrHand
    
    If mint编辑状态 = 3 Then
        With vsfBill
            If .rows > 1 Then
                Call FS.ShowFlash("正在药品变动,请稍候 ...", Me)
                
                lng库房ID = txtStock.Tag
                .Redraw = flexRDNone
                For intRow = 1 To .rows - 1
                    If Val(.TextMatrix(intRow, 0)) <> 0 Then
                        lng药品ID = Val(.TextMatrix(intRow, 0))
                        lng批次 = Val(.TextMatrix(intRow, mconIntCol批次))
                        dat盘点时间 = CDate(txtCheckDate.Caption)
                        dbl原账面数量 = Val(.TextMatrix(intRow, mconintCol库存数量))
                        
                        gstrSQL = "Select 库房id, 药品id, 批次, Nvl(Sum(实际数量), 0) As 账面数量, Nvl(Sum(盘点数量), 0) As 盘点数量, Nvl(Sum(实际金额), 0) As 实际金额," & vbNewLine & _
                                    "       Nvl(Sum(实际差价), 0) As 实际差价, Nvl(Sum(可用数量), 0) As 可用数量" & vbNewLine & _
                                    "From (Select a.库房id, a.药品id, Nvl(批次, 0) As 批次, Nvl(a.实际数量, 0) 实际数量, 0 盘点数量, Nvl(a.实际金额, 0) 实际金额, Nvl(a.实际差价, 0) 实际差价," & vbNewLine & _
                                    "              Nvl(a.可用数量, 0) 可用数量" & vbNewLine & _
                                    "       From 药品库存 A" & vbNewLine & _
                                    "       Where a.性质 = 1 And a.库房id = [1] And a.药品id = [2] And Nvl(a.批次, 0) = [3]" & vbNewLine & _
                                    "       Union All" & vbNewLine & _
                                    "       Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, Sum(-1 * a.入出系数 * a.实际数量 * a.付数) As 实际数量, 0 盘点数量," & vbNewLine & _
                                    "              Sum(-1 * a.入出系数 * a.零售金额) As 实际金额, Sum(-1 * a.入出系数 * a.差价) As 实际差价, 0 As 可用数量" & vbNewLine & _
                                    "       From 药品收发记录 A" & vbNewLine & _
                                    "       Where a.库房id + 0 = [1] And a.药品id + 0 = [2] And Nvl(a.批次, 0) = [3] And a.审核日期 > [4]" & vbNewLine & _
                                    "       Group By a.库房id, a.药品id, Nvl(a.批次, 0))" & vbNewLine & _
                                    "Group By 库房id, 药品id, 批次"

                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "当前库存检查", lng库房ID, lng药品ID, lng批次, dat盘点时间)
                        
                        If rsTemp.RecordCount > 0 Then
                            dbl现账面数量 = rsTemp!账面数量
                            If dbl原账面数量 <> dbl现账面数量 Then
                                bln变动 = True
                                                                
                                .TextMatrix(intRow, mconintCol库存数量) = Nvl(rsTemp!账面数量, 0)
                                .TextMatrix(intRow, mconIntCol实际金额) = zlStr.Nvl(rsTemp!实际金额, 0)
                                .TextMatrix(intRow, mconIntCol实际差价) = zlStr.Nvl(rsTemp!实际差价, 0)
                                If mintUnit > 0 Then
                                    .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsTemp!账面数量, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数)), mintNumberDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsTemp!账面数量, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), mintNumberDigit0, , True)
                                    
                                    .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsTemp!账面数量, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大))), mintNumberDigit0, , True)
                                    .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsTemp!账面数量) - Val(.TextMatrix(intRow, mconintCol大包装帐面数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), mintNumberDigit0, , True)
                                    .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsTemp!可用数量, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), mintNumberDigit0, , True)
                                End If

                                If Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol售价)) = Val(.TextMatrix(intRow, mconintCol成本价))) Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                                    intMoneyBit = mintMaxMoneyBit
                                Else
                                    intMoneyBit = mintMoneyDigit
                                End If
                                
                                .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx(Abs(Val(.TextMatrix(intRow, mconintCol实盘数量)) - Val(.TextMatrix(intRow, mconintCol帐面数量))), mintNumberDigit, , True)
                                If Val(.TextMatrix(intRow, mconintCol实盘数量)) > Val(.TextMatrix(intRow, mconintCol帐面数量)) Then
                                    .TextMatrix(intRow, mconintCol标志) = "盈"
                                ElseIf Val(.TextMatrix(intRow, mconintCol实盘数量)) < Val(.TextMatrix(intRow, mconintCol帐面数量)) Then
                                    .TextMatrix(intRow, mconintCol标志) = "亏"
                                Else
                                    .TextMatrix(intRow, mconintCol标志) = "平"
                                End If
                                
                                '实盘数量为0与库存数量比较判断盈亏
                                If Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 And Val(.TextMatrix(intRow, mconintCol库存数量)) > 0 Then
                                    .TextMatrix(intRow, mconintCol标志) = "亏"
                                End If

                                '金额差=当前售价*实盘数量-实际金额
                                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                                .TextMatrix(intRow, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(intRow, mconIntCol实际金额)), intMoneyBit, , True)
                                .TextMatrix(intRow, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol售价)) - Val(.TextMatrix(intRow, mconintCol成本价))) * Val(.TextMatrix(intRow, mconintCol实盘数量)) - Val(.TextMatrix(intRow, mconIntCol实际差价)), intMoneyBit, , True)
                                dbl金额差 = Val(.TextMatrix(intRow, mconintCol金额差))
                                dbl差价差 = Val(.TextMatrix(intRow, mconintCol差价差))
                                If .TextMatrix(intRow, mconintCol标志) = "亏" Then
                                    .TextMatrix(intRow, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol金额差)), intMoneyBit, , True)
                                    .TextMatrix(intRow, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)
                                End If
                            
                                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                                .TextMatrix(intRow, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsTemp!实际金额, 0) + dbl金额差) - (zlStr.Nvl(rsTemp!实际差价, 0) + dbl差价差), mintMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol金额差)) - Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)

                            End If
                        End If
                    End If
                Next
                .Redraw = flexRDDirect
                If bln变动 = True Then
                    MsgBox "库存发生变化，将自动更新界面数据，请检查！", vbInformation, gstrSysName
                    mbln检查变动 = True
                End If
            End If
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsSave(ByVal lngControlId As Long)
'参数：lngControlId表示保存的模式，参数结合 mint编辑状态 使用
    Dim blnSuccess As Boolean
    Dim intLop As Integer
    Dim str药品 As String '记录可用数量不足时的药品，充足则为空
    Dim intNumberDigit As Integer
    Dim intNumberDigit0 As Integer
    Dim dbl换算系数 As Double
    
    '设置排序数据集
    Call SetSortRecord
    
    If mint编辑状态 = 3 Then        '审核
    
        '自动批量检查并执行调价
        Call AutoAdjustPrice_ByNO(12, mstr单据号)
        
        mstrTime_End = GetBillInfo(12, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        '发生了变动现将原始盘点单删除然后再产生NO相同的新的盘点单
        If mbln检查变动 = True Then
            blnSuccess = SaveCard
        End If
        If mbln检查变动 = False Then
            '检查库存是否发生变化
            Call CheckDataUpdate
            If mbln检查变动 = True Then
                Exit Sub
            End If
        End If
        
        '零差价管理：检查是否存在不满足零差价的药品
        For intLop = 1 To vsfBill.rows - 1
            If Val(vsfBill.TextMatrix(intLop, mconIntCol新批次)) = 0 Then
                '不是新增批次时
                If vsfBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                    If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                        If CheckPriceAdjust(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol批次))) = False Then
                            MsgBox "第" & intLop & "行药品已启用零差价管理，但库存记录中售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                            vsfBill.SetFocus
                            vsfBill.Row = intLop
                            vsfBill.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            Else
                '新增时
                If vsfBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                    If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                        '如果是零差价管理，检查界面售价和成本价关系
                        If Val(vsfBill.TextMatrix(intLop, mconintCol成本价)) <> Val(vsfBill.TextMatrix(intLop, mconIntCol售价)) Then
                            MsgBox "第" & intLop & "行药品已启用零差价管理，但盘点界面的售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                            vsfBill.SetFocus
                            vsfBill.Row = intLop
                            vsfBill.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            If vsfBill.TextMatrix(intLop, mconintCol标志) = "亏" Then '盘亏时出库，检查库存是否足够
                'mintUnit-单位系数：1-售价;2-门诊;3-住院;4-药库（说明，等于0时有大小包装区分，大于0时为默认包装）
                dbl换算系数 = IIf(mintUnit > 0, Val(vsfBill.TextMatrix(intLop, mconIntCol比例系数)), Val(vsfBill.TextMatrix(intLop, mconIntCol比例系数小)))
            
                If Not 库存实际数量检查(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol批次)), _
                Val(vsfBill.TextMatrix(intLop, mconintCol数量差)), dbl换算系数, IIf(mintUnit > 0, intNumberDigit, intNumberDigit0)) Then
                    mlngSum = mlngSum + 1
                    If mlngSum <= 3 Then '拼提示信息串
                        If mlngSum = 3 Then mstrMsg = mstrMsg & "【" & vsfBill.TextMatrix(intLop, mconIntCol药名) & "(" & vsfBill.TextMatrix(intLop, mconIntCol批号) & "）" & "】，"
                        If mlngSum <> 3 Then mstrMsg = mstrMsg & "【" & vsfBill.TextMatrix(intLop, mconIntCol药名) & "(" & vsfBill.TextMatrix(intLop, mconIntCol批号) & "）" & "】" & Chr(10)
		    End If
                    
                    If mlngSum = 1 Then vsfBill.Row = intLop: vsfBill.TopRow = intLop
                End If
                
            End If
        Next
        
        
        '库存不足提示信息
        If mlngSum > 0 Then
            If mint库存检查 = 1 Then '不足提醒
                If MsgBox(mstrMsg & IIf(mlngSum <= 3, mlngSum & "个药品库存不足，是否继续？", "等" & mlngSum & "个药品库存不足，是否继续？"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mlngSum = 0
                    mstrMsg = ""
                    Exit Sub
                End If
            ElseIf mint库存检查 = 2 Then '不足禁止
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个药品库存不足，不能审核！", "等" & mlngSum & "个药品库存不足，不能审核！"), vbInformation, gstrSysName
                mlngSum = 0
                mstrMsg = ""
                Exit Sub
            End If
        End If
        mlngSum = 0
        mstrMsg = ""
        
        If SaveCheck = True Then
            If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品盘点)) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    
    If Not 盘亏减可用数量检查 Then Exit Sub
        
    blnSuccess = SaveCard
    
    Me.MousePointer = vbDefault
        
    If blnSuccess = True Then
            
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.药品盘点)) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then    '修改保存后处理
            If lngControlId = mcon确定 Then
                Unload Me
                Exit Sub
            Else
                mblnChange = False '即时保存后mblnChange = False
                MsgBox "保存成功！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    
    If mint编辑状态 = 1 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Then    '新增保存后处理
        If lngControlId = mcon保存退出 Then
            Unload Me
            Exit Sub
        End If
        
        If lngControlId = mcon保存 Then
            txtNo.Caption = txtNo.Tag
            mstr单据号 = txtNo.Tag
            mbln即时保存 = True
            mblnChange = False '即时保存后mblnChange = False
            MsgBox "保存成功！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If lngControlId = mcon确定 And (mint编辑状态 <> 7 Or mint编辑状态 <> 8) Then '全盘(特殊药品盘点)无需保存后新增（保存新增按钮也不会显示）
            txtNo.Caption = ""
            mstr单据号 = ""
            mbln即时保存 = False
            txtCheckDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
            mblnSave = False
            mblnEdit = True
            vsfBill.rows = 2
            vsfBill.Cell(flexcpText, 1, 0, 1, vsfBill.Cols - 1) = ""
        
            Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
            txt摘要.Text = ""
            mblnChange = False
            
            If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub
'功能：检查盘亏药品在减可用数量后数量是否为负，为负则根据库存检查参数是提醒还是禁止
Private Function 盘亏减可用数量检查() As Boolean
    Dim dbl帐面数量 As Double
    Dim dbl数量差 As Double
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl比例系数 As Double
    Dim rsData As ADODB.Recordset
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim strMsg As String
    Dim intSum As Integer
    Dim intFirstRow As Integer '记录第一次出现提示的行
    
    On Error GoTo errHandle
    
    盘亏减可用数量检查 = False
    
    If mint库存检查 = 0 Then '不检查库存则盘亏不检查减可用数量后是否为负
        盘亏减可用数量检查 = True
        Exit Function
    Else
        If Not mbln盘亏减可用数量检查 Then '检查库存但未勾选“盘亏减可用数量检查”
            盘亏减可用数量检查 = True
            Exit Function
        End If
    End If
    
    
    
    lng库房ID = txtStock.Tag
    
    With vsfBill
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            
            If .TextMatrix(intRow, 0) <> "" And Trim(.TextMatrix(intRow, mconintCol标志)) = "亏" Then '只有盘亏才检查
                lng药品ID = Val(.TextMatrix(intRow, 0))
                lng批次 = IIf(.TextMatrix(intRow, mconIntCol批次) = "", 0, Val(.TextMatrix(intRow, mconIntCol批次)))
                
                '获取数量差
                dbl比例系数 = IIf(mintUnit > 0, Val(.TextMatrix(intRow, mconIntCol比例系数)), Val(.TextMatrix(intRow, mconIntCol比例系数小)))
                dbl帐面数量 = Val(.TextMatrix(intRow, mconintCol库存数量))
                If Trim(.TextMatrix(intRow, mconintCol标志)) = "平" Then
                    If dbl帐面数量 <> Val(.TextMatrix(intRow, mconintCol帐面数量)) * dbl比例系数 Then
                        '真实库存账面数量和界面账面数量换算后的不一致时(由于精度取舍导致的，可能导致盘点后无法得到预期的实盘数量)
                        '使用真实库存数量来和实盘数量计算数量差
                        dbl数量差 = Val(.TextMatrix(intRow, mconintCol实盘数量)) * dbl比例系数 - dbl帐面数量
                    Else
                        dbl数量差 = 0
                    End If
                Else
                    dbl数量差 = zlStr.FormatEx(Abs(.TextMatrix(intRow, mconintCol实盘数量) * dbl比例系数 - Val(.TextMatrix(intRow, mconintCol库存数量))), gtype_UserDrugDigits.Digit_数量, , True)
                End If
                
'                gstrSQL = "select 可用数量 from 药品库存 where 库房id = [1] and 药品id = [2] and nvl(批次,0) = [3] and 性质 = 1"
                gstrSQL = "Select Sum(可用数量) As 可用数量" & vbNewLine & _
                        " From (Select Nvl(可用数量, 0) As 可用数量" & vbNewLine & _
                        "       From 药品库存" & vbNewLine & _
                        "       Where 性质=1 And 库房id = [1] And 药品id = [2]  And nvl(批次,0) = [3]" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select Abs(a.实际数量 * Nvl(a.付数, 1)) As 可用数量" & vbNewLine & _
                        "       From 药品收发记录 A" & vbNewLine & _
                        "       Where a.审核日期 Is Null And a.库房id = [1] And a.药品id + 0 = [2]  And a.No = [4] And a.单据 = 12  And nvl(批次,0) = [3] )"

                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", lng库房ID, lng药品ID, lng批次, mstr单据号)
                
                If rsData.RecordCount = 0 Then '无库存记录
                    intSum = intSum + 1
                    If intSum = 1 Then intFirstRow = intRow
                    
                    If intSum <= 3 Then '只记录前三条药品药名
                        If intSum = 1 Then
                            strMsg = .TextMatrix(intRow, mconIntCol药名)
                        Else
                            strMsg = strMsg & "，" & .TextMatrix(intRow, mconIntCol药名)
                        End If
                    End If
                    
                Else
                    If zlStr.Nvl(rsData!可用数量, 0) < dbl数量差 Then '库存减可用数量后为负
                        intSum = intSum + 1
                        If intSum = 1 Then intFirstRow = intRow
                        
                        If intSum <= 3 Then '只记录前三条药品药名
                            If intSum = 1 Then
                                strMsg = .TextMatrix(intRow, mconIntCol药名)
                            Else
                                strMsg = strMsg & "，" & .TextMatrix(intRow, mconIntCol药名)
                            End If
                        End If
                        
                    End If
                End If
                
            End If
            recSort.MoveNext
        Next
        
    End With
    
    If intSum <> 0 Then
        vsfBill.Row = intFirstRow
        vsfBill.TopRow = intFirstRow
        If mint库存检查 = 1 Then
            If MsgBox("药品" & strMsg & IIf(intSum <= 3, "保存后可用数量将小于零，", "等" & intSum & "个保存后可用数量将小于零，") & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            MsgBox "药品" & strMsg & IIf(intSum <= 3, "保存后可用数量将小于零，", "等" & intSum & "个保存后可用数量将小于零，") & "不能保存！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    盘亏减可用数量检查 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    Dim str用途ID As String, str库房货位 As String, str剂型编码 As String, strALL剂型编码 As String
    Dim str材质分类 As String, lng库房ID As Long, int盘点方式 As Integer, str盘点时间 As String
    Dim int盘无库存药品 As Integer, bln盘点单 As Boolean   '是否只针对盘点单中的药品进行盘点，FALSE-表示对所有药品进行盘点，盘点单中不存在的药品自动盘为零
    Dim bln盘无库存有金额药品 As Boolean
    Dim cbrMenuControl As CommandBarControl
    
    
    If mblnFirst = False Then Exit Sub
    If mblnLoad = False Then Exit Sub
    
    
    mstr分类ID = ""
    mblnLoadData = False
    mintBatchNoLen = GetBatchNoLen()
    If mintParallelRecord <> 1 Then mblnChange = False
    
    mbln盘停用药品 = IIf(Val(zlDataBase.GetPara("盘已停用的药品", glngSys, 1307, 0)) = 0, False, True)
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 5
            MsgBox "还存在未审核的药品单据，请全部审核后再试！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
     
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    '初始化变量
    str用途ID = "": str剂型编码 = ""
    
    If mint编辑状态 = 1 Then
        '自动搜索或手工输入盘点表
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str剂型编码, lng库房ID, int盘点方式, str盘点时间, int盘无库存药品, str库房货位, bln盘无库存有金额药品, mstr分类ID, mbln忽略盘点时间) = True Then
            If mlng库房 = 0 Then
                mlng库房 = lng库房ID
            End If
            Call Get大小单位
            Call SearchData(str剂型编码, lng库房ID, int盘点方式, str盘点时间, (int盘无库存药品 = 1), str库房货位, bln盘无库存有金额药品)
        Else
            Unload Me
            Exit Sub
        End If
        
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon取消, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
        
    ElseIf mint编辑状态 = 5 Then
        '产生盘点表（汇总指定时刻的盘点记录单与指定时刻的库存）
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng库房ID, mstr盘点单号, bln盘点单, mbln删除盘点单) = True Then
            If mlng库房 = 0 Then
                mlng库房 = lng库房ID
            End If
            Call Get大小单位
            Call SearchTableData(lng库房ID, bln盘点单)
        Else
            Unload Me
            Exit Sub
        End If
        
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon取消, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    ElseIf mint编辑状态 = 6 Then
        '全部盘为零
        str盘点时间 = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房ID
        mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng库房 = 0 Then
            mlng库房 = lng库房ID
        End If
        Call Get大小单位
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        Call SearchTableData(lng库房ID)
        
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon取消, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True)
        If cbrMenuControl.Enabled = False Then
            cbrMenuControl.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    ElseIf mint编辑状态 = 7 Then '库房全部药品盘点
        
        If Not frmCheckDate.GetCondition(Me, str盘点时间, 7, 0, "", "") Then
            Unload Me
            Exit Sub
        End If
        
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房ID
        mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng库房 = 0 Then
            mlng库房 = lng库房ID
        End If
        Call Get大小单位
        
        SearchAllData lng库房ID
    ElseIf mint编辑状态 = 8 Then '特殊药品盘点
        If Not frmCheckDate.GetCondition(Me, str盘点时间, 8, mint盘点类型, mstr近效期, mstr毒理) Then
            Unload Me
            Exit Sub
        End If
        
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房ID
        mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng库房 = 0 Then
            mlng库房 = lng库房ID
        End If
        Call Get大小单位
        
        SearchSpecialData lng库房ID, mint盘点类型, mstr近效期, mstr毒理
    ElseIf mint编辑状态 = 9 Then '自动生成有账面数量未盘点的药品
        If Not frmCheckDate.GetCondition(Me, str盘点时间, 9, 0, "", "") Then
            Unload Me
            Exit Sub
        End If
        
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房ID
        mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng库房 = 0 Then
            mlng库房 = lng库房ID
        End If
        Call Get大小单位
        
        SearchMisslData mlng库房, mstr药品信息
    End If
    
    mblnLoadData = True
End Sub

Private Sub SetSortCode()
    '根据药品编码返回格式化的排序编码
    '编码中可能含有"-"符号，查找所有编码中"-"前最多几位，"-"后最多几位，所有编码都按最大位数进行格式化处理
    Dim str编码 As String
    Dim lngRow As Long
    Dim int前缀 As Integer
    Dim int后缀 As Integer
    Dim str编码前缀 As String
    Dim str编码后缀 As String
    Dim blnLine As Boolean
    
    With vsfBill
       For lngRow = 1 To vsfBill.rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                str编码 = Replace(.TextMatrix(lngRow, mconIntCol药品编码), "[", "")
                str编码 = Replace(str编码, "]", "")
                
                If InStr(1, str编码, "-") > 0 Then
                    blnLine = True
                    If Len(Mid(str编码, 1, InStr(str编码, "-") - 1)) > int前缀 Then
                        int前缀 = Len(Mid(str编码, 1, InStr(str编码, "-") - 1))
                    End If
                    
                    If Len(Mid(str编码, InStr(str编码, "-") + 1)) > int后缀 Then
                        int后缀 = Len(Mid(str编码, InStr(str编码, "-") + 1))
                    End If
                Else
                    If Len(str编码) > int前缀 Then
                        int前缀 = Len(str编码)
                    End If
                End If
            End If
        Next
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                str编码 = Replace(.TextMatrix(lngRow, mconIntCol药品编码), "[", "")
                str编码 = Replace(str编码, "]", "")
                
                If blnLine = False Then
                    .TextMatrix(lngRow, mconIntCol排序编码) = Format(str编码, String(int前缀, "0"))
                Else
                    If InStr(str编码, "-") > 0 Then
                        str编码前缀 = Mid(str编码, 1, InStr(str编码, "-") - 1)
                        str编码后缀 = Mid(str编码, InStr(str编码, "-") + 1)
                        
                        str编码前缀 = Format(str编码前缀, String(int前缀, "0"))
                        str编码后缀 = Format(str编码后缀, String(int后缀, "0"))
                    Else
                        str编码前缀 = Format(str编码, String(int前缀, "0"))
                        str编码后缀 = String(int后缀, "0")
                    End If
                    
                    .TextMatrix(lngRow, mconIntCol排序编码) = str编码前缀 & "-" & str编码后缀
                End If
            End If
        Next
    End With
End Sub
Private Sub SearchData(ByVal str剂型编码 As String, ByVal lng库房ID As Long, _
    ByVal int盘点方式 As Integer, ByVal str盘点时间 As String, ByVal bln盘无库存药品 As Boolean, ByVal str库房货位 As String, ByVal bln盘无库存有金额药品 As Boolean)
    
    Dim rsPhysic As ADODB.Recordset '药品库存记录集
    Dim rsDetail As ADODB.Recordset
    Dim str盘点属性 As String
    Dim dbl成本价 As Double, dbl零售价 As Double, dbl加成率 As Double
    Dim bln库房 As Boolean
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim n As Integer
    Dim rs时价分批 As ADODB.Recordset
    Dim str药名 As String
    Dim rsTemp As ADODB.Recordset
    Dim strArry As Variant
    Dim x As Long
    Dim strTemp As String
    Dim j As Long
    Dim str货位id As String
    Dim str货位 As String
    Dim dbl盘点金额 As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    
    '初始化数据集
    Set rsPhysic = New ADODB.Recordset
    With rsPhysic
        If .State = 1 Then .Close
        .Fields.Append "药品id", adDouble, 18, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '设置界面显示内容
    Select Case int盘点方式
        Case 1
            staThis.Panels(2).Text = "现在对" & txtStock & "的药品进行日盘点"
        Case 2
            staThis.Panels(2).Text = "现在对" & txtStock & "的药品进行周盘点"
        Case 3
            staThis.Panels(2).Text = "现在对" & txtStock & "的药品进行月盘点"
        Case 4
            staThis.Panels(2).Text = "现在对" & txtStock & "的药品进行季度盘点"
        Case 5
            staThis.Panels(2).Text = "现在对所有的药品进行季度盘点"
    End Select
    str盘点属性 = " And Substr(A.盘点属性," & int盘点方式 & ",1)='1'"
    If int盘点方式 = 5 Then str盘点属性 = "所有"
    Call FS.ShowFlash("正在计算药品库存数据,请稍候 ...", Me)
    DoEvents
    
    x = 1
    strArry = Array()
    str货位id = ""
    For j = 0 To UBound(Split(str库房货位, ",")) - 1
        str货位 = Mid(str库房货位, x, InStr(x, str库房货位, ",") - x)
        x = InStr(x, str库房货位, ",") + 1
        If Len(IIf(str货位id = "", "", str货位id & ",") & str货位) > 4000 Then
            ReDim Preserve strArry(UBound(strArry) + 1)
            strArry(UBound(strArry)) = str货位id
            str货位id = str货位
        Else
            str货位id = IIf(str货位id = "", "", str货位id & ",") & str货位
        End If
    Next
    
    If str货位id <> "" Then
        ReDim Preserve strArry(UBound(strArry) + 1)
        strArry(UBound(strArry)) = str货位id
    End If
    
    If str库房货位 = "" Then
        Set rsPhysic = GetPhysic(lng库房ID, str盘点属性, str剂型编码, str库房货位, bln盘无库存药品, False, False, bln盘无库存有金额药品)
    Else
        For j = 0 To UBound(strArry)
            Set rsTemp = GetPhysic(lng库房ID, str盘点属性, str剂型编码, CStr(strArry(j)), bln盘无库存药品, False, False, bln盘无库存有金额药品)
            If Not rsTemp.EOF Then
                Do While Not rsTemp.EOF
                    With rsPhysic
                        .AddNew
                        !药品id = rsTemp!药品id
                        !编码 = rsTemp!编码
                        !名称 = rsTemp!名称
                        !库房货位 = rsTemp!库房货位
                        
                        .Update
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        Next
    End If
    
    Call FS.StopFlash
    
    If rsPhysic.RecordCount = 0 Then
        If mint编辑状态 = 6 Then
            MsgBox "未能正确读取药品库存数据,请重试！", vbInformation, gstrSysName: Exit Sub
        Else
            MsgBox "未能正确读取药品库存数据,请重试或手工输入药品！", vbInformation, gstrSysName
            vsfBill.Row = 1
            vsfBill.Col = mconIntCol药名
            Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("正在装入药品数据,请稍候 ...", Me)
    DoEvents
    vsfBill.Redraw = flexRDNone
    
    bln库房 = CheckPartProp(lng库房ID)
    With vsfBill
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            '取该药品的详细信息（可能分多个批次）
            Set rsDetail = GetPhysicDetail(lng库房ID, rsPhysic!药品id, bln盘无库存药品, False, bln盘无库存有金额药品)
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                '时价药品重算售价
                dbl成本价 = zlStr.Nvl(rsDetail!平均成本价, 0)
                dbl零售价 = zlStr.Nvl(rsDetail!售价, 0)
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(CLng(rsPhysic!药品id), lng库房ID, CLng(rsDetail!批次), 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = rsPhysic!药品id
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = zlStr.Nvl(rsDetail!产地, zlStr.Nvl(rsDetail!缺省产地))
                .TextMatrix(.rows - 1, mconIntCol原产地) = zlStr.Nvl(rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                .TextMatrix(.rows - 1, mconintCol标志) = "平"
                .TextMatrix(.rows - 1, mconintCol数量差) = "0"
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol帐面数量), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = .TextMatrix(.rows - 1, mconintCol大包装帐面数量)

                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & .TextMatrix(.rows - 1, mconIntCol实盘数量单位小)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If
                
                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                '如果是分批药品，将批次改填为-1，表示新增批次
                .TextMatrix(.rows - 1, mconIntCol批次) = zlStr.Nvl(rsDetail!批次, 0)
                If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(.rows - 1, mconIntCol批次)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol批次) = -1
'                    '调试用，自动为新增批次设置批号与效期
'                    .TextMatrix(.Rows - 1, mconIntCol批号) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntCol效期) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                                
                If mbln盘停用药品 = True Then
                    '如果是停用药品，该行粗体显示
                    If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                    End If
                End If
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + Val(.TextMatrix(.rows - 1, mconintCol金额差))) - (zlStr.Nvl(rsDetail!实际差价, 0) + Val(.TextMatrix(.rows - 1, mconintCol差价差))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmNewCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
    Else
        vsfBill.Col = mconIntCol药名
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
        vsfBill.EditCell
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SearchTableData(ByVal lng库房ID As Long, Optional ByVal bln盘点单 As Boolean = False)
    Dim strPhysic As String
    Dim dbl成本价 As Double, dbl零售价 As Double, dbl加成率 As Double
    Dim lngPhysic As Long
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsPhysic As New ADODB.Recordset '药品库存记录集
    Dim rsDetail As New ADODB.Recordset
    Dim n As Integer
    Dim intOld As Integer
    Dim rs时价分批 As ADODB.Recordset
    Dim str药名 As String
    Dim lngDrugID As Long
    Dim rsDingPrice As ADODB.Recordset
    Dim intMoneyBit As Integer
    Dim dbl金额差, dbl差价差 As Double
    Dim str盘点时间 As String
    Dim dbl盘点金额 As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    
    str盘点时间 = txtCheckDate.Caption
    
    Call FS.ShowFlash("正在计算药品库存数据,请稍候 ...", Me)
    DoEvents
    Set rsPhysic = GetPhysic(lng库房ID, "所有", "所有", "所有", False, IIf(mint编辑状态 = 5, True, False), bln盘点单)
    Call FS.StopFlash
    
    If rsPhysic.RecordCount = 0 Then
        If mint编辑状态 = 6 Then
            MsgBox "未能正确读取药品库存数据,请重试！", vbInformation, gstrSysName: Exit Sub
        Else
            MsgBox "未能正确读取药品库存数据,请重试或手工输入药品！", vbInformation, gstrSysName: Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("正在装入药品数据,请稍候 ...", Me)
    DoEvents
    
    With vsfBill
        .Redraw = flexRDNone
        Do While Not rsPhysic.EOF
            Set rsDetail = GetPhysicDetail(lng库房ID, rsPhysic!药品id, False, IIf(mint编辑状态 = 5, True, False))
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                dbl成本价 = zlStr.Nvl(rsDetail!成本价, 0)
                dbl零售价 = IIf(IsNull(rsDetail!售价), 0, rsDetail!售价)
                '处理在盘点后又新增了的药品
                If rsDetail!是否变价 = 0 And IsNull(rsDetail!售价) Then
                    gstrSQL = "select 现价 from 收费价目 where 收费细目id=[1] and sysdate between 执行日期 and 终止日期" & _
                            GetPriceClassString("")
                    lngDrugID = rsPhysic!药品id
                    
                    Set rsDingPrice = zlDataBase.OpenSQLRecord(gstrSQL, "定价价格", lngDrugID)
                    If rsDingPrice.EOF = False Then
                        dbl零售价 = rsDingPrice!现价
                    End If
                End If
                
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(CLng(rsDetail!药品id), lng库房ID, CLng(rsDetail!批次), 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                If Nvl(rsDetail!批次, 0) = -1 Then
                    '分批药品没有批次就是新增盘点入库
                    .TextMatrix(.rows - 1, mconIntCol新批次) = "1"
                ElseIf CheckNoStock(Val(txtStock.Tag), Val(rsDetail!药品id), Nvl(rsDetail!批次, 0)) = True Then
                    '无库存时盘点就是新增盘点入库
                    .TextMatrix(.rows - 1, mconIntCol新批次) = "1"
                End If
                
                '零差价管理：新增盘点入库时对价格进行处理
                If gtype_UserSysParms.P275_零差价管理模式 = 2 And .TextMatrix(.rows - 1, mconIntCol新批次) = "1" Then
                    If IsPriceAdjustMod(Val(rsDetail!药品id)) = True Then
                        If rsDetail!是否变价 = 1 Then
                            '时价时，售价=成本价
                            dbl零售价 = dbl成本价
                        Else
                            '定价时，成本价=售价
                            dbl成本价 = dbl零售价
                        End If
                    End If
                End If

                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = rsDetail!药品id
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = IIf(IsNull(rsDetail!基本药物), "", rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = IIf(IsNull(rsDetail!产地), "", rsDetail!产地)
                .TextMatrix(.rows - 1, mconIntCol原产地) = IIf(IsNull(rsDetail!原产地), "", rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                End If
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol批次) = IIf(IsNull(rsDetail!批次), "0", rsDetail!批次)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                
'                If mint编辑状态 <> 5 Then
'                    .TextMatrix(.rows - 1, mconintCol数量差) =Str.FormatEx(rsDetail!数量差, mintNumberDigit)
'                End If
                If mint编辑状态 = 5 Then
                    If mintUnit > 0 Then
                        .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!盘点数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    Else
                        .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!盘点数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                        .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & rsDetail!小包装单位
                    End If
                Else
                    '单独处理盘为0时的数量的精度位数，以最大显示
                    mintNumberDigit = 5
                    mintNumberDigit0 = 5
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                End If
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(Val(zlStr.Nvl(rsDetail!当前库存, 0)) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!盘点数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx((Val(rsDetail!盘点数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装实盘数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    '单独处理盘为0时的数量的精度位数，以最大显示
                    If mint编辑状态 = 6 Then
                        mintNumberDigit = 5
                        mintNumberDigit0 = 5
                        .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                        .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                    End If
                    
                End If
                
                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                
                If Val(.TextMatrix(.rows - 1, mconintCol帐面数量)) > Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                ElseIf Val(.TextMatrix(.rows - 1, mconintCol帐面数量)) < Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "盈"
                Else
                    .TextMatrix(.rows - 1, mconintCol标志) = "平"
                End If
                
                .TextMatrix(.rows - 1, mconintCol数量差) = zlStr.FormatEx(Abs(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconintCol帐面数量))), mintNumberDigit, , True)
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                dbl金额差 = Val(.TextMatrix(.rows - 1, mconintCol金额差))
                dbl差价差 = Val(.TextMatrix(.rows - 1, mconintCol差价差))
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + dbl金额差) - (zlStr.Nvl(rsDetail!实际差价, 0) + dbl差价差), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                .Col = mconintCol实盘数量
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmNewCheckCard)
            rsPhysic.MoveNext
        Loop
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1: vsfBill.Col = mconintCol实盘数量
    If Me.Visible = True Then
        vsfBill.SetFocus
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim n As Integer
    Dim intOld As Integer
    Dim intMoneyBit As Integer
    Dim str药名 As String
    Dim strSqlOrder As String
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim dbl当前库存大 As Double, dbl当前库存小 As Double
    Dim dbl数量占用大 As Double, dbl数量占用小 As Double
    Dim bln当前库存 As Boolean, bln可用数量占用 As Boolean '对应列是否隐藏
    Dim lng药品定位 As Long
    Dim intNumberDigit As Integer
    Dim intNumberDigit0 As Integer
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        strSqlOrder = "药品编码"
    ElseIf strCompare = "2" Then
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            strSqlOrder = "通用名"
        Else
            strSqlOrder = "Nvl(商品名, 通用名)"
        End If
    ElseIf strCompare = "3" Then
        strSqlOrder = "库房货位"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",药品编码,序号"
    
    mfrmMain.MousePointer = vbHourglass
    Select Case mint编辑状态
        Case 1, 5, 6, 7, 8, 9
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'            Txt修改人 = UserInfo.用户姓名
'            Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid

            Dim cbrMenuControl As CommandBarControl
            Set cbrMenuControl = Me.cbsMain(1).Controls.Find(xtpControlPopup, mcon固定列, , True)
            cbrMenuControl.Visible = mint编辑状态 = 1
'            cmd固定列.Visible = (mint编辑状态 = 1)
        Case 2, 3, 4
            initGrid
            
            bln当前库存 = vsfBill.colHidden(mconintCol当前库存)
            bln可用数量占用 = vsfBill.colHidden(mconintCol可用数量占用)
            
            If mint编辑状态 <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
                mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
            Else
                gstrSQL = "select distinct b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id " _
                    & "and A.单据 = 12 and a.no=[1] "
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!名称
                txtStock.Tag = rsInitCard!id
                mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
                rsInitCard.Close
            End If
            
            If mintUnit > 0 Then
                '大小包装相同时
                Select Case mintUnit
                    Case mconint售价单位
                        strUnitQuantity = "I.计算单位 AS 单位, A.填写数量 AS 帐面数量,A.扣率 AS 实盘数量, A.实际数量 AS 数量差,'1' as 比例系数,a.零售价 as 售价,A.单量 成本价," & IIf(bln当前库存, "", "d.实际数量 as 当前库存 ,") & IIf(bln可用数量占用, "", "y.可用数量占用 as 可用数量占用,")
                    Case mconint门诊单位
                        strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量/ B.门诊包装) AS 帐面数量,(A.扣率/ B.门诊包装) AS 实盘数量, (A.实际数量 / B.门诊包装) AS 数量差,B.门诊包装 as 比例系数,a.零售价*B.门诊包装 as 售价,(A.单量* B.门诊包装) 成本价," & IIf(bln当前库存, "", "d.实际数量/ B.门诊包装 as 当前库存 ,") & IIf(bln可用数量占用, "", "(y.可用数量占用/ B.门诊包装) as 可用数量占用 ,")
                    Case mconint住院单位
                        strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量/ B.住院包装) AS 帐面数量,(A.扣率/ B.住院包装) AS 实盘数量, (A.实际数量 / B.住院包装) AS 数量差,B.住院包装 as 比例系数,a.零售价*B.住院包装 as 售价,(A.单量*B.住院包装) 成本价," & IIf(bln当前库存, "", "d.实际数量/ B.住院包装 as 当前库存 ,") & IIf(bln可用数量占用, "", "(y.可用数量占用/ B.住院包装) as 可用数量占用 ,")
                    Case mconint药库单位
                        strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量/ B.药库包装) AS 帐面数量,(A.扣率/ B.药库包装) AS 实盘数量, (A.实际数量 / B.药库包装) AS 数量差,B.药库包装 as 比例系数,a.零售价*B.药库包装 as 售价,(A.单量* B.药库包装) 成本价," & IIf(bln当前库存, "", "d.实际数量 / B.药库包装 as 当前库存 ,") & IIf(bln可用数量占用, "", "(y.可用数量占用/ B.药库包装) as 可用数量占用 ,")
                End Select
            Else
                '取全部单位，包装；数量，售价，成本价取原始值
                strUnitQuantity = "I.计算单位 As 售价单位, B.门诊单位, B.住院单位, B.药库单位, A.填写数量 AS 帐面数量, A.扣率 AS 实盘数量, A.实际数量 AS 数量差, " & _
                            " '1' As 比例系数售价, B.门诊包装 As 比例系数门诊, B.住院包装 as 比例系数住院, B.药库包装 as 比例系数药库, a.零售价 as 售价, A.单量 成本价, " & IIf(bln当前库存, "", "d.实际数量  as 当前库存 ,") & IIf(bln可用数量占用, "", "y.可用数量占用 as 可用数量占用 ,")
            End If
            
            gstrSQL = "Select *" _
                    & " From " _
                    & "     (SELECT DISTINCT a.药品id,A.序号,a.入出系数,'[' || I.编码 || ']' As 药品编码, I.名称 As 通用名, N.名称 As 商品名," _
                    & "             B.药品来源,B.基本药物,I.规格,A.产地,A.原产地,Nvl(A.库房货位,C.库房货位) As 库房货位,A.批号,a.效期,a.批次," & strUnitQuantity _
                    & "             A.零售金额 as 金额差,A.差价 as 差价差, " _
                    & "             a.摘要,填制人,填制日期,修改人,修改日期,审核人,审核日期,a.频次 as 盘点时间,a.成本价 as 库存金额,a.成本金额 as 库存差价,Nvl(b.加成率,0) as 加成率,I.是否变价,b.药房分批 as 药房分批核算,A.填写数量,A.批准文号,Nvl(A.发药方式,0) As 新批次, " _
                    & " Nvl(I.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间,NVL(B.最大效期,0) 最大效期,a.盘点模式 " _
                    & "      From (Select a.库房id,a.药品id,A.序号,a.入出系数,A.产地,A.原产地,A.库房货位,A.批号,a.效期,a.批次,A.填写数量,A.扣率,A.实际数量,a.零售价,A.单量,A.零售金额,A.差价,a.摘要,填制人,填制日期,修改人,修改日期,审核人,审核日期,a.频次,a.成本价,a.成本金额,A.批准文号,A.发药方式,A.用法 盘点模式 " _
                    & "            From 药品收发记录 A" _
                    & "            Where A.记录状态 =[2] AND A.单据 =12 AND A.No = [1]) A," _
                    & IIf(bln可用数量占用, "", "(Select sum(y.实际数量) 可用数量占用 ,y.库房id,y.药品id,y.批次 From 药品收发记录 Y Where y.入出系数 = -1 And y.审核日期 is null and y.填制日期 > (sysdate - 30)  Group By y.库房id, y.药品id, y.批次) Y,") _
                    & "           药品规格 b,收费项目目录 I ,收费项目别名 n,药品储备限额 C,药品库存 D" _
                    & "      Where A.药品id = B.药品id And A.药品id = I.id" _
                    & "            And A.药品id=n.收费细目id(+) And n.性质(+)=3 " _
                    & IIf(bln可用数量占用, "", "And a.药品id = y.药品id(+) and a.库房id = y.库房id(+) and nvl(a.批次,0) =  nvl(y.批次(+),0)") _
                    & "            And A.药品ID=C.药品ID(+) And A.库房ID=C.库房ID(+) And a.药品id = d.药品id(+) and a.库房id = d.库房id(+) and nvl(a.批次,0) =  nvl(d.批次(+),0) and d.性质(+) = 1)" _
                    & " ORDER BY " & strSqlOrder
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            
            Txt填制人 = rsInitCard!填制人
            Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
            
            Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
            Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
            
            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            txtCheckDate.Caption = rsInitCard!盘点时间
            
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            intRow = 0
            With vsfBill
                .Redraw = flexRDNone
                Do While Not rsInitCard.EOF
                    If rsInitCard.AbsolutePosition = 1 Then mint盘点模式 = IIf(IsNull(rsInitCard!盘点模式), 0, rsInitCard!盘点模式)
                    
                    intRow = intRow + 1
                    'intRow = rsInitCard!序号
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = rsInitCard!通用名
                    Else
                        str药名 = IIf(IsNull(rsInitCard!商品名), rsInitCard!通用名, rsInitCard!商品名)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol药品编码和名称) = rsInitCard!药品编码 & str药名
                    .TextMatrix(intRow, mconIntCol药品编码) = rsInitCard!药品编码
                    .TextMatrix(intRow, mconIntCol药品名称) = str药名
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
                    Else
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsInitCard!商品名), "", rsInitCard!商品名)
                    
                    .TextMatrix(intRow, mconIntCol来源) = zlStr.Nvl(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = zlStr.Nvl(rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsInitCard!库房货位), "", rsInitCard!库房货位)
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx(rsInitCard!加成率, mintMoneyDigit, , True) / 100 & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol新批次) = IIf(IsNull(rsInitCard!新批次), "0", rsInitCard!新批次)
                    
                    If mlng药品ID > 0 Then
                        If Val(.TextMatrix(intRow, 0)) = mlng药品ID And Val(.TextMatrix(intRow, mconIntCol批次)) = mlng批次 Then lng药品定位 = intRow
                    End If
                    
                    If rsInitCard!实盘数量 = 0 Then
                        intNumberDigit = 5
                        intNumberDigit0 = 5
                    Else
                        intNumberDigit = mintNumberDigit
                        intNumberDigit0 = mintNumberDigit0
                    End If
                    .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(rsInitCard!帐面数量, intNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(rsInitCard!实盘数量, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!售价, mintPriceDigit, , True)
                    .TextMatrix(intRow, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol实盘数量)) * Val(.TextMatrix(intRow, mconIntCol售价)), mintMoneyDigit, , True)
                    
                    If mintUnit > 0 Then
                        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!成本价, 0), mintCostDigit, , True)
                    Else
                        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!成本价, 0), mintCostDigit0, , True)
                    End If
                    
                    If mintUnit > 0 Then
                        .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                        .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                        .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx(rsInitCard!数量差, intNumberDigit, , True)
                        
                        If Not .colHidden(mconintCol当前库存) Then .TextMatrix(intRow, mconintCol当前库存) = zlStr.FormatEx(Val(zlStr.Nvl(rsInitCard!当前库存, 0)), intNumberDigit, , True) & rsInitCard!单位
                        If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(intRow, mconintCol可用数量占用) = zlStr.FormatEx(Val(zlStr.Nvl(rsInitCard!可用数量占用, 0)), intNumberDigit, , True) & rsInitCard!单位
                    Else
                        Select Case mint大单位
'                            Case mconint售价单位
'                                .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsintcard!售价单位
'                                .TextMatrix(intRow, mconIntCol盘点数量单位大) = rsintcard!售价单位
'                                .TextMatrix(intRow, mconIntCol比例系数大) = rsInitCard!比例系数售价
'                                .TextMatrix(intRow, mconintCol大包装帐面数量) =Str.FormatEx(rsInitCard!帐面数量, mintNumberDigit)
'                                .TextMatrix(intRow, mconintCol大包装实盘数量) =Str.FormatEx(rsInitCard!实盘数量, mintNumberDigit)
                            Case mconint门诊单位
                                .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsInitCard!门诊单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位大) = rsInitCard!门诊单位
                                .TextMatrix(intRow, mconIntCol比例系数大) = rsInitCard!比例系数门诊
                                .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(rsInitCard!帐面数量 / rsInitCard!比例系数门诊), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol大包装实盘数量) = zlStr.FormatEx(Int(rsInitCard!实盘数量 / rsInitCard!比例系数门诊), intNumberDigit0, , True)
                            Case mconint住院单位
                                .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsInitCard!住院单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位大) = rsInitCard!住院单位
                                .TextMatrix(intRow, mconIntCol比例系数大) = rsInitCard!比例系数住院
                                .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(rsInitCard!帐面数量 / rsInitCard!比例系数住院), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol大包装实盘数量) = zlStr.FormatEx(Int(rsInitCard!实盘数量 / rsInitCard!比例系数住院), intNumberDigit0, , True)
                            Case mconint药库单位
                                .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsInitCard!药库单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位大) = rsInitCard!药库单位
                                .TextMatrix(intRow, mconIntCol比例系数大) = rsInitCard!比例系数药库
                                .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(rsInitCard!帐面数量 / rsInitCard!比例系数药库), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol大包装实盘数量) = zlStr.FormatEx(Int(rsInitCard!实盘数量 / rsInitCard!比例系数药库), intNumberDigit0, , True)
                        End Select
                        
                        Select Case mint小单位
                            Case mconint售价单位
                                .TextMatrix(intRow, mconIntCol帐面数量单位小) = rsInitCard!售价单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位小) = rsInitCard!售价单位
                                .TextMatrix(intRow, mconIntCol比例系数小) = rsInitCard!比例系数售价

                                .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx(Val(rsInitCard!帐面数量) - Val(.TextMatrix(intRow, mconintCol大包装帐面数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大)), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol小包装实盘数量) = zlStr.FormatEx(Val(rsInitCard!实盘数量) - Val(.TextMatrix(intRow, mconintCol大包装实盘数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大)), intNumberDigit0, , True)
                                
                                If Not .colHidden(mconintCol当前库存) Then
                                    dbl当前库存大 = Fix(zlStr.Nvl(rsInitCard!当前库存, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsInitCard!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl当前库存小 = Abs(Val(zlStr.Nvl(rsInitCard!当前库存, 0)) - dbl当前库存大 * Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol当前库存) = .TextMatrix(intRow, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, intNumberDigit0, , True) & rsInitCard!售价单位
                                End If
                                If Not .colHidden(mconintCol可用数量占用) Then
                                    dbl数量占用大 = Fix(zlStr.Nvl(rsInitCard!可用数量占用, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsInitCard!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl数量占用小 = Abs(Val(zlStr.Nvl(rsInitCard!可用数量占用, 0)) - dbl数量占用大 * Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol可用数量占用) = .TextMatrix(intRow, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, intNumberDigit0, , True) & rsInitCard!售价单位
                                End If
                                
                                .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx(rsInitCard!数量差, intNumberDigit, , True)
                                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!售价 * rsInitCard!比例系数售价, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!成本价, 0) * rsInitCard!比例系数售价, mintCostDigit0, , True)
                                .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsInitCard!售价单位
                            Case mconint门诊单位
                                .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(rsInitCard!帐面数量 / rsInitCard!比例系数门诊, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(rsInitCard!实盘数量 / rsInitCard!比例系数门诊, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsInitCard!门诊单位
                                
                                If Not .colHidden(mconintCol当前库存) Then
                                    dbl当前库存大 = Fix(zlStr.Nvl(rsInitCard!当前库存, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsInitCard!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl当前库存小 = Abs((Val(zlStr.Nvl(rsInitCard!当前库存, 0)) - dbl当前库存大 * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / rsInitCard!比例系数门诊)
                                    .TextMatrix(intRow, mconintCol当前库存) = .TextMatrix(intRow, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, intNumberDigit0, , True) & rsInitCard!门诊单位
                                End If
                                If Not .colHidden(mconintCol可用数量占用) Then
                                    dbl数量占用大 = Fix(zlStr.Nvl(rsInitCard!可用数量占用, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsInitCard!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl数量占用小 = Abs((Val(zlStr.Nvl(rsInitCard!可用数量占用, 0)) - dbl数量占用大 * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / rsInitCard!比例系数门诊)
                                    .TextMatrix(intRow, mconintCol可用数量占用) = .TextMatrix(intRow, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, intNumberDigit0, , True) & rsInitCard!门诊单位
                                End If
                                
                                .TextMatrix(intRow, mconIntCol帐面数量单位小) = rsInitCard!门诊单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位小) = rsInitCard!门诊单位
                                .TextMatrix(intRow, mconIntCol比例系数小) = rsInitCard!比例系数门诊
                                
                                .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsInitCard!帐面数量) - Val(.TextMatrix(intRow, mconintCol大包装帐面数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol小包装实盘数量) = zlStr.FormatEx((Val(rsInitCard!实盘数量) - Val(.TextMatrix(intRow, mconintCol大包装实盘数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), intNumberDigit0, , True)

                                .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx(rsInitCard!数量差 / rsInitCard!比例系数门诊, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!售价 * rsInitCard!比例系数门诊, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!成本价, 0) * rsInitCard!比例系数门诊, mintCostDigit0, , True)
                            Case mconint住院单位
                                .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(rsInitCard!帐面数量 / rsInitCard!比例系数住院, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(rsInitCard!实盘数量 / rsInitCard!比例系数住院, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsInitCard!住院单位
                                
                                If Not .colHidden(mconintCol当前库存) Then
                                    dbl当前库存大 = Fix(zlStr.Nvl(rsInitCard!当前库存, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsInitCard!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl当前库存小 = Abs((Val(zlStr.Nvl(rsInitCard!当前库存, 0)) - dbl当前库存大 * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / rsInitCard!比例系数住院)
                                    .TextMatrix(intRow, mconintCol当前库存) = .TextMatrix(intRow, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, intNumberDigit0, , True) & rsInitCard!住院单位
                                End If
                                If Not .colHidden(mconintCol可用数量占用) Then
                                    dbl数量占用大 = Fix(zlStr.Nvl(rsInitCard!可用数量占用, 0) / Val(.TextMatrix(intRow, mconIntCol比例系数大)))
                                    .TextMatrix(intRow, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsInitCard!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, intNumberDigit0, , True) & .TextMatrix(intRow, mconIntCol帐面数量单位大)
                                    dbl数量占用小 = Abs((Val(zlStr.Nvl(rsInitCard!可用数量占用, 0)) - dbl数量占用大 * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / rsInitCard!比例系数住院)
                                    .TextMatrix(intRow, mconintCol可用数量占用) = .TextMatrix(intRow, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, intNumberDigit0, , True) & rsInitCard!住院单位
                                End If
                                
                                .TextMatrix(intRow, mconIntCol帐面数量单位小) = rsInitCard!住院单位
                                .TextMatrix(intRow, mconIntCol实盘数量单位小) = rsInitCard!住院单位
                                .TextMatrix(intRow, mconIntCol比例系数小) = rsInitCard!比例系数住院
                                
                                .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsInitCard!帐面数量) - Val(.TextMatrix(intRow, mconintCol大包装帐面数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol小包装实盘数量) = zlStr.FormatEx((Val(rsInitCard!实盘数量) - Val(.TextMatrix(intRow, mconintCol大包装实盘数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数大))) / Val(.TextMatrix(intRow, mconIntCol比例系数小)), intNumberDigit0, , True)
                                
                                .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx(rsInitCard!数量差 / rsInitCard!比例系数住院, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!售价 * rsInitCard!比例系数住院, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!成本价, 0) * rsInitCard!比例系数住院, mintCostDigit0, , True)
                            Case mconint药库单位
                                .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(rsInitCard!帐面数量 / rsInitCard!比例系数药库, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(rsInitCard!实盘数量 / rsInitCard!比例系数药库, intNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsInitCard!药库单位
                                
                                If Not .colHidden(mconintCol当前库存) Then .TextMatrix(intRow, mconintCol当前库存) = zlStr.FormatEx(Val(zlStr.Nvl(rsInitCard!当前库存, 0)) / rsInitCard!比例系数药库, intNumberDigit0, , True) & rsInitCard!药库单位
                                If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(intRow, mconintCol可用数量占用) = zlStr.FormatEx(Val(zlStr.Nvl(rsInitCard!可用数量占用, 0)) / rsInitCard!比例系数药库, intNumberDigit0, , True) & rsInitCard!药库单位
'                                .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsintcard!药库单位
'                                .TextMatrix(intRow, mconIntCol盘点数量单位大) = rsintcard!药库单位
'                                .TextMatrix(intRow, mconIntCol比例系数大) = rsInitCard!比例系数药库
'                                .TextMatrix(intRow, mconintCol大包装帐面数量) =Str.FormatEx(Int(rsInitCard!帐面数量 / rsInitCard!比例系数药库), mintNumberDigit)
'                                .TextMatrix(intRow, mconintCol大包装实盘数量) =Str.FormatEx(Int(rsInitCard!实盘数量 / rsInitCard!比例系数药库), mintNumberDigit)
                        End Select
                    End If
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        .Cell(flexcpFontBold, intRow, mconintCol当前库存, intRow, mconintCol当前库存) = False
                        If zlStr.Nvl(rsInitCard!当前库存, 0) <> zlStr.Nvl(rsInitCard!帐面数量, 0) Then .Cell(flexcpFontBold, intRow, mconintCol当前库存, intRow, mconintCol当前库存) = True
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        If zlStr.Nvl(rsInitCard!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, intRow, mconintCol可用数量占用, intRow, mconintCol可用数量占用) = True
                    End If
                    
                    If rsInitCard!实盘数量 > rsInitCard!帐面数量 Then
                        .TextMatrix(intRow, mconintCol标志) = "盈"
                    ElseIf rsInitCard!实盘数量 < rsInitCard!帐面数量 Then
                        .TextMatrix(intRow, mconintCol标志) = "亏"
                    Else
                        .TextMatrix(intRow, mconintCol标志) = "平"
                    End If
                    
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
                    '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
                    '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
                    If Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol售价)) = Val(.TextMatrix(intRow, mconintCol成本价))) Then
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                    .TextMatrix(intRow, mconIntCol实际差价) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!库存差价, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconIntCol实际金额) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!库存金额, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconintCol金额差) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!金额差, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconintCol差价差) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!差价差, 0), intMoneyBit, , True)
                    '保持与主界面金额差和差价差算法一致
                    dbl金额差 = Val(.TextMatrix(intRow, mconintCol金额差)) * rsInitCard!入出系数 * IIf(mint记录状态 = 1, 1, IIf(mint记录状态 Mod 3 = 0, 1, -1))
                    dbl差价差 = Val(.TextMatrix(intRow, mconintCol差价差)) * rsInitCard!入出系数 * IIf(mint记录状态 = 1, 1, IIf(mint记录状态 Mod 3 = 0, 1, -1))
                    
                    '.TextMatrix(intRow, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mintMoneyDigit)
                    '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                    .TextMatrix(intRow, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsInitCard!库存金额, 0) + dbl金额差) - (zlStr.Nvl(rsInitCard!库存差价, 0) + dbl差价差), mintMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol金额差)) - Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(intRow, mconintCol库存数量) = zlStr.Nvl(rsInitCard!填写数量, 0)
                    
                    '实盘数量为0与库存数量比较判断盈亏
                    If Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 And Val(.TextMatrix(intRow, mconintCol库存数量)) > 0 Then
                        .TextMatrix(intRow, mconintCol标志) = "亏"
                    End If
                    
                    
                    '设置分批属性
                    Call Get药品分批属性(intRow)
                                        
                    .Row = intRow
                    
                    '盘亏盘盈行用颜色区分
                    Call SetStocktakingColor(vsfBill, intRow)
                   
                    '如果是停用药品，该行粗体显示
                    If Format(rsInitCard!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                    End If
                    
                    .RowData(intRow) = Val(IIf(IsNull(rsInitCard!最大效期), 0, rsInitCard!最大效期))
                    
                    rsInitCard.MoveNext
                Loop
                
                If mintUnit > 0 Then
                    .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
                Else
                    .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
                    .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
                End If
                
                Call SetSortCode
                
                .Redraw = flexRDDirect
            End With
            rsInitCard.Close
    End Select
    mfrmMain.MousePointer = vbDefault
    
    Call vsfColHidden '中药类库房才显示"原产地"列
    Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
    Call 显示合计金额
    
    mblnChange = False
    If lng药品定位 > 0 Then
        vsfBill.Row = lng药品定位
        vsfBill.TopRow = lng药品定位
    End If
    
    mblnLoadData = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfColHidden()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    '只有中药类库房才显示"原产地"列
    str库房性质 = ""
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", txtStock.Tag)
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
    vsfBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'初始化编辑控件
Private Sub initGrid()
    Dim i As Integer
    
    With vsfBill
        .Redraw = flexRDNone
        .Cols = mconIntColS
        .Editable = flexEDNone
        .RowHeightMax = 315
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "生产商"
        .TextMatrix(0, mconIntCol原产地) = "原产地"
        .TextMatrix(0, mconIntCol库房货位) = "库房货位"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        
        .TextMatrix(0, mconIntCol比例系数大) = "比例系数大"
        .TextMatrix(0, mconIntCol比例系数小) = "比例系数小"
        
        .TextMatrix(0, mconIntcol加成率) = "加成率"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        
        .TextMatrix(0, mconintCol帐面数量) = "帐面数量"
        
        .TextMatrix(0, mconintCol大包装帐面数量) = "账面数量(大)"
        .TextMatrix(0, mconIntCol帐面数量单位大) = "单位"
        
        .TextMatrix(0, mconintCol小包装帐面数量) = "账面数量(小)"
        .TextMatrix(0, mconIntCol帐面数量单位小) = "单位"
        
        .TextMatrix(0, mconintCol实盘数量) = "实盘数量"
                
        .TextMatrix(0, mconintCol大包装实盘数量) = "实盘数量(大)"
        .TextMatrix(0, mconIntCol实盘数量单位大) = "单位"
        
        .TextMatrix(0, mconintCol小包装实盘数量) = "实盘数量(小)"
        .TextMatrix(0, mconIntCol实盘数量单位小) = "单位"
        
        .TextMatrix(0, mconintCol合计) = "合计"
        
        .TextMatrix(0, mconintCol当前库存) = "当前库存"
        .TextMatrix(0, mconintCol可用数量占用) = "可用数量占用"
        
        .TextMatrix(0, mconintCol标志) = "标志"
        .TextMatrix(0, mconintCol数量差) = "数量差"
        .TextMatrix(0, mconintCol成本价) = "成本价"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconintCol金额差) = "金额差"
        .TextMatrix(0, mconintCol差价差) = "差价差"
        .TextMatrix(0, mconintCol盘点金额) = "盘点金额"
        .TextMatrix(0, mconintCol盘点成本金额) = "盘点成本金额"
        .TextMatrix(0, mconintCol盘点成本金额差) = "盘点成本金额差"
        .TextMatrix(0, mconintCol库存数量) = "库存数量"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconIntCol新批次) = "新批次"
        .TextMatrix(0, mconIntCol排序编码) = "排序编码"
        .TextMatrix(0, mconIntCol分批属性) = "分批属性"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        mblnChange = False '加载表格的时候mblnChange应该为false
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol来源) = 900
        .colHidden(mconIntCol来源) = True '默认不显示
        .ColWidth(mconIntCol基本药物) = 900
        .colHidden(mconIntCol基本药物) = True '默认不显示
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol可用数量) = 0
        
        .ColWidth(mconIntCol比例系数) = 0
        
        .ColWidth(mconIntCol比例系数大) = 0
        .ColWidth(mconIntCol比例系数小) = 0
        
        .ColWidth(mconIntcol加成率) = 0
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol原产地) = 0
        .ColWidth(mconIntCol库房货位) = 2000
        .colHidden(mconIntCol库房货位) = True '默认不显示
        .ColWidth(mconIntCol单位) = IIf(mintUnit = 0, 0, 600)
        
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .colHidden(mconIntCol批准文号) = True '默认不显示
        
        .ColWidth(mconintCol帐面数量) = IIf(mintUnit = 0, 0, 1200)
        
        .ColWidth(mconintCol大包装帐面数量) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconIntCol帐面数量单位大) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintCol小包装帐面数量) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconIntCol帐面数量单位小) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintCol实盘数量) = IIf(mintUnit = 0, 0, 1200)
        
        .ColWidth(mconintCol大包装实盘数量) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconIntCol实盘数量单位大) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintCol小包装实盘数量) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconIntCol实盘数量单位小) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintCol合计) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconintCol当前库存) = IIf(mintUnit = 0, 2000, 1000)
        .colHidden(mconintCol当前库存) = False '默认显示
        .ColWidth(mconintCol可用数量占用) = IIf(mintUnit = 0, 2000, 1200)
        .colHidden(mconintCol可用数量占用) = True '默认不显示
        .ColWidth(mconintCol标志) = 500
        .ColWidth(mconintCol数量差) = 800
        .ColWidth(mconintCol成本价) = 900
        .ColWidth(mconIntCol售价) = 900
        .ColWidth(mconintCol金额差) = 900
        .colHidden(mconintCol金额差) = True '默认不显示
        .ColWidth(mconintCol差价差) = 900
        .colHidden(mconintCol差价差) = True '默认不显示
        .ColWidth(mconintCol盘点金额) = 900
        .ColWidth(mconintCol盘点成本金额) = 1400
        .ColWidth(mconintCol盘点成本金额差) = 1500
        .colHidden(mconintCol盘点成本金额差) = True '默认不显示
        .ColWidth(mconintCol库存数量) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        .ColWidth(mconIntCol新批次) = 0
        .ColWidth(mconIntCol排序编码) = 0
        .ColWidth(mconIntCol分批属性) = 0
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconintCol帐面数量) = flexAlignRightCenter
        .ColAlignment(mconintCol大包装帐面数量) = flexAlignRightCenter
        .ColAlignment(mconintCol小包装帐面数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol帐面数量单位大) = flexAlignCenterCenter
        .ColAlignment(mconIntCol帐面数量单位小) = flexAlignCenterCenter
        .ColAlignment(mconintCol实盘数量) = flexAlignRightCenter
        .ColAlignment(mconintCol大包装实盘数量) = flexAlignRightCenter
        .ColAlignment(mconintCol小包装实盘数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol实盘数量单位大) = flexAlignCenterCenter
        .ColAlignment(mconIntCol实盘数量单位小) = flexAlignCenterCenter
        
        .ColAlignment(mconintCol合计) = flexAlignRightCenter
        .ColAlignment(mconintCol当前库存) = flexAlignRightCenter
        .ColAlignment(mconintCol可用数量占用) = flexAlignRightCenter
        .ColAlignment(mconintCol标志) = flexAlignCenterCenter
        .ColAlignment(mconintCol数量差) = flexAlignRightCenter
        .ColAlignment(mconintCol成本价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconintCol金额差) = flexAlignRightCenter
        .ColAlignment(mconintCol差价差) = flexAlignRightCenter
        .ColAlignment(mconintCol盘点金额) = flexAlignRightCenter
        .ColAlignment(mconintCol盘点成本金额) = flexAlignRightCenter
        .ColAlignment(mconintCol盘点成本金额差) = flexAlignRightCenter
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Or mint编辑状态 = 9 Then
            txt摘要.Enabled = True
        Else
            txt摘要.Enabled = False
        End If
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .Redraw = flexRDDirect
    End With
    txt摘要.MaxLength = Sys.FieldsLength("药品收发记录", "摘要")
    
    '恢复个性化设置，但部分列不受影响
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '权限控制的，在个性化恢复后还需要进一步控制
    vsfBill.ColWidth(mconintCol成本价) = IIf(mblnViewCost = True, 900, 0)
    vsfBill.ColWidth(mconintCol差价差) = IIf(mblnViewCost = True, 900, 0)
    vsfBill.ColWidth(mconintCol盘点成本金额) = IIf(mblnViewCost = True, 1400, 0)
    vsfBill.ColWidth(mconintCol盘点成本金额差) = IIf(mblnViewCost = True, 1400, 0)
    
    vsfBill.ColWidth(mconIntCol单位) = IIf(mintUnit = 0, 0, 600)
    vsfBill.ColWidth(mconintCol帐面数量) = IIf(mintUnit = 0, 0, 1200)
    vsfBill.ColWidth(mconintCol大包装帐面数量) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol帐面数量单位大) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintCol小包装帐面数量) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol帐面数量单位小) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintCol实盘数量) = IIf(mintUnit = 0, 0, 1200)
    vsfBill.ColWidth(mconintCol大包装实盘数量) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol实盘数量单位大) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintCol小包装实盘数量) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol实盘数量单位小) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintCol合计) = IIf(mintUnit = 0, 1000, 0)
    vsfBill.ColWidth(mconintCol当前库存) = IIf(mintUnit = 0, 2000, 1000)
    vsfBill.ColWidth(mconintCol可用数量占用) = IIf(mintUnit = 0, 2000, 1200)
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        vsfBill.ColWidth(mconIntCol商品名) = IIf(vsfBill.ColWidth(mconIntCol商品名) = 0, 2000, vsfBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        vsfBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    vsfHidden vsfBill
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品盘点管理", "药品名称显示方式", mintDrugNameShow)
    
    mbln检查变动 = False
    mbln即时保存 = False
    mint盘点模式 = 0
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        vsfBill.SetFocus
        vsfBill.Row = 1
        vsfBill.Col = mconIntCol药名
        If txtCheckDate.Caption = "" Then txtCheckDate.Caption = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
    
End Sub

Private Function SaveCheck() As Boolean
    Dim strNo As String
    Dim str审核人 As String
    
    mblnSave = False
    SaveCheck = False
    
    str审核人 = UserInfo.用户姓名
    strNo = txtNo.Tag
    On Error GoTo errHandle
    
    gstrSQL = "zl_药品盘点_Verify('" & strNo & "','" & str审核人 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
        
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With vsfBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码和名称)
                End If
            End If
        Next
    End With
End Sub
Private Sub cbsDefault()
    vsfBill.FixedCols = 1
End Sub

Private Sub cbsFirst()
    vsfBill.Redraw = flexRDNone
    vsfBill.FixedCols = mconIntCol单位
    vsfBill.Refresh
    vsfBill.Redraw = flexRDDirect
End Sub

Private Sub cbsSecond()
    vsfBill.Redraw = flexRDNone
    vsfBill.FixedCols = mconIntCol效期
    vsfBill.Refresh
    vsfBill.Redraw = flexRDDirect
End Sub



Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Then
        Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象)
    End If
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntCol药名
                .ColComboList(mconIntCol药名) = "..."
        End Select
    End With
End Sub

Private Sub vsfBill_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Dim lngColor As Long
    
    With vsfBill
        If NewRowSel > 0 And NewRowSel <> OldRowSel Then
            If .TextMatrix(NewRowSel, mconintCol标志) = "平" Then
                lngColor = mlngColor_盘平
            ElseIf .TextMatrix(NewRowSel, mconintCol标志) = "盈" Then
                lngColor = mlngColor_盘盈
            ElseIf .TextMatrix(NewRowSel, mconintCol标志) = "亏" Then
                lngColor = mlngColor_盘亏
            End If
            
            .ForeColorSel = lngColor
        End If
    End With
End Sub

Private Sub vsfBill_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    With vsfBill
        If Col = mconIntCol药名 Then
            .Col = mconIntCol排序编码
            .Sort = Order
        End If
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) = "" And .rows <> 2 Then
                .RemoveItem lngRow
                .rows = .rows + 1
                .TextMatrix(.rows - 1, mconIntCol行号) = .rows - 1
                Exit For
            End If
        Next
    End With
    
    Call RefreshListSN
End Sub

Private Sub vsfBill_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    If Button = 1 Then
        If y <= vsfBill.RowHeight(0) Then '当点击列头时，从列头开始重新查询
            mlngFindCurrRow = 1
            If Not mrsFindName Is Nothing Then
                mrsFindName.MoveFirst
            End If
        End If
    End If
End Sub


Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    
    intOldRow = vsfBill.Row
    With vsfBill
        Select Case Col
        Case mconIntCol药名
            If mblnNotTrigger <> True Then
                mblnNotTrigger = True
                
                If grsMaster.State = adStateClosed Or mstrLast盘点时间 <> txtCheckDate.Caption Then
                    mstrLast盘点时间 = txtCheckDate.Caption
                    Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象)
                End If
                Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , txtStock.Tag, txtStock.Tag, , 0, False, True, True, IIf(mbln盘停用药品, 1, 0))
                If RecReturn.RecordCount > 0 Then
                    Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
                End If
                
                mblnNotTrigger = False
            Else
                Exit Sub
            End If
        
            '让"Frm药品选择器"中的代码先执行完
            DoEvents
                            
            If RecReturn.RecordCount > 0 Then
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intCurRow = .Row
                    Call SetPhiscRows(RecReturn!药品id, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), Val(RecReturn!成本价), IIf(mintUnit > 0, Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), 0), _
                            IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号))
                    
                    vsfBill_MoveNextCell Row, Col
                    
                    If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End If
        Case mconIntCol产地
            vRect = zlControl.GetControlRect(vsfBill.hWnd)
            dblLeft = vRect.Left + vsfBill.CellLeft
            dblTop = vRect.Top + vsfBill.CellTop
            
            gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
            Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
            True, dblLeft, dblTop, 300, blnCancel, False, True, gstrNodeNo)
            
            If rsProvider Is Nothing Then
                Exit Sub
            End If
            If Not rsProvider.EOF Then
                .TextMatrix(.Row, mconIntCol产地) = rsProvider!名称
            End If
        Case mconIntCol原产地
            vRect = zlControl.GetControlRect(vsfBill.hWnd)
            dblLeft = vRect.Left + vsfBill.CellLeft
            dblTop = vRect.Top + vsfBill.CellTop
            
            gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
            Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "原产地", False, "", "", False, False, _
            True, dblLeft, dblTop, 300, blnCancel, False, True, gstrNodeNo)
            
            If rsProvider Is Nothing Then
                Exit Sub
            End If
            If Not rsProvider.EOF Then
                .TextMatrix(.Row, mconIntCol原产地) = rsProvider!名称
            End If
        Case mconintCol当前库存
            With vsfBill
                If mintUnit > 0 Then '大于0时为默认包装
                    frm库存变动.ShowME 1, Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), txtCheckDate.Caption, Me, .TextMatrix(.Row, mconIntCol单位), Val(.TextMatrix(.Row, mconIntCol比例系数)), mintNumberDigit
                Else
                    frm库存变动.ShowME 1, Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), txtCheckDate.Caption, Me, .TextMatrix(.Row, mconIntCol帐面数量单位大), Val(.TextMatrix(.Row, mconIntCol比例系数大)), .TextMatrix(.Row, mconIntCol帐面数量单位小), Val(.TextMatrix(.Row, mconIntCol比例系数小)), mintNumberDigit0
                End If
            End With
        Case mconintCol可用数量占用
            With vsfBill
                If mintUnit > 0 Then '大于0时为默认包装
                    frm库存变动.ShowME 2, Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), txtCheckDate.Caption, Me, .TextMatrix(.Row, mconIntCol单位), Val(.TextMatrix(.Row, mconIntCol比例系数)), mintNumberDigit
                Else
                    frm库存变动.ShowME 2, Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), txtCheckDate.Caption, Me, .TextMatrix(.Row, mconIntCol帐面数量单位大), Val(.TextMatrix(.Row, mconIntCol比例系数大)), .TextMatrix(.Row, mconIntCol帐面数量单位小), Val(.TextMatrix(.Row, mconIntCol比例系数小)), mintNumberDigit0
                End If
            End With
        End Select
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：用来检查列表中已有药品与新选择的药品是否重复和时价药品是否有库存

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim rs建档时间 As ADODB.Recordset
    Dim str库存 As String
    Dim strSQL As String
    Dim strDub As String    '重复药品
    Dim str重复药名 As String
    Dim strNotPrice As String  '无价格药品
    Dim strNotPrice药名 As String   '用来记录重复选择了的药品名称
    Dim strPrice药名 As String
    Dim rsDetail As ADODB.Recordset
    Dim str盘点时间 As String
    Dim str盘点时间后药品 As String       '纪录在盘点时间后建立的药品
    Dim strSql盘点 As String   '过滤盘点时间后建立的药品
    
    rsTemp.MoveFirst
    str盘点时间后药品 = ""
    strSql盘点 = ""
    str批次 = ""
    strTemp = ""
    str盘点时间 = txtCheckDate.Caption
    
    On Error GoTo errHandle
    Do While Not rsTemp.EOF
        str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            If Val(str批次) <> -1 Then strTemp = strTemp & rsTemp!药品id & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
        
        gstrSQL = "select 现价 from 收费价目 where 执行日期(+)<=[1] AND NVL(终止日期(+),SYSDATE)>=[1] and 收费细目id=[2]" & _
                GetPriceClassString("")
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "查询现价", CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")), rsTemp!药品id)
        If Not rsDetail.EOF Then
            If IsNull(rsDetail!现价) Then
                strNotPrice = strNotPrice & rsTemp!药品id & "," & rsTemp!通用名 & "|"
            End If
        End If
        
        gstrSQL = "Select a.建档时间 From 收费项目目录 A Where a.Id =[1]"
        Set rs建档时间 = zlDataBase.OpenSQLRecord(gstrSQL, "查询建档时间", rsTemp!药品id)
        If Format(rs建档时间!建档时间, "yyyy-MM-dd HH:mm:ss") > Format(txtCheckDate.Caption, "yyyy-MM-dd HH:mm:ss") Then
            str盘点时间后药品 = str盘点时间后药品 & ";" & "[" & rsTemp!药品编码 & "]" & rsTemp!通用名
            strSql盘点 = strSql盘点 & "药品id<>" & rsTemp!药品id & " and "
        End If
        
        rsTemp.MoveNext
    Loop
           
    If strSql盘点 <> "" Then
        MsgBox Mid(str盘点时间后药品, 2) & vbCrLf & "以上药品为盘点时间后建立，所以不会被添加！", vbInformation, gstrSysName
        rsTemp.Filter = Mid(strSql盘点, 1, Len(strSql盘点) - 4)
    End If
     
    With vsfBill    '把重复的查询出来
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
            End If
        Next
        
        If strInfo <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        If strNotPrice <> "" Then
            strPrice药名 = ""
            For i = 0 To UBound(Split(strNotPrice, "|")) - 1
                strPrice药名 = strPrice药名 & "药品id<>" & Split(Split(strNotPrice, "|")(i), ",")(0) & " and "
                If UBound(Split(strNotPrice药名, ",")) <= 2 Then
                    strNotPrice药名 = strNotPrice药名 & Split(Split(strNotPrice, "|")(i), ",")(1) & ","
                End If
            Next
            If strPrice药名 <> "" Then
                strPrice药名 = Mid(strPrice药名, 1, Len(strPrice药名) - 4)
            End If
        End If
        '判断以什么方式拼接sql
        
        If str重复药名 <> "" And strNotPrice药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & strNotPrice药名 & "在本次盘点时间时无售价信息！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub & " and " & strPrice药名
        End If
        If str重复药名 <> "" And strNotPrice药名 = "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If str重复药名 = "" And strNotPrice药名 <> "" Then
            MsgBox strNotPrice药名 & "在本次盘点时间时无售价信息！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strPrice药名
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsfBill_EnterCell()
    Dim lng批次  As Long
    Dim bln新批次 As Boolean
    Dim intRow As Integer
        
    With vsfBill
        .Editable = flexEDNone
        
        .FocusRect = flexFocusLight
        
        Select Case .Col
            Case mconIntCol药名
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol药名) = "..."
                    
                End If
                
            Case mconIntCol批号
                .EditMaxLength = mintBatchNoLen
                
                lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
                bln新批次 = (Val(.TextMatrix(.Row, mconIntCol新批次)) = 1 And (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8))
                
                If IIf(lng批次 = -1 Or bln新批次 = True, 1, 0) = 1 Then
                    .Editable = flexEDKbdMouse
                End If
                
            Case mconIntCol产地
                lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
                bln新批次 = (Val(.TextMatrix(.Row, mconIntCol新批次)) = 1 And (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8))
                
                If IIf(lng批次 = -1 Or bln新批次 = True, 1, 0) = 1 Then
                    .EditMaxLength = mlng生产商长度
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol产地) = "..."
                    
                End If
                
            Case mconIntCol原产地
                lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
                bln新批次 = (Val(.TextMatrix(.Row, mconIntCol新批次)) = 1 And (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8))
                
                If IIf(lng批次 = -1 Or bln新批次 = True, 1, 0) = 1 Then
                    .EditMaxLength = mlng原产地长度
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol原产地) = "..."
                    
                End If
                
            Case mconIntCol效期
                .EditMaxLength = 10
                
                lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
                bln新批次 = (Val(.TextMatrix(.Row, mconIntCol新批次)) = 1 And (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8))
                
                If IIf(lng批次 = -1 Or bln新批次 = True, 1, 0) = 1 Then
                    .Editable = flexEDKbdMouse
                    
                End If
                 
                If .TextMatrix(.Row, mconIntCol批号) <> "" And .TextMatrix(.Row, mconIntCol效期) = "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                        
                        If Len(strxq) = 8 Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                If .RowData(.Row) = 0 Then Exit Sub '未设置最大效期则不自动生成效期
                                .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                                If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                                    '换算为有效期
                                    .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                                End If
                                
                                Call CheckLapse(.TextMatrix(.Row, mconIntCol效期)) '检查是否失效
                            End If
                        End If
                        
                    End If
                End If
                
            Case mconintCol实盘数量, mconintCol大包装实盘数量, mconintCol小包装实盘数量
                .EditMaxLength = 16
                If Val(.TextMatrix(.Row, 0)) <> 0 Then
                    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Or mint编辑状态 = 9 Then
                        If (.Col = mconintCol实盘数量 And mintUnit > 0) Or ((.Col = mconintCol大包装实盘数量 Or .Col = mconintCol小包装实盘数量) And mintUnit = 0) Then
                            .Editable = flexEDKbdMouse
                            
                        End If
                    End If
                End If
            Case mconintCol当前库存
                If Not .Cell(flexcpFontBold, .Row, .Col, .Row, .Col) Then Exit Sub
                .Editable = flexEDKbdMouse
                .ColComboList(mconintCol当前库存) = "..."
            Case mconintCol可用数量占用
                If Not .Cell(flexcpFontBold, .Row, .Col, .Row, .Col) Then Exit Sub
                .Editable = flexEDKbdMouse
                .ColComboList(mconintCol可用数量占用) = "..."
            Case mconintCol成本价
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Then
                    If Val(.TextMatrix(.Row, mconintCol帐面数量)) = 0 Then
                       .Editable = flexEDKbdMouse
                       
                    End If
                End If
                
        End Select
        
        
        If mlongCurrRow <> .Row Then
            mlongCurrRow = .Row
            Call 显示合计金额
            Call 提示库存数
        End If
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBill
        If KeyCode = vbKeyDelete Then
            If .rows = 2 Then Exit Sub
            If .TextMatrix(.Row, mconIntCol行号) = "" Then Exit Sub
            If InStr(1, "469", mint编辑状态) <> 0 Then Exit Sub 'mint编辑状态=3 or 5，让其可以删除（主要是删除库存不足的或可用数量不足的）
            
            If MsgBox("是否删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshRowNO(vsfBill, mconIntCol行号, .Row)
                
                If mint编辑状态 = 3 Then mbln检查变动 = True '审核删除了数据，则mbln检查变动=true
            End If
        End If
        
        If txtCode.Visible And KeyCode = vbKeyF3 Then
            Call txtCode_KeyPress(13)
        End If
        
        Select Case .Col
            Case mconIntCol药名
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntCol药名) = ""
                ElseIf .EditText = "" Then
'                    mblnNotTrigger = True
                    If .TextMatrix(.Row, mconIntCol药名) = "" Then
                        txt摘要.SetFocus
                    End If
                End If
            Case mconIntCol产地
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntCol产地) = ""
                End If
            Case mconIntCol原产地
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntCol原产地) = ""
                End If
        End Select
    End With
End Sub

Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strkey As String
    Dim strTmp As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim rsProvider As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    intOldRow = vsfBill.Row
    With vsfBill
        .Redraw = flexRDNone
        
        .EditText = Trim(.EditText)
        strkey = UCase(Trim(.EditText))
        
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        Select Case Col
            Case mconIntCol药名
                strTmp = .TextMatrix(Row, Col)
                If strkey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + vsfBill.Left + vsfBill.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + vsfBill.Top + vsfBill.CellTop + vsfBill.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - vsfBill.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , strkey, sngLeft, sngTop, False, True, True, True, True, 0, mblnNoStock, 0, mbln盘停用药品, mbln忽略服务对象)
                    If grsMaster.State = adStateClosed Or mstrLast盘点时间 <> txtCheckDate.Caption Then
                        mstrLast盘点时间 = txtCheckDate.Caption
                        Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象)
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, txtStock.Tag, txtStock.Tag, , 0, False, True, True, IIf(mbln盘停用药品, 1, 0))
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
                    End If
                    '让"Frm药品多选选择器"中的代码先执行完
                    DoEvents
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            Call SetPhiscRows(RecReturn!药品id, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), Val(RecReturn!成本价), IIf(mintUnit > 0, Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), 0), IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号))
                            
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    End If

                    Call 提示库存数
                End If
            Case mconIntCol产地
            
                vRect = zlControl.GetControlRect(vsfBill.hWnd)
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top + vsfBill.CellTop
                
                gstrSQL = "Select 编码 as id,简码,名称 From 药品生产商 " _
                            & "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) Order By 编码"
                
                Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & UCase(strkey) & "%", UCase(strkey) & "%", gstrNodeNo)
                
                If rsProvider Is Nothing Then
                    .EditText = ""
                    .TextMatrix(.Row, .Col) = ""
                    Exit Sub
                End If
                If Not rsProvider.EOF Then
                    .TextMatrix(.Row, mconIntCol产地) = rsProvider!名称
                    .EditText = rsProvider!名称
                End If
            Case mconIntCol原产地
                vRect = zlControl.GetControlRect(vsfBill.hWnd)
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top + vsfBill.CellTop
                
                gstrSQL = "Select 编码 as id,简码,名称 From 药品生产商 " _
                            & "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) Order By 编码"
                
                Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "原产地", False, "", "", False, False, _
                True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & UCase(strkey) & "%", UCase(strkey) & "%", gstrNodeNo)
                
                If rsProvider Is Nothing Then
                    .EditText = ""
                    .TextMatrix(.Row, .Col) = ""
                    Exit Sub
                End If
                If Not rsProvider.EOF Then
                    .TextMatrix(.Row, mconIntCol原产地) = rsProvider!名称
                    .EditText = rsProvider!名称
                End If
        End Select
        
        vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
    End If
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        mblnKeyPressReturn = True
    Else
        mblnKeyPressReturn = False
    End If
    
    With vsfBill
        Select Case Col
            Case mconintCol实盘数量, mconintCol大包装实盘数量, mconintCol小包装实盘数量
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     '只能存在一个小数点
                        KeyAscii = 0
                    End If
                End If
                
                strkey = .EditText
                If strkey = "" Then
                    strkey = .TextMatrix(.Row, .Col)
                End If
                Select Case .Col
                    Case mconintCol实盘数量
                        intDigit = mintNumberDigit
                    Case mconintCol大包装实盘数量
                        intDigit = mintNumberDigit1
                    Case mconintCol小包装实盘数量
                        intDigit = mintNumberDigit0
                End Select
                
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Case mconIntCol效期
                If InStr("1234567890-" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case mconintCol成本价
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= mintCostDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
        End Select
    End With
End Sub

Private Sub vsfBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With vsfBill
            If .Col = mconIntCol药名 Then
                If .Row < 1 Then Exit Sub
                mobjPopup.ShowPopup
            End If
        End With
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        If Trim(txtCode.Text) = "" Then Exit Sub
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long
    
    '查找药品
    On Error GoTo errHandle
    If strInput <> txtCode.Tag Then
        '表示新的查找
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If
        lngFindRow = vsfBill.FindRow(str药名, mlngFindCurrRow, CLng(mconIntCol药品编码和名称), True, True)
        
        If lngFindRow > 0 Then '查询到数据后就移动下到下一行，继续检查下一行是否有相同的药品
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
                        
            If lngFindRow < vsfBill.rows - 1 Then
                mlngFindCurrRow = lngFindRow + 1
            Else
                mlngFindCurrRow = 1
                mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
            End If
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
            mlngFindCurrRow = 1 '继续从第一行开始比较其他药品
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim lng效期 As Long
    Dim dbl未发药数量 As Double
    Dim dbl比例系数 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim lng药品ID As Long
    Dim str产地 As String, str批号 As String, dbl成本价 As Double
    Dim intRow As Integer
    
    On Error GoTo errHandle
    With vsfBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, 0)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol实盘数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的实盘数量为空了，请检查！", vbInformation, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintCol实盘数量
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintCol实盘数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的实盘数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintCol实盘数量
                        .EditCell
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol金额差)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的金额差大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintCol实盘数量
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintCol数量差)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的数量差大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintCol实盘数量
                        .EditCell
                        Exit Function
                    End If
                    
                    '分批药品必须录入产地和批号
                    If Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 And (Val(.TextMatrix(intLop, mconIntCol批次)) = -1 Or Val(.TextMatrix(intLop, mconIntCol新批次)) = 1) And (.TextMatrix(intLop, mconIntCol产地) = "" Or .TextMatrix(intLop, mconIntCol批号) = "") Then
                        MsgBox "第" & intLop & "行的药品是新增批次分批药品,请把它的生产商和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol产地) = "" Then
                            .Col = mconIntCol产地
                        Else
                            .Col = mconIntCol批号
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol批次)) = -1 Then
                        If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                            MsgBox "第" & intLop & "行药品的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                            .SetFocus
                            .Row = intLop
                            .TopRow = intLop
                            .Col = mconIntCol批号
                            .EditCell
                            Exit Function
                        End If
                        
                        '判断是否为效期药品
                        gstrSQL = "Select Nvl(最大效期,0) 效期 From 药品规格 Where 药品ID=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是否为效期药品]", Val(.TextMatrix(intLop, 0)))
                        
                        lng效期 = rsTemp!效期
                        If lng效期 <> 0 Then
                            If Val(.TextMatrix(intLop, mconintCol实盘数量)) <> 0 Then
                                If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Or Trim(.TextMatrix(intLop, mconIntCol效期)) = "" Then
                                    MsgBox "第" & intLop & "行的药品是效期药品,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！", vbInformation, gstrSysName
                                    vsfBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, mconIntCol批号) = "" Then
                                        .Col = mconIntCol批号
                                    Else
                                        .Col = mconIntCol效期
                                    End If
                                    .EditCell
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol新批次)) = 0 Then
                        '零差价管理：检查是否存在不满足零差价的药品
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And (Val(.TextMatrix(intLop, mconIntCol批次)) >= 0 And Val(.TextMatrix(intLop, mconIntCol新批次)) = 0) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                                If CheckPriceAdjust(Val(.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(.TextMatrix(intLop, mconIntCol批次))) = False Then
                                    MsgBox "第" & intLop & "行药品已启用零差价管理，但库存记录中售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                                    .SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        '新增时
                        If .TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                            If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                                '如果是零差价管理，检查界面售价和成本价关系
                                If Val(.TextMatrix(intLop, mconintCol成本价)) <> Val(.TextMatrix(intLop, mconIntCol售价)) Then
                                    MsgBox "第" & intLop & "行药品已启用零差价管理，但盘点界面的售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                                    .SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                                        
                End If
            Next
            
            '检查分批药品新增批次的产地，批号是否重复
            For intLop = 1 To .rows - 1
                If Val(.TextMatrix(intLop, mconIntCol批次)) = -1 Or Val(.TextMatrix(intLop, mconIntCol新批次)) = 1 Then
                    lng药品ID = Val(.TextMatrix(intLop, 0))
                    str产地 = .TextMatrix(intLop, mconIntCol产地)
                    str批号 = .TextMatrix(intLop, mconIntCol批号)
                    dbl成本价 = Val(.TextMatrix(intLop, mconintCol成本价))
                    
                    For intRow = 1 To .rows - 1
                        If intLop <> intRow And _
                            lng药品ID = Val(.TextMatrix(intRow, 0)) And _
                            str产地 = .TextMatrix(intRow, mconIntCol产地) And _
                            str批号 = .TextMatrix(intRow, mconIntCol批号) And _
                            dbl成本价 = Val(.TextMatrix(intRow, mconintCol成本价)) Then
                            
                            MsgBox "第" & intLop & "行的药品(" & Trim(.TextMatrix(intLop, mconIntCol药名)) & ")新增批次的生产商，批号，成本价和第" & intRow & "行重复了！" & vbCrLf & "请重新录入生产商和批号信息！", vbInformation, gstrSysName
                            
                            vsfBill.SetFocus
                            .Row = intLop
                            .TopRow = intLop
                            .Col = mconIntCol批号
                            .EditCell
                            Exit Function
                        End If
                    Next
                End If
                
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function SaveCard() As Boolean
    Dim lng入出类别id As Long
    Dim int入出系数 As Integer
    Dim lng入库类别ID As Integer
    Dim lng出库类别ID As Integer
    
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim str批号 As String
    Dim lng批次ID As Long
    Dim str产地 As String
    Dim str原产地 As String
    Dim dat效期 As String
    Dim dbl帐面数量 As Double
    Dim dbl实盘数量 As Double
    Dim dbl数量差 As Double
    Dim dbl售价 As Double
    Dim dbl成本价 As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim dat填制日期 As String
    Dim str修改人 As String
    Dim dat修改日期 As String
    Dim str盘点时间 As String
    Dim dbl库存金额 As Double
    Dim dbl库存差价 As Double
    Dim rs入出类别 As New Recordset
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim int新批次 As Integer
    Dim arrSql As Variant
    Dim i As Integer
    
    Dim str单据号() As String
    Dim n As Long
    
    Dim intMoneyBit As Integer
    Dim dbl比例系数 As Double
    Dim str库房货位 As String
    
    arrSql = Array()
    SaveCard = False
    On Error GoTo errHandle
    '在外面设置入出类别ID，主要是所有药品都要用他
    gstrSQL = "SELECT b.系数,b.id AS 类别id " _
            & "FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID AND a.单据 = 12 "
    Set rs入出类别 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
    If rs入出类别.EOF Then
        MsgBox "对不起，没有设置药品盘点管理的入出类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng入库类别ID = 0
    lng出库类别ID = 0
    
    rs入出类别.MoveFirst
    Do While Not rs入出类别.EOF
        If rs入出类别!系数 = 1 Then
            lng入库类别ID = rs入出类别!类别id
        Else
            lng出库类别ID = rs入出类别!类别id
        End If
        rs入出类别.MoveNext
    Loop
    rs入出类别.Close
    
    If lng入库类别ID = 0 Then
        MsgBox "对不起，没有设置药品盘点管理的入库类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If lng出库类别ID = 0 Then
        MsgBox "对不起，没有设置药品盘点管理的出库类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Me.MousePointer = vbHourglass
    
    With vsfBill
        chrNo = Trim(txtNo)
        lng库房ID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(29, lng库房ID)
        
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        dat填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        str盘点时间 = txtCheckDate.Caption
        
        If mint盘点模式 = 0 And mint编辑状态 <> 2 Then mint盘点模式 = mint编辑状态   '新增mint编辑状态 = 0
        
        If mint编辑状态 = 2 Or mbln检查变动 = True Or mbln即时保存 Then       '修改
            gstrSQL = "zl_药品盘点_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            str填制人 = Txt填制人
            dat填制日期 = Format(Txt填制日期, "yyyy-mm-dd hh:mm:ss")
            str修改人 = UserInfo.用户姓名
            dat修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                int新批次 = 0
                If Val(.TextMatrix(intRow, mconIntCol批次)) = -1 Or Val(.TextMatrix(intRow, mconIntCol新批次)) = 1 Then
                    int新批次 = 1
                End If
                
                lng药品ID = .TextMatrix(intRow, 0)
                dbl比例系数 = IIf(mintUnit > 0, Val(.TextMatrix(intRow, mconIntCol比例系数)), Val(.TextMatrix(intRow, mconIntCol比例系数小)))
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str原产地 = .TextMatrix(intRow, mconIntCol原产地)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                lng批次ID = IIf(.TextMatrix(intRow, mconIntCol批次) = "", 0, .TextMatrix(intRow, mconIntCol批次))
                dat效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                
                dbl帐面数量 = Val(.TextMatrix(intRow, mconintCol库存数量))
                dbl实盘数量 = zlStr.FormatEx(.TextMatrix(intRow, mconintCol实盘数量) * dbl比例系数, gtype_UserDrugDigits.Digit_数量, , True)

                If Trim(.TextMatrix(intRow, mconintCol标志)) = "平" Then
                    If dbl帐面数量 <> Val(.TextMatrix(intRow, mconintCol帐面数量)) * dbl比例系数 Then
                        '真实库存账面数量和界面账面数量换算后的不一致时(由于精度取舍导致的，可能导致盘点后无法得到预期的实盘数量)
                        '使用真实库存数量来和实盘数量计算数量差
                        dbl数量差 = Val(.TextMatrix(intRow, mconintCol实盘数量)) * dbl比例系数 - dbl帐面数量
                    Else
                        dbl数量差 = 0
                    End If
                    dbl实盘数量 = Val(.TextMatrix(intRow, mconintCol库存数量))
                Else
                    dbl数量差 = zlStr.FormatEx(Abs(.TextMatrix(intRow, mconintCol实盘数量) * dbl比例系数 - Val(.TextMatrix(intRow, mconintCol库存数量))), gtype_UserDrugDigits.Digit_数量, , True)
                End If
                
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                              
                dbl售价 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / dbl比例系数, gtype_UserDrugDigits.Digit_零售价)
                dbl成本价 = zlStr.FormatEx(.TextMatrix(intRow, mconintCol成本价) / dbl比例系数, gtype_UserDrugDigits.Digit_成本价)

                If Val(Split(.TextMatrix(intRow, mconIntcol加成率), "||")(1)) = 0 Or int新批次 = 0 Then
                    '定价药品或不是新增批次取原始售价
                    dbl售价 = Get盘点时刻售价(Split(.TextMatrix(intRow, mconIntcol加成率), "||")(1) = 1, lng药品ID, lng库房ID, lng批次ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                Else
                    '新增批次时价按界面价格换算后保存
                    dbl售价 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / dbl比例系数, gtype_UserDrugDigits.Digit_零售价)
                End If

                If int新批次 = 0 Then
                    '不是新增批次取原始成本价
                    dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                Else
                    If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(lng药品ID) = True Then
                        dbl成本价 = dbl售价
                    Else
                        '新增批次按界面价格换算后保存
                        dbl成本价 = zlStr.FormatEx(.TextMatrix(intRow, mconintCol成本价) / dbl比例系数, gtype_UserDrugDigits.Digit_成本价)
                    End If
                End If
      
                str库房货位 = IIf(Trim(.TextMatrix(intRow, mconIntCol库房货位)) = "", "", .TextMatrix(intRow, mconIntCol库房货位))
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
                '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
                '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
                If int新批次 = 1 Then
                    intMoneyBit = mintMoneyDigit
                ElseIf dbl实盘数量 = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol售价)) = Val(.TextMatrix(intRow, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
        
                dbl金额差 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol金额差)), intMoneyBit, , True)
                dbl差价差 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)
                dbl库存金额 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol实际金额)), intMoneyBit, , True)
                dbl库存差价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol实际差价)), intMoneyBit, , True)
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If dbl帐面数量 <= dbl实盘数量 Then
                    lng入出类别id = lng入库类别ID
                    int入出系数 = 1
                Else
                    lng入出类别id = lng出库类别ID
                    int入出系数 = -1
                End If
                
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lng序号 = intRow
'                Else
'                    lng序号 = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                lng序号 = intRow
                
                gstrSQL = "zl_药品盘点_INSERT('" & chrNo & "'," & lng序号 & "," & lng库房ID & "," & lng批次ID & "," _
                    & lng入出类别id & "," & int入出系数 & "," & lng药品ID & "," & dbl帐面数量 & "," _
                    & dbl实盘数量 & "," & dbl数量差 & "," & dbl售价 & "," & dbl金额差 & "," & dbl差价差 & ",'" _
                    & str填制人 & "',to_date('" & dat填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & str摘要 & "','" & str产地 & "','" & str批号 & "'," & IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & str盘点时间 & "'," & dbl库存金额 & "," & dbl库存差价 & "," & dbl成本价 & ",'" & str批准文号 & "'," & int新批次 & ",'" & str库房货位 & "','" & str原产地 & "'," & IIf(mint盘点模式 = 0, "Null", mint盘点模式) & ",'" _
                    & str修改人 & "'," & IIf(dat修改日期 = "", "Null", "to_date('" & dat修改日期 & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                
            End If
            recSort.MoveNext
        Next
        
        If mint编辑状态 = 5 Then
            If InStr(mstr盘点单号, ",") = 0 Then
                ReDim str单据号(0)
                str单据号(0) = mstr盘点单号
            Else
                str单据号 = Split(mstr盘点单号, ",")
            End If
            
            If mbln删除盘点单 Then
                For n = 0 To UBound(str单据号)
                    gstrSQL = "Zl_药品盘点记录单_DELETE(" & str单据号(n) & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                Next
            Else
                For n = 0 To UBound(str单据号)
                    gstrSQL = "Zl_药品盘点记录单_Update(" & str单据号(n) & ",'" & UserInfo.用户姓名 & "',to_date('" & Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                Next
            End If
        End If
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    Me.MousePointer = vbDefault
    
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub 显示合计金额()
    Dim dbl金额差 As Double
    Dim dbl盘点金额 As Double
    Dim intLop As Integer
    Dim dbl成本金额 As Double
    
    dbl金额差 = 0
    dbl盘点金额 = 0
    dbl成本金额 = 0
    
    With vsfBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl金额差 = dbl金额差 + Val(.TextMatrix(intLop, mconintCol金额差)) * IIf(.TextMatrix(intLop, mconintCol标志) = "亏", -1, 1)
                dbl盘点金额 = dbl盘点金额 + Val(.TextMatrix(intLop, mconintCol盘点金额))
                dbl成本金额 = dbl成本金额 + Val(.TextMatrix(intLop, mconintCol盘点成本金额))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "金额差合计：" & zlStr.FormatEx(dbl金额差, mintMoneyDigit, , True)
    lblPurchasePrice.Width = Pic单据.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200

    lblCheckSum.Caption = "盘点金额合计：" & zlStr.FormatEx(dbl盘点金额, mintMoneyDigit, , True)
    lblCheckSum.Width = Pic单据.TextWidth(lblCheckSum.Caption)
    
    lblCostPrice.Top = lblCheckSum.Top
    lblCostPrice.Left = lblCheckSum.Left + lblCheckSum.Width + 200
    lblCostPrice.Caption = "盘点成本金额合计：" & zlStr.FormatEx(dbl成本金额, mintMoneyDigit, , True)
    lblCostPrice.Width = Pic单据.TextWidth(lblCostPrice.Caption)
End Sub

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    Dim dbl大包装数量 As Double
    Dim dbl小包装数量 As Double
    Dim dbl大包装实际数量 As Double
    Dim dbl小包装实际数量 As Double
    
    On Error GoTo errHandle
    With vsfBill
        If .TextMatrix(.Row, mconIntCol药名) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(vsfBill.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "select Nvl(可用数量,0) 可用数量,nvl(实际数量,0) 实际数量 from 药品库存 " _
                & "where 库房id=[1] " _
                & "  and 药品id=[2] " _
                & "  and 性质=1 " _
                & "  and nvl(批次,0)=[3]"
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", txtStock.Tag, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            If mintUnit > 0 Then
                dbl大包装数量 = rsUseCount!可用数量 / Val(.TextMatrix(.Row, mconIntCol比例系数))
                dbl大包装实际数量 = rsUseCount!实际数量 / Val(.TextMatrix(.Row, mconIntCol比例系数))
                
                .TextMatrix(.Row, mconIntCol可用数量) = dbl大包装数量
            Else
                dbl大包装数量 = Int(rsUseCount!可用数量 / Val(.TextMatrix(.Row, mconIntCol比例系数大)))
                dbl大包装实际数量 = Int(rsUseCount!实际数量 / Val(.TextMatrix(.Row, mconIntCol比例系数大)))
                
                dbl小包装数量 = zlStr.FormatEx((Val(rsUseCount!可用数量) - dbl大包装数量 * Val(.TextMatrix(.Row, mconIntCol比例系数大))) / Val(.TextMatrix(.Row, mconIntCol比例系数小)), mintNumberDigit0, , True)
                dbl小包装实际数量 = zlStr.FormatEx((Val(rsUseCount!实际数量) - dbl大包装实际数量 * Val(.TextMatrix(.Row, mconIntCol比例系数大))) / Val(.TextMatrix(.Row, mconIntCol比例系数小)), mintNumberDigit0, , True)
                
                .TextMatrix(.Row, mconIntCol可用数量) = rsUseCount!可用数量 / Val(.TextMatrix(.Row, mconIntCol比例系数小))
            End If
        End If
        rsUseCount.Close
        
        If mintUnit > 0 Then
            staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(dbl大包装实际数量, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol单位)
        Else
            staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(dbl大包装实际数量, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol帐面数量单位大) & _
                ",[" & zlStr.FormatEx(dbl小包装实际数量, mintNumberDigit0, , True) & "]" & .TextMatrix(.Row, mconIntCol帐面数量单位小)
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme True
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
     OS.OpenIme
End Sub

Private Function SetPhiscRows(ByVal lngID As Long, ByVal lng批次 As Long, ByVal dbl初始成本价 As Double, ByVal dbl比例系数 As Double, ByVal str批准文号 As String) As Boolean
'功能：根据药品ID在盘存表上显示并处理该药品的初始盘存信息
'说明：
'   1.如果是非分批核算药,且已经输入了,则提示并退出。
'   2.如果是分批核算药，则分别处理该药的未处理的各批次库存行。
    Dim i As Integer, lngRow As Long
    Dim rsDetail As ADODB.Recordset
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim bln库房 As Boolean
    Dim dbl成本价 As Double, dbl零售价 As Double, dbl加成率 As Double
    Dim str产地 As String
    Dim lngBatch As Long
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim n As Integer
    Dim rs时价分批 As ADODB.Recordset
    Dim rsDingPrice As ADODB.Recordset
    Dim str药名 As String
    Dim bln盘点入库 As Boolean
    Dim str盘点时间 As String
    Dim dbl盘点金额 As Double

    On Error GoTo errH
    
    str盘点时间 = txtCheckDate.Caption
    
    Set rsDetail = GetPhysicDetail(txtStock.Tag, lngID)
    intRecordCount = rsDetail.RecordCount
    If intRecordCount = 0 Then Exit Function
    
    mstrMsg = ""
    
    '新增批次药品
    If lng批次 <> -1 Then
        rsDetail.MoveFirst
        rsDetail.Find "批次=" & lng批次
        If rsDetail.EOF Then Exit Function
    End If
    
    bln库房 = CheckPartProp(Val(txtStock.Tag))
    With vsfBill
        vsfBill.Redraw = flexRDNone
        intRow = .Row
        .TextMatrix(intRow, 0) = rsDetail!药品id
        
        dbl成本价 = zlStr.Nvl(rsDetail!平均成本价, 0)
        dbl零售价 = IIf(IsNull(rsDetail!售价), 0, rsDetail!售价)
        '处理在盘点后又新增了的药品
        If rsDetail!是否变价 = 0 And IsNull(rsDetail!售价) Then
            gstrSQL = "select 现价 from 收费价目 where 收费细目id=[1] and sysdate between 执行日期 and 终止日期" & _
                    GetPriceClassString("")
            
            Set rsDingPrice = zlDataBase.OpenSQLRecord(gstrSQL, "定价价格", rsDetail!药品id)
            If rsDingPrice.EOF = False Then
                dbl零售价 = rsDingPrice!现价
            End If
        End If
        
        If rsDetail!是否变价 = 1 Then
            dbl零售价 = Get盘点时刻零售价(Val(.TextMatrix(intRow, 0)), Val(txtStock.Tag), lng批次, 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
        End If
        
        '判断有无库存，如果无库存作为新增药品
        If lng批次 = 0 Then
            If CheckNoStock(Val(txtStock.Tag), Val(.TextMatrix(intRow, 0))) = True Then
                '无库存时为盘点入库
                bln盘点入库 = True
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    If rsDetail!是否变价 = 1 Then
                        '零差价管理，时价药品售价要等于成本价
                        dbl零售价 = dbl成本价
                    Else
                        '零差价管理，定价药品成本价要等于售价
                        dbl成本价 = dbl零售价
                    End If
                End If
            End If
        End If
        
        '如果是新增批次时
        If lng批次 = -1 Then
            If rsDetail!是否变价 = 0 Then
                '定价
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    '零差价管理：成本价要等于售价
                    dbl成本价 = dbl零售价
                End If
            Else
                '时价
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    '零差价管理：售价要等于成本价
                    dbl零售价 = dbl成本价
                Else
                    dbl零售价 = Get盘点时刻零售价(Val(.TextMatrix(intRow, 0)), Val(txtStock.Tag), lng批次, 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
            End If
        End If
        
        str产地 = zlStr.Nvl(rsDetail!缺省产地, "")
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = rsDetail!通用名
        Else
            str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = rsDetail!药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
        
        If .Col = mconIntCol药名 Then
            .EditText = .TextMatrix(intRow, mconIntCol药名)
        End If
        
        .TextMatrix(intRow, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
        .TextMatrix(intRow, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
        .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsDetail!产地), "", rsDetail!产地)
        .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsDetail!原产地), "", rsDetail!原产地)
        If .TextMatrix(intRow, mconIntCol产地) = "" Then .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
        
        If mintUnit > 0 Then
            '按常量定义进行格式化
            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
            
            .TextMatrix(intRow, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
            .TextMatrix(intRow, mconIntCol比例系数) = rsDetail!比例系数
            
            If rsDetail!是否变价 = 1 Then
                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Get盘点时刻成本价(rsDetail!药品id, Val(txtStock.Tag), CLng(rsDetail!批次), CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss"))) * dbl比例系数, mintCostDigit, , True)
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    '零差价管理：售价要等于成本价
                    .TextMatrix(intRow, mconIntCol售价) = .TextMatrix(intRow, mconintCol成本价)
                End If
            Else
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    .TextMatrix(intRow, mconintCol成本价) = .TextMatrix(intRow, mconIntCol售价)
                Else
                    .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(IIf(dbl初始成本价 = 0, dbl成本价, dbl初始成本价) * dbl比例系数, mintCostDigit, , True)
                End If
            End If
        Else
            '按常量定义进行格式化
            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
            
            .TextMatrix(intRow, mconIntCol帐面数量单位大) = rsDetail!大包装单位
            .TextMatrix(intRow, mconIntCol帐面数量单位小) = rsDetail!小包装单位
            .TextMatrix(intRow, mconIntCol实盘数量单位大) = rsDetail!大包装单位
            .TextMatrix(intRow, mconIntCol实盘数量单位小) = rsDetail!小包装单位
            
            .TextMatrix(intRow, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
            .TextMatrix(intRow, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
            
            If rsDetail!是否变价 = 1 Then
                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Get盘点时刻成本价(rsDetail!药品id, Val(txtStock.Tag), CLng(rsDetail!批次), CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss"))) * rsDetail!比例系数小, mintCostDigit0, , True)
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    '零差价管理：售价要等于成本价
                    .TextMatrix(intRow, mconIntCol售价) = .TextMatrix(intRow, mconintCol成本价)
                End If
            Else
                If IsPriceAdjustMod(rsDetail!药品id) = True Then
                    .TextMatrix(intRow, mconintCol成本价) = .TextMatrix(intRow, mconIntCol售价)
                Else
                    .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(IIf(dbl初始成本价 = 0, dbl成本价, dbl初始成本价) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
            End If
        End If
            
        .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsDetail!批次), "0", rsDetail!批次)
        If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(intRow, mconIntCol批次)) = 0 Then
            lng批次 = -1
        End If
        
        If lng批次 = -1 Or bln盘点入库 = True Then
            .TextMatrix(intRow, mconIntCol新批次) = 1
            .TextMatrix(intRow, mconIntCol批次) = lng批次
            .TextMatrix(intRow, mconIntCol批号) = ""
            .TextMatrix(intRow, mconIntCol效期) = ""
            .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
            
            .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol帐面数量), mintNumberDigit, , True)
            
            If mintUnit = 0 Then
                .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(0, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx(0, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol大包装实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol大包装帐面数量), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                
                If Not .colHidden(mconintCol当前库存) Then .TextMatrix(intRow, mconintCol当前库存) = zlStr.FormatEx(0, mintNumberDigit0, , True) & rsDetail!大包装单位 & zlStr.FormatEx(0, mintNumberDigit0, , True) & rsDetail!小包装单位
                If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(intRow, mconintCol可用数量占用) = zlStr.FormatEx(0, mintNumberDigit0, , True) & rsDetail!大包装单位 & zlStr.FormatEx(0, mintNumberDigit0, , True) & rsDetail!小包装单位
            End If
            
            .TextMatrix(intRow, mconintCol盘点金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconIntCol实际金额) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintCol库存数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconIntCol实际差价) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            
            If mintUnit <= 0 Then
                .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsDetail!小包装单位
            Else
                If Not .colHidden(mconintCol当前库存) Then .TextMatrix(intRow, mconintCol当前库存) = zlStr.FormatEx(0, mintNumberDigit, , True) & rsDetail!单位
                If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(intRow, mconintCol可用数量占用) = zlStr.FormatEx(0, mintNumberDigit, , True) & rsDetail!单位
            End If
        Else
            .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsDetail!批次), "0", rsDetail!批次)
            .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
            .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                '换算为有效期
                .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
            End If
            
            .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
            
            If mintUnit > 0 Then
                .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol帐面数量), mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0), mintNumberDigit, , True)
                
                If Not .colHidden(mconintCol当前库存) Then .TextMatrix(intRow, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(intRow, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                
                .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
            Else
                .TextMatrix(intRow, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol帐面数量), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0), mintNumberDigit0, , True)
                
                .TextMatrix(intRow, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(rsDetail!帐面数量 / rsDetail!比例系数大), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol大包装实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol大包装帐面数量), mintNumberDigit0, , True)

                .TextMatrix(intRow, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(intRow, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol小包装帐面数量), mintNumberDigit0, , True)
            
                If Not .colHidden(mconintCol当前库存) Then
                    Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                    dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                    .TextMatrix(intRow, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                    dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                    .TextMatrix(intRow, mconintCol当前库存) = .TextMatrix(intRow, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                    dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                    .TextMatrix(intRow, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                    dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                    .TextMatrix(intRow, mconintCol可用数量占用) = .TextMatrix(intRow, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                End If
                
                If mintUnit <= 0 Then
                    .TextMatrix(intRow, mconintCol合计) = .TextMatrix(intRow, mconintCol实盘数量) & rsDetail!小包装单位
                End If
            End If
            
            
            .TextMatrix(intRow, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol实盘数量)) * Val(.TextMatrix(intRow, mconIntCol售价)), mintMoneyDigit, , True)
            .TextMatrix(intRow, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
            .TextMatrix(intRow, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
            .TextMatrix(intRow, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
        End If
        
        If Not .colHidden(mconintCol当前库存) Then
            .Cell(flexcpFontBold, intRow, mconintCol当前库存, intRow, mconintCol当前库存) = False
            If lng批次 <> -1 Then
                If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, intRow, mconintCol当前库存, intRow, mconintCol当前库存) = True
            End If
        End If
        If Not .colHidden(mconintCol可用数量占用) Then
            If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, intRow, mconintCol可用数量占用, intRow, mconintCol可用数量占用) = True
        End If
        
        
        .TextMatrix(intRow, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
        .TextMatrix(intRow, mconintCol标志) = "平"
        
        '实盘数量为0与库存数量比较判断盈亏
        If Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 And Val(.TextMatrix(intRow, mconintCol库存数量)) > 0 Then
            .TextMatrix(intRow, mconintCol标志) = "亏"
        End If
                
        .TextMatrix(intRow, mconintCol数量差) = zlStr.FormatEx("0", mintNumberDigit, , True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
        '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
        '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
        If (Val(.TextMatrix(intRow, mconintCol实盘数量)) = 0 And lng批次 <> -1 And bln盘点入库 = False) Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol售价)) = Val(.TextMatrix(intRow, mconintCol成本价))) Then
            intMoneyBit = mintMaxMoneyBit
        Else
            intMoneyBit = mintMoneyDigit
        End If
        
        '金额差=当前售价*实盘数量-实际金额
        '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
        dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
        .TextMatrix(intRow, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(intRow, mconIntCol实际金额)), intMoneyBit, , True)
        .TextMatrix(intRow, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol售价)) - Val(.TextMatrix(intRow, mconintCol成本价))) * Val(.TextMatrix(intRow, mconintCol实盘数量)) - Val(.TextMatrix(intRow, mconIntCol实际差价)), intMoneyBit, , True)
        
        If .TextMatrix(intRow, mconintCol标志) = "亏" Then
            .TextMatrix(intRow, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol金额差)), intMoneyBit, , True)
            .TextMatrix(intRow, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)
        End If
                
        '.TextMatrix(intRow, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mintMoneyDigit)
        '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
        .TextMatrix(intRow, mconintCol盘点成本金额) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol实际金额)) + Val(.TextMatrix(intRow, mconintCol金额差))) - (Val(.TextMatrix(intRow, mconIntCol实际差价)) + Val(.TextMatrix(intRow, mconintCol差价差))), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol金额差)) - Val(.TextMatrix(intRow, mconintCol差价差)), intMoneyBit, , True)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If mbln盘停用药品 = True Then
            '如果是停用药品，该行粗体显示
            If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
            End If
        End If
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, intRow, mconintCol实盘数量, intRow, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, intRow, mconintCol大包装实盘数量, intRow, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, intRow, mconintCol小包装实盘数量, intRow, mconintCol小包装实盘数量) = True
        End If
        .RowData(intRow) = Val(IIf(IsNull(rsDetail!最大效期), 0, rsDetail!最大效期))
        
        '盘亏盘盈行用颜色区分
        Call SetStocktakingColor(vsfBill, intRow)
                
        '设置分批属性
        Call Get药品分批属性(intRow)
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        vsfBill.Redraw = flexRDDirect
    End With
    rsDetail.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'在一行中插入
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim intOldRows As Integer
    Dim intLop As Integer
    Dim intExchange As Integer
    Dim intCol As Integer
    
    With vsfBill
        blnHaveData = False
        intOldRows = .rows - 1
        .rows = .rows + intRecordCount
        For intLop = intRow + 1 To intRecordCount
            If .TextMatrix(intLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For intExchange = .rows - 1 To intOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(intExchange, intCol) = .TextMatrix(intExchange - intRecordCount, intCol)
                    .TextMatrix(intExchange - intRecordCount, intCol) = ""
                Next
            Next
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'打印单据
Private Sub printbill()
    Dim int单位系数 As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint售价单位
            int单位系数 = 4
        Case mconint门诊单位
            int单位系数 = 2
        Case mconint住院单位
            int单位系数 = 1
        Case mconint药库单位
            int单位系数 = 3
    End Select
    
    strNo = txtNo.Tag
    Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), mint记录状态, int单位系数, 1307, "药品盘点表", strNo)
End Sub

Private Function CheckPartProp(ByVal lng库房ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查库房属性，如果是药库，返回真
    On Error GoTo errHandle
    gstrSQL = "SELECT count(*) " _
            & "From 部门性质说明 " _
            & "WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) " _
            & "  AND 部门id =[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是药库/药房]", lng库房ID)
    
    If rsTemp.Fields(0) > 0 Then
        CheckPartProp = False
    Else
        CheckPartProp = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPhysicBatch(ByVal bln库房 As Boolean, ByVal int药库分批 As Integer, ByVal int药房分批 As Integer) As Boolean
    '返回该药品是否分批的标识
    CheckPhysicBatch = (bln库房 And (int药库分批 = 1)) Or (Not bln库房 And (int药房分批 = 1))
End Function

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Set rsBatchNolen = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-取批号长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPhysic(ByVal lng库房ID As Long, ByVal str盘点属性 As String, _
        ByVal str剂型 As String, Optional ByVal str库房货位 As String = "所有", _
        Optional ByVal bln盘无库存药品 As Boolean = True, _
        Optional ByVal bln汇总盘点单 As Boolean = False, _
        Optional ByVal bln盘点单 As Boolean = False, _
        Optional ByVal bln盘无库存有金额药品 As Boolean = False) As ADODB.Recordset
    '读取出符合条件的药品（同时提出单位与包装系数）
    'bln盘无库存药品=是否将无库存药品也提取出来
    'bln汇总盘点单=是否需要汇总指定盘点时间的盘点单形成盘点表
    'bln盘点单=是否仅针对盘点单产生盘点表，如果为假，说明要将现有库存一并提出来汇总，不在盘点单中的药品的实盘数量以零显示
    Dim str单位 As String, str盘点时间 As String, str汇总盘点单 As String
    Dim strOrder As String, strCompare As String
    Dim rsTemp As New ADODB.Recordset
    Dim strNo串 As String
    Dim str盘点单NO As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If str库房货位 = "" Then
        str库房货位 = "所有"
    ElseIf str库房货位 <> "所有" Then
        str库房货位 = Replace(str库房货位, "'", "")
    End If
    
    If str剂型 = "" Then str剂型 = "'zyb'"          '保证传入的剂型为空时，不查出任何药品
    
    str盘点时间 = txtCheckDate.Caption
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
    strCompare = Mid(strOrder, 1, 1)

    '汇总指定盘点时刻的盘点单
    str汇总盘点单 = " Union " & _
             " Select A.药品ID,B.编码,B.名称,E.库房货位" & _
             " From (select DISTINCT a.药品ID,a.库房ID FROM 药品收发记录 a " & _
             " Where a.单据=14 And a.库房ID+0=[1] And a.No in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))) A, " & _
             " 收费项目目录 B,药品储备限额 E " & _
             " Where A.药品ID+0=B.ID and A.库房id=E.库房id(+) and A.药品id+0=E.药品id(+) "
    If mbln忽略服务对象 = False Then
         str汇总盘点单 = str汇总盘点单 & " And(Decode(B.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3))" & _
                " or Decode(B.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]) )"
    End If
    
    '提取符合盘点条件的所有药品清单
    gstrSQL = "SELECT " & IIf(str库房货位 <> "所有", " /*+rule*/ ", "") & " Distinct A.药品ID,B.编码,B.名称,E.库房货位" & _
             " FROM 药品规格 A,收费项目目录 B,药品特性 C,诊疗项目目录 K,诊疗分类目录 L," & _
             "     (SELECT 药品ID,Nvl(实际数量,0) 实际数量,Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价 " & _
             "      FROM 药品库存 WHERE 库房ID=[1] AND 性质=1 " & IIf(bln盘无库存有金额药品 = True, " And 实际数量=0 And (实际金额<>0 Or 实际差价<>0)", " And (Nvl(实际数量,0)<>0 Or Nvl(实际金额,0)<>0 Or Nvl(实际差价,0)<>0 )") & ") D, "
    If bln汇总盘点单 Then
        gstrSQL = gstrSQL & "(SELECT 库房id, 药品id, 上限, 下限, 盘点属性, 库房货位 FROM 药品储备限额 WHERE 库房ID=[1]) E, " & _
             "     (SELECT 收费细目id, 病人来源, 开单科室id, 执行科室id FROM 收费执行科室 WHERE 执行科室ID=[1]) F " & _
             " WHERE A.药品ID=B.ID And A.药名ID=K.ID And K.分类ID=L.ID and L.类型 in (1,2,3) And A.药名ID=C.药名ID AND A.药品ID=F.收费细目ID" & IIf(mblnNoStock, "(+)", "") & _
             "  AND (B.撤档时间=TO_DATE('3000-01-01','yyyy-MM-dd') OR B.撤档时间 IS NULL Or B.撤档时间 BETWEEN To_Date('" & str盘点时间 & "', 'yyyy-mm-dd hh24:mi:ss') AND SYSDATE) " & _
             IIf(mstr分类ID = "", "", " AND L.ID in (select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) ") & _
             IIf(str剂型 = "所有", "", " AND C.药品剂型 in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) ") & _
             "  AND A.药品ID=D.药品ID" & IIf(bln盘无库存药品, "(+)", "") & " AND A.药品ID=E.药品ID(+)"
        If mbln忽略服务对象 = False Then
            gstrSQL = gstrSQL & " And(Decode(B.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3))" & _
                " or Decode(B.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]) )"
        End If
    Else
        If str库房货位 <> "所有" Then
'            gstrSQL = gstrSQL & "(SELECT A.药品id, A.库房货位 FROM 药品储备限额 A WHERE A.库房ID=[1] " & IIf(str盘点属性 = "所有", "", str盘点属性) & " And A.库房货位 in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))) E, "
            gstrSQL = gstrSQL & "(Select a.药品id, a.库房货位" & vbNewLine & _
                            "From 药品储备限额 A, (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) B" & vbNewLine & _
                            "Where a.库房id = [1] " & IIf(str盘点属性 = "所有", "", str盘点属性) & " And (Instr(',' || a.库房货位 || ',', ',' || b.Column_Value || ',') > 0)) E, "
        Else
            gstrSQL = gstrSQL & "(SELECT A.药品id, A.库房货位 FROM 药品储备限额 A WHERE A.库房ID=[1] " & IIf(str盘点属性 = "所有", "", str盘点属性) & " ) E, "
        End If
        
        gstrSQL = gstrSQL & " (SELECT 收费细目id, 病人来源, 开单科室id, 执行科室id FROM 收费执行科室 WHERE 执行科室ID=[1]) F " & _
             " WHERE A.药品ID=B.ID And A.药名ID=K.ID And K.分类ID=L.ID and L.类型 in (1,2,3) And A.药名ID=C.药名ID AND A.药品ID=F.收费细目ID" & IIf(mblnNoStock, "(+)", "") & " " & _
             IIf(mbln盘停用药品 = True, "", " AND (B.撤档时间=TO_DATE('3000-01-01','yyyy-MM-dd') OR B.撤档时间 IS NULL Or B.撤档时间 BETWEEN To_Date('" & str盘点时间 & "', 'yyyy-mm-dd hh24:mi:ss') AND SYSDATE) ") & _
             IIf(mstr分类ID = "", "", " AND L.ID in (select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) ") & _
             IIf(str剂型 = "所有", "", " AND C.药品剂型 in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) ") & _
             "  AND A.药品ID=D.药品ID" & IIf(bln盘无库存药品, "(+)", "") & " AND" & IIf(str盘点属性 = "所有", " A.药品ID=E.药品ID(+)", " A.药品ID=E.药品ID")
        If mbln忽略服务对象 = False Then
            gstrSQL = gstrSQL & " And(Decode(B.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3))" & _
                " or Decode(B.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]) )"
        End If
    End If
    If bln汇总盘点单 Then
        str盘点单NO = mstr盘点单号 & ","
        For i = 0 To UBound(Split(str盘点单NO, ","))
            If Split(str盘点单NO, ",")(i) <> "" Then
                strNo串 = IIf(strNo串 = "", "", strNo串 & ",") & Replace(Split(str盘点单NO, ",")(i), "'", "")
            End If
        Next
        
        If bln盘点单 = False Then
            gstrSQL = gstrSQL & str汇总盘点单
        Else
            gstrSQL = Replace(str汇总盘点单, " Union", "")
        End If
    End If
    
    gstrSQL = gstrSQL & " and b.建档时间 <=To_Date('" & str盘点时间 & "', 'yyyy-mm-dd hh24:mi:ss') "
    
    gstrSQL = gstrSQL & " Order by " & _
              IIf(strCompare = "0", "编码", IIf(strCompare = "1", "编码", IIf(strCompare = "2", "名称", "库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",编码"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取出符合条件的药品]", lng库房ID, str库房货位, mstr分类ID, str剂型, strNo串)
    
    Set GetPhysic = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPhysicDetail(ByVal lng库房ID As Long, ByVal lng药品ID As Long, _
    Optional ByVal bln盘无库存药品 As Boolean = True, Optional ByVal bln汇总盘点单 As Boolean = False, Optional ByVal bln盘无库存有金额药品 As Boolean = False) As ADODB.Recordset
    'bln盘无库存药品=是否将无库存药品也提取出来
    'bln汇总盘点单=是否需要汇总指定盘点时间的盘点单形成盘点表
    '提取该药品当前库房所有批次明细记录
    Dim str单位 As String, str盘点时间 As String, str汇总盘点单 As String, str汇总盘点单新增批次 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql大包装 As String
    Dim strSql小包装 As String
    Dim strSql盘点时间之后发生 As String
    Dim str盘点单NO As String
    Dim strNo串 As String
    Dim i As Integer
    Dim bln当前库存 As Boolean, bln可用数量占用 As Boolean '对应列是否隐藏
    
    On Error GoTo ErrHand
    bln当前库存 = vsfBill.colHidden(mconintCol当前库存)
    bln可用数量占用 = vsfBill.colHidden(mconintCol可用数量占用)
    
    str盘点时间 = txtCheckDate.Caption
    
    If mintUnit > 0 Then
        Select Case mintUnit
            Case mconint售价单位
                str单位 = ",E.计算单位 As 单位,1 As 比例系数"
            Case mconint门诊单位
                str单位 = ",A.门诊单位 As 单位,A.门诊包装 As 比例系数"
            Case mconint住院单位
                str单位 = ",A.住院单位 As 单位,A.住院包装 As 比例系数"
            Case mconint药库单位
                str单位 = ",A.药库单位 As 单位,A.药库包装 As 比例系数"
        End Select
    Else
        Select Case mint大单位
            Case mconint售价单位
                strSql大包装 = ",E.计算单位 As 大包装单位,1 As 比例系数大"
            Case mconint门诊单位
                strSql大包装 = ",A.门诊单位 As 大包装单位,A.门诊包装 As 比例系数大"
            Case mconint住院单位
                strSql大包装 = ",A.住院单位 As 大包装单位,A.住院包装 As 比例系数大"
            Case mconint药库单位
                strSql大包装 = ",A.药库单位 As 大包装单位,A.药库包装 As 比例系数大"
        End Select
        
        Select Case mint小单位
            Case mconint售价单位
                strSql小包装 = ",E.计算单位 As 小包装单位,1 As 比例系数小"
            Case mconint门诊单位
                strSql小包装 = ",A.门诊单位 As 小包装单位,A.门诊包装 As 比例系数小"
            Case mconint住院单位
                strSql小包装 = ",A.住院单位 As 小包装单位,A.住院包装 As 比例系数小"
            Case mconint药库单位
                strSql小包装 = ",A.药库单位 As 小包装单位,A.药库包装 As 比例系数小"
        End Select
        
        str单位 = strSql大包装 & strSql小包装
    End If
    
    '汇总盘点单的SQL
    If bln汇总盘点单 Then
        str盘点单NO = mstr盘点单号 & ","
        For i = 0 To UBound(Split(str盘点单NO, ","))
            If Split(str盘点单NO, ",")(i) <> "" Then
                strNo串 = IIf(strNo串 = "", "", strNo串 & ",") & Replace(Split(str盘点单NO, ",")(i), "'", "")
            End If
        Next
        
        '35.60支持盘点单录入多个新增批次
        str汇总盘点单 = "" & _
            " UNION ALL" & _
            " SELECT A.库房ID,A.药品ID,NVL(A.批次, 0) AS 批次,0 AS 实际数量,A.扣率 As 盘点数量," & _
                    " 0 AS 实际金额,0 AS 实际差价,0 AS 可用数量,A.批号,A.产地,A.原产地,A.效期,A.批准文号" & IIf(bln当前库存, "", " , 0 当前库存") & _
            " FROM 药品收发记录 A " & _
            " Where A.单据=14 AND A.库房ID+0=[1] And Nvl(a.批次, 0) <> -1 AND a.No in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) "
            
        
        str汇总盘点单新增批次 = "" & _
            " UNION ALL" & _
            " Select 库房id, 药品id, 批次, Sum(实际数量) As 帐面数量, Sum(盘点数量) As 盘点数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价," & _
            " Sum(可用数量) As 可用数量, Max(批号) As 批号, Max(产地) As 产地, Max(原产地) As 原产地, Max(效期) As 效期, Max(批准文号) As 批准文号, 成本价" & IIf(bln当前库存, "", " , 0 当前库存") & _
            " from (SELECT A.库房ID,A.药品ID,NVL(A.批次, 0) AS 批次,0 AS 实际数量,A.扣率 As 盘点数量," & _
                    " 0 AS 实际金额,0 AS 实际差价,0 AS 可用数量,A.批号,A.产地,A.原产地,A.效期,A.批准文号, a.单量 As 成本价 " & _
            " FROM 药品收发记录 A " & _
            " Where A.单据=14 AND A.库房ID+0=[1] And Nvl(a.批次, 0) = -1 AND a.No in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))) " & _
            " GROUP BY 库房ID, 药品ID, 批次, 产地,原产地,批号, 成本价 "
    End If
    
    If mbln忽略盘点时间 = False Then
        strSql盘点时间之后发生 = "" & _
            " Union All" & _
            " SELECT A.库房ID,A.药品ID,NVL(A.批次,0) AS 批次,-1*A.入出系数*A.实际数量*A.付数 AS 实际数量,0 盘点数量," & _
            " -1*A.入出系数*A.零售金额 AS 实际金额, -1*A.入出系数*A.差价 AS 实际差价,0 AS 可用数量,A.批号,A.产地,a.原产地,A.效期,A.批准文号" & IIf(bln当前库存, "", " , 0 当前库存") & _
            " FROM 药品收发记录 A" & _
            " Where A.库房ID+0=[1] And A.药品ID+0=[2] " & _
            " AND A.审核日期 > [3] "
    End If
    
    '取药品当前库存及盘点时间以后的净发生额
    gstrSQL = "" & _
        " SELECT DISTINCT A.药品ID,A.成本价 As 平均成本价,E.产地 缺省产地,'[' || E.编码 || ']' As 药品编码, E.名称 As 通用名, C.名称 As 商品名,A.药库分批 AS 分批核算,A.药房分批 AS 药房分批核算,E.是否变价,A.加成率," & _
        "        NVL(B.实际金额,0) 实际金额,NVL(B.实际差价,0) 实际差价,D.现价 售价,NVL(B.批次,0) 批次,A.药品来源,A.基本药物,Decode(b.批号, Null, a.上次批号, b.批号) As 批号,B.效期,F.库房货位,E.规格,decode(b.产地,null,decode(a.上次产地,null,e.产地,a.上次产地),b.产地) as 产地," & _
        "        Decode(b.原产地, Null,a.原产地,b.原产地) As 原产地,B.批准文号,Nvl(B.帐面数量,0) 帐面数量,B.盘点数量,B.可用数量" & str单位 & ",Decode(b.批次, -1, b.成本价, Decode(x.现价, Null, Decode(k.成本价, Null, a.成本价, k.成本价), x.现价)) As 成本价, " & _
        "        Nvl(E.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间,a.最大效期" & IIf(bln当前库存, "", ",nvl(b.当前库存,0) as 当前库存") & IIf(bln可用数量占用, "", " ,nvl(y.可用数量占用,0) 可用数量占用 ") & _
        " FROM (SELECT 库房ID, 药品ID, 批次, SUM (实际数量) AS 帐面数量,SUM (盘点数量) AS 盘点数量,SUM (实际金额) AS 实际金额," & _
        "         SUM (实际差价) AS 实际差价, SUM(可用数量) AS 可用数量,MAX(批号) As 批号, MAX(产地) AS 产地 ,Max(原产地) As 原产地,MAX(效期) AS 效期, MAX(批准文号) As 批准文号, 0 As 成本价 " & IIf(bln当前库存, "", " ,Sum(当前库存) As 当前库存 ") & _
        "         From" & _
        "             ( SELECT A.库房ID,A.药品ID,NVL(批次,0) AS 批次,Nvl(A.实际数量,0) 实际数量,0 盘点数量,Nvl(A.实际金额,0) 实际金额,Nvl(A.实际差价,0) 实际差价,Nvl(A.可用数量,0) 可用数量,A.上次批号 AS 批号,A.上次产地 AS 产地,a.原产地,A.效期,A.批准文号 " & IIf(bln当前库存, "", ",Nvl(a.实际数量, 0) 当前库存") & _
        "             FROM 药品库存 A" & _
        "             Where A.性质 = 1 And A.库房ID=[1] And A.药品ID=[2] " & IIf(bln盘无库存有金额药品 = True, " And Nvl(A.实际数量,0)=0 And (Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0)", " And (Nvl(A.实际数量,0)<>0 Or Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0 )") & _
        IIf(mbln忽略盘点时间 = True, "", strSql盘点时间之后发生) & _
        IIf(Not bln汇总盘点单, "", str汇总盘点单) & _
        "     ) GROUP BY 库房ID, 药品ID, 批次 " & IIf(Not bln汇总盘点单, "", str汇总盘点单新增批次) & _
        ") B, 收费价目 D, 收费项目别名 C, 收费项目目录 E, 药品规格 A," & _
        "      (Select x.药品id,x.库房id,x.批次,nvl(x.现价,0) 现价 from 药品价格记录 x where x.价格类型 = 2 and [3] between x.执行日期 and x.终止日期) X," & _
        IIf(bln可用数量占用, "", "(Select sum(y.实际数量) 可用数量占用 ,y.库房id,y.药品id,y.批次 From 药品收发记录 Y Where y.入出系数 = -1 And y.审核日期 is null and y.填制日期 > (sysdate - 30)  Group By y.库房id, y.药品id, y.批次) Y,") & _
        "      (Select 药品id,批次,平均成本价 成本价 From 药品库存 Where 性质 = 1 And 库房id =[1] " & IIf(bln盘无库存有金额药品 = True, " And 实际数量=0 And (实际金额<>0 Or 实际差价<>0)", "") & ") K,药品储备限额 F " & _
        " Where A.药品ID=E.ID And A.药品ID=B.药品ID" & IIf(bln盘无库存药品, "(+)", "") & _
        " AND A.药品ID=F.药品ID(+) And B.药品id=K.药品id(+) And Nvl(B.批次, 0)=nvl(K.批次(+),0) " & _
        " AND b.药品id = x.药品id(+) And b.库房id = x.库房id(+) And Nvl(b.批次, 0) = Nvl(x.批次(+), 0) " & IIf(bln可用数量占用, "", " AND b.药品id = y.药品id(+) And b.库房id = y.库房id(+) And Nvl(b.批次, 0) = Nvl(y.批次(+), 0) ") & _
        " AND A.药品ID=C.收费细目ID(+) AND C.性质(+)=3 AND A.药品ID=D.收费细目ID(+)  " & _
        " AND F.库房ID(+)=[1] And A.药品ID+0=[2] AND D.执行日期(+)<=[3] AND NVL(D.终止日期(+),SYSDATE)>=[3] " & _
        GetPriceClassString("D") & _
        " and e.建档时间<=[3]  Order by 批次 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品当前库房所有批次明细记录]", lng库房ID, lng药品ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")), strNo串)
    
    Set GetPhysicDetail = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function 时价药品零售价(ByVal lng药品ID As Long, ByVal sin加成率 As Double, ByVal sin采购价 As Single) As Double
    Dim sin零售价 As Single, sin指导零售价 As Single, sin差价让利比 As Single
    Dim rsTemp As New ADODB.Recordset
    '时价药品零售价计算公式:采购价*(1+加成率)
    '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    On Error GoTo errHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    
    时价药品零售价 = 0
    
    sin零售价 = sin采购价 * (1 + sin加成率)
    sin零售价 = sin零售价 + (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    时价药品零售价 = IIf(sin零售价 > sin指导零售价, sin指导零售价, sin零售价)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub vsfBill_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfBill.EditSelStart = 0
    vsfBill.EditSelLength = zlStr.ActualLen(vsfBill.EditText)
End Sub

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String
    Dim intMoneyBit As Integer
    Dim intNumber As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim dbl成本价 As Double
    Dim dblSumNum As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim dbl盘点金额 As Double
    
    On Error GoTo errHandle
    With vsfBill
        .Redraw = flexRDNone
        
        .EditText = Trim(.EditText)
        strkey = Trim(.EditText)
        
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        Select Case Col
            Case mconIntCol批号
                .TextMatrix(Row, Col) = strkey
            Case mconIntCol效期
                '有处理
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            .EditText = ""
                            MsgBox "对不起，失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Exit Sub
                        End If
                        .EditText = strkey
                    ElseIf Not IsDate(strkey) Then
                        .EditText = ""
                        MsgBox "对不起，失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(Row, Col) = strkey
                Call CheckLapse(.TextMatrix(.Row, mconIntCol效期)) '检查是否失效
            Case mconintCol实盘数量
                If .TextMatrix(Row, Col) = "" Or strkey = "" Then
                    MsgBox "对不起，实盘数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，实盘数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                
                If strkey <> "" And .TextMatrix(Row, 0) <> "" And Val(strkey) <> Val(.TextMatrix(Row, mconintCol实盘数量)) Then
                    strkey = zlStr.FormatEx(strkey, mintNumberDigit, , True)
                    .EditText = strkey
                    
                    .TextMatrix(Row, mconintCol数量差) = zlStr.FormatEx(Abs(Val(strkey) - Val(.TextMatrix(Row, mconintCol帐面数量))), mintNumberDigit, , True)
                    If Val(strkey) > Val(.TextMatrix(Row, mconintCol帐面数量)) Then
                        .TextMatrix(Row, mconintCol标志) = "盈"
                    ElseIf Val(strkey) < Val(.TextMatrix(Row, mconintCol帐面数量)) Then
                        .TextMatrix(Row, mconintCol标志) = "亏"
                    Else
                        .TextMatrix(Row, mconintCol标志) = "平"
                    End If
                    
                    '实盘数量为0与库存数量比较判断盈亏
                    If Val(strkey) = 0 And Val(.TextMatrix(Row, mconintCol库存数量)) > 0 Then
                        .TextMatrix(Row, mconintCol标志) = "亏"
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
                    '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
                    '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
                    If Val(.TextMatrix(Row, mconIntCol新批次)) = 1 Then
                        intMoneyBit = mintMoneyDigit
                    ElseIf Val(strkey) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True And Val(.TextMatrix(Row, mconIntCol售价)) = Val(.TextMatrix(Row, mconintCol成本价))) Then
                        '盘0或者零差价药品盘点时
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                    
                    '金额差=当前售价*实盘数量-实际金额
                    '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                    dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * Val(strkey), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                    .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(Row, mconIntCol实际金额)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol售价)) - Val(.TextMatrix(Row, mconintCol成本价))) * Val(strkey) - Val(.TextMatrix(Row, mconIntCol实际差价)), intMoneyBit, , True)
                    dbl金额差 = Val(.TextMatrix(Row, mconintCol金额差))
                    dbl差价差 = Val(.TextMatrix(Row, mconintCol差价差))
                    If .TextMatrix(Row, mconintCol标志) = "亏" Then
                        .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol金额差)), intMoneyBit, , True)
                        .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(Row, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * Val(strkey), mintMoneyDigit, , True)
                    .TextMatrix(Row, mconintCol实盘数量) = strkey
                    
                    '.TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol成本价)) * Val(.TextMatrix(Row, mconintCol实盘数量)), mintMoneyDigit)
                    '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                    .TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol实际金额)) + dbl金额差) - (Val(.TextMatrix(Row, mconIntCol实际差价)) + dbl差价差), mintMoneyDigit, , True)
                    .TextMatrix(Row, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol金额差)) - Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
                    
                    '盘亏盘盈行用颜色区分
                    Call SetStocktakingColor(vsfBill, .Row)
                End If
                
                Call 显示合计金额
        Case mconintCol大包装实盘数量, mconintCol小包装实盘数量
            If .TextMatrix(Row, Col) = "" Or strkey = "" Then
                MsgBox "对不起，实盘数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Not IsNumeric(strkey) And strkey <> "" Then
                MsgBox "对不起，实盘数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If mintUnit > 0 Then
                intNumber = mintNumberDigit
            Else
                intNumber = mintNumberDigit0
            End If
               
            If strkey <> "" And .TextMatrix(Row, 0) <> "" Then
                strkey = zlStr.FormatEx(strkey, intNumber, , True)
                .EditText = strkey
                
                '换算成小包装单位来汇总实盘数量
                If .Col = mconintCol大包装实盘数量 Then
                    dblSumNum = Val(strkey) * Val(.TextMatrix(Row, mconIntCol比例系数大)) / Val(.TextMatrix(Row, mconIntCol比例系数小)) + Val(.TextMatrix(Row, mconintCol小包装实盘数量))
                Else
                    dblSumNum = Val(.TextMatrix(Row, mconintCol大包装实盘数量)) * Val(.TextMatrix(Row, mconIntCol比例系数大)) / Val(.TextMatrix(Row, mconIntCol比例系数小)) + Val(strkey)
                End If
                
                .TextMatrix(Row, mconintCol实盘数量) = zlStr.FormatEx(dblSumNum, intNumber, , True)
                .TextMatrix(Row, mconintCol合计) = .TextMatrix(Row, mconintCol实盘数量) & .TextMatrix(Row, mconIntCol实盘数量单位小)
                
                .TextMatrix(Row, mconintCol数量差) = zlStr.FormatEx(Abs(Val(.TextMatrix(Row, mconintCol实盘数量)) - Val(.TextMatrix(Row, mconintCol帐面数量))), intNumber, , True)
                
                If dblSumNum > Val(.TextMatrix(Row, mconintCol帐面数量)) Then
                    .TextMatrix(Row, mconintCol标志) = "盈"
                ElseIf dblSumNum < Val(.TextMatrix(Row, mconintCol帐面数量)) Then
                    .TextMatrix(Row, mconintCol标志) = "亏"
                Else
                    .TextMatrix(Row, mconintCol标志) = "平"
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(Row, mconintCol实盘数量)) = 0 And Val(.TextMatrix(Row, mconintCol库存数量)) > 0 Then
                    .TextMatrix(Row, mconintCol标志) = "亏"
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '解决药品库存中数量为0，金额或差价不为0的药品无法通过盘点清除库存记录的问题
                '这种情况下的通常药品库存金额或差价的实际位数多于系统参数中设置的金额位数
                '解决办法是如果实盘数量为0，则金额差和差价差小数位数保持和药品库存表中金额和差价位数一致
                If Val(.TextMatrix(Row, mconIntCol新批次)) = 1 Then
                    intMoneyBit = mintMoneyDigit
                ElseIf dblSumNum = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True And Val(.TextMatrix(Row, mconIntCol售价)) = Val(.TextMatrix(Row, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * dblSumNum, mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(Row, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol售价)) - Val(.TextMatrix(Row, mconintCol成本价))) * Val(dblSumNum) - Val(.TextMatrix(Row, mconIntCol实际差价)), intMoneyBit, , True)
                dbl金额差 = Val(.TextMatrix(Row, mconintCol金额差))
                dbl差价差 = Val(.TextMatrix(Row, mconintCol差价差))
                If .TextMatrix(Row, mconintCol标志) = "亏" Then
                    .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                .TextMatrix(Row, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * dblSumNum, mintMoneyDigit, , True)
                '.TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol成本价)) * Val(.TextMatrix(Row, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol实际金额)) + dbl金额差) - (Val(.TextMatrix(Row, mconIntCol实际差价)) + dbl差价差), mintMoneyDigit, , True)
                .TextMatrix(Row, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol金额差)) - Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
                
                 '盘亏盘盈行用颜色区分
                 Call SetStocktakingColor(vsfBill, .Row)
            End If
            
            Call 显示合计金额
        Case mconintCol成本价
            If .TextMatrix(Row, Col) = "" Or strkey = "" Then
                MsgBox "对不起，价格必须输入！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Not IsNumeric(strkey) And strkey <> "" Then
                MsgBox "对不起，价格必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Val(strkey) < 0 Then
                MsgBox "对不起，价格不能为负数,请重输！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
                
            If strkey <> "" And .TextMatrix(Row, 0) <> "" Then
                strkey = zlStr.FormatEx(strkey, mintCostDigit, , True)
                .EditText = strkey
                
                If Split(.TextMatrix(Row, mconIntcol加成率), "||")(1) = 1 Then
                    '时价药品时
                    If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                        '零差价管理，售价等于成本价
                        .TextMatrix(Row, mconIntCol售价) = strkey
                    End If
                Else
                    '定价药品
                    If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                        '零差价管理，要判断成本价是否等于售价
                        If Val(strkey) <> Val(.TextMatrix(Row, mconIntCol售价)) Then
                            MsgBox "该定价药品已启用零差价管理模式，入库成本价应和售价(" & .TextMatrix(Row, mconIntCol售价) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                            strkey = .TextMatrix(Row, mconIntCol售价)
                            .TextMatrix(.Row, mconintCol成本价) = zlStr.FormatEx(strkey, mintCostDigit, , True)
                            .EditText = strkey
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(Row, mconIntCol新批次)) = 1 Then
                    intMoneyBit = mintMoneyDigit
                ElseIf IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                If mintUnit > 0 Then
                    dblSumNum = Val(.TextMatrix(Row, mconintCol实盘数量))
                Else
                    dblSumNum = Val(.TextMatrix(Row, mconintCol大包装实盘数量)) * Val(.TextMatrix(Row, mconIntCol比例系数大)) / Val(.TextMatrix(Row, mconIntCol比例系数小)) + Val(.TextMatrix(Row, mconintCol小包装实盘数量))
                End If
                                   
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * dblSumNum, mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(Row, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol售价)) - Val(strkey)) * Val(dblSumNum) - Val(.TextMatrix(Row, mconIntCol实际差价)), intMoneyBit, , True)
                dbl金额差 = Val(.TextMatrix(Row, mconintCol金额差))
                dbl差价差 = Val(.TextMatrix(Row, mconintCol差价差))
                If .TextMatrix(Row, mconintCol标志) = "亏" Then
                    .TextMatrix(Row, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
                End If
                                    
                .TextMatrix(Row, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol售价)) * dblSumNum, mintMoneyDigit, , True)
                '.TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol成本价)) * Val(.TextMatrix(Row, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(Row, mconintCol盘点成本金额) = zlStr.FormatEx(Val(strkey) * dblSumNum, mintMoneyDigit, , True)
                .TextMatrix(Row, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol金额差)) - Val(.TextMatrix(Row, mconintCol差价差)), intMoneyBit, , True)
            
            End If
        End Select
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, Row, mconintCol实盘数量, Row, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, Row, mconintCol大包装实盘数量, Row, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, Row, mconintCol小包装实盘数量, Row, mconintCol小包装实盘数量) = True
        End If
        
        If mblnKeyPressReturn = True Then
            vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
        End If

        mblnKeyPressReturn = False
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Get药品分批属性(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int分批属性 As Integer      '0-不分批;1-分批
    Dim int药库分批 As Integer      '0-不分批;1-分批
    Dim int药房分批 As Integer      '0-不分批;1-分批
    Dim bln是否具有药房性质 As Boolean  'True-具有药房性质;False-不具有药房性质
    
    If Val(vsfBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo errHandle
    strSQL = "SELECT NVL(药库分批, 0) 药库分批,NVL(药房分批, 0) 药房分批 " & _
            " From 药品规格 WHERE 药品ID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取药品库房分批属性", Val(vsfBill.TextMatrix(intBillRow, 0)))
    
    If rsTemp.RecordCount > 0 Then
        int药库分批 = rsTemp!药库分批
        int药房分批 = rsTemp!药房分批
    End If
    
    If int药房分批 = 1 Then     '如果药房分批，则分批属性为1
        int分批属性 = 1
    Else
        If int药库分批 = 1 Then
            strSQL = "SELECT 部门ID From 部门性质说明 " & _
                    " WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门ID = [1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取部门性质", txtStock.Tag)
            
            bln是否具有药房性质 = (rsTemp.RecordCount > 0)
                    
            If bln是否具有药房性质 Then
                int分批属性 = 0
            Else
                int分批属性 = 1
            End If
        End If
    End If
    
    vsfBill.TextMatrix(intBillRow, mconIntCol分批属性) = int分批属性
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get盘点时刻零售价(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double, ByVal date盘点时刻 As Date) As Double
    '功能：获取指定时刻时价药品当前药品的零售价
    '参数:药品id,库房id,批次,盘点时刻
    '返回值：零售价
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle
    '1、判断药品价格记录是否有数据
    gstrSQL = "select nvl(现价,0) as 零售价 from 药品价格记录 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 价格类型 = 1 and [4] between 执行日期 and 终止日期"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID, lng库房ID, lng批次, date盘点时刻)
    
    If rsData.EOF Then '无对应的药品价格记录
    
        gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, Nvl(实际金额, 0) / 实际数量), 零售价) as 零售价 from 药品库存 where 性质 = 1 and 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID, lng库房ID, lng批次)
        
        If rsData.EOF Then
            '时价药品零售价计算公式:采购价*(1+加成率)
            '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get盘点时刻零售价 = 0
            dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次, date盘点时刻)
            dbl加成率 = rsData!加成率 / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get盘点时刻零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价) * dbl比例系数
        Else
            If rsData!零售价 = 0 Then
                gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get盘点时刻零售价 = 0
                dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次, date盘点时刻)
                dbl加成率 = rsData!加成率 / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get盘点时刻零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价) * dbl比例系数
            Else
                Get盘点时刻零售价 = rsData!零售价 * dbl比例系数
            End If
        End If
    Else '有对应药品价格记录
        Get盘点时刻零售价 = rsData!零售价 * dbl比例系数
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get盘点时刻售价(ByVal bln是否时价 As Boolean, lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, ByVal date盘点时刻 As Date) As Double
    '功能：获取原始的售价单位售价，主要用于出库
    '参数: bln是否时价:false-定价,true-时价
    '返回值：最小单位的价格
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle

    '取定价药品售价
    If bln是否时价 = False Then
        gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And [2] Between A.执行日期 And Nvl(A.终止日期,Sysdate) " & GetPriceClassString("A")
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "Get盘点时刻售价-取定价药品售价", lng药品ID, date盘点时刻)
        
        If Not rsData.EOF Then
            Get盘点时刻售价 = rsData!现价
        End If
    Else
        '取时价药品售价
        '1、判断药品价格记录是否有数据
        gstrSQL = "select nvl(现价,0) as 零售价 from 药品价格记录 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 价格类型 = 1 and [4] between 执行日期 and 终止日期"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID, lng库房ID, lng批次, date盘点时刻)
        
        If rsData.EOF Then '无对应的药品价格记录
        
            gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 " & _
                " from 药品库存 where 性质=1 and  药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "GetOriPrice-零售价", lng药品ID, lng库房ID, lng批次)
            
            If rsData.EOF Then
                '时价药品零售价计算公式:采购价*(1+加成率)
                '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
                '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
                gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get盘点时刻售价 = 0
                dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次, date盘点时刻)
                dbl加成率 = rsData!加成率 / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get盘点时刻售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
            Else
                If rsData!零售价 = 0 Then
                    gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
                    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
                    dbl指导零售价 = rsData!指导零售价
                    dbl差价让利比 = rsData!差价让利比
                    
                    Get盘点时刻售价 = 0
                    dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次, date盘点时刻)
                    dbl加成率 = rsData!加成率 / 100
                    dbl零售价 = dbl成本价 * (1 + dbl加成率)
                    dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                    Get盘点时刻售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
                Else
                    Get盘点时刻售价 = rsData!零售价
                End If
            End If
        Else
            Get盘点时刻售价 = rsData!零售价
        End If
        
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get盘点时刻成本价(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, ByVal date盘点时刻 As Date) As Double
'功能：获取当前药品的成本价格
'参数：药品id,库房id,批次
'返回值： 成本价格
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    '1、判断药品价格记录是否有数据
    gstrSQL = "select nvl(现价,0) as 成本价 from 药品价格记录 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 价格类型 = 2 and [4] between 执行日期 and 终止日期"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "成本价", lng药品ID, lng库房ID, lng批次, date盘点时刻)
    
    If rsData.EOF Then '无对应的药品价格记录
    
        gstrSQL = "select 平均成本价 from 药品库存 where 性质=1 and 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "成本价", lng药品ID, lng库房ID, lng批次)
        
        If rsData.EOF Then
            blnNullPrice = True
        ElseIf IsNull(rsData!平均成本价) = True Then
            blnNullPrice = True
        ElseIf Val(rsData!平均成本价) < 0 Then
            blnNullPrice = True
        End If
        
        If Not blnNullPrice Then
            Get盘点时刻成本价 = rsData!平均成本价
        Else
            '如果无法从库存中取成本价，则从药品规格中取
            gstrSQL = "select 成本价 from 药品规格 where 药品id=[1]"
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "成本价", lng药品ID)
            If Not rsData.EOF Then
                If Val(Nvl(rsData!成本价, 0)) > 0 Then
                    Get盘点时刻成本价 = rsData!成本价
                End If
            End If
        End If
    Else
        Get盘点时刻成本价 = rsData!成本价
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'功能：将vsf表格存在的列并且在列设置中未勾选隐藏的列进行隐藏
Private Sub vsfHidden(ByRef objVSF As Object)
    Dim strColsName As String
    Dim strColName() As String
    Dim i As Integer
    Dim n As Integer
    Dim strDefaultColsName As String '默认的列
    Dim strTempColName As String
    
    strDefaultColsName = ":药品来源,0:基本药物,0:库房货位,0:批准文号,0:金额差,0:差价差,0:盘点成本金额差,0:账面金额差,0:成本金额差,0:当前库存,1:" '所有可以隐藏的列
    
    With objVSF
        strColsName = zlDataBase.GetPara("列设置", glngSys, 1307, "")
        
        '兼容处理
        If strColsName = "" Then '未提取到列设置信息
            strColsName = strDefaultColsName
        Else
            '判断提取的列与默认列个数，不一致则取默认的
            If UBound(Split(strColsName, ":")) <> UBound(Split(strDefaultColsName, ":")) Then strColsName = strDefaultColsName
            
            '判断提取的列名是否与默认的一致，不一致取默认的
            For i = LBound(Split(strColsName, ":")) + 1 To UBound(Split(strColsName, ":")) - 1
                strTempColName = Split(Split(strColsName, ":")(i), ",")(0) '获取单个列名
                
                If InStr(1, strDefaultColsName, ":" & strTempColName) = 0 Then '列名不存在于默认列名中
                    strColsName = strDefaultColsName
                    Exit For
                End If
            Next
            
        End If
        
        strColName = Split(strColsName, ":") '格式:C,1
        
        For i = 0 To .Cols - 1
            '判断界面对应列是否是可屏蔽列
            If InStr(1, strColsName, ":" & .TextMatrix(0, i)) > 0 Then
                For n = LBound(strColName) + 1 To UBound(strColName) - 1
                    If Split(strColName(n), ",")(0) = .TextMatrix(0, i) Then .colHidden(i) = Split(strColName(n), ",")(1) <> 1
                Next
            End If
             
        Next
    End With
End Sub


Private Sub SearchAllData(ByVal lng库房ID As Long)
'功能：获取当前库房的所有药品
'参数：库房id
    Dim rsPhysic As ADODB.Recordset '记录该库房的所有药品ID（包括设置存储属性和有库存未设置存储属性的药品）
    Dim rsDetail As ADODB.Recordset
    Dim bln库房 As Boolean
    Dim dbl成本价 As Double, dbl零售价 As Double
    Dim str盘点时间 As String
    Dim str药名 As String
    Dim intMoneyBit As Integer
    Dim dbl盘点金额 As Double
    
    On Error GoTo errHandle
    
    Call FS.ShowFlash("正在装入药品数据,请稍候 ...", Me)
    
    str盘点时间 = txtCheckDate
    
    gstrSQL = "Select a.Id 药品ID" & vbNewLine & _
            "From 收费项目目录 A, 药品规格 B, 收费执行科室 C" & vbNewLine & _
            "Where a.Id = b.药品id And a.Id = c.收费细目id And c.执行科室id = [1]" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Distinct a.药品id" & vbNewLine & _
            "From 药品库存 A" & vbNewLine & _
            "Where a.库房id = [1] and a.性质 = 1 And Not Exists (Select 1 From 收费执行科室 C Where c.收费细目id = a.药品id And c.执行科室id = [1])"
        
    Set rsPhysic = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取出符合条件的药品]", lng库房ID)
    
    
    If rsPhysic.RecordCount = 0 Then
        MsgBox "未能正确读取药品库存数据,请重试！", vbInformation, gstrSysName: Exit Sub
    End If
    
    DoEvents
    vsfBill.Redraw = flexRDNone
    
    bln库房 = CheckPartProp(lng库房ID)
    With vsfBill
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            '取该药品的详细信息（可能分多个批次）
            
            Set rsDetail = GetPhysicDetail(lng库房ID, rsPhysic!药品id, True, False, False)
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                '时价药品重算售价
                dbl成本价 = zlStr.Nvl(rsDetail!平均成本价, 0)
                dbl零售价 = zlStr.Nvl(rsDetail!售价, 0)
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(CLng(rsPhysic!药品id), lng库房ID, CLng(rsDetail!批次), 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = rsPhysic!药品id
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = zlStr.Nvl(rsDetail!产地, zlStr.Nvl(rsDetail!缺省产地))
                .TextMatrix(.rows - 1, mconIntCol原产地) = zlStr.Nvl(rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                .TextMatrix(.rows - 1, mconintCol标志) = "平"
                .TextMatrix(.rows - 1, mconintCol数量差) = "0"
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol帐面数量), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = .TextMatrix(.rows - 1, mconintCol大包装帐面数量)

                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & .TextMatrix(.rows - 1, mconIntCol实盘数量单位小)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If

                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False '先恢复不加粗
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                '如果是分批药品，将批次改填为-1，表示新增批次
                .TextMatrix(.rows - 1, mconIntCol批次) = zlStr.Nvl(rsDetail!批次, 0)
                If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(.rows - 1, mconIntCol批次)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol批次) = -1
'                    '调试用，自动为新增批次设置批号与效期
'                    .TextMatrix(.Rows - 1, mconIntCol批号) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntCol效期) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                '如果是停用药品，该行粗体显示
                If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                End If
                    
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + Val(.TextMatrix(.rows - 1, mconintCol金额差))) - (zlStr.Nvl(rsDetail!实际差价, 0) + Val(.TextMatrix(.rows - 1, mconintCol差价差))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
        
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                .RowData(.rows - 1) = Val(IIf(IsNull(rsDetail!最大效期), 0, rsDetail!最大效期))
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmNewCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
    Else
        vsfBill.Col = mconIntCol药名
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
        vsfBill.EditCell
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SearchSpecialData(ByVal lng库房ID As Long, ByVal int盘点类型 As Integer, ByVal str近效期 As String, ByVal str毒理 As String)
'功能：获取当前库房的特殊药品
'参数：库房id
'      int盘点类型：0.近效期药品，1.毒麻精神药品,2.停用药品,3.无数量但有库存金额的药品,5.基本药物
'      str近效期：格式C1:C2；C1表示近效期的天数，为0时C2为1则表示只选择了已失效的
'      str毒理：C1:C2:C3:C4(麻醉药，毒性药，精神I类药，精神II类药),例：1:0:0:0,表示只选择了麻醉药
    Dim rsPhysic As ADODB.Recordset '记录该库房的所有药品ID（包括设置存储属性和有库存未设置存储属性的药品）
    Dim rsDetail As ADODB.Recordset
    Dim bln库房 As Boolean
    Dim dbl成本价 As Double, dbl零售价 As Double
    Dim str盘点时间 As String
    Dim str药名 As String
    Dim intMoneyBit As Integer
    Dim strSql条件 As String
    Dim int近效期天数 As Integer
    Dim bln已失效 As Boolean '是否读取失效药品
    Dim str毒理分类 As String
    Dim bln盘无库存药品 As Boolean
    Dim bln盘无库存有金额药品 As Boolean
    Dim dbl盘点金额 As Double
    
    On Error GoTo errHandle
    
    Call FS.ShowFlash("正在装入药品数据,请稍候 ...", Me)
    
    str盘点时间 = txtCheckDate
    
    '组装sql
    If int盘点类型 = 0 Then '近效期药品
        int近效期天数 = Split(str近效期, ":")(0)
        bln已失效 = Split(str近效期, ":")(1) = 1
        
        bln盘无库存药品 = False
        bln盘无库存有金额药品 = False
    
        If int近效期天数 <> 0 Then strSql条件 = " (效期 Between Trunc(Sysdate) And (Trunc(Sysdate) + [1])) "
        If bln已失效 Then strSql条件 = strSql条件 & IIf(int近效期天数 = 0, "", "or") & " 效期 < sysdate "
        
        gstrSQL = "Select distinct 药品id" & vbNewLine & _
            "From 药品库存" & vbNewLine & _
            "Where (" & strSql条件 & ") And 效期 Is Not Null and 库房id = [2] and 性质 = 1 "
    ElseIf int盘点类型 = 1 Then '毒麻药品
        bln盘无库存药品 = True
        bln盘无库存有金额药品 = False
    
        If Split(str毒理, ":")(0) = 1 Then str毒理分类 = "'麻醉药'"
        If Split(str毒理, ":")(1) = 1 Then str毒理分类 = str毒理分类 & IIf(str毒理分类 = "", "'", ",'") & "毒性药'"
        If Split(str毒理, ":")(2) = 1 Then str毒理分类 = str毒理分类 & IIf(str毒理分类 = "", "'", ",'") & "精神I类'"
        If Split(str毒理, ":")(3) = 1 Then str毒理分类 = str毒理分类 & IIf(str毒理分类 = "", "'", ",'") & "精神II类'"
        
        gstrSQL = "Select distinct id 药品ID" & vbNewLine & _
                "From 收费项目目录 A, 药品规格 B, 药品特性 C,收费执行科室 D" & vbNewLine & _
                "Where a.Id = b.药品id And b.药名id = c.药名id and a.id = d.收费细目id And c.毒理分类 in (" & str毒理分类 & ") and d.执行科室id = [2]"
    ElseIf int盘点类型 = 2 Then '停用药品
        bln盘无库存药品 = True
        bln盘无库存有金额药品 = False
        
        gstrSQL = "Select distinct id 药品id" & vbNewLine & _
                "From 收费项目目录 A, 药品规格 B, 收费执行科室 C" & vbNewLine & _
                "Where a.Id = b.药品id And a.Id = c.收费细目id And Sysdate > 撤档时间 And c.执行科室id = [2]"
'    ElseIf int盘点类型 = 3 Then '无库存记录的药品
'        bln盘无库存药品 = True
'        bln盘无库存有金额药品 = False
'
'        gstrSQL = "Select Distinct ID 药品id" & vbNewLine & _
'                "From 收费项目目录 A, 药品规格 B, 收费执行科室 C" & vbNewLine & _
'                "Where a.Id = b.药品id And a.Id = c.收费细目id And c.执行科室id = [2] And Not Exists" & vbNewLine & _
'                " (Select 1 From 药品库存 D Where a.Id = d.药品id And d.库房id = [2])"
    ElseIf int盘点类型 = 3 Then '无数量但有库存金额的药品
        bln盘无库存药品 = False
        bln盘无库存有金额药品 = True
        
        gstrSQL = "Select Distinct 药品id From 药品库存 Where 性质 = 1 and Nvl(实际数量, 0) = 0 And (Nvl(实际金额, 0) <> 0 or Nvl(实际差价, 0) <> 0) And 库房id = [2]"
    ElseIf int盘点类型 = 4 Then '基本药物
        bln盘无库存药品 = True
        bln盘无库存有金额药品 = False
    
        gstrSQL = "Select Distinct a.药品id" & vbNewLine & _
            "From 药品规格 A, 收费执行科室 B" & vbNewLine & _
            "Where a.药品id = b.收费细目id And a.基本药物 Is Not Null And b.执行科室id = [2]"

    End If
    
    

        
    Set rsPhysic = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取出符合条件的药品]", int近效期天数, lng库房ID)
    
    
    If rsPhysic.RecordCount = 0 Then
        MsgBox "未能找到任何药品数据,请重试！", vbInformation, gstrSysName: Unload Me
        Exit Sub
    End If
    
    DoEvents
    vsfBill.Redraw = flexRDNone
    
    bln库房 = CheckPartProp(lng库房ID)
    With vsfBill
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            '取该药品的详细信息（可能分多个批次）
            
            If int盘点类型 = 0 Then '近效期药品
                Set rsDetail = Get近效期PhysicDetail(lng库房ID, rsPhysic!药品id, int近效期天数, bln已失效)
            Else
                Set rsDetail = GetPhysicDetail(lng库房ID, rsPhysic!药品id, bln盘无库存药品, False, bln盘无库存有金额药品)
            End If
            
            Do While Not rsDetail.EOF
                If (rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1) And .TextMatrix(.rows - 1, 0) <> "" Then .rows = .rows + 1
                '时价药品重算售价
                dbl成本价 = zlStr.Nvl(rsDetail!平均成本价, 0)
                dbl零售价 = zlStr.Nvl(rsDetail!售价, 0)
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(CLng(rsPhysic!药品id), lng库房ID, CLng(rsDetail!批次), 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = rsPhysic!药品id
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = zlStr.Nvl(rsDetail!产地, zlStr.Nvl(rsDetail!缺省产地))
                .TextMatrix(.rows - 1, mconIntCol原产地) = zlStr.Nvl(rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                
                If int盘点类型 = 0 Then '近效期药品
                    If Format(rsDetail!效期, "yyyy-MM-dd") < Format(zlDataBase.Currentdate, "yyyy-MM-dd") Then
                        .Cell(flexcpPicture, .rows - 1, mconIntCol效期, .rows - 1, mconIntCol效期) = imglvw.ListImages(4).Picture
                        .Cell(flexcpPictureAlignment, .rows - 1, mconIntCol效期, .rows - 1, mconIntCol效期) = flexPicAlignLeftCenter
                    End If
                End If
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                .TextMatrix(.rows - 1, mconintCol标志) = "平"
                .TextMatrix(.rows - 1, mconintCol数量差) = "0"
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol帐面数量), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = .TextMatrix(.rows - 1, mconintCol大包装帐面数量)

                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & .TextMatrix(.rows - 1, mconIntCol实盘数量单位小)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If
                
                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                '如果是分批药品，将批次改填为-1，表示新增批次
                .TextMatrix(.rows - 1, mconIntCol批次) = zlStr.Nvl(rsDetail!批次, 0)
                If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(.rows - 1, mconIntCol批次)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol批次) = -1
'                    '调试用，自动为新增批次设置批号与效期
'                    .TextMatrix(.Rows - 1, mconIntCol批号) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntCol效期) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                '如果是停用药品，该行粗体显示
                If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                End If
                
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + Val(.TextMatrix(.rows - 1, mconintCol金额差))) - (zlStr.Nvl(rsDetail!实际差价, 0) + Val(.TextMatrix(.rows - 1, mconintCol差价差))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                .RowData(.rows - 1) = Val(IIf(IsNull(rsDetail!最大效期), 0, rsDetail!最大效期))
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmNewCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
    Else
        vsfBill.Col = mconIntCol药名
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
        vsfBill.EditCell
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get近效期PhysicDetail(ByVal lng库房ID As Long, ByVal lng药品ID As Long, ByVal int近效期天数 As Integer, ByVal bln已失效 As Boolean) As ADODB.Recordset
     '提取该药品当前库房所有批次明细记录
    Dim str单位 As String, str盘点时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql大包装 As String
    Dim strSql小包装 As String
    Dim strSql盘点时间之后发生 As String
    Dim str盘点单NO As String
    Dim strNo串 As String
    Dim i As Integer
    Dim bln当前库存 As Boolean, bln可用数量占用 As Boolean '对应列是否隐藏
    Dim strSql条件 As String

    On Error GoTo ErrHand
    bln当前库存 = vsfBill.colHidden(mconintCol当前库存)
    bln可用数量占用 = vsfBill.colHidden(mconintCol可用数量占用)

    str盘点时间 = txtCheckDate.Caption

    If mintUnit > 0 Then
        Select Case mintUnit
            Case mconint售价单位
                str单位 = ",E.计算单位 As 单位,1 As 比例系数"
            Case mconint门诊单位
                str单位 = ",A.门诊单位 As 单位,A.门诊包装 As 比例系数"
            Case mconint住院单位
                str单位 = ",A.住院单位 As 单位,A.住院包装 As 比例系数"
            Case mconint药库单位
                str单位 = ",A.药库单位 As 单位,A.药库包装 As 比例系数"
        End Select
    Else
        Select Case mint大单位
            Case mconint售价单位
                strSql大包装 = ",E.计算单位 As 大包装单位,1 As 比例系数大"
            Case mconint门诊单位
                strSql大包装 = ",A.门诊单位 As 大包装单位,A.门诊包装 As 比例系数大"
            Case mconint住院单位
                strSql大包装 = ",A.住院单位 As 大包装单位,A.住院包装 As 比例系数大"
            Case mconint药库单位
                strSql大包装 = ",A.药库单位 As 大包装单位,A.药库包装 As 比例系数大"
        End Select

        Select Case mint小单位
            Case mconint售价单位
                strSql小包装 = ",E.计算单位 As 小包装单位,1 As 比例系数小"
            Case mconint门诊单位
                strSql小包装 = ",A.门诊单位 As 小包装单位,A.门诊包装 As 比例系数小"
            Case mconint住院单位
                strSql小包装 = ",A.住院单位 As 小包装单位,A.住院包装 As 比例系数小"
            Case mconint药库单位
                strSql小包装 = ",A.药库单位 As 小包装单位,A.药库包装 As 比例系数小"
        End Select

        str单位 = strSql大包装 & strSql小包装
    End If

    If int近效期天数 <> 0 Then strSql条件 = " (效期 Between Trunc(Sysdate) And (Trunc(Sysdate) + [5])) "
    If bln已失效 Then strSql条件 = strSql条件 & IIf(int近效期天数 = 0, "", "or") & " 效期 < sysdate "
    
    '取药品当前库存及盘点时间以后的净发生额
    gstrSQL = "" & _
        " SELECT DISTINCT A.药品ID,A.成本价 As 平均成本价,E.产地 缺省产地,'[' || E.编码 || ']' As 药品编码, E.名称 As 通用名, C.名称 As 商品名,A.药库分批 AS 分批核算,A.药房分批 AS 药房分批核算,E.是否变价,A.加成率," & _
        "        NVL(B.实际金额,0) 实际金额,NVL(B.实际差价,0) 实际差价,D.现价 售价,NVL(B.批次,0) 批次,A.药品来源,A.基本药物,Decode(b.批号, Null, a.上次批号, b.批号) As 批号,B.效期,F.库房货位,E.规格,decode(b.产地,null,decode(a.上次产地,null,e.产地,a.上次产地),b.产地) as 产地," & _
        "        Decode(b.原产地, Null,a.原产地,b.原产地) As 原产地,B.批准文号,Nvl(B.帐面数量,0) 帐面数量,B.盘点数量,B.可用数量" & str单位 & ",Decode(b.批次, -1, b.成本价, Decode(x.现价, Null, Decode(k.成本价, Null, a.成本价, k.成本价), x.现价)) As 成本价, " & _
        "        Nvl(E.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间,a.最大效期" & IIf(bln当前库存, "", ",nvl(b.当前库存,0) as 当前库存") & IIf(bln可用数量占用, "", " ,nvl(y.可用数量占用,0) 可用数量占用 ") & _
        " FROM (SELECT 库房ID, 药品ID, 批次, SUM (实际数量) AS 帐面数量,SUM (盘点数量) AS 盘点数量,SUM (实际金额) AS 实际金额," & _
        "         SUM (实际差价) AS 实际差价, SUM(可用数量) AS 可用数量,MAX(批号) As 批号, MAX(产地) AS 产地 ,Max(原产地) As 原产地,MAX(效期) AS 效期, MAX(批准文号) As 批准文号, 0 As 成本价 " & IIf(bln当前库存, "", " ,Sum(当前库存) As 当前库存 ") & _
        "         From" & _
        "             ( SELECT A.库房ID,A.药品ID,NVL(批次,0) AS 批次,Nvl(A.实际数量,0) 实际数量,0 盘点数量,Nvl(A.实际金额,0) 实际金额,Nvl(A.实际差价,0) 实际差价,Nvl(A.可用数量,0) 可用数量,A.上次批号 AS 批号,A.上次产地 AS 产地,a.原产地,A.效期,A.批准文号 " & IIf(bln当前库存, "", ",Nvl(a.实际数量, 0) 当前库存") & _
        "             FROM 药品库存 A" & _
        "             Where A.性质 = 1 And A.库房ID=[1] And A.药品ID=[2]  And (Nvl(A.实际数量,0)<>0 Or Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0 ) And (" & strSql条件 & ")" & _
        "     ) GROUP BY 库房ID, 药品ID, 批次 " & _
        ") B, 收费价目 D, 收费项目别名 C, 收费项目目录 E, 药品规格 A," & _
        "      (Select x.药品id,x.库房id,x.批次,nvl(x.现价,0) 现价 from 药品价格记录 x where x.价格类型 = 2 and [3] between x.执行日期 and x.终止日期) X," & _
        IIf(bln可用数量占用, "", "(Select sum(y.实际数量) 可用数量占用 ,y.库房id,y.药品id,y.批次 From 药品收发记录 Y Where y.入出系数 = -1 And y.审核日期 is null and y.填制日期 > (sysdate - 30)  Group By y.库房id, y.药品id, y.批次) Y,") & _
        "      (Select 药品id,批次,平均成本价 成本价 From 药品库存 Where 性质 = 1 And 库房id =[1] ) K,药品储备限额 F " & _
        " Where A.药品ID=E.ID And A.药品ID=B.药品ID" & _
        " AND A.药品ID=F.药品ID(+) And B.药品id=K.药品id(+) And Nvl(B.批次, 0)=nvl(K.批次(+),0) " & _
        " AND b.药品id = x.药品id(+) And b.库房id = x.库房id(+) And Nvl(b.批次, 0) = Nvl(x.批次(+), 0) " & IIf(bln可用数量占用, "", " AND b.药品id = y.药品id(+) And b.库房id = y.库房id(+) And Nvl(b.批次, 0) = Nvl(y.批次(+), 0) ") & _
        " AND A.药品ID=C.收费细目ID(+) AND C.性质(+)=3 AND A.药品ID=D.收费细目ID(+)  " & _
        " AND F.库房ID(+)=[1] And A.药品ID+0=[2] AND D.执行日期(+)<=[3] AND NVL(D.终止日期(+),SYSDATE)>=[3] " & _
        GetPriceClassString("D") & _
        " and e.建档时间<=[3]  Order by 批次 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品当前库房所有批次明细记录]", lng库房ID, lng药品ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")), strNo串, int近效期天数)

    Set Get近效期PhysicDetail = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SearchMisslData(ByVal lng库房ID As Long, ByVal str药品信息 As String)
'功能：获取指定药品
'参数：str药品信息；格式：药品id:批次;药品id:批次...
'
    Dim rsPhysic As ADODB.Recordset '记录该库房的所有药品ID（包括设置存储属性和有库存未设置存储属性的药品）
    Dim rsDetail As ADODB.Recordset
    Dim bln库房 As Boolean
    Dim dbl成本价 As Double, dbl零售价 As Double
    Dim str盘点时间 As String
    Dim str药名 As String
    Dim intMoneyBit As Integer
    Dim bln盘无库存药品 As Boolean
    Dim bln盘无库存有金额药品 As Boolean
    Dim strID批次() As String '药品id:批次
    Dim i As Long
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim dbl盘点金额 As Double
    
    On Error GoTo errHandle
    
    Call FS.ShowFlash("正在装入药品数据,请稍候 ...", Me)
    
    str盘点时间 = txtCheckDate
    strID批次 = Split(str药品信息, ";")
    
    DoEvents
    vsfBill.Redraw = flexRDNone
    
    bln库房 = CheckPartProp(lng库房ID)
    
    With vsfBill
        For i = LBound(strID批次) To UBound(strID批次)
            lng药品ID = Val(Split(strID批次(i), ":")(0))
            lng批次 = Val(Split(strID批次(i), ":")(1))
            
            Set rsDetail = GetPhysicDetail(lng库房ID, lng药品ID, False, False, False)
            rsDetail.Filter = "批次 = " & lng批次
            
            Do While Not rsDetail.EOF
                If .TextMatrix(.rows - 1, 0) <> "" Then .rows = .rows + 1
                '时价药品重算售价
                dbl成本价 = zlStr.Nvl(rsDetail!平均成本价, 0)
                dbl零售价 = zlStr.Nvl(rsDetail!售价, 0)
                If rsDetail!是否变价 = 1 Then
                    dbl零售价 = Get盘点时刻零售价(lng药品ID, lng库房ID, lng批次, 1, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '按常量定义进行格式化
                .TextMatrix(.rows - 1, 0) = lng药品ID
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsDetail!通用名
                Else
                    str药名 = IIf(IsNull(rsDetail!商品名), rsDetail!通用名, rsDetail!商品名)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol药品编码和名称) = rsDetail!药品编码 & str药名
                .TextMatrix(.rows - 1, mconIntCol药品编码) = rsDetail!药品编码
                .TextMatrix(.rows - 1, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品名称)
                Else
                    .TextMatrix(.rows - 1, mconIntCol药名) = .TextMatrix(.rows - 1, mconIntCol药品编码和名称)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsDetail!商品名), "", rsDetail!商品名)
                
                .TextMatrix(.rows - 1, mconIntCol来源) = zlStr.Nvl(rsDetail!药品来源)
                .TextMatrix(.rows - 1, mconIntCol基本药物) = zlStr.Nvl(rsDetail!基本药物)
                .TextMatrix(.rows - 1, mconIntCol规格) = IIf(IsNull(rsDetail!规格), "", rsDetail!规格)
                .TextMatrix(.rows - 1, mconIntCol产地) = zlStr.Nvl(rsDetail!产地, zlStr.Nvl(rsDetail!缺省产地))
                .TextMatrix(.rows - 1, mconIntCol原产地) = zlStr.Nvl(rsDetail!原产地)
                .TextMatrix(.rows - 1, mconIntCol库房货位) = IIf(IsNull(rsDetail!库房货位), "", rsDetail!库房货位)
                .TextMatrix(.rows - 1, mconIntCol批号) = IIf(IsNull(rsDetail!批号), "", rsDetail!批号)
                .TextMatrix(.rows - 1, mconIntCol效期) = IIf(IsNull(rsDetail!效期), "", Format(rsDetail!效期, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(.rows - 1, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(.rows - 1, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol批准文号) = IIf(IsNull(rsDetail!批准文号), "", rsDetail!批准文号)
                .TextMatrix(.rows - 1, mconIntCol实际金额) = zlStr.Nvl(rsDetail!实际金额, 0)
                .TextMatrix(.rows - 1, mconIntCol实际差价) = zlStr.Nvl(rsDetail!实际差价, 0)
                .TextMatrix(.rows - 1, mconIntcol加成率) = rsDetail!加成率 / 100 & "||" & rsDetail!是否变价 & "||" & rsDetail!药房分批核算
                .TextMatrix(.rows - 1, mconintCol标志) = "平"
                .TextMatrix(.rows - 1, mconintCol数量差) = "0"
                .TextMatrix(.rows - 1, mconintCol库存数量) = zlStr.Nvl(rsDetail!帐面数量, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol单位) = IIf(IsNull(rsDetail!单位), "", rsDetail!单位)
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol比例系数) = zlStr.Nvl(rsDetail!比例系数, 0)
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol帐面数量), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数, mintNumberDigit, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then .TextMatrix(.rows - 1, mconintCol当前库存) = zlStr.FormatEx(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    If Not .colHidden(mconintCol可用数量占用) Then .TextMatrix(.rows - 1, mconintCol可用数量占用) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数, mintNumberDigit, , True) & rsDetail!单位
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol售价) = zlStr.FormatEx(dbl零售价 * rsDetail!比例系数小, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol帐面数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol比例系数大) = zlStr.Nvl(rsDetail!比例系数大, 0)
                    .TextMatrix(.rows - 1, mconIntCol比例系数小) = zlStr.Nvl(rsDetail!比例系数小, 0)
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol帐面数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位大) = rsDetail!大包装单位
                    .TextMatrix(.rows - 1, mconIntCol实盘数量单位小) = rsDetail!小包装单位
                    .TextMatrix(.rows - 1, mconintCol大包装帐面数量) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数大), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol大包装实盘数量) = .TextMatrix(.rows - 1, mconintCol大包装帐面数量)

                    .TextMatrix(.rows - 1, mconintCol小包装帐面数量) = zlStr.FormatEx((Val(rsDetail!帐面数量) - Val(.TextMatrix(.rows - 1, mconintCol大包装帐面数量)) * Val(rsDetail!比例系数大)) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol小包装实盘数量) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol小包装帐面数量), mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconintCol实盘数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!帐面数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol合计) = .TextMatrix(.rows - 1, mconintCol实盘数量) & .TextMatrix(.rows - 1, mconIntCol实盘数量单位小)
                    .TextMatrix(.rows - 1, mconintCol盘点金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) * Val(.TextMatrix(.rows - 1, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsDetail!可用数量, 0) / rsDetail!比例系数小, mintNumberDigit0, , True)
                    
                    If Not .colHidden(mconintCol当前库存) Then
                        Dim dbl当前库存大 As Double, dbl当前库存小 As Double
                        dbl当前库存大 = Fix(zlStr.Nvl(rsDetail!当前库存, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = IIf(dbl当前库存大 = 0 And rsDetail!当前库存 < 0, "-", "") & zlStr.FormatEx(dbl当前库存大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl当前库存小 = Abs((Val(zlStr.Nvl(rsDetail!当前库存, 0)) - dbl当前库存大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol当前库存) = .TextMatrix(.rows - 1, mconintCol当前库存) & zlStr.FormatEx(dbl当前库存小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    If Not .colHidden(mconintCol可用数量占用) Then
                        Dim dbl数量占用大 As Double, dbl数量占用小 As Double
                        dbl数量占用大 = Fix(zlStr.Nvl(rsDetail!可用数量占用, 0) / rsDetail!比例系数大)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = IIf(dbl数量占用大 = 0 And rsDetail!可用数量占用 < 0, "-", "") & zlStr.FormatEx(dbl数量占用大, mintNumberDigit0, , True) & rsDetail!大包装单位
                        dbl数量占用小 = Abs((Val(zlStr.Nvl(rsDetail!可用数量占用, 0)) - dbl数量占用大 * Val(rsDetail!比例系数大)) / rsDetail!比例系数小)
                        .TextMatrix(.rows - 1, mconintCol可用数量占用) = .TextMatrix(.rows - 1, mconintCol可用数量占用) & zlStr.FormatEx(dbl数量占用小, mintNumberDigit0, , True) & rsDetail!小包装单位
                    End If
                    
                    .TextMatrix(.rows - 1, mconintCol成本价) = zlStr.FormatEx(zlStr.Nvl(rsDetail!成本价, 0) * rsDetail!比例系数小, mintCostDigit0, , True)
                End If
                
                '实盘数量为0与库存数量比较判断盈亏
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 And Val(.TextMatrix(.rows - 1, mconintCol库存数量)) > 0 Then
                    .TextMatrix(.rows - 1, mconintCol标志) = "亏"
                End If
                
                If Not .colHidden(mconintCol当前库存) Then
                    .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = False
                    If zlStr.Nvl(rsDetail!当前库存, 0) <> zlStr.Nvl(rsDetail!帐面数量, 0) Then .Cell(flexcpFontBold, .rows - 1, mconintCol当前库存, .rows - 1, mconintCol当前库存) = True
                End If
                If Not .colHidden(mconintCol可用数量占用) Then
                    If zlStr.Nvl(rsDetail!可用数量占用, 0) <> 0 Then .Cell(flexcpFontBold, .rows - 1, mconintCol可用数量占用, .rows - 1, mconintCol可用数量占用) = True
                End If
                
                '如果是分批药品，将批次改填为-1，表示新增批次
                .TextMatrix(.rows - 1, mconIntCol批次) = zlStr.Nvl(rsDetail!批次, 0)
                If CheckPhysicBatch(bln库房, rsDetail!分批核算, rsDetail!药房分批核算) And Val(.TextMatrix(.rows - 1, mconIntCol批次)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol批次) = -1
'                    '调试用，自动为新增批次设置批号与效期
'                    .TextMatrix(.Rows - 1, mconIntCol批号) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntCol效期) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol售价)) = Val(.TextMatrix(.rows - 1, mconintCol成本价))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                dbl盘点金额 = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol售价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit, , True) '当前售价*实盘数量 = 盘点金额
                .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(dbl盘点金额 - Val(.TextMatrix(.rows - 1, mconIntCol实际金额)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol售价)) - Val(.TextMatrix(.rows - 1, mconintCol成本价))) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)) - Val(.TextMatrix(.rows - 1, mconIntCol实际差价)), intMoneyBit, , True)
                
                If .TextMatrix(.rows - 1, mconintCol标志) = "亏" Then
                    .TextMatrix(.rows - 1, mconintCol金额差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol金额差)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol差价差) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                End If
                
                '如果是停用药品，该行粗体显示
                If Format(rsDetail!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                End If
                
                '.TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol成本价)) * Val(.TextMatrix(.rows - 1, mconintCol实盘数量)), mintMoneyDigit)
                '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                .TextMatrix(.rows - 1, mconintCol盘点成本金额) = zlStr.FormatEx((zlStr.Nvl(rsDetail!实际金额, 0) + Val(.TextMatrix(.rows - 1, mconintCol金额差))) - (zlStr.Nvl(rsDetail!实际差价, 0) + Val(.TextMatrix(.rows - 1, mconintCol差价差))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol盘点成本金额差) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol金额差)) - Val(.TextMatrix(.rows - 1, mconintCol差价差)), intMoneyBit, , True)
                
                '盘亏盘盈行用颜色区分
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '设置分批属性
                Call Get药品分批属性(.rows - 1)
                
                .RowData(.rows - 1) = Val(IIf(IsNull(rsDetail!最大效期), 0, rsDetail!最大效期))
                
                rsDetail.MoveNext
                
            Loop
            
            Call zlControl.StaShowPercent((i + 1) / (UBound(strID批次) + 1), staThis.Panels(2), frmNewCheckCard) '进度条
        Next
        
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintCol实盘数量, .rows - 1, mconintCol实盘数量) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol大包装实盘数量, .rows - 1, mconintCol大包装实盘数量) = True
            .Cell(flexcpFontBold, 1, mconintCol小包装实盘数量, .rows - 1, mconintCol小包装实盘数量) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
        
    End With
    
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol大包装实盘数量, mconintCol实盘数量)
    Else
        vsfBill.Col = mconIntCol药名
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
        vsfBill.EditCell
    End If
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.上次产地 as 生产商, t.原产地 as 原产地 From 药品规格 T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng生产商长度 = rsTmp.Fields("生产商").DefinedSize
    mlng原产地长度 = rsTmp.Fields("原产地").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
