VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSelfMakeCard 
   AutoRedraw      =   -1  'True
   Caption         =   "卫材自制入库单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmSelfMakeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   5970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   7
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   8
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   11655
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VSFlex8Ctl.VSFlexGrid vs组成材料 
         Height          =   2220
         Left            =   210
         TabIndex        =   30
         Top             =   2610
         Width           =   11145
         _cx             =   19659
         _cy             =   3916
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelfMakeCard.frx":014A
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
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   29
         Top             =   180
         Width           =   1410
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDrug 
         Height          =   2235
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   3942
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   1230
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2170
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4920
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   26
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:asdfasdfasdfsadfsadfsdfasdfsadfasdfsdf"
         Height          =   180
         Left            =   2040
         TabIndex        =   25
         Top             =   2280
         Width           =   4590
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7230
         TabIndex        =   22
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   21
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   20
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   19
         Top             =   5280
         Width           =   915
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
         Left            =   9480
         TabIndex        =   18
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   4995
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料自制入库单"
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
         TabIndex        =   17
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   16
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   15
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   6645
         TabIndex        =   14
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   8520
         TabIndex        =   13
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制剂室(&T)"
         Height          =   180
         Left            =   8220
         TabIndex        =   2
         Top             =   660
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
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
            Picture         =   "frmSelfMakeCard.frx":02AA
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":04C4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":06DE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":08F8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0B12
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0D2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0F46
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1160
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
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
            Picture         =   "frmSelfMakeCard.frx":137A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1594
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":17AE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":19C8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1BE2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1DFC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":2016
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":2230
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSelfMakeCard.frx":244A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSelfMakeCard.frx":2CDE
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSelfMakeCard.frx":31E0
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label lblCode 
      Caption         =   "材料"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmSelfMakeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln单据增加 As Boolean
Private mblnFirst As Boolean
Private mintUnit  As Integer                '0-散装单位,1-包装单位
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mint库存检查 As Integer             '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Dim mstrPrivs As String                     '权限
Private mintBatchNoLen As Integer           '数据库中批号定义长度
'刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材自制入库单"
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
Private Const mlngModule = 1713

'----------------------------------------------------------------------------------------------------------

Private mcolUseCount As Collection

'=========================================================================================

Private Const mconIntCol材料 As Integer = 1
Private Const mconIntCol规格 As Integer = 2
Private Const mconIntCol原销期 As Integer = 3
Private Const mconIntCol比例系数 As Integer = 4
Private Const mconIntCol单位 As Integer = 5
Private Const mconIntCol批号 As Integer = 6
Private Const mconIntCol效期 As Integer = 7
Private Const mconIntCol一次性材料 As Integer = 8
Private Const mconIntCol灭菌效期 As Integer = 9
Private Const mconIntCol灭菌日期   As Integer = 10
Private Const mconIntCol灭菌失效期 As Integer = 11

Private Const mconIntCol数量 As Integer = 12
Private Const mconIntCol采购价 As Integer = 13
Private Const mconIntCol采购金额 As Integer = 14
Private Const mconIntCol售价 As Integer = 15
Private Const mconIntCol售价金额 As Integer = 16
Private Const mconintCol差价 As Integer = 17


Private Const mconIntColS As Integer = 18       '总列数
'=========================================================================================
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSQL As String
    GetDepend = False
    
    On Error GoTo ErrHandle
    strSQL = "" & _
        "   SELECT B.Id,b.名称 " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID " & _
        "       AND A.单据 = 31  and b.系数=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-检查入类别")
    If rsTemp.EOF Then
        MsgBox "没有设置卫材自制入库的入库类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = "" & _
        "   SELECT B.Id,b.名称 " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID " & _
        "           AND A.单据 = 31  and b.系数=-1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-检查出类别")
    If rsTemp.EOF Then
        MsgBox "没有设置卫材自制入库的出库类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 " & _
        "           AND b.编码 ='K'" & _
        "           AND a.id = c.部门id " & _
        "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-检查制剂室")
    If rsTemp.EOF Then
        MsgBox "部门性质中没有性质为制剂室的部门,请查看部门管理！", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = " SELECT a.自制材料id FROM 自制材料构成 a, 材料特性 b Where a.自制材料id = b.材料id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-获取自制卫材构成")
    If rsTemp.EOF Then
        MsgBox "没有一种具有原料卫材组成的自制卫材,请查看卫生卫材目录管理！", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
        Optional int记录状态 As Integer = 1, Optional strPrivs As String, _
        Optional blnSuccess As Boolean = False)
        
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:显示或编辑卡片,是唯一入库
    '--入参数:
    '--出参数:
    '--返  回:blnSuccess
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
        
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain

    Call GetRegInFor(g私有模块, "卫材自制入库管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
        txtNo = mstr单据号
        txtNo.Tag = txtNo
        txtNo.Locked = True
        txtNo.TabStop = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
        txtNo.Locked = True
        txtNo.TabStop = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    If Not GetDepend Then Exit Sub
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
Private Sub cboStock_Click()
    mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变库房，有可能要改变相应卫材的单位，" & vbCrLf & "且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫材单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                            
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With mshBill
        .SetFocus
        .Row = 1
        .Col = mconIntCol材料
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'查找
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mconIntCol材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
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
    End Select
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mconIntCol材料, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
        
    If mint编辑状态 = 3 Then        '审核
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub

        '刘兴宏:2007/06/10:问题10813
        mstrTime_End = GetBillInfo(16, txtNo.Tag)
        If mstrTime_End = "" Then
            MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("注意:" & vbCrLf & "  该单据已经被其他操作员编辑，不能继续!" & vbCrLf & "  是否重新刷新单据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
                
        
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
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
    
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
            
        strReg = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0)
        
        If Val(strReg) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    If txtNo.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
    
    mblnSave = False
'    mblnEdit = True
    mshBill.ClearBill
    vs组成材料.Clear (1)
    vs组成材料.Rows = 2
    Call 显示合计金额
'    SetEdit
    txt摘要.Text = ""
    cboType.SetFocus
    mblnChange = False
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    mblnFirst = True
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mintUnit = Val(strReg)
    
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
        
    
    mintBatchNoLen = GetBatchNoLen()
    
    txtNo = mstr单据号
    txtNo.Tag = txtNo.Text
    With cboType
    
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.名称 " & _
            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "   Where c.工作性质 = b.名称 " & _
            "           AND b.编码 ='K'" & _
            "           AND a.id = c.部门id " & _
            "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption)
        
        If rsTemp.EOF Then Exit Sub
        
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    Call initCard
    
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mconIntCol采购价) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconIntCol采购金额) = IIf(mblnCostView = True, 1200, 0)
        .ColWidth(mconintCol差价) = IIf(mblnCostView = True, 1200, 0)
    End With
    With vs组成材料
        .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 1200, 0)
        .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1200, 0)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim str换算系数 As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim strReg As String
    
    On Error GoTo ErrHandle
    strReg = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    
    strOrder = strReg
    
    '库房
    strCompare = Mid(strOrder, 1, 1)
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
                
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 16 and a.no=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsTemp!名称
                    .ItemData(.NewIndex) = rsTemp!Id
                    .ListIndex = 0
                End With
                rsTemp.Close
            End If
            
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.计算单位 AS 单位,(A.填写数量) AS 数量,1 as 比例系数,"
                    str换算系数 = "1"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.填写数量 / B.换算系数) AS 数量,B.换算系数 as 比例系数, "
                    str换算系数 = "B.换算系数"
            End Select
            
            gstrSQL = "" & _
                "   SELECT * " & _
                "   FROM (  SELECT DISTINCT 序号,a.药品id as 材料id, ('[' || c.编码 || ']' || c.名称) AS 材料信息,c.规格,a.产地, a.批号, a.效期," & _
                "                   zlSpellCode(c.名称) 名称," & strUnitQuantity & _
                "                   (a.成本价*" & str换算系数 & ") AS 成本价," & _
                "                   a.成本金额 ,(a.零售价*" & str换算系数 & ") AS 零售价,a.零售金额 AS 零售金额," & _
                "                   a.差价 AS 差价,a.填制人,a.填制日期,a.审核人,a.审核日期,a.摘要,c.产地 as 原产地,b.最大效期,b.一次性材料,b.灭菌效期," & _
                "                   a.灭菌日期,a.灭菌效期 as 灭菌失效期,a.对方部门id,c.是否变价,b.指导差价率/100 as 指导差价率,b.在用分批 " & _
                "           FROM 药品收发记录 a, 材料特性 b,收费项目目录 c" & _
                "           Where a.药品id = b.材料id and a.药品id=c.id " & _
                "                   AND a.记录状态 = [2]" & _
                "                   AND a.单据 = 16 AND 入出系数=1 " & _
                "                   AND a.no =[1]  " & _
                "           )" & _
                " ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "材料信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mint记录状态)
                
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            '刘兴宏:2007/06/10:问题10813
            mstrTime_Start = GetBillInfo(16, mstr单据号)
            
            Txt填制人 = rsTemp!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户名
            End If
            
            Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd hh:mm:ss")
            
            Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
            Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!对方部门id Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = rsTemp.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp!材料ID
                    .TextMatrix(intRow, mconIntCol材料) = rsTemp!材料信息
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)

                    .TextMatrix(intRow, mconIntCol单位) = rsTemp!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol一次性材料) = zlStr.Nvl(rsTemp!一次性材料)
                    .TextMatrix(intRow, mconIntCol灭菌效期) = zlStr.Nvl(rsTemp!灭菌效期)
                    .TextMatrix(intRow, mconIntCol灭菌日期) = IIf(IsNull(rsTemp!灭菌日期), "", Format(rsTemp!灭菌日期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol灭菌失效期) = IIf(IsNull(rsTemp!灭菌失效期), "", Format(rsTemp!灭菌失效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol数量) = Format(zlStr.Nvl(rsTemp!数量, 0), mFMT.FM_数量)
                    .TextMatrix(intRow, mconIntCol采购价) = Format(zlStr.Nvl(rsTemp!成本价, 0), mFMT.FM_成本价)
                    .TextMatrix(intRow, mconIntCol采购金额) = Format(zlStr.Nvl(rsTemp!成本金额, 0), mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol售价) = Format(zlStr.Nvl(rsTemp!零售价, 0), mFMT.FM_零售价)
                    .TextMatrix(intRow, mconIntCol售价金额) = Format(zlStr.Nvl(rsTemp!零售金额, 0), mFMT.FM_金额)
                    .TextMatrix(intRow, mconintCol差价) = Format(zlStr.Nvl(rsTemp!差价, 0), mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(rsTemp!最大效期), "0", rsTemp!最大效期) & "||" & rsTemp!指导差价率 & "||" & rsTemp!是否变价 & "||" & rsTemp!在用分批
                    .TextMatrix(intRow, mconIntCol比例系数) = rsTemp!比例系数
                    rsTemp.MoveNext
                Loop
                
                Dim dblCostPrice As Double
                If .TextMatrix(1, 0) <> "" Then
                    Call Set组成材料(Val(.TextMatrix(1, 0)), Val(.TextMatrix(1, mconIntCol数量) * .TextMatrix(1, mconIntCol比例系数)), False, dblCostPrice)
                End If
                
            End With
            rsTemp.Close
                 
    End Select
    SetEdit         '设置编辑属性
    Call 显示合计金额
    
    If mint编辑状态 = 2 And mint库存检查 <> 0 Then
        SetUseCountCol
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'设置修改前原料药的使用数量，以便于在修改过程中对库存数量的判断更准确
Private Sub SetUseCountCol()
    Dim rsTemp As New Recordset
    Dim numUsedCount As Double
    Dim vardrug As Variant
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select 药品id as 材料id,填写数量,费用id,批次 " & _
        "   From 药品收发记录 " & _
        "   Where no=[1] and 单据=16 and 记录状态=1 and 入出系数=-1 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
    
    If rsTemp.EOF Then Exit Sub
    
    Set mcolUseCount = New Collection
    With mcolUseCount
        Do While Not rsTemp.EOF
            numUsedCount = 0
            For Each vardrug In mcolUseCount
                If vardrug(0) = zlStr.Nvl(rsTemp!费用ID) & "!" & zlStr.Nvl(rsTemp!材料ID) & "!" & Val(zlStr.Nvl(rsTemp!批次)) Then
                    numUsedCount = vardrug(1)
                    .Remove vardrug(0)
                    Exit For
                End If
            Next
            .Add Array(zlStr.Nvl(rsTemp!费用ID) & "!" & zlStr.Nvl(rsTemp!材料ID) & "!" & Val(zlStr.Nvl(rsTemp!批次)), Val(rsTemp!填写数量)), zlStr.Nvl(rsTemp!费用ID) & "!" & zlStr.Nvl(rsTemp!材料ID) & "!" & Val(zlStr.Nvl(rsTemp!批次))
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboType.Enabled = False
            txt摘要.Enabled = False
        Else
            .ColData(0) = 5
            .ColData(mconIntCol材料) = 1
            .ColData(mconIntCol规格) = 5
            
            .ColData(mconIntCol单位) = 5
            .ColData(mconIntCol批号) = 4
            .ColData(mconIntCol效期) = 5
            .ColData(mconIntCol灭菌日期) = 2
            .ColData(mconIntCol灭菌效期) = 5
            .ColData(mconIntCol数量) = 4
            .ColData(mconIntCol采购价) = 5
            .ColData(mconIntCol采购金额) = 5
            .ColData(mconIntCol售价) = 5
            .ColData(mconIntCol售价金额) = 5
            .ColData(mconintCol差价) = 5
            
            
            .ColData(mconIntCol原销期) = 5
            .ColData(mconIntCol比例系数) = 5
            
            .ColAlignment(mconIntCol材料) = flexAlignLeftCenter
            .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
            
            .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
            .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
            .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol数量) = flexAlignRightCenter
            .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
            .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
            .ColAlignment(mconintCol差价) = flexAlignRightCenter
            
            cboStock.Enabled = True
            cboType.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol材料) = "名称与编码"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = "失效期"

        .TextMatrix(0, mconIntCol一次性材料) = "一次性材料"
        .TextMatrix(0, mconIntCol灭菌效期) = "灭菌效期"
        .TextMatrix(0, mconIntCol灭菌日期) = "灭菌日期"
        .TextMatrix(0, mconIntCol灭菌失效期) = "灭菌失效期"
                
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol采购价) = "成本价"
        .TextMatrix(0, mconIntCol采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        
        .TextMatrix(0, mconIntCol原销期) = "原效期"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        
        
        .TextMatrix(1, 0) = ""
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol材料) = 2000
        .ColWidth(mconIntCol规格) = 900
        
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        
        
        .ColWidth(mconIntCol一次性材料) = 0
        .ColWidth(mconIntCol灭菌效期) = 0
        .ColWidth(mconIntCol灭菌日期) = 1000
        .ColWidth(mconIntCol灭菌失效期) = 1000
                
        .ColWidth(mconIntCol数量) = 800
        .ColWidth(mconIntCol采购价) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol采购金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol售价) = 900
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = IIf(mblnCostView = False, 0, 800)
        
        
        
        .ColWidth(mconIntCol原销期) = 0
        .ColWidth(mconIntCol比例系数) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol材料) = 1
        .ColData(mconIntCol规格) = 5
        
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 4
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol一次性材料) = 5
        .ColData(mconIntCol灭菌效期) = 5
        .ColData(mconIntCol灭菌日期) = 2
        .ColData(mconIntCol灭菌失效期) = 5
        
        .ColData(mconIntCol数量) = 4
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 0
        
        
        .ColData(mconIntCol原销期) = 5
        .ColData(mconIntCol比例系数) = 5
        
        .ColAlignment(mconIntCol材料) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol一次性材料) = flexAlignCenterCenter
        .ColAlignment(mconIntCol灭菌效期) = flexAlignCenterCenter
        .ColAlignment(mconIntCol灭菌日期) = flexAlignCenterCenter
        .ColAlignment(mconIntCol灭菌失效期) = flexAlignCenterCenter
        
             
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol材料
        .LocateCol = mconIntCol材料
    End With
    
    With vs组成材料
        .RowHeight(0) = .RowHeight(0) * 2
        .ExplorerBar = .ExplorerBar + &H1000&
    End With
    txt摘要.MaxLength = sys.FieldsLength("药品收发记录", "摘要")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
    
    LblType.Left = cboType.Left - LblType.Width - 100
    
    
    With Lbl填制人
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
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
    
    With Txt审核日期
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
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
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
    End With
    
    With vs组成材料
        .Left = mshBill.Left
        .Width = mshBill.Width
        .Top = txt摘要.Top - 60 - .Height
    End With
            
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = vs组成材料.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mshDrug.Visible Then
        mshDrug.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    
End Sub

Private Function SaveCheck() As Boolean
    mblnSave = False
    SaveCheck = False
    gstrSQL = "zl_自制材料入库_verify('" & txtNo.Tag & "','" & UserInfo.用户名 & "')"
    On Error GoTo ErrHandle
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mshBill_AfterDeleteRow()
    With mshBill
        If .Row > 1 Then
            .Row = .Row - 1
        Else
            .Row = 1
        End If
        If .TextMatrix(.Row, 0) = "" Then
            vs组成材料.Clear (1)
        Else
            Dim dblCostPrice As Double
            Call Set组成材料(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol比例系数)), False, dblCostPrice)
        End If
        
    End With
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行卫材？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim rsTemp As New Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intStockID As Long
    Dim strUnitQuantity As String
    
    On Error GoTo ErrHandle
    Select Case mintUnit
        Case 0
            strUnitQuantity = "D.计算单位 AS 单位, (to_char(s.库存数量 ," & mOraFMT.FM_数量 & ")) AS 数量,1 as 比例系数," _
                & "to_char(p.售价," & mOraFMT.FM_零售价 & ") as 售价,"
        Case Else
            strUnitQuantity = "d.包装单位 AS 单位, (to_char(s.库存数量 / d.换算系数," & mOraFMT.FM_数量 & ")) AS 数量,d.换算系数 as 比例系数," _
                & "to_char(p.售价*d.换算系数," & mOraFMT.FM_零售价 & ") as 售价, "
    End Select
        
    intStockID = cboStock.ItemData(cboStock.ListIndex)
    
    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50
    
    
    '卫材条件
    gstrSQL = "" & _
        "   SELECT  D.编码, D.名称,D.规格, d.产地,d.材料id as 药品id, " & strUnitQuantity & "  s.库存金额, d.最大效期,d.是否变价,d.指导差价率/100 as 指导差价率,d.在用分批,e.库房货位 " & _
        "   FROM  ( SELECT DISTINCT L.编码, L.名称,L.规格, L.产地, d.材料id, L.计算单位,NVL (TO_CHAR (d.最大效期, '9999990'), 0) 最大效期," & _
        "                   d.包装单位,TO_CHAR (d.换算系数, " & GFM_XS & ") 换算系数,l.是否变价,d.指导差价率,d.在用分批 " & _
        "           FROM 自制材料构成 f,  材料特性 d,收费项目目录 L,收费执行科室 R" & _
        "           Where f.自制材料id = d.材料id  and F.自制材料id=L.id And (L.站点=[2] or L.站点 is null) AND nvl(d.自制材料,0)=1 and f.自制材料ID=R.收费细目id and R.执行科室ID=[1]" & _
        "                   AND (   EXISTS (SELECT 1 From 部门性质说明 WHERE  工作性质 In ('卫材库','制剂室', '虚拟库房') AND 部门id = [1]) " & _
        "                           OR L.服务对象 =(SELECT distinct '1' From 部门性质说明 WHERE 工作性质 LIKE '发料部门' AND 部门id =[1] AND 服务对象 IN (1, 3)) " & _
        "                           OR L.服务对象 =(SELECT distinct '2' From 部门性质说明 WHERE 工作性质 LIKE '发料部门' AND 部门id =[1] AND 服务对象 IN (2, 3))) " & _
        "                   AND (L.撤档时间 IS NULL OR TO_CHAR (L.撤档时间, 'yyyy-MM-dd') = '3000-01-01') " & _
        "           ) d,"
        
    '连收费价目，主要找售价
    gstrSQL = gstrSQL & _
        "   (   SELECT 收费细目id, TO_CHAR (现价, '999999999990.9999') 售价 " & _
        "       From 收费价目 " & _
        "       WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (    SYSDATE >= 执行日期 AND 终止日期 IS NULL)) " & _
        GetPriceClassString("") & _
        "   ) p,"
             
    
    '连药品库存
    gstrSQL = gstrSQL & _
        "   (   SELECT 药品id, TO_CHAR (SUM (可用数量), " & mOraFMT.FM_数量 & ") 可用数量,TO_CHAR (SUM (实际数量)," & mOraFMT.FM_数量 & ") 库存数量,TO_CHAR (SUM (实际金额), " & mOraFMT.FM_金额 & ") 库存金额 " & _
        "       From 药品库存 " & _
        "       Where 库房id =[1]  and 性质=1 " & _
        "       GROUP BY 药品id) s, "
    
    gstrSQL = gstrSQL & _
        "   (   Select 材料ID,库房ID,库房货位 From 材料储备限额 " & _
        "       Where 库房ID=[1]) E ,收费项目目录 M"
        
    '整个条件
    gstrSQL = gstrSQL & _
        "   Where d.材料id = p.收费细目id  And D.材料id=M.id And (M.站点=[2] or M.站点 is null) And M.是否变价<>1 AND d.材料id = s.药品id (+) and D.材料ID=E.材料ID(+)" & _
        "   ORDER BY d.编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, intStockID, gstrNodeNo)
           
    If rsTemp.EOF Then
        ShowMsgBox "不存在自制材料,可能未设置存储库房或未设置自制材料,请在[卫材目录管理]中设置!"
        Exit Sub
    End If
    
    Set mshDrug.Recordset = rsTemp
    rsTemp.Close
    Call SetDrugWidth(sngLeft, sngTop)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'设置卫材选择器的宽度及相关属性
Private Sub SetDrugWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    With mshDrug
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
        If RestoreFlexState(mshDrug, mstrCaption) = False Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 0
            
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(.Cols - 1) = 1500
        End If
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol数量 Or .Col = mconIntCol采购价 Or .Col = mconIntCol售价 Or .Col = mconIntCol采购金额 Or .Col = mconIntCol售价金额 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol数量
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
                Case mconIntCol采购价, mconIntCol售价
                   intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.成本价小数, g_小数位数.obj_散装小数.成本价小数)
                Case mconIntCol采购金额, mconIntCol售价金额
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.金额小数, g_小数位数.obj_散装小数.金额小数)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            SetInputFormat .Row
            Dim dblCostPrice As Double
            
            If .TextMatrix(.Row, 0) <> "" Then
                Call Set组成材料(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol比例系数)), False, dblCostPrice)
            Else
                vs组成材料.Rows = 2
                vs组成材料.Clear (1)
            End If
                
        End If
        
        Select Case .Col
            Case mconIntCol材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                            
            Case mconIntCol批号
                .TxtCheck = False
                '.TextMask = "1234567890"
                .MaxLength = mintBatchNoLen
            
            Case mconIntCol效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol批号) <> "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) And .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0) <> "0" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                                 Call CheckLapse(.TextMatrix(.Row, mconIntCol效期))
                            End If
                        End If
                    End If
                End If
            Case mconIntCol采购价
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol采购金额
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim strSerach As String
    Dim strLike As String
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        With mshBill
            .Text = Trim(.Text)
            strKey = Trim(.Text)
            
            If Mid(strKey, 1, 1) = "[" Then
                If InStr(2, strKey, "]") <> 0 Then
                    strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                Else
                    strKey = Mid(strKey, 2)
                End If
            End If
            Select Case .Col
                
                Case mconIntCol材料
                    If strKey <> "" Then
                        Dim rsTemp As New Recordset
                        Dim sngLeft As Single
                        Dim sngTop As Single
                        Dim intStockID As Long
                        
                        Select Case mintUnit
                            Case 0
                                strUnitQuantity = "d.计算单位 AS 单位, (to_char(s.库存数量 ," & mOraFMT.FM_数量 & ")) AS 数量,1 as 比例系数," & _
                                    "   to_char(p.售价," & mOraFMT.FM_零售价 & ") as 售价,"
                            Case 1
                                strUnitQuantity = "d.包装单位 AS 单位, (to_char(s.库存数量 / d.换算系数," & mOraFMT.FM_数量 & ")) AS 数量,d.换算系数 as 比例系数," _
                                    & "to_char(p.售价*d.换算系数," & mOraFMT.FM_零售价 & ") as 售价, "
                        End Select
                            
                        intStockID = cboStock.ItemData(cboStock.ListIndex)
                        
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50
                        
                        strSerach = " And (A.编码 Like [2] OR B.名称 Like [2] OR B.简码 LIKE [2])"
                        
                        If IsNumeric(strKey) Then                         '如果是数字,则只取编码
                            If Mid(gSystem_Para.Para_输入方式, 1, 1) = "1" Then strSerach = " And (A.编码 Like [2])"
                            strLike = "" & GetMatchingSting(UCase(strKey)) & ""
                        ElseIf zlStr.IsCharAlpha(strKey) Then          '输入全是字母时只匹配简码
                            If Mid(gSystem_Para.Para_输入方式, 2, 1) = "1" Then strSerach = " And B.简码 Like [2] "
                            strLike = "" & GetMatchingSting(UCase(strKey)) & ""
                        ElseIf zlStr.IsCharChinese(strKey) Then
                            strSerach = " And B.名称 Like [2] "
                            strLike = "" & GetMatchingSting(strKey) & ""
                        End If
                        
                        
                        '卫材条件
                          gstrSQL = "" & _
                              "   SELECT  d.编码, d.名称,d.规格, d.产地,d.材料id , " & strUnitQuantity & "  s.库存金额, d.最大效期,d.是否变价,d.指导差价率/100 as 指导差价率,d.在用分批,e.库房货位 " & _
                              "   FROM  ( SELECT DISTINCT l.编码, l.名称,l.规格, L.产地, d.材料id, l.计算单位,NVL (TO_CHAR (d.最大效期, '9999990'), 0) 最大效期," & _
                              "                   d.包装单位,TO_CHAR (d.换算系数, " & GFM_XS & ") 换算系数,l.是否变价,d.指导差价率,d.在用分批 " & _
                              "           FROM 自制材料构成 f,  材料特性 d, " & _
                              "                    (  Select A.ID,A.编码,A.名称,A.规格,A.产地,A.计算单位,A.服务对象,a.是否变价  " & _
                              "                       From 收费项目目录 A,收费项目别名 B  " & _
                              "                       Where A.ID=B.收费细目ID And (A.站点=[4] or A.站点 is null) " & _
                              "                           And B.码类=[5] " & _
                              "                           AND A.类别 ='4' And (A.撤档时间 is null Or A.撤档时间>=[3]) " & strSerach & ")  L," & _
                              "                 收费执行科室 R" & _
                              "           Where f.自制材料id = d.材料id and f.自制材料id=l.id AND nvl(d.自制材料,0)=1 and F.自制材料ID=R.收费细目ID and R.执行科室ID=[1]" & _
                              "                   AND (   EXISTS (SELECT 1 From 部门性质说明 WHERE  工作性质 In ('卫材库','制剂室', '虚拟库房') AND 部门id = [1]) " & _
                              "                           OR L.服务对象 =(SELECT distinct '1' From 部门性质说明 WHERE 工作性质 LIKE '发料部门' AND 部门id =[1] AND 服务对象 IN (1, 3)) " & _
                              "                           OR l.服务对象 =(SELECT distinct '2' From 部门性质说明 WHERE 工作性质 LIKE '发料部门' AND 部门id =[1] AND 服务对象 IN (2, 3))) " & _
                              "                     " & _
                              "           ) d,"
                              
                          
                          '连收费价目，主要找售价
                          gstrSQL = gstrSQL & _
                              "   (   SELECT 收费细目id, TO_CHAR (现价, '999999999990.9999') 售价 " & _
                              "       From 收费价目 " & _
                              "       WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (    SYSDATE >= 执行日期 AND 终止日期 IS NULL)) " & _
                              GetPriceClassString("") & _
                              "   ) p,"
                                   
                          
                          '连药品库存
                          gstrSQL = gstrSQL & _
                              "   (   SELECT 药品id, TO_CHAR (SUM (可用数量), " & mOraFMT.FM_数量 & ") 可用数量,TO_CHAR (SUM (实际数量)," & mOraFMT.FM_数量 & ") 库存数量,TO_CHAR (SUM (实际金额), " & mOraFMT.FM_金额 & ") 库存金额 " & _
                              "       From 药品库存 " & _
                              "       Where 库房id =[1]  and 性质=1 " & _
                              "       GROUP BY 药品id) s, "
                          
                          gstrSQL = gstrSQL & _
                              "   (   Select 材料ID,库房ID,库房货位 From 材料储备限额 " & _
                              "       Where 库房ID=[1]) E,收费项目目录 M"
                              
                          '整个条件
                          gstrSQL = gstrSQL & _
                              "   Where d.材料id = p.收费细目id  And D.材料id=M.id And (M.站点=[4] or M.站点 is null) And M.是否变价<>1 AND d.材料id  = s.药品id (+) and D.材料id =E.材料ID(+)" & _
                              "   ORDER BY d.编码"
                              '(Select 收费细目id From 收费执行科室 Where 执行科室id = [1])

                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, intStockID, strLike, CDate("3000-01-01"), gstrNodeNo, IIf(gSystem_Para.int简码方式 = 1, 2, 1))
                                                  
                        
                        If rsTemp.EOF Then
                            MsgBox "没有匹配的自制卫材！", vbInformation + vbOKOnly, gstrSysName
                            rsTemp.Close
                            Cancel = True
                            Exit Sub
                        ElseIf rsTemp.RecordCount = 1 Then
                            If SetColValue(.Row, rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, IIf(IsNull(rsTemp!规格), "", rsTemp!规格), _
                               rsTemp!单位, _
                               IIf(IsNull(rsTemp!售价), 0, rsTemp!售价), _
                               IIf(IsNull(rsTemp!最大效期), "0", rsTemp!最大效期), rsTemp!比例系数, rsTemp!是否变价, rsTemp!指导差价率, rsTemp!在用分批) = False Then
                               rsTemp.Close
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            rsTemp.Close
                        Else
                            Set mshDrug.Recordset = rsTemp
                            rsTemp.Close
                            Call SetDrugWidth(sngLeft, sngTop)
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    Call 提示库存数
                    'End If
                
                Case mconIntCol批号
                    '无处理
                    If strKey = "" Then
                        If .TxtVisible = True Then
                            .TextMatrix(.Row, mconIntCol批号) = ""
                        End If
                        If .ColData(mconIntCol效期) = 2 Then
                            .Col = mconIntCol效期
                        Else
                            .Col = mconIntCol数量
                        End If
                        
                        
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If zlCommFun.ActualLen(strKey) > mintBatchNoLen Then
                        MsgBox "批号长度不能超过" & mintBatchNoLen & "位或" & Int(mintBatchNoLen / 2) & "个汉字! ,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                Case mconIntCol效期
                    '有处理
                    If strKey <> "" Then
                        If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                            strKey = TranNumToDate(strKey)
                            If strKey = "" Then
                                MsgBox "失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = strKey
                            Exit Sub
                        End If
                        If Not IsDate(strKey) Then
                            MsgBox "失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                    ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol效期) Then
                    
                        If .TxtVisible = True Then
                            .Text = " "
                            Exit Sub
                        End If
                        
                        Exit Sub
                    End If
                    
                Case mconIntCol采购价
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "采购价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    
                    '设置金额
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol采购价) And .TextMatrix(.Row, mconIntCol数量) <> "" Then
                        .TextMatrix(.Row, mconIntCol采购金额) = Format(.TextMatrix(.Row, mconIntCol数量) * strKey, mFMT.FM_金额)
                        .TextMatrix(.Row, mconintCol差价) = Format(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol采购金额) = "", 0, .TextMatrix(.Row, mconIntCol采购金额)), mFMT.FM_金额)
                    End If
                    
                    显示合计金额
                Case mconIntCol采购金额
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "采购金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol采购金额) Then
                        If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                            .TextMatrix(.Row, mconIntCol采购价) = Format(strKey / .TextMatrix(.Row, mconIntCol数量), mFMT.FM_成本价)
                        End If
                        
                        .TextMatrix(.Row, mconintCol差价) = Format(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - strKey, mFMT.FM_金额)
                        .TextMatrix(.Row, mconIntCol采购金额) = Format(strKey, mFMT.FM_金额)
                    End If
                    显示合计金额
            Case mconIntCol灭菌日期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "灭菌日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "灭菌日期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("该卫材已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '计算失效期
                    .TextMatrix(.Row, mconIntCol灭菌失效期) = Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol灭菌效期)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol灭菌日期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
                Case mconIntCol数量
                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                        MsgBox "数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        If Val(strKey) = 0 Then
                            MsgBox "数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Abs(Val(strKey)) < 0.001 Then
                            MsgBox "数量的必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        Dim dblCostPrice As Double
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If .TextMatrix(.Row, 0) = "" Then Exit Sub
                        
                        '取组成卫材的数量,并设置自制卫材的采购价 等
                        If Set组成材料(Val(.TextMatrix(.Row, 0)), Val(strKey) * Val(.TextMatrix(.Row, mconIntCol比例系数)), True, dblCostPrice) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        .TextMatrix(.Row, mconIntCol采购价) = Format(dblCostPrice * Val(.TextMatrix(.Row, mconIntCol比例系数)), mFMT.FM_成本价)
                                
                        strKey = Format(strKey, mFMT.FM_数量)
                        .Text = strKey
                        If .TextMatrix(.Row, mconIntCol采购价) <> "" Then
                            .TextMatrix(.Row, mconIntCol采购金额) = Format(.TextMatrix(.Row, mconIntCol采购价) * strKey, mFMT.FM_金额)
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol采购金额)) >= 10 ^ 14 - 1 Then
                            MsgBox "采购金额必须小于" & (10 ^ 14 - 1) & ",请重新输入数量!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                            .TextMatrix(.Row, mconIntCol售价) = Format(.TextMatrix(.Row, mconIntCol采购价) / (1 - Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1)), mFMT.FM_零售价)
                        End If
                            
                        If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                            .TextMatrix(.Row, mconIntCol售价金额) = Format(.TextMatrix(.Row, mconIntCol售价) * strKey, mFMT.FM_金额)
                              
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol售价金额)) >= 10 ^ 14 - 1 Then
                            MsgBox "售价金额必须小于" & (10 ^ 14 - 1) & ",请重新输入数量!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .TextMatrix(.Row, mconintCol差价) = Format(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol采购金额) = "", 0, .TextMatrix(.Row, mconIntCol采购金额)), mFMT.FM_金额)
                        
                    End If
                    显示合计金额
                
            End Select
        End With
    ElseIf KeyCode = vbKeyDown And Shift = vbAltMask Then
        mshbill_CommandClick
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'从卫生卫材目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, ByVal str材料 As String, _
    ByVal str规格 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal int原效期 As Integer, ByVal num比例系数 As Double, _
    ByVal int是否变价 As Integer, ByVal dbl指导差价率 As Double, ByVal int在用分批 As Integer) As Boolean
    
    Dim intCount As Integer
    Dim rsStructure As New Recordset
    Dim intCol As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 一次性材料,灭菌效期 from 材料特性 where 材料id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            '.TextMatrix(intRow, intCol) = ""
            '2010-5-5 有数量时不赋空值
            If mconIntCol数量 <> intCol Or Trim(.TextMatrix(intRow, mconIntCol数量)) = "" Then
                .TextMatrix(intRow, intCol) = ""
            End If
        Next
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID Then
                    Call MsgBox("该卫生材料已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mconIntCol材料) = str材料
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol一次性材料) = zlStr.Nvl(rsTemp!一次性材料)
        .TextMatrix(intRow, mconIntCol灭菌效期) = zlStr.Nvl(rsTemp!灭菌效期)
        
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol售价) = Format(num售价, mFMT.FM_零售价)
        
        .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl指导差价率 & "||" & int是否变价 & "||" & int在用分批
        
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        SetInputFormat intRow
        
        If Set组成材料(lng材料ID, 0, True, 0) = False Then
            For intCol = 0 To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
            Exit Function
        End If
    End With
    Call 提示库存数
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'
'Private Function SetStructure(ByVal int材料id As Long) As Boolean
'    Dim rsTemp As New Recordset
'
'    SetStructure = False
'    vs组成材料.Redraw = False
'
'    If mint编辑状态 <> 4 Then
'        gstrSQL = "" & _
'            "   SELECT DISTINCT b.材料id, b.编码,b.名称 AS 商品名称, b.规格, c.上次产地, b.计算单位 as 单位, c.实际差价,c.实际金额, d.售价, " & _
'            "             (a.分子 / a.分母) AS 组成,c.可用数量, b.指导差价率,b.是否变价,b.在用分批 " & _
'            "   FROM 自制材料构成 a,(select b.材料id,a.编码,a.名称, a.规格,b.指导差价率,a.是否变价,a.计算单位,b.在用分批 from 收费项目目录 a,材料特性 b where a.id=b.材料id and nvl(a.是否变价,0)=0) b," & _
'            "        (SELECT 药品id, 实际差价,实际金额, 上次产地,可用数量 From 药品库存 WHERE 库房id =[2] and 性质=1) c," & _
'            "        (SELECT 收费细目id, TO_CHAR (现价," & mOraFMT.FM_零售价 & ") 售价 From 收费价目 WHERE ( (SYSDATE BETWEEN 执行日期 AND 终止日期) OR (    SYSDATE >= 执行日期 AND 终止日期 IS NULL) )) d " & _
'            "   Where a.原料材料id = b.材料ID " & _
'            "AND a.原料材料id = d.收费细目id " & _
'            "AND a.原料材料id = c.药品id (+) " & _
'            "AND a.自制材料id =[1]"
'
'        gstrSQL = gstrSQL & " union " & _
'             "  SELECT DISTINCT b.材料id, b.编码,b.名称 AS 商品名称, b.规格, c.上次产地, b.计算单位 as 单位, c.实际差价,c.实际金额,TO_CHAR ((c.实际金额/c.实际数量), " & mOraFMT.FM_零售价 & ")  as 售价, " & _
'             "          (a.分子 / a.分母) AS 组成,c.可用数量,b.指导差价率,b.是否变价,b.在用分批 " & _
'             "  FROM 自制材料构成 a,(select b.材料id,a.编码,a.名称,a.规格,b.指导差价率,a.是否变价,a.计算单位,b.在用分批 from 收费项目目录 a,材料特性 b where a.id=b.材料id and nvl(a.是否变价,0)=1) b," & _
'             "      (SELECT 药品id, 实际差价,实际金额, 上次产地,可用数量,实际数量 From 药品库存 WHERE 库房id =[2]  and 性质=1 and 实际数量>0 ) c " & _
'             "  Where a.原料材料id = b.材料ID " & _
'             "          AND a.原料材料id = c.药品id  " & _
'             "          AND a.自制材料id =[1]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, int材料id, cboType.ItemData(cboType.ListIndex))
'
'
'        If rsTemp.EOF Then
'            vs组成材料.Redraw = True
'            Exit Function
'        End If
'
'        With vs组成材料
'            .ClearBill
'            Do While Not rsTemp.EOF
'                If rsTemp!在用分批 = 1 Then
'                    MsgBox "组成卫材是一个在用分批卫材，但当前版本不支持在用分批的组成卫材，请检查！", vbInformation + vbOKOnly, gstrSysName
'                    vs组成材料.Redraw = True
'                    Exit Function
'                End If
'
'                .TextMatrix(.Row, mconIntCol材料名) = "[" & rsTemp!编码 & "]" & rsTemp!商品名称
'                .TextMatrix(.Row, mconIntCol构规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
'                .TextMatrix(.Row, mconIntCol构产地) = IIf(IsNull(rsTemp!上次产地), "", rsTemp!上次产地)
'                .TextMatrix(.Row, mconIntCol构单位) = rsTemp!单位
'                .TextMatrix(.Row, mconIntCol构售价) = Format(rsTemp!售价, mFMT.FM_零售价)
'                .TextMatrix(.Row, mconIntCol构可用数量) = Format(IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), mFMT.FM_数量)
'                .TextMatrix(.Row, mconIntCol构组成数量) = rsTemp!组成
'                .TextMatrix(.Row, mconintcol构指导差价率) = rsTemp!指导差价率 & "||" & IIf(IsNull(rsTemp!是否变价), 0, rsTemp!是否变价) & "||" & IIf(IsNull(rsTemp!在用分批), 0, rsTemp!在用分批)
'                .TextMatrix(.Row, mconintcol构实际差价) = IIf(IsNull(rsTemp!实际差价), "0", rsTemp!实际差价)
'                .TextMatrix(.Row, mconintcol构实际金额) = IIf(IsNull(rsTemp!实际金额), "0", rsTemp!实际金额)
'                .TextMatrix(.Row, mconintcol构材料id) = rsTemp!材料ID
'
'
'                If .Row = .Rows - 1 Then
'                    .Rows = .Rows + 1
'                End If
'                .Row = .Row + 1
'                rsTemp.MoveNext
'            Loop
'        End With
'    Else            '查看
'        gstrSQL = "" & _
'            "   SELECT DISTINCT a.材料id, c.编码,c.名称 AS 商品名称,b.一次性材料,b.灭菌效期, c.规格," & _
'            "           a.产地, c.计算单位 as 单位,a.实际数量,a.成本价,a.成本金额,a.零售价,a.零售金额,a.差价 " & _
'            "   FROM (  Select 药品id as 材料id,产地,实际数量,成本价,成本金额,零售价,零售金额,差价 " & _
'            "           From 药品收发记录 " & _
'            "           Where   no=[1] and 单据=16 and 记录状态=[2]" & _
'            "                   and 入出系数=-1 and 扣率=[4] AND 费用id =[3]) a," & _
'            "       材料特性 b,收费项目目录 c " & _
'            "Where a.材料id = b.材料ID and a.材料id=c.id "
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNo.Tag, mint记录状态, int材料id, mshBill.Row)
'
'        If rsTemp.EOF Then
'            vs组成材料.Redraw = True
'            Exit Function
'        End If
'        With vs组成材料
'            .ClearBill
'            Do While Not rsTemp.EOF
'                .TextMatrix(.Row, mconIntCol材料名) = "[" & rsTemp!编码 & "]" & rsTemp!商品名称
'                .TextMatrix(.Row, mconIntCol构规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
'                .TextMatrix(.Row, mconIntCol构产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
'                .TextMatrix(.Row, mconIntCol构单位) = rsTemp!单位
'                .TextMatrix(.Row, mconIntCol构数量) = Format(rsTemp!实际数量, mFMT.FM_数量)
'                .TextMatrix(.Row, mconIntCol构采购价) = Format(rsTemp!成本价, mFMT.FM_成本价)
'                .TextMatrix(.Row, mconIntCol构采购金额) = Format(IIf(IsNull(rsTemp!成本金额), 0, rsTemp!成本金额), mFMT.FM_金额)
'                .TextMatrix(.Row, mconIntCol构售价) = Format(rsTemp!零售价, mFMT.FM_零售价)
'                .TextMatrix(.Row, mconIntCol构售价金额) = Format(IIf(IsNull(rsTemp!零售金额), 0, rsTemp!零售金额), mFMT.FM_金额)
'                .TextMatrix(.Row, mconintCol构差价) = Format(IIf(IsNull(rsTemp!差价), 0, rsTemp!差价), mFMT.FM_金额)
'                .TextMatrix(.Row, mconintcol构材料id) = rsTemp!材料ID
'
'                If .Row = .Rows - 1 Then
'                    .Rows = .Rows + 1
'                End If
'                .Row = .Row + 1
'                rsTemp.MoveNext
'            Loop
'
'        End With
'        rsTemp.Close
'        vs组成材料.Redraw = True
'        Exit Function
'    End If
'    rsTemp.Close
'    SetStructure = True
'    vs组成材料.Redraw = True
'    Exit Function
'errHandle:
'    vs组成材料.Redraw = True
'    Exit Function
'
'End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    If mblnEdit = False Then Exit Sub
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then
            .ColData(mconIntCol效期) = 5
            Exit Sub
        End If
        
        If .TextMatrix(intRow, mconIntCol一次性材料) = "1" Then
            .ColData(mconIntCol灭菌日期) = 2
            .ColData(mconIntCol灭菌失效期) = 5
        Else
            .ColData(mconIntCol灭菌日期) = 5              '禁止
            .ColData(mconIntCol灭菌失效期) = 5
        End If
        
         
        If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
            If Split(.TextMatrix(intRow, mconIntCol原销期), "||")(0) = "0" Then
                .ColData(mconIntCol效期) = 5
            Else
                .ColData(mconIntCol效期) = 2                '日期输入框
            End If
        Else
            .ColData(mconIntCol效期) = 5
        End If
    End With
End Sub


Private Sub mshDrug_DblClick()
    mshDrug_KeyPress 13
    
End Sub

Private Sub mshDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshDrug
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshDrug_KeyPress(KeyAscii As Integer)
    With mshDrug
        If KeyAscii = 13 Then
            If Not SetColValue(mshBill.Row, .TextMatrix(.Row, 4), "[" & .TextMatrix(.Row, 0) & "]" & .TextMatrix(.Row, 1), _
                 .TextMatrix(.Row, 2), .TextMatrix(.Row, 5), Val(.TextMatrix(.Row, 8)), _
                 IIf(IsNull(.TextMatrix(.Row, 10)), "0", .TextMatrix(.Row, 10)), .TextMatrix(.Row, 7), Val(.TextMatrix(.Row, 11)), Val(.TextMatrix(.Row, 12)), Val(.TextMatrix(.Row, 13))) Then
                mshBill.SetFocus
                mshBill.Col = mconIntCol材料
                .Visible = False
                Exit Sub
            End If
            .Visible = False
            mshBill.Text = "[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 4)
            
            mshBill.Col = mconIntCol批号
            
            mshBill.SetFocus
        End If
    End With
                
            
End Sub

Private Sub mshDrug_LostFocus()
    SaveFlexState mshDrug, mstrCaption
    If mshDrug.Visible Then mshDrug.Visible = False
End Sub

Private Sub vs组成材料_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub
 

Private Sub r_DecideInput(strInput As String, Cancel As Boolean)

End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    If txtNo.Locked = False Then
        If Trim(txtNo.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        
        If InStr(1, txtNo.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
        
        If LenB(StrConv(txtNo.Text, vbFromUnicode)) > txtNo.MaxLength Then
            ShowMsgBox "单据号超长,最多能输入" & CInt(txtNo.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNo.MaxLength & "个字符!"
            txtNo.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        MsgBox "第" & intLop & "行卫材的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行卫材的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol批号
                        Exit Function
                    End If
                    
          
                    If Split(.TextMatrix(intLop, mconIntCol原销期), "||")(0) <> "0" Then
                        If .TextMatrix(intLop, mconIntCol批号) = "" Or .TextMatrix(intLop, mconIntCol效期) = "" Then
                            MsgBox "第" & intLop & "行的卫材是效期卫材,请把它的批号及效期信息完整输入单据中！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mconIntCol批号) = "" Then
                                .Col = mconIntCol批号
                            Else
                                .Col = mconIntCol效期
                            End If
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol采购金额
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    If Check原料库存(Val(.TextMatrix(intLop, 0)), Val(.TextMatrix(intLop, mconIntCol数量)) * Val(.TextMatrix(intLop, mconIntCol比例系数))) = False Then
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    ValidData = True
End Function


Private Function SaveCard() As Boolean
    Dim chrNo As Variant
    Dim lng序号 As Long, lngStockID As Long, lng记录数 As Long, lng材料ID As Long, lng制剂室ID As Long
    Dim str批号 As String, str效期 As String, str填制日期 As String, str灭菌日期 As String, str灭菌效期 As String
    Dim dbl填制数量 As Double, dbl成本价 As Double, dbl成本金额 As Double
    Dim dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim str摘要 As String, str填制人 As String
    Dim intRow As Integer, cllProc As Collection
    
    SaveCard = False
    Set cllProc = New Collection
    With mshBill
        chrNo = Trim(txtNo)
        '算总记录数
        lng记录数 = 0
        For intRow = 1 To .Rows - 1
             If .TextMatrix(intRow, 0) <> "" Then
                    lng记录数 = lng记录数 + 1
             End If
        Next
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint编辑状态 = 1 Then 'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(69, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(69, lngStockID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNo.Tag = chrNo
        
        
        lng制剂室ID = cboType.ItemData(cboType.ListIndex)
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_自制材料入库_Delete('" & mstr单据号 & "')"
            AddArray cllProc, gstrSQL
        End If
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                str效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                dbl填制数量 = Round(Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl成本价 = Round(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mconIntCol采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零售价 = Round(Val(.TextMatrix(intRow, mconIntCol售价)) / .TextMatrix(intRow, mconIntCol比例系数), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mconIntCol售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mconintCol差价)), g_小数位数.obj_最大小数.金额小数)
                
                str灭菌日期 = IIf(.TextMatrix(intRow, mconIntCol灭菌日期) = "", "", .TextMatrix(intRow, mconIntCol灭菌日期))
                str灭菌效期 = IIf(.TextMatrix(intRow, mconIntCol灭菌失效期) = "", "", .TextMatrix(intRow, mconIntCol灭菌失效期))
                
                lng序号 = intRow
                'Zl_自制材料入库_Insert
                gstrSQL = "Zl_自制材料入库_Insert("
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  对方部门id_In In 药品收发记录.对方部门id%Type,
                gstrSQL = gstrSQL & "" & lng制剂室ID & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  实际数量_In   In 药品收发记录.实际数量%Type,
                gstrSQL = gstrSQL & "" & dbl填制数量 & ","
                '  零售价_In     In 药品收发记录.零售价%Type,
                gstrSQL = gstrSQL & "" & dbl零售价 & ","
                '  零售金额_In   In 药品收发记录.零售金额%Type,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '  填制人_In     In 药品收发记录.填制人%Type,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '  批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '  效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                gstrSQL = gstrSQL & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                gstrSQL = gstrSQL & IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  摘要_In       In 药品收发记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "'" & str摘要 & "',"
                '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
                '  记录数_In     In Integer := 0
                gstrSQL = gstrSQL & "" & lng记录数 & ")"
                AddArray cllProc, gstrSQL
            End If
        Next
    End With
    
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rsTemp As New ADODB.Recordset
    Dim dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntCol材料) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    Select Case mintUnit
        Case 1
            strQuantity = "可用数量/换算系数 "
        Case 0
            strQuantity = "可用数量 "
    End Select
    
    gstrSQL = "" & _
        "   Select b.材料ID , Sum(" & strQuantity & ") as 数量 " & _
        "   From 药品库存 a,材料特性 b " & _
        "   Where a.性质=1 and a.药品id=b.材料id and 可用数量<>0 And 库房ID=[1]" & _
        "       and b.材料ID=[2]" & _
        "   Group by b.材料ID "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--提示库存数", cboStock.ItemData(cboStock.ListIndex), intID)
    
    With rsTemp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        dbl数量 = IIf(IsNull(!数量), 0, !数量)
        stbThis.Panels(2).Text = "该卫材当前库存数为[" & dbl数量 & "]"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub 提示原料库存数()
    Dim rsTemp As New ADODB.Recordset
    Dim dbl数量 As Double, lng材料ID As Long
    
    On Error GoTo ErrHandle
    With vs组成材料
        If .TextMatrix(.Row, .ColIndex("原材料编码及名称")) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        lng材料ID = Val(.Cell(flexcpData, .Row, .ColIndex("原材料编码及名称")))
    End With
    
    gstrSQL = "" & _
        "   Select b.ID as 材料id, Sum(可用数量) as 数量,b.计算单位 as 单位 " & _
        "   From 药品库存 a,收费项目目录 b " & _
        "   Where a.性质=1 and a.药品id=b.id and 可用数量<>0 And 库房ID=[1]" & _
        "       and b.ID=[2]" & _
        "   Group by b.ID,b.计算单位 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "---提示原料库存数", cboType.ItemData(cboType.ListIndex), lng材料ID)
    
    With rsTemp
        If .EOF Then
            stbThis.Panels(2).Text = "当前无库存"
            Exit Sub
        End If
        dbl数量 = !数量
        stbThis.Panels(2).Text = "该卫材当前库存数为[" & dbl数量 & "]" & zlStr.Nvl(!单位)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    ImeLanguage True
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
    ImeLanguage False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'打印单据
Private Sub printbill()
    Dim strNo As String
    strNo = txtNo.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1713", mint记录状态, mintUnit, 1713, "卫材自制入库单", strNo
End Sub


'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption & "--取字段长度"
    GetBatchNoLen = rsTemp.Fields(0).DefinedSize
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Set组成材料(ByVal lng材料ID As Long, Optional dbl自制数量 As Double = 0, Optional bln检查库存 As Boolean = True, Optional ByRef dblOut成本价 As Double = 0) As Boolean
    '------------------------------------------------------------------------------
    '功能:设置相关的组成材料]
    '    int材料id-材料ID
    '    dbl自制数量-自制材料输入的数量
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/03/21
    '------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, rsSort As New ADODB.Recordset
    Dim lngRow As Long, lng库房ID As Long, arrtemp As Variant
    Dim dbl剩于数量 As Double, dbl当前数量 As Double, dbl可用数量 As Double
    Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double, dblSum成本金额 As Double
    Dim blnContinue As Boolean '继续
    Dim bln实价 As Boolean
    err = 0: On Error GoTo ErrHand:
    blnContinue = False
    Set组成材料 = False
    
    vs组成材料.Redraw = flexRDNone
    
    On Error GoTo ErrHand
    If mint编辑状态 <> 4 Then
        gstrSQL = "" & _
        "   SELECT DISTINCT b.ID as 材料id, b.编码,b.名称 AS 商品名称, b.规格, b.计算单位 as 单位,d.售价, " & _
        "             (a.分子 / a.分母) AS 组成, C.指导差价率,B.是否变价,C.在用分批 " & _
        "   FROM 自制材料构成 a,收费项目目录 B,材料特性 C," & _
        "        (  SELECT 收费细目id,现价 as 售价 From 收费价目 WHERE  ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL))" & _
        GetPriceClassString("") & ") d " & _
        "   Where a.原料材料id = b.ID And (B.站点=[2] or B.站点 is null) and A.原料材料id=c.材料ID  AND a.原料材料id = d.收费细目id(+)" & _
        "         AND a.自制材料id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID, gstrNodeNo)
        If rsTemp.EOF Then
            vs组成材料.Redraw = flexRDBuffered
            Exit Function
        End If

        lng库房ID = cboType.ItemData(cboType.ListIndex)
        
        With vs组成材料
            .Clear (1)
            .Rows = 2
            lngRow = 1
            Do While Not rsTemp.EOF
                    If mint编辑状态 <> 1 Then
                        gstrSQL = "" & _
                            "   SELECT nvl(批次,0) 批次," & _
                            "          nvl(可用数量,0)  as 可用数量, nvl(实际数量,0)  as 实际数量, " & _
                            "          nvl(实际差价,0)  as 实际差价,nvl(实际金额,0) as 实际金额,零售价," & _
                            "         上次产地,上次批号,上次生产日期, 效期 ,nvl(可用数量,0) as 实际可用数量" & _
                            "   From 药品库存 " & _
                            "   WHERE 药品id=[1] and 库房id =[2]  and 性质=1 " & _
                            "   Union ALL " & _
                            "   Select nvl(批次,0) as 批次,填写数量 as 可用数量,0 as 实际数量,0 as 实际差价,0 as 实际金额,零售价,产地,批号,生产日期,效期,0 as 实际可用数量 " & _
                            "   From 药品收发记录 " & _
                            "   where 单据=16 and 药品id=[1] and NO=[3] and 入出系数=-1"
                        gstrSQL = "" & _
                            "   SELECT nvl(批次,0) 批次," & _
                            "           sum(nvl(可用数量,0)) as 可用数量,sum(nvl(实际数量,0)) as 实际数量, " & _
                            "           Sum(nvl(实际差价,0)) as 实际差价,sum(nvl(实际金额,0)) as 实际金额, " & _
                            "           Sum(nvl(实际可用数量,0)) as 实际可用数量,max(零售价) as 零售价," & _
                            "           max(上次产地) as 上次产地,max(上次批号) as 上次批号,max(上次生产日期) as 上次生产日期,max(效期) as 效期 " & _
                            "   From (" & gstrSQL & ") " & _
                            "   Group by nvl(批次,0) " & _
                            "   Order by 批次"
                    Else
                        gstrSQL = "" & _
                            "   SELECT nvl(批次,0) 批次," & _
                            "           sum(nvl(可用数量,0)) as 可用数量,sum(nvl(实际数量,0)) as 实际数量, " & _
                            "           Sum(nvl(实际差价,0)) as 实际差价,sum(nvl(实际金额,0)) as 实际金额, " & _
                            "           sum(nvl(可用数量,0)) as 实际可用数量,Max(零售价) as 零售价," & _
                            "           max(上次产地) as 上次产地,max(上次批号) as 上次批号,max(上次生产日期) as 上次生产日期,max(效期) as 效期 " & _
                            "   From 药品库存 " & _
                            "   WHERE 药品id=[1] and 库房id =[2]  and 性质=1 " & _
                            "   Group by nvl(批次,0) " & _
                            "   Order by 批次"
                    End If
                    
                    Set rsSort = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!材料ID)), lng库房ID, mstr单据号)
                    dbl剩于数量 = dbl自制数量 * Val(zlStr.Nvl(rsTemp!组成))
                    .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!商品名称)
                    .Cell(flexcpData, lngRow, .ColIndex("原材料编码及名称")) = zlStr.Nvl(rsTemp!材料ID)
                    .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
                    .TextMatrix(lngRow, .ColIndex("单位")) = zlStr.Nvl(rsTemp!单位)
                    dblSum成本金额 = 0
                    .Cell(flexcpData, lngRow, .ColIndex("单位")) = 0 & "," & 0 & "," & 0 & "," & 0 & "," & Val(zlStr.Nvl(rsTemp!指导差价率))
                    Do While Not rsSort.EOF
                        .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!商品名称)
                        .Cell(flexcpData, lngRow, .ColIndex("原材料编码及名称")) = zlStr.Nvl(rsTemp!材料ID)
                        .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
                        .TextMatrix(lngRow, .ColIndex("单位")) = zlStr.Nvl(rsTemp!单位)
                        .Cell(flexcpData, lngRow, .ColIndex("单位")) = zlStr.Nvl(rsSort!实际可用数量) & "," & zlStr.Nvl(rsSort!实际数量) & "," & zlStr.Nvl(rsSort!实际差价) & "," & zlStr.Nvl(rsSort!实际金额) & "," & Val(zlStr.Nvl(rsTemp!指导差价率))
                        .TextMatrix(lngRow, .ColIndex("批号")) = zlStr.Nvl(rsSort!上次批号)
                        
                        .Cell(flexcpData, lngRow, .ColIndex("批号")) = zlStr.Nvl(rsSort!批次)
                        
                        If Val(zlStr.Nvl(rsTemp!是否变价)) = 0 Then
                            '定价
                            .TextMatrix(lngRow, .ColIndex("售价")) = Format(Val(zlStr.Nvl(rsTemp!售价)), mFMT.FM_零售价)
                            .Cell(flexcpData, lngRow, .ColIndex("售价")) = zlStr.Nvl(rsTemp!售价)
                        ElseIf Val(zlStr.Nvl(rsSort!实际数量)) <> 0 Then
                            If Val(zlStr.Nvl(rsSort!零售价)) <> 0 Then
                                .TextMatrix(lngRow, .ColIndex("售价")) = Format(Val(zlStr.Nvl(rsSort!零售价)), mFMT.FM_零售价)
                                .Cell(flexcpData, lngRow, .ColIndex("售价")) = Val(zlStr.Nvl(rsSort!零售价))
                            Else
                                .TextMatrix(lngRow, .ColIndex("售价")) = Format(Val(zlStr.Nvl(rsSort!实际金额)) / Val(zlStr.Nvl(rsSort!实际数量)), mFMT.FM_零售价)
                                .Cell(flexcpData, lngRow, .ColIndex("售价")) = Val(zlStr.Nvl(rsSort!实际金额)) / Val(zlStr.Nvl(rsSort!实际数量))
                            End If
                        Else
                            .TextMatrix(lngRow, .ColIndex("售价")) = ""
                            .Cell(flexcpData, lngRow, .ColIndex("售价")) = ""
                        End If
                        dbl可用数量 = Val(zlStr.Nvl(rsSort!可用数量))
                        If dbl可用数量 >= dbl剩于数量 Then
                            dbl当前数量 = dbl剩于数量
                        Else
                            dbl当前数量 = dbl可用数量
                        End If
                        
                        .TextMatrix(lngRow, .ColIndex("数量")) = Format(dbl当前数量, mFMT.FM_数量)
                        .Cell(flexcpData, lngRow, .ColIndex("数量")) = dbl当前数量
                        .TextMatrix(lngRow, .ColIndex("售价金额")) = Format(dbl当前数量 * Val(.Cell(flexcpData, lngRow, .ColIndex("售价"))), mFMT.FM_金额)
                
'                        Call 验证出库差价计算(lng库房ID, Val(zlStr.Nvl(rsTemp!材料ID)), Val(.Cell(flexcpData, lngRow, .ColIndex("批号"))), 1, Val(zlStr.Nvl(rsSort!实际差价)), Val(zlStr.Nvl(rsSort!实际金额)), Val(zlStr.Nvl(rsTemp!指导差价率)) / 100, dbl当前数量, Val(.TextMatrix(lngRow, .ColIndex("售价金额"))), dbl差价, dbl购价, dbl成本金额)
                        
'                        .TextMatrix(lngRow, .ColIndex("成本价")) = Format(dbl购价, mFMT.FM_成本价)
'                        .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(dbl成本金额, mFMT.FM_金额)
'                        .TextMatrix(lngRow, .ColIndex("差价")) = Format(dbl差价, mFMT.FM_金额)
                        
                        .TextMatrix(lngRow, .ColIndex("成本价")) = Format(Get成本价(Val(zlStr.Nvl(rsTemp!材料ID)), lng库房ID, Val(.Cell(flexcpData, lngRow, .ColIndex("批号")))), mFMT.FM_成本价)
                        .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) * dbl当前数量, mFMT.FM_金额)
                        .TextMatrix(lngRow, .ColIndex("差价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价金额"))) - Val(.TextMatrix(lngRow, .ColIndex("成本金额"))), mFMT.FM_金额)
                        
                        dbl剩于数量 = dbl剩于数量 - Val(zlStr.Nvl(rsSort!可用数量))
                        If dbl剩于数量 <= 0 Then
                            Exit Do
                        End If
                        rsSort.MoveNext
                        If rsSort.EOF Then Exit Do
                        .Rows = .Rows + 1
                        lngRow = lngRow + 1
                    Loop
                    
                    If Round(dbl剩于数量, 7) > 0 Then
                         .TextMatrix(lngRow, .ColIndex("数量")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))) + dbl剩于数量, mFMT.FM_数量)
                         .Cell(flexcpData, lngRow, .ColIndex("数量")) = Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))) + dbl剩于数量
                         .TextMatrix(lngRow, .ColIndex("售价金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))) * Val(.TextMatrix(lngRow, .ColIndex("售价"))), mFMT.FM_金额)
                         arrtemp = Split(.Cell(flexcpData, lngRow, .ColIndex("单位")) & ",,,,,", ",")
                         ' NVL(rsSort!实际可用数量) & "," & NVL(rsSort!实际数量) & "," & NVL(rsSort!实际差价) & "," & NVL(rsSort!实际金额),指导差价率
'                        Call 验证出库差价计算(lng库房ID, Val(NVL(rsTemp!材料ID)), Val(.Cell(flexcpData, lngRow, .ColIndex("批号"))), 1, Val(ArrTemp(2)), Val(ArrTemp(3)), Val(ArrTemp(4)) / 100, Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))), Val(.TextMatrix(lngRow, .ColIndex("售价金额"))), dbl差价, dbl购价, dbl成本金额)
'                        .TextMatrix(lngRow, .ColIndex("成本价")) = Format(dbl购价, mFMT.FM_成本价)
'                        .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(dbl成本金额, mFMT.FM_金额)
'                        .TextMatrix(lngRow, .ColIndex("差价")) = Format(dbl差价, mFMT.FM_金额)
                        
                        .TextMatrix(lngRow, .ColIndex("成本价")) = Format(Get成本价(Val(zlStr.Nvl(rsTemp!材料ID)), lng库房ID, Val(.Cell(flexcpData, lngRow, .ColIndex("批号")))), mFMT.FM_成本价)
                        .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) * Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))), mFMT.FM_金额)
                        .TextMatrix(lngRow, .ColIndex("差价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价金额"))) - Val(.TextMatrix(lngRow, .ColIndex("成本金额"))), mFMT.FM_金额)
                        
                        bln实价 = Val(zlStr.Nvl(rsTemp!是否变价)) = 1
                        
                        If bln检查库存 Then
                            If mint库存检查 = 0 Then
                                '不检查
                                If bln实价 Or Val(zlStr.Nvl(rsTemp!在用分批)) = 1 Then
                                    vs组成材料.Redraw = flexRDBuffered
                                    MsgBox "该自制卫材的原料卫材“" & .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                            ElseIf mint库存检查 = 1 Then
                                '检查，提醒
                                If bln实价 Or Val(zlStr.Nvl(rsTemp!在用分批)) = 1 Then
                                    vs组成材料.Redraw = flexRDBuffered
                                    MsgBox "该自制卫材的原料卫材“" & .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                ElseIf blnContinue = False Then
                                    If MsgBox("该自制卫材的原料卫材“" & .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) & "”可用库存数不够，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        vs组成材料.Redraw = flexRDBuffered
                                        Exit Function
                                    End If
                                    blnContinue = True
                                End If
                            ElseIf mint库存检查 = 2 Then
                                '禁止
                                vs组成材料.Redraw = flexRDBuffered
                                MsgBox "该自制卫材的原料卫材“" & .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
                dblSum成本金额 = 0
                '算成本价
                For lngRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, lngRow, .ColIndex("原材料编码及名称"))) <> 0 Then
                        If Val(.Cell(flexcpData, lngRow, .ColIndex("数量"))) <> 0 Then
                            dblSum成本金额 = dblSum成本金额 + Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                        End If
                    End If
                Next
            End With
            If dbl自制数量 <> 0 Then
                dblOut成本价 = dblSum成本金额 / dbl自制数量
            Else
                dblOut成本价 = 0
            End If
    Else            '查看
        gstrSQL = "" & _
            "   SELECT DISTINCT a.材料id,a.批号, c.编码,c.名称 AS 商品名称,b.一次性材料,b.灭菌效期, c.规格," & _
            "           a.产地, c.计算单位 as 单位,a.实际数量,a.成本价,a.成本金额,a.零售价,a.零售金额,a.差价 " & _
            "   FROM (  Select 药品id as 材料id,批号,产地,实际数量,成本价,成本金额,零售价,零售金额,差价 " & _
            "           From 药品收发记录 " & _
            "           Where   no=[1] and 单据=16 and 记录状态=[2]" & _
            "                   and 入出系数=-1 and 扣率=[4] AND 费用id =[3]) a," & _
            "       材料特性 b,收费项目目录 c " & _
            "Where a.材料id = b.材料ID and a.材料id=c.id And (C.站点=[5] or C.站点 is null) "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNo.Tag, mint记录状态, lng材料ID, mshBill.Row, gstrNodeNo)
        
        If rsTemp.EOF Then
            vs组成材料.Redraw = flexRDBuffered
            Exit Function
        End If
        With vs组成材料
            .Clear (1)
            .Rows = 2
            lngRow = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, .ColIndex("原材料编码及名称")) = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!商品名称)
                .Cell(flexcpData, .ColIndex("原材料编码及名称")) = zlStr.Nvl(rsTemp!材料ID)
                .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
                .TextMatrix(lngRow, .ColIndex("单位")) = zlStr.Nvl(rsTemp!单位)
                .TextMatrix(lngRow, .ColIndex("批号")) = zlStr.Nvl(rsTemp!批号)
                .TextMatrix(lngRow, .ColIndex("数量")) = Format(rsTemp!实际数量, mFMT.FM_数量)
                .TextMatrix(lngRow, .ColIndex("售价")) = Format(Val(zlStr.Nvl(rsTemp!零售价)), mFMT.FM_零售价)
                .TextMatrix(lngRow, .ColIndex("售价金额")) = Format(Val(zlStr.Nvl(rsTemp!零售金额)), mFMT.FM_金额)
                .TextMatrix(lngRow, .ColIndex("成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)), mFMT.FM_成本价)
                .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(Val(zlStr.Nvl(rsTemp!成本金额)), mFMT.FM_金额)
                .TextMatrix(lngRow, .ColIndex("差价")) = Format(Val(zlStr.Nvl(rsTemp!差价)), mFMT.FM_金额)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
        End With
        rsTemp.Close
        vs组成材料.Redraw = flexRDBuffered
        Exit Function
    End If
    rsTemp.Close
    Set组成材料 = True
    vs组成材料.Redraw = flexRDBuffered
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    vs组成材料.Redraw = flexRDBuffered
    Exit Function
End Function

Private Sub vs组成材料_BeforeSort(ByVal Col As Long, Order As Integer)
    
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim intRow As Integer
    
    
    With vs组成材料
        '自定义排序
        If .ExplorerBar > &H1000& Then Exit Sub
'
'        .GetSelection lngRow, lngCol, lngRows, lngCols
'        .Redraw = flexRDNone
'        '应用到非空行
'        For intRow = .Rows - 1 To .FixedRows Step -1
'            If Len(.TextMatrix(intRow, Col)) Then Exit For
'        Next
'
'        If intRow > .FixedRows Then
'            .Select .FixedRows, Col, intRow, Col
'            .Sort = Order
'        End If
'
'        ' 恢复选择
'        .Select lngRow, lngCol, lngRows, lngCols
'        .Redraw = flexRDDirect
        Order = 0
    End With
End Sub

Private Sub vs组成材料_EnterCell()
    Call 提示原料库存数
End Sub
Private Function Check原料库存(ByVal lng材料ID As Long, ByVal dbl自制数量 As Double) As Boolean
    '------------------------------------------------------------------------------
    '功能:检查原料库存是否合法
    '返回:有库存,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/23
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, rsStorck As New ADODB.Recordset
    Dim dbl可用数量 As Double
    Dim lng库房ID As Long, bln实价 As Boolean, blnContinue As Boolean
    
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
    "   SELECT DISTINCT b.ID as 材料id, b.编码,b.名称 AS 商品名称, b.规格, b.计算单位 as 单位, " & _
    "             (a.分子 / a.分母) AS 组成,B.是否变价,C.在用分批 " & _
    "   FROM 自制材料构成 a,收费项目目录 B,材料特性 C" & _
    "   Where a.原料材料id = b.ID and A.原料材料id=c.材料ID" & _
    "         AND a.自制材料id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    If rsTemp.EOF Then
        gstrSQL = "Select 编码,名称 From 收费项目目录 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
        If Not rsTemp.EOF Then
            ShowMsgBox "自制材料:" & zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称) & vbCrLf & " 没有相关的组成材料,请检查!"
        End If
        Exit Function
    End If
    
    lng库房ID = cboType.ItemData(cboType.ListIndex)
    
    If mint编辑状态 = 2 Then
        gstrSQL = "" & _
            "   SELECT nvl(批次,0) 批次," & _
            "          nvl(可用数量,0)  as 可用数量, nvl(实际数量,0)  as 实际数量, " & _
            "          nvl(实际差价,0)  as 实际差价,nvl(实际金额,0) as 实际金额, " & _
            "         上次产地,上次批号,上次生产日期, 效期 ,nvl(可用数量,0) as 实际可用数量" & _
            "   From 药品库存 " & _
            "   WHERE 药品id=[1] and 库房id =[2]  and 性质=1 " & _
            "   Union ALL " & _
            "   Select nvl(批次,0) as 批次,填写数量 as 可用数量,0 as 实际数量,0 as 实际差价,0 as 实际金额,产地,批号,生产日期,效期,0 as 实际可用数量 " & _
            "   From 药品收发记录 " & _
            "   where 单据=16 and 药品id=[1] and NO=[3] and 入出系数=-1"
        gstrSQL = "" & _
            "   SELECT  sum(nvl(可用数量,0)) as 可用数量 " & _
            "   From (" & gstrSQL & ") "
    Else
        gstrSQL = "" & _
            "   SELECT sum(nvl(可用数量,0)) as 可用数量 " & _
            "   From 药品库存 " & _
            "   WHERE 药品id=[1] and 库房id =[2]  and 性质=1 "
    End If
    
    Do While Not rsTemp.EOF
        Set rsStorck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!材料ID)), lng库房ID, mstr单据号)
        If rsStorck.EOF Then
            dbl可用数量 = 0
        Else
            dbl可用数量 = Val(zlStr.Nvl(rsStorck!可用数量))
        End If
        If Round(dbl可用数量, 7) < Round(dbl自制数量 * Val(zlStr.Nvl(rsTemp!组成)), 7) Then
            bln实价 = Val(zlStr.Nvl(rsTemp!是否变价)) = 1
            If mint库存检查 = 0 Then
                '不检查
                If bln实价 Or Val(zlStr.Nvl(rsTemp!在用分批)) = 1 Then
                    vs组成材料.Redraw = flexRDBuffered
                    MsgBox "该自制卫材的原料卫材“" & zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!商品名称) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint库存检查 = 1 Then
                '检查，提醒
                If bln实价 Or Val(zlStr.Nvl(rsTemp!在用分批)) Then
                    MsgBox "该自制卫材的原料卫材“" & zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!商品名称) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                ElseIf blnContinue = False Then
                    If MsgBox("该自制卫材的原料卫材“" & zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!商品名称) & "”可用库存数不够，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    blnContinue = True
                End If
            ElseIf mint库存检查 = 2 Then
                '禁止
                MsgBox "该自制卫材的原料卫材“" & zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!商品名称) & "”可用库存数不够，请检查该原料卫材的库存！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    Check原料库存 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

