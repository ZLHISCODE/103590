VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOtherInputCard 
   Caption         =   "卫材其他入库单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmOtherInputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   3030
      TabIndex        =   30
      Top             =   5730
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   28
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4080
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
      Begin VSFlex8Ctl.VSFlexGrid mshBill 
         Height          =   2730
         Left            =   225
         TabIndex        =   4
         Top             =   1035
         Width           =   11085
         _cx             =   19553
         _cy             =   4815
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   32
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOtherInputCard.frx":014A
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
         ExplorerBar     =   7
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
         Begin VB.Image imgLeft 
            Height          =   240
            Left            =   30
            Picture         =   "frmOtherInputCard.frx":058D
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6510
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9210
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   4440
         Width           =   1005
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
         TabIndex        =   19
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
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料其他入库单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   5925
         TabIndex        =   15
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   8400
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入出类别(&T)"
         Height          =   180
         Left            =   8040
         TabIndex        =   2
         Top             =   660
         Width           =   990
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
            Picture         =   "frmOtherInputCard.frx":0B17
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0D31
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0F4B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1165
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":137F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1599
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":17B3
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":19CD
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
            Picture         =   "frmOtherInputCard.frx":1BE7
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1E01
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":201B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2235
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":244F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2669
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2883
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2A9D
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
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
            Picture         =   "frmOtherInputCard.frx":2CB7
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
            Picture         =   "frmOtherInputCard.frx":354B
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":3A4D
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
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmOtherInputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln不强制控制指导价格 As Boolean
Private mbln时价卫材直接确定售价 As Boolean '外购入库时,时价卫材直接确定售价

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mstr单据号 As String                '具体的单据号;
Private mbln加价率 As Boolean               '时价卫材是否必须输入加价率
Private mdbl加价率 As Double
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintErrInfor As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录

Private mintBatchNoLen As Integer           '数据库中批号定义长度

Private mrsInOutType As Recordset           '入出类别
Dim mstrPrivs As String                     '权限
Private mbln分段加成率 As Boolean   '以分段加成率为依据
'刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private Const mlngModule = 1714
Private mbln库房  As Boolean    '该库房是否为卫材库!
Private mblnSort As Boolean     '存在排序,不触发相关事件
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看

Private mbln分批卫材批号产地控制 As Boolean  '是否检查分批卫材批号产地是否录入

Private Const mstrCaption As String = "卫材其他入库单"

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraMaxFmt As g_FmtString

'----------------------------------------------------------------------------------------------------------
'=========================================================================================
Private mblnFirst As Boolean    '第一次运行时
'=========================================================================================



'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSql As String
    
    GetDepend = False
    
    On Error GoTo ErrHandle
    strSql = "" & _
        "   SELECT B.Id,b.名称 " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID   AND A.单据 = 32 "

    zlDatabase.OpenRecordset rsTemp, strSql, "卫材其他入库管理-获取入出类别"
    If rsTemp.EOF Then
        MsgBox "没有设置卫材其他入库的入出类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    Set mrsInOutType = rsTemp
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
   '-----------------------------------------------------------------------------------------------------------
    '--功  能:显示或编辑卡片,是唯一入口
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
    mintErrInfor = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    
    Call GetRegInFor(g私有模块, "卫材其他入库管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
        txtNO.Locked = True
        txtNO.TabStop = True
        txtNO = mstr单据号
        txtNO.Tag = txtNO.Text
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
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
    ElseIf mint编辑状态 = 6 Then
        mblnEdit = False
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
 

Private Sub cboStock_Click()
    Call 当前仅为库房
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
                If mshBill.TextMatrix(i, mshBill.ColIndex("材料ID")) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变库房，有可能要改变相应卫材的单位，" & vbCrLf & "且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫材单位改变
                    mintcboIndex = .ListIndex
                    mshBill.Rows = 2: mshBill.Cell(flexcpData, 1, mshBill.Cols - 1) = ""
                    mshBill.Clear 1
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
        .Col = .ColIndex("卫材信息")
    End With
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("材料ID"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("冲销数量")) = Format(0, mFMT.FM_数量)
                .TextMatrix(intRow, .ColIndex("采购金额")) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, .ColIndex("售价金额")) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, .ColIndex("差价")) = Format(0, mFMT.FM_金额)
                '刘兴宏:零售价处理
                .TextMatrix(intRow, .ColIndex("零售金额")) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, .ColIndex("零售差价")) = Format(0, mFMT.FM_金额)
            End If
        Next
    End With
    Call 显示合计金额
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    With mshBill
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("材料ID"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("冲销数量")) = Format(Val(.TextMatrix(intRow, .ColIndex("数量"))), mFMT.FM_数量)
                .TextMatrix(intRow, .ColIndex("采购金额")) = Format(Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("采购价"))), mFMT.FM_金额)
                .TextMatrix(intRow, .ColIndex("售价金额")) = Format(Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("售价"))), mFMT.FM_金额)
                .TextMatrix(intRow, .ColIndex("差价")) = Format(Val(.TextMatrix(intRow, .ColIndex("售价金额"))) - Val(.TextMatrix(intRow, .ColIndex("采购金额"))), mFMT.FM_金额)
                '刘兴宏:零售价处理,主要是确定时价定价问题
                Call 计算零售价及零售差价(intRow, False)
            End If
        Next
    End With
    Call 显示合计金额
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'查找
Private Sub cmdFind_Click()
    Dim lngRow As Integer
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindVsRowNew mshBill, mshBill.ColIndex("卫材信息"), txtCode.Text, True
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
    Select Case mintErrInfor
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的卫材，请检查！", vbOKOnly, gstrSysName
            Else
                '单据已被删除
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            End If
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
        FindVsRowNew mshBill, mshBill.ColIndex("卫材信息"), txtCode.Text, False
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
    
    '设置排序数据集
    Call SetSortRecord
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        If Not 检查单价(17, txtNO.Tag) Then Exit Sub
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        '刘兴宏:2007/06/10:问题10813
        mstrTime_End = GetBillInfo(17, txtNO.Tag)
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
    
    If mint编辑状态 = 6 Then '冲销
        '检查库存是否充足
        If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
            MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
            txt摘要.SetFocus
            Exit Sub
        End If
        
        If SaveStrike Then Unload Me
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
    txtNO.Text = ""
    mblnSave = False
    mblnEdit = True
    mshBill.Rows = 2: mshBill.Clear 1
    
    Call RefreshRowNO(mshBill, mshBill.ColIndex("行号"), 1)
    SetEdit
    
    txt摘要.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
End Sub

Private Sub Form_Load()
    Dim strReg As String

    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mintUnit = Val(strReg)
 

    mblnFirst = True
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(2, g_售价)
    End With
    

    mintBatchNoLen = GetBatchNoLen()
    
    mbln加价率 = Get加价率
    mbln分段加成率 = IS分段加成率()
    mbln不强制控制指导价格 = ISCHECK不强制控制指导价格()
    mbln时价卫材直接确定售价 = is时价卫材直接确定售价()
    
    mbln分批卫材批号产地控制 = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    With cboType
        .Clear
        Do While Not mrsInOutType.EOF
            .AddItem mrsInOutType.Fields(1)
            .ItemData(.NewIndex) = mrsInOutType.Fields(0)
            mrsInOutType.MoveNext
        Loop
        .ListIndex = 0
    End With
      
    Call initCard
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(.ColIndex("采购价")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("采购金额")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 900, 0)
    End With
        
    mshBill_LostFocus
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    
    On Error GoTo ErrHandle

    '库房
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
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
    
    '初始化网格控件
    Call initGrid
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Case 2, 3, 4, 6
            If mint编辑状态 = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.名称  " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   Where a.库房id=b.id and A.单据 = 17 and a.no=[1]"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                
                If rsTemp.EOF Then: mintErrInfor = 2: Exit Sub
                
                With cboStock
                    .AddItem rsTemp!名称: .ItemData(.NewIndex) = rsTemp!Id: .ListIndex = 0
                End With
                rsTemp.Close
            End If
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.计算单位 AS 单位 ,c.计算单位 AS 零售单位,(A.填写数量 ) AS 数量,b.指导批发价 as 指导批发价 , a.成本价 as 成本价  ,  1 as 比例系数,"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,c.计算单位 AS 零售单位,(A.填写数量 / B.换算系数) AS 数量,b.指导批发价*B.换算系数 as 指导批发价 , a.成本价*B.换算系数 as 成本价 ,B.换算系数 as 比例系数,"
            End Select
            
            If mint编辑状态 <> 6 Then
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.药品id as 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息, " & _
                    "               zlSpellCode(c.名称) 名称,c.规格,c.产地 as 原产地,A.产地,A.批准文号, A.批号,to_char(a.生产日期,'yyyy-mm-dd') 生产日期," & _
                    "               b.最大效期,A.效期,a.灭菌日期,a.灭菌效期 as 灭菌失效期,a.商品条码,b.一次性材料,nvl(b.是否条码管理,0) as 条码管理,b.库房分批,b.灭菌效期," & strUnitQuantity & _
                    "               A.成本金额,A.零售价,to_number(nvl(to_char(a.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ") as 零售差价, " & _
                    "               A.零售金额,A.差价,b.指导差价率/100 as 指导差价率,nvl(b.加成率,0)/100 as 加成率,c.是否变价,b.在用分批, " & _
                    "               a.摘要,填制人,填制日期,审核人,审核日期,a.库房id,g.名称 as 部门,a.入出类别id " & _
                    "           FROM 药品收发记录 A,材料特性 b,收费项目目录 c,部门表 g " & _
                    "           Where A.药品id = B.材料id and a.药品id=c.id and a.库房id=g.id " & _
                    "                   AND A.记录状态 =[2]" & _
                    "                   AND A.单据 = 17 AND A.No =[1] )" & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Else
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.药品id as 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息, " & _
                    "                   zlSpellCode(c.名称) 名称,c.规格,c.产地 as 原产地,A.产地,A.批准文号, A.批号,to_char(a.生产日期,'yyyy-mm-dd') 生产日期," & _
                    "                   b.最大效期,A.效期,a.灭菌日期,a.灭菌效期 as 灭菌失效期,b.一次性材料,nvl(b.是否条码管理,0) as 条码管理,b.库房分批,b.灭菌效期," & strUnitQuantity & _
                    "                   A.成本金额,A.零售价,A.零售差价," & _
                    "                   A.零售金额,A.差价,b.指导差价率/100 as 指导差价率,nvl(b.加成率,0)/100 as 加成率,c.是否变价,b.在用分批,A.填写数量 as 真实数量, " & _
                    "                   a.库房id,g.名称 as 部门,a.入出类别id,a.商品条码 " & _
                    "           FROM (  Select min(id) as id, sum(填写数量) as 填写数量,sum(成本金额) as 成本金额, " & _
                    "                       药品id,序号,产地,批准文号, 批号,生产日期,效期,灭菌日期,灭菌效期,扣率,成本价," & _
                    "                       零售价,sum(零售金额) as 零售金额,Sum(差价) as 差价,Sum(to_number(nvl(to_char(x.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ")) as 零售差价," & _
                    "                       库房ID,入出类别ID,商品条码" & _
                    "                   From 药品收发记录 x " & _
                    "                   WHERE NO=[1] AND 单据=17  " & _
                    "                   group by 药品ID,序号,产地,批准文号, 批号,生产日期,效期,灭菌日期,灭菌效期,扣率,成本价,零售价,库房ID,入出类别ID,商品条码" & _
                    "                   having sum(填写数量)<>0 " & _
                    "                 ) A,材料特性 b,收费项目目录 c,部门表 g " & _
                    "           Where A.药品id = B.材料id and a.药品id=c.id and a.库房id=g.id ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mint记录状态)
            
            If rsTemp.EOF Then: mintErrInfor = 2: Exit Sub
            
            '刘兴宏:2007/06/10:问题10813
            mstrTime_Start = GetBillInfo(17, mstr单据号)
            
            Select Case mint编辑状态
                Case 2, 6
                    Txt填制人 = UserInfo.用户名
                    Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    If mint编辑状态 = 2 Then
                        Txt审核人 = ""
                        Txt审核日期 = ""
                    Else
                        Txt审核人 = UserInfo.用户名
                        Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case Else
                    Txt填制人 = rsTemp!填制人
                    Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
                    Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            If mint编辑状态 <> 6 Then
                txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            Else
                txt摘要.Text = Get摘要(mstr单据号)
            End If
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintErrInfor = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!入出类别ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            intRow = 0
            With mshBill
                .Clear 1
                .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
                Do While Not rsTemp.EOF
                    intRow = intRow + 1
                    .TextMatrix(intRow, .ColIndex("材料ID")) = zlStr.NVL(rsTemp!材料ID)
                    .TextMatrix(intRow, .ColIndex("卫材信息")) = zlStr.NVL(rsTemp!卫材信息)
                    .TextMatrix(intRow, .ColIndex("序号")) = zlStr.NVL(rsTemp!序号)
                    .TextMatrix(intRow, .ColIndex("规格")) = zlStr.NVL(rsTemp!规格)
                    .TextMatrix(intRow, .ColIndex("产地")) = zlStr.NVL(rsTemp!产地)
                    .TextMatrix(intRow, .ColIndex("批准文号")) = zlStr.NVL(rsTemp!批准文号)
                    .TextMatrix(intRow, .ColIndex("单位")) = zlStr.NVL(rsTemp!单位)
                    .TextMatrix(intRow, .ColIndex("批号")) = zlStr.NVL(rsTemp!批号)
                    .TextMatrix(intRow, .ColIndex("生产日期")) = zlStr.NVL(rsTemp!生产日期)
                    .TextMatrix(intRow, .ColIndex("效期")) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("一次性材料")) = Val(zlStr.NVL(rsTemp!一次性材料))
                    .TextMatrix(intRow, .ColIndex("条码管理")) = Val(zlStr.NVL(rsTemp!条码管理))
                    .TextMatrix(intRow, .ColIndex("灭菌效期")) = zlStr.NVL(rsTemp!灭菌效期)
                    .TextMatrix(intRow, .ColIndex("灭菌日期")) = IIf(IsNull(rsTemp!灭菌日期), "", Format(rsTemp!灭菌日期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("灭菌失效期")) = IIf(IsNull(rsTemp!灭菌失效期), "", Format(rsTemp!灭菌失效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("数量")) = Format(rsTemp!数量, mFMT.FM_数量)
                    .TextMatrix(intRow, .ColIndex("商品条码")) = zlStr.NVL(rsTemp!商品条码)
                    If rsTemp!数量 <> 0 Then
                        .TextMatrix(intRow, .ColIndex("采购价")) = Format(rsTemp!成本金额 / rsTemp!数量, mFMT.FM_成本价)
                    Else
                        .TextMatrix(intRow, .ColIndex("采购价")) = "0.00"
                    End If
   
                    '刘兴宏:零售价处理:零售价-->零售价;零售金额-->零售金额;差价-->零售差价;用途-->库房单位差价
                    ' 零售金额＝入库数量×零售价；
                    ' 零售差价"＝售价金额－零售金额，即按入库单位计算的金额和按零售单位计算的金额的差值；
                    .TextMatrix(intRow, .ColIndex("零售价")) = Format(Val(zlStr.NVL(rsTemp!零售价)), mFMT.FM_散装零售价)          'If Val(.TextMatrix(.row, .colindex("零售价"))) = 0 Then
                    
                    '反算售价
                    .TextMatrix(intRow, .ColIndex("售价")) = Format((Val(zlStr.NVL(rsTemp!零售金额)) - Val(zlStr.NVL(rsTemp!零售差价))) / Val(zlStr.NVL(rsTemp!数量)), mFMT.FM_零售价)
                    .TextMatrix(intRow, .ColIndex("零售单位")) = zlStr.NVL(rsTemp!零售单位)
                    
                    If mint编辑状态 = 6 Then
                        '冲销没有相关的差价
                        .TextMatrix(intRow, .ColIndex("零售差价")) = ""
                        .TextMatrix(intRow, .ColIndex("零售金额")) = ""
                        .TextMatrix(intRow, .ColIndex("差价")) = ""
                        .TextMatrix(intRow, .ColIndex("售价金额")) = ""
                        .TextMatrix(intRow, .ColIndex("采购金额")) = ""
                    Else
                        .TextMatrix(intRow, .ColIndex("零售差价")) = Format(Val(zlStr.NVL(rsTemp!差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, .ColIndex("零售金额")) = Format(Val(zlStr.NVL(rsTemp!零售金额)), mFMT.FM_金额)
                        '反算售价及售价金额
                        .TextMatrix(intRow, .ColIndex("差价")) = Format(Val(zlStr.NVL(rsTemp!差价)) - Val(zlStr.NVL(rsTemp!零售差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, .ColIndex("售价金额")) = Format(Val(zlStr.NVL(rsTemp!零售金额)) - Val(zlStr.NVL(rsTemp!零售差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, .ColIndex("采购金额")) = Format(Val(zlStr.NVL(rsTemp!成本金额)), mFMT.FM_金额)
                    End If
                    .TextMatrix(intRow, .ColIndex("原产地")) = IIf(IsNull(rsTemp!原产地), "!", rsTemp!原产地)
                    
                    '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                    .TextMatrix(intRow, .ColIndex("原销期")) = IIf(IsNull(rsTemp!最大效期), "0", rsTemp!最大效期) & "||" & rsTemp!加成率 & "||" & IIf(IsNull(rsTemp!是否变价), 0, rsTemp!是否变价) & "||" & IIf(IsNull(rsTemp!在用分批), 0, rsTemp!在用分批) & "||" & zlStr.NVL(rsTemp!库房分批, 0)
                    .TextMatrix(intRow, .ColIndex("比例系数")) = rsTemp!比例系数
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, .ColIndex("冲销数量")) = Format(0, mFMT.FM_数量)
                        .TextMatrix(intRow, .ColIndex("真实数量")) = zlStr.NVL(rsTemp!真实数量)
                    End If
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    SetEdit         '设置编辑属性
    Call RefreshRowNO(mshBill, mshBill.ColIndex("行号"), 1)
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get摘要(ByVal strNo As String) As String
    '获取新的摘要
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
         '冲销(取最后一次冲销的摘要)
    gstrSQL = "Select 摘要 From 药品收发记录 Where 单据=17 And No=[1] and (记录状态 =1 or mod(记录状态,3)=0) Order By 审核日期 Desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取摘要信息", strNo)
    
    If Not rsTemp.EOF Then
        Get摘要 = zlStr.NVL(rsTemp!摘要)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub imgLeft_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(mshBill.hwnd)
    lngLeft = vRect.Left + imgLeft.Left
    lngTop = vRect.Top + imgLeft.Height
    Call frmVsColSel.ShowColSet(Me, mstrCaption, mshBill, lngLeft, lngTop, imgLeft.Height)
    zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "列头信息", True
End Sub

Private Sub mshBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        If mblnSort = True Then Exit Sub
        Call zl_VsGridRowChange(mshBill, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub mshBill_AfterSort(ByVal Col As Long, Order As Integer)
    With mshBill
    End With
End Sub

Private Sub mshBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With mshBill
        If mblnEdit = False Then
            If mint编辑状态 = 6 Then
                If .ColIndex("冲销数量") = Col Then: Exit Sub
                Cancel = True: Exit Sub
            End If
            Cancel = True
        End If
        Select Case Col
        Case .ColIndex("卫材信息"), .ColIndex("批号")
        Case .ColIndex("灭菌日期")
            '非一次性材料不能编辑灭菌日期
            If Val(.TextMatrix(Row, .ColIndex("一次性材料"))) <> "1" Then Cancel = True: Exit Sub
        Case .ColIndex("商品条码")
            If Val(.TextMatrix(Row, .ColIndex("条码管理"))) <> 1 Then Cancel = True: Exit Sub
        Case .ColIndex("生产日期"), .ColIndex("批准文号")
        
        Case .ColIndex("数量"), .ColIndex("采购价"), .ColIndex("采购金额")
        Case .ColIndex("产地")
            If Val(.TextMatrix(Row, .ColIndex("材料ID"))) <= 0 Then '空行禁止输入
                Cancel = True
            End If
        Case .ColIndex("效期")
             '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
             If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                If Val(arrSplit(4)) = 0 Then Cancel = True: Exit Sub    '非库房分批,不能编辑效期
             Else
                Cancel = True
             End If
        Case .ColIndex("售价")
            '如果是时价卫材，则允许输入售价,且参数"时价卫材直接确定售价有效
             If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                If Not (Val(arrSplit(2)) = 1 And mbln时价卫材直接确定售价) Then Cancel = True: Exit Sub     '非库房分批,不能编辑效期
             Else
                Cancel = True
             End If
        Case .ColIndex("零售价")
             If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                If Val(arrSplit(2)) = 1 And (IIf(mbln库房, Val(arrSplit(4)) = 1, Val(arrSplit(3)) = 1)) Then
                   '实价卫材且库房分批的
                    Exit Sub
                End If
             End If
            Cancel = True: Exit Sub
        Case Else: Cancel = True
        End Select
    End With
End Sub
Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim arrSplit As Variant
    If mblnEdit = False Then Exit Sub
    
    With mshBill
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("产地")) = "0||" & IIf(Val(.TextMatrix(intRow, .ColIndex("材料ID"))) > 0, 0, 2) '& IIf(.TextMatrix(intRow, .ColIndex("原产地")) = "!", 0, 2)
        .ColData(.ColIndex("灭菌日期")) = "0||" & IIf(Val(.TextMatrix(intRow, .ColIndex("一次性材料"))) = 1, 0, 2)
        .ColData(.ColIndex("效期")) = "0||2"
        .ColData(.ColIndex("售价")) = "0||2"
        .ColData(.ColIndex("零售价")) = "0||2"
        .ColData(.ColIndex("商品条码")) = "0||" & IIf(.TextMatrix(intRow, .ColIndex("条码管理")) = "1", 0, 2)
        If .TextMatrix(intRow, .ColIndex("原销期")) <> "" Then
            '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
            arrSplit = Split(.TextMatrix(intRow, .ColIndex("原销期")), "||")
            .ColData(.ColIndex("效期")) = "0||" & IIf(Val(arrSplit(4)) = 1, 0, 2)
            '如果是时价卫材，则允许输入售价
            .ColData(.ColIndex("售价")) = "0||" & IIf(Val(arrSplit(2)) = 1 And mbln时价卫材直接确定售价, 0, 2)
            If Val(arrSplit(2)) = 1 And arrSplit(4) = 1 Then
                '实价且分批的
                .ColData(.ColIndex("零售价")) = "0||0"
            End If
        End If
    End With
End Sub
Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            cboStock.Enabled = False
            cboType.Enabled = False
            txt摘要.Enabled = True
            If mint编辑状态 = 6 Then .Editable = flexEDKbdMouse
            
            If mint编辑状态 <> 6 Then
                txt摘要.Enabled = False
            End If
        Else
            cboStock.Enabled = True

            cboType.Enabled = True
            txt摘要.Enabled = True
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的默认属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-12-02 11:39:14
    '-----------------------------------------------------------------------------------------------------------

    With mshBill
        
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, mshBill, mstrCaption, "列头信息", True, True
        
        .FixedCols = 1
        If mintUnit = 0 Then
            .ColHidden(.ColIndex("零售价")) = True
            .ColHidden(.ColIndex("零售单位")) = True
            .ColHidden(.ColIndex("零售金额")) = True
            .ColHidden(.ColIndex("零售差价")) = True
            .ColData(.ColIndex("零售价")) = -1
            .ColData(.ColIndex("零售单位")) = -1
            .ColData(.ColIndex("零售金额")) = -1
            .ColData(.ColIndex("零售差价")) = -1
        End If
        '隐藏冲销列
        .ColHidden(.ColIndex("冲销数量")) = IIf(mint编辑状态 = 6, False, True)
        If .ColWidth(.ColIndex("冲销数量")) = 0 And .ColHidden(.ColIndex("冲销数量")) = False Then .ColWidth(.ColIndex("冲销数量")) = 800
        .ColHidden(.ColIndex("材料ID")) = True
        .ColHidden(.ColIndex("序号")) = True
        .ColHidden(.ColIndex("真实数量")) = True
        .ColHidden(.ColIndex("一次性材料")) = True
        .ColHidden(.ColIndex("条码管理")) = True
        .ColHidden(.ColIndex("灭菌效期")) = True
        .ColHidden(.ColIndex("原产地")) = True
        .ColHidden(.ColIndex("原销期")) = True
        .ColHidden(.ColIndex("比例系数")) = True
        If mblnCostView = False Then
            .ColHidden(.ColIndex("采购价")) = True
            .ColHidden(.ColIndex("采购金额")) = True
            .ColHidden(.ColIndex("差价")) = True
        Else
            .ColHidden(.ColIndex("采购价")) = False
            .ColHidden(.ColIndex("采购金额")) = False
            .ColHidden(.ColIndex("差价")) = False
        End If
        .ColData(.ColIndex("冲销数量")) = "-1|0"
        .ColData(.ColIndex("卫材信息")) = "1|0"
        .ColData(.ColIndex("数量")) = "1|0"
        If mblnCostView = False Then
            .ColData(.ColIndex("采购价")) = "-1|1"
            .ColData(.ColIndex("采购金额")) = "-1|1"
            .ColData(.ColIndex("差价")) = "-1|1"
        Else
            .ColData(.ColIndex("采购价")) = "1|0"
            .ColData(.ColIndex("差价")) = "0||2"
        End If
        .ColData(.ColIndex("售价")) = "1|0"
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("材料ID")) = -1
        .ColData(.ColIndex("序号")) = -1
        .ColData(.ColIndex("真实数量")) = -1
        .ColData(.ColIndex("一次性材料")) = -1
        .ColData(.ColIndex("条码管理")) = -1
        .ColData(.ColIndex("灭菌效期")) = -1
        .ColData(.ColIndex("原产地")) = -1
        .ColData(.ColIndex("原销期")) = -1
        .ColData(.ColIndex("比例系数")) = -1
        
        .ColData(.ColIndex("规格")) = "0||2"
        .ColData(.ColIndex("单位")) = "0||2"
        .ColData(.ColIndex("灭菌失效期")) = "0||2"
        .ColData(.ColIndex("售价金额")) = "0||2"
        .ColData(.ColIndex("零售单位")) = "0||2"
        .ColData(.ColIndex("零售金额")) = "0||2"
        .ColData(.ColIndex("零售差价")) = "0||2"
        
        If gblnCode Then
            .ColData(.ColIndex("商品条码")) = "0||2"
        Else
            .ColData(.ColIndex("商品条码")) = "-1||1"
            .ColHidden(.ColIndex("商品条码")) = True
        End If

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
    With txtNO
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
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
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
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
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

    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "列头信息", True, True
        
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "列头信息", True, True
End Sub
Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:审核其他入库单
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-02 11:40:30
    '-----------------------------------------------------------------------------------------------------------
    mblnSave = False: SaveCheck = False
    
    gstrSQL = "zl_材料其他入库_Verify('" & txtNO.Tag & "','" & UserInfo.用户名 & "')"
    
    On Error GoTo ErrHandle
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
    SaveCheck = True: mblnSave = True: mblnSuccess = True: mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mshBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置相关的格式
    '返回:
    '编制:刘兴洪
    '日期:2008-12-02 11:43:38
    '-----------------------------------------------------------------------------------------------------------
    Dim arrSplit As Variant, str批号 As String, strxq As String
    With mshBill
        Select Case Col
        Case .ColIndex("卫材信息")
            .ColComboList(Col) = "..."
        Case .ColIndex("产地")
        Case .ColIndex("采购价"), .ColIndex("采购金额")
            Call 计算零售价及零售差价(Row, True)
              显示合计金额
        Case .ColIndex("售价")
            Call 计算零售价及零售差价(Row, True)
              显示合计金额
        Case .ColIndex("零售价")
            Call 计算零售价及零售差价(Row, False)
              显示合计金额
        Case .ColIndex("数量"), .ColIndex("冲销数量")
            Call 计算零售价及零售差价(Row, True)
              显示合计金额
        Case .ColIndex("商品条码")
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
        Case .ColIndex("批号")
                If Trim(.TextMatrix(Row, .ColIndex("批号"))) = "" Or IsNumeric(.TextMatrix(Row, .ColIndex("批号"))) = False Then
                    If Not IsDate(Trim(.TextMatrix(Row, .ColIndex("生产日期")))) Then
                        str批号 = ""
                    Else
                        str批号 = Format(.TextMatrix(Row, .ColIndex("生产日期")), "yyyymmdd")
                    End If
                Else
                    str批号 = Trim(.TextMatrix(Row, .ColIndex("批号")))
                End If
                If str批号 <> "" Then
                    '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                    arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")) & "||||||||||", "||")
                    If IsNumeric(str批号) And Val(arrSplit(0)) <> 0 Then
                        strxq = UCase(str批号)
                        If Trim(.TextMatrix(Row, .ColIndex("生产日期"))) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(Row, .ColIndex("生产日期")) = Format(strxq, "yyyy-mm-dd")
                                'Call CheckLapse(.TextMatrix(row, .ColIndex("效期")))
                            End If
                        End If
                    End If
                End If
        Case .ColIndex("生产日期")
'                If Trim(.TextMatrix(row, .ColIndex("批号"))) = "" Or IsNumeric(.TextMatrix(row, .ColIndex("批号"))) = False Then
                    If Not IsDate(Trim(.TextMatrix(Row, .ColIndex("生产日期")))) Then
                        str批号 = ""
                    Else
                        str批号 = Format(.TextMatrix(Row, .ColIndex("生产日期")), "yyyymmdd")
                    End If
'                Else
'                    str批号 = Trim(.TextMatrix(row, .ColIndex("批号")))
'                End If
                
                If str批号 <> "" Then
                    '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                    arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")) & "||||||||||", "||")
                    If IsNumeric(str批号) And Val(arrSplit(0)) <> 0 Then
                        strxq = UCase(str批号)
                        If Trim(.TextMatrix(Row, .ColIndex("效期"))) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(Row, .ColIndex("效期")) = Format(DateAdd("M", Val(arrSplit(0)), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(Row, .ColIndex("效期")))
                            End If
                        End If
                    End If
                End If
        End Select
    End With
End Sub
Private Sub AfterAddRow(Row As Long)
    '增加行后
    Call RefreshRowNO(mshBill, mshBill.ColIndex("行号"), Row)
End Sub

Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If Val(.TextMatrix(Row, .ColIndex("材料ID"))) <> 0 Then
            If MsgBox("你是否真的要删除卫生材料为“" & .TextMatrix(.Row, .ColIndex("卫材信息")) & "”的记录吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub
Private Sub mshBill_BeforeSort(ByVal Col As Long, Order As Integer)
    mblnSort = True
    Call zl_VsGridBeforeSort(mshBill, Col, Order, mshBill.ColIndex("行号"))
    With mshBill
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColorBkg
        Call zl_VsGridRowChange(mshBill, .FixedRows, .Row, .FixedCols, .Col)
        If InStr(1, "12", mint编辑状态) > 0 Then Call RefreshRowNO(mshBill, .ColIndex("行号"), 1)
    End With
    mblnSort = False
End Sub

Private Sub mshBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '功能:按钮选择
    '参数:
    '--------------------------------------------------------------------------
    Dim lngRow As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    With mshBill
        Select Case Col
        Case .ColIndex("卫材信息")
            If Select卫材信息("") = False Then Exit Sub
            Call zlVsMoveGridCell(mshBill, .ColIndex("卫材信息"), , IIf(mint编辑状态 = 1 Or mint编辑状态 = 2, True, False), lngRow)
            
        Case .ColIndex("产地")
            If SelectAndNotAddItem(Me, mshBill, "", "材料生产商", "材料生产商选择器", True, True, , zl_获取站点限制(True)) = True Then
                Call zlVsMoveGridCell(mshBill, .ColIndex("卫材信息"), , IIf(mint编辑状态 = 1 Or mint编辑状态 = 2, True, False), lngRow)
            End If
            
            If .TextMatrix(.Row, .ColIndex("产地")) <> "" Then
                gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("产地")), .TextMatrix(.Row, .ColIndex("材料ID")))
                If rsTemp.RecordCount Then
                    .TextMatrix(.Row, .ColIndex("批准文号")) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                Else
                    .TextMatrix(.Row, .ColIndex("批准文号")) = ""
                End If
            End If
        Case .ColIndex("效期")
            If SelDate(Col) = False Then Exit Sub
        Case .ColIndex("生产日期")
            If SelDate(Col) = False Then Exit Sub
        Case .ColIndex("灭菌日期")
            If SelDate(Col) = False Then Exit Sub
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_ChangeEdit()
    mblnChange = True
End Sub
 
Private Sub mshBill_GotFocus()
    Call zl_VsGridGotFocus(mshBill)
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    
    With mshBill
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                mshBill_CellButtonClick .Row, .Col
            Else
            
            Select Case .Col
            Case .ColIndex("卫材信息"), .ColIndex("产地"), .ColIndex("效期"), .ColIndex("生产日期"), .ColIndex("灭菌日期")
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If
 
        If KeyCode = vbKeyDelete Then
            blnCancel = False
            '删除行前
            Call BeforeDeleteRow(.Row, blnCancel)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            '删除行后
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If Val(.TextMatrix(.Row, .ColIndex("材料ID"))) = 0 And .Col = .ColIndex("卫材信息") Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(mshBill, .ColIndex("卫材信息"), , IIf(mint编辑状态 = 1 Or mint编辑状态 = 2, True, False), lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub

Private Sub mshBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    Dim rsProvider As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        Select Case Col
        Case .ColIndex("卫材信息")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If Select卫材信息(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case .ColIndex("产地")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If SelectAndNotAddItem(Me, mshBill, strKey, "材料生产商", "材料生产商选择器", True, True, , zl_获取站点限制(True)) = True Then
                gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("产地")), .TextMatrix(.Row, .ColIndex("材料ID")))
                If rsProvider.RecordCount > 0 Then
                    .TextMatrix(.Row, .ColIndex("批准文号")) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                Else
                    .TextMatrix(.Row, .ColIndex("批准文号")) = ""
                End If
            Else
                .EditText = ""
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
        Case Else
        
        End Select
        Call zlVsMoveGridCell(mshBill, .ColIndex("卫材信息"), -1, True, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub mshBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        Select Case .Col
            Case .ColIndex("卫材信息")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m文本式
            Case .ColIndex("产地"), .ColIndex("批号")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m文本式
            Case .ColIndex("效期"), .ColIndex("生产日期"), .ColIndex("灭菌日期")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m文本式
            Case .ColIndex("采购价"), .ColIndex("采购金额"), .ColIndex("售价"), _
                 .ColIndex("数量"), .ColIndex("零售价"), .ColIndex("冲销数量")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m金额式
                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                Select Case .Col
                    Case .ColIndex("数量"), .ColIndex("冲销数量")
                        intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
                    Case .ColIndex("采购价")
                       intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.成本价小数, g_小数位数.obj_散装小数.成本价小数)
                    Case .ColIndex("采购金额")
                        intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.金额小数, g_小数位数.obj_散装小数.金额小数)
                    Case .ColIndex("零售价"), .ColIndex("售价")
                        intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.零售价小数, g_小数位数.obj_散装小数.零售价小数)
                End Select
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Case .ColIndex("商品条码")
                Select Case KeyAscii
                    Case vbKeyBack, vbKeyEscape, 3, 22
                        Exit Sub
                    Case vbKeyReturn
'                        Call OS.PressKey(vbKeyTab)
                        Exit Sub
                    Case Else
                        '仅能录入数字和字母
                        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Then Exit Sub
                End Select
                KeyAscii = 0
        End Select
    End With
End Sub
Private Sub mshBill_LeaveCell()
    If mblnSort Then Exit Sub
    OS.OpenIme False
End Sub
Private Sub mshBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '设置单元格的编辑长度
        With mshBill
           Select Case .Col
               Case .ColIndex("卫材信息")
                   .EditMaxLength = 40
               Case .ColIndex("产地")
                   .EditMaxLength = 30
               Case .ColIndex("批号")
                   .EditMaxLength = mintBatchNoLen
              Case .ColIndex("效期")
                   .EditMaxLength = 10
               Case .ColIndex("生产日期")
                   .EditMaxLength = 10
               Case .ColIndex("灭菌日期")
                    .EditMaxLength = 10
               Case .ColIndex("采购价"), .ColIndex("采购金额"), .ColIndex("售价"), .ColIndex("零售价"), .ColIndex("数量"), .ColIndex("冲销数量")
                   .EditMaxLength = 16
           End Select
    End With
End Sub

Private Sub mshbill_EnterCell()
    If mblnSort = True Then Exit Sub
    '新增或修改才存在设置
    If mint编辑状态 <> 1 And mint编辑状态 <> 2 Then Exit Sub
    With mshBill
        SetInputFormat .Row
        OS.OpenIme (False)
        Select Case .Col
        Case .ColIndex("卫材信息")
             .ColComboList(.Col) = "..."
            '只在诊疗id列才显示合计信息和库存数
            Call 显示合计金额
            Call 提示库存数
        Case .ColIndex("效期"), .ColIndex("灭菌日期"), .ColIndex("生产日期")
            .ColComboList(.Col) = "..."
            If .ColIndex("效期") = .Col Then
                If Trim(.TextMatrix(.Row, .Col)) <> "" Then Exit Sub
                Dim str生产日期 As String, strxq As String
                If Not IsDate(.TextMatrix(.Row, .ColIndex("生产日期"))) Then
                    str生产日期 = ""
                Else
                    str生产日期 = Format(.TextMatrix(.Row, .ColIndex("生产日期")), "yyyymmdd")
                End If
                
                If str生产日期 <> "" And Trim(.TextMatrix(.Row, .ColIndex("原销期"))) <> "" Then
                    '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批

                    If IsNumeric(str生产日期) And Split(.TextMatrix(.Row, .ColIndex("原销期")), "||")(0) <> "0" Then
                        strxq = UCase(str生产日期)
'                            If Trim(.TextMatrix(.Row, mCol效期)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(.Row, .ColIndex("效期")) = Format(DateAdd("M", Split(.TextMatrix(.Row, .ColIndex("原销期")), "||")(0), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(.Row, .ColIndex("效期")))
                            End If
'                            End If
                    End If
                End If
             End If
        Case .ColIndex("产地")
            OS.OpenIme (True)
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub
 
 '从卫生材料信息中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, ByVal str诊疗id As String, ByVal str规格 As String, _
    ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal num指导批发价 As Double, ByVal str原产地 As String, _
    ByVal int原效期 As Integer, dbl比例系数 As Double, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal dbl指导差价率 As Double, ByVal str批准文号 As String) As Boolean
    Dim sng分段售价 As Double
    Dim intCount As Integer, intCol As Integer, lngDepartid As Long
    Dim rsTemp As New ADODB.Recordset
    Dim dbl成本价  As Double, dbl加成率 As Double, int库房分批 As Integer
    Dim str散装单位 As String
    Dim lngRow As Long
    
    On Error GoTo ErrHandle

    SetColValue = False
    lngDepartid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    gstrSQL = "SELECT a.加成率 from 材料特性 a where a.材料id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "加成率", lng材料ID)
    dbl加成率 = NVL(rsTemp!加成率, 0) / 100
        
    gstrSQL = "SELECT nvl(A.扣率,0) 扣率,A.灭菌效期,A.一次性材料,A.成本价,A.库房分批,A.注册证号,B.计算单位 散装单位,Nvl(A.是否条码管理,0) As 条码管理 " & _
              "From 材料特性 A, 收费项目目录 B Where a.材料ID=b.id and  A.材料id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    
    int库房分批 = Val(zlStr.NVL(rsTemp!库房分批))
    dbl成本价 = zlStr.NVL(rsTemp!成本价, 0)
    str散装单位 = zlStr.NVL(rsTemp!散装单位)
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> .ColIndex("行号") Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, .ColIndex("行号")) = intRow
        .TextMatrix(intRow, .ColIndex("材料ID")) = lng材料ID
        If Trim(.EditText) <> "" Then .EditText = str诊疗id
        .TextMatrix(intRow, .ColIndex("卫材信息")) = str诊疗id
        .TextMatrix(intRow, .ColIndex("规格")) = str规格
        .TextMatrix(intRow, .ColIndex("一次性材料")) = zlStr.NVL(rsTemp!一次性材料)
        .TextMatrix(intRow, .ColIndex("条码管理")) = zlStr.NVL(rsTemp!条码管理)
        .TextMatrix(intRow, .ColIndex("灭菌效期")) = zlStr.NVL(rsTemp!灭菌效期)
        .TextMatrix(intRow, .ColIndex("产地")) = IIf(IsNull(str产地), "", str产地)
        .TextMatrix(intRow, .ColIndex("批准文号")) = IIf(IsNull(str批准文号), "", str批准文号)
        .TextMatrix(intRow, .ColIndex("单位")) = str单位
        .TextMatrix(intRow, .ColIndex("售价")) = Format(num售价 * dbl比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, .ColIndex("原产地")) = IIf(IsNull(str原产地), "", str原产地)
        
        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
        .TextMatrix(intRow, .ColIndex("原销期")) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl加成率 & "||" & int是否变价 & "||" & int在用分批 & "||" & int库房分批
        .TextMatrix(intRow, .ColIndex("采购价")) = Format(num指导批发价 * dbl比例系数, mFMT.FM_成本价)
        .TextMatrix(intRow, .ColIndex("比例系数")) = dbl比例系数
        
        SetInputFormat intRow
        
        '说明：这里区分分批核算和不分批核算的目的是提高运行速度。
        '本来可以不分这些，直接用第一条SQL语句实现，但不分批的卫材就多在数据库中扫描一次。
        If Val(int库房分批) > 0 Then
'            If mintUnit = 1 Then
                gstrSQL = "" & _
                    "   Select 上次采购价,上次产地,上次生产日期 " & _
                    "   From 药品库存 " & _
                    "   Where 性质=1 and 库房id=[3] and 药品id=" & lng材料ID & _
                    "        and nvl(批次,0) =( select max(nvl(批次,0)) " & _
                    "                           from 药品库存 " & _
                    "                           where 性质=1 and 库房id=[1] and 药品id=[2] )"
'            Else
'            End If
        Else
            gstrSQL = "" & _
                "   Select 上次采购价,上次产地,上次生产日期 " & _
                "   From 药品库存 " & _
                "   Where 性质=1 and 库房id=[1] and 药品id=[2]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--取上次采购价", lngDepartid, lng材料ID, lngDepartid)
        
        If Not rsTemp.EOF Then
            If .TextMatrix(intRow, .ColIndex("产地")) = "" Then
                .TextMatrix(intRow, .ColIndex("产地")) = IIf(IsNull(rsTemp.Fields(1)), "", rsTemp.Fields(1))
            End If
            .TextMatrix(intRow, .ColIndex("采购价")) = Format(IIf(IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0)) * dbl比例系数 = 0, .TextMatrix(intRow, .ColIndex("采购价")), IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0)) * dbl比例系数), mFMT.FM_成本价)
            If IsNull(rsTemp!上次生产日期) Then
                .TextMatrix(intRow, .ColIndex("生产日期")) = ""
            Else
                .TextMatrix(intRow, .ColIndex("生产日期")) = Format(rsTemp!上次生产日期, "yyyy-mm-dd")
            End If
            
        Else
            If dbl成本价 <> 0 Then .TextMatrix(intRow, .ColIndex("采购价")) = Format(dbl成本价 * dbl比例系数, mFMT.FM_成本价)
        End If
        
        If .TextMatrix(intRow, .ColIndex("产地")) <> "" Then
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("产地")), lng材料ID)
            If rsTemp.RecordCount > 0 Then
               .TextMatrix(intRow, .ColIndex("批准文号")) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
            End If
        End If
        
        '时价材料处理
        If int是否变价 = 1 Then
            .TextMatrix(intRow, .ColIndex("售价")) = Format(校正零售价(sng分段售价 + _
                                                             时价材料零售价(lng材料ID, Val(.TextMatrix(intRow, .ColIndex("采购价"))), 0, -1, sng分段售价)) _
                                                             , mFMT.FM_零售价)
        End If
        .TextMatrix(intRow, .ColIndex("零售单位")) = str散装单位
        '刘兴宏:零售价处理
        Call 计算零售价及零售差价(intRow)
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
Private Sub mshBill_LostFocus()
    OS.OpenIme False
     Call zl_VsGridLOSTFOCUS(mshBill)
End Sub
Private Sub mshBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String, arrSplit As Variant
    Dim dbl加成率 As Double, sng分段售价 As Double, dbl采购价 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim dbl采购限价 As Double
    
    '数据验证
    On Error GoTo ErrHandle
    With mshBill
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
          Case .ColIndex("灭菌日期")
                '有处理
                If strKey = "" Then Exit Sub
                strKey = zlCheckIsDate(strKey, .ColKey(Col))
                If strKey = "" Then Cancel = True: Exit Sub
                
                If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("灭菌效期"))), CDate(strKey)), "yyyy-mm-dd") Then
                    If MsgBox("该卫材已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("灭菌效期"))), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
                '计算失效期
                .TextMatrix(Row, .ColIndex("灭菌失效期")) = Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("灭菌效期"))), CDate(strKey)), "yyyy-mm-dd")
                .EditText = strKey
           Case .ColIndex("生产日期"), .ColIndex("效期")
                '有处理
                If strKey = "" Then Exit Sub
                strKey = zlCheckIsDate(strKey, .ColKey(Col))
                If strKey = "" Then Cancel = True: Exit Sub
                .EditText = strKey
            Case .ColIndex("产地")
                '如果找不到对应的产地，则以输入做为产地
                If strKey = "" Then Exit Sub
                If zlCommFun.StrIsValid(strKey, .EditMaxLength, , .ColKey(Col)) = False Then
                    Cancel = True
                End If
            Case .ColIndex("批号")
                '如果找不到对应的产地，则以输入做为产地
                If strKey = "" Then Exit Sub
                If zlCommFun.StrIsValid(strKey, .EditMaxLength, , .ColKey(Col)) = False Then
                    Cancel = True
                End If
            Case .ColIndex("采购价")
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    '检查价格不能大于了指导批发价
                    gstrSQL = "Select nvl(a.指导批发价,0) as 指导批发价, b.现价" & vbNewLine & _
                                " From 材料特性 A, 收费价目 B" & vbNewLine & _
                                " Where a.材料id = b.收费细目id And Sysdate Between b.执行日期 And b.终止日期 And a.材料id = [1]" & _
                                GetPriceClassString("B")
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询指导批发价", Val(.TextMatrix(.Row, .ColIndex("材料ID"))))
                    
                    dbl采购限价 = Format(rsTemp!指导批发价 * Val(.TextMatrix(Row, .ColIndex("比例系数"))), mFMT.FM_成本价)
                    If mbln不强制控制指导价格 = False Then
                        If dbl采购限价 < Val(Format(Val(strKey), mFMT.FM_成本价)) Then
                            MsgBox "当前价格大于了指导批发价" & dbl采购限价 & "！", vbInformation, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = Format(Val(strKey), mFMT.FM_成本价)
                    
                    If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                        '对时价卫材的处理
                         '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                        arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                         If arrSplit(2) = 1 Then
                            '实价卫材处理
                             .EditText = Format(Val(strKey), mFMT.FM_成本价)
                            If mbln加价率 And mbln分段加成率 = False Then
                                If Show加成率(Col) = False Then Cancel = True: Exit Sub
                            Else
                                If mbln分段加成率 Then
                                    dbl加成率 = 0 'Get分段加成率(Val(strkey)) / 100
                                    If Get分段加成售价(Val(strKey), Val(.TextMatrix(Row, .ColIndex("比例系数"))), mstrCaption, sng分段售价) = False Then
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(sng分段售价 + _
                                                                          时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(strKey), dbl加成率, -1, sng分段售价)) _
                                                                          , mFMT.FM_零售价)
                                Else
                                    '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                                    dbl加成率 = Val(arrSplit(1))
                                    .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(strKey * (1 + dbl加成率) + _
                                                                          时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), strKey, dbl加成率)) _
                                                                          , mFMT.FM_零售价)
                                End If
                                If .TextMatrix(Row, .ColIndex("数量")) <> "" Then
                                    .TextMatrix(Row, .ColIndex("售价金额")) = Format(.TextMatrix(Row, .ColIndex("数量")) * .TextMatrix(Row, .ColIndex("售价")), mFMT.FM_金额)
                                End If
                            End If
                         Else
                            '定价检查成本价大于了售价提示
                            If Val(Format(rsTemp!现价, mFMT.FM_零售价)) < Val(Format(Val(strKey), mFMT.FM_成本价)) Then
                                MsgBox "当前价格大于了售价！", vbInformation, gstrSysName
                            End If
                         End If
                    End If
                End If
                '设置金额
                If strKey <> "" And strKey <> .TextMatrix(Row, .ColIndex("采购价")) And .TextMatrix(Row, .ColIndex("数量")) <> "" Then
                    .TextMatrix(Row, .ColIndex("采购金额")) = Format(.TextMatrix(Row, .ColIndex("数量")) * strKey, mFMT.FM_金额)
                    .TextMatrix(Row, .ColIndex("差价")) = Format(IIf(.TextMatrix(Row, .ColIndex("售价金额")) = "", 0, .TextMatrix(Row, .ColIndex("售价金额"))) - IIf(.TextMatrix(Row, .ColIndex("采购金额")) = "", 0, .TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                End If
       
                
            Case .ColIndex("采购金额")
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(Row, .ColIndex("采购金额")) Then
                     If .TextMatrix(Row, .ColIndex("数量")) <> "" Then
                            If mbln加价率 Then
                                '取得改变采购金额前的加价率
                                mdbl加价率 = 15
                                If Val(.TextMatrix(Row, .ColIndex("售价"))) <> 0 And Val(.TextMatrix(Row, .ColIndex("采购价"))) <> 0 Then
                                    mdbl加价率 = 计算加成率(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("售价"))), Val(.TextMatrix(Row, .ColIndex("采购价"))))
                                End If
                            End If
                            .TextMatrix(Row, .ColIndex("采购价")) = Format(strKey / .TextMatrix(Row, .ColIndex("数量")), mFMT.FM_成本价)
                            
                            '对时价卫材的处理
                            If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                                '重新计算零售价、差价
                                arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                                '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                                If Val(arrSplit(2)) = 1 Then
                                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                    If mbln加价率 And mbln分段加成率 = False Then
                                        .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(Val(.TextMatrix(Row, .ColIndex("采购价"))) * (1 + (mdbl加价率 / 100)) + _
                                                                              时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("采购价"))), (mdbl加价率 / 100))) _
                                                                              , mFMT.FM_零售价)
                                        .TextMatrix(Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(Row, .ColIndex("售价"))) * Val(.TextMatrix(Row, .ColIndex("数量"))), mFMT.FM_金额)
                                        .TextMatrix(Row, .ColIndex("差价")) = Format(IIf(.TextMatrix(Row, .ColIndex("售价金额")) = "", 0, .TextMatrix(Row, .ColIndex("售价金额"))) - IIf(.TextMatrix(Row, .ColIndex("采购金额")) = "", 0, .TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                                    Else
                                        Dim sng采购价 As Double
                                        sng采购价 = Val(.TextMatrix(Row, .ColIndex("采购价")))
    
                                        If mbln分段加成率 Then
                                            dbl加成率 = 0 ' Get分段加成率(Val(.TextMatrix(row, .colindex("采购价")))) / 100
                                            If Get分段加成售价(sng采购价, Val(.TextMatrix(Row, .ColIndex("比例系数"))), mstrCaption, sng分段售价) = False Then
                                                Cancel = True
                                                Exit Sub
                                            End If
                                            .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(sng分段售价 + _
                                                                                  时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), sng采购价, dbl加成率, -1, sng分段售价)) _
                                                                                  , mFMT.FM_零售价)
                                        Else
                                            '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                                            dbl加成率 = Val(arrSplit(1))
                                            .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(.TextMatrix(Row, .ColIndex("采购价")) * (1 + dbl加成率) + _
                                                                                  时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("采购价"))), dbl加成率)) _
                                                                                  , mFMT.FM_零售价)
                                        End If
                                        .TextMatrix(Row, .ColIndex("售价金额")) = Format(.TextMatrix(Row, .ColIndex("数量")) * .TextMatrix(Row, .ColIndex("售价")), mFMT.FM_金额)
                                    End If
                                End If
                            End If
                            .TextMatrix(Row, .ColIndex("差价")) = Format(Val(.TextMatrix(Row, .ColIndex("售价金额"))) - Val(strKey), mFMT.FM_金额)
                            .EditText = Format(strKey, mFMT.FM_金额)
                            .TextMatrix(Row, .ColIndex("采购金额")) = .EditText
                    End If
                End If
 
            Case .ColIndex("数量")
            
                If zlCommFun.DblIsValid(strKey, 16, True, True, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(strKey, mFMT.FM_数量)
                .EditText = strKey
                If .TextMatrix(Row, .ColIndex("采购价")) <> "" Then
                    .TextMatrix(Row, .ColIndex("采购金额")) = Format(Val(.TextMatrix(Row, .ColIndex("采购价"))) * Val(strKey), mFMT.FM_金额)
                    '时价卫材的处理
                    If .TextMatrix(Row, .ColIndex("原销期")) <> "" Then
                        arrSplit = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")
                        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                        If Val(arrSplit(2)) = 1 Then
                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                            If mbln加价率 Then
                                mdbl加价率 = Round(arrSplit(1) * 100, 2)
                                If Val(.TextMatrix(Row, .ColIndex("售价"))) <> 0 And Val(.TextMatrix(Row, .ColIndex("采购价"))) <> 0 Then
                                    mdbl加价率 = 计算加成率(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("售价"))), Val(.TextMatrix(Row, .ColIndex("采购价"))))
                                End If
                                .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(Val(.TextMatrix(Row, .ColIndex("采购价"))) * (1 + (mdbl加价率 / 100)) + _
                                                                      时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("采购价"))), (mdbl加价率 / 100))) _
                                                                      , mFMT.FM_零售价)
                                .TextMatrix(Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(Row, .ColIndex("售价"))) * strKey, mFMT.FM_金额)
                                .TextMatrix(Row, .ColIndex("差价")) = Format(IIf(.TextMatrix(Row, .ColIndex("售价金额")) = "", 0, .TextMatrix(Row, .ColIndex("售价金额"))) - IIf(.TextMatrix(Row, .ColIndex("采购金额")) = "", 0, .TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                            Else
                                If mbln分段加成率 Then
                                    dbl加成率 = 0 ' Get分段加成率(Val(.TextMatrix(row, .colindex("采购价")))) / 100
                                    If Get分段加成售价(Val(.TextMatrix(Row, .ColIndex("采购价"))), Val(.TextMatrix(Row, .ColIndex("比例系数"))), mstrCaption, sng分段售价) = False Then
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(sng分段售价 + _
                                                                          时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("采购价"))), dbl加成率, -1, sng分段售价)) _
                                                                          , mFMT.FM_零售价)
                                Else
                                    dbl加成率 = Split(.TextMatrix(Row, .ColIndex("原销期")), "||")(1)
                                    .TextMatrix(Row, .ColIndex("售价")) = Format(校正零售价(.TextMatrix(Row, .ColIndex("采购价")) * (1 + dbl加成率) + _
                                                                          时价材料零售价(Val(.TextMatrix(Row, .ColIndex("材料ID"))), Val(.TextMatrix(Row, .ColIndex("采购价"))), dbl加成率)) _
                                                                          , mFMT.FM_零售价)
                                End If
                            End If
                        End If
                    End If
                    If .TextMatrix(Row, .ColIndex("售价")) <> "" Then
                        .TextMatrix(Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(Row, .ColIndex("售价"))) * Val(strKey), mFMT.FM_金额)
                    End If
                    .TextMatrix(Row, .ColIndex("差价")) = Format(Val(.TextMatrix(Row, .ColIndex("售价金额"))) - Val(.TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                End If
            Case .ColIndex("冲销数量")
                
                If strKey = "" Then
                    MsgBox "冲销数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    
                    If Val(strKey) > Val(.TextMatrix(Row, .ColIndex("数量"))) Then
                        MsgBox "冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    
                    strKey = Format(Val(strKey), mFMT.FM_数量)
                    .Text = Val(strKey)
                    If .TextMatrix(Row, .ColIndex("采购价")) <> "" Then
                        .TextMatrix(Row, .ColIndex("采购金额")) = Format(Val(.TextMatrix(Row, .ColIndex("采购价"))) * Val(strKey), mFMT.FM_金额)
                    End If
                    If .TextMatrix(Row, .ColIndex("售价")) <> "" Then
                        .TextMatrix(Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(Row, .ColIndex("售价"))) * Val(strKey), mFMT.FM_金额)
                    End If
                    .TextMatrix(Row, .ColIndex("差价")) = Format(Val(.TextMatrix(Row, .ColIndex("售价金额"))) - Val(.TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                End If

            Case .ColIndex("售价")
                '检查条件:
                ' 1.售价不能大于指导零售价(根据参数:不强制控制指导价格决定)
                ' 2.检查了结算价与售价

                If Val(.TextMatrix(Row, .ColIndex("材料ID"))) = 0 Then Exit Sub
                
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                
                If strKey <> "" Then
                    
                    If mbln不强制控制指导价格 = False Then
                        '判断输入的零售价与指导零售价
                        gstrSQL = "Select 指导零售价 From 材料特性 Where 材料ID=[1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[读取指导零售价]", Val(.TextMatrix(Row, .ColIndex("材料ID"))))
                        Dim dbl指导零售价 As Double
                        dbl指导零售价 = Val(zlStr.NVL(rsTemp!指导零售价))
                        dbl指导零售价 = Val(Format(dbl指导零售价 * Val(.TextMatrix(Row, .ColIndex("比例系数"))), mFMT.FM_零售价))
                        If Val(Format(Val(strKey), mFMT.FM_零售价)) > dbl指导零售价 Then
                            ShowMsgBox "售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < Val(.TextMatrix(Row, .ColIndex("采购价"))) Then
                        If MsgBox("注意：" & vbCrLf & "     售价(￥" & Format(Val(strKey), mFMT.FM_零售价) & " 小于了" & vbCrLf & "     采购价（￥" & Format(Val(.TextMatrix(Row, .ColIndex("采购价"))), mFMT.FM_成本价) & "）,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                strKey = Format(Val(strKey), mFMT.FM_零售价)
                .EditText = strKey
                '重算差价
                .TextMatrix(Row, .ColIndex("售价金额")) = Format(Val(strKey) * Val(.TextMatrix(Row, .ColIndex("数量"))), mFMT.FM_金额)
                .TextMatrix(Row, .ColIndex("差价")) = Format(Val(.TextMatrix(Row, .ColIndex("售价金额"))) - Val(.TextMatrix(Row, .ColIndex("采购金额"))), mFMT.FM_金额)
        Case .ColIndex("零售价")
                '检查条件:
                ' 1.售价不能大于指导零售价(根据参数:不强制控制指导价格决定)
                ' 2.检查了结算价与售价
                If Val(.TextMatrix(Row, .ColIndex("材料ID"))) = 0 Then Exit Sub
                
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    If mbln不强制控制指导价格 = False Then
                        '判断输入的零售价与指导零售价
                        gstrSQL = "Select 指导零售价 From 材料特性 Where 材料ID=[1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[读取指导零售价]", Val(.TextMatrix(Row, .ColIndex("材料ID"))))
                        dbl指导零售价 = Val(zlStr.NVL(rsTemp!指导零售价))
                        dbl指导零售价 = Val(Format(dbl指导零售价, mFMT.FM_散装零售价))
                        If Val(Format(Val(strKey), mFMT.FM_散装零售价)) > dbl指导零售价 Then
                            ShowMsgBox "零售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    If Val(.TextMatrix(Row, .ColIndex("比例系数"))) = 0 Then
                        dbl采购价 = Val(.TextMatrix(Row, .ColIndex("采购价")))
                    Else
                        dbl采购价 = Val(.TextMatrix(Row, .ColIndex("采购价"))) / Val(.TextMatrix(Row, .ColIndex("比例系数")))
                    End If
                    
                    If Val(strKey) < dbl采购价 Then
                        If MsgBox("注意：" & vbCrLf & "     零售价(￥" & Format(Val(strKey), mFMT.FM_散装零售价) & " 小于了" & vbCrLf & "     结算价（￥" & Format(dbl采购价, mFMT.FM_成本价) & "）,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    strKey = Format(Val(strKey), mFMT.FM_散装零售价)
                    .EditText = strKey
                    .TextMatrix(.Row, .Col) = strKey
                    '刘兴宏:零售价处理
                    Call 计算零售价及零售差价(.Row, False)
                    If strKey <> "" Then
                        .TextMatrix(.Row, .ColIndex("售价")) = Format(Val(strKey) * Val(.TextMatrix(.Row, .ColIndex("比例系数"))), mFMT.FM_零售价)
                        .TextMatrix(.Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价"))) * Val(.TextMatrix(.Row, .ColIndex("数量"))), mFMT.FM_金额)
                        .TextMatrix(.Row, .ColIndex("差价")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价金额"))) - Val(.TextMatrix(.Row, .ColIndex("采购金额"))), mFMT.FM_金额)
                    End If
                    显示合计金额
                Else
                    strKey = Format(Val(strKey), mFMT.FM_散装零售价)
                    .EditText = strKey
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .ColIndex("产地")) = msh产地.TextMatrix(msh产地.Row, 2)
            
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("产地")), .ColIndex("材料ID"))
            If rsTemp.RecordCount Then
                .TextMatrix(.Row, .ColIndex("批准文号")) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
            Else
                .TextMatrix(.Row, .ColIndex("批准文号")) = ""
            End If
            
            msh产地.Visible = False
            .Col = .ColIndex("批号")
            .SetFocus
        End If
    End With
End Sub

Private Sub msh产地_LostFocus()
    If msh产地.Visible Then
        msh产地.Visible = False
    End If
End Sub
Private Function 当前仅为库房() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:判断当前库房仅为库房
    '入参:
    '出参:
    '返回:返回true表示仅为库房,否则为(发料部门或制剂室)
    '编制:刘兴洪
    '日期:2008-12-03 11:23:18
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From 部门性质说明 " & _
        "   WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) " & _
        "        AND 部门id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rsTemp.Fields(0) > 0 Then
        当前仅为库房 = False
        mbln库房 = False
    Else
        当前仅为库房 = True
        mbln库房 = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:验证数据的合法性
    '入参:
    '出参:
    '返回:变压器伸曲,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-03 09:49:18
    '-----------------------------------------------------------------------------------------------------------
    Dim intLop As Integer, rsTemp As New Recordset, blnStock As Boolean, arrSplit As Variant
    ValidData = False
    blnStock = 当前仅为库房()
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "单据号超长,最多能输入" & CInt(txtNO.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNO.MaxLength & "个字符!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, .ColIndex("材料ID")) <> "" Then        '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, .ColIndex("卫材信息"))) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, .ColIndex("数量")))) = "" Then
                        MsgBox "第" & intLop & "行卫材的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("数量")
                        Exit Function
                    End If
'
 
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, .ColIndex("批号")))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行卫材的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("批号")
                        Exit Function
                    End If
                    
                    If Len(Trim(.TextMatrix(intLop, .ColIndex("商品条码")))) > 50 Then
                        MsgBox "第" & intLop & "行卫材的商品条码超长,最多能输入50个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("商品条码")
                        Exit Function
                    End If
                    
                    If blnStock = True Then
                        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("原销期")) & "||||||||||", "||")
                        If Val(arrSplit(0)) <> 0 Then
                            If .TextMatrix(intLop, .ColIndex("批号")) = "" Or .TextMatrix(intLop, .ColIndex("效期")) = "" Then
                                MsgBox "第" & intLop & "行的卫材是效期卫材,请把它的批号及效期信息完整输入单据中！", vbInformation, gstrSysName
                                mshBill.SetFocus: .Row = intLop: .TopRow = intLop
                                If .TextMatrix(intLop, .ColIndex("批号")) = "" Then
                                    .Col = .ColIndex("批号")
                                Else
                                    .Col = .ColIndex("效期")
                                End If
                                Exit Function
                            End If
                        End If
                        
                        If Val(arrSplit(4)) <> 0 Then '库房分批
                            If mbln分批卫材批号产地控制 = True Then
                                If .TextMatrix(intLop, .ColIndex("产地")) = "" Or .TextMatrix(intLop, .ColIndex("批号")) = "" Then
                                    MsgBox "第" & intLop & "行的卫材是分批卫材,请把它的产地和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("产地")) = "" Then
                                        .Col = .ColIndex("产地")
                                    Else
                                        .Col = .ColIndex("批号")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If

                    Else '性质是“发料部门”
                        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("原销期")) & "||||||||||", "||")
                        If Val(arrSplit(3)) <> 0 Then '在用分批
                            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                            If arrSplit(0) <> "0" Then
                                If .TextMatrix(intLop, .ColIndex("批号")) = "" Or .TextMatrix(intLop, .ColIndex("效期")) = "" Then
                                    MsgBox "第" & intLop & "行的卫生材料是效期材料,请把它的批号及效期" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("批号")) = "" Then
                                        .Col = .ColIndex("批号")
                                    Else
                                        .Col = .ColIndex("效期")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    
                        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("原销期")) & "||||||||||", "||")
                        If Val(arrSplit(3)) <> 0 Then '在用分批
                            If mbln分批卫材批号产地控制 = True Then
                                If .TextMatrix(intLop, .ColIndex("产地")) = "" Or .TextMatrix(intLop, .ColIndex("批号")) = "" Then
                                    MsgBox "第" & intLop & "行的卫材是分批卫材,请把它的产地和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("产地")) = "" Then
                                        .Col = .ColIndex("产地")
                                    Else
                                        .Col = .ColIndex("批号")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("采购价"))) > 9999999999# Then
                        MsgBox "  第" & intLop & "行卫材的采购价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("采购价")
                        mshBill.SetFocus
                        Exit Function
                    End If
                    
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("零售价")), 16, False, False, , "第" & intLop & "行卫生材料的零售价") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("零售价")
                        Exit Function
                    End If
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("零售金额")), 16, False, False, , "第" & intLop & "行卫生材料的零售金额") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("零售价")
                        Exit Function
                    End If
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("零售差价")), 16, False, False, , "第" & intLop & "行卫生材料的零售差价") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("零售价")
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("数量"))) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("数量")
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("采购金额"))) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的采购金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("采购金额")
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, .ColIndex("售价金额"))) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("数量")
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
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存其他入库单信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-03 10:00:08
    '-----------------------------------------------------------------------------------------------------------

    Dim chrNo As Variant, cllPro As New Collection
    Dim lng序号 As Long, lng库房id As Long, lng入出类别ID As Long, lng材料ID As Long, intRow As Integer
    Dim dbl数量 As Double, dbl采购价 As Double, dbl采购金额 As Double
    Dim dbl零售金额 As Double, dbl差价 As Double, str零售差价 As String, dbl零售价 As Double
    Dim str摘要 As String, str填制人 As String, str填制日期 As String, str生产日期 As String
    Dim str审核人 As String, str灭菌日期 As String, str灭菌效期 As String
    Dim str批号 As String, str产地 As String, str效期 As String, str商品条码 As String
    Dim str批准文号 As String
    Dim n As Long
    
    
    SaveCard = False
    With mshBill
        chrNo = Trim(txtNO)
        lng库房id = cboStock.ItemData(cboStock.ListIndex)
        If mint编辑状态 = 1 Then   'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(70, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(70, lng库房id)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        lng入出类别ID = cboType.ItemData(cboType.ListIndex)
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str审核人 = Txt审核人
        
        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_材料其他入库_Delete('" & mstr单据号 & "')"
            AddArray cllPro, gstrSQL
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("材料ID")) <> "" Then
                lng材料ID = Val(.TextMatrix(intRow, .ColIndex("材料ID")))
                str产地 = .TextMatrix(intRow, .ColIndex("产地"))
                str批准文号 = .TextMatrix(intRow, .ColIndex("批准文号"))
                str批号 = .TextMatrix(intRow, .ColIndex("批号"))
                str效期 = IIf(.TextMatrix(intRow, .ColIndex("效期")) = "", "", .TextMatrix(intRow, .ColIndex("效期")))
                
                dbl数量 = Round(Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_最大小数.数量小数)
                dbl采购价 = Round(Val(.TextMatrix(intRow, .ColIndex("采购价"))) / Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_最大小数.成本价小数)
                dbl采购金额 = Round(Val(.TextMatrix(intRow, .ColIndex("采购金额"))), g_小数位数.obj_最大小数.金额小数)
                
                
   
                '刘兴宏:零售价处理
                '数据库中的:差价 = 零售金额 - 结算金额
                '数据库中的:用法 = 零售金额-售价金额或零售差价-差价(库房单位的差价)

                dbl零售价 = Round(Val(.TextMatrix(intRow, .ColIndex("零售价"))), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, .ColIndex("零售金额"))), g_小数位数.obj_最大小数.零售价小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, .ColIndex("零售差价"))), g_小数位数.obj_最大小数.零售价小数)
                str零售差价 = Round(Val(.TextMatrix(intRow, .ColIndex("零售差价"))) - Val(.TextMatrix(intRow, .ColIndex("差价"))), g_小数位数.obj_最大小数.零售价小数)
'                dbl售价 = Round(Val(.TextMatrix(intRow, .ColIndex("售价"))) / Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_散装小数.零售价小数)
'                dbl零售金额 = Round(Val(.TextMatrix(intRow, .ColIndex("售价金额"))), g_小数位数.obj_散装小数.金额小数)
'                dbl差价 = Round(Val(.TextMatrix(intRow, .ColIndex("差价"))), g_小数位数.obj_散装小数.金额小数)
                
                str生产日期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("生产日期")) = "", "", .TextMatrix(intRow, .ColIndex("生产日期"))))
                str灭菌日期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("灭菌日期")) = "", "", .TextMatrix(intRow, .ColIndex("灭菌日期"))))
                str灭菌效期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("灭菌失效期")) = "", "", .TextMatrix(intRow, .ColIndex("灭菌失效期"))))
                If gblnCode = True Then str商品条码 = Trim(IIf(.TextMatrix(intRow, .ColIndex("商品条码")) = "", "", .TextMatrix(intRow, .ColIndex("商品条码"))))
                
                lng序号 = intRow
                
                'Zl_材料其他入库_Insert
                gstrSQL = "zl_材料其他入库_INSERT("
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lng库房id & ","
                '  入出类别id_In In 药品收发记录.入出类别id%Type,
                gstrSQL = gstrSQL & "" & lng入出类别ID & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  实际数量_In   In 药品收发记录.实际数量%Type,
                gstrSQL = gstrSQL & "" & dbl数量 & ","
                '  成本价_In     In 药品收发记录.成本价%Type,
                gstrSQL = gstrSQL & "" & dbl采购价 & ","
                '  成本金额_In   In 药品收发记录.成本金额%Type,
                gstrSQL = gstrSQL & "" & dbl采购金额 & ","
                '  零售价_In     In 药品收发记录.零售价%Type,
                gstrSQL = gstrSQL & "" & dbl零售价 & ","
                '  零售金额_In   In 药品收发记录.零售金额%Type,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '  差价_In       In 药品收发记录.差价%Type,
                gstrSQL = gstrSQL & "" & dbl差价 & ","
                '  零售差价_In   In 药品收发记录.差价%Type,
                gstrSQL = gstrSQL & "" & str零售差价 & ","
                '  填制人_In     In 药品收发记录.填制人%Type,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '  填制日期_In   In 药品收发记录.填制日期%Type,
                gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
                '  摘要_In       In 药品收发记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "'" & str摘要 & "',"
                '  产地_In       In 药品收发记录.产地%Type := Null,
                gstrSQL = gstrSQL & "'" & str产地 & "',"
                '  批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '  生产日期_In   In 药品收发记录.生产日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str生产日期 = "", "Null", "to_date('" & Format(str生产日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  商品条码_In   In 药品收发记录.商品条码%Type := Null
                gstrSQL = gstrSQL & "'" & str商品条码 & "',"
                '  批准文号_In   In 药品收发记录.批准文号%Type := Null
                gstrSQL = gstrSQL & IIf(str批准文号 = "", "NULL", "'" & str批准文号 & "'")
                gstrSQL = gstrSQL & ")"
                AddArray cllPro, gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    err = 0: On Error GoTo ErrHandle:
    ExecuteProcedureArrAy cllPro, mstrCaption, True
    If Not 检查单价(17, txtNO.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    gcnOracle.CommitTrans
    
    mblnSave = True: mblnSuccess = True: mblnChange = False: SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:冲销其他入库单
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-03 10:44:27
    '-----------------------------------------------------------------------------------------------------------

    Dim cllProc As New Collection, bln全冲 As Boolean
    Dim lng行次 As Integer, lng原记录状态 As Integer, lng序号 As Integer, intRow As Integer
    Dim lng材料ID As Long, dbl冲销数量 As Double, dbl实际数量 As Double, dbl零售价 As Double
    Dim chrNo As String, str填制人 As String, str填制日期 As String
    Dim int库存检查 As Integer, lng库房id As Long, lng批次 As Long
    Dim str摘要 As String
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '检查冲销数量，不能小于零
        lng库房id = cboStock.ItemData(cboStock.ListIndex)
        int库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        chrNo = Trim(txtNO.Tag)
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("冲销数量"))) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, .ColIndex("数量"))), Val(.TextMatrix(intRow, .ColIndex("冲销数量")))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
                If int库存检查 <> 0 Then
                    dbl冲销数量 = Round(Val(.TextMatrix(intRow, .ColIndex("冲销数量"))) * Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_散装小数.数量小数)
                    dbl实际数量 = Round(Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_散装小数.数量小数)
                    bln全冲 = (dbl冲销数量 = dbl实际数量)
                    If bln全冲 Then
                        dbl冲销数量 = Val(.TextMatrix(intRow, .ColIndex("真实数量")))
                    End If
                    lng批次 = 取单据批次(17, chrNo, Val(.TextMatrix(intRow, .ColIndex("材料ID"))), Val(.TextMatrix(intRow, .ColIndex("序号"))))
                    If Check可用数量(lng库房id, Val(.TextMatrix(intRow, .ColIndex("材料ID"))), lng批次, dbl冲销数量, int库存检查) = False Then Exit Function
                End If
            End If
        Next
    
        str填制人 = UserInfo.用户名
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        lng原记录状态 = mint记录状态
        
        lng行次 = 0
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("材料ID")) <> "" And Val(.TextMatrix(intRow, .ColIndex("冲销数量"))) <> 0 Then
                lng行次 = lng行次 + 1
                
                lng材料ID = .TextMatrix(intRow, .ColIndex("材料ID"))
                dbl冲销数量 = Round(Val(.TextMatrix(intRow, .ColIndex("冲销数量"))) * Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_散装小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("比例系数"))), g_小数位数.obj_散装小数.数量小数)
                str摘要 = txt摘要.Text
                
                bln全冲 = (dbl冲销数量 = dbl实际数量)
                lng序号 = Val(.TextMatrix(intRow, .ColIndex("序号")))
                'Zl_材料其他入库_Strike
                gstrSQL = "Zl_材料其他入库_Strike("
                '  行次_In       In Integer,
                gstrSQL = gstrSQL & "" & lng行次 & ","
                '  原记录状态_In In 药品收发记录.记录状态%Type,
                gstrSQL = gstrSQL & "" & lng原记录状态 & ","
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  冲销数量_In   In 药品收发记录.实际数量%Type,
                gstrSQL = gstrSQL & "" & dbl冲销数量 & ","
                '  填制人_In     In 药品收发记录.填制人%Type,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '  填制日期_In   In 药品收发记录.填制日期%Type,
                gstrSQL = gstrSQL & "to_date('" & Format(str填制日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
                '  全部冲销_In   In 药品收发记录.实际数量%Type := 0 --1-全部冲销,0-部分冲销
                gstrSQL = gstrSQL & "" & IIf(bln全冲, 1, 0)
                gstrSQL = gstrSQL & ",'" & str摘要 & "')"
                AddArray cllProc, gstrSQL
            End If
            
            recSort.MoveNext
        Next
        If lng行次 = 0 Then
            MsgBox "没有选择一行卫材来冲销，不能冲销，请检查！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    err = 0: On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllProc, mstrCaption
    mblnSave = True: mblnSuccess = True: mblnChange = False
    
    SaveStrike = True
    Exit Function
ErrHandle:
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
            curTotal = curTotal + Val(.TextMatrix(intLop, .ColIndex("采购金额")))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, .ColIndex("售价金额")))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
    
    
End Sub

Private Sub 提示库存数()
    Dim recTmp As New ADODB.Recordset
    Dim dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("卫材信息")) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("材料ID")) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("材料ID"))
    Select Case mintUnit
        Case 0
            strQuantity = "a.可用数量"
        Case Else
            strQuantity = "a.可用数量/b.换算系数 "
    End Select
        
    gstrSQL = "" & _
        "   Select b.材料ID," & IIf(mintUnit = 0, "M.计算单位", "b.包装单位") & " as 单位, Sum(" & strQuantity & ") as 数量 " & _
        "   From 药品库存 a,材料特性 b,收费项目目录 M " & _
        "   Where a.性质=1 and a.药品id=b.材料id and a.药品id=M.id and a.可用数量<>0 And " & _
        "           a.库房ID=[1]" & _
        "           and b.材料ID=[2]" & _
        "   Group by b.材料ID," & IIf(mintUnit = 0, "m.计算单位", "b.包装单位")
    Set recTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提示库存数", cboStock.ItemData(cboStock.ListIndex), intID)
        
    With recTmp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        dbl数量 = IIf(IsNull(!数量), 0, !数量)
        
        stbThis.Panels(2).Text = "该卫材当前库存数为[" & Format(dbl数量, mFMT.FM_数量) & "]" & zlStr.NVL(!单位)
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
    Dim strUnit As String
    Dim int单位系数 As Integer
    Dim strNo As String
    
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1714", mint记录状态, mintUnit, 1714, "卫材其它入库单", strNo
End Sub

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
        
    zlDatabase.OpenRecordset rsBatchNolen, gstrSQL, "取字段长度"
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function 时价材料零售价(ByVal lng材料ID As Long, ByVal sin采购价 As Double, ByVal sin加成率 As Double, _
    Optional LngLastRow As Long = -1, Optional sng售价 As Double = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '功能:根据指导价格或差价比计算出时价材料的差价让利情况
    '入参:lng材料ID-材料ID
    '     sin采购价-采购价格
    '     sin加成率-加成率(如果传入0,同时又传入dbl零售价,则将按传入的零售价进行计算)
    '     LngLastRow-单据的行号
    '     sng售价-传入的零售价
    '出参:
    '返回:零售价的让利情况
    '修改人:刘兴宏
    '修改时间:2007/2/25
    '------------------------------------------------------------------------------------------------------
    '时价材料零售价计算公式:采购价*(1+加成率)
    '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    
    Dim sin零售价 As Double, sin指导零售价 As Double, sin差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", lng材料ID)
    
    If rsTemp.EOF Then Exit Function
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    
    时价材料零售价 = 0
    If sin差价让利比 = 100 Then Exit Function
    If sin指导零售价 = 0 Then Exit Function
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    sin零售价 = sin采购价 * (1 + sin加成率)
    If sin零售价 / Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("比例系数"))) >= sin指导零售价 Then Exit Function
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("比例系数")))
    时价材料零售价 = (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 计算加成率(ByVal lng材料ID As Long, ByVal sin零售价 As Double, ByVal sin成本价 As Double) As Double
    Dim sin指导零售价 As Double, sin差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    '根据零售价反算成本价,由于时价材料公式的变化,导致原来计算加成率的公式无效,需重新计算
    '原公式:(零售价/成本价-1)*100
    '现公式的理论:由于零售价是按加成率算出来后,再加上了让利外那部分金额,因此实际按加成率算出的零售价=指导零售价-(指导零售价-零售价)/差价让利比
    '再套用原公式算出实际的加成率
    计算加成率 = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select A.指导零售价,Nvl(A.差价让利比,100) 差价让利比,Nvl(B.是否变价,0) 时价 From 材料特性 A,收费项目目录 B Where A.材料id=B.id and A.材料ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", lng材料ID)
    
    If rsTemp.EOF Then Exit Function
    
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    If rsTemp!时价 = 0 Then Exit Function
'   If mbln分段加成率 Then
'            计算加成率 = Get分段加成率(sin成本价)
'   Else
        '指导零售价-(指导零售价-零售价)/差价让利比
        sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("比例系数")))
        If sin差价让利比 <> 100 And sin差价让利比 > 0 Then
            sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价) / sin差价让利比 * 100
        Else
            sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价)
        End If
        计算加成率 = (sin零售价 / sin成本价 - 1) * 100
'    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 校正零售价(ByVal sin零售价 As Double, Optional LngLastRow As Long = -1) As Double
    '得到按当前单位系数计算出来的指导零售价，如果时价卫材强制控制指导价计算出来的零售价大于指导零售价，以指导零售价为准
    Dim sin指导零售价 As Double
    Dim rsTemp As New ADODB.Recordset
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("材料ID"))))
    
    If rsTemp.EOF Then Exit Function
    sin指导零售价 = zlStr.NVL(rsTemp!指导零售价, mshBill.ColIndex("材料ID"))
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("比例系数")))
    If sin指导零售价 = 0 Then sin指导零售价 = sin零售价
    校正零售价 = IIf(sin零售价 > sin指导零售价 And Not mbln不强制控制指导价格, sin指导零售价, sin零售价)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 计算零售价及零售差价(ByVal lngRow As Long, Optional bln零售价 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据库房单位计算散装单位的零售价及金额
    '入参:lngRow -指定计算的行
    '     bln零售价-零售价为售价
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-28 12:09:04
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl比例系数 As Double, arrSplit As Variant
    Dim dbl数量 As Double
    
    With mshBill
        dbl比例系数 = Val(.TextMatrix(lngRow, mshBill.ColIndex("比例系数")))
        dbl数量 = IIf(mint编辑状态 = 6, Val(.TextMatrix(lngRow, .ColIndex("冲销数量"))), Val(.TextMatrix(lngRow, .ColIndex("数量"))))
        If dbl数量 = 0 Or Val(.TextMatrix(lngRow, .ColIndex("材料ID"))) = 0 Then
            .TextMatrix(lngRow, .ColIndex("零售金额")) = 0
            .TextMatrix(lngRow, .ColIndex("零售差价")) = 0
            .TextMatrix(lngRow, .ColIndex("差价")) = 0
            .TextMatrix(lngRow, .ColIndex("售价金额")) = 0
            .TextMatrix(lngRow, .ColIndex("零售价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价"))) / IIf(dbl比例系数 = 0, 1, dbl比例系数), mFMT.FM_散装零售价)
            Exit Function
        End If
        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
        If .TextMatrix(lngRow, .ColIndex("原销期")) <> "" Then
           arrSplit = Split(.TextMatrix(lngRow, .ColIndex("原销期")), "||")
           If Val(arrSplit(2)) = 1 And (IIf(mbln库房, arrSplit(4) = 1, arrSplit(3) = 1)) Then
                '实价卫材
                '刘兴宏:零售价处理
                If bln零售价 Then
                    .TextMatrix(lngRow, .ColIndex("零售价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价"))) / dbl比例系数, mFMT.FM_散装零售价)
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("零售价"))) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("零售价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价"))) / dbl比例系数, mFMT.FM_散装零售价)
                End If
                .TextMatrix(lngRow, .ColIndex("零售金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
                '零售差价=零售金额-结算金额
                .TextMatrix(lngRow, .ColIndex("零售差价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售金额"))) - Val(.TextMatrix(lngRow, .ColIndex("采购金额"))), mFMT.FM_金额)
           Else '定价
                .TextMatrix(lngRow, .ColIndex("零售价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价"))) / dbl比例系数, mFMT.FM_散装零售价)
                .TextMatrix(lngRow, .ColIndex("零售金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
                '零售差价=零售金额-结算金额
                .TextMatrix(lngRow, .ColIndex("零售差价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售金额"))) - Val(.TextMatrix(lngRow, .ColIndex("采购金额"))), mFMT.FM_金额)
           End If
        Else
                .TextMatrix(lngRow, .ColIndex("零售价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("售价"))) / dbl比例系数, mFMT.FM_散装零售价)
                .TextMatrix(lngRow, .ColIndex("零售金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
                '零售差价=零售金额-结算金额
                .TextMatrix(lngRow, .ColIndex("零售差价")) = Format(Val(.TextMatrix(lngRow, .ColIndex("零售金额"))) - Val(.TextMatrix(lngRow, .ColIndex("采购金额"))), mFMT.FM_金额)
        End If
    End With
    计算零售价及零售差价 = True
End Function

Private Sub AfterDeleteRow()
    '删除行后
    
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mshBill.ColIndex("行号"), mshBill.Row)
End Sub
Private Function Select卫材信息(Optional strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:卫生材料信息选择
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-12-02 11:50:35
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, lng库房id As Long
    Dim sngLeft As Single, sngTop As Single
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    With mshBill
        lng库房id = cboStock.ItemData(cboStock.ListIndex)
        If strSearch = "" Then
            Set rsTemp = Frm材料选择器.ShowMe(Me, 1, , lng库房id, lng库房id, , , , , , , , , , , 1714, , mstrPrivs, , False)
        Else
            Call CalcPosition(sngLeft, sngTop, mshBill)
            Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, lng库房id, lng库房id, lng库房id, strSearch, sngLeft, sngTop, mshBill.CellWidth, mshBill.CellHeight, , , , , , , , , , 1714, , mstrPrivs, , False)
        End If
        
        If rsTemp.RecordCount <= 0 Then Exit Function
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            SetColValue .Row, rsTemp!材料ID, _
                "[" & rsTemp!编码 & "]" & rsTemp!名称, IIf(IsNull(rsTemp!规格), "", rsTemp!规格), _
                IIf(IsNull(rsTemp!产地), "", rsTemp!产地), IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
                IIf(IsNull(rsTemp!售价), 0, rsTemp!售价), rsTemp!指导批发价 / IIf(mintUnit = 0, 1, rsTemp!换算系数), _
                IIf(IsNull(rsTemp!产地), "!", rsTemp!产地), rsTemp!最大效期, IIf(mintUnit = 0, 1, rsTemp!换算系数), _
                rsTemp!时价, rsTemp!在用分批, rsTemp!指导差价率 / 100, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
            
            
            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Next
        
        mshBill.Row = int点击行
        
'        If rsTemp.RecordCount = 1 Then
'            SetColValue .Row, rsTemp!材料ID, _
'                "[" & rsTemp!编码 & "]" & rsTemp!名称, IIf(IsNull(rsTemp!规格), "", rsTemp!规格), _
'                IIf(IsNull(rsTemp!产地), "", rsTemp!产地), IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
'                IIf(IsNull(rsTemp!售价), 0, rsTemp!售价), rsTemp!指导批发价 / IIf(mintUnit = 0, 1, rsTemp!换算系数), _
'                IIf(IsNull(rsTemp!产地), "!", rsTemp!产地), rsTemp!最大效期, IIf(mintUnit = 0, 1, rsTemp!换算系数), _
'                rsTemp!时价, rsTemp!在用分批, rsTemp!指导差价率 / 100, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
'        End If
        rsTemp.Close
        Select卫材信息 = True
    End With
    Call 提示库存数
End Function
 
Private Function SelDate(ByVal intCol As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择日期
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, blnreturn As Boolean
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim strMaxDate As String
    Dim lngRow As Long
    With mshBill
        strDate = .TextMatrix(.Row, intCol)
        If strDate = "" Then strDate = Format(sys.Currentdate, "yyyy-mm-dd")
        lngH = .CellHeight
        Call CalcPosition(sngX, sngY, mshBill)
        If intCol = .ColIndex("效期") Then strMaxDate = "3000-01-01"
        If intCol = .ColIndex("灭菌失效期") Then strMaxDate = "3000-01-01"
        If intCol = .ColIndex("生产日期") Then strMaxDate = Format(sys.Currentdate, "yyyy-mm-dd")
    End With
    blnreturn = frmDateSel.SelectDate(Me, sngX, sngY, lngH, strDate, , strMaxDate)
    If blnreturn = False Then Exit Function
    With mshBill
        .TextMatrix(.Row, intCol) = strDate
    End With
    zlVsMoveGridCell mshBill, mshBill.ColIndex("卫材信息"), 0, True, lngRow
    SelDate = True
End Function

Private Function Show加成率(ByVal intCol As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:显示加成率
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim blnreturn As Boolean, dbl零售价 As Double, dbl结算价 As Double, dbl加成率 As Double, lng材料ID As Long
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim dbl原加成率 As Double
    
    dbl加成率 = 15
    With mshBill
        lng材料ID = Val(.TextMatrix(.Row, .ColIndex("材料ID")))
        If lng材料ID = 0 Then Exit Function
        
        dbl零售价 = Val(.TextMatrix(.Row, .ColIndex("售价"))) '
        If intCol = .ColIndex("采购价") Then
            dbl结算价 = Val(.EditText)
        Else
            dbl结算价 = Val(.TextMatrix(.Row, .ColIndex("采购价")))
        End If
        If dbl零售价 <> 0 And dbl结算价 <> 0 Then
            dbl加成率 = Format(计算加成率(lng材料ID, dbl零售价, dbl结算价), "####0.0000000;-###0.0000000;0;0")
        End If
        lngH = .CellHeight
        Call CalcPosition(sngX, sngY, mshBill)
    End With
    dbl原加成率 = dbl加成率
    
    blnreturn = frm扣率Set.ShowCalc(Me, sngX, sngY, lngH, lng材料ID, mintUnit, dbl零售价, dbl结算价, dbl加成率, mbln不强制控制指导价格)
    With mshBill
        If blnreturn = False Then
            mdbl加价率 = dbl原加成率
            '重新计算零售价、差价
            .TextMatrix(.Row, .ColIndex("售价")) = Format(Val(.TextMatrix(.Row, .ColIndex("采购价"))) * (1 + (mdbl加价率 / 100)), mFMT.FM_零售价)
            .TextMatrix(.Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价"))) * Val(.TextMatrix(.Row, .ColIndex("数量"))), mFMT.FM_金额)
            .TextMatrix(.Row, .ColIndex("差价")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价金额"))) - Val(.TextMatrix(.Row, .ColIndex("采购金额"))), mFMT.FM_金额)
            Exit Function
        End If
        .TextMatrix(.Row, .ColIndex("售价")) = Format(dbl零售价, mFMT.FM_零售价)
        .TextMatrix(.Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价"))) * Val(.TextMatrix(.Row, .ColIndex("数量"))), mFMT.FM_金额)
        .TextMatrix(.Row, .ColIndex("差价")) = Format(Val(.TextMatrix(.Row, .ColIndex("售价金额"))) - Val(.TextMatrix(.Row, .ColIndex("采购金额"))), mFMT.FM_金额)
        mdbl加价率 = dbl加成率
    End With
    Show加成率 = True
    'debug.Print "aaa"
End Function

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mshBill.ColIndex("序号"))) = 0, n, Val(mshBill.TextMatrix(n, mshBill.ColIndex("序号"))))
                !药品id = Val(mshBill.TextMatrix(n, mshBill.ColIndex("材料id")))
                
                .Update
            End If
        Next
        
    End With
End Sub
