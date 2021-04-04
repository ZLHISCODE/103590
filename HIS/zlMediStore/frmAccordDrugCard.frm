VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAccordDrugCard 
   AutoRedraw      =   -1  'True
   Caption         =   "协定药品入库单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmAccordDrugCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   5970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   6
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDrug 
         Height          =   3000
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Visible         =   0   'False
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5292
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
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   1230
         Left            =   195
         TabIndex        =   2
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
         TabIndex        =   4
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
      Begin ZL9BillEdit.BillEdit mshStructure 
         Height          =   1875
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   3307
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
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7230
         TabIndex        =   21
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   20
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Top             =   593
         Width           =   1425
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
         TabIndex        =   16
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4995
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "协定药品入库单"
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   5340
         Width           =   720
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
            Picture         =   "frmAccordDrugCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":01A8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0206
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0264
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":02C2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0320
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":037E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":03DC
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
            Picture         =   "frmAccordDrugCard.frx":043A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0498
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":04F6
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0554
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":05B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":0610
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":066E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccordDrugCard.frx":06CC
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
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
            Picture         =   "frmAccordDrugCard.frx":072A
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
            Picture         =   "frmAccordDrugCard.frx":0FBE
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAccordDrugCard.frx":14C0
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
      Caption         =   "编码"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "列名"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(编码和名称)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅编码)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅名称)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAccordDrugCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnChange As Boolean               '是否进行过编辑
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Dim mstrPrivs As String                     '权限
Private mbln下可用数量 As Boolean           '填单是否下可用数量

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录

Private mlng库房id As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

Private Const MStrCaption As String = "协定药品入库"

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4


Private mcolUseCount As Collection

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

'=========================================================================================
Private Const mconIntCol药名 As Integer = 1
Private Const mconIntCol商品名 As Integer = 2
Private Const mconIntCol来源 As Integer = 3
Private Const mconIntCol基本药物 As Integer = 4
Private Const mconIntCol规格 As Integer = 5
Private Const mconIntCol比例系数 As Integer = 6
Private Const mconIntCol原销期 As Integer = 7
Private Const mconIntCol单位 As Integer = 8
Private Const mconIntCol数量 As Integer = 9
Private Const mconIntCol采购价 As Integer = 10
Private Const mconIntCol采购金额 As Integer = 11
Private Const mconIntCol售价 As Integer = 12
Private Const mconIntCol售价金额 As Integer = 13
Private Const mconintCol差价 As Integer = 14
Private Const mconIntCol药品编码和名称 As Integer = 15
Private Const mconIntCol药品编码 As Integer = 16
Private Const mconIntCol药品名称 As Integer = 17
Private Const mconIntColS As Integer = 18      '总列数
'=========================================================================================


'=========================================================================================
'构成药品各列
Private Const mconIntCol构药名 As Integer = 0
Private Const mconIntCol构商品名 As Integer = 1
Private Const mconIntCol构规格 As Integer = 2
Private Const mconIntCol构产地 As Integer = 3
Private Const mconIntCol构单位 As Integer = 4
Private Const mconIntCol构数量 As Integer = 5
Private Const mconIntCol构组成数量 As Integer = 6
Private Const mconIntCol构可用数量 As Integer = 7
Private Const mconIntcol加成率 As Integer = 8
Private Const mconintcol构实际差价 As Integer = 9
Private Const mconintcol构实际金额 As Integer = 10
Private Const mconintcol构药品id As Integer = 11

Private Const mconIntCol构采购价 As Integer = 12
Private Const mconIntCol构采购金额 As Integer = 13
Private Const mconIntCol构售价 As Integer = 14
Private Const mconIntCol构售价金额 As Integer = 15
Private Const mconintCol构差价 As Integer = 16
Private Const mconIntCol构药品编码和名称 As Integer = 17
Private Const mconIntCol构药品编码 As Integer = 18
Private Const mconIntCol构药品名称 As Integer = 19

Private Const mconInt构ColS As Integer = 20             '总列数
'=========================================================================================

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
                
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !药品ID = Val(mshBill.TextMatrix(n, 0))
                                
                .Update
            End If
        Next
    End With
End Sub
Private Sub GetSysParm()
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)

End Sub
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim int入系数 As Integer, int出系数 As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.系数, b.名称 " _
        & " FROM 药品单据性质 A, 药品入出类别 B " _
        & "Where A.类别id = B.ID " _
      & "AND A.单据 = 3  "
'    Call SQLTest(App.Title, "协定药品入库管理", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
'    Call SQLTest
    
    If rsDepend.EOF Then
        MsgBox "没有设置协定药品入库的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = "系数=-1"
    If rsDepend.EOF Then
        MsgBox "没有设置协定药品入库的出库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = "系数=1"
    If rsDepend.EOF Then
        MsgBox "没有设置协定药品入库的入库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Filter = adFilterNone
    rsDepend.Close
    
    
    gstrSQL = " SELECT a.药品id FROM 协定药品对照 a, 药品规格 b Where a.药品id = b.药品id "
'    Call SQLTest(App.Title, "协定药品入库管理", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
'    Call SQLTest
    
    
    If rsDepend.EOF Then
        MsgBox "没有一种具有协定药品对照的协定药品,请查看药品目录管理！", vbInformation, gstrSysName
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

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1344)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    If mint编辑状态 = 1 Then
    ElseIf mint编辑状态 = 2 Then
    ElseIf mint编辑状态 = 3 Then
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        CmdSave.Caption = "打印(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
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
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("如果改变库房，有可能要改变相应药品的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill

                    mlng库房id = .ItemData(.ListIndex)
                    Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                    
                    '组织格式化串
                    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
                    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
                    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
                    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
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
        FindRow mshBill, mconIntCol药品编码和名称, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
'    mblnChange = False
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
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zldatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    
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
        mstrTime_End = GetBillInfo(3, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If

        If SaveCheck = True Then
            If Val(zldatabase.GetPara("审核打印", glngSys, 1344)) = 1 Then
                '打印
                If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
       
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zldatabase.GetPara("存盘打印", glngSys, 1344)) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
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
    
    mblnSave = False
    mshBill.ClearBill
    mshStructure.ClearBill
    Call 显示合计金额
    txt摘要.Text = ""
    mshBill.SetFocus
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
End Sub

Private Sub Form_Load()
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")

    txtNo = mstr单据号
    txtNo.Tag = txtNo
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "协定药品入库", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    mlng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(3, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '根据系统参数决定药房人员查看单据时，是否显示成本价
    If mblnViewCost = False Then
        mshBill.ColWidth(mconIntCol采购价) = 0
        mshBill.ColWidth(mconIntCol采购金额) = 0
        mshBill.ColWidth(mconintCol差价) = 0
        mshStructure.ColWidth(mconIntCol构采购价) = 0
        mshStructure.ColWidth(mconIntCol构采购金额) = 0
        mshStructure.ColWidth(mconintCol构差价) = 0
    Else
        mshBill.ColWidth(mconIntCol采购价) = 900
        mshBill.ColWidth(mconIntCol采购金额) = 900
        mshBill.ColWidth(mconintCol差价) = 800
        mshStructure.ColWidth(mconIntCol构采购价) = 1200
        mshStructure.ColWidth(mconIntCol构采购金额) = 1200
        mshStructure.ColWidth(mconintCol构差价) = 1000
    End If
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
        mshStructure.ColWidth(mconIntCol构商品名) = IIf(mshStructure.ColWidth(mconIntCol构商品名) = 0, 2000, mshStructure.ColWidth(mconIntCol构商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
        mshStructure.ColWidth(mconIntCol构商品名) = 0
    End If
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim str包装系数 As String
    Dim intRow As Integer
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPricedigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    
    On Error GoTo errHandle
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
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
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
                
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 3 and a.no=[1] "
                Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                With cboStock
                    .AddItem rsInitCard!名称
                    .ItemData(.NewIndex) = rsInitCard!Id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量 AS 数量,'1' as 比例系数,"
                    str包装系数 = "1"
                Case mconint门诊单位
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 数量,B.门诊包装 as 比例系数, "
                    str包装系数 = "B.门诊包装"
                Case mconint住院单位
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 数量,B.住院包装 as 比例系数,"
                    str包装系数 = "B.住院包装"
                Case mconint药库单位
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 数量, b.药库包装 as 比例系数, "
                    str包装系数 = "B.药库包装"
            End Select
            
            gstrSQL = " SELECT * FROM " & _
                "    (SELECT DISTINCT 序号,A.药品ID, '[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名,F.规格," & _
                strUnitQuantity & _
                "    (A.成本价*" & str包装系数 & ") AS 成本价,A.成本金额 AS 成本金额," & _
                "    (A.零售价*" & str包装系数 & ") AS 零售价,A.零售金额 AS 零售金额," & _
                "    A.差价 AS 差价,A.填制人,A.填制日期,A.审核人,A.审核日期,A.摘要,B.最大效期,B.药品来源,B.基本药物," & _
                "    F.是否变价,B.加成率/100 AS 加成率 ,B.药房分批 AS 药房分批核算 " & _
                "    FROM 药品收发记录 A, 药品规格 B,收费项目别名 E,收费项目目录 F " & _
                "    WHERE A.药品ID = B.药品ID AND B.药品ID = F.ID " & _
                "    AND B.药品ID = E.收费细目ID(+) And E.性质(+)=3 " & _
                "    AND 记录状态 = [2] AND A.单据 = 3 AND 入出系数=1 " & _
                "    AND A.NO = [1])" & _
                " ORDER BY 序号 "
            Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt填制人 = rsInitCard!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = gstrUserName
            End If
            Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
            
            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            
            With mshBill
                Do While Not rsInitCard.EOF
                    
                    intRow = rsInitCard!序号
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard!药品ID
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = IIf(IsNull(rsInitCard!通用名), "", rsInitCard!通用名)
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
                    .TextMatrix(intRow, mconIntCol来源) = IIf(IsNull(rsInitCard!药品来源), "", rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = IIf(IsNull(rsInitCard!基本药物), "", rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    
                    .TextMatrix(intRow, mconIntCol数量) = zlStr.FormatEx(rsInitCard!数量, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(rsInitCard!成本价, intCostDigit)
                    .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(rsInitCard!成本金额, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!零售价, intPricedigit)
                    .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(rsInitCard!零售金额, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(rsInitCard!差价, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!加成率 & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    rsInitCard.MoveNext
                Loop
                Dim dblCostPrice As Double
                
                If .TextMatrix(1, 0) <> "" Then
                    If SetStructure(.TextMatrix(1, 0)) <> False Then
                        If .TextMatrix(1, mconIntCol数量) <> "" Then
                            GetStructureNum .TextMatrix(1, mconIntCol数量) * .TextMatrix(1, mconIntCol比例系数), dblCostPrice, False
                        End If
                    End If
                End If
            End With
            rsInitCard.Close
                 
    End Select
    Call 显示合计金额
    If mint编辑状态 = 2 And mint库存检查 <> 0 Then
        SetUseCountCol
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'设置修改前原料药的使用数量，以便于在修改过程中对库存数量的判断更准确
Private Sub SetUseCountCol()
    Dim rsUseCount As New Recordset
    Dim numUsedCount As Double
    Dim vardrug As Variant
    
'    gstrSQL = "select 药品id,填写数量,费用id from 药品收发记录 where no='" & mstr单据号 & "' and 单据=3 and 记录状态=1 and 入出系数=-1 "
'    Call SQLTest(App.Title, mstrCaption, gstrSQL)
'    rsUseCount.Open gstrSQL, gcnOracle
'    Call SQLTest
    On Error GoTo errHandle
    gstrSQL = "select 药品id,填写数量,费用id from 药品收发记录 where no=[1] and 单据=3 and 记录状态=1 and 入出系数=-1 "
    Set rsUseCount = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
    
    If rsUseCount.EOF Then Exit Sub
    Set mcolUseCount = New Collection
    With mcolUseCount
        Do While Not rsUseCount.EOF
            numUsedCount = 0
            For Each vardrug In mcolUseCount
                If vardrug(0) = rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)) Then
                    numUsedCount = vardrug(1)
                    .Remove vardrug(0)
                    Exit For
                End If
            Next
            
            .Add Array(rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)), numUsedCount + rsUseCount.Fields(1)), rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0))
            rsUseCount.MoveNext
        Loop
        rsUseCount.Close
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol采购价) = "购价"
        .TextMatrix(0, mconIntCol采购金额) = "购价金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconIntCol原销期) = "原销期"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        
        .TextMatrix(1, 0) = ""
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol原销期) = 0
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol数量) = 1000
        If mblnViewCost = False Then
            .ColWidth(mconIntCol采购价) = 0
            .ColWidth(mconIntCol采购金额) = 0
            .ColWidth(mconintCol差价) = 0
        Else
            .ColWidth(mconIntCol采购价) = 900
            .ColWidth(mconIntCol采购金额) = 900
            .ColWidth(mconintCol差价) = 800
        End If
        .ColWidth(mconIntCol售价) = 900
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol药名) = 1
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol原销期) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            .ColData(mconIntCol数量) = 4
            
            If cboStock.Enabled = True Then
                cboStock.Enabled = True
            End If
            txt摘要.Enabled = True
        Else
            .ColData(mconIntCol数量) = 5
            .ColData(mconIntCol药名) = 0
            cboStock.Enabled = False
            txt摘要.Enabled = False
        End If
            
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        
        
        .ColData(mconIntCol比例系数) = 5
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
    End With
    
    With mshStructure
        
        .Cols = mconInt构ColS
        
        .TextMatrix(0, mconIntCol构药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol构商品名) = "商品名"
        .TextMatrix(0, mconIntCol构规格) = "规格"
        .TextMatrix(0, mconIntCol构产地) = "产地"
        .TextMatrix(0, mconIntCol构单位) = "单位"
        .TextMatrix(0, mconIntCol构数量) = "数量"
        .TextMatrix(0, mconIntCol构组成数量) = "组成数量"
        .TextMatrix(0, mconIntCol构可用数量) = "可用数量"
        .TextMatrix(0, mconIntcol加成率) = "加成率"
        .TextMatrix(0, mconintcol构实际差价) = "实际差价"
        .TextMatrix(0, mconintcol构实际金额) = "实际金额"
        .TextMatrix(0, mconintcol构药品id) = "药品id"
        
        .TextMatrix(0, mconIntCol构采购价) = "成本价"
        .TextMatrix(0, mconIntCol构采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol构售价) = "售价"
        .TextMatrix(0, mconIntCol构售价金额) = "售价金额"
        .TextMatrix(0, mconintCol构差价) = "差价"
        .TextMatrix(0, mconIntCol构药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol构药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol构药品名称) = "药品名称"
        
        
        .ColWidth(mconIntCol构药名) = 2500
        .ColWidth(mconIntCol构商品名) = 2000
        .ColWidth(mconIntCol构规格) = 1000
        .ColWidth(mconIntCol构产地) = 1000
        .ColWidth(mconIntCol构单位) = 500
        .ColWidth(mconIntCol构数量) = 1000
        .ColWidth(mconIntCol构组成数量) = 0
        .ColWidth(mconIntCol构可用数量) = 0
        .ColWidth(mconIntcol加成率) = 0
        .ColWidth(mconintcol构实际差价) = 0
        .ColWidth(mconintcol构实际金额) = 0
        .ColWidth(mconintcol构药品id) = 0
        
        If mblnViewCost = False Then
            .ColWidth(mconIntCol构采购价) = 0
            .ColWidth(mconIntCol构采购金额) = 0
            .ColWidth(mconintCol构差价) = 0
        Else
            .ColWidth(mconIntCol构采购价) = 1000
            .ColWidth(mconIntCol构采购金额) = 1200
            .ColWidth(mconintCol构差价) = 1000
        End If
        .ColWidth(mconIntCol构售价) = 1000
        .ColWidth(mconIntCol构售价金额) = 1200
                
        .ColWidth(mconIntCol构药品编码和名称) = 0
        .ColWidth(mconIntCol构药品编码) = 0
        .ColWidth(mconIntCol构药品名称) = 0
        
        .ColAlignment(mconIntCol构商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol构单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol构采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol构采购金额) = flexAlignRightCenter
        .ColAlignment(mconintCol构差价) = flexAlignRightCenter
        .ColAlignment(mconIntCol构售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol构售价金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol构数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol构规格) = flexAlignLeftCenter
    End With
    txt摘要.MaxLength = Sys.FieldsLength("药品收发记录", "摘要")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
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
    End With
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
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
    
    With mshStructure
        .Left = mshBill.Left
        .Width = mshBill.Width
        .Top = txt摘要.Top - 60 - .Height
    End With
    
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = mshStructure.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 50
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
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\协定药品入库", "药品名称显示方式", mintDrugNameShow)
    
    If mshDrug.Visible Then
        mshDrug.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
End Sub

Private Function CheckBuildupNumStore() As String
    '检查协定药品的原料药库存数量是否足够
    '返回值：空-表示数量足够，不为空-表示数量不够
    Dim intRow As Integer
    Dim dblNum组合 As Double
    Dim dblNum As Double
    Dim rstemp As ADODB.Recordset
    Dim strKey As String
    Dim collNum As Collection
    Dim vardrug As Variant
    Dim strArray As String
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim lng药品id As Long
    
    With mshBill
        If .rows <= 1 Then Exit Function
        
        Set collNum = New Collection
        
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, 0)) <> 0 Then
                gstrSQL = "Select Distinct b.药品id As 原料药id, (a.分子 / a.分母) As 组成, b.剂量系数 As 原料药剂量系数, c.实际数量 As 原料药库存" & vbNewLine & _
                    "From 协定药品对照 A, 药品规格 B, 药品库存 C" & vbNewLine & _
                    "Where a.协定药品id = b.药品id And b.药品id = c.药品id(+) And a.药品id = [1] And c.库房id = [2]"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "查询组成系数", Val(.TextMatrix(intRow, 0)), cboStock.ItemData(cboStock.ListIndex))
                If rstemp.RecordCount > 0 Then
                    If rstemp!原料药剂量系数 <> 0 Then
                        dblNum组合 = rstemp!组成 * Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                    End If
                    
                    For Each vardrug In collNum
                        If vardrug(0) = rstemp!原料药id & "" Then
                            dblNum = vardrug(1)
                            collNum.Remove vardrug(0)
                            Exit For
                        End If
                    Next
                    strKey = rstemp!原料药id
                    '以最小单位保存数量，方便审核时数量与库存数据比较
                    strArray = dblNum + dblNum组合
                    collNum.Add Array(strKey, strArray), strKey
                End If
            End If
        Next
        
        For Each varNum In collNum
            lng药品id = varNum(0)  '格式是药品id,批次
            dblNum = varNum(1)
            
            '只有有数量才判断
            If dblNum > 0 Then
                gstrSQL = "Select (a.实际数量 - [1]) As 剩余数量, b.名称" & vbNewLine & _
                            "From 药品库存 A, 收费项目目录 B" & vbNewLine & _
                            "Where a.药品id = b.Id And a.药品id = [2] And a.库房id = [3] And Nvl(a.批次, 0) = [4] And b.类别 In ('5', '6', '7') And a.性质 = 1"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "库存检查", dblNum, lng药品id, cboStock.ItemData(cboStock.ListIndex), 0)
                If rstemp.RecordCount = 0 Then
                    gstrSQL = "select 名称 from 收费项目目录 where id=[1]"
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "库存检查", lng药品id)
                    CheckBuildupNumStore = rstemp!名称
                    Exit Function
                Else
                    If rstemp!剩余数量 >= 0 Then
                        CheckBuildupNumStore = ""
                    Else
                        CheckBuildupNumStore = rstemp!名称
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Function SaveCheck() As Boolean
    Dim str药品 As String
    Dim mbln提示方式  As Boolean
    '检查库存
    str药品 = CheckBuildupNumStore
    If str药品 <> "" Then
        If mint库存检查 = 1 Then '不足提醒
            If MsgBox("原料药品【" & str药品 & "】库存不足，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                mbln提示方式 = True
            End If
        ElseIf mint库存检查 = 2 Then '不足禁止
            MsgBox "原料药品【" & str药品 & "】库存不足，不能审核！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    mblnSave = False
    SaveCheck = False
    gstrSQL = "zl_协定入库_Verify('" & txtNo.Tag & "','" & gstrUserName & "')"
    On Error GoTo errHandle
    Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
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
    
    With mshStructure
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol构药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol构药名) = .TextMatrix(lngRow, mconIntCol构药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol构药名) = .TextMatrix(lngRow, mconIntCol构药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol构药名) = .TextMatrix(lngRow, mconIntCol构药品编码和名称)
                End If
            End If
        Next
    End With
End Sub
Private Sub mshBill_AfterDeleteRow()
    With mshBill
        If .Row > 1 Then
            .Row = .Row - 1
        Else
            .Row = 1
        End If
        If .TextMatrix(.Row, 0) = "" Then
            mshStructure.ClearBill
        Else
            Dim dblCostPrice As Double
            
            If SetStructure(.TextMatrix(.Row, 0)) Then
                If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                    GetStructureNum .TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol比例系数), dblCostPrice, False
                End If
            End If
            
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
            If MsgBox("你确实要删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As New Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intStockID As Long
    Dim strUnitQuantity As String
    
    On Error GoTo errHandle
    mblnChange = True
    
    Select Case mintUnit
        Case mconint售价单位
            strUnitQuantity = "D.计算单位 AS 单位, trim(to_char(s.库存数量," & mstrNumberFormat & ")) AS 数量,'1' as 比例系数," _
                & "trim(to_char(p.现价," & mstrPriceFormat & ")) as 售价,"
        Case mconint门诊单位
            strUnitQuantity = "d.门诊单位 AS 单位, trim(to_char(s.库存数量 / d.门诊包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.门诊包装) as 比例系数," _
                & "trim(to_char(p.现价*d.门诊包装," & mstrPriceFormat & ")) as 售价, "
        Case mconint住院单位
            strUnitQuantity = "d.住院单位 AS 单位, trim(to_char(s.库存数量 / d.住院包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.住院包装) as 比例系数," _
                & "trim(to_char(p.现价*d.住院包装," & mstrPriceFormat & ")) as 售价,"
        Case mconint药库单位
            strUnitQuantity = "d.药库单位 AS 单位, trim(to_char(s.库存数量 / d.药库包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.药库包装) as 比例系数," _
                & "trim(to_char(p.现价*d.药库包装," & mstrPriceFormat & ")) as 售价 , "
    End Select

    intStockID = cboStock.ItemData(cboStock.ListIndex)

    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50

    gstrSQL = "" & _
        " SELECT DECODE(D.类别,5,'西成药',6,'中成药','中草药') AS 材质分类,D.剂型,D.编码,D.通用名称" & _
        "   ,D.商品名,D.药品来源,D.基本药物,D.规格,D.产地,D.药品ID," & _
             strUnitQuantity & _
        "    S.库存金额, D.最大效期,D.是否变价,D.加成率,D.药房分批核算,E.库房货位 " & _
        " FROM  " & _
        "    (SELECT DISTINCT J.名称 剂型,M.类别,M.编码,M.名称 通用名称,A.名称 商品名,D.药品来源,D.基本药物," & _
        "        M.规格,M.产地, D.药名ID, D.药品ID, M.计算单位,NVL (TO_CHAR (D.最大效期, '9999990'), 0) 最大效期,D.门诊单位," & _
        "        TO_CHAR (D.门诊包装, '999999999990.99999') 门诊包装,D.住院单位,TO_CHAR (D.住院包装, '999999999990.99999') 住院包装," & _
        "        D.药库单位,TO_CHAR(D.药库包装, '999999999990.99999') 药库包装,M.是否变价,D.加成率,D.药房分批 AS 药房分批核算 " & _
        "    FROM 协定药品对照 F, 药品特性 C, 药品规格 D,收费项目目录 M,收费项目别名 A, 药品剂型 J " & _
        "    WHERE F.药品ID = D.药品ID AND D.药品ID=M.ID AND D.药名ID=C.药名ID AND C.药品剂型 = J.名称(+)" & _
        "    AND D.药品ID = A.收费细目ID(+) AND A.性质(+)=3 AND NVL(D.协定药品,0)=1 And (M.站点 = '" & gstrNodeNo & "' Or M.站点 is Null) " & _
        "    AND (EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 = '制剂室' AND 部门ID =[1] ) " & _
        "        OR M.类别 =(SELECT DISTINCT 5 FROM 部门性质说明 WHERE 工作性质 LIKE '西药%' AND 部门ID =[1]) " & _
        "        OR M.类别 =(SELECT DISTINCT 6 FROM 部门性质说明 WHERE 工作性质 LIKE '成药%' AND 部门ID =[1]) "
    gstrSQL = gstrSQL & _
        "        OR M.类别 =(SELECT DISTINCT 7 FROM 部门性质说明 WHERE 工作性质 LIKE '中药%' AND 部门ID =[1])) " & _
        "    AND ( EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 LIKE '%药库' AND 部门ID = [1]) " & _
        "        OR EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 = '制剂室' AND 部门ID =[1]) " & _
        "        OR DECODE (服务对象,1,1,3,1,0) =(SELECT DISTINCT '1' FROM 部门性质说明 WHERE 工作性质 LIKE '%药房' AND 部门ID =[1] AND 服务对象 IN (1, 3)) " & _
        "        OR DECODE (服务对象,2,1,3,1,0) =(SELECT DISTINCT '1' FROM 部门性质说明 WHERE 工作性质 LIKE '%药房' AND 部门ID =[1] AND 服务对象 IN (2, 3))) " & _
        "    AND ( M.撤档时间 IS NULL OR TO_CHAR (M.撤档时间, 'YYYY-MM-DD') = '3000-01-01') ) D,收费价目 P," & _
        "    (SELECT 药品ID,TRIM(TO_CHAR(SUM(可用数量)," & mstrNumberFormat & ")) 可用数量," & _
        "        TRIM(TO_CHAR(SUM (实际数量), " & mstrNumberFormat & ")) 库存数量," & _
        "        TRIM(TO_CHAR(SUM (实际金额), " & mstrMoneyFormat & ")) 库存金额 " & _
        "    FROM 药品库存 " & _
        "    WHERE 库房ID =[1] AND 性质=1 " & _
        "    GROUP BY 药品ID) S,药品储备限额 E,(Select 收费细目id From 收费执行科室 Where 执行科室id = [1]) F " & _
        " WHERE D.药品ID=P.收费细目ID AND SYSDATE BETWEEN P.执行日期 AND NVL(P.终止日期,SYSDATE)" & _
        GetPriceClassString("P") & _
        " AND D.药品ID=S.药品ID(+) AND D.药品ID=E.药品ID(+) And D.药品id = F.收费细目id AND E.库房ID(+)=[1] " & _
        " ORDER BY D.编码"
    Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID)

    If RecReturn.EOF Then Exit Sub
    Set mshDrug.Recordset = RecReturn
    RecReturn.Close
    Call SetDrugWidth(sngLeft, sngTop)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'设置药品选择器的宽度及相关属性
Private Sub SetDrugWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshDrug
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
        If RestoreFlexState(mshDrug, MStrCaption) = False Then
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
    With mshBill
        mshBill.Text = UCase(curText)
        mshBill.SelStart = Len(mshBill.Text)
    End With
    mblnChange = True
End Sub


Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            Dim dblCostPrice As Double
            
            If .TextMatrix(.Row, 0) <> "" Then
                If SetStructure(.TextMatrix(.Row, 0)) <> False Then
                    If IIf(.TextMatrix(.Row, mconIntCol数量) = "", 0, .TextMatrix(.Row, mconIntCol数量)) <> 0 Then
                        GetStructureNum .TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol比例系数), dblCostPrice, False
                    End If
                End If
            Else
                mshStructure.ClearBill
            End If
                
        End If
        
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mconIntCol采购价
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol采购金额
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        With mshBill
            strKey = UCase(Trim(.Text))
            
            If Mid(strKey, 1, 1) = "[" Then
                If InStr(2, strKey, "]") <> 0 Then
                    strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                Else
                    strKey = Mid(strKey, 2)
                End If
            End If
            Select Case .Col
                
                Case mconIntCol药名
                    If strKey <> "" Then
                        Dim RecReturn As New Recordset
                        Dim sngLeft As Single
                        Dim sngTop As Single
                        Dim intStockID As Long
                        
                        Select Case mintUnit
                            Case mconint售价单位
                                strUnitQuantity = "d.计算单位 AS 单位, TRIM(to_char(s.库存数量," & mstrNumberFormat & ")) AS 数量,'1' as 比例系数," _
                                    & "TRIM(to_char(p.现价," & mstrPriceFormat & ")) as 售价,"
                            Case mconint门诊单位
                                strUnitQuantity = "d.门诊单位 AS 单位, TRIM(to_char(s.库存数量 / d.门诊包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.门诊包装) as 比例系数," _
                                    & "TRIM(to_char(p.现价*d.门诊包装," & mstrPriceFormat & ")) as 售价, "
                            Case mconint住院单位
                                strUnitQuantity = "d.住院单位 AS 单位, TRIM(to_char(s.库存数量 / d.住院包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.住院包装) as 比例系数," _
                                    & "TRIM(to_char(p.现价*d.住院包装," & mstrPriceFormat & ")) as 售价,"
                            Case mconint药库单位
                                strUnitQuantity = "d.药库单位 AS 单位, TRIM(to_char(s.库存数量 / d.药库包装," & mstrNumberFormat & ")) AS 数量,TRIM(d.药库包装) as 比例系数," _
                                    & "TRIM(to_char(p.现价*d.药库包装," & mstrPriceFormat & ")) as 售价 , "
                        End Select
                        
                        intStockID = cboStock.ItemData(cboStock.ListIndex)
                        
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50

                        gstrSQL = "" & _
                        " SELECT DECODE(D.类别,5,'西成药',6,'中成药','中草药') AS 材质分类,D.剂型,D.编码,D.通用名称,D.商品名," & _
                        "      D.药品来源,D.基本药物,D.规格,D.产地,D.药品ID," & _
                               strUnitQuantity & _
                        "      S.库存金额, D.最大效期,D.是否变价,D.加成率,D.药房分批核算,E.库房货位  " & _
                        " FROM  " & _
                        "     (SELECT DISTINCT J.名称 剂型,M.类别,M.编码,M.名称 通用名称,A.名称 商品名,d.药品来源,d.基本药物," & _
                        "         M.规格,M.产地, D.药名ID, D.药品ID, M.计算单位,NVL (TO_CHAR (D.最大效期, '9999990'), 0) 最大效期,D.门诊单位," & _
                        "         TO_CHAR (D.门诊包装, '999999999990.99999') 门诊包装,D.住院单位,TO_CHAR (D.住院包装, '999999999990.99999') 住院包装," & _
                        "         D.药库单位,TO_CHAR(D.药库包装, '999999999990.99999') 药库包装,M.是否变价,D.加成率,D.药房分批 AS 药房分批核算 " & _
                        "     FROM 协定药品对照 F, 药品特性 C, 药品规格 D, 药品剂型 J,收费项目目录 M," & _
                        "         (Select A.* From 收费项目别名 A,收费项目目录 B" & _
                        "     Where A.收费细目ID=B.ID ANd (A.简码 Like [2] Or A.名称 Like [2] Or B.编码 Like [2]) And A.码类=" & IIf(gint简码方式 = 1, 2, 1) & _
                        "         And (B.站点 = '" & gstrNodeNo & "' Or B.站点 is Null)) A,收费项目别名 N " & _
                        "     WHERE F.药品ID = D.药品ID AND D.药品ID=M.ID And D.药品ID=A.收费细目ID AND D.药名ID=C.药名ID AND C.药品剂型 = J.名称(+)" & _
                        "     AND D.药品ID = N.收费细目ID(+) AND N.性质(+)=3 AND NVL(D.协定药品,0)=1 " & _
                        "     AND (EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 = '制剂室' AND 部门ID = [1])"
                        gstrSQL = gstrSQL & _
                        "         OR M.类别 =(SELECT DISTINCT 5 FROM 部门性质说明 WHERE 工作性质 LIKE '西药%' AND 部门ID = [1] ) " & _
                        "         OR M.类别 =(SELECT DISTINCT 6 FROM 部门性质说明 WHERE 工作性质 LIKE '成药%' AND 部门ID = [1] ) " & _
                        "         OR M.类别 =(SELECT DISTINCT 7 FROM 部门性质说明 WHERE 工作性质 LIKE '中药%' AND 部门ID = [1] )) " & _
                        "     AND ( EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 LIKE '%药库' AND 部门ID =  [1] ) " & _
                        "         OR EXISTS (SELECT 1 FROM 部门性质说明 WHERE 工作性质 = '制剂室' AND 部门ID = [1] ) " & _
                        "         OR DECODE (服务对象,1,1,3,1,0) =(SELECT DISTINCT '1' FROM 部门性质说明 WHERE 工作性质 LIKE '%药房' AND 部门ID = [1]  AND 服务对象 IN (1, 3)) " & _
                        "         OR DECODE (服务对象,2,1,3,1,0) =(SELECT DISTINCT '1' FROM 部门性质说明 WHERE 工作性质 LIKE '%药房' AND 部门ID = [1]  AND 服务对象 IN (2, 3))) " & _
                        "     AND ( M.撤档时间 IS NULL OR TO_CHAR (M.撤档时间, 'YYYY-MM-DD') = '3000-01-01') ) D,收费价目 P," & _
                        "     (SELECT 药品ID,TO_CHAR(SUM(可用数量), " & mstrNumberFormat & ") 可用数量," & _
                        "         TO_CHAR (SUM (实际数量), " & mstrNumberFormat & ") 库存数量," & _
                        "         TO_CHAR (SUM (实际金额), " & mstrMoneyFormat & ") 库存金额 " & _
                        "     FROM 药品库存 " & _
                        "     WHERE 库房ID = [1]  AND 性质=1 " & _
                        "     GROUP BY 药品ID) S,药品储备限额 E,(Select 收费细目id From 收费执行科室 Where 执行科室id = [1]) F " & _
                        " WHERE D.药品ID=P.收费细目ID AND SYSDATE BETWEEN P.执行日期 AND NVL(P.终止日期,SYSDATE)" & _
                        GetPriceClassString("P") & _
                        " AND D.药品ID=S.药品ID(+) AND D.药品ID=E.药品ID(+) And D.药品id = F.收费细目id AND E.库房ID(+)= [1]"
                        
                        Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                        
                        If RecReturn.EOF Then
                            MsgBox "没有匹配的协定药品！", vbInformation + vbOKOnly, gstrSysName
                            RecReturn.Close
                            Cancel = True
                            Exit Sub
                        ElseIf RecReturn.RecordCount = 1 Then
                            If SetColValue(.Row, RecReturn!药品ID, "[" & RecReturn!编码 & "]", RecReturn!通用名称, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                               "" & RecReturn!药品来源, "" & RecReturn!基本药物, IIf(IsNull(RecReturn!规格), "", RecReturn!规格), _
                               RecReturn!单位, IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                               RecReturn!比例系数, RecReturn!是否变价, RecReturn!加成率, RecReturn!药房分批核算) = False Then
                               RecReturn.Close
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            RecReturn.Close
                        Else
                            Set mshDrug.Recordset = RecReturn
                            RecReturn.Close
                            Call SetDrugWidth(sngLeft, sngTop)
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    Call 提示库存数
                    'End If
                Case mconIntCol数量
                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                        MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        If Val(strKey) = 0 Then
                            MsgBox "对不起，数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) < 0.001 Then
                            MsgBox "对不起，数量必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        Dim dblCostPrice As Double
                        
                        
                        If .TextMatrix(.Row, 0) = "" Then Exit Sub
                        
                        If GetStructureNum(strKey * .TextMatrix(.Row, mconIntCol比例系数), dblCostPrice) = False Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, mconIntCol采购价) = zlStr.FormatEx(dblCostPrice * .TextMatrix(.Row, mconIntCol比例系数), mintCostDigit)
                        End If
                                
                        strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                        .Text = strKey
                        If .TextMatrix(.Row, mconIntCol采购价) <> "" Then
                            .TextMatrix(.Row, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol采购价) * strKey, mintMoneyDigit, , True)
                            If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                                .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol采购价) / (1 - Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1)), mintPriceDigit)
                            End If
                        End If
                        
                        If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                            .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strKey, mintMoneyDigit, , True)
                        End If
                        .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol采购金额) = "", 0, .TextMatrix(.Row, mconIntCol采购金额)), mintMoneyDigit, , True)
                        
                    End If
                    显示合计金额
                
            End Select
        End With
    ElseIf KeyCode = vbKeyDown And Shift = vbAltMask Then
        mshbill_CommandClick
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'根据药品ID判断是否有够用的组成药品，如有，则填上相应的数量
Private Function GetStructureNum(ByVal dblNum As Double, ByRef dblCostPrice As Double, _
         Optional bln判断库存 As Boolean = True) As Boolean
    Dim rsDrug As New Recordset
    Dim intReturn As Integer
    Dim blnContinue As Boolean      '用户的选择：0，退出，1继续
    Dim dblConstruct As Double      '实际数量对应的组成数量
    Dim dblPurchase As Double       '协定药品的成本价：所有（组成药品的进价*组成数量）
    Dim intRow As Integer
    Dim dbl原填写数量 As Double
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim numUseCount As Double
    Dim vardrug As Variant
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           差价和成本价在出库处理中的公式
    '   出库金额=数量*售价
    '   出库差价=出库金额*（实际差价/实际金额）
    '          如果实际差价和实际金额不存在时，为：
    '       出库差价=出库金额*指导差价率
    '   购价（成本价)=(出库金额-出库差价)/数量
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    intCostDigit = mintCostDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    
    GetStructureNum = False
    blnContinue = False
    With mshStructure
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dblConstruct = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol构组成数量) * dblNum, intNumberDigit, , True)
                
                .TextMatrix(intRow, mconIntCol构数量) = zlStr.FormatEx(dblConstruct, intNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol构售价金额) = zlStr.FormatEx(dblConstruct * .TextMatrix(intRow, mconIntCol构售价), intNumberDigit, , True)
                If .TextMatrix(intRow, mconintcol构实际金额) <= "0" Then
'                    .TextMatrix(intRow, mconintCol构差价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol构售价金额) * Split(.TextMatrix(intRow, mconIntcol加成率), "||")(0) / 100, intMoneyDigit)
'                    .TextMatrix(intRow, mconIntCol构采购价) =Str.FormatEx((.TextMatrix(intRow, mconIntCol构售价金额) - .TextMatrix(intRow, mconintCol构差价)) / (IIf(dblConstruct = 0, 1, dblConstruct)), intCostDigit)
                    .TextMatrix(intRow, mconIntCol构采购金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol构采购价) * dblConstruct, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol构差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol构售价金额)) - Val(.TextMatrix(intRow, mconIntCol构采购金额)), intMoneyDigit, , True)
                Else
'                    .TextMatrix(intRow, mconintCol构差价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol构售价金额) * (.TextMatrix(intRow, mconintcol构实际差价) / .TextMatrix(intRow, mconintcol构实际金额)), intMoneyDigit)
'                    .TextMatrix(intRow, mconIntCol构采购价) =Str.FormatEx((.TextMatrix(intRow, mconIntCol构售价金额) - .TextMatrix(intRow, mconintCol构差价)) / (IIf(dblConstruct = 0, 1, dblConstruct)), intCostDigit)
                    .TextMatrix(intRow, mconIntCol构采购金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol构采购价) * dblConstruct, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol构差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol构售价金额)) - Val(.TextMatrix(intRow, mconIntCol构采购金额)), intMoneyDigit, , True)
                End If
                dblPurchase = zlStr.FormatEx(dblPurchase + .TextMatrix(intRow, mconIntCol构采购金额) / dblNum, intCostDigit)
                
            End If
        Next
        dblCostPrice = dblPurchase
            
    End With
    
    GetStructureNum = True
End Function


'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal int药品id As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, _
    ByVal str基本药物 As String, ByVal str规格 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal int原效期 As Integer, ByVal num比例系数 As Double, _
    ByVal int是否变价 As Integer, ByVal dbl加成率 As Double, ByVal int药房分批核算 As Integer) As Boolean
    
    Dim intCount As Integer
    Dim rsStructure As New Recordset
    Dim intCol As Integer
    Dim str药名 As String
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
        
        If Not SetStructure(int药品id) Then Exit Function
        .TextMatrix(intRow, 0) = int药品id
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = str药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = str商品名
        .TextMatrix(intRow, mconIntCol来源) = str药品来源
        .TextMatrix(intRow, mconIntCol基本药物) = str基本药物
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(num售价, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl加成率 / 100 & "||" & int是否变价 & "||" & int药房分批核算
            
    End With
    SetColValue = True
End Function

Private Function SetStructure(ByVal int药品id As Long) As Boolean
    Dim rsStructure As New Recordset
    Dim str药名 As String
    Dim rs成本价 As ADODB.Recordset
    
    SetStructure = False
    mshStructure.Redraw = False
    
    On Error GoTo errHandle
    If mint编辑状态 <> 4 Then
        gstrSQL = "SELECT DISTINCT B.药品ID,'[' || F.编码 || ']' As 编码,F.名称 As 通用名称,E.名称 AS 商品名称, F.规格, C.上次产地,F.计算单位 AS 单位,c.平均成本价," & _
                  " C.实际差价,C.实际金额, to_char(D.现价, " & mstrPriceFormat & ") 售价, " & _
                  " (A.分子 / A.分母) AS 组成,C.可用数量,B.加成率,F.是否变价,B.药房分批 AS 药房分批核算, Nvl(F.是否变价, 0) 定价 " & _
                  " FROM 协定药品对照 A,药品规格 B,收费项目别名 E,收费项目目录 F,药品库存 C, 收费价目 D" & _
                  " WHERE A.协定药品ID = B.药品ID AND B.药品ID=F.ID " & _
                  " AND A.协定药品ID = D.收费细目ID AND (SYSDATE BETWEEN 执行日期 AND NVL(终止日期,SYSDATE))" & _
                  GetPriceClassString("D") & _
                  " AND B.药品ID = E.收费细目ID(+) AND E.性质(+)=3 " & _
                  " AND A.协定药品ID = C.药品ID(+) AND C.库房ID(+) =[1] AND C.性质(+)=1" & _
                  " AND (F.站点 = [3] Or F.站点 is Null) And A.药品ID =[2] "

        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), int药品id, gstrNodeNo)
        If rsStructure.EOF Then
            mshStructure.Redraw = True
            Exit Function
        End If
        With mshStructure
            .ClearBill
            Do While Not rsStructure.EOF
                If rsStructure!药房分批核算 = 1 Then
                    MsgBox "组成药品是一个药房分批药品，但当前版本不支持药房分批的组成药品，请检查！", vbInformation + vbOKOnly, gstrSysName
                    mshStructure.Redraw = True
                    Exit Function
                End If
                If rsStructure!定价 = 1 And Nvl(rsStructure!可用数量, 0) = 0 Then
                    MsgBox "该协定药品的组成药品是时价药品，可用库存为0。当前版本不支持，请检查！", vbInformation + vbOKOnly, gstrSysName
                    mshStructure.Redraw = True
                    Exit Function
                End If
                
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsStructure!通用名称
                Else
                    str药名 = IIf(IsNull(rsStructure!商品名称), rsStructure!通用名称, rsStructure!商品名称)
                End If
                                                
                .TextMatrix(.Row, mconIntCol构药品编码和名称) = rsStructure!编码 & str药名
                .TextMatrix(.Row, mconIntCol构药品编码) = rsStructure!编码
                .TextMatrix(.Row, mconIntCol构药品名称) = str药名
                
                If mintDrugNameShow = 0 Then
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品编码和名称)
                ElseIf mintDrugNameShow = 1 Then
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品编码)
                Else
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品名称)
                End If
                
                .TextMatrix(.Row, mconIntCol构商品名) = IIf(IsNull(rsStructure!商品名称), "", rsStructure!商品名称)
                
                .TextMatrix(.Row, mconIntCol构规格) = IIf(IsNull(rsStructure!规格), "", rsStructure!规格)
                .TextMatrix(.Row, mconIntCol构产地) = IIf(IsNull(rsStructure!上次产地), "", rsStructure!上次产地)
                .TextMatrix(.Row, mconIntCol构单位) = rsStructure!单位
                .TextMatrix(.Row, mconIntCol构售价) = zlStr.FormatEx(rsStructure!售价, mintPriceDigit)
                .TextMatrix(.Row, mconIntCol构可用数量) = zlStr.FormatEx(IIf(IsNull(rsStructure!可用数量), "0", rsStructure!可用数量), mintNumberDigit, , True)
                .TextMatrix(.Row, mconIntCol构组成数量) = rsStructure!组成
                .TextMatrix(.Row, mconIntcol加成率) = rsStructure!加成率 / 100 & "||" & IIf(IsNull(rsStructure!是否变价), 0, rsStructure!是否变价) & "||" & IIf(IsNull(rsStructure!药房分批核算), 0, rsStructure!药房分批核算)
                .TextMatrix(.Row, mconintcol构实际差价) = IIf(IsNull(rsStructure!实际差价), "0", rsStructure!实际差价)
                .TextMatrix(.Row, mconintcol构实际金额) = IIf(IsNull(rsStructure!实际金额), "0", rsStructure!实际金额)
                .TextMatrix(.Row, mconintcol构药品id) = rsStructure!药品ID
                
                If IsNull(rsStructure!平均成本价) Then
                    gstrSQL = "select 成本价 from 药品规格 where 药品id=[1]"
                    Set rs成本价 = zldatabase.OpenSQLRecord(gstrSQL, "查询成本价", Val(rsStructure!药品ID))
                    If rs成本价.RecordCount > 0 Then
                        .TextMatrix(.Row, mconIntCol构采购价) = zlStr.FormatEx(rs成本价!成本价, mintCostDigit, , True)
                    End If
                Else
                    .TextMatrix(.Row, mconIntCol构采购价) = zlStr.FormatEx(rsStructure!平均成本价, mintCostDigit, , True)
                End If
                
                If .Row = .rows - 1 Then
                    .rows = .rows + 1
                End If
                .Row = .Row + 1
                rsStructure.MoveNext
            Loop
        End With
        rsStructure.Close
    Else            '查看
        gstrSQL = " SELECT DISTINCT A.药品ID,'[' || F.编码 || ']' As 编码,F.名称 As 通用名称,E.名称 AS 商品名称,F.规格," & _
                  "     A.产地,F.计算单位 AS 单位,A.实际数量,A.成本价,A.成本金额,A.零售价,A.零售金额,A.差价 " & _
                  " FROM " & _
                  "     (SELECT 药品ID,产地,实际数量,成本价,成本金额,零售价,零售金额,差价 FROM 药品收发记录 " & _
                  "     WHERE NO=[1] AND 单据=3 AND 记录状态=[3] " & _
                  "     AND 入出系数=-1 AND 扣率=[4] AND 费用ID =[2]) A," & _
                  "     药品规格 B,收费项目别名 E,收费项目目录 F " & _
                  " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                  " AND B.药品ID = E.收费细目ID(+) AND E.性质(+)=3 "
        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNo.Tag, int药品id, mint记录状态, mshBill.Row)
        
        If rsStructure.EOF Then
            mshStructure.Redraw = True
            Exit Function
        End If
        With mshStructure
            .ClearBill
            Do While Not rsStructure.EOF
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = rsStructure!通用名称
                Else
                    str药名 = IIf(IsNull(rsStructure!商品名称), rsStructure!通用名称, rsStructure!商品名称)
                End If
                                                
                .TextMatrix(.Row, mconIntCol构药品编码和名称) = rsStructure!编码 & str药名
                .TextMatrix(.Row, mconIntCol构药品编码) = rsStructure!编码
                .TextMatrix(.Row, mconIntCol构药品名称) = str药名
                
                If mintDrugNameShow = 0 Then
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品编码和名称)
                ElseIf mintDrugNameShow = 1 Then
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品编码)
                Else
                    .TextMatrix(.Row, mconIntCol构药名) = .TextMatrix(.Row, mconIntCol构药品名称)
                End If
                
                .TextMatrix(.Row, mconIntCol构商品名) = IIf(IsNull(rsStructure!商品名称), "", rsStructure!商品名称)
                
                .TextMatrix(.Row, mconIntCol构规格) = IIf(IsNull(rsStructure!规格), "", rsStructure!规格)
                .TextMatrix(.Row, mconIntCol构产地) = IIf(IsNull(rsStructure!产地), "", rsStructure!产地)
                .TextMatrix(.Row, mconIntCol构单位) = rsStructure!单位
                .TextMatrix(.Row, mconIntCol构数量) = zlStr.FormatEx(rsStructure!实际数量, mintNumberDigit, , True)
                .TextMatrix(.Row, mconIntCol构采购价) = zlStr.FormatEx(rsStructure!成本价, mintCostDigit)
                .TextMatrix(.Row, mconIntCol构采购金额) = zlStr.FormatEx(rsStructure!成本金额, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconIntCol构售价) = zlStr.FormatEx(rsStructure!零售价, mintPriceDigit)
                .TextMatrix(.Row, mconIntCol构售价金额) = zlStr.FormatEx(rsStructure!零售金额, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintCol构差价) = zlStr.FormatEx(rsStructure!差价, mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintcol构药品id) = rsStructure!药品ID
                
                If .Row = .rows - 1 Then
                    .rows = .rows + 1
                End If
                .Row = .Row + 1
                rsStructure.MoveNext
            Loop
                
        End With
        rsStructure.Close
        mshStructure.Redraw = True
        Exit Function
    End If
    
    SetStructure = True
    mshStructure.Redraw = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
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
            If Not SetColValue(mshBill.Row, .TextMatrix(.Row, 9), "[" & .TextMatrix(.Row, 2) & "]", .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), _
                .TextMatrix(.Row, 5), .TextMatrix(.Row, 6), .TextMatrix(.Row, 7), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), _
                IIf(IsNull(.TextMatrix(.Row, 15)), "0", .TextMatrix(.Row, 15)), .TextMatrix(.Row, 12), Val(.TextMatrix(.Row, 16)), _
                Val(.TextMatrix(.Row, 17)), Val(.TextMatrix(.Row, 18))) Then
                mshBill.SetFocus
                mshBill.Col = mconIntCol药名
                .Visible = False
                Exit Sub
            End If
            .Visible = False
            mshBill.Text = "[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 4)
            mshBill.Col = mconIntCol数量
            mshBill.SetFocus
        End If
    End With
End Sub

Private Sub mshDrug_LostFocus()
    SaveFlexState mshDrug, MStrCaption
    If mshDrug.Visible Then mshDrug.Visible = False
End Sub

Private Sub mshStructure_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub mshStructure_EnterCell(Row As Long, Col As Long)
    Call 提示组成库存数
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
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
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
    Dim rsDepend As New Recordset
    
    Dim lng制剂室 As Long
    Dim lng入库类别ID As Long
    Dim lng出库类别ID As Long
    Dim chrNo As Variant
    Dim int序号 As String
    Dim lng库房id As Long
    Dim lng药品id As Long
    Dim dbl填写数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str摘要 As String
    
    Dim intRow As Integer
    Dim n As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    
    arrSql = Array()
    
    On Error GoTo errHandle
    
    SaveCard = False
    With mshBill
        gstrSQL = "SELECT B.Id,b.系数, b.名称 " _
                  & " FROM 药品单据性质 A, 药品入出类别 B " _
                  & "Where A.类别id = B.ID " _
                & "AND A.单据 = 3  "
        Call SQLTest(App.Title, "协定药品入库管理", gstrSQL)
        Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "SaveCard")
        Call SQLTest
        If rsDepend.EOF Then
            MsgBox "没有设置协定药品入库的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        End If
        rsDepend.Filter = "系数=-1"
        If rsDepend.EOF Then
            MsgBox "没有设置协定药品入库的出库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        Else
            lng出库类别ID = rsDepend!Id
        End If
        rsDepend.Filter = "系数=1"
        If rsDepend.EOF Then
            MsgBox "没有设置协定药品入库的入库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
            rsDepend.Close
            Exit Function
        Else
            lng入库类别ID = rsDepend!Id
        End If
        rsDepend.Filter = adFilterNone
        rsDepend.Close
        
        chrNo = Trim(txtNo)
        lng库房id = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = Sys.GetNextNo(23, lng库房id)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")

        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_协定入库_Delete('" & mstr单据号 & "')"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lng药品id = .TextMatrix(intRow, 0)
                dbl填写数量 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                dbl成本价 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol采购价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dbl成本金额 = .TextMatrix(intRow, mconIntCol采购金额)
                dbl零售价 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dbl零售金额 = .TextMatrix(intRow, mconIntCol售价金额)
                dbl差价 = .TextMatrix(intRow, mconintCol差价)
                int序号 = intRow
                              
                gstrSQL = "zl_协定入库_INSERT("
                '入出类别ID
                gstrSQL = gstrSQL & lng入库类别ID
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & int序号
                '库房ID
                gstrSQL = gstrSQL & "," & lng库房id
                '药品ID
                gstrSQL = gstrSQL & "," & lng药品id
                '填写数量
                gstrSQL = gstrSQL & "," & dbl填写数量
                '成本价
                gstrSQL = gstrSQL & "," & dbl成本价
                '成本金额
                gstrSQL = gstrSQL & "," & dbl成本金额
                '零售价
                gstrSQL = gstrSQL & "," & dbl零售价
                '零售金额
                gstrSQL = gstrSQL & "," & dbl零售金额
                '差价
                gstrSQL = gstrSQL & "," & dbl差价
                '填制人
                gstrSQL = gstrSQL & ",'" & str填制人 & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS')"
                '摘要
                gstrSQL = gstrSQL & ",'" & str摘要 & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gstrSQL = "zl_药品协定对照出库_insert('" & chrNo & "'," & lng出库类别ID & "," & lng库房id & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        '集中处理退药事务
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim cur购价金额 As Double, Cur售价金额 As Double, Cur差价 As Double
    Dim intLop As Integer
    
    cur购价金额 = 0
    Cur售价金额 = 0
    Cur差价 = 0
    
    With mshBill
        For intLop = 1 To .rows - 1
            cur购价金额 = cur购价金额 + Val(.TextMatrix(intLop, mconIntCol采购金额))
            Cur售价金额 = Cur售价金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
        Next
    End With
    
    Cur差价 = Cur售价金额 - cur购价金额
    lblPurchasePrice.Caption = "购价金额合计：" & zlStr.FormatEx(cur购价金额, mintMoneyDigit, , True)
    lblSalePrice.Caption = "售价金额合计：" & zlStr.FormatEx(Cur售价金额, mintMoneyDigit, , True)
    lblDifference.Caption = "差价合计：" & zlStr.FormatEx(Cur差价, mintMoneyDigit, , True)
    
End Sub


Private Sub 提示库存数()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntCol药名) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    
    If RecTmp.State = 1 Then RecTmp.Close
    
    Select Case mintUnit
        Case mconint售价单位
            strUnit = "计算单位"
            strQuantity = "可用数量 "
        Case mconint门诊单位
            strUnit = "门诊单位"
            strQuantity = "可用数量/门诊包装 "
        Case mconint住院单位
            strUnit = "住院单位"
            strQuantity = "可用数量/住院包装 "
        Case mconint药库单位
            strUnit = "药库单位"
            strQuantity = "可用数量/药库包装 "
    End Select
    
    gstrSQL = " Select b.药品ID," & strUnit & " as 单位, Sum(" & strQuantity & ") as 数量 " & _
              " From 药品库存 a,药品规格 b,收费项目目录 C " & _
              " Where a.药品id=b.药品id and b.药品ID=C.ID " & _
              " and nvl(a.可用数量,0)<>0 and a.性质=1 And a.库房ID=[1] and b.药品ID=[2] " & _
              " Group by b.药品ID," & strUnit
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), intID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl数量 = IIf(IsNull(RecTmp!数量), 0, RecTmp!数量)
    
    staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(Dbl数量, mintNumberDigit, , True) & "]" & RecTmp!单位
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub 提示组成库存数()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshStructure.TextMatrix(mshStructure.Row, mconIntCol构药名) = "" Then
        Exit Sub
    End If
    
    intID = mshStructure.TextMatrix(mshStructure.Row, mconintcol构药品id)

    gstrSQL = "Select b.药品ID, nvl(Sum(可用数量),0) as 数量,c.计算单位 as 单位 " & _
        " from 药品库存 a,药品规格 b,收费项目目录 C " & _
        " Where b.药品ID=C.ID and b.药品id=a.药品id(+) AND a.库房ID(+)=[1] and a.性质(+)=1" & _
        " and b.药品ID=[2] " & _
        " Group by b.药品ID,c.计算单位 "
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), intID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl数量 = IIf(IsNull(RecTmp!数量), 0, RecTmp!数量)
    
    staThis.Panels(2).Text = "该药品当前库存数为[" & Dbl数量 & "]" & RecTmp!单位
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1344", "zl8_bill_1344"), mint记录状态, int单位系数, 1344, "药品协定入库单", strNo
End Sub


