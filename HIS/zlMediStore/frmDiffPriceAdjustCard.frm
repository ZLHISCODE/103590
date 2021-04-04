VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiffPriceAdjustCard 
   Caption         =   "库存差价调整单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDiffPriceAdjustCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   10
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   11
      Top             =   0
      Width           =   11715
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   5940
         TabIndex        =   31
         Top             =   945
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
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
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   1380
         TabIndex        =   1
         Top             =   615
         Width           =   2895
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   4290
         TabIndex        =   29
         Top             =   615
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   180
         TabIndex        =   3
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
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
         TabIndex        =   5
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位(&G)"
         Height          =   180
         Left            =   240
         TabIndex        =   30
         Top             =   660
         Width           =   990
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "调整额合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   26
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   1920
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "库存差价合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   22
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   19
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   18
         Top             =   158
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
         TabIndex        =   17
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
         TabIndex        =   4
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "库存差价调整单"
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   13
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   12
         Top             =   4500
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
            Picture         =   "frmDiffPriceAdjustCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1000
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
            Picture         =   "frmDiffPriceAdjustCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   27
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
            Picture         =   "frmDiffPriceAdjustCard.frx":22EA
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
            Picture         =   "frmDiffPriceAdjustCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDiffPriceAdjustCard.frx":3080
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
      TabIndex        =   23
      Top             =   5160
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
Attribute VB_Name = "frmDiffPriceAdjustCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '是否可选库房
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mblnFirst As Boolean                '第一次显示
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mint业务模式 As Integer             '1-库存差价调整;2-成本价调价
Private mlng供药单位ID As Long              '供药单位ID

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mstrPrivs As String                     '权限

Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mlng库房 As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称
Private Const MStrCaption As String = "库存差价调整管理"

'从参数表中取药品价格、数量、金额小数位数 精度
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol来源 As Integer = 4
Private Const mconIntCol基本药物 As Integer = 5
Private Const mconIntCol规格 As Integer = 6
Private Const mconIntCol批次 As Integer = 7
Private Const mconIntCol可用数量 As Integer = 8
Private Const mconIntCol比例系数 As Integer = 9
Private Const mconIntCol产地 As Integer = 10
Private Const mconIntCol单位 As Integer = 11
Private Const mconIntCol批号 As Integer = 12
Private Const mconIntCol效期 As Integer = 13
Private Const mconIntCol库存金额 As Integer = 14
Private Const mconIntCol库存差价 As Integer = 15
Private Const mconintCol成本价 As Integer = 16
Private Const mconintCol新成本价 As Integer = 17
Private Const mconintCol调整额 As Integer = 18
Private Const mconIntCol实际数量 As Integer = 19
Private Const mconIntCol药品编码和名称 As Integer = 20
Private Const mconIntCol药品编码 As Integer = 21
Private Const mconIntCol药品名称 As Integer = 22
Private Const mconIntColS  As Integer = 23              '总列数

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
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
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = n
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
Private Function Check应付记录(ByVal lng药品ID As Long, ByVal lng供药单位ID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
        " Where 系统标识=1 And 记录性质=0 And 收发id In (Select ID From 药品收发记录 " & _
        " Where 单据 = 1 And (Mod(记录状态, 3) = 0 Or 记录状态 = 1) And 药品id = [1] And 供药单位id = [2]) "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查应付记录]", lng药品ID, lng供药单位ID)
    
    If rsCheck.EOF Then
        Check应付记录 = True
        Exit Function
    Else
        Check应付记录 = (rsCheck!付款序号 = 0)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check库存(ByVal lng药品ID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    On Error GoTo errHandle
    strsql = "select Count(药品id) 库存 from 药品库存 Where 药品ID=[1] And 性质=1 And 实际数量>0 "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, MStrCaption & "[检查药品库存]", lng药品ID)
    
    Check库存 = (rsCheck!库存 > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check同一药品(ByVal lng药品ID As Long, ByVal intRow As Integer) As Boolean
    Dim n As Integer
        
    If intRow = 1 Then
        Check同一药品 = True
        Exit Function
    End If
    
    For n = 1 To mshBill.rows - 1
        If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
            If Val(mshBill.TextMatrix(n, 0)) = lng药品ID And n <> intRow Then
                Check同一药品 = False
                Exit Function
            End If
        End If
    Next
    
    Check同一药品 = True
End Function
Private Function Check药品供应商(ByVal lng药品ID As Long, ByVal lng供药单位ID As Long) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "Select Nvl(上次供应商ID,0) 上次供应商ID  From 药品库存 Where 药品id=[1] And 上次供应商id Is Not Null Order By nvl(批次,0) Desc "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, MStrCaption & "[检查药品供应商]", lng药品ID)
    If rsCheck.RecordCount = 0 Then
        Check药品供应商 = False
    Else
        Check药品供应商 = (rsCheck!上次供应商ID = lng供药单位ID)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'=========================================================================================
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    GetDepend = False
    strsql = "SELECT B.Id " _
           & "FROM 药品单据性质 A, 药品入出类别 B " _
           & "Where A.类别id = B.ID AND A.单据 = 5 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strsql, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品库存差价调整的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
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


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False, Optional int业务模式 As Integer = 1)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mblnFirst = True
    mint业务模式 = int业务模式
    mstrPrivs = GetPrivFunc(glngSys, 1303)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    If mint业务模式 = 1 Then
        LblTitle.Caption = "库存差价调整单"
        LblProvider.Visible = False
        txtProvider.Visible = False
        cmdProvider.Visible = False
    Else
        LblTitle.Caption = "成本价调整单"
        LblStock.Visible = False
        txtStock.Visible = False
        If mint编辑状态 <> 1 And mint编辑状态 <> 2 Then
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
        End If
    End If
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(IIf(mint业务模式 = 1, 2, 1), MStrCaption, IIf(mint业务模式 = 1, txtStock.Tag, 0), IIf(mint业务模式 = 1, txtStock.Tag, 0))
    End If
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

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID " & _
              "Order by level,ID"
    Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "取药品供应商", gstrNodeNo)
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "所有药品供应商"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = False Then Exit Sub
        
        Me.txtProvider.Tag = .CurrentID
        Me.txtProvider = .CurrentName
    End With
    Unload FrmSelect
    mshBill.SetFocus
    
    If Val(txtProvider.Tag) <> mlng供药单位ID Then
        mlng供药单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol行号) = "1"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    'mblnChange = False
    If mblnFirst = False Then Exit Sub
    
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint编辑状态 = 1 Then
        mshBill.ClearBill
        
        Dim str用途ID As String, str剂型编码 As String, strALL剂型编码 As String
        Dim str材质分类 As String, lng库房ID As Long, int差价波动率 As Integer
        
        If mint业务模式 = 1 Then
            If frmDiffPriceAdjustCondition.GetCondition(mfrmMain, str用途ID, str剂型编码, lng库房ID, int差价波动率) = True Then
                Screen.MousePointer = 11
                SearchData str用途ID, str剂型编码, lng库房ID, int差价波动率
                Screen.MousePointer = 0
            Else
                Unload Me
                Exit Sub
            End If
        Else
            Call RefreshRowNO(mshBill, mconIntCol行号, 1)
        End If
        
        If cmdCancel.Enabled = False Then
            cmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
        
        If txtProvider.Visible = True And mint业务模式 = 2 Then txtProvider.SetFocus
    Else
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
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
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
        mstrTime_End = GetBillInfo(5, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If

        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        If SaveCheck = True Then
            If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.差价调整)) = 1 Then
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
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.差价调整)) = 1 Then
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
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    txt摘要.Text = ""
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
End Sub

Private Sub Form_Load()
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    
    mlng库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    '为了处理特殊情况，把金额位数默认为最大位数
'    mintMoneyDigit = gtype_UserDrugDigits.Digit_金额
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "库存差价调整管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(5, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    If mint业务模式 = 1 Then
        mshBill.ColWidth(mconIntCol可用数量) = 0
        mshBill.ColWidth(mconIntCol批号) = 1000
        mshBill.ColWidth(mconIntCol效期) = 1000
        mshBill.ColWidth(mconintCol成本价) = 1200
    Else
        mshBill.ColWidth(mconIntCol可用数量) = 1000
        mshBill.ColWidth(mconIntCol批号) = 0
        mshBill.ColWidth(mconIntCol效期) = 0
        mshBill.ColWidth(mconintCol成本价) = 1200
    End If
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim strPrice As String
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPricedigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    
    '库房
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.差价调整)
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
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            txtStock = .List(.ListIndex)
            txtStock.Tag = .ItemData(.ListIndex)
            
        End With
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select distinct b.id,b.名称 from 药品收发记录 a,部门表 b  " _
                    & " where a.库房id=b.id and A.单据 =5 and  a.no=[1]"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!名称
                txtStock.Tag = rsInitCard!id
                
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量 as 可用数量,'1' as 比例系数,"
                    strPrice = ",A.新成本价 "
                Case mconint门诊单位
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 可用数量,B.门诊包装 as 比例系数,"
                    strPrice = ",A.新成本价*B.门诊包装 AS 新成本价 "
                Case mconint住院单位
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 可用数量,B.住院包装 as 比例系数,"
                    strPrice = ",A.新成本价*B.住院包装 AS 新成本价 "
                Case mconint药库单位
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 可用数量,B.药库包装 as 比例系数,"
                    strPrice = ",A.新成本价*B.药库包装 AS 新成本价 "
            End Select
            If mint业务模式 = 1 Then
                gstrSQL = "SELECT * " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    "     B.药品来源,B.基本药物,F.规格,A.产地, A.批号,A.效期,A.批次," & _
                    "     NVL(E.名称,F.名称) 名称," & strUnitQuantity & _
                    "     A.成本价 AS 库存差价,NVL(A.零售价,0) AS 库存金额,A.差价 AS 调整额, " & _
                    "     A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.填写数量 实际数量,A.单量 As 新成本价 " & _
                    "     FROM 药品收发记录 A, 药品规格 B,收费项目别名 E ,收费项目目录 F " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    "     AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    "     AND A.记录状态 =[2] AND A.单据 =5 AND A.NO = [1]) " & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT a.*,Rownum 序号 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    "     B.药品来源,B.基本药物,F.规格," & _
                    "     NVL(E.名称,F.名称) 名称," & strUnitQuantity & _
                    "     A.成本价 AS 库存差价,NVL(A.零售价,0) AS 库存金额,A.差价 AS 调整额, " & _
                    "     A.摘要,填制人,填制日期,审核人,审核日期,A.填写数量 实际数量 " & strPrice & ",G.名称 供应商,G.Id 供药单位id" & _
                    "     FROM (Select Sum(填写数量) 填写数量, 药品id, Sum(成本价) 成本价, Nvl(Sum(零售价), 0) 零售价, Sum(差价) 差价, 摘要, 填制人," & _
                    "     填制日期 , 审核人, 审核日期,单量 新成本价,供药单位id" & _
                    "     From 药品收发记录 " & _
                    "     Where 单据 = 5 And No = [1] And 记录状态 = [2] " & _
                    "     Group By 药品id, 摘要, 填制人, 填制日期, 审核人, 审核日期,单量,供药单位id) A, 药品规格 B,收费项目别名 E ,收费项目目录 F,供应商 G " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    "     AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 And A.供药单位id=G.Id ) A " & _
                    "  ORDER BY " & strSqlOrder
            End If
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt填制人 = rsInitCard!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户姓名
            End If
            Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
            
            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            With mshBill
                If mint业务模式 = 2 Then
                    txtProvider.Text = rsInitCard!供应商
                    mlng供药单位ID = rsInitCard!供药单位ID
                End If
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
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
                    
                    .TextMatrix(intRow, mconIntCol来源) = Nvl(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = Nvl(rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    .TextMatrix(intRow, mconIntCol库存金额) = zlStr.FormatEx(rsInitCard!库存金额, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol库存差价) = zlStr.FormatEx(IIf(IsNull(rsInitCard!库存差价), 0, rsInitCard!库存差价), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol调整额) = zlStr.FormatEx(rsInitCard!调整额, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量), intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol实际数量) = zlStr.FormatEx(IIf(IsNull(rsInitCard!实际数量), "0", rsInitCard!实际数量), intNumberDigit, , True)
                    If mint业务模式 = 1 Then
                        .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                        .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                        .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                        If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                            '换算为有效期
                            .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                        End If
                        If Not IsNull(rsInitCard!新成本价) Then
                            .TextMatrix(intRow, mconintCol新成本价) = zlStr.FormatEx(rsInitCard!新成本价 * rsInitCard!比例系数, intCostDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx((rsInitCard!库存金额 - rsInitCard!库存差价) / rsInitCard!可用数量, intCostDigit, , True)
                        If Not IsNull(rsInitCard!新成本价) Then
                            .TextMatrix(intRow, mconintCol新成本价) = zlStr.FormatEx(rsInitCard!新成本价 * rsInitCard!比例系数, intCostDigit, , True)
                        End If
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'初始化编辑控件
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol库存差价) = "库存差价"
        .TextMatrix(0, mconIntCol库存金额) = "库存金额"
        .TextMatrix(0, mconintCol调整额) = "调整额"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconintCol成本价) = "成本价"
        .TextMatrix(0, mconintCol新成本价) = "新成本价"
        .TextMatrix(0, mconIntCol实际数量) = "实际数量"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol药名) = 2500
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol规格) = 1000
        .ColWidth(mconIntCol产地) = 1000
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol库存金额) = 1200
        .ColWidth(mconIntCol库存差价) = 1200
        .ColWidth(mconintCol调整额) = 1200
        .ColWidth(mconintCol新成本价) = 1200
        .ColWidth(mconIntCol实际数量) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        
        If mint业务模式 = 1 Then
            .ColWidth(mconIntCol可用数量) = 0
            .ColWidth(mconIntCol批号) = 1000
            .ColWidth(mconIntCol效期) = 1000
            .ColWidth(mconintCol成本价) = 1200
        Else
            .ColWidth(mconIntCol可用数量) = 1000
            .ColWidth(mconIntCol批号) = 0
            .ColWidth(mconIntCol效期) = 0
            .ColWidth(mconintCol成本价) = 1200
        End If
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol库存差价) = 5
        .ColData(mconIntCol库存金额) = 5
        .ColData(mconIntCol批次) = 5
        .ColData(mconIntCol可用数量) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol实际数量) = 5
        .ColData(mconintCol成本价) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mconIntCol药名) = 1
            .ColData(mconintCol新成本价) = 4
            If mint业务模式 = 1 Then
                .ColData(mconintCol调整额) = 4
            Else
                .ColData(mconintCol调整额) = 5
            End If
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            .ColData(mconintCol调整额) = 5
            .ColData(mconintCol新成本价) = 5
        End If
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol库存金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol库存差价) = flexAlignRightCenter
        .ColAlignment(mconintCol调整额) = flexAlignRightCenter
        .ColAlignment(mconintCol成本价) = flexAlignRightCenter
        .ColAlignment(mconintCol新成本价) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
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
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - cmdCancel.Height - 200
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
        LblNO.Left = .Left - LblNO.Width - 100
        .Top = LblTitle.Top
        LblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
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
        '.Width = .Left - .Left
        Debug.Print .Width
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
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
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With cmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With
    
    With CmdSave
        .Left = cmdCancel.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = cmdCancel.Top
    End With
        
    With cmdFind
        .Top = cmdCancel.Top
    End With
    
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\库存差价调整管理", "药品名称显示方式", mintDrugNameShow)
    
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelLength = Len(txtProvider.Text)
        txtProvider.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
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
    Dim n As Integer
    
    mblnSave = False
    SaveCheck = False
    
    str审核人 = UserInfo.用户姓名
    strNo = txtNo.Tag
    On Error GoTo errHandle
    
    '检查应付记录，如果已经付款，则不能调整成本价
    If mint业务模式 = 2 Then
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
                If Not Check应付记录(Val(mshBill.TextMatrix(n, 0)), mlng供药单位ID) Then
                    MsgBox mshBill.TextMatrix(n, mconIntCol药名) & " 已全付款，不能调整成本价！", vbInformation + vbOKOnly, gstrSysName
                    mshBill.SetFocus
                    mshBill.Col = mconIntCol药名
                    Exit Function
                End If
            End If
        Next
    End If
                            
    gstrSQL = "zl_药品库存差价调整_Verify('" & strNo & "','" & str审核人 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
   
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "审核失败！", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
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
End Sub
Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol药名) = 0 Then
        'Cancel = True    '等待加CANCEL参数
        Exit Sub
    End If
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
    Dim RecReturn As Recordset
    Dim str药名 As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
'    Set RecReturn = Frm药品选择器.ShowME(Me, IIf(mint业务模式 = 1, 2, 1), IIf(mint业务模式 = 1, txtStock.Tag, 0), , , False)
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(IIf(mint业务模式 = 1, 2, 1), MStrCaption, IIf(mint业务模式 = 1, txtStock.Tag, 0), IIf(mint业务模式 = 1, txtStock.Tag, 0))
    End If
    Set RecReturn = frmSelector.showMe(Me, 0, IIf(mint业务模式 = 1, 2, 1), , , , IIf(mint业务模式 = 1, txtStock.Tag, 0), , , False, , , , , , mstrPrivs & ";查看成本价;")
    mshBill.CmdEnable = True
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
'        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
'            str药名 = RecReturn!通用名
'        Else
'            str药名 = IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
'        End If
            
        If mint业务模式 = 2 Then
            '检查药品重复
'            If Not Check同一药品(RecReturn!药品ID, mshBill.Row) Then
'                MsgBox "药品" & str药名 & "已存在，请重新输入！", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntCol药名
'                Exit Sub
'            End If
'
'            '检查药品库存
'            If Not Check库存(RecReturn!药品ID) Then
'                MsgBox "药品" & str药名 & " 在所有库房都无库存，不能调整成本价！", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntCol药名
'                Exit Sub
'            End If
'
'            '检查药品供应商关系
'            If Not Check药品供应商(RecReturn!药品ID, mlng供药单位ID) Then
'                MsgBox txtProvider.Text & "不是药品" & str药名 & " 的供药单位，请重新选择药品或者供药单位！", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntCol药名
'                Exit Sub
'            End If
'
'            '检查应付记录，如果已经付款，则不能调整成本价
'            If Not Check应付记录(RecReturn!药品ID, mlng供药单位ID) Then
'                MsgBox str药名 & " 已全付款，不能调整成本价！", vbInformation + vbOKOnly, gstrSysName
'                mshBill.SetFocus
'                mshBill.Col = mconIntCol药名
'                Exit Sub
'            End If
            Set RecReturn = CheckData(RecReturn)
        End If
        If RecReturn.RecordCount > 0 Then
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                intRow = mshBill.Row
                With mshBill
                    .TextMatrix(intRow, mconIntCol行号) = .Row
                    SetColValue .Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", _
                        RecReturn!商品名), Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                        IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                        Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                        IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                        IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                        IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
                        IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
                        Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                        IIf(IsNull(RecReturn!库存数量), "0", RecReturn!库存数量)
                    
                    .Col = mconintCol新成本价
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                    RecReturn.MoveNext
                End With
            Next
            mshBill.Row = intOldRow
            RecReturn.Close
        End If
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
        If strkey = "" Then
            strkey = .TextMatrix(.Row, .Col)
        End If
        Select Case .Col
            Case mconintCol新成本价
               intDigit = mintCostDigit
            Case mconintCol调整额
                intDigit = mintMoneyDigit
        End Select
        
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If .SelLength = Len(strkey) Then Exit Sub
            If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        Select Case .Col
            Case mconIntCol药名
                .TxtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mconintCol调整额
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890-"
            Case mconintCol成本价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconintCol新成本价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim str药名 As String
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntCol药名
                If strkey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    Dim i As Integer
                    Dim intCurRow As Integer
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, IIf(mint业务模式 = 1, 2, 1), IIf(mint业务模式 = 1, txtStock.Tag, 0), , , strkey, sngLeft, sngTop, False)
                    
                    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                        Call SetSelectorRS(IIf(mint业务模式 = 1, 2, 1), MStrCaption, IIf(mint业务模式 = 1, txtStock.Tag, 0), IIf(mint业务模式 = 1, txtStock.Tag, 0))
                    End If
                    Set RecReturn = frmSelector.showMe(Me, 1, IIf(mint业务模式 = 1, 2, 1), strkey, sngLeft, sngTop, IIf(mint业务模式 = 1, txtStock.Tag, 0), , , False, , , , , , mstrPrivs & ";查看成本价;")
'                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
'                        str药名 = RecReturn!通用名
'                    Else
'                        str药名 = IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
'                    End If
            
                    If mint业务模式 = 2 Then
                    '检查药品重复
'                        If Not Check同一药品(RecReturn!药品ID, mshBill.Row) Then
'                            MsgBox "药品" & str药名 & "已存在，请重新输入！", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntCol药名
'                            Cancel = True
'                            Exit Sub
'                        End If
'
'                        '检查药品库存
'                        If Not Check库存(RecReturn!药品ID) Then
'                            MsgBox "药品" & str药名 & " 在所有库房都无库存，不能调整成本价！", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            mshBill.Col = mconIntCol药名
'                            Exit Sub
'                        End If
'
'                        '检查药品供应商关系
'                        If Not Check药品供应商(RecReturn!药品ID, mlng供药单位ID) Then
'                            MsgBox txtProvider.Text & "不是药品" & str药名 & " 的供药单位，请重新选择药品或者供药单位！", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntCol药名
'                            Exit Sub
'                        End If
'
'                        '检查应付记录，如果已经付款，则不能调整成本价
'                        If Not Check应付记录(RecReturn!药品ID, mlng供药单位ID) Then
'                            MsgBox str药名 & " 已全付款，不能调整成本价！", vbInformation + vbOKOnly, gstrSysName
'                            mshBill.SetFocus
'                            .Col = mconIntCol药名
'                            Exit Sub
'                        End If
                        If RecReturn.RecordCount > 0 Then
                            Set RecReturn = CheckData(RecReturn)
                        End If
                    End If
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                    Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                    IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                    IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
                                    IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
                                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), IIf(IsNull(RecReturn!库存数量), "0", RecReturn!库存数量)) = False Then
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            
                            Call 提示库存数
                        
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        Cancel = True
                    End If
                End If
            Case mconintCol新成本价
                If strkey = "" And mint业务模式 = 1 Then
                    .Col = mconintCol调整额
                    Cancel = True
                    Exit Sub
                End If
                
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，成本价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) < 0.001 Then
                        MsgBox "对不起，成本价必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "成本价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strkey, mintCostDigit, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
      
                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintCostDigit, , True)
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol新成本价) = .Text
                End If
                                
                '重算差价调整额(调整额＝库存金额－可用数量*成本价-库存差价)
                If strkey <> "" Then
                    .TextMatrix(.Row, mconintCol调整额) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol库存金额) = "", 0, .TextMatrix(.Row, mconIntCol库存金额)) - Val(IIf(.TextMatrix(.Row, mconIntCol可用数量) = "", 0, .TextMatrix(.Row, mconIntCol可用数量))) * Val(IIf(.TextMatrix(.Row, mconintCol新成本价) = "", 0, .TextMatrix(.Row, mconintCol新成本价))) _
                        - Val(IIf(.TextMatrix(.Row, mconIntCol库存差价) = "", 0, .TextMatrix(.Row, mconIntCol库存差价))), mintMoneyDigit, , True)
                End If
                
            Case mconintCol调整额
                If .TextMatrix(.Row, .Col) = "" And strkey = "" Then
                    MsgBox "对不起，调整额必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，调整额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) = 0 Then
                        MsgBox "对不起，调整额不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strkey)) < 0.00001 Then
                        MsgBox "对不起，调整额的绝对值必须不小于0.00001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "调整额必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = zlStr.FormatEx(strkey, mintMoneyDigit, , True)
                    .Text = strkey
                    
                    '重算成本价(成本价=(库存金额-库存差价-调整额)/可用数量)
                    If strkey <> "" And Val(.TextMatrix(.Row, mconIntCol可用数量)) <> 0 Then
                        .TextMatrix(.Row, mconintCol新成本价) = zlStr.FormatEx((IIf(.TextMatrix(.Row, mconIntCol库存金额) = "", 0, .TextMatrix(.Row, mconIntCol库存金额)) - Val(IIf(.TextMatrix(.Row, mconIntCol库存差价) = "", 0, .TextMatrix(.Row, mconIntCol库存差价))) - Val(strkey)) / Val(IIf(.TextMatrix(.Row, mconIntCol可用数量) = "", 0, .TextMatrix(.Row, mconIntCol可用数量))), mintCostDigit, , True)
                    End If
                End If
                Call 显示合计金额
        End Select
    End With
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal int药品id As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, _
    ByVal str基本药物 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal str批号 As String, ByVal str效期 As String, _
    ByVal num库存差价 As Double, ByVal lng批次 As Long, ByVal num可用数量 As Double, _
    ByVal num比例系数 As Double, ByVal num库存金额 As Double, ByVal num实际数量 As Double) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim str药名 As String
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
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
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol单位) = str单位
        
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        .TextMatrix(intRow, mconIntCol库存金额) = zlStr.FormatEx(num库存金额, mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol库存差价) = zlStr.FormatEx(num库存差价, mintMoneyDigit, , True)
        
        If mint业务模式 = 1 Then
            If lng批次 > 0 Then
                .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(num可用数量, mintNumberDigit, , True)
            Else
                .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(num可用数量 / num比例系数, mintNumberDigit, , True)
            End If
            .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Get成本价(int药品id, txtStock.Tag, Val(lng批次)) * num比例系数, mintCostDigit, , True)
        Else
            .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(num可用数量 / num比例系数, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx((((num库存金额 - num库存差价)) / num可用数量) / num比例系数, mintCostDigit, , True)
        End If
        .TextMatrix(intRow, mconIntCol实际数量) = .TextMatrix(intRow, mconIntCol可用数量)
    End With
    SetColValue = True
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

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub


Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelStart = 0
        txtProvider.SelLength = Len(txtProvider.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtProvider.Text = mshProvider.TextMatrix(mshProvider.Row, 2)
        txtProvider.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        mshBill.SetFocus
    End If

    If Val(txtProvider.Tag) <> mlng供药单位ID Then
        mshBill.ClearBill
        mlng供药单位ID = Val(txtProvider.Tag)
        mshBill.TextMatrix(1, mconIntCol行号) = "1"
    End If
End Sub


Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
    End If
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
            If mint业务模式 = 2 Then
                If mlng供药单位ID = 0 Then
                    MsgBox "对不起，供药单位不能为空！", vbOKOnly + vbInformation, gstrSysName
                    txtProvider.SetFocus
                    Exit Function
                End If
            End If
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If mint业务模式 = 2 Then
                    If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                        If Trim(.TextMatrix(intLop, mconintCol新成本价)) = "" Then
                            MsgBox "对不起，新成本价不能为空！", vbOKOnly + vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconintCol新成本价
                            Exit Function
                        End If
                    End If
                End If
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol调整额))) = "" Then
                        MsgBox "第" & intLop & "行药品的调整额为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol调整额
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol调整额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的调整额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol调整额
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
    Dim lng入出类别id As Long
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim str批号 As String
    Dim lng批次ID As Long
    Dim str产地 As String
    Dim dat效期 As String
    Dim dbl可用数量 As Double
    Dim dbl库存差价 As Double
    Dim dbl库存金额 As Double
    Dim dbl调整额 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim dat填制日期 As String
    Dim rs入出类别 As New Recordset
    Dim dbl新成本价 As Double
    
    Dim intRow As Integer
    Dim n As Integer
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveCard = False
    arrSql = Array()
    On Error GoTo errHandle
    '在外面设置入出类别ID，主要是所有药品都要用他
    gstrSQL = "SELECT B.Id " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 5 "
    Set rs入出类别 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
    
    If rs入出类别.EOF Then
        MsgBox "没有设置药品库存差价调整的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rs入出类别.Close
        Exit Function
    End If
    lng入出类别id = rs入出类别.Fields(0)
    rs入出类别.Close
   
    With mshBill
        chrNo = Trim(txtNo)
        lng库房ID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(25, lng库房ID)
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        dat填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_药品库存差价调整_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lng药品ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                lng批次ID = Val(.TextMatrix(intRow, mconIntCol批次))
                dat效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                
                dbl可用数量 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol实际数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_数量, , True)
                dbl库存金额 = .TextMatrix(intRow, mconIntCol库存金额)
                dbl库存差价 = .TextMatrix(intRow, mconIntCol库存差价)
                dbl调整额 = .TextMatrix(intRow, mconintCol调整额)
                dbl新成本价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol新成本价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_成本价, , True)
                lng序号 = intRow
                
                'zl_药品库存差价调整_INSERT( /*入出类别ID_IN*/, /*NO_IN*/, /*序号_IN*/,
                    '/*库房ID_IN*/, /*药品ID_IN*/, /*批次_IN*/, /*可用数量_IN*/,
                    '/*库存差价_IN*/, /*调整额_IN*/, /*填制人_IN*/, /*填制日期_IN*/,
                    '/*产地_IN*/, /*批号_IN*/, /*效期_IN*/, /*摘要_IN*/ );
                    
                gstrSQL = "zl_药品库存差价调整_INSERT(" & lng入出类别id & ",'" & chrNo & "'," & lng序号 & "," _
                    & lng库房ID & "," & lng药品ID & "," & lng批次ID & "," & dbl可用数量 & "," _
                    & dbl库存金额 & "," & dbl库存差价 & "," & dbl调整额 & ",'" & str填制人 & "',to_date('" & dat填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & str产地 & "','" & str批号 & "'," & IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & str摘要 & "'," & mlng供药单位ID & "," & dbl新成本价 & "," & IIf(mint业务模式 = 1, 0, 1) & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim dbl库存差价 As Double
    Dim dbl调整额 As Double
    Dim dbl库存金额 As Double
    
    Dim intLop As Integer
    
    dbl库存差价 = 0
    dbl调整额 = 0
    
    With mshBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl库存差价 = dbl库存差价 + Val(.TextMatrix(intLop, mconIntCol库存差价))
                dbl库存金额 = dbl库存金额 + Val(.TextMatrix(intLop, mconIntCol库存金额))
                dbl调整额 = dbl调整额 + Val(.TextMatrix(intLop, mconintCol调整额))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "库存金额合计：" & zlStr.FormatEx(dbl库存金额, mintMoneyDigit, , True)
    lblSalePrice.Caption = "库存差价合计：" & zlStr.FormatEx(dbl库存差价, mintMoneyDigit, , True)
    lblDifference.Caption = "调整额合计：" & zlStr.FormatEx(dbl调整额, mintMoneyDigit, , True)
    
End Sub

Private Sub 提示库存数()
    
    If mint编辑状态 = 4 Then Exit Sub
    With mshBill
        If .TextMatrix(.Row, mconIntCol药名) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol可用数量), mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol单位)
    End With
End Sub

Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    mblnChange = True
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub


Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑状态 = 3 Or mint编辑状态 = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"
        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If adoProvider.EOF Then
            MsgBox "没有你输入的供药单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1000
                .ColWidth(2) = 2700
                .ColWidth(3) = 1200
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Top = txtProvider.Top + txtProvider.Height
                .Left = cmdProvider.Left + cmdProvider.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!名称
            .Tag = adoProvider!id
        End If
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If Val(.Tag) <> mlng供药单位ID Then
            mlng供药单位ID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mconIntCol行号) = "1"
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub


Private Sub txtProvider_Validate(Cancel As Boolean)
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng供药单位ID Then
        mlng供药单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol行号) = "1"
    End If
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
    FrmBillPrint.showMe Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1303", "zl8_bill_1303"), mint记录状态, int单位系数, 1303, "药品差价调整单", strNo
End Sub

Private Sub SearchData(ByVal str用途ID, ByVal str剂型编码 As String, _
    ByVal lng库房ID As Long, ByVal intRate As Integer)
    
    Dim rsData As New Recordset  '药品库存记录集
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim intRecordCount As Integer
    Dim strUnitQuantity As String
    Dim str药名 As String
    Dim strUseID As String, strClassID As String
    
    On Error GoTo errHandle:
    '设置界面显示内容
    staThis.Panels(2).Text = "现在对" & txtStock & "的药品进行自动差价计算"
    '构造药品查询条件(药品目录 A)
    strPhysic = " And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)"
    If str剂型编码 = "" Then str剂型编码 = "'ZYB'"
    
    If str用途ID <> "" Then
        If InStr(1, "'中成药','中草药','西成药'", str用途ID) <> 0 Then
            Select Case str用途ID
            Case "'西成药'"
                strClassID = "1"
            Case "'中成药'"
                strClassID = "2"
            Case Else
                strClassID = "3"
            End Select
            strPhysic = strPhysic & " And F.类型 = [5] "
        Else
            strUseID = str用途ID
            strPhysic = strPhysic & " And M.分类ID in (select * from Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And F.类型 In ('1','2','3') "     '数据量不大 In 未作优化处理
        End If
    End If
    
    DoEvents    ': Me.Refresh

    Select Case mintUnit
        Case mconint售价单位
            strUnitQuantity = "C.计算单位 AS 单位, nvl(b.实际数量,0) AS 可用数量, '1' as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价,b.平均成本价) 成本价,"
        Case mconint门诊单位
            strUnitQuantity = "a.门诊单位 AS 单位,(nvl(b.实际数量,0)/a.门诊包装) AS 可用数量,a.门诊包装 as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价*a.门诊包装,b.平均成本价*a.门诊包装) 成本价,"
        Case mconint住院单位
            strUnitQuantity = "a.住院单位 AS 单位, (nvl(b.实际数量,0)/a.住院包装) AS 可用数量, a.住院包装 as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价*a.住院包装,b.平均成本价*a.住院包装) 成本价,"
        Case mconint药库单位
            strUnitQuantity = "a.药库单位 AS 单位, (nvl(b.实际数量,0)/a.药库包装) AS 可用数量,a.药库包装 as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价*a.药库包装,b.平均成本价*a.药库包装) 成本价,"
    End Select

    gstrSQL = "SELECT DISTINCT B.药品ID,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, D.名称 As 商品名," & _
        " A.药品来源,A.基本药物,C.规格,NVL(B.上次产地,C.产地) AS 产地,B.批次,B.上次批号 AS 批号, B.效期," & _
        " B.实际金额, B.实际差价," & strUnitQuantity & _
        " DECODE(SIGN (B.实际差价/B.实际金额*100-(A.指导差价率+[3])),1,-(实际差价-B.实际金额*A.指导差价率/100)," & _
        " DECODE(SIGN(B.实际差价/B.实际金额*100-(A.指导差价率-[3])),-1,B.实际金额*A.指导差价率/100-实际差价)) AS 差价调整额,NVL(b.实际数量,0) 实际数量 " & _
        " FROM 药品规格 A,(SELECT 库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率,平均成本价 FROM 药品库存 WHERE NVL(实际金额,0)<>0) B," & _
        " 收费项目目录 C,收费项目别名 D,药品特性 T,诊疗分类目录 F,诊疗项目目录 M"
    
    gstrSQL = gstrSQL & " WHERE A.药品ID = C.ID and A.药名ID=T.药名ID " & _
        " And T.药名ID=M.ID And M.分类ID=F.ID " & _
        " AND A.药品ID=D.收费细目ID(+) AND D.性质(+)=3 AND D.码类(+)=1 " & _
        " AND B.性质=1 AND B.库房ID=[1] AND A.药品ID=B.药品ID " & _
        " AND (B.实际差价/NVL(B.实际金额,1)*100>(A.指导差价率+[3]) OR B.实际差价/NVL(B.实际金额,1)*100<A.指导差价率-[3])" & strPhysic
        
    If str剂型编码 <> "" Then
        gstrSQL = gstrSQL & " And T.药品剂型 in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) "
    End If
    
    gstrSQL = gstrSQL & " ORDER BY 药品编码"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[正在计算药品库存数据]", lng库房ID, str剂型编码, intRate, strUseID, strClassID)
    
    intRecordCount = rsData.RecordCount
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    If intRecordCount = 0 Then
        MsgBox "未能正确读取药品库存数据,请重试或手工输入药品！", vbInformation, gstrSysName: Exit Sub
    End If
    
    DoEvents: 'Me.Refresh
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .rows = .rows + 1
            .TextMatrix(i, 0) = rsData!药品id
           
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                str药名 = rsData!通用名
            Else
                str药名 = IIf(IsNull(rsData!商品名), rsData!通用名, rsData!商品名)
            End If
            
            .TextMatrix(i, mconIntCol药品编码和名称) = rsData!药品编码 & str药名
            .TextMatrix(i, mconIntCol药品编码) = rsData!药品编码
            .TextMatrix(i, mconIntCol药品名称) = str药名
            
            If mintDrugNameShow = 1 Then
                .TextMatrix(i, mconIntCol药名) = .TextMatrix(i, mconIntCol药品编码)
            ElseIf mintDrugNameShow = 2 Then
                .TextMatrix(i, mconIntCol药名) = .TextMatrix(i, mconIntCol药品名称)
            Else
                .TextMatrix(i, mconIntCol药名) = .TextMatrix(i, mconIntCol药品编码和名称)
            End If
            
            .TextMatrix(i, mconIntCol商品名) = IIf(IsNull(rsData!商品名), "", rsData!商品名)
            .TextMatrix(i, mconIntCol来源) = IIf(IsNull(rsData!药品来源), "", rsData!药品来源)
            .TextMatrix(i, mconIntCol基本药物) = IIf(IsNull(rsData!基本药物), "", rsData!基本药物)

            .TextMatrix(i, mconIntCol规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
            .TextMatrix(i, mconIntCol产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
            .TextMatrix(i, mconIntCol单位) = IIf(IsNull(rsData!单位), "", rsData!单位)
            .TextMatrix(i, mconIntCol批次) = IIf(IsNull(rsData!批次), "0", rsData!批次)
            .TextMatrix(i, mconIntCol批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
            .TextMatrix(i, mconIntCol效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(i, mconIntCol效期) <> "" Then
                '换算为有效期
                .TextMatrix(i, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(i, mconIntCol效期)), "yyyy-mm-dd")
            End If
           
            .TextMatrix(i, mconIntCol可用数量) = rsData!可用数量
            .TextMatrix(i, mconIntCol实际数量) = rsData!实际数量 / rsData!比例系数
            .TextMatrix(i, mconIntCol库存金额) = zlStr.FormatEx(rsData!实际金额, mintMoneyDigit, , True)
            .TextMatrix(i, mconIntCol库存差价) = zlStr.FormatEx(rsData!实际差价, mintMoneyDigit, , True)
            .TextMatrix(i, mconintCol调整额) = zlStr.FormatEx(rsData!差价调整额, mintMoneyDigit, , True)
            .TextMatrix(i, mconIntCol比例系数) = rsData!比例系数
            .TextMatrix(i, mconintCol成本价) = zlStr.FormatEx(rsData!成本价, mintCostDigit, , True)
                
            Call zlControl.StaShowPercent(i / intRecordCount, staThis.Panels(2), frmDiffPriceAdjustCard)
            i = i + 1
            rsData.MoveNext
        Loop
        .Redraw = True
    End With
    rsData.Close
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    
    staThis.Panels(2).Text = ""
    mshBill.Row = 1
    mshBill.Col = mconintCol调整额
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    mshBill.Redraw = True
    Call SaveErrLog
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：用来检查列表中已有药品与新选择的药品是否重复和时价药品是否有库存

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str库存 As String
    Dim strsql As String
    Dim strDub As String    '重复药品
    Dim strNotNum As String  '无库存药品
    Dim str重复药名 As String   '用来记录重复选择了的药品名称
    Dim strNot药名 As String    '用来记录哪些药品是时价但无库存
    Dim bln供应商 As Boolean    '验证该药品是否和选择的供应商相同
    Dim str供应商 As String
    Dim strPro As String
    Dim strProvider As String
    Dim bln是否付款 As Boolean
    Dim str是否付款 As String
    Dim strPay As String
    Dim strToP As String
    Dim strmsg供应商 As String
    Dim strmsg是否重复 As String
    Dim strmsg库存 As String
    Dim strmsg是否付款 As String
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str批次 = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
        
        If rsTemp!时价 = 1 Then '将时价无库存的记录找出来
            gstrSQL = "select Decode(Nvl(批次,0),0,实际金额/实际数量,Nvl(零售价,实际金额/实际数量))*" & Choose(mintUnit, 1, rsTemp!门诊包装, rsTemp!住院包装, rsTemp!药库包装) & " as  售价 " _
                & "  from 药品库存 " _
                & " where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1 and 实际数量>0 and " _
                & " nvl(批次,0)=[3]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!药品id, IIf(IsNull(rsTemp!批次), 0, rsTemp!批次))
            If rsPrice.EOF Then
                str库存 = str库存 & rsTemp!药品id & "," & rsTemp!通用名 & "|"
            End If
        End If
        
        bln供应商 = Check药品供应商(rsTemp!药品id, mlng供药单位ID)  '检查药品的供应商
        If bln供应商 = False Then
            str供应商 = str供应商 & rsTemp!药品id & "," & rsTemp!通用名 & "|"
        End If
        
        bln是否付款 = Check应付记录(rsTemp!药品id, mlng供药单位ID)  '检查是否付款
        If bln是否付款 = False Then
            str是否付款 = str是否付款 & rsTemp!药品id & "," & rsTemp!通用名 & "|"
        End If
        
        rsTemp.MoveNext
    Loop
        
    With mshBill    '把重复的查询出来
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
        If str库存 <> "" Then
            strNotNum = ""
            For i = 0 To UBound(Split(str库存, "|")) - 1
                strNotNum = strNotNum & "药品id<>" & Split(Split(str库存, "|")(i), ",")(0) & " and "
                If UBound(Split(strNot药名, ",")) <= 2 Then
                    strNot药名 = strNot药名 & Split(Split(str库存, "|")(i), ",")(1) & ","
                End If
            Next
            If strNotNum <> "" Then
                strNotNum = Mid(strNotNum, 1, Len(strNotNum) - 4)
            End If
        End If
        If str供应商 <> "" Then
            strProvider = ""
            For i = 0 To UBound(Split(str供应商, "|")) - 1
                strProvider = strProvider & "药品id<>" & Split(Split(str供应商, "|")(i), ",")(0) & " and "
                If UBound(Split(strPro, ",")) <= 2 Then
                    strPro = strPro & Split(Split(str供应商, "|")(i), ",")(1) & ","
                End If
            Next
            If strProvider <> "" Then
                strProvider = Mid(strProvider, 1, Len(strProvider) - 4)
            End If
        End If
        If str是否付款 <> "" Then
            strProvider = ""
            For i = 0 To UBound(Split(str是否付款, "|")) - 1
                strPay = strPay & "药品id<>" & Split(Split(str是否付款, "|")(i), ",")(0) & " and "
                If UBound(Split(strToP, ",")) <= 2 Then
                    strToP = strToP & Split(Split(str是否付款, "|")(i), ",")(1) & ","
                End If
            Next
            If strPay <> "" Then
                strPay = Mid(strPay, 1, Len(strPay) - 4)
            End If
        End If
        
        
        '判断以什么方式拼接sql
        strsql = strDub & " " & strNotNum & " " & strProvider & " " & strPay
        If str重复药名 <> "" Then
            strmsg是否重复 = str重复药名 & "列表中已经含有了！"
            strsql = strDub
        End If
        If strNot药名 <> "" Then
            strmsg库存 = vbCrLf & strNot药名 & "是时价药品，没有库存不允许出库！"
            If strsql = "" Then
                strsql = strNotNum
            Else
                strsql = strsql & " and " & strNotNum
            End If
        End If
        If strPro <> "" Then
            strmsg供应商 = vbCrLf & strPro & "是时价药品，没有库存不允许出库！"
            If strsql = "" Then
                strsql = strProvider
            Else
                strsql = strsql & " and " & strProvider
            End If
        End If
        If strToP <> "" Then
            strmsg是否付款 = vbCrLf & strToP & "是时价药品，没有库存不允许出库！"
            If strsql = "" Then
                strsql = strPay
            Else
                strsql = strsql & " and " & strPay
            End If
        End If
        If strmsg是否重复 <> "" Or strmsg库存 <> "" Or strmsg供应商 <> "" Or strmsg是否付款 <> "" Then
            MsgBox strmsg是否重复 & strmsg库存 & strmsg供应商 & strmsg是否付款 & "...以上药品将不再添加！", vbInformation, gstrSysName
        End If
        
        If strsql <> "" Then
            rsTemp.Filter = strsql
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

