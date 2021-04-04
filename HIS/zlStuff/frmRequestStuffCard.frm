VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmRequestStuffCard 
   Caption         =   "卫材申领单"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmRequestStuffCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdRequest 
      Caption         =   "按申购单申领(&R)"
      Height          =   350
      Left            =   3840
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6345
      TabIndex        =   28
      Top             =   5535
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7665
      TabIndex        =   27
      Top             =   5535
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
         Top             =   950
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
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   557
         Width           =   1515
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9960
         TabIndex        =   17
         Top             =   550
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
         Top             =   587
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
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料申领单"
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
         Caption         =   "发料库房(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
            Picture         =   "frmRequestStuffCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1000
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
            Picture         =   "frmRequestStuffCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   6495
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestStuffCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestStuffCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestStuffCard.frx":3080
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.Label lblCode 
      Caption         =   "材料"
      Height          =   255
      Left            =   3255
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmRequestStuffCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、通过向导新增；6、接受（接收后记录接收登记人，可以取消错误的接收）；7、拒收
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑

Private mint明确批次 As Integer             '表示在填写申领单时，是否明确卫材的批次
Private mint库存检查 As Integer             '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                     '权限
Private mlngStockID As Long                 '当前用户所选的发料部门ID
Private rsDepend As New ADODB.Recordset

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private Const mlngModule = 1722
Private mint仅显示有库存物资 As Boolean
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容
Private Const mstrCaption As String = "卫材申领单"
Private mbln申领核查 As Boolean             '移库是否需要核查,true-需要，false-不需要
Private mint处理方式 As Integer             '冲销时：0－正常冲销；1－产生冲销申请单据
Private mstr重复卫材 As String '记录重复的卫材

Private mstrRequestNO As String     '按申购单移库NO ，空代表不按照申购单方式申领，否则按照申购单申领
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mstrTime_Start As String                        '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

'=========================================================================================
Private Enum mBillCol
    C_行号 = 1
    C_材料 = 2
    c_规格 = 3
    C_序号 = 4
    C_库房分批 = 5
    C_最大效期 = 6
    C_可用数量 = 7
    C_指导差价率 = 8
    C_实际金额 = 9
    C_实际差价 = 10
    C_比例系数 = 11
    c_批次 = 12
    C_产地 = 13
    C_批准文号 = 14
    c_单位 = 15
    c_批号 = 16
    C_效期 = 17
    C_灭菌失效期 = 18
    C_当前库存 = 19
    C_对方库存 = 20
    C_填写数量 = 21
    C_实际数量 = 22
    c_原始数量 = 23
    C_采购价 = 24
    C_采购金额 = 25
    C_售价 = 26
    C_售价金额 = 27
    C_差价 = 28
End Enum

Private Const mBillCols  As Integer = 29              '总列数
'=========================================================================================
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strMsg As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    
    '检查药品入出类别是否完整
    strMsg = "没有设置卫材移库的入库及出库类别，请在入出分类中设置！"
    
    gstrSQL = "" & _
        "   SELECT B.Id,B.系数 " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID AND A.单据 = 34"
    
    zlDatabase.OpenRecordset rsDepend, gstrSQL, "卫材移库管理"
        
    With rsDepend
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "系数=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的入库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = "系数=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的出库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    
    Set rsDepend = ReturnSQL(mlngStockID, "卫材移库管理", False, , 1722)
    rsDepend.Filter = "ID<>" & mlngStockID
    With rsDepend
        strMsg = "没有任何库房允许申领，请在[卫材参数设置]的卫材流向中设置！"
        If .RecordCount = 0 Then GoTo ErrHand
    End With
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsDepend.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, Optional lngStockID As Long = 0, Optional int处理方式 As Integer = 0)
    Dim strSQL As String
    Dim rsPara As New ADODB.Recordset
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
    mlngStockID = IIf(lngStockID = 0, glngDeptId, lngStockID)
    mint处理方式 = int处理方式

    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    
    mintUnit = Val(strReg)
    mint库存检查 = Get出库检查(mlngStockID)
    
    mint明确批次 = IIf(IS批次申领, 1, 0)

    If mint明确批次 = 0 Then mint库存检查 = 0
    
    If mint编辑状态 = 1 Or mint编辑状态 = 5 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint编辑状态 = 3 Then
        CmdSave.Caption = "核查(&C)"
        Lbl填制人.Caption = "核查人"
        Lbl填制日期.Caption = "核查日期"
    ElseIf mint编辑状态 = 4 Then
        mblnFirst = True
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 7 Then
        mblnEdit = False
        mblnFirst = True
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint处理方式 = 1 Then
            CmdSave.Caption = "申请冲销(&O)"
            CmdSave.Width = CmdSave.Width + 200
        Else
            CmdSave.Caption = "冲销(&O)"
            CmdSave.Width = cmdCancel.Width
        End If

    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
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
                If MsgBox("如果改变库房，有可能要改变相应卫材的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
        
        mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        If mint明确批次 = 0 Then mint库存检查 = 0
        
    End With
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_实际数量) = Format(0, mFMT.FM_数量)
                .TextMatrix(intRow, mBillCol.C_采购金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_售价金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_差价) = Format(0, mFMT.FM_金额)
            End If
        Next
    End With
    Call 显示合计金额
End Sub


Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_实际数量) = .TextMatrix(intRow, mBillCol.C_填写数量)
                .TextMatrix(intRow, mBillCol.C_采购金额) = Format(.TextMatrix(intRow, mBillCol.C_填写数量) * .TextMatrix(intRow, mBillCol.C_采购价), mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_售价金额) = Format(.TextMatrix(intRow, mBillCol.C_填写数量) * .TextMatrix(intRow, mBillCol.C_售价), mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_差价) = Format(.TextMatrix(intRow, mBillCol.C_售价金额) - .TextMatrix(intRow, mBillCol.C_采购金额), mFMT.FM_金额)
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
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
        
        cmdRequest.Left = txtCode.Left + txtCode.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    Else
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
        cmdRequest.Left = cmdFind.Left + cmdFind.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub



Private Sub cmdRequest_Click()
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim blnDo As Boolean
    Dim str灭菌效期 As String
    Dim bln药房 As Boolean
    Dim dblPrice As Double
    Dim str效期 As String
    Dim dbl数量 As Double
    Dim lng材料ID As Long
    Dim bln分批 As Boolean
    Dim dbl申购数量  As Double
    Dim dbl已导数量 As Double
    
    If mlngStockID = 0 Then  '无移入库房
        MsgBox "移入库房不能为空！", vbInformation, gstrSysName
        Exit Sub
    End If


    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.Text, Val(cboStock.ItemData(cboStock.ListIndex)), mfrmMain.cboStock.Text, mlngStockID)
    If mstrRequestNO <> "" Then
        blnDo = False
        mstrRequestNO = Mid(mstrRequestNO, 1, LenB(StrConv(mstrRequestNO, vbFromUnicode)) - 1)

        bln药房 = True
        gstrSQL = "Select Distinct 0 " & _
                                    "From 部门性质说明 " & _
                                    "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
        If rsTemp.RecordCount = 0 Then
            bln药房 = False
        End If

        gstrSQL = "Select a.Id as 材料id, d.数量 as 计划数量,a.编码,a.名称 ,a.规格,c.现价 as 售价,a.计算单位 as 散装单位,a.是否变价 as 时价,b.包装单位,b.换算系数,b.指导差价率,b.最大效期,b.一次性材料" & vbNewLine & _
                    ",e.上次产地 as 产地,e.上次批号 as 批号,nvl(e.批次,0) as 批次,e.效期,e.灭菌效期,e.可用数量,nvl(e.实际数量,0) as 实际数量,e.实际金额,e.实际差价,e.零售价,e.平均成本价,e.批准文号,b.库房分批,b.在用分批, nvl(b.跟踪病人,0) as 跟踪病人" & vbNewLine & _
                    "From 收费项目目录 A, 材料特性 B, 收费价目 C," & vbNewLine & _
                    "     (Select  b.材料id, Sum(b.计划数量) As 数量" & vbNewLine & _
                    "       From 材料采购计划 A, 材料计划内容 B" & vbNewLine & _
                    "       Where a.Id = b.计划id  and a.单据=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.材料id) D,药品库存 e" & vbNewLine & _
                    "Where a.Id = b.材料id And b.材料id = c.收费细目id And a.Id = d.材料id and b.材料id=e.药品id(+)  and e.库房id=[2] and e.实际数量>0 and e.性质=1 And Sysdate Between c.执行日期 And c.终止日期"

        If gSystem_Para.P156_出库算法 = 0 Then '批次还是效期优先先出库
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.批次, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.效期,Nvl(e.批次, 0)"
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestTransfer_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))

        Do While Not rsTemp.EOF
            With mshBill
                For lngRow = 1 To .Rows - 1
                    If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                        If Val(.TextMatrix(lngRow, 0)) = rsTemp!材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = rsTemp!批次 Then
                            blnDo = True
                            MsgBox "重复材料" & "[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "不再添加！", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                Next

                If Val(.TextMatrix(.Rows - 1, 0)) = 0 Then
                    lngRow = .Rows - 1
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                End If

                str灭菌效期 = IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-MM-dd"))
                If Format(str灭菌效期, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") And Trim(str灭菌效期) <> "" Then
                   If MsgBox("[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫材已经过了灭菌效期,是否还要领用！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If

'                str效期 = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd"))
'                If IsDate(str效期) Then
'                    If Format(str效期, "yyyy-MM-dd") < Format(zldatabase.Currentdate, "yyyy-MM-dd") Then
'                        MsgBox "[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫生材料已经失效了！", vbInformation, gstrSysName
'                    End If
'                End If

                '取售价
                If rsTemp!时价 = 1 Then
                    If rsTemp!在用分批 = 0 Then
                        If rsTemp!库房分批 = 1 And bln药房 = False Then
                            bln分批 = True
                        Else
                            bln分批 = False
                        End If
                    Else
                        bln分批 = True
                    End If

                    If bln分批 = True Then
                        If IsNull(rsTemp!零售价) Then
                            If rsTemp!实际数量 = 0 Then
                                dblPrice = 0
                            Else
                                dblPrice = rsTemp!实际金额 / rsTemp!实际数量
                            End If
                        Else
                            dblPrice = rsTemp!零售价
                        End If
                    Else
                        If rsTemp!实际数量 = 0 Then
                            dblPrice = 0
                        Else
                            dblPrice = rsTemp!实际金额 / rsTemp!实际数量
                        End If
                    End If
                Else
                    dblPrice = IIf(IsNull(rsTemp!售价), 0, rsTemp!售价)
                End If

                If lng材料ID = rsTemp!材料ID Then
                    If rsTemp!可用数量 + dbl已导数量 > rsTemp!计划数量 Then
                        If rsTemp!计划数量 - dbl已导数量 <> 0 Then
                            dbl数量 = rsTemp!计划数量 - dbl已导数量
                            dbl已导数量 = dbl已导数量 + dbl数量
                        Else
                            blnDo = True
                        End If
                    Else
                        dbl数量 = rsTemp!可用数量
                        dbl已导数量 = dbl已导数量 + dbl数量
                    End If
                Else
                    If rsTemp!可用数量 > rsTemp!计划数量 Then
                        dbl数量 = rsTemp!计划数量
                    Else
                        dbl数量 = rsTemp!可用数量
                    End If
                    dbl已导数量 = dbl数量
                End If
                lng材料ID = rsTemp!材料ID

                If dbl数量 = 0 Then
                    blnDo = True
                End If

                '只有不重复的才添加到表格中去
                If blnDo = False Then
                
                    SetRequestColValue lngRow, rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
                    IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                    IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
                    IIf(IsNull(rsTemp!售价), 0, rsTemp!售价), IIf(IsNull(rsTemp!批号), "", rsTemp!批号), _
                    IIf(IsNull(rsTemp!效期), "", rsTemp!效期), _
                    IIf(IsNull(rsTemp!最大效期), "0", rsTemp!最大效期), _
                    IIf(rsTemp!一次性材料 = 1, True, False), _
                    IIf(IsNull(rsTemp!灭菌效期), "", rsTemp!灭菌效期), _
                    rsTemp!库房分批, _
                    IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), _
                    IIf(IsNull(rsTemp!实际金额), "0", rsTemp!实际金额), _
                    IIf(IsNull(rsTemp!实际差价), "0", rsTemp!实际差价), _
                    IIf(IsNull(rsTemp!指导差价率), "0", rsTemp!指导差价率), _
                    IIf(mintUnit = 0, 1, rsTemp!换算系数), IIf(IsNull(rsTemp!批次), 0, rsTemp!批次), rsTemp!时价, rsTemp!在用分批, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                    
                    With mshBill
                        .Row = lngRow
                        .TextMatrix(lngRow, mBillCol.C_行号) = lngRow
                    
                        .TextMatrix(lngRow, mBillCol.C_填写数量) = Format(dbl数量 / IIf(mintUnit = 0, 1, rsTemp!换算系数), mFMT.FM_数量)

                        If .TextMatrix(lngRow, mBillCol.C_售价) <> "" Then
                            .TextMatrix(lngRow, mBillCol.C_售价金额) = Format(.TextMatrix(lngRow, mBillCol.C_售价) * .TextMatrix(lngRow, mBillCol.C_填写数量), mFMT.FM_金额)
                        End If
                        
                        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
                        
                        Call 验证出库差价计算(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mBillCol.c_批次)), _
                            Val(.TextMatrix(lngRow, mBillCol.C_比例系数)), Val(.TextMatrix(lngRow, mBillCol.C_实际差价)), Val(.TextMatrix(lngRow, mBillCol.C_实际金额)), _
                            Val(Split(.TextMatrix(lngRow, mBillCol.C_指导差价率), "||")(0)) / 100, Val(.TextMatrix(lngRow, mBillCol.C_填写数量)), Val(.TextMatrix(lngRow, mBillCol.C_售价金额)), dbl差价, dbl购价, dbl成本金额)
                        
                        .TextMatrix(lngRow, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
                        .TextMatrix(lngRow, mBillCol.C_采购价) = Format(dbl购价, mFMT.FM_成本价)
                        .TextMatrix(lngRow, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)

                        .TextMatrix(lngRow, mBillCol.C_实际数量) = Format(dbl数量 / IIf(mintUnit = 0, 1, rsTemp!换算系数), mFMT.FM_数量)
                    End With

                End If

                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub



Private Sub Form_Activate()
    If mblnFirst = False Then
        If mshBill.Rows > 50 Then
            Call AviShow(Me) '提示用户正在查询数据
        End If
        Call get库存数量    '为当前库存数量和对方库存数量列赋值
        If mshBill.Rows > 50 Then
            Call AviShow(Me, False)
        End If
        Exit Sub
    End If
    
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
    
    mblnFirst = False
    If mint编辑状态 = 5 Then
        
        If Not frmRequestNavigation.ShowNavigation(Me, mlngStockID) = True Then
            Unload Me
            Exit Sub
        End If
        mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        If mint明确批次 = 0 Then mint库存检查 = 0
        mshBill.SetFocus
    End If
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            If mint编辑状态 = 7 Then
                MsgBox "该单据已没有可以冲销的材料，请检查！", vbOKOnly, gstrSysName
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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Function DeleteNo() As Boolean
    '删除单据
    If txtNO.Caption <> "" Then
        gstrSQL = "zl_材料申领_delete('" & txtNO.Caption & "')"
        
        zlDatabase.ExecuteProcedure gstrSQL, "删除单据"
        DeleteNo = True
        Exit Function
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    '设置排序数据集
    Call SetSortRecord
    
    If mint编辑状态 = 3 Then
        '核查
        Call SaveCard
    End If
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If

    If mint编辑状态 = 6 Then       '审核
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        If Not 检查单价(19, txtNO.Tag, False) And Not mblnUpdate Then
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        If SaveCheck() = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    
    If mint编辑状态 = 7 Then '冲销
        If SaveStrike Then Unload Me
        Exit Sub
    End If
    
    
'    If mint编辑状态 = 6 Or mint编辑状态 = 7 Then '接受，更新接受人
'        gstrSQL = "ZL_材料移库_RECEIVE('" & txtNo.Caption & "'," & IIf(mint编辑状态 = 6, "'" & gstrUserName & "'", "NULL") & ")"
'        Call ExecuteProcedure("接受或拒收库房发出的单据")
'        mblnSuccess = True
'        Unload Me
'        Exit Sub
'    End If
    
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
    
'    mstr单据号 = NextNo(72)
    txtNO = ""
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    
    txt摘要.Text = ""
    If cboStock.Enabled Then cboStock.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
End Sub

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng材料ID As Long
    Dim dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsprice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = 19 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(b.现价, " & g_小数位数.obj_散装小数.零售价小数 & ") And" & _
              "    NVL(c.是否变价, 0) = 0" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 19 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & g_小数位数.obj_散装小数.零售价小数 & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 19 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & g_小数位数.obj_散装小数.成本价小数 & ")<>round(b.平均成本价," & g_小数位数.obj_散装小数.成本价小数 & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 材料id, 序号"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取当前价格]", CStr(Me.txtNO.Caption))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
        dbl成本价 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购价))
        dbl零售价 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价))
        dbl成本金额 = dbl成本价 * dbl数量
        dbl零售金额 = dbl零售价 * dbl数量
        dbl差价 = dbl零售金额 - dbl成本金额
'
        If lng材料ID <> 0 Then
            rsprice.Filter = "类型='售价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数)), mFMT.FM_零售价))
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            rsprice.Filter = "类型='成本价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl成本价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数)), mFMT.FM_金额))
                dbl成本金额 = Val(Format(dbl成本价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mBillCol.C_售价) = Format(dbl零售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mBillCol.C_售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(dbl成本价, mFMT.FM_成本价)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Load()
    Dim strStock As String
    Dim rsStock As New Recordset
    Dim strReg As String
    
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mbln申领核查 = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, mlngModule, "0")) = 0, False, True)
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    
    txtNO = mstr单据号
    txtNO.Tag = mstr单据号
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            If mlngStockID <> rsDepend!Id Then
                .AddItem rsDepend!名称
                .ItemData(.NewIndex) = rsDepend!Id
            End If
            rsDepend.MoveNext
        Loop
        .ListIndex = 0
    End With
    mstrTime_Start = GetBillInfo(19, mstr单据号)
    
    Call initCard
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = True, 1000, 0)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_差价) = IIf(mblnCostView = True, 800, 0)
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim strUnitQuantity_Stock As String
    Dim intRow As Integer
    Dim varStuff As Variant
    Dim numUseAbleCount As Double
    Dim lngStockID  As Long
    Dim intCount As Integer
    '库房
    On Error GoTo ErrHandle
   With cboStock
        If Not (mint编辑状态 = 1 Or mint编辑状态 = 5) Then
            '取指定单据的出库库房与入库库房
            gstrSQL = " Select 库房ID,对方部门ID From 药品收发记录" & _
                      " Where NO=[1] And 单据=19 And 入出系数=-1 And Rownum<2"
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, "取指定单据的出库库房与入库库房", mstr单据号)
                      
            If rsInitCard.RecordCount <> 0 Then
                lngStockID = rsInitCard!库房ID
            End If
        End If
        For intCount = 0 To .ListCount - 1
            If .ItemData(intCount) = lngStockID Then
                .ListIndex = intCount: Exit For
            End If
        Next
        mintcboIndex = .ListIndex
    End With
    
    
    Select Case mint编辑状态
        Case 1, 5
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 6, 7
            initGrid

                        
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量,Z.实际金额,Z.实际差价"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.填写数量 / B.换算系数) AS 填写数量,(A.实际数量 / B.换算系数) AS 实际数量,a.成本价*B.换算系数 as 成本价,a.零售价*B.换算系数 as 零售价,B.换算系数 as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量/B.换算系数 As 可用数量,Z.实际金额,Z.实际差价"
            End Select
            
            
            
            If mint编辑状态 <> 7 Then
                gstrSQL = "" & _
                    "   SELECT DISTINCT A.药品ID 材料id,A.序号,('['||D.编码||']'||D.名称) AS 卫材信息," & _
                    "                   B.材料来源,D.规格,D.产地 AS 原产地,A.产地,A.批准文号,A.批号,A.批次,B.指导差价率,B.库房分批 ," & _
                    "                   B.最大效期,A.效期,A.灭菌效期,A.填写数量 as 原始数量," & strUnitQuantity & _
                    "                   A.成本金额,A.零售金额, A.差价, " & strUnitQuantity_Stock & _
                    "                   ,A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.对方部门ID,D.是否变价,B.在用分批 " & _
                    "   FROM 药品收发记录 A, 材料特性 B,收费项目目录 D, " & _
                    "       (   SELECT 药品ID 材料ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "           FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                    " WHERE A.药品ID = B.材料ID AND b.材料ID=D.ID " & _
                    "       AND A.单据 = 19 AND A.入出系数=-1 AND A.NO =[1] AND A.记录状态 =[3]" & _
                    "       AND A.药品ID=Z.材料ID(+) AND NVL(A.批次,0)=Z.批次(+) " & _
                    " ORDER BY A.序号 "
            Else
                gstrSQL = "" & _
                    "   SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                    "   FROM (" & _
                    "           SELECT DISTINCT A.药品ID 材料ID,A.序号,('['||D.编码||']'||D.名称) AS 卫材信息," & _
                    "                   B.材料来源,D.规格,D.产地 AS 原产地,A.产地,A.批准文号, A.批号,A.批次,B.指导差价率,B.库房分批 ," & _
                    "                   B.最大效期,A.效期,A.灭菌效期,A.填写数量 as 原始数量," & strUnitQuantity & _
                    "                   0 成本金额,0 零售金额, 0 差价,A.摘要,A.库房ID,A.对方部门ID,D.是否变价,B.在用分批" & _
                    "           FROM ( " & _
                    "                   SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,0 实际数量,SUM(成本金额) AS 成本金额," & _
                    "                           药品ID,序号,产地,批准文号, 批号,效期,灭菌效期,NVL(批次,0) 批次,扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID" & _
                    "                   FROM 药品收发记录 X " & _
                    "                   WHERE NO=[1] AND 单据=19 AND 入出系数=-1 " & _
                    "                   GROUP BY 药品ID,序号,产地,批准文号,批号,效期,灭菌效期,NVL(批次,0),扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID" & _
                    "                   Having SUM(实际数量)<>0 ) A," & _
                    "               材料特性 B,收费项目目录 D" & _
                    "           WHERE A.药品ID = B.材料ID AND B.材料ID=D.ID ) W," & _
                    "           (   SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "               FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                    "   WHERE W.材料ID=Z.药品ID(+) AND NVL(W.批次,0)=Z.批次(+) " & _
                    "   ORDER BY 序号"
            End If
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, lngStockID, mint记录状态)
               
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            
            
            If mint编辑状态 = 7 Then
                Txt填制人 = gstrUserName
                Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = gstrUserName
                Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            Else
                Txt填制人 = rsInitCard!填制人
                If mint编辑状态 = 2 Then
                    Txt填制人 = gstrUserName
                End If
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End If
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            
            If mint编辑状态 = 2 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    'IntRow = rsInitCard!序号
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_材料) = rsInitCard!卫材信息
                    .TextMatrix(intRow, mBillCol.c_规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mBillCol.C_序号) = zlStr.NVL(rsInitCard!序号)
                    
                    .TextMatrix(intRow, mBillCol.C_产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mBillCol.C_批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mBillCol.c_单位) = rsInitCard!单位
                    .TextMatrix(intRow, mBillCol.c_批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mBillCol.C_效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mBillCol.C_灭菌失效期) = IIf(IsNull(rsInitCard!灭菌效期), "", Format(rsInitCard!灭菌效期, "yyyy-mm-dd"))
                                
                    .TextMatrix(intRow, mBillCol.C_填写数量) = Format(rsInitCard!填写数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mBillCol.C_实际数量) = Format(rsInitCard!实际数量, mFMT.FM_数量)
                                
                    .TextMatrix(intRow, mBillCol.C_采购价) = Format(rsInitCard!成本价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mBillCol.C_采购金额) = Format(rsInitCard!成本金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_售价) = Format(rsInitCard!零售价, mFMT.FM_零售价)
                    .TextMatrix(intRow, mBillCol.C_售价金额) = Format(rsInitCard!零售金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_差价) = Format(rsInitCard!差价, mFMT.FM_金额)
                    
                    .TextMatrix(intRow, mBillCol.C_最大效期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!是否变价 & "||" & rsInitCard!在用分批
                    .TextMatrix(intRow, mBillCol.c_批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mBillCol.C_比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mBillCol.C_指导差价率) = rsInitCard!指导差价率
                    .TextMatrix(intRow, mBillCol.C_库房分批) = IIf(IsNull(rsInitCard!库房分批), "0", rsInitCard!库房分批)
                    .TextMatrix(intRow, mBillCol.C_可用数量) = IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量)
                    .TextMatrix(intRow, mBillCol.C_实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mBillCol.C_实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    .TextMatrix(intRow, mBillCol.c_原始数量) = Val(zlStr.NVL(rsInitCard!原始数量))
                    
                    If mint编辑状态 = 2 Then
                        numUseAbleCount = 0
                        For Each varStuff In mcolUsedCount
                            If varStuff(0) = CStr(rsInitCard!材料ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)) Then
                                numUseAbleCount = varStuff(1)
                                mcolUsedCount.Remove varStuff(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!材料ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!填写数量), "0", rsInitCard!填写数量))), CStr(rsInitCard!材料ID) & CStr(IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次))
                        
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    
    Call get库存数量
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_行号) = ""
        .TextMatrix(0, mBillCol.C_序号) = "序号"
                
        .TextMatrix(0, mBillCol.C_材料) = "名称与编码"
        .TextMatrix(0, mBillCol.c_规格) = "规格"
        .TextMatrix(0, mBillCol.C_产地) = "产地"
        .TextMatrix(0, mBillCol.C_批准文号) = "批准文号"
        .TextMatrix(0, mBillCol.c_单位) = "单位"
        .TextMatrix(0, mBillCol.c_批号) = "批号"
        .TextMatrix(0, mBillCol.C_效期) = "效期"
        .TextMatrix(0, mBillCol.C_灭菌失效期) = "灭菌失效期"
        
        .TextMatrix(0, mBillCol.C_当前库存) = "当前库存"
        .TextMatrix(0, mBillCol.C_对方库存) = "对方库存"
        
        .TextMatrix(0, mBillCol.C_填写数量) = IIf(mint编辑状态 = 7, "数量", "填写数量")
        .TextMatrix(0, mBillCol.C_实际数量) = IIf(mint编辑状态 = 7, "冲销数量", "实际数量")
        .TextMatrix(0, mBillCol.c_原始数量) = "原始数量"
    
        .TextMatrix(0, mBillCol.C_采购价) = "成本价"
        .TextMatrix(0, mBillCol.C_采购金额) = "成本金额"
        .TextMatrix(0, mBillCol.C_售价) = "售价"
        .TextMatrix(0, mBillCol.C_售价金额) = "售价金额"
        .TextMatrix(0, mBillCol.C_差价) = "差价"
        
        .TextMatrix(0, mBillCol.C_可用数量) = "可用数量"
        .TextMatrix(0, mBillCol.C_库房分批) = "库房分批"
        .TextMatrix(0, mBillCol.C_最大效期) = "最大效期"
        .TextMatrix(0, mBillCol.C_实际差价) = "实际差价"
        .TextMatrix(0, mBillCol.C_实际金额) = "实际金额"
        .TextMatrix(0, mBillCol.C_指导差价率) = "指导差价率"
        .TextMatrix(0, mBillCol.C_比例系数) = "比例系数"
        .TextMatrix(0, mBillCol.c_批次) = "批次"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_序号) = 0
        
        .ColWidth(mBillCol.C_行号) = 300
        .ColWidth(mBillCol.C_材料) = 2200
        .ColWidth(mBillCol.c_规格) = 900
        .ColWidth(mBillCol.C_产地) = 800
        .ColWidth(mBillCol.C_批准文号) = 1000
        .ColWidth(mBillCol.c_单位) = 400
        .ColWidth(mBillCol.c_批号) = 800
        .ColWidth(mBillCol.C_效期) = 1000
        .ColWidth(mBillCol.C_灭菌失效期) = 1000
        .ColWidth(mBillCol.C_当前库存) = 1100
        .ColWidth(mBillCol.C_对方库存) = 1100
        .ColWidth(mBillCol.C_填写数量) = 1100
        .ColWidth(mBillCol.C_实际数量) = 1100
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_售价) = 1000
        .ColWidth(mBillCol.C_售价金额) = 900
        .ColWidth(mBillCol.C_差价) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.c_原始数量) = 0
        
        .ColWidth(mBillCol.C_库房分批) = 0
        .ColWidth(mBillCol.C_可用数量) = 0
        .ColWidth(mBillCol.C_最大效期) = 0
        .ColWidth(mBillCol.C_实际差价) = 0
        .ColWidth(mBillCol.C_实际金额) = 0
        .ColWidth(mBillCol.C_指导差价率) = 0
        .ColWidth(mBillCol.C_比例系数) = 0
        .ColWidth(mBillCol.c_批次) = 0
        
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mBillCol.C_序号) = 5
        .ColData(mBillCol.C_行号) = 5
        .ColData(mBillCol.c_规格) = 5
        .ColData(mBillCol.C_产地) = 5
        .ColData(mBillCol.C_批准文号) = 5
        .ColData(mBillCol.c_单位) = 5
        .ColData(mBillCol.c_批号) = 5
        .ColData(mBillCol.C_效期) = 5
        .ColData(mBillCol.C_灭菌失效期) = 5
        .ColData(mBillCol.c_原始数量) = 5
        .ColData(mBillCol.C_当前库存) = 5
        .ColData(mBillCol.C_对方库存) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
            cboStock.Enabled = True
            txt摘要.Enabled = True
            .ColData(mBillCol.C_材料) = 1
            .ColData(mBillCol.C_填写数量) = 4
            .ColData(mBillCol.C_实际数量) = 5
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Or mint编辑状态 = 6 Or mint编辑状态 = 7 Then
            cboStock.Enabled = False
            txt摘要.Enabled = False
            .ColData(mBillCol.C_填写数量) = 5
            .ColData(mBillCol.C_实际数量) = IIf(mint编辑状态 <> 6, 4, 5)
            .ColData(mBillCol.C_材料) = 0
        End If
        
        
        .ColData(mBillCol.C_采购价) = 5
        .ColData(mBillCol.C_采购金额) = 5
        .ColData(mBillCol.C_售价) = 5
        .ColData(mBillCol.C_售价金额) = 5
        .ColData(mBillCol.C_差价) = 5
        
        .ColData(mBillCol.C_库房分批) = 5
        .ColData(mBillCol.C_可用数量) = 5
        .ColData(mBillCol.C_最大效期) = 5
        .ColData(mBillCol.C_实际差价) = 5
        .ColData(mBillCol.C_实际金额) = 5
        .ColData(mBillCol.C_指导差价率) = 5
        .ColData(mBillCol.C_比例系数) = 5
        .ColData(mBillCol.c_批次) = 5
        
        .ColAlignment(mBillCol.C_材料) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_规格) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_产地) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_批准文号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_单位) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_批号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_效期) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_当前库存) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_对方库存) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_填写数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_实际数量) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_采购价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_采购金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_差价) = flexAlignRightCenter
        
        .PrimaryCol = mBillCol.C_材料
        .LocateCol = mBillCol.C_材料
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mBillCol.C_材料) = 0
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
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - cmdCancel.Height - 200
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
    
    With lbl审核人
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
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = cmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
        
    With cmdFind
        .Top = cmdCancel.Top
    End With
    
    With cmdRequest
        .Top = cmdFind.Top
        
        .Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2) '新增和修改才可见
        
    End With
    
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    If mblnChange = False Or mint编辑状态 = 4 Then
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

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mBillCol.C_行号, mshBill.Row)
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
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    mint仅显示有库存物资 = gSystem_Para.para_卫材填单下可用库存 And mint库存检查 = 2
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
        mlngStockID, mlngStockID, IIf(mint库存检查 = 0, False, IIf(mint明确批次 = 0, False, True)), IIf(mint明确批次 = 0, False, True), _
        False, False, (InStr(1, mstrPrivs, "显示对方库存")), , , , , mint仅显示有库存物资, , , mstrPrivs, IIf(mint明确批次 = 0, False, True), False)
    If RecReturn.RecordCount > 0 Then
    
        With mshBill
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                mblnChange = True
                
                If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
                    IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                    IIf(RecReturn!一次性材料 = 1, True, False), _
                    IIf(IsNull(RecReturn!灭菌失效期), "", RecReturn!灭菌失效期), _
                    RecReturn!库房分批, _
                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                    .Row = .Row + 1
                End If
                
                .Col = mBillCol.C_填写数量
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int点击行
            
            If mstr重复卫材 <> "" Then
                MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                mstr重复卫材 = ""
            End If
        
'            If RecReturn.RecordCount = 1 Then
'                mblnChange = True
'
'                SetColValue .Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                    IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
'                    IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
'                    IIf(RecReturn!一次性材料 = 1, True, False), _
'                    IIf(IsNull(RecReturn!灭菌失效期), "", RecReturn!灭菌失效期), _
'                    RecReturn!库房分批, _
'                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
'                .Col = mBillCol.C_填写数量
'            End If
        End With
        RecReturn.Close
    End If
End Sub


Private Sub mshbill_EditChange(curText As String)
    With mshBill
        If .Col <> mBillCol.C_产地 Then
            mshBill.Text = UCase(curText)
            mshBill.SelStart = Len(mshBill.Text)
        End If
    End With
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_填写数量 Or .Col = mBillCol.C_实际数量 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_填写数量, mBillCol.C_实际数量
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
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
        Select Case .Col
            Case mBillCol.C_材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
'                Call 提示库存数
                
            Case mBillCol.c_批号
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 8
            
            Case mBillCol.C_效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mBillCol.c_批号) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mBillCol.c_批号)) And .TextMatrix(.Row, mBillCol.C_最大效期) <> "" Then
                        If Split(.TextMatrix(.Row, mBillCol.C_最大效期), "||")(0) <> 0 Then
                            strxq = .TextMatrix(.Row, mBillCol.c_批号)
                            strxq = TranNumToDate(strxq)
                            If strxq = "" Then Exit Sub
                            
                            .TextMatrix(.Row, mBillCol.C_效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mBillCol.C_最大效期), "||")(0), strxq), "yyyy-mm-dd")
                        End If
                    End If
                End If
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
                
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
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
            
            Case mBillCol.C_材料
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    

                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
                    mint仅显示有库存物资 = gSystem_Para.para_卫材填单下可用库存 And mint库存检查 = 2
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                        mlngStockID, mlngStockID, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, IIf(mint库存检查 = 0, False, IIf(mint明确批次 = 0, False, True)), _
                        IIf(mint明确批次 = 0, False, True), False, False, (InStr(1, mstrPrivs, "显示对方库存")), , , , mint仅显示有库存物资, , , mstrPrivs, IIf(mint明确批次 = 0, False, True), False)
                        
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
                                IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                                IIf(zlStr.NVL(RecReturn!一次性材料, 0) = 1, True, False), _
                                IIf(zlStr.NVL(RecReturn!灭菌失效期) = "", "", Format(RecReturn!灭菌失效期, "yyyy-mm-dd")), _
                                RecReturn!库房分批, _
                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        RecReturn.MoveNext
                    Next
    
                    mshBill.Row = int点击行
                    
                    If mstr重复卫材 <> "" Then
                        MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                        mstr重复卫材 = ""
                    End If
'
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                                IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
'                                IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
'                                IIf(zlStr.NVL(RecReturn!一次性材料, 0) = 1, True, False), _
'                                IIf(zlStr.NVL(RecReturn!灭菌失效期) = "", "", Format(RecReturn!灭菌失效期, "yyyy-mm-dd")), _
'                                RecReturn!库房分批, _
'                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
'                    Call 提示库存数
                End If
            Case mBillCol.c_批号
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.c_批号) = ""
                    End If
                    If .ColData(mBillCol.C_效期) = 2 Then
                        .Col = mBillCol.C_效期
                    Else
                        .Col = mBillCol.C_填写数量
                    End If
                    
                    
                    Cancel = True
                    Exit Sub
                End If
                
                If Len(strKey) < 8 Then
                    MsgBox "批号长度不够，必须为8位,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mBillCol.C_效期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_效期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                If .TextMatrix(.Row, 0) = "" Then .Text = "": .TextMatrix(.Row, mBillCol.C_填写数量) = "": Exit Sub
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
                        MsgBox "数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) < 0.001 Then
                        MsgBox "数量必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
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
                    
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '成本价的公式：     出库金额=数量*售价
                    '                  出库差价=出库金额*（实际差价/实际金额）
                    '                  if 实际金额=0 then  出库差价=出库金额*指导差价率
                    '                  购价（成本价）=（出库金额-出库差价）/数量
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mBillCol.C_售价) <> "" Then
                        .TextMatrix(.Row, mBillCol.C_售价金额) = Format(.TextMatrix(.Row, mBillCol.C_售价) * strKey, mFMT.FM_金额)
                    End If
                    
                    If mint编辑状态 <> 7 Then
                        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
                        'cboStock.ItemData(cboStock.ListIndex)
                        
                        Call 验证出库差价计算(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)), _
                            Val(.TextMatrix(.Row, mBillCol.C_比例系数)), Val(.TextMatrix(.Row, mBillCol.C_实际差价)), Val(.TextMatrix(.Row, mBillCol.C_实际金额)), _
                            Val(Split(.TextMatrix(.Row, mBillCol.C_指导差价率), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mBillCol.C_售价金额)), dbl差价, dbl购价, dbl成本金额)
                        .TextMatrix(.Row, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
                        .TextMatrix(.Row, mBillCol.C_采购价) = Format(dbl购价, mFMT.FM_成本价)
                        .TextMatrix(.Row, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
                    Else
                        .TextMatrix(.Row, mBillCol.C_采购金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_采购价)) * strKey, mFMT.FM_金额)
                        .TextMatrix(.Row, mBillCol.C_差价) = Format(Val(.TextMatrix(.Row, mBillCol.C_售价金额)) - Val(.TextMatrix(.Row, mBillCol.C_采购金额)), mFMT.FM_金额)
                    End If
                    
                    If .Col = mBillCol.C_填写数量 Then
                        .TextMatrix(.Row, mBillCol.C_实际数量) = strKey
                    End If
                End If
                显示合计金额
            
        End Select
    End With
End Sub

'从材料目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
    ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal bln一次性材料 As Boolean, _
    ByVal str灭菌失效期 As String, ByVal int库房分批 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsprice As New Recordset
    Dim bln分批 As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If bln一次性材料 = True Then
        If Format(str灭菌失效期, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str灭菌失效期) <> "" Then
           If MsgBox("卫材【" & str材料 & "(" & lng批次 & ")】已经过了灭菌失效期,是否还要申领！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                Exit Function
           End If
        End If
    End If
    
    SetColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If int是否变价 = 1 Then
            If int在用分批 = 0 Then
                If int库房分批 = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From 部门性质说明 " & _
                            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln分批 = True
                    End If
                End If
            Else
                bln分批 = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "       and 药品id=[2]" & _
                "       and 性质=1 and 实际数量>0 and " & _
                "       nvl(批次,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
            If rsprice.EOF Then
                If mint明确批次 = 1 Then
                    MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num售价 * num比例系数
                End If
            Else
                If bln分批 = True Then
                    dblPrice = rsprice!分批售价
                Else
                    dblPrice = rsprice!平均零售价
                End If
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_库房分批) = int库房分批
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_最大效期) = int最大效期 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.C_实际差价) = num实际差价
        .TextMatrix(intRow, mBillCol.C_实际金额) = num实际金额
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率
        .TextMatrix(intRow, mBillCol.C_比例系数) = num比例系数
        If mint明确批次 = 1 Then
            .TextMatrix(intRow, mBillCol.c_批次) = lng批次
            .TextMatrix(intRow, mBillCol.c_批号) = str批号
            .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
            .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
        Else
            .TextMatrix(intRow, mBillCol.c_批次) = lng批次 '手工加载卫材，批次传入为0；按申购单提取会传具体批次
            .TextMatrix(intRow, mBillCol.c_批号) = ""
            .TextMatrix(intRow, mBillCol.C_效期) = ""
            .TextMatrix(intRow, mBillCol.C_灭菌失效期) = ""
        End If
        '需要考虑时价分批和不分批情况
        If int是否变价 = 1 Then .TextMatrix(intRow, mBillCol.C_售价) = Format(dblPrice, mFMT.FM_零售价)
        Call CheckLapse(str效期)
        
        Call get库存数量(intRow)
    End With
'    Call 提示库存数
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'从材料目录中取值并附给相应的列
Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
    ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal bln一次性材料 As Boolean, _
    ByVal str灭菌失效期 As String, ByVal int库房分批 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsprice As New Recordset
    Dim bln分批 As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If bln一次性材料 = True Then
        If Format(str灭菌失效期, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str灭菌失效期) <> "" Then
           If MsgBox("卫材【" & str材料 & "(" & lng批次 & ")】已经过了灭菌失效期,是否还要申领！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                Exit Function
           End If
        End If
    End If
    
    SetRequestColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If int是否变价 = 1 Then
            If int在用分批 = 0 Then
                If int库房分批 = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From 部门性质说明 " & _
                            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln分批 = True
                    End If
                End If
            Else
                bln分批 = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "       and 药品id=[2]" & _
                "       and 性质=1 and 实际数量>0 and " & _
                "       nvl(批次,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
            If rsprice.EOF Then
                If mint明确批次 = 1 Then
                    MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num售价 * num比例系数
                End If
            Else
                If bln分批 = True Then
                    dblPrice = rsprice!分批售价
                Else
                    dblPrice = rsprice!平均零售价
                End If
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_库房分批) = int库房分批
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_最大效期) = int最大效期 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.C_实际差价) = num实际差价
        .TextMatrix(intRow, mBillCol.C_实际金额) = num实际金额
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率
        .TextMatrix(intRow, mBillCol.C_比例系数) = num比例系数
        '按申购单申领是明确批次的，不用判断是否按批次申领
        .TextMatrix(intRow, mBillCol.c_批次) = lng批次
        .TextMatrix(intRow, mBillCol.c_批号) = str批号
        .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
       
        '需要考虑时价分批和不分批情况
        If int是否变价 = 1 Then .TextMatrix(intRow, mBillCol.C_售价) = Format(dblPrice, mFMT.FM_零售价)
        Call CheckLapse(str效期)
        
        Call get库存数量(intRow)
    End With
'    Call 提示库存数
    SetRequestColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
  
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Val(.TextMatrix(intLop, 0)) > 0 And Trim(.TextMatrix(intLop, mBillCol.C_材料)) = "" Then
                    MsgBox "第" & intLop & "行卫材的名称为空了，请检查！", vbInformation, gstrSysName
                    mshBill.SetFocus
                    .Row = intLop
                    .MsfObj.TopRow = intLop
                    .Col = mBillCol.C_材料
                    Exit Function
                End If
                
                If Trim(.TextMatrix(intLop, mBillCol.C_材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_填写数量))) = "" Then
                        MsgBox "第" & intLop & "行卫材的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_填写数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的填写数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_实际数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的实际数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_实际数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
                    End If
                               
                    '新增/修改时检查可用数量，防止并发
                    If Not CompareUsableQuantity(intLop, Val(Trim(.TextMatrix(intLop, mBillCol.C_填写数量))), True) Then
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
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
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lng部门ID As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次 As Long
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl填写数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl实际数量 As Double
    Dim dbl差价 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str核查日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str灭菌效期 As String
    Dim str核查人 As String
    Dim n As Long
    
    Dim intRow As Integer
    Dim arrSQL As Variant
    
    '自动分解申领记录时使用
    Dim blnAuto As Boolean              '是否需要自动分解
    Dim dbl填写数量_Cur As Double
    Dim rsStock As New ADODB.Recordset
    
    SaveCard = False
    arrSQL = Array()
    
    With mshBill
        chrNo = Trim(txtNO)
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        If chrNo <> "" Then
            If CheckNOExists(72, chrNo) Then Exit Function
        End If
        
        If chrNo = "" Then chrNo = sys.GetNextNo(72, lng库房ID)
        If IsNull(chrNo) Then Exit Function
        txtNO.Tag = chrNo
        
        lng部门ID = mlngStockID
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mbln申领核查 = True And mint编辑状态 = 3 Then
            str填制日期 = Txt填制日期
        End If
        str核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str审核人 = Txt审核人
        
        If mbln申领核查 = True And mint编辑状态 = 3 Then
            str核查人 = Txt填制人
        End If
        
        On Error GoTo ErrHandle
        
        If mint编辑状态 = 2 Or mint编辑状态 = 3 Then      '修改和核查
            gstrSQL = "zl_材料申领_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & vbCrLf & gstrSQL
        End If
        
        Dim intTmp As Integer
        lng序号 = -1
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                '如果当前批次卫材不够，自动取其它批次的卫材，产生多笔申领记录
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mBillCol.C_产地)
                str批号 = .TextMatrix(intRow, mBillCol.c_批号)
                intTmp = Val(.TextMatrix(intRow, mBillCol.C_比例系数))
                lng批次 = Val(.TextMatrix(intRow, mBillCol.c_批次))
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "", .TextMatrix(intRow, mBillCol.C_效期))
                str灭菌效期 = IIf(.TextMatrix(intRow, mBillCol.C_灭菌失效期) = "", "", .TextMatrix(intRow, mBillCol.C_灭菌失效期))
                dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * intTmp, g_小数位数.obj_最大小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * intTmp, g_小数位数.obj_最大小数.数量小数)
                dbl成本价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / IIf(intTmp = 0, 1, intTmp), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零售价 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价)) / IIf(intTmp = 0, 1, intTmp), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mBillCol.C_差价)), g_小数位数.obj_最大小数.金额小数)
                lng序号 = lng序号 + 2  '求奇数：公式为：2n-1;出库序号为偶数
                'zl_材料移库_INSERT( /*NO_IN*/, /*序号_IN*/, /*库房ID_IN*/,
                '/*对方部门ID_IN*/, /*材料ID_IN*/, /*批次_IN*/, /*填写数量_IN*/实际数量/,
                '/*成本价_IN*/, /*成本金额_IN*/, /*零售价_IN*/, /*零售金额_IN*/,
                '/*差价_IN*/, /*填制人_IN*/, /*产地_IN*/, /*批号_IN*/, /*效期_IN*/,
                '/*摘要_IN*/填制日期_in );
                gstrSQL = "zl_材料申领_INSERT('" & _
                    chrNo & "'," & _
                    lng序号 & "," & _
                    lng库房ID & "," & _
                    lng部门ID & "," & _
                    lng材料ID & "," & _
                    lng批次 & "," & _
                    dbl填写数量 & "," & _
                    dbl实际数量 & "," & _
                    dbl成本价 & "," & _
                    dbl成本金额 & "," & _
                    dbl零售价 & "," & _
                    dbl零售金额 & "," & _
                    dbl差价 & ",'" & _
                    str填制人 & "','" & _
                    str产地 & "','" & _
                    str批号 & "'," & _
                    IIf(str效期 = "", "Null", "to_date('" & str效期 & "','yyyy-mm-dd')") & "," & _
                    IIf(str灭菌效期 = "", "Null", "to_date('" & str灭菌效期 & "','yyyy-mm-dd')") & ",'" & _
                    str摘要 & "',to_date('" & _
                    str填制日期 & "','yyyy-mm-dd HH24:MI:SS')," & _
                    IIf(str核查人 <> "", "'" & str核查人 & "'", "Null") & "," & _
                    IIf(str核查人 <> "", "to_date('" & str核查日期 & "','yyyy-mm-dd')", "null") & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
        If Not ExecuteSql(arrSQL, mstrCaption) Then Exit Function
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mBillCol.C_售价金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim blnIs显示对方库存 As Boolean
    Dim str对方库存数 As String
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        
        '发出库存的当前卫材的可用数量
        If mint明确批次 = 1 Then
            gstrSQL = " Select 可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as 可用数量 from 药品库存 " & _
                      " Where 库房id=[1]" & _
                      " And 药品id=[2] And 性质=1 " & _
                      " And Nvl(批次,0)=[3]"
        Else
            gstrSQL = " Select Sum(可用数量)/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as 可用数量 from 药品库存 " & _
                      " Where 库房id=[1]" & _
                      " And 药品id=[2] And 性质=1 "
        End If
        
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "发出库房可用数量", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)))
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mBillCol.C_可用数量) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
        rsUseCount.Close
        
        '当前发料部门的可用数量,申领库房数量始终为该库房所有批次库存
'        If mint明确批次 = 1 Then
'            gstrSQL = " Select Sum(可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & ") as 可用数量 from 药品库存 where 库房id=[1]" & _
'                      " And 药品id=[2] And 性质=1 " & _
'                      " And nvl(批次,0)=[3]"
'        Else
            gstrSQL = " Select Sum(可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & ") as 可用数量 from 药品库存 where 库房id=[1]" & _
                      " And 药品id=[2] And 性质=1 "
'        End If
        
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "当前在用的可用数量", mlngStockID, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)))
        
        If rsUseCount.EOF Then
            dblStock = 0
        Else
            dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
'        stbThis.Panels(2).Text = "该卫材当前库存数为[" & Format(dblStock, mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.C_单位)
    
        
        blnIs显示对方库存 = zlStr.IsHavePrivs(mstrPrivs, "显示对方库存")
        str对方库存数 = "；" & Me.cboStock.Text & "库存数为[" & Format(.TextMatrix(.Row, mBillCol.C_可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.c_单位)
        
        stbThis.Panels(2).Text = "该卫材" & mfrmMain.cboStock.Text & "库存数为[" & Format(dblStock, mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.c_单位) _
            & IIf(blnIs显示对方库存, str对方库存数, "")
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get库存数量(Optional ByVal intRow As Integer = 0)
'''''''''''''''''''''''''''''''''''''
'获取库存数量的方法
'''''''''''''''''''''''''''''''''''''
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim blnIs显示对方库存 As Boolean
    Dim intStart As Integer, intEnd As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    blnIs显示对方库存 = zlStr.IsHavePrivs(mstrPrivs, "显示对方库存")
    
    If intRow > 0 Then
        intStart = intRow
        intEnd = intRow
    Else
        intStart = 1
        intEnd = mshBill.Rows - 1
    End If
    
    With mshBill
        For i = intStart To intEnd
            If .TextMatrix(i, 0) = "" Then Exit Sub

            If blnIs显示对方库存 Then
                If Val(.TextMatrix(i, c_批次)) > 0 Then
                    gstrSQL = " Select Nvl(可用数量,0)/" & .TextMatrix(i, C_比例系数) & " as 可用数量, Nvl(实际数量,0)/" & .TextMatrix(i, C_比例系数) & " as 实际数量 from 药品库存 " & _
                              " Where 库房id=[1] " & _
                              " And 药品id=[2] And 性质=1 " & _
                              " And Nvl(批次,0)=[3] "
                Else
                    If Get分批属性(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(i, 0))) = 1 Then
                        '如果出库库房是分批，则统计所有批次的合计数量
                        gstrSQL = " Select Sum(Nvl(可用数量,0))/" & .TextMatrix(i, C_比例系数) & " as 可用数量, Sum(Nvl(实际数量,0))/" & .TextMatrix(i, C_比例系数) & " as 实际数量 from 药品库存 " & _
                              " Where 库房id=[1] " & _
                              " And 药品id=[2] And 性质=1 And Nvl(批次,0)>0 "
                    Else
                        '如果出库库房是不分批的，则统计总的数量
                        gstrSQL = " Select Sum(Nvl(可用数量,0))/" & .TextMatrix(i, C_比例系数) & " as 可用数量, Sum(Nvl(实际数量,0))/" & .TextMatrix(i, C_比例系数) & " as 实际数量 from 药品库存 " & _
                              " Where 库房id=[1] " & _
                              " And 药品id=[2] And 性质=1 "
                    End If
                End If
                Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[发出库房数量]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, c_批次)))
                
                If rsUseCount.EOF Then
                    dblStock = 0
                Else
                    If mint编辑状态 = 6 Then
                        '接收(审核)时显示实际数量
                        dblStock = NVL(rsUseCount!实际数量, 0)
                    Else
                        '其他状态时显示可用数量
                        dblStock = NVL(rsUseCount!可用数量, 0)
                    End If
                End If
                .TextMatrix(i, C_对方库存) = Format(dblStock, mFMT.FM_数量)
                rsUseCount.Close
            End If
                
            '发料部门始终显示所有数量
            gstrSQL = " Select Sum(Nvl(可用数量,0))/" & .TextMatrix(i, C_比例系数) & " as 可用数量, Sum(Nvl(实际数量,0))/" & .TextMatrix(i, C_比例系数) & " as 实际数量 from 药品库存 where 库房id=[1] " & _
                      " And 药品id=[2] And 性质=1 "
            Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[申领部门数量]", mlngStockID, Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, c_批次)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                If mint编辑状态 = 6 Then
                    '接收(审核)时显示实际数量
                    dblStock = NVL(rsUseCount!实际数量, 0)
                Else
                    '其他状态时显示可用数量
                    dblStock = NVL(rsUseCount!可用数量, 0)
                End If
            End If
            .TextMatrix(i, C_当前库存) = Format(dblStock, mFMT.FM_数量)
       Next
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

'转换数值为日期
Private Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 2000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    
End Function

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double, Optional ByVal blnSave As Boolean = False) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim numUsedCount As Double
    Dim varStuff As Variant
    Dim rsCheck As ADODB.Recordset
    Dim strSaveCheck As String
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '只要是分批卫材，允许输入比当前批次大的数量，程序自动分解，而仅仅是时价卫材属性的不允许
    CompareUsableQuantity = False
    If mint明确批次 = 0 Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If Not blnSave Then
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_可用数量), mFMT.FM_数量)
        Else
            '新增，修改保存时重新取数据库中的可用数量，主要防止并发或多人同时填单对可用数量取值的影响
            gstrSQL = "Select Nvl(可用数量, 0) 可用数量 From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3] "
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CompareUsableQuantity", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_批次)))
            
            If rsCheck.EOF Then
                dblUsableQuantity = 0
            Else
                dblUsableQuantity = Val(Format(rsCheck!可用数量 / Val(.TextMatrix(intRow, mBillCol.C_比例系数)), mFMT.FM_数量))
                
                If dblUsableQuantity <> Val(Format(.TextMatrix(intRow, mBillCol.C_可用数量), mFMT.FM_数量)) Then
                    .TextMatrix(intRow, mBillCol.C_可用数量) = dblUsableQuantity
                End If
                
                strSaveCheck = "，可能是其他操作员填单占用了"
            End If
        End If
        
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If mint编辑状态 = 1 Or mint编辑状态 = 5 Then
                If dbl填写数量 > dblUsableQuantity Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”" & strSaveCheck & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                numUsedCount = 0
                For Each varStuff In mcolUsedCount
                    If varStuff(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_批次) Then
                        numUsedCount = varStuff(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount & "”" & strSaveCheck & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Or mint编辑状态 = 5 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”" & strSaveCheck & "，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                numUsedCount = 0
                For Each varStuff In mcolUsedCount
                    If varStuff(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_批次) Then
                        numUsedCount = varStuff(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount & "”" & strSaveCheck & "，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ExecuteSql(ByRef arrSQL As Variant, strTitle As String, Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer

    ExecuteSql = False
    If UBound(arrSQL) >= 0 Then
        '对SQL序列材料ID升序排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '执行SQL语句
        On Error GoTo errH
        If Not bln强制保存 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(Split(arrSQL(i), ";")(1)), mstrCaption
                        
'            Call SQLTest(App.ProductName, strTitle, CStr(Split(arrSql(i), ";")(1)))
'            Debug.Print CStr(Split(arrSql(i), ";")(1))
'            gcnOracle.Execute CStr(Split(arrSql(i), ";")(1)), , adCmdStoredProc
'            Call SQLTest
        Next
        If Not bln强制保存 Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
errH:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'打印单据
Private Sub printbill()
    
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1722", mint记录状态, mintUnit, 1722, "卫材申领单", strNo
End Sub


Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng库房ID As Long
    Dim lng对方部门id As Long
    Dim str审核人 As String
    
    Dim lng材料ID As Long
    Dim str产地 As String
    Dim lng出批次 As Long
    Dim dbl填写数量 As Double
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim lng出类别id As Long
    Dim lng入类别id As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str灭菌效期 As String
    Dim str审核日期 As String
    Dim int序列号 As Integer
    Dim n As Long
    
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    arrSQL = Array()
    mblnSave = False
    SaveCheck = False
    
    '检查该单据是否在进入编辑界面后，被其他操作员修改
    mstrTime_End = GetBillInfo(19, mstr单据号)
    If mstrTime_End = "" Then
        MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrTime_End > mstrTime_Start Then
        MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    '检查该单据是否被正常发送
    gstrSQL = " Select 配药日期 From 药品收发记录 " & _
            " Where 单据=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查该单据是否被正常发送", Me.txtNO.Tag)
    
    If IsNull(rsTemp!配药日期) Then
        MsgBox "该单据被其他操作员取消发送，不允许接收！", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = mlngStockID
    str审核人 = gstrUserName
    strNo = txtNO.Tag
    
    
    gstrSQL = "" & _
        "   SELECT b.系数,b.id AS 类别id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID AND a.单据 = 34 "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "材料移库管理")
    
    If rsTemp.EOF Then
        MsgBox "卫材入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.RecordCount < 2 Then
        MsgBox "卫材入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!系数 = 1 Then
            lng入类别id = rsTemp!类别ID
        Else
            lng出类别id = rsTemp!类别ID
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    str审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo ErrHandle
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mBillCol.C_产地)
                lng出批次 = .TextMatrix(intRow, mBillCol.c_批次)
                dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.数量小数)
                dbl成本价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_散装小数.金额小数)
                dbl售价 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价)) / .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价金额)), g_小数位数.obj_散装小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mBillCol.C_差价)), g_小数位数.obj_散装小数.金额小数)
                str批号 = .TextMatrix(intRow, mBillCol.c_批号)
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_效期) & "','yyyy-mm-dd')")
                str灭菌效期 = IIf(.TextMatrix(intRow, mBillCol.C_灭菌失效期) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_灭菌失效期) & "','yyyy-mm-dd')")

                int序列号 = Val(.TextMatrix(intRow, mBillCol.C_序号))
                
                'zl_材料移库_VERIFY( /*库房ID_IN*/, /*对方部门ID_IN*/, /*药品ID_IN*/,
                    '产地_IN*/, /*出批次_IN*/, /*填写数量_IN*/, /*实际数量_IN*/, /*成本价_IN*/,
                    '/*成本金额_IN*/, /*零售金额_IN*/, /*差价_IN*/, /*出类别ID_IN*/, /*入类别ID_IN*/,
                    '/*NO_IN*/, /*审核人_IN*/, /*批号_IN*/, /*效期_IN*/灭菌效期_IN );
                        
                gstrSQL = "zl_材料移库_Verify(" & int序列号 & "," & lng库房ID & "," & lng对方部门id & "," & _
                     lng材料ID & ",'" & str产地 & "'," & lng出批次 & "," & dbl填写数量 & "," & _
                     dbl实际数量 & "," & dbl成本价 & "," & dbl成本金额 & "," & dbl零售金额 & "," & _
                     dbl差价 & "," & lng出类别id & "," & lng入类别id & ",'" & _
                     strNo & "','" & str审核人 & "','" & str批号 & "'," & str效期 & "," & str灭菌效期 & ",to_date('" & str审核日期 & "','yyyy-mm-dd HH24:MI:SS')" & _
                    ",1," & dbl售价 & " )"
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    gcnOracle.BeginTrans
    If Not ExecuteSql(arrSQL, mstrCaption, True) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
'    If Not 检查单价(19, txtNo.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Function SaveStrike() As Boolean
    Dim 行次_IN As Integer
    Dim 原记录状态_IN As Integer
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 材料ID_IN As Long
    Dim 冲销数量_IN As Double
    Dim 填制人_IN As String
    Dim 填制日期_IN  As String
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim n As Long
    
    SaveStrike = False
    
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mBillCol.C_填写数量)), Val(.TextMatrix(intRow, mBillCol.C_实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        NO_IN = Trim(txtNO.Tag)
        填制人_IN = gstrUserName
        填制日期_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        
        err = 0: On Error GoTo ErrHandle
        
        gcnOracle.BeginTrans
        
        行次_IN = 0
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                材料ID_IN = .TextMatrix(intRow, 0)
                冲销数量_IN = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * Val(.TextMatrix(intRow, mBillCol.C_比例系数)), g_小数位数.obj_散装小数.数量小数)
                If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                End If
                
                序号_IN = .TextMatrix(intRow, mBillCol.C_序号)
                
                'ZL_材料移库_STRIKE(/*行次_IN*/,/*原记录状态_IN*/,/*NO_IN*/,/*序号_IN*/, /*材料ID_IN*/,
                '/*冲销数量_IN*/,/*填制人_IN*/, /*填制日期_IN*/);
                gstrSQL = "" & _
                    "   ZL_材料移库_STRIKE(" & _
                            行次_IN & "," & _
                            原记录状态_IN & ",'" & _
                            NO_IN & "'," & _
                            序号_IN & "," & _
                            材料ID_IN & "," & _
                            冲销数量_IN & ",'" & _
                            填制人_IN & "',to_date('" & _
                            Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & _
                            mint处理方式 & ")"
                zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                
            End If
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行材料来冲销，不能冲销，请检查！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mBillCol.C_序号)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.C_序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mBillCol.c_批次))
                
                .Update
            End If
        Next
        
    End With
End Sub

