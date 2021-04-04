VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmNewCheckCourseCard 
   Caption         =   "药品盘点记录单"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15075
   Icon            =   "frmNewCheckCourseCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15075
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   6045
      Left            =   0
      ScaleHeight     =   5985
      ScaleWidth      =   15015
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   15075
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
         Top             =   960
         Width           =   11235
         _cx             =   19817
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
         FormatString    =   $"frmNewCheckCourseCard.frx":06EA
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
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
      Begin MSMask.MaskEdBox TxtCheckDate 
         Height          =   315
         Left            =   9480
         TabIndex        =   21
         Top             =   593
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Txt修改日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7920
         TabIndex        =   26
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt修改人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5880
         TabIndex        =   25
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label lbl修改人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改人"
         Height          =   180
         Left            =   5280
         TabIndex        =   24
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label lbl修改日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改日期"
         Height          =   180
         Left            =   7140
         TabIndex        =   23
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   12240
         TabIndex        =   20
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   10365
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   660
         Width           =   720
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品盘点记录单"
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
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   14
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
         Left            =   9480
         TabIndex        =   13
         Top             =   195
         Width           =   480
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   12
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   11
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   10
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   13050
         TabIndex        =   9
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10950
         TabIndex        =   8
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额差合计："
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   8640
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
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   9960
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   8070
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmNewCheckCourseCard.frx":075F
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20241
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmNewCheckCourseCard.frx":0FF3
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmNewCheckCourseCard.frx":14F5
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Bindings        =   "frmNewCheckCourseCard.frx":19F7
      Left            =   1320
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNewCheckCourseCard.frx":1A0B
   End
End
Attribute VB_Name = "frmNewCheckCourseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mcon按库存提取 As Integer = 101
Private Const mcon按分类货位提取 As Integer = 102
Private Const mcon确定 As Integer = 103
Private Const mcon帮助 As Integer = 104
Private Const mcon取消 As Integer = 105
Private Const mcon查找 As Integer = 106
Private Const mconFind As Integer = 110
Private Const mcon保存 As Integer = 111
Private Const mcon保存退出 As Integer = 112

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


Private mblnLoad As Boolean
Private mblnAutoLoadData As Boolean '是否自动加载完数据

Private mintSelectStock As Integer           '是否可选库房
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
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
Private mintDefault As Integer              '缺省单位
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Dim mstrPrivs As String                     '权限
Private mblnNoStock As Boolean              '本地参数：是否允许盘点没有设置存储库房的药品
Private mblnLoadData As Boolean             '用于检查是否已装入数据（对于已存在单据）
Private mlngCurrRow As Long
Private mbln忽略服务对象 As Boolean         '为真时忽略药品的服务对象
Private mrsTemp As ADODB.Recordset
Private mbln盘停用药品 As Boolean
Private mstr货位 As String                  '用来保存所选择的货位
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private Const MStrCaption As String = "药品盘点记录单"
Private mbln即时保存 As Boolean         '编辑后保存不退出编辑界面，还可以继续编辑保存
Private mstrLast盘点时间 As String      '记录上次盘点时间，判断是否需要重新加载记录集

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mlngFindFirst As Long
Private mlngFind As Long                            '用于查找
Private mrsFindName As ADODB.Recordset              '用于查找

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

'从参数表中取药品价格、数量、金额小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintCostDigit As Integer           '成本价小数位数
Private mintNumberDigit0 As Integer         '数量小数位数-大单位
Private mintNumberDigit1 As Integer         '数量小数位数-小单位
Private mintMoneyDigit As Integer           '金额小数位数

Private mstr单位 As String
Private mbln相同单位 As Boolean             '大小包装相同，界面只显示一个包装单位

Private mblnNotTrigger As Boolean
Private mblnBatch As Boolean

Private mlng生产商长度 As Long                 '生产商字段长度
Private mlng原产地长度 As Long                 '原产地字段长度

Private Type Type_药品id
    str药品ID As String
    int退出 As Integer
End Type

Private SQLCondition As Type_药品id

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
Private Const mconIntcol加成率 As Integer = 11
Private Const mconIntCol实际差价 As Integer = 12
Private Const mconIntCol实际金额 As Integer = 13
Private Const mconIntCol产地 As Integer = 14
Private Const mconIntCol原产地 As Integer = 15
Private Const mconIntCol库房货位 As Integer = 16
Private Const mconIntCol单位 As Integer = 17
Private Const mconIntCol批号 As Integer = 18
Private Const mconIntCol效期 As Integer = 19
Private Const mconIntCol批准文号 As Integer = 20
Private Const mconintCol成本价 As Integer = 21
Private Const mconIntCol售价 As Integer = 22
Private Const mconintCol帐面数量 As Integer = 23
Private Const mconIntCol大单位数量 As Integer = 24
Private Const mconintCol大单位 As Integer = 25
Private Const mconIntCol小单位数量 As Integer = 26
Private Const mconintCol小单位 As Integer = 27
Private Const mconintCol数量_合计 As Integer = 28
Private Const mconintCol单位_合计 As Integer = 29
Private Const mconintCol标志 As Integer = 30
Private Const mconintCol数量差 As Integer = 31
Private Const mconintCol金额差 As Integer = 32
Private Const mconintCol差价差 As Integer = 33
Private Const mconintCol盘点金额 As Integer = 34
Private Const mconIntCol药品编码和名称 As Integer = 35
Private Const mconIntCol药品编码 As Integer = 36
Private Const mconIntCol药品名称 As Integer = 37
Private Const mconIntColS  As Integer = 38            '总列数
'=========================================================================================







Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not mblnAutoLoadData Then Exit Sub  '数据没加载完不做其他处理
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
        Case mcon按分类货位提取
            cbsClass
        Case mcon确定, mcon保存, mcon保存退出
            cbsSave Control.id
        Case mcon取消
            cbsCancel
    End Select
End Sub



Private Sub cbsClass()
    Dim lngValue As Long
    Dim intCol As Integer
    
    lngValue = MsgBox("提取分类记录，界面已有数据将清除，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If lngValue = vbYes Then
        mblnAutoLoadData = False
        
        frmCheckClass.ShowME Me, Val(txtStock.Tag), mstr货位, SQLCondition.str药品ID, SQLCondition.int退出
        If SQLCondition.int退出 = 1 Then    '1-选择了条件，0-没有选择条件 退出不执行刷新操作
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
            Call GetBatchRec
        End If
    End If
    vsfBill_EnterCell
    
    mblnAutoLoadData = True
End Sub

Private Sub cbsBatch()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim lngRows As Long
    Dim strCompare As String
    Dim strOrder As String
    
    If MsgBox("提取当前库存记录，界面已有数据将清除，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    mblnAutoLoadData = False
    
    vsfBill.rows = 2
    vsfBill.Row = 1
    
    gstrSQL = "Select b.编码 As 药品编码, a.药品id, b.名称, Nvl(a.批次, 0) As 批次, a.批准文号, c.库房货位" & vbNewLine & _
                "From 药品库存 A, 收费项目目录 B, 药品储备限额 C" & vbNewLine & _
                "Where a.性质 = 1 And a.药品id = b.Id And b.Id = c.药品id(+) And a.库房id = [1] And c.库房id = [1] " & _
                " And (Nvl(A.实际数量,0)<>0 Or Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0 )"

    If mbln忽略服务对象 = False Then
        gstrSQL = gstrSQL & _
            " and (Decode(B.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3)) " & _
                " or Decode(B.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]))"
    End If
          
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
    strCompare = Mid(strOrder, 1, 1)
    gstrSQL = gstrSQL & " Order by " & IIf(strCompare = "0", "b.编码", IIf(strCompare = "1", "b.编码", IIf(strCompare = "2", "b.名称", "c.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",b.编码"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "GetBatchRec", Val(txtStock.Tag))
    
    If rsData.RecordCount = 0 Then
        Exit Sub
    End If
    rsData.MoveFirst
    With rsData
        If .RecordCount = 0 Then Exit Sub
        
        mblnBatch = True
        
        Do While Not .EOF
            Call SetPhiscRows(!药品id, !批次, Nvl(!批准文号, ""), True)

            DoEvents
            Call zlControl.StaShowPercent(.AbsolutePosition / .RecordCount, staThis.Panels(2), frmNewCheckCourseCard)
            DoEvents

            If Trim(vsfBill.TextMatrix(vsfBill.Row, 0)) <> "" Then
                vsfBill.rows = vsfBill.rows + 1
                vsfBill.Row = vsfBill.rows - 1
            End If

            .MoveNext
        Loop
        
'        lngRows = .RecordCount
'
'        vsfBill.rows = lngRows + 1
'
'        For lngRow = 1 To lngRows
'            vsfBill.Row = lngRow
'            Call SetPhiscRows(!药品id, !批次, Nvl(!批准文号, ""), True)
'
'            DoEvents
'
'            Call zlControl.StaShowPercent(lngRow / lngRows, staThis.Panels(2), frmNewCheckCourseCard)
'            DoEvents
'
'            .MoveNext
'        Next
    End With
    
    staThis.Panels(2).Text = ""
    
    Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
    
    mblnBatch = False
    mblnAutoLoadData = True
    
    vsfBill_EnterCell
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub cbsHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
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
    Form_Resize
End Sub


Private Sub Form_Load()
    
    mblnLoad = False
    
    InitComandBars
    Call GetDefineSize
    
    mintMoneyDigit = GetDigit(0, 1, 4)
    mblnNoStock = (Val(zlDataBase.GetPara("存储库房", glngSys, 模块号.药品盘点)) = 1)
    mbln忽略服务对象 = (Val(zlDataBase.GetPara("忽略药品服务对象", glngSys, 模块号.药品盘点)) = 1)
    mintBatchNoLen = GetBatchNoLen()
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    mblnLoadData = False
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品盘点管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    'mnuColDrug.Item(mintDrugNameShow).Checked = True
    mobjPopup.Controls.Item(mintDrugNameShow + 1).Checked = True
    
    Call initCard
    
    mblnLoad = True
    mblnAutoLoadData = True
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
        Set cbrControlMain = .Add(xtpControlButton, mcon按分类货位提取, "按分类、货位提取")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        Set cbrControlMain = .Add(xtpControlButton, mcon保存, "保存")
        cbrControlMain.BeginGroup = True
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
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Me.Pic单据.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - staThis.Height

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
    
    TxtCheckDate.Left = vsfBill.Left + vsfBill.Width - TxtCheckDate.Width
    lblCheckDate.Left = TxtCheckDate.Left - lblCheckDate.Width - 100
    
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
    
    With vsfBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    Me.cbsMain(1).Controls.Find(xtpControlButton, mcon保存退出, , True).Visible = mint编辑状态 = 1
End Sub


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional blnSuccess As Boolean = False)
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

    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "保存退出"
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "审核(&V)"
        'CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Caption = "打印(&P)"
        'CmdSave.Caption = "打印(&P)"
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon确定, , True).Visible = False
        'CmdSave.Visible = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon按库存提取, , True).Visible = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon按分类货位提取, , True).Visible = False
        Me.cbsMain(1).Controls.Find(xtpControlButton, mcon保存, , True).Visible = False
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
           & "FROM 药品单据性质 A, 药品入出类别 B " _
           & "Where A.类别id = B.ID AND A.单据 = 14  and b.系数=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品盘点记录单的入库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
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

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim str药名 As String
    Dim strSqlOrder As String
    
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
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
'            Txt修改人 = UserInfo.用户姓名
'            Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
            TxtCheckDate.Text = Txt填制日期.Caption
            txtStock = mfrmMain.cboStock.Text
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Call 获取单位
            initGrid
        Case 2, 3, 4
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Call 获取单位
            initGrid
            If mint编辑状态 <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "select distinct b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id " _
                        & "  and A.单据 = 14 and a.no=[1]"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!名称
                txtStock.Tag = rsInitCard!id
                
                rsInitCard.Close
            End If
            
            strUnitQuantity = "A.扣率 实盘数量,A.填写数量 帐面数量,A.实际数量 数量差,B.住院单位 AS 住院单位,B.住院包装 as 住院系数,a.零售价*B.住院包装 as 住院售价,"
            strUnitQuantity = strUnitQuantity & "B.门诊单位 AS 门诊单位,B.门诊包装 as 门诊系数,a.零售价*B.门诊包装 as 门诊售价,"
            strUnitQuantity = strUnitQuantity & "B.药库单位 AS 药库单位,B.药库包装 as 药库系数,a.零售价*B.药库包装 as 药库售价,"
            strUnitQuantity = strUnitQuantity & "D.计算单位 AS 售价单位,'1' as 售价系数,a.零售价 as 售价售价,"

            gstrSQL = "SELECT * " & _
                " FROM " & _
                " (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                " NVL(B.最大效期,0) 最大效期,B.药品来源,B.基本药物,D.规格,A.产地,A.原产地,Nvl(A.库房货位,C.库房货位) As 库房货位, A.批号,A.效期,A.批次," & strUnitQuantity & _
                " A.零售金额 AS 金额差,A.差价 AS 差价差,A.零售价,A.单量 As 成本价, " & _
                " A.摘要,填制人,填制日期,修改人,修改日期,审核人,审核日期,A.频次 AS 盘点时间,A.成本价 AS 库存金额,A.成本金额 AS 库存差价,B.加成率,D.是否变价,B.药房分批 AS 药房分批核算,A.批准文号 " & _
                " FROM 药品收发记录 A, 药品规格 B,收费项目别名 E ,收费项目目录 D,药品储备限额 C " & _
                " WHERE A.药品ID = B.药品ID AND b.药品ID=D.ID " & _
                " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                " AND A.药品ID=C.药品ID(+) AND A.库房ID=C.库房ID(+) AND A.记录状态 =[2] " & _
                " AND A.单据 =14 AND A.NO = [1]) " & _
                " ORDER BY " & strSqlOrder
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt填制人 = rsInitCard!填制人
            Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd HH:mm:ss")
            
            Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
            Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd HH:mm:ss"))
            
            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd HH:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            TxtCheckDate.Text = rsInitCard!盘点时间
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            intRow = 0
            With vsfBill
                Do While Not rsInitCard.EOF
                    
                    intRow = intRow + 1
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
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsInitCard!库房货位), "", rsInitCard!库房货位)
                    .TextMatrix(intRow, mconIntCol单位) = IIf(IsNull(rsInitCard.Fields(Split(mstr单位, "|")(1)).Value), "", rsInitCard.Fields(Split(mstr单位, "|")(1)).Value)
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx(rsInitCard!加成率 / 100, 2, , True) & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = 获取比例系数(rsInitCard)

                    If mbln相同单位 = True Then
                        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Nvl(rsInitCard!成本价, 0) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0)), mintPriceDigit, , True)
                        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Nvl(rsInitCard!零售价, 0) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0)), mintPriceDigit, , True)
                    
                        .TextMatrix(intRow, mconIntCol大单位数量) = zlStr.FormatEx(rsInitCard.Fields("实盘数量").Value / Split(获取比例系数(rsInitCard), "|")(0), mintNumberDigit0, , True)
                        .TextMatrix(intRow, mconintCol大单位) = IIf(IsNull(rsInitCard.Fields(Split(mstr单位, "|")(0)).Value), "", rsInitCard.Fields(Split(mstr单位, "|")(0)).Value)
                    Else
                        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Nvl(rsInitCard!成本价, 0) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1)), mintPriceDigit, , True)
                        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Nvl(rsInitCard!零售价, 0) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1)), mintPriceDigit, , True)
                    
                        .TextMatrix(intRow, mconIntCol大单位数量) = zlStr.FormatEx(Int(rsInitCard.Fields("实盘数量").Value / Split(获取比例系数(rsInitCard), "|")(0)), mintNumberDigit0, , True)
                        .TextMatrix(intRow, mconintCol大单位) = IIf(IsNull(rsInitCard.Fields(Split(mstr单位, "|")(0)).Value), "", rsInitCard.Fields(Split(mstr单位, "|")(0)).Value)
                        
                        .TextMatrix(intRow, mconIntCol小单位数量) = zlStr.FormatEx((rsInitCard.Fields("实盘数量").Value / Split(获取比例系数(rsInitCard), "|")(0) - Val(.TextMatrix(intRow, mconIntCol大单位数量))) * Split(获取比例系数(rsInitCard), "|")(0) / Val(Split(获取比例系数(rsInitCard), "|")(1)), mintNumberDigit1, , True)
                        .TextMatrix(intRow, mconintCol小单位) = IIf(IsNull(rsInitCard.Fields(Split(mstr单位, "|")(1)).Value), "", rsInitCard.Fields(Split(mstr单位, "|")(1)).Value)
                        
                        .TextMatrix(intRow, mconintCol数量_合计) = zlStr.FormatEx(rsInitCard.Fields("实盘数量").Value, mintNumberDigit1, , True)
                        .TextMatrix(intRow, mconintCol单位_合计) = IIf(IsNull(rsInitCard.Fields("售价单位")), "", rsInitCard.Fields("售价单位"))
                    End If
                    
                    .RowData(intRow) = Val(IIf(IsNull(rsInitCard!最大效期), 0, rsInitCard!最大效期))
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    Call vsfColHidden '中药类库房才显示"原产地"列
    Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
    Call 显示合计金额
    mint库存检查 = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    mblnLoadData = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function 获取单位() As String
    Dim intUnit As Integer, strUnit As String, strDefault As String
    Dim strCompare As String
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
    '取得缺省单位
    strDefault = GetDrugUnit(Val(txtStock.Tag), "药品盘点管理")
    
    '取盘点单的指定单位
    intUnit = Val(zlDataBase.GetPara("小包装单位", glngSys, 模块号.药品盘点))
    
    If intUnit = 0 Then
        strUnit = strDefault
    Else
        strUnit = Split(strCompare, ";")(intUnit - 1)
    End If
    
    '将指定单位与缺省单位按大单位、小单位的顺序排列
    mintDefault = 1
    If strUnit <> strDefault Then
        If InStr(1, strCompare, strUnit) < InStr(1, strCompare, strDefault) Then
            获取单位 = strUnit & "|" & strDefault
        Else
            mintDefault = 0
            获取单位 = strDefault & "|" & strUnit
        End If
    Else
        获取单位 = strUnit & "|" & strDefault
    End If
    
    mstr单位 = 获取单位
    
    '取大单位的精度（售价、数量、金额）
    Select Case Split(mstr单位, "|")(0)
        Case "售价单位"
            intUnit = conint售价单位
        Case "门诊单位"
            intUnit = conint门诊单位
        Case "住院单位"
            intUnit = conint住院单位
        Case "药库单位"
            intUnit = conint药库单位
    End Select

    mintCostDigit = GetDigit(int性质, conInt药品, conInt成本价, intUnit)
    mintPriceDigit = GetDigit(int性质, conInt药品, conInt售价, intUnit)
    mintNumberDigit0 = GetDigit(int性质, conInt药品, conInt数量, intUnit)
    mintMoneyDigit = GetDigit(int性质, conInt药品, conInt金额)
    
    '取小单位的精度（数量）
    Select Case Split(mstr单位, "|")(1)
        Case "售价单位"
            intUnit = conint售价单位
        Case "门诊单位"
            intUnit = conint门诊单位
        Case "住院单位"
            intUnit = conint住院单位
        Case "药库单位"
            intUnit = conint药库单位
    End Select
    mintNumberDigit1 = GetDigit(int性质, conInt药品, conInt数量, intUnit)
    
    mbln相同单位 = False
    If Split(mstr单位, "|")(0) = Split(mstr单位, "|")(1) Then
        mbln相同单位 = True
    End If
End Function

'初始化编辑控件
Private Sub initGrid()
    Dim i As Integer
    
    With vsfBill
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntColS
        .Editable = flexEDNone
        .RowHeightMax = 315
        .ExplorerBar = flexExSortShow
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
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
        .TextMatrix(0, mconIntcol加成率) = "加成率"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconintCol帐面数量) = "帐面数量"
        .TextMatrix(0, mconIntCol大单位数量) = IIf(mbln相同单位, "数量", "大包装")
        .TextMatrix(0, mconintCol大单位) = "单位"
        .TextMatrix(0, mconIntCol小单位数量) = "小包装"
        .TextMatrix(0, mconintCol小单位) = "单位"
        .TextMatrix(0, mconintCol数量_合计) = "合计"
        .TextMatrix(0, mconintCol单位_合计) = "单位"
        .TextMatrix(0, mconintCol标志) = "标志"
        .TextMatrix(0, mconintCol数量差) = "数量差"
        .TextMatrix(0, mconintCol成本价) = "成本价"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconintCol金额差) = "金额差"
        .TextMatrix(0, mconintCol差价差) = "差价差"
        .TextMatrix(0, mconintCol盘点金额) = "盘点金额"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
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
        .ColWidth(mconIntCol单位) = 0
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .colHidden(mconIntCol批准文号) = True '默认不显示
        .ColWidth(mconintCol帐面数量) = 0
        .ColWidth(mconIntCol大单位数量) = 1000
        .ColWidth(mconintCol大单位) = 500
        .ColWidth(mconIntCol小单位数量) = IIf(mbln相同单位, 0, 1000)
        .ColWidth(mconintCol小单位) = IIf(mbln相同单位, 0, 500)
        .ColWidth(mconintCol数量_合计) = IIf(mbln相同单位, 0, 1000)
        .ColWidth(mconintCol单位_合计) = IIf(mbln相同单位, 0, 500)
        .ColWidth(mconintCol标志) = 0
        .ColWidth(mconintCol数量差) = 0
        .ColWidth(mconintCol成本价) = IIf(mblnViewCost, 1000, 0)
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconintCol金额差) = 0
        .colHidden(mconintCol金额差) = True '默认不显示
        .ColWidth(mconintCol差价差) = 0
        .colHidden(mconintCol差价差) = True '默认不显示
        .ColWidth(mconintCol盘点金额) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
                
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            txt摘要.Enabled = False
        End If
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconintCol大单位) = flexAlignCenterCenter
        .ColAlignment(mconintCol小单位) = flexAlignCenterCenter
        .ColAlignment(mconintCol单位_合计) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconintCol帐面数量) = flexAlignRightCenter
        .ColAlignment(mconintCol标志) = flexAlignCenterCenter
        .ColAlignment(mconintCol数量差) = flexAlignRightCenter
        .ColAlignment(mconintCol成本价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconintCol金额差) = flexAlignRightCenter
        .ColAlignment(mconintCol差价差) = flexAlignRightCenter
        .ColAlignment(mconintCol盘点金额) = flexAlignRightCenter
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .Cell(flexcpFontBold, 1, mconIntCol大单位数量, 1, mconIntCol大单位数量) = True
        .Cell(flexcpFontBold, 1, mconIntCol小单位数量, 1, mconIntCol小单位数量) = True
        
        .Redraw = flexRDDirect
    End With
    txt摘要.MaxLength = Sys.FieldsLength("药品收发记录", "摘要")
    
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, MStrCaption
    '恢复个性化参数设置后，权限控制的列需要进一步控制显示
    vsfBill.ColWidth(mconintCol成本价) = IIf(mblnViewCost, 1000, 0)
    
    vsfBill.ColWidth(mconIntCol小单位数量) = IIf(mbln相同单位, 0, 1000)
    vsfBill.ColWidth(mconintCol小单位) = IIf(mbln相同单位, 0, 500)
    
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

Private Function 获取比例系数(ByVal rsData As ADODB.Recordset) As String
    获取比例系数 = Replace(mstr单位, "单位", "系数")
    获取比例系数 = rsData.Fields(Split(获取比例系数, "|")(0)).Value & "|" & rsData.Fields(Split(获取比例系数, "|")(1)).Value
End Function

Private Sub 显示合计金额()
End Sub


Private Sub GetBatchRec()
    '提取库存所有记录
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim lngRows As Long
    Dim i As Integer
    Dim strTemp As Variant
    Dim rsProperty As ADODB.Recordset           '药品规格
    Dim rs货位 As ADODB.Recordset       '货位
    Dim arrDrugID As Variant
    Dim j As Integer
    Dim lng药品ID As Long
    Dim x As Integer
    Dim str药品ID As String
    Dim strArry As Variant  '保存货位的数组
    Dim str货位id As String
    Dim str货位 As String
    Dim str货位sql As String
    Dim strCompare As String
    Dim strOrder As String
    
    On Error GoTo errHandle
    Set rsProperty = New ADODB.Recordset
    With rsProperty
        If .State = 1 Then .Close
        .Fields.Append "药品编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品id", adDouble, 50, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批准文号", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rs货位 = New ADODB.Recordset
    
    With rs货位
        If .State = 1 Then .Close
        .Fields.Append "药品id", adDouble, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    x = 1
    strArry = Array()
    str货位id = ""
    For j = 0 To UBound(Split(mstr货位, ",")) - 1
        str货位 = Mid(mstr货位, x, InStr(x, mstr货位, ",") - x)
        x = InStr(x, mstr货位, ",") + 1
        If Len(IIf(str货位id = "", "", str货位id & ",") & str货位) > 4000 Then
            ReDim Preserve strArry(UBound(strArry) + 1)
            strArry(UBound(strArry)) = str货位id
            str货位id = str货位
        Else
            str货位id = IIf(str货位id = "", "", str货位id & ",") & str货位
        End If
    Next
    
    If str货位id <> "" Then
'        SQLCondition.str药品ID = ""
        ReDim Preserve strArry(UBound(strArry) + 1)
        strArry(UBound(strArry)) = str货位id
        
        gstrSQL = " Select distinct a.药品id" & _
                    " From 药品储备限额 A," & _
                         "收费项目目录 C,(select * from Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) B" & _
                    " Where a.库房id = [1] and a.药品id=c.id And (Instr(',' || a.库房货位 || ',', ',' || b.Column_Value || ',') > 0) "
        
        If mbln忽略服务对象 = False Then
            gstrSQL = gstrSQL & _
                " and (Decode(c.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3)) " & _
                    " or Decode(c.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                    " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]))"
        End If
        
        For i = 0 To UBound(strArry)
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "依据货位查询药品", Val(txtStock.Tag), CStr(strArry(i)))
            
            If Not rsData.EOF Then
                Do While Not rsData.EOF
                    With rs货位
                        .AddNew
                        !药品id = rsData!药品id
                        
                        .Update
                    End With
                    rsData.MoveNext
                Loop
            End If
        Next
    End If
    
'    If rs货位.RecordCount > 0 Then
'        rsData.MoveFirst
'        For i = 0 To rsData.RecordCount - 1 '如果选择了货位，则按照货位优先取药品，然后按照优先取出的药品在从库存表中取数据
'            SQLCondition.str药品ID = rsData!药品ID & "," & SQLCondition.str药品ID
'            rsData.MoveNext
'        Next
'    End If
    
'    If SQLCondition.str药品ID = "" Then
'        MsgBox "未查询到数据！", vbInformation, gstrSysName
'        Exit Sub
'    Else
        If SQLCondition.str药品ID <> "" And str货位id <> "" Then
            strTemp = Split(SQLCondition.str药品ID, ",")
            SQLCondition.str药品ID = ""
            
            For i = 0 To UBound(strTemp) - 1
                rs货位.MoveFirst
                For j = 0 To rs货位.RecordCount - 1
                    If rs货位.EOF Then Exit For
                    If Val(strTemp(i)) = Val(rs货位!药品id) Then
                        SQLCondition.str药品ID = strTemp(i) & "," & SQLCondition.str药品ID
                    End If
                    If j <> rs货位.RecordCount - 1 Then
                        rs货位.MoveNext
                    End If
                Next
            Next
        ElseIf SQLCondition.str药品ID = "" And str货位id <> "" Then
            If rs货位.RecordCount > 0 Then
                rs货位.MoveFirst
            End If
            
            Do While Not rs货位.EOF
                SQLCondition.str药品ID = rs货位!药品id & "," & SQLCondition.str药品ID
                rs货位.MoveNext
            Loop
        ElseIf SQLCondition.str药品ID = "" And str货位id = "" Then
            Exit Sub
        End If
        
        x = 1
        arrDrugID = Array()
        str药品ID = ""
        For j = 0 To UBound(Split(SQLCondition.str药品ID, ",")) - 1
            lng药品ID = Mid(SQLCondition.str药品ID, x, InStr(x, SQLCondition.str药品ID, ",") - x)
            x = InStr(x, SQLCondition.str药品ID, ",") + 1
            If Len(IIf(str药品ID = "", "", str药品ID & ",") & lng药品ID) > 4000 Then
                ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
                arrDrugID(UBound(arrDrugID)) = str药品ID
                str药品ID = lng药品ID
            Else
                str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & lng药品ID
            End If
        Next
        
        If str药品ID = "" And UBound(arrDrugID) < 0 Then
            Exit Sub
        ElseIf str药品ID <> "" Then
            ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
            arrDrugID(UBound(arrDrugID)) = str药品ID
        End If
        
        gstrSQL = "Select b.编码 As 药品编码, a.药品id, b.名称, Nvl(a.批次, 0) As 批次, a.批准文号, c.库房货位" & vbNewLine & _
                " From 药品库存 A, 收费项目目录 B, 药品储备限额 C" & vbNewLine & _
                " Where a.性质 = 1 And a.药品id = b.Id And b.Id = c.药品id(+) And a.库房id = [1] And c.库房id = [1] And " & _
                " b.Id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
                " And (Nvl(A.实际数量,0)<>0 Or Nvl(A.实际金额,0)<>0 Or Nvl(A.实际差价,0)<>0 )"

        If mbln忽略服务对象 = False Then
            gstrSQL = gstrSQL & _
                " and (Decode(b.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(1,3)) " & _
                    " or Decode(b.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[1] and 服务对象 in(2,3)) " & _
                    " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[1]))"
        End If
        
        strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
        strCompare = Mid(strOrder, 1, 1)
        gstrSQL = gstrSQL & " Order by " & IIf(strCompare = "0", "b.编码", IIf(strCompare = "1", "b.编码", IIf(strCompare = "2", "b.名称", "c.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",b.编码"

        For i = 0 To UBound(arrDrugID)
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "GetBatchRec", Val(txtStock.Tag), CStr(arrDrugID(i)))
            
            If Not rsData.EOF Then
                Do While Not rsData.EOF
                    With rsProperty
                        .AddNew
                        !药品编码 = rsData!药品编码
                        !药品id = rsData!药品id
                        !批次 = rsData!批次
                        !批准文号 = rsData!批准文号
                        
                        .Update
                    End With
                    rsData.MoveNext
                Loop
            End If
        Next
'    End If
    
    If rsProperty.RecordCount = 0 Then
        Exit Sub
    End If
    rsProperty.MoveFirst
    With rsProperty
        If .RecordCount = 0 Then Exit Sub
        
        mblnBatch = True
        
        lngRows = .RecordCount
        
        vsfBill.rows = lngRows + 1
        
        For lngRow = 1 To lngRows
            vsfBill.Row = lngRow
            Call SetPhiscRows(!药品id, !批次, Nvl(!批准文号, ""), True)
            
            DoEvents
            Call zlControl.StaShowPercent(lngRow / lngRows, staThis.Panels(2), frmCheckCourseCard)
            DoEvents
            
            .MoveNext
        Next
    End With
    
    staThis.Panels(2).Text = ""
    
    Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
    
    mblnBatch = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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



Private Sub cbsCancel()
    Unload Me
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

Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long, lngRows As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim str其他名称 As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim str药名 As String
    
    '查找药品
    On Error GoTo errHandle
    If strInput = txtCode.Tag Then
        '表示查找下一条记录
        If mlngFind >= vsfBill.rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '表示新的查找
        mlngFindFirst = 0
        lngStart = 0
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    mlngFind = 0
    lngStart = lngStart + 1
    lngRows = vsfBill.rows - 1
    
    mrsFindName.MoveFirst
    For n = 1 To mrsFindName.RecordCount
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If
    
        lngFindRow = vsfBill.FindRow(str药名, lngStart, mconIntCol药品编码和名称, True, True)
        If lngFindRow > 0 Then
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
            mlngFind = lngFindRow
            
            '记录找到的第1条记录
            If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
            
            Exit For
        End If
        mrsFindName.MoveNext
        
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF And lngFindRow = -1 And mlngFindFirst <> 0 Then
            vsfBill.Select mlngFindFirst, 1, mlngFindFirst, vsfBill.Cols - 1
            vsfBill.TopRow = mlngFindFirst
            mlngFind = mlngFindFirst
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cbsSave(ByVal lngControlId As Long)
'参数：lngControlId表示保存的模式，参数结合 mint编辑状态 使用
    Dim blnSuccess As Boolean
    
    '设置排序数据集
    Call SetSortRecord
    
    Me.txtNo.Tag = ""
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.药品盘点)) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then   '修改保存后处理
            If lngControlId = mcon确定 Then
                Unload Me
                Exit Sub
            Else
                MsgBox "保存成功！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    
    If mint编辑状态 = 1 Then '新增保存后处理
        If lngControlId = mcon保存退出 Then
            Unload Me
            Exit Sub
        End If
        
        If lngControlId = mcon保存 Then
            txtNo.Caption = txtNo.Tag
            mstr单据号 = txtNo.Tag
            mbln即时保存 = True
            MsgBox "保存成功！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If lngControlId = mcon确定 Then
            txtNo.Caption = ""
            mstr单据号 = ""
            mbln即时保存 = False
            TxtCheckDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
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
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
            
    mblnFirst = False
    mbln盘停用药品 = IIf(Val(zlDataBase.GetPara("盘已停用的药品", glngSys, glngModul, 0)) = 0, False, True)
    If mint编辑状态 = 1 Then
        Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
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
    End If
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
    
    vsfBill.SetFocus
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) = "" Then
        vsfBill.Col = mconIntCol药名
    Else
        vsfBill.Col = mconIntCol大单位数量
        vsfBill.EditCell
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnAutoLoadData Then Cancel = 1: Exit Sub '数据没加载完不准退出
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品盘点管理", "药品名称显示方式", mintDrugNameShow)
    
    mbln即时保存 = False
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
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象)
    End If
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntCol药名
                .ColComboList(mconIntCol药名) = "..."
            Case mconIntCol产地
                .ColComboList(mconIntCol产地) = "..."
            Case mconIntCol原产地
                .ColComboList(mconIntCol原产地) = "..."
        End Select
    End With
End Sub

Private Sub vsfBill_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    With vsfBill
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

Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    Dim rsProvider As ADODB.Recordset
    
    intOldRow = vsfBill.Row
    With vsfBill
        Select Case Col
            Case mconIntCol药名
'                If mblnNotTrigger = True Then
'                    mblnNotTrigger = False
'                    Exit Sub
'                End If
                
                If mblnNotTrigger <> True Then
                    mblnNotTrigger = True
'                    Set RecReturn = Frm药品选择器.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , False, True, False, True, zlStr.IsHavePrivs(mstrPrivs, "查看盘点单库存"), 0, mblnNoStock, 0, False, mbln忽略服务对象)
                    If grsMaster.State = adStateClosed Or mstrLast盘点时间 <> TxtCheckDate.Text Then
                        mstrLast盘点时间 = TxtCheckDate.Text
                        Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象, TxtCheckDate.Text)
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , txtStock.Tag, txtStock.Tag, , 0, False, True, zlStr.IsHavePrivs(mstrPrivs, "查看盘点单库存"), IIf(mbln盘停用药品, 1, 0), , mstrPrivs, TxtCheckDate.Text)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
                    End If
                    mblnNotTrigger = False
                Else
                    Exit Sub
                End If
                
                If RecReturn.RecordCount > 0 Then
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        intCurRow = .Row
                        Call SetPhiscRows(RecReturn!药品id, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号))
'                        .EditCell
                        
                        If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                            .rows = .rows + 1
                        End If
                        .Row = .rows - 1
                        RecReturn.MoveNext
                    Next
                    .Row = intOldRow
                    If Val(.TextMatrix(Row, mconIntCol批次)) = -1 And .TextMatrix(Row, mconIntCol批号) = "" Then
                        .Col = mconIntCol批号
                    Else
                        .Col = mconIntCol大单位数量
                    End If
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
    Dim strNotNum As String  '无库存药品
    Dim str重复药名 As String   '用来记录重复选择了的药品名称
    Dim strNot药名 As String    '用来记录哪些药品是时价但无库存
    Dim str盘点时间后药品 As String       '纪录在盘点时间后建立的药品
    Dim strSql盘点 As String   '过滤盘点时间后建立的药品
    
    rsTemp.MoveFirst
    str盘点时间后药品 = ""
    strSql盘点 = ""
    str批次 = ""
    strTemp = ""
    Do While Not rsTemp.EOF
    
        str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
        
        gstrSQL = "Select a.建档时间 From 收费项目目录 A Where a.Id =[1]"
        Set rs建档时间 = zlDataBase.OpenSQLRecord(gstrSQL, "查询建档时间", rsTemp!药品id)
        If Format(rs建档时间!建档时间, "yyyy-MM-dd HH:mm:ss") > Format(TxtCheckDate.Text, "yyyy-MM-dd HH:mm:ss") Then
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
            '35.60版本支持同个分批药品录入多个批次，不检查批次=-1(新增批次)的数据
            If Val(.TextMatrix(i, mconIntCol批次)) >= 0 Then
                If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 Then
                    strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
                End If
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
        
        '判断以什么方式拼接sql
        If str重复药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
End Function

Private Sub vsfBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsfBill_EnterCell()
    Dim lng批次 As Long
    Dim intRow As Integer
    
    If mblnBatch = True Then Exit Sub
    
    With vsfBill
        .Editable = flexEDNone
        .FocusRect = flexFocusLight
        
        If mint编辑状态 = 4 Then Exit Sub
        
        lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
        
        Select Case .Col
            Case mconIntCol药名
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol药名) = "..."
                    
                End If
                
            Case mconIntCol批号
                .EditMaxLength = mintBatchNoLen
                
                If lng批次 = -1 Then
                    .Editable = flexEDKbdMouse
                    
                End If
            Case mconIntCol产地
                If lng批次 = -1 And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
                    .EditMaxLength = mlng生产商长度
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol产地) = "..."
                    
                End If
            Case mconIntCol原产地
                If lng批次 = -1 And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
                    .EditMaxLength = mlng原产地长度
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol原产地) = "..."
                    
                End If
            Case mconIntCol效期
'                .TextMask = "1234567890-"
                .EditMaxLength = 10
                
                If lng批次 = -1 Then
                    .Editable = flexEDKbdMouse
                    
                End If
                
                If .TextMatrix(.Row, mconIntCol批号) <> "" And Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                        
                        If Len(strxq) = 8 Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                If .RowData(.Row) = 0 Then Exit Sub
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
            Case mconIntCol大单位数量, mconIntCol小单位数量
                .EditMaxLength = 16
'                .TextMask = ".1234567890"
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                    .Editable = flexEDKbdMouse
                    
                End If
            Case mconintCol成本价
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                    If Val(.TextMatrix(.Row, mconIntCol批次)) = -1 Then
                       .Editable = flexEDKbdMouse
                       
                    End If
                End If
        End Select
        
        
        If mlngCurrRow <> .Row Then
            mlngCurrRow = .Row
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
            If InStr(1, "34", mint编辑状态) <> 0 Then Exit Sub
            
            If MsgBox("是否删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshRowNO(vsfBill, mconIntCol行号, .Row)
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
                    Else
                        If Val(.TextMatrix(.Row, mconIntCol批次)) = -1 And .TextMatrix(.Row, mconIntCol批号) = "" Then
                            .Col = mconIntCol批号
                        Else
                            .Col = mconIntCol大单位数量
                        End If
                        .EditCell
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
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    Dim rsProvider As ADODB.Recordset
    
    intOldRow = vsfBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfBill
        .EditText = Trim(.EditText)
        strkey = UCase(Trim(.EditText))
        
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
                    
                    sngLeft = Me.Left + Pic单据.Left + vsfBill.Left + vsfBill.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + vsfBill.Top + vsfBill.CellTop + vsfBill.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - vsfBill.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , strkey, sngLeft, sngTop, False, True, False, True, zlStr.IsHavePrivs(mstrPrivs, "查看盘点单库存"), 0, mblnNoStock, 0, False, mbln忽略服务对象)
                    If grsMaster.State = adStateClosed Or mstrLast盘点时间 <> TxtCheckDate.Text Then
                        mstrLast盘点时间 = TxtCheckDate.Text
                        Call SetSelectorRS(2, "药品盘点管理", txtStock.Tag, txtStock.Tag, , , , mbln盘停用药品, mblnNoStock, 1, , , mbln忽略服务对象, TxtCheckDate.Text)
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, txtStock.Tag, txtStock.Tag, , 0, False, True, zlStr.IsHavePrivs(mstrPrivs, "查看盘点单库存"), IIf(mbln盘停用药品, 1, 0), , mstrPrivs, TxtCheckDate.Text)
                    
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            Call SetPhiscRows(RecReturn!药品id, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号))
                            
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                        If Val(.TextMatrix(Row, mconIntCol批次)) = -1 And .TextMatrix(Row, mconIntCol批号) = "" Then
                            .Col = mconIntCol批号
                        Else
                            .Col = mconIntCol大单位数量
                        End If
                    End If
                    Call 提示库存数
                End If
            Case mconIntCol批号
                '无处理
                .TextMatrix(.Row, mconIntCol批号) = strkey
                
                If .TextMatrix(.Row, mconIntCol效期) = "" Then
                    .Col = mconIntCol效期
                Else
                    .Col = mconIntCol大单位数量
                End If
                .EditCell
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
                
                .TextMatrix(.Row, mconIntCol效期) = strkey
                Call CheckLapse(.TextMatrix(.Row, mconIntCol效期)) '检查是否失效
                
                .Col = mconIntCol大单位数量
                .EditCell
            Case mconIntCol大单位数量, mconIntCol小单位数量
                If strkey <> "" Then
                    If Not IsNumeric(strkey) And strkey <> "" Then
                        MsgBox "对不起，实盘数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    End If
                Else
                    .EditText = IIf(.TextMatrix(.Row, .Col) = "", " ", .TextMatrix(.Row, .Col))
                    .TextMatrix(.Row, .Col) = .EditText
                End If
                
                If strkey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    If .Col = mconIntCol大单位数量 Then
                        strkey = zlStr.FormatEx(strkey, mintNumberDigit0, , True)
                    Else
                        strkey = zlStr.FormatEx(strkey, mintNumberDigit1, , True)
                    End If
                    .EditText = strkey
                End If
                
                '显示合计数量
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If .Col = mconIntCol大单位数量 Then
                    strkey = Val(.TextMatrix(.Row, mconIntCol小单位数量)) + Val(strkey) * Val(Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(0)) / Val(Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(1))
                Else
                    strkey = Val(strkey) + Val(.TextMatrix(.Row, mconIntCol大单位数量)) * Val(Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(0)) / Val(Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(1))
                End If
                .TextMatrix(.Row, mconintCol数量_合计) = zlStr.FormatEx(strkey * Val(Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(1)), mintNumberDigit1, , True)
                
                Call 显示合计金额
                
                If Col = mconIntCol大单位数量 Then
                    If .ColWidth(mconIntCol小单位数量) > 0 Then
                        .Col = mconIntCol小单位数量
                    Else
                        '如果下一行为空或者药名列为空则返回到药名列，否则返回到实盘数量列
                        If .Row < .rows - 1 Then
                            .Row = .Row + 1
                            If .TextMatrix(.Row, mconIntCol药名) <> "" Then
                                .Col = mconIntCol大单位数量
                            Else
                                .Col = mconIntCol药名
                            End If
                        Else
                            .rows = .rows + 1
                            .Row = .rows - 1
                            .TextMatrix(.Row, mconIntCol行号) = .Row
                            .Col = mconIntCol药名
                            
                            .Cell(flexcpFontBold, .rows - 1, mconIntCol大单位数量, .rows - 1, mconIntCol大单位数量) = True
                            .Cell(flexcpFontBold, .rows - 1, mconIntCol小单位数量, .rows - 1, mconIntCol小单位数量) = True
                        End If
                    End If
                Else
                    '如果下一行为空或者药名列为空则返回到药名列，否则返回到实盘数量列
                    If .Row < .rows - 1 Then
                        .Row = .Row + 1
                        If .TextMatrix(.Row, mconIntCol药名) <> "" Then
                            .Col = mconIntCol大单位数量
                        Else
                            .Col = mconIntCol药名
                        End If
                    Else
                        .rows = .rows + 1
                        .Row = .rows - 1
                        .TextMatrix(.Row, mconIntCol行号) = .Row
                        .Col = mconIntCol药名
                        
                        .Cell(flexcpFontBold, .rows - 1, mconIntCol大单位数量, .rows - 1, mconIntCol大单位数量) = True
                        .Cell(flexcpFontBold, .rows - 1, mconIntCol小单位数量, .rows - 1, mconIntCol小单位数量) = True
                    End If
                End If
            End Select
    End With
End Sub


Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBill
        Select Case Col
            Case mconIntCol大单位数量, mconIntCol小单位数量
                strkey = .EditText
                If strkey = "" Then
                    strkey = .TextMatrix(.Row, .Col)
                End If
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(.EditText, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    
                    Select Case .Col
                        Case mconIntCol大单位数量
                            intDigit = mintNumberDigit0
                        Case mconIntCol小单位数量
                            intDigit = mintNumberDigit1
                    End Select
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

Private Sub TxtCheckDate_GotFocus()
    With TxtCheckDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheckDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub TxtCheckDate_LostFocus()
    If Not IsDate(TxtCheckDate.Text) Then
        MsgBox "请输入正确的日期格式。"
        TxtCheckDate.SetFocus
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim lng药品ID As Long
    Dim str产地 As String, str批号 As String, dbl成本价 As Double
    Dim intRow As Integer
    
    With vsfBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    '分批药品必须录入产地和批号
                    If Val(.TextMatrix(intLop, mconIntCol批次)) = -1 And (.TextMatrix(intLop, mconIntCol产地) = "" Or .TextMatrix(intLop, mconIntCol批号) = "") Then
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
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行药品的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntCol批号
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol大单位数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的大包装数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntCol大单位数量
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol小单位数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的小包装数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconIntCol小单位数量
                        .EditCell
                        Exit Function
                    End If
                End If
            Next
            
            '检查分批药品新增批次的产地，批号是否重复
            For intLop = 1 To .rows - 1
                If Val(.TextMatrix(intLop, mconIntCol批次)) = -1 Then
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
    Dim dbl成本价 As Double
    Dim n As Integer
    Dim str库房货位 As String
    Dim arrSql As Variant
    Dim i As Integer
    
    SaveCard = False
    arrSql = Array()
    On Error GoTo errHandle
    '在外面设置入出类别ID，主要是所有药品都要用他
    gstrSQL = "SELECT b.系数,b.id AS 类别id " _
            & "FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID AND a.单据 = 14 "
    Set rs入出类别 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
    
    If rs入出类别.EOF Then
        MsgBox "对不起，没有设置药品盘点管理的入出类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng入库类别ID = 0
    lng出库类别ID = 0
    
    If rs入出类别!系数 = 1 Then lng入库类别ID = rs入出类别!类别id
    rs入出类别.Close
    
    If lng入库类别ID = 0 Then
        MsgBox "对不起，没有设置药品盘点记录单的入库类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If CheckTime = False Then Exit Function
    
    With vsfBill
        chrNo = Trim(txtNo)
        lng库房ID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(62, lng库房ID)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        dat填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
        
        str盘点时间 = TxtCheckDate.Text
        
        If mint编辑状态 = 2 Or mbln即时保存 Then         '修改
            gstrSQL = "zl_药品盘点记录单_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            str填制人 = Txt填制人
            dat填制日期 = Format(Txt填制日期, "yyyy-mm-dd HH:mm:ss")
            str修改人 = UserInfo.用户姓名
            dat修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lng药品ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str原产地 = .TextMatrix(intRow, mconIntCol原产地)
                str批号 = Trim(.TextMatrix(intRow, mconIntCol批号))
                lng批次ID = IIf(.TextMatrix(intRow, mconIntCol批次) = "", 0, .TextMatrix(intRow, mconIntCol批次))
                dat效期 = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                dbl帐面数量 = Val(.TextMatrix(intRow, mconintCol帐面数量)) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1))
                
                dbl实盘数量 = Val(.TextMatrix(intRow, mconIntCol大单位数量)) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0))
                
                If mbln相同单位 = False Then
                    dbl实盘数量 = dbl实盘数量 + Val(.TextMatrix(intRow, mconIntCol小单位数量)) * Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1))
                End If
                
                dbl数量差 = 0
                
'                If mbln相同单位 = False Then
'                    dbl成本价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1)), gtype_UserDrugDigits.Digit_零售价)
'                    dbl售价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1)), gtype_UserDrugDigits.Digit_零售价)
'                Else
'                    dbl成本价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0)), gtype_UserDrugDigits.Digit_零售价)
'                    dbl售价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0)), gtype_UserDrugDigits.Digit_零售价)
'                End If


                dbl售价 = Get盘点时刻售价(Split(.TextMatrix(intRow, mconIntcol加成率), "||")(1) = 1, lng药品ID, lng库房ID, lng批次ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                                
                '新增价格时去界面价格，不是新增批次时取原始价格
                If lng批次ID = -1 Then
                    If mbln相同单位 = False Then
                        dbl成本价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(1)), gtype_UserDrugDigits.Digit_零售价)
                    Else
                        dbl成本价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol成本价)) / Val(Split(.TextMatrix(intRow, mconIntCol比例系数), "|")(0)), gtype_UserDrugDigits.Digit_零售价)
                    End If
                Else
                    dbl成本价 = Get盘点时刻成本价(lng药品ID, lng库房ID, lng批次ID, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                dbl金额差 = Val(.TextMatrix(intRow, mconintCol金额差))
                dbl差价差 = Val(.TextMatrix(intRow, mconintCol差价差))
                dbl库存金额 = Val(.TextMatrix(intRow, mconIntCol实际金额))
                dbl库存差价 = Val(.TextMatrix(intRow, mconIntCol实际差价))
                str库房货位 = IIf(Trim(.TextMatrix(intRow, mconIntCol库房货位)) = "", "", .TextMatrix(intRow, mconIntCol库房货位))
                
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
                
                'zl_药品盘点记录单_INSERT( /*NO_IN*/, /*序号_IN*/, /*库房ID_IN*/, /*批次_IN*/,
                    '/*入出类别ID_IN*/, /*入出系数_IN*/, /*药品ID_IN*/, /*帐面数量_IN*/,
                    '/*实盘数量_IN*/, /*数量差_IN*/, /*售价_IN*/, /*金额差_IN*/, /*差价差_IN*/,
                    '/*填制人_IN*/, /*填制日期_IN*/, /*摘要_IN*/, /*产地_IN*/, /*批号_IN*/,
                    '/*效期_IN*/, /*盘点时间_IN*/'/*修改人_IN*/, /*修改日期_IN*/ );
                
                gstrSQL = "zl_药品盘点记录单_INSERT('" & chrNo & "'," & lng序号 & "," & lng库房ID & "," & lng批次ID & "," _
                    & lng入出类别id & "," & int入出系数 & "," & lng药品ID & "," & dbl帐面数量 & "," _
                    & dbl实盘数量 & "," & dbl数量差 & "," & dbl售价 & "," & dbl金额差 & "," & dbl差价差 & ",'" _
                    & str填制人 & "',to_date('" & dat填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & str摘要 & "','" & str产地 & "','" & str批号 & "'," & IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & str盘点时间 & "'," & dbl库存金额 & "," & dbl库存差价 & ",'" & str批准文号 & "'," & dbl成本价 & ",'" & str库房货位 & "','" & str原产地 & "','" _
                    & str修改人 & "'," & IIf(dat修改日期 = "", "Null", "to_date('" & dat修改日期 & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                
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
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function



Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    Dim dbl实际数量 As Double
    
    On Error GoTo errHandle
    If Not zlStr.IsHavePrivs(mstrPrivs, "查看盘点单库存") Then Exit Sub
    
    With vsfBill
        If .TextMatrix(.Row, mconIntCol药名) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(vsfBill.Row, 0) = "" Then Exit Sub
        gstrSQL = "select 可用数量/" & Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(1) & " as  可用数量, " & _
            " 实际数量/" & Split(.TextMatrix(.Row, mconIntCol比例系数), "|")(1) & " as  实际数量 " & _
            " from 药品库存 where 库房id=[1] " _
            & " and 药品id=[2] " _
            & " and 性质=1 and " _
            & " nvl(批次,0)=[3]"
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", txtStock.Tag, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            dbl实际数量 = IIf(IsNull(rsUseCount!实际数量), 0, rsUseCount!实际数量)
        End If
        rsUseCount.Close
        
        staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(dbl实际数量, mintNumberDigit1, , True) & "]" & .TextMatrix(.Row, mconIntCol单位)
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

Private Function SetPhiscRows(ByVal lngID As Long, ByVal lng批次 As Long, ByVal str批准文号 As String, Optional ByVal blnBatch As Boolean = False) As Boolean
'功能：根据药品ID在盘存表上显示并处理该药品的初始盘存信息
'说明：
'   1.如果是非分批核算药,且已经输入了,则提示并退出。
'   2.如果是分批核算药，则分别处理该药的未处理的各批次库存行。
    Dim i As Integer, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsPrice As New Recordset
    Dim str药名 As String
    Dim str盘点时间 As String
     
    On Error GoTo errH
    
    str盘点时间 = TxtCheckDate.Text
    
    SetPhiscRows = False
    Set rsData = GetPhysicDetail(txtStock.Tag, lngID, Not blnBatch)
    intRecordCount = rsData.RecordCount
    If intRecordCount = 0 Then Exit Function
    '新增批次药品
    If lng批次 <> -1 Then
        rsData.MoveFirst
        rsData.Find "批次=" & lng批次
        If rsData.EOF Then Exit Function
    End If
    
    With vsfBill
        intRow = .Row
        intCurrentRow = .Row
        
        vsfBill.Redraw = flexRDNone
        
        .TextMatrix(intRow, 0) = rsData!药品id
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = rsData!通用名
        Else
            str药名 = IIf(IsNull(rsData!商品名), rsData!通用名, rsData!商品名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = rsData!药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = rsData!药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsData!商品名), "", rsData!商品名)
        
        If .Col = mconIntCol药名 Then
            .EditText = .TextMatrix(intRow, mconIntCol药名)
        End If
 
        .TextMatrix(intRow, mconIntCol来源) = Nvl(rsData!药品来源)
        .TextMatrix(intRow, mconIntCol基本药物) = Nvl(rsData!基本药物)
        .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
        .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsData!原产地), "", rsData!原产地)
        
        '取该药品的产地
        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
        If .TextMatrix(intRow, mconIntCol产地) = "" Then .TextMatrix(intRow, mconIntCol产地) = Nvl(rsData!缺省产地)
        
        .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
        .TextMatrix(intRow, mconIntCol单位) = IIf(IsNull(rsData.Fields(Split(mstr单位, "|")(1)).Value), "", rsData.Fields(Split(mstr单位, "|")(1)).Value)
        
        If lng批次 = -1 Then
            .TextMatrix(intRow, mconIntCol批次) = lng批次
            .TextMatrix(intRow, mconIntCol批号) = ""
            .TextMatrix(intRow, mconIntCol效期) = ""
            .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        Else
            .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsData!批次), "0", rsData!批次)
            .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
            .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                '换算为有效期
                .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
            End If
                
            .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsData!批准文号), "", rsData!批准文号)
        End If
        
        .TextMatrix(intRow, mconIntCol大单位数量) = ""
        .TextMatrix(intRow, mconIntCol小单位数量) = ""
        .TextMatrix(intRow, mconintCol大单位) = IIf(IsNull(rsData.Fields(Split(mstr单位, "|")(0)).Value), "", rsData.Fields(Split(mstr单位, "|")(0)).Value)
        .TextMatrix(intRow, mconintCol小单位) = IIf(IsNull(rsData.Fields(Split(mstr单位, "|")(1)).Value), "", rsData.Fields(Split(mstr单位, "|")(1)).Value)
        .TextMatrix(intRow, mconintCol数量_合计) = ""
        .TextMatrix(intRow, mconintCol单位_合计) = IIf(IsNull(rsData!售价单位), "", rsData!售价单位)
        .TextMatrix(intRow, mconIntCol比例系数) = 获取比例系数(rsData)
        .TextMatrix(intRow, mconIntcol加成率) = rsData!加成率 / 100 & "||" & rsData!是否变价 & "||" & rsData!药房分批核算
        
        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Nvl(rsData!售价, 0) * rsData.Fields(Replace(Split(mstr单位, "|")(1), "单位", "系数")).Value, mintPriceDigit, , True)
        .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(Nvl(rsData!成本价, 0) * rsData.Fields(Replace(Split(mstr单位, "|")(1), "单位", "系数")).Value, mintPriceDigit, , True)
        
        If rsData!是否变价 = 1 Then
            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Get盘点时刻零售价(CLng(rsData!药品id), txtStock.Tag, CLng(IIf(IsNull(rsData!批次), "0", rsData!批次)), rsData.Fields(Replace(Split(mstr单位, "|")(1), "单位", "系数")).Value, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss"))), mintPriceDigit, , True)
        End If

        .RowData(intRow) = Val(IIf(IsNull(rsData!最大效期), 0, rsData!最大效期))
        rsData.MoveNext
        
        If blnBatch = False Then
            Call RefreshRowNO(vsfBill, mconIntCol行号, 1)
        End If
        
        .Col = IIf(lng批次 = -1, mconIntCol批号, mconIntCol大单位数量)
        .EditCell
        
        vsfBill.Redraw = flexRDDirect
    End With
    
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'打印单据
Private Sub printbill()
'    Dim strUnit As String
'    Dim int单位系数 As Integer
'    Dim StrNo As String
'
'    strUnit = GetDrugUnit(txtStock.Tag)
'    Select Case strUnit
'        Case "住院单位"
'            int单位系数 = 1
'        Case "门诊单位"
'            int单位系数 = 2
'        Case "药库单位"
'            int单位系数 = 3
'        Case "售价单位"             '售价单位：主要是制剂室
'            int单位系数 = 4
'    End Select
'    StrNo = txtNo
'    Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), mint记录状态, int单位系数, 1307, "药品盘点单", StrNo)
End Sub





Private Function GetPhysicDetail(ByVal lng库房ID As Long, ByVal lng药品ID As Long, _
    Optional ByVal bln盘无库存药品 As Boolean = True, Optional ByVal bln汇总盘点单 As Boolean = False) As ADODB.Recordset
    'bln盘无库存药品=是否将无库存药品也提取出来
    'bln汇总盘点单=是否需要汇总指定盘点时间的盘点单形成盘点表
    '提取该药品当前库房所有批次明细记录
    Dim str单位 As String, str盘点时间 As String, str汇总盘点单 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    str盘点时间 = TxtCheckDate.Text
    str单位 = ",A.住院单位,A.住院包装 AS 住院系数"
    str单位 = str单位 & ",A.门诊单位,A.门诊包装 AS 门诊系数"
    str单位 = str单位 & ",A.药库单位,A.药库包装 AS 药库系数"
    str单位 = str单位 & ",E.计算单位 AS 售价单位,1 As 售价系数"
    
    '汇总盘点单的SQL
    If bln汇总盘点单 Then
        str汇总盘点单 = "" & _
            " UNION ALL" & _
            " SELECT A.库房ID,A.药品ID,NVL(A.批次, 0) AS 批次,0 AS 实际数量,SUM(A.扣率) 盘点数量," & _
                    " 0 AS 实际金额,0 AS 实际差价,0 AS 可用数量,A.批号,A.产地 ,A.原产地,A.效期,A.批准文号" & _
            " FROM 药品收发记录 A" & _
            " Where A.单据=14 AND A.库房ID=[1] AND A.频次=[3] " & _
            " GROUP BY A.库房ID,A.药品ID,NVL(A.批次, 0),A.批号,A.产地,A.原产地,A.效期,A.批准文号"
    End If
    
    '取药品当前库存及盘点时间以后的净发生额
    gstrSQL = "" & _
        " SELECT DISTINCT A.药品ID,A.成本价 As 平均成本价,E.产地 缺省产地,'[' || E.编码 || ']' As 药品编码, E.名称 As 通用名, C.名称 As 商品名," & _
        "   A.药品来源,A.基本药物,A.药库分批 AS 分批核算,A.药房分批 AS 药房分批核算,E.是否变价,A.加成率," & _
        "   NVL(B.实际金额,0) 实际金额,NVL(B.实际差价,0) 实际差价,D.现价 售价,NVL(B.批次,0) 批次,B.批号,B.效期,F.库房货位,E.规格, decode(b.产地,null,decode(a.上次产地,null,e.产地,a.上次产地),b.产地) as 产地,Decode(b.原产地, Null, a.原产地, b.原产地) As 原产地,A.最大效期," & _
        "   B.批准文号,B.帐面数量,B.盘点数量,B.可用数量" & str单位 & ",Decode(sign(NVL(b.帐面数量,0)), 1,Decode(x.现价,Null,Decode(k.成本价, Null, a.成本价, k.成本价),x.现价), Decode(x.现价,Null,a.成本价,x.现价)) 成本价 " & _
        " FROM 药品规格 A,收费项目目录 E,收费项目别名 C,收费价目 D,药品储备限额 F," & _
        "     (SELECT 库房ID, 药品ID, 批次, SUM (实际数量) AS 帐面数量,SUM (盘点数量) AS 盘点数量,SUM (实际金额) AS 实际金额," & _
        "         SUM (实际差价) AS 实际差价, SUM(可用数量) AS 可用数量,MAX(批号) AS 批号,MAX(产地) AS 产地 ,Max(原产地) As 原产地,MAX(效期) AS 效期,批准文号" & _
        "         From" & _
        "             ( SELECT A.库房ID,A.药品ID,NVL(批次,0) AS 批次,Nvl(A.实际数量,0) 实际数量,0 盘点数量,Nvl(A.实际金额,0) 实际金额,Nvl(A.实际差价,0) 实际差价,Nvl(A.可用数量,0) 可用数量,A.上次批号 AS 批号,A.上次产地 AS 产地,a.原产地,A.效期,A.批准文号" & _
        "             FROM 药品库存 A" & _
        "             Where A.性质 = 1 And A.库房ID=[1] And A.药品ID=[2] " & _
        "             Union All" & _
        "             SELECT A.库房ID,A.药品ID,NVL(A.批次,0) AS 批次,SUM(-1*A.入出系数*A.实际数量*A.付数) AS 实际数量,0 盘点数量," & _
        "             SUM (-1*A.入出系数*A.零售金额) AS 实际金额, SUM(-1*A.入出系数*A.差价) AS 实际差价,0 AS 可用数量,A.批号,A.产地,a.原产地,A.效期,A.批准文号" & _
        "             FROM 药品收发记录 A" & _
        "             Where A.库房ID+0=[1] And A.药品ID+0=[2] " & _
        "             AND A.审核日期 >[4] " & _
        "             GROUP BY A.库房ID, A.药品ID, NVL(A.批次, 0),A.批号,A.产地,a.原产地,A.效期,A.批准文号 " & IIf(Not bln汇总盘点单, "", str汇总盘点单) & _
        "     ) GROUP BY 库房ID, 药品ID, 批次,批准文号) B,(Select x.药品id,x.库房id,x.批次,nvl(x.现价,0) 现价 from 药品价格记录 x where x.价格类型 = 2 and [4] between x.执行日期 and x.终止日期) X," & _
        "      (Select 药品id,批次,平均成本价 成本价 From 药品库存 Where 性质 = 1 And 库房id =[1]) K " & _
        " Where A.药品ID+0=[2] And A.药品ID=E.ID And A.药品ID=B.药品ID" & IIf(bln盘无库存药品, "(+)", "") & _
        " AND A.药品ID=F.药品ID(+) AND F.库房ID(+)=[1] And B.药品id=K.药品id(+) And Nvl(B.批次, 0)=nvl(K.批次(+),0)" & _
        " AND A.药品ID=C.收费细目ID(+) AND C.性质(+)=3 And b.药品id = x.药品id(+) And b.库房id = x.库房id(+) And Nvl(b.批次, 0) = Nvl(x.批次(+), 0) " & _
        " AND A.药品ID=D.收费细目ID(+) AND D.执行日期(+)<=SYSDATE AND NVL(D.终止日期(+),SYSDATE)>=SYSDATE"
    gstrSQL = gstrSQL & GetPriceClassString("D")
    gstrSQL = gstrSQL & " and e.建档时间 <= [4] "
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品当前库房所有批次明细记录]", lng库房ID, lng药品ID, str盘点时间, CDate(Format(str盘点时间, "yyyy-mm-dd hh:mm:ss")))
    
    Set GetPhysicDetail = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Double
    
    With vsfBill
        Select Case Col
            Case mconIntCol大单位数量, mconIntCol小单位数量
                '显示合计数量
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                If .EditText <> "" Then .TextMatrix(Row, Col) = Val(.EditText)
                If .Col = mconIntCol大单位数量 Then
                    lngSum = Val(.TextMatrix(Row, mconIntCol小单位数量)) + Val(.TextMatrix(Row, mconIntCol大单位数量)) * Val(Split(.TextMatrix(Row, mconIntCol比例系数), "|")(0)) / Val(Split(.TextMatrix(Row, mconIntCol比例系数), "|")(1))
                Else
                    lngSum = Val(.TextMatrix(Row, mconIntCol小单位数量)) + Val(.TextMatrix(Row, mconIntCol大单位数量)) * Val(Split(.TextMatrix(Row, mconIntCol比例系数), "|")(0)) / Val(Split(.TextMatrix(Row, mconIntCol比例系数), "|")(1))
                End If
                .TextMatrix(Row, mconintCol数量_合计) = zlStr.FormatEx(lngSum * Val(Split(.TextMatrix(Row, mconIntCol比例系数), "|")(1)), mintNumberDigit1, , True)
            Case mconintCol成本价
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                .EditText = zlStr.FormatEx(Val(.EditText), mintCostDigit, , True)
                .TextMatrix(Row, Col) = .EditText
        End Select
    End With
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
    
        gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, Nvl(实际金额,0) / 实际数量), 零售价) as 零售价 from 药品库存 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 性质 = 1"
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

Private Function CheckTime() As Boolean
    '检查是否有药品建档时间大于盘点时间
    Dim rs建档时间 As Recordset
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errHandle
    CheckTime = False
    
    With vsfBill
        For i = 1 To .rows - 1
            If .TextMatrix(i, 0) <> "" Then
                strSQL = "Select a.建档时间 From 收费项目目录 A Where a.Id =[1]"
                Set rs建档时间 = zlDataBase.OpenSQLRecord(strSQL, "查询建档时间", .TextMatrix(i, 0))
                If Format(rs建档时间!建档时间, "yyyy-MM-dd HH:mm:ss") > Format(TxtCheckDate.Text, "yyyy-MM-dd HH:mm:ss") Then
                    MsgBox "第" & i & "行药品为盘点时间后建立，请检查！", vbInformation + vbOKOnly, gstrSysName
                    .Row = i
                    Exit Function
                End If
            End If
        Next
        
    End With
    
    CheckTime = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
