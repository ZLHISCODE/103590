VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiffPriceRecal 
   Caption         =   "药品初始结存"
   ClientHeight    =   7305
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   11760
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   5
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   10410
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
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
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   23
         Top             =   4500
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   22
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   21
         Top             =   4500
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   20
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结存库房"
         Height          =   180
         Left            =   270
         TabIndex        =   19
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品初始结存"
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
         TabIndex        =   18
         Top             =   45
         Width           =   11535
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   4155
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   14
         Top             =   4440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   13
         Top             =   4440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额差合计："
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9600
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   8640
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "盘点金额合计："
         Height          =   180
         Left            =   1920
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8730
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7410
      TabIndex        =   3
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   5100
      Width           =   1815
   End
   Begin VB.CommandButton cmd固定列 
      Caption         =   "固定列(&L)"
      Height          =   350
      Left            =   6090
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1100
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
            Picture         =   "frmDiffPriceRecal.frx":0000
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":021A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0434
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":064E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0868
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0A82
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0C9C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0EB6
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
            Picture         =   "frmDiffPriceRecal.frx":10D0
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":12EA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1504
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":171E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1938
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1B52
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1D6C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1F86
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   6945
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiffPriceRecal.frx":21A0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
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
      Caption         =   "查找药品"
      Height          =   180
      Left            =   1530
      TabIndex        =   25
      Top             =   5145
      Width           =   720
   End
End
Attribute VB_Name = "frmDiffPriceRecal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSelectStock As String           '是否可选库房
Private mint编辑状态 As Integer             '1、初始结存；2、手工录入差价；
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean                '第一次显示
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑

Private Const mlngColorRed As Long = vbRed
Private Const mlngColorBlue As Long = vbBlue
Private Const mlngColorBlack As Long = vbBlack
Private mlngCurrColor As Long
Private mlngNextColor As Long
Private blnColorRefresh As Boolean

Private mstrMsg As String

Private mlng库房 As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数
Private mintMaxMoneyBit As Integer          '药品库存表中金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol序号 As Integer = 4
Private Const mconIntCol规格 As Integer = 5
Private Const mconIntCol售价单位 As Integer = 6
Private Const mconIntCol小单位数量 As Integer = 7
Private Const mconIntCol药库单位 As Integer = 8
Private Const mconIntCol大单位数量 As Integer = 9
Private Const mconIntCol结存金额 As Integer = 10
Private Const mconIntCol结存差价 As Integer = 11
Private Const mconIntCol实际差价 As Integer = 12
Private Const mconIntColS  As Integer = 13              '总列数

Private Function GetAllDrug() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str药名 As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct A.药品id, '[' || E.编码 || ']' As 药品编码, E.名称 As 通用名, C.名称 As 商品名, E.规格, E.计算单位 As 售价单位, S.结存数量 As 小包装数量," & _
        " A.药库单位, S.结存数量 / A.药库包装 As 大包装数量, S.结存金额, S.结存差价 " & _
        " From 药品规格 A, 收费项目目录 E, 收费项目别名 C, " & _
        " (Select 药品id, Sum(实际数量) 结存数量, Sum(实际金额) 结存金额, Sum(实际差价) 结存差价 " & _
        " From 药品结存 Where Nvl(是否初始,0) = 1 " & _
        " Group By 药品id) S " & _
        " Where A.药品id = E.ID And A.药品id = C.收费细目id(+) And C.性质(+) = 3 And A.药品id = S.药品id " & _
        " Order By 药品编码"
        
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取所有结存药品]")
    
    If rsTmp.EOF Then
        GetAllDrug = False
        Exit Function
    End If
    
    Call initGrid
    
    Call FS.StopFlash
    
    With mshBill
        .Redraw = False
        Do While Not rsTmp.EOF
            .TextMatrix(.rows - 1, 0) = rsTmp!药品id
            
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                str药名 = rsTmp!通用名
            Else
                str药名 = IIf(IsNull(rsTmp!商品名), rsTmp!通用名, rsTmp!商品名)
            End If
           
            .TextMatrix(.rows - 1, mconIntCol药名) = rsTmp!药品编码 & str药名
          
            .TextMatrix(.rows - 1, mconIntCol商品名) = IIf(IsNull(rsTmp!商品名), "", rsTmp!商品名)
                    
            .TextMatrix(.rows - 1, mconIntCol规格) = rsTmp!规格
            .TextMatrix(.rows - 1, mconIntCol售价单位) = rsTmp!售价单位
            .TextMatrix(.rows - 1, mconIntCol小单位数量) = rsTmp!小包装数量
            .TextMatrix(.rows - 1, mconIntCol药库单位) = rsTmp!药库单位
            .TextMatrix(.rows - 1, mconIntCol大单位数量) = rsTmp!大包装数量
            .TextMatrix(.rows - 1, mconIntCol结存金额) = zlStr.FormatEx(rsTmp!结存金额, gtype_UserSysParms.P9_费用金额保留位数)
            .TextMatrix(.rows - 1, mconIntCol结存差价) = zlStr.FormatEx(rsTmp!结存差价, gtype_UserSysParms.P9_费用金额保留位数)

            Call zlControl.StaShowPercent(rsTmp.AbsolutePosition / rsTmp.RecordCount, staThis.Panels(2), frmDiffPriceRecal)
            rsTmp.MoveNext
            If Not rsTmp.EOF Then .rows = .rows + 1
        Loop
        Call RefreshRowNO(mshBill, mconIntCol行号, 1)
        .Redraw = True
    End With
    
    DoEvents
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    
    GetAllDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal int编辑状态 As Integer)
    mblnSave = False
    mblnSuccess = False
    mint编辑状态 = int编辑状态
    mblnChange = False
    mblnFirst = True
        
    Set mfrmMain = FrmMain
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption & IIf(mint编辑状态 = 2, "(差价录入)", "")
    
    Me.Show vbModal, FrmMain
    
End Sub
'初始化编辑控件
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol售价单位) = "售价单位"
        .TextMatrix(0, mconIntCol小单位数量) = "小包装数量"
        .TextMatrix(0, mconIntCol药库单位) = "药库单位"
        .TextMatrix(0, mconIntCol大单位数量) = "大包装数量"
        .TextMatrix(0, mconIntCol结存金额) = "结存金额"
        .TextMatrix(0, mconIntCol结存差价) = "结存差价"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 500
        .ColWidth(mconIntCol药名) = 3000
        
        '商品名列处理
        If gint药品名称显示 = 2 Then
            '显示商品名列
            .ColWidth(mconIntCol商品名) = 2000
        Else
            '不单独显示商品名列
            .ColWidth(mconIntCol商品名) = 0
        End If
        
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol售价单位) = 800
        .ColWidth(mconIntCol小单位数量) = 1000
        .ColWidth(mconIntCol药库单位) = 800
        .ColWidth(mconIntCol大单位数量) = 1000
        .ColWidth(mconIntCol结存金额) = 1000
        .ColWidth(mconIntCol结存差价) = 1000
        .ColWidth(mconIntCol实际差价) = 1000
        
        
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol药名) = 5
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol售价单位) = 5
        .ColData(mconIntCol小单位数量) = 5
        .ColData(mconIntCol药库单位) = 5
        .ColData(mconIntCol大单位数量) = 5
        .ColData(mconIntCol结存金额) = 5
        .ColData(mconIntCol结存差价) = 5
        .ColData(mconIntCol实际差价) = 4
                
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol售价单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol小单位数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol药库单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol大单位数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol结存金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol结存差价) = flexAlignRightCenter
        .ColAlignment(mconIntCol实际差价) = flexAlignRightCenter
                
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    MsgBox "对不起，暂时没有帮助！"
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    Dim dbl差价 As Double
    Dim lng药品ID As Long
    Dim strTmp As String
    Dim intDrugCount As Integer
    
    On Error GoTo errHandle
    
    gcnOracle.BeginTrans
        
    With mshBill
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, 0)) > 0 And Val(.TextMatrix(lngRow, mconIntCol实际差价)) > 0 Then
                intDrugCount = intDrugCount + 1
                lng药品ID = Val(.TextMatrix(lngRow, 0))
                
                '按最小包装计算差价
                dbl差价 = Round(Val(.TextMatrix(lngRow, mconIntCol实际差价)) / Val(.TextMatrix(lngRow, mconIntCol小单位数量)), 7)
                
                strTmp = IIf(strTmp = "", "", strTmp & "|")
                strTmp = strTmp & lng药品ID & "," & dbl差价
                
                If intDrugCount > 99 Then
                    gstrSQL = "Zl_药品结存_Update('" & strTmp & "' )"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
                    strTmp = ""
                    intDrugCount = 0
                End If
            End If
        Next
        If strTmp <> "" Then
            gstrSQL = "Zl_药品结存_Update('" & strTmp & "' )"
            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End With
    
    gcnOracle.CommitTrans
    MsgBox "差价保存完毕！", vbInformation + vbOKOnly, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Activate()
    mshBill.ClearBill
    If GetAllDrug = False Then
        Exit Sub
        Unload Me
    End If
End Sub

'=========================================================================================
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
    
    txtCheckDate.Left = mshBill.Left + mshBill.Width - txtCheckDate.Width
    lblCheckDate.Left = txtCheckDate.Left - lblCheckDate.Width - 100
    
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
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = Pic单据.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic单据.TextWidth(lblCheckSum.Caption) + 200
    End With
    
    With mshBill
        .Height = Pic单据.Height - .Top - 100
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
        
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
    With cmd固定列
        .Left = CmdSave.Left - .Width - 150
        .Top = CmdSave.Top
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        Select Case .Col
            Case mconIntCol实际差价
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select

    End With
End Sub


Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        Select Case .Col
            Case mconIntCol实际差价
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，差价金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Abs(Val(strkey)) < 0.00001 Then
                        MsgBox "对不起，差价金额的绝对值必须不小于0.00001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "差价金额必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = zlStr.FormatEx(strkey, gtype_UserSysParms.P9_费用金额保留位数, , True)
                    .Text = strkey
                    
                End If
        End Select
    End With
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, True
    End If
End Sub


