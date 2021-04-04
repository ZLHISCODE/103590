VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCheckCard 
   Caption         =   "卫材盘点表"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCheckCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd固定列 
      Caption         =   "固定列(&L)"
      Height          =   350
      Left            =   6090
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7425
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8730
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   210
         TabIndex        =   1
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
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCheckCostSum 
         AutoSize        =   -1  'True
         Caption         =   "盘点成本金额合计："
         Height          =   180
         Left            =   3960
         TabIndex        =   29
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "盘点金额合计："
         Height          =   180
         Left            =   1920
         TabIndex        =   26
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   8640
         TabIndex        =   24
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9600
         TabIndex        =   23
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额差合计："
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   19
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   17
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Top             =   4440
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
         TabIndex        =   15
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
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料盘点表"
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
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点库房"
         Height          =   180
         Left            =   270
         TabIndex        =   0
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
            Picture         =   "frmCheckCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1000
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
            Picture         =   "frmCheckCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
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
            Picture         =   "frmCheckCard.frx":22EA
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
            Picture         =   "frmCheckCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCard.frx":3080
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
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu PopMenu 
      Caption         =   "固定列"
      Visible         =   0   'False
      Begin VB.Menu mnuFirst 
         Caption         =   "从材料信息到单位列(&1)"
      End
      Begin VB.Menu mnuSecond 
         Caption         =   "从材料信息到效期列(&2)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "恢复(&D)"
      End
   End
End
Attribute VB_Name = "frmCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、汇总盘点记录单,产生盘点表;6、全部盘为零
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

Private mint库存检查 As Integer             '表示卫生材料出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Dim mstrPrivs As String                     '权限
Private Const mstrCaption As String = "卫材盘点表"
Private mstr重复卫材 As String '记录重复的卫材

Private recSort As ADODB.Recordset          '按药品ID、批次排序的专用记录集

'刘兴宏:2007/06/10
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private Const mlngModule = 1719
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mstr盘点单号 As String  '以NO,NO为分隔
Private mbln只统计盘点单卫材 As Boolean
Private mbln删除盘点单 As Boolean
Private mbln盘无存储库房材料 As Boolean
Private mbln分批卫材批号产地控制 As Boolean  '是否检查分批卫材批号产地是否录入

'=========================================================================================
Private Enum mBillCol
     C_行号 = 1
     C_材料 = 2
     C_序号 = 3
     c_规格 = 4
     C_批次 = 5
     C_可用数量 = 6
     c_比例系数 = 7
     C_指导差价率 = 8
     C_实际差价 = 9
     C_实际金额 = 10
     C_产地 = 11
     C_批准文号 = 12
     C_库房货位 = 13
     c_单位 = 14
     c_批号 = 15
     C_效期 = 16
     C_帐面数量 = 17
     C_实盘数量 = 18
     C_标志 = 19
     C_数量差 = 20
     C_成本价 = 21
     C_售价 = 22
     c_金额差 = 23
     c_差价差 = 24
     C_盘点金额 = 25
     C_盘点成本金额 = 26
     C_盘点成本金额差 = 27
     c_新批次 = 28
     c_批号编辑 = 29
     c_产地编辑 = 30
     C_Cols = 31               '总列数
End Enum

'=========================================================================================


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    GetDepend = False
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID " & _
        "           AND A.单据 = [1]  and b.系数=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "卫生材料盘点管理", 37, 1)
    
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫生材料盘点表的入库类别，请在入出分类中设置！"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID " & _
        "           AND A.单据 = [1]  and b.系数=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "卫生材料盘点管理", 37, -1)

    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫生材料盘点表的出库类别，请在入出分类中设置！"
        rsTemp.Close
        Exit Function
    End If
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, _
    Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:编辑单据或显示单据,是单据的唯一入口
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    Call GetRegInFor(g私有模块, "卫材盘点管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint编辑状态 = 1 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Then
        mblnEdit = True
        If mbln单据增加 Then
            'mstr单据号 = NextNo(75)
        End If

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
        If mint编辑状态 = 4 Then
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
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
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmd固定列_Click()
    Call PopupMenu(PopMenu, 2)
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


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    '设置排序数据集
    Call SetSortRecord
    
    If mint编辑状态 = 5 Then    '汇总产生盘点表
        If ValidData = False Then Exit Sub
        blnSuccess = SaveCard
        
        If blnSuccess Then
            '对它进行审核
'            If SaveCheck Then
'                strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
'                If Val(strReg) = 1 Then
'                    '打印
'                    If InStr(mstrPrivs, "单据打印") <> 0 Then
'                        printbill
'                    End If
'                End If
'            End If
            strReg = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        
        mstrTime_End = GetBillInfo(22, mstr单据号)
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
        
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
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
'
'    If mbln单据增加 Then
'        'mstr单据号 = NextNo(75)
'        txtNO = mstr单据号
'    End If
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    txt摘要.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
End Sub

Private Sub Form_Activate()
    
    Dim str分类ID As String, lng库房id As Long, int盘点方式 As Integer, str盘点时间 As String, str库房货位 As String
    Dim int盘无库存材料 As Integer, bln盘点零数量且有金额 As Boolean
    
    If mblnFirst = False Then Exit Sub
        
    mint库存检查 = Get出库检查(lng库房id)
    mintBatchNoLen = GetBatchNoLen()
    
    If mintParallelRecord <> 1 Then mblnChange = False
    
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
            MsgBox "还存在未审核的卫生材料单据，请全部审核后再试！", vbOKOnly, gstrSysName
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
    
    mblnFirst = False
    '初始化变量
    str分类ID = ""
    
    If mint编辑状态 = 1 Then
        '自动搜索或手工输入盘点表
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str分类ID, lng库房id, int盘点方式, str盘点时间, int盘无库存材料, bln盘点零数量且有金额, str库房货位) = True Then
            If str分类ID <> "" Then
                If str分类ID = "所有卫生材料" Then
                    str分类ID = ""
                End If
                Call SearchData(str分类ID, lng库房id, int盘点方式, str盘点时间, int盘无库存材料, bln盘点零数量且有金额, str库房货位)
            End If
        Else
            Unload Me
            Exit Sub
        End If
        
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    ElseIf mint编辑状态 = 5 Then
        '产生盘点表（汇总指定时刻的盘点记录单与指定时刻的库存）
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng库房id, str盘点时间, mstr盘点单号, mbln只统计盘点单卫材, mbln删除盘点单) = True Then
            Call SearchTableData(lng库房id, str盘点时间)
        
        Else
            Unload Me
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    ElseIf mint编辑状态 = 6 Then
        '全部盘为零
        str盘点时间 = Format(sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        txtCheckDate = str盘点时间
        txtStock.Caption = mfrmMain.cboStock.Text
        lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng库房id
        
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        
        Call SearchTableData(lng库房id, str盘点时间)
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    End If
End Sub

Private Sub SearchData(ByVal str分类ID As String, ByVal lng库房id As Long, _
    ByVal int盘点方式 As Integer, ByVal str盘点时间 As String, ByVal int盘无库存材料 As Integer, _
    ByVal bln盘点零数量且有金额 As Boolean, ByVal str库房货位 As String)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:根据条件，获取相关数据
    '--入参数:str分类ID-分类ID(1,2)
    '         lng库房ID-库房id
    '         int盘点方式:日盘,月盘...
    '         str盘点时间-盘点日期
    '         int盘无库存材料-包含盘点无库存数量的材料
    '         bln盘点零数量且有金额-仅仅盘点无库存数量但有金额的卫生材料
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

    
    Dim rsData As ADODB.Recordset '库存记录集
    Dim rsTemp As ADODB.Recordset
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim lngRecordCount As Long
    Dim dbl成本价 As Double
    Dim bln库房 As Boolean
    Dim rsprice As New Recordset
    Dim strMoneyDigit As String
    Dim dbl金额差, dbl差价差 As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    '设置界面显示内容
    Select Case int盘点方式
        Case 1
            stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行日盘点"
        Case 2
            stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行周盘点"
        Case 3
            stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行月盘点"
        Case 4
            stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行季度盘点"
        Case 5
            stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行忽略盘点方式盘点"
    End Select
    
  
    Call FS.ShowFlash("正在计算卫生材料库存数据,请稍候 ...", Me)

    DoEvents    ': Me.Refresh
    Set rsData = GetDateStock(str盘点时间, lng库房id, int盘点方式, IIf(int盘无库存材料 = 0, False, True), , str分类ID, , bln盘点零数量且有金额, str库房货位)
    
    Call FS.StopFlash    ': Me.Refresh
    
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then
        If mint编辑状态 = 6 Then
            ShowMsgBox "未能正确读取卫生材料库存数据,请重试！": Exit Sub
        Else
            ShowMsgBox "未能正确读取卫生材料库存数据,请重试或手工输入卫生材料！": Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("正在装入卫生材料数据,请稍候 ...", Me)
    DoEvents: 'Me.Refresh
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1
    bln库房 = CheckPartProp(lng库房id)
    
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .Rows = .Rows + 1
            .TextMatrix(i, 0) = rsData!材料ID
            
            '取该材料的成本价（当库存数量为零且是时价材料时，用于计算差价）
'            gstrSQL = "Select Nvl(成本价,0) 成本价 From 材料特性 Where 材料ID=[1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--取该卫生材料的成本价", Val(NVL(rsData!材料ID)))
'
            dbl成本价 = Val(zlStr.NVL(rsData!最后进价)) ' rsTemp!成本价
            
            '时价材料重算售价
            If rsData!是否变价 = 1 Then
                .TextMatrix(i, mBillCol.C_售价) = Format(Get零售价(Val(zlStr.NVL(rsData!材料ID)), Val(txtStock.Tag), Val(zlStr.NVL(rsData!批次)), rsData!比例系数), mFMT.FM_零售价)
            Else
                .TextMatrix(i, mBillCol.C_售价) = Format(IIf(IsNull(rsData!售价), 0, rsData!售价), mFMT.FM_零售价)
            End If
           
            .TextMatrix(i, mBillCol.C_材料) = "[" & rsData!编码 & "]" & rsData!商品名称
            .TextMatrix(i, mBillCol.c_规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
            .TextMatrix(i, mBillCol.C_产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
            .TextMatrix(i, mBillCol.C_批准文号) = IIf(IsNull(rsData!批准文号), "", rsData!批准文号)
            .TextMatrix(i, mBillCol.C_库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
            .TextMatrix(i, mBillCol.c_单位) = IIf(IsNull(rsData!单位), "", rsData!单位)
            .TextMatrix(i, mBillCol.c_批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
            
            '如果是分批材料，将批次改填为-1，表示新增批次
            .TextMatrix(i, mBillCol.C_批次) = IIf(IsNull(rsData!批次), "", rsData!批次)
            
            If Val(.TextMatrix(i, mBillCol.C_批次)) <> 0 Then
                .TextMatrix(i, mBillCol.c_批号编辑) = rsData!批号编辑
                .TextMatrix(i, mBillCol.c_产地编辑) = rsData!产地编辑
            End If
            
            If CheckPhysicBatch(bln库房, rsData!库房分批, rsData!在用分批) And Val(.TextMatrix(i, mBillCol.C_批次)) = 0 Then
                .TextMatrix(i, mBillCol.C_批次) = -1
            End If
            If Val(.TextMatrix(i, mBillCol.C_批次)) = -1 Then
                .TextMatrix(i, mBillCol.C_成本价) = Format(dbl成本价, mFMT.FM_成本价)
            Else
                .TextMatrix(i, mBillCol.C_成本价) = Format(Val(zlStr.NVL(rsData!成本价)), mFMT.FM_成本价)
            End If
            .TextMatrix(i, mBillCol.C_效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
            .TextMatrix(i, mBillCol.C_帐面数量) = Format(Val(zlStr.NVL(rsData!帐面数量)), mFMT.FM_数量)
            .TextMatrix(i, mBillCol.C_实盘数量) = .TextMatrix(i, mBillCol.C_帐面数量)
            .TextMatrix(i, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(i, mBillCol.C_实盘数量)) * Val(.TextMatrix(i, mBillCol.C_售价)), mFMT.FM_金额)
            .TextMatrix(i, mBillCol.C_可用数量) = rsData!可用数量
            .TextMatrix(i, mBillCol.C_实际金额) = rsData!实际金额
            .TextMatrix(i, mBillCol.C_实际差价) = rsData!实际差价
            .TextMatrix(i, mBillCol.c_比例系数) = rsData!比例系数
            
            .TextMatrix(i, mBillCol.C_指导差价率) = rsData!指导差价率 & "||" & rsData!是否变价 & "||" & rsData!在用分批
            .TextMatrix(i, mBillCol.C_标志) = "平"
            .TextMatrix(i, mBillCol.C_数量差) = Format("0", mFMT.FM_数量)
            
            If Val(.TextMatrix(i, mBillCol.C_帐面数量)) = 0 Then
                strMoneyDigit = "#0.00000"
            Else
                strMoneyDigit = mFMT.FM_金额
            End If
             
             '金额差=当前售价*实盘数量-实际金额
             '差价差=金额差*iif(实际金额<=0,指导差价率,(实际差价/实际金额))
            .TextMatrix(i, mBillCol.c_金额差) = Format(Val(.TextMatrix(i, mBillCol.C_售价)) * Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_实际金额)), strMoneyDigit)
            .TextMatrix(i, mBillCol.c_差价差) = Format((Val(.TextMatrix(i, mBillCol.C_售价)) - Val(.TextMatrix(i, mBillCol.C_成本价))) * Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_实际差价)), strMoneyDigit)
            dbl金额差 = Val(.TextMatrix(i, mBillCol.c_金额差))
            dbl差价差 = Val(.TextMatrix(i, mBillCol.c_差价差))
            
            .TextMatrix(i, mBillCol.C_盘点成本金额) = Format(Val(.TextMatrix(i, mBillCol.C_实际金额)) + dbl金额差 - (Val(.TextMatrix(i, mBillCol.C_实际差价)) + dbl差价差), mFMT.FM_金额)
            .TextMatrix(i, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(i, mBillCol.c_金额差)) - Val(.TextMatrix(i, mBillCol.c_差价差)), mFMT.FM_金额)
            Call ShowPercent(i / lngRecordCount)
            i = i + 1
nextloop:
            rsData.MoveNext
        Loop
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        .Redraw = True
    End With
    Call FS.StopFlash
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1: mshBill.Col = mBillCol.C_实盘数量
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SearchTableData(ByVal lng库房id As Long, ByVal str盘点时间 As String)
    Dim rsData As ADODB.Recordset '卫生材料库存记录集
    Dim rsTemp As ADODB.Recordset
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim lngRecordCount As Long
    Dim sinPrice As Single
    Dim dbl成本价 As Double
    Dim lngPhysic As Long
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsprice As New Recordset
    Dim str盘点单NO串 As String
    Dim strMoneyDigit As String
    Dim dbl金额差, dbl差价差 As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    
    Call FS.ShowFlash("正在计算卫生材料库存数据,请稍候 ...", Me)

    DoEvents
    
    If mint编辑状态 = 5 Then
        Set rsData = Get汇总记录单(lng库房id, str盘点时间)
    Else 'mint编辑状态 = 6（只有5、6才调用了该过程）
        Set rsData = GetDateStock(str盘点时间, lng库房id, 0, False, IIf(mint编辑状态 = 5, True, False))
    End If
    Call FS.StopFlash
    
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then
        If mint编辑状态 = 6 Then
            ShowMsgBox "未能正确读取卫生材料库存数据,请重试！": Exit Sub
        Else
            ShowMsgBox "未能正确读取卫生材料库存数据,请重试或手工输入材料！": Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("正在装入材料数据,请稍候 ...", Me)
    DoEvents
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1: lngPhysic = 0
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .Rows = .Rows + 1
            '如果材料ID不同，如果是时价材料，则取实际的零售价
            .TextMatrix(i, 0) = rsData!材料ID
            lngPhysic = rsData!材料ID
            sinPrice = IIf(rsData!是否变价 = 1, 0, IIf(IsNull(rsData!售价), 0, rsData!售价))
            
            
            dbl成本价 = Val(zlStr.NVL(rsData!最后进价))
            
            '如果是时价材料，重算其售价
            If rsData!是否变价 = 1 Then
                sinPrice = Get零售价(Val(zlStr.NVL(rsData!材料ID)), lng库房id, Val(zlStr.NVL(rsData!批次)), rsData!比例系数)
                .TextMatrix(i, mBillCol.C_售价) = Format(sinPrice, mFMT.FM_零售价)
            Else
                .TextMatrix(i, mBillCol.C_售价) = Format(sinPrice, mFMT.FM_零售价)
            End If
            
            If (rsData!批次 = -1) Then
                '表示初始化操作，可以打开记录单
                Select Case mintUnit
                    Case 0
                        strUnitQuantity = ",Sum(A.扣率) AS 盘点数量"
                    Case Else
                        strUnitQuantity = ",Sum(A.扣率/b.换算系数) AS 盘点数量"
                End Select
                
                str盘点单NO串 = Replace(mstr盘点单号, "'", "")
                
                gstrSQL = "" & _
                    "   Select /*+rule*/ Nvl(A.批次,0) 批次,A.批号,A.效期,A.产地,A.单量 成本价" & strUnitQuantity & _
                    "   From 药品收发记录 A,材料特性 B,Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) C" & _
                    "   Where A.药品ID+0=[1]" & " And Nvl(A.批次,0)=-1 " & _
                    "           And A.NO=C.Column_Value And A.单据=23 And A.药品ID=B.材料ID" & _
                    "   Group By Nvl(批次,0),批号,效期,产地,序号,A.单量"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--读取记录单批次数据", lngPhysic, str盘点时间, str盘点单NO串)
                
                Do While Not rsTemp.EOF
'                    If i > 1 Then .Rows = .Rows + 1
                    If rsTemp.AbsolutePosition > 1 Then .Rows = .Rows + 1 '加载第一条不需要.rows+1
                    .TextMatrix(i, 0) = rsData!材料ID
                    .TextMatrix(i, mBillCol.C_售价) = Format(sinPrice, mFMT.FM_零售价)
                    .TextMatrix(i, mBillCol.C_材料) = "[" & rsData!编码 & "]" & rsData!商品名称
                    .TextMatrix(i, mBillCol.c_规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
                    .TextMatrix(i, mBillCol.C_产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(i, mBillCol.C_批准文号) = IIf(IsNull(rsData!批准文号), "", rsData!批准文号)
                    .TextMatrix(i, mBillCol.C_库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
                    .TextMatrix(i, mBillCol.c_单位) = IIf(IsNull(rsData!单位), "", rsData!单位)
                    .TextMatrix(i, mBillCol.c_批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(i, mBillCol.C_批次) = IIf(IsNull(rsData!批次), "", rsData!批次)
                    .TextMatrix(i, mBillCol.C_效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd"))
                    .TextMatrix(i, mBillCol.C_帐面数量) = Format(Val(zlStr.NVL(rsData!帐面数量)), mFMT.FM_数量)
                    .TextMatrix(i, mBillCol.C_实盘数量) = Format(IIf(IsNull(rsTemp!盘点数量), 0, rsTemp!盘点数量), mFMT.FM_数量)
                    .TextMatrix(i, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(i, mBillCol.C_实盘数量)) * Val(.TextMatrix(i, mBillCol.C_售价)), mFMT.FM_金额)
                    .TextMatrix(i, mBillCol.C_可用数量) = rsData!可用数量
                    .TextMatrix(i, mBillCol.C_实际金额) = rsData!实际金额
                    .TextMatrix(i, mBillCol.C_实际差价) = rsData!实际差价
                    .TextMatrix(i, mBillCol.c_比例系数) = rsData!比例系数
                    .TextMatrix(i, mBillCol.C_指导差价率) = rsData!指导差价率 & "||" & rsData!是否变价 & "||" & rsData!在用分批
                    .TextMatrix(i, mBillCol.C_成本价) = Format(Val(rsTemp!成本价) * Val(rsData!比例系数), mFMT.FM_成本价)
                    
                    If Val(.TextMatrix(i, mBillCol.C_帐面数量)) > Val(.TextMatrix(i, mBillCol.C_实盘数量)) Then
                        .TextMatrix(i, mBillCol.C_标志) = "亏"
                    ElseIf Val(.TextMatrix(i, mBillCol.C_帐面数量)) < Val(.TextMatrix(i, mBillCol.C_实盘数量)) Then
                        .TextMatrix(i, mBillCol.C_标志) = "盈"
                    Else
                        .TextMatrix(i, mBillCol.C_标志) = "平"
                    End If
                    .TextMatrix(i, mBillCol.C_数量差) = Format(Abs(Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_帐面数量))), mFMT.FM_数量)
                    
                    If Val(.TextMatrix(i, mBillCol.C_帐面数量)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_金额
                    End If
                     '金额差=当前售价*实盘数量-实际金额
                    '差价差=金额差*iif(实际金额<=0,指导差价率,(实际差价/实际金额))
                    .TextMatrix(i, mBillCol.c_金额差) = Format(Val(.TextMatrix(i, mBillCol.C_售价)) * Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_实际金额)), strMoneyDigit)
'                    If rsData!是否变价 = 1 And Val(.TextMatrix(i, mBillCol.C_帐面数量)) = 0 Then
'                        .TextMatrix(i, mBillCol.C_差价差) = Format(Val(.TextMatrix(i, mBillCol.C_数量差)) * (Val(.TextMatrix(i, mBillCol.C_售价)) - dbl成本价 * rsData!比例系数), strMoneyDigit)
'                    Else
                        .TextMatrix(i, mBillCol.c_差价差) = Format(Val(.TextMatrix(i, mBillCol.C_实盘数量)) * (Val(.TextMatrix(i, mBillCol.C_售价)) - Val(.TextMatrix(i, mBillCol.C_成本价))) - Val(.TextMatrix(i, mBillCol.C_实际差价)), strMoneyDigit)
'                    End If
                    
                    dbl金额差 = .TextMatrix(i, mBillCol.c_金额差)
                    dbl差价差 = .TextMatrix(i, mBillCol.c_差价差)
                    
                    If .TextMatrix(i, mBillCol.C_标志) = "亏" Then
                        '保证实际金额与金额差相同的符号（因为入出系数为-1，这样就能保证完全冲销为零）
                        If Not 相同符号(Val(.TextMatrix(i, mBillCol.c_金额差)), Val(.TextMatrix(i, mBillCol.C_实际金额))) Then
                            .TextMatrix(i, mBillCol.c_金额差) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_金额差)), strMoneyDigit)
                        End If
                        If Not 相同符号(Val(.TextMatrix(i, mBillCol.c_差价差)), Val(.TextMatrix(i, mBillCol.C_实际差价))) Then
                            .TextMatrix(i, mBillCol.c_差价差) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_差价差)), strMoneyDigit)
                        End If
                    End If
                    .TextMatrix(i, mBillCol.C_盘点成本金额) = Format(Val(.TextMatrix(i, mBillCol.C_实际金额)) + dbl金额差 - (Val(.TextMatrix(i, mBillCol.C_实际差价)) + dbl差价差), mFMT.FM_金额)
                    .TextMatrix(i, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(i, mBillCol.c_金额差)) - Val(.TextMatrix(i, mBillCol.c_差价差)), mFMT.FM_金额)
                    
                    i = i + 1
                    rsTemp.MoveNext
                Loop
                i = i - 1
            Else
                .TextMatrix(i, mBillCol.C_材料) = "[" & rsData!编码 & "]" & rsData!商品名称
                .TextMatrix(i, mBillCol.c_规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
                .TextMatrix(i, mBillCol.C_产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
                .TextMatrix(i, mBillCol.C_批准文号) = IIf(IsNull(rsData!批准文号), "", rsData!批准文号)
                .TextMatrix(i, mBillCol.C_库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
                .TextMatrix(i, mBillCol.c_单位) = IIf(IsNull(rsData!单位), "", rsData!单位)
                .TextMatrix(i, mBillCol.c_批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
                .TextMatrix(i, mBillCol.C_批次) = IIf(IsNull(rsData!批次), "", rsData!批次)
                
                If Val(.TextMatrix(i, mBillCol.C_批次)) <> 0 Then
                    .TextMatrix(i, mBillCol.c_批号编辑) = rsData!批号编辑
                    .TextMatrix(i, mBillCol.c_产地编辑) = rsData!产地编辑
                End If
                
                .TextMatrix(i, mBillCol.C_效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
                .TextMatrix(i, mBillCol.C_帐面数量) = Format(IIf(IsNull(rsData!帐面数量), 0, rsData!帐面数量), mFMT.FM_数量)
                If mint编辑状态 = 5 Then
                    .TextMatrix(i, mBillCol.C_实盘数量) = Format(IIf(IsNull(rsData!盘点数量), 0, rsData!盘点数量), mFMT.FM_数量)
                Else
                    .TextMatrix(i, mBillCol.C_实盘数量) = Format(0, mFMT.FM_数量)
                End If
                .TextMatrix(i, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(i, mBillCol.C_实盘数量)) * Val(.TextMatrix(i, mBillCol.C_售价)), mFMT.FM_金额)
                .TextMatrix(i, mBillCol.C_可用数量) = rsData!可用数量
                .TextMatrix(i, mBillCol.C_实际金额) = rsData!实际金额
                .TextMatrix(i, mBillCol.C_实际差价) = rsData!实际差价
                .TextMatrix(i, mBillCol.c_比例系数) = rsData!比例系数
                .TextMatrix(i, mBillCol.C_成本价) = Format(Val(zlStr.NVL(rsData!成本价)), mFMT.FM_成本价)
                
                .TextMatrix(i, mBillCol.C_指导差价率) = rsData!指导差价率 & "||" & rsData!是否变价 & "||" & rsData!在用分批
                If Val(.TextMatrix(i, mBillCol.C_帐面数量)) > Val(.TextMatrix(i, mBillCol.C_实盘数量)) Then
                    .TextMatrix(i, mBillCol.C_标志) = "亏"
                ElseIf Val(.TextMatrix(i, mBillCol.C_帐面数量)) < Val(.TextMatrix(i, mBillCol.C_实盘数量)) Then
                    .TextMatrix(i, mBillCol.C_标志) = "盈"
                Else
                    .TextMatrix(i, mBillCol.C_标志) = "平"
                End If
                .TextMatrix(i, mBillCol.C_数量差) = Format(Abs(Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_帐面数量))), mFMT.FM_数量)
                
                
                If Val(.TextMatrix(i, mBillCol.C_帐面数量)) = 0 Then
                    strMoneyDigit = "#0.00000"
                Else
                    strMoneyDigit = mFMT.FM_金额
                End If
                 '金额差=当前售价*实盘数量-实际金额
                '差价差=金额差*iif(实际金额<=0,指导差价率,(实际差价/实际金额))
                .TextMatrix(i, mBillCol.c_金额差) = Format(Val(.TextMatrix(i, mBillCol.C_售价)) * Val(.TextMatrix(i, mBillCol.C_实盘数量)) - Val(.TextMatrix(i, mBillCol.C_实际金额)), strMoneyDigit)
                If rsData!是否变价 = 1 And Val(.TextMatrix(i, mBillCol.C_帐面数量)) = 0 Then
                    .TextMatrix(i, mBillCol.c_差价差) = Format(Val(.TextMatrix(i, mBillCol.C_数量差)) * (Val(.TextMatrix(i, mBillCol.C_售价)) - dbl成本价) - Val(.TextMatrix(i, mBillCol.C_实际差价)), strMoneyDigit)
                Else
                    .TextMatrix(i, mBillCol.c_差价差) = Format(Val(.TextMatrix(i, mBillCol.C_实盘数量)) * (Val(.TextMatrix(i, mBillCol.C_售价)) - Val(.TextMatrix(i, mBillCol.C_成本价))) - Val(.TextMatrix(i, mBillCol.C_实际差价)), strMoneyDigit)
                End If
                dbl金额差 = .TextMatrix(i, mBillCol.c_金额差)
                dbl差价差 = .TextMatrix(i, mBillCol.c_差价差)
                
                If .TextMatrix(i, mBillCol.C_标志) = "亏" Then
                    '保证实际金额与金额差相同的符号（因为入出系数为-1，这样就能保证完全冲销为零）
                    If Not 相同符号(Val(.TextMatrix(i, mBillCol.c_金额差)), Val(.TextMatrix(i, mBillCol.C_实际金额))) Then
                        .TextMatrix(i, mBillCol.c_金额差) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_金额差)), strMoneyDigit)
                    End If
                    If Not 相同符号(Val(.TextMatrix(i, mBillCol.c_差价差)), Val(.TextMatrix(i, mBillCol.C_实际差价))) Then
                        .TextMatrix(i, mBillCol.c_差价差) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_差价差)), strMoneyDigit)
                    End If
                End If
            End If
            
            .TextMatrix(i, mBillCol.C_盘点成本金额) = Format(Val(.TextMatrix(i, mBillCol.C_实际金额)) + dbl金额差 - (Val(.TextMatrix(i, mBillCol.C_实际差价)) + dbl差价差), mFMT.FM_金额)
            .TextMatrix(i, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(i, mBillCol.c_金额差)) - Val(.TextMatrix(i, mBillCol.c_差价差)), mFMT.FM_金额)
            
            Call ShowPercent(i / lngRecordCount)
            i = i + 1
nextloop:
            rsData.MoveNext
        Loop
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        .Redraw = True
    End With
    Call FS.StopFlash
    Call 显示合计金额
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1: mshBill.Col = mBillCol.C_实盘数量
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowPercent(sngPercent As Single)
    '功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub



Private Function GetDateStock(str盘存时间 As String, lng库房id As Long, int盘点方式 As Integer, _
    Optional blnZero As Boolean = False, Optional ByVal bln汇总 As Boolean = False, Optional str分类ID As String = "", _
    Optional lng材料ID As Long = 0, Optional bln盘点零数量且有金额 As Boolean = False, Optional ByVal str库房货位 As String = "所有") As ADODB.Recordset
    '功能：获取指定条件材料在指定时间点的库存及相关信息
    '参数：str盘存时间=要求以YYYY-MM-DD HH24:MI:SS为格式的时间字符串
    '      int盘点方式: 非0-自动生成盘点表（1-每日 ;2-每周 ;3-每月 ;4-每季度 ;5-忽略盘点方式）；0-表示非自动生成盘点表
    '      bln材料ID:为str材料条件，表示按材料ID进行过滤
    '      blnZero=是否读取库存数结果为0的材料,缺省否.当强行输入该材料时,才设为是。
    Dim rsTmp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim strUnit As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    Dim strRule As String
    Dim str材料条件 As String
    
    On Error GoTo errH
    
    '构造材料查询条件(材料特性 B)
    str材料条件 = " And (c.撤档时间>[2] Or c.撤档时间 is NULL)"
    
    If int盘点方式 <> 5 And int盘点方式 <> 0 Then '忽略盘点方式
        str材料条件 = str材料条件 & " And Substr(E.盘点属性," & int盘点方式 & ",1)='1' "
    End If

    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    blnStock = CheckPartProp(lng库房id)

    If str库房货位 = "" Then
        str库房货位 = "所有"
    ElseIf str库房货位 <> "所有" Then
        str库房货位 = Replace(str库房货位, "'", "")
        str库房货位 = "," & str库房货位 & ","
    End If
    
    If int盘点方式 <> 0 Then '自动生成盘点表时才有 “库房货位”
        str材料条件 = str材料条件 & IIf(str库房货位 = "所有", "", " and (Instr([6], ',' || e.库房货位 || ',') > 0) ")
    End If
    
    If lng材料ID > 0 Then '具体的材料，材料查询条件重写
        str材料条件 = " And B.材料ID=[4] "
    End If
    
    '取得当前库存
    gstrSQL = "" & _
        "   SELECT a.库房id, b.材料id, NVL (a.批次, 0) AS 批次, a.实际数量,0 盘点数量,a.实际金额, a.实际差价, a.可用数量, a.平均成本价 成本价,a.上次批号 AS 批号,a.上次产地 AS 产地,a.批准文号,a.效期, e.库房货位 " & _
        "   FROM 药品库存 a, 材料特性 b,收费项目目录 c ,诊疗项目目录 D, 材料储备限额 E" & _
        "   Where a.药品id = b.材料id and a.药品id=c.id and b.诊疗id=d.id  " & _
        "           and a.库房id = e.库房id" & IIf(int盘点方式 = 5, "(+)", "") & " And a.药品id = e.材料id" & IIf(int盘点方式 = 5, "(+)", "") & _
        "           AND a.性质=1 " & _
        "           AND a.库房id =[1] " & str材料条件 & IIf(str分类ID = "", "", " and D.分类id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ")
    
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "   SELECT a.库房id, b.材料id, NVL (a.批次, 0) AS 批次, " & _
        "           -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量,0 盘点数量, " & _
        "           -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
        "           -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价,0 AS 可用数量,max(decode(a.单据,22,a.单量,23,0,a.成本价)) as 成本价,a.批号,a.批准文号,a.产地,a.效期,Max(e.库房货位) As 库房货位 " & _
        "   FROM 药品收发记录 a,  材料特性 b,收费项目目录 c ,诊疗项目目录 D,材料储备限额 E,收费执行科室 G " & _
        "   Where a.药品id + 0 = b.材料id and a.药品id +0 =c.id and b.诊疗id=d.id " & _
        "           and a.库房id + 0 = e.库房id" & IIf(int盘点方式 = 5, "(+)", "") & " And a.药品id + 0 = e.材料id" & IIf(int盘点方式 = 5, "(+)", "") & _
        "           AND a.库房id + 0 =[1] " & _
        "           and b.材料id=g.收费细目id " & IIf(mbln盘无存储库房材料, "(+)", "") & _
        "           and G.执行科室id" & IIf(mbln盘无存储库房材料, "(+)", "") & "=[1] " & _
        "           AND a.审核日期 >[2] " & str材料条件 & IIf(str分类ID = "", "", " and D.分类id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & _
        " GROUP BY a.库房id, b.材料id, a.批次,a.批号,a.产地,a.批准文号,a.效期 "
    
    If bln汇总 Then
        gstrSQL = gstrSQL & _
            "   UNION ALL" & _
            "   SELECT A.库房ID,B.材料ID,NVL(A.批次, 0) AS 批次,0 AS 实际数量,SUM(A.扣率) 盘点数量," & _
            "           0 AS 实际金额,0 AS 实际差价,0 AS 可用数量,0 as 成本价,A.批号,A.产地,A.批准文号,A.效期,e.库房货位 " & _
            "   FROM 药品收发记录 A, 材料特性 b,收费项目目录 c ,诊疗项目目录 D,材料储备限额 E" & _
            "   Where A.药品ID+0 = B.材料ID And A.单据 = 23 and a.药品id+0=c.id and b.诊疗id=d.id" & _
            "           AND a.库房id + 0 = e.库房id" & IIf(int盘点方式 = 5, "(+)", "") & " And a.药品id + 0 = e.材料id" & IIf(int盘点方式 = 5, "(+)", "") & " AND A.库房ID + 0 =[1] " & _
            "           AND A.频次 =[3] " & str材料条件 & IIf(str分类ID = "", "", " and D.分类id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & _
            "           AND (c.撤档时间 >[2] OR c.撤档时间 IS NULL)" & _
            " GROUP BY A.库房ID,B.材料ID,A.批次,A.批号,A.产地,A.批准文号,A.效期,e.库房货位"
    End If
    
    '取得盘点时间那一刻的帐面数量
    gstrSQL = "" & _
        "   SELECT 库房id, 材料id, 批次, SUM (实际数量) AS 帐面数量,SUM (盘点数量) AS 盘点数量," & _
        "           SUM (实际金额) AS 实际金额, SUM (实际差价) AS 实际差价, " & _
        "           SUM(可用数量) As 可用数量,max(成本价) as 成本价,max(批号) as 批号,max(产地) as 产地 ,max(批准文号) as 批准文号,max(效期) as 效期,Max(库房货位) As 库房货位 " & _
        "   FROM ( " & gstrSQL & ") " & _
        "   GROUP BY 库房id, 材料id, 批次 " & _
       IIf(bln盘点零数量且有金额, "   Having sum(实际数量)=0 and (sum(实际金额)<>0 or sum(实际差价)<>0 )", "")
    
    
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.计算单位 AS 单位, nvl(a.帐面数量,0) AS 帐面数量, nvl(a.盘点数量,0) AS 盘点数量,nvl(a.可用数量,0) AS 可用数量, '1' as 比例系数," & _
             " f.现价 as 售价,decode(nvl(a.成本价,0),0,B.成本价,A.成本价) as 成本价,b.成本价 as 最后进价,"
        Case Else
            strUnitQuantity = "b.包装单位 AS 单位, (nvl(a.帐面数量,0) / b.换算系数) AS 帐面数量, (nvl(a.盘点数量,0) / b.换算系数) AS 盘点数量,(nvl(a.可用数量,0) / b.换算系数) AS 可用数量,b.换算系数 as 比例系数," & _
             "f.现价*b.换算系数 as 售价,decode(nvl(a.成本价,0),0,B.成本价,A.成本价)*b.换算系数 as 成本价,b.成本价*b.换算系数 as 最后进价,"
    End Select
    '成本价计算的方式:
    'a.如果不分批
    '          1.库存数量:成本价=(库存金额-库存差价)/库存数量,
    '          2.无库存数量:上次成本价:即取材料特性的成本价
    'b.如果分批
    '          1.有库存,取库存的上次采购价,
    '          2.无库存数量:上次成本价:即取材料特性的成本价
    gstrSQL = "" & _
        "   SELECT  DISTINCT b.材料id, c.编码, c.名称 AS 商品名称," & _
        "           zlSpellCode(c.名称) 名称,c.规格, Decode(a.产地, Null, decode(b.上次产地,null,c.产地,b.上次产地), a.产地) As 产地,A.批准文号,a.库房货位,nvl(a.批次,0) 批次, a.批号, a.效期," & strUnitQuantity & _
        "           nvl(a.实际金额,0) as 实际金额 ,nvl(a.实际差价,0) as 实际差价, b.指导差价率,c.是否变价,b.库房分批,b.在用分批,nvl(b.最大效期,0) 最大效期,decode(a.批号,null,1,0) 批号编辑,decode(a.产地,null,1,0) 产地编辑 " & _
        "   From (" & gstrSQL & ") A , 材料特性 b,收费项目目录 c ,收费执行科室 G,收费价目 F "
        
    gstrSQL = gstrSQL & IIf(blnZero And str分类ID <> "", ", (select D.ID From 诊疗项目目录 D  where 1=1 " & IIf(str分类ID = "", "", " and D.分类id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & ") D ", "") & _
        "   Where " & IIf(blnZero = False, "a.材料id = b.材料id and a.材料id=c.id  ", " b.材料id = a.材料id(+) and b.材料id=c.id(+) " & IIf(str分类ID <> "", " and b.诊疗id=d.id", "")) & _
        "           AND b.材料id=f.收费细目id " & _
        "           and b.材料id=g.收费细目id " & IIf(mbln盘无存储库房材料, "(+)", "") & _
        "           and G.执行科室id" & IIf(mbln盘无存储库房材料, "(+)", "") & "=[1] " & _
        "           and ((SYSDATE BETWEEN f.执行日期 AND f.终止日期) OR (SYSDATE >= f.执行日期 AND f.终止日期 IS NULL)) " & _
        GetPriceClassString("F") & _
                    IIf(blnZero = False, " AND (a.帐面数量<>0 or nvl(a.实际金额,0)<>0 or nvl(a.实际差价,0)<>0 Or nvl(a.盘点数量,0)<>0)", "") & _
                    IIf(lng材料ID > 0, " And B.材料ID=[4] ", "") & _
        " ORDER BY " & IIf(strCompare = "0", "c.编码", IIf(strCompare = "1", "c.编码", IIf(strCompare = "2", "c.名称", "a.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
        
    Screen.MousePointer = 11
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "卫生材料盘点管理", lng库房id, CDate(str盘存时间), str盘存时间, lng材料ID, str分类ID, str库房货位)
    
    Set GetDateStock = rsTmp
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog

End Function

Private Function Get汇总记录单(ByVal lng库房id As Long, ByVal str盘点时间 As String) As ADODB.Recordset
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    '功能：按盘点记录单进行汇总统计
    '参数：lng库房ID-库房ID
    '      str盘点时间 -盘点时间:格式yyyy-mm-dd hh24:mi:ss
    '返回：返回符合条件的记录
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim strUnit As String
    Dim str盘点单 As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    Dim str盘点单NO串 As String
    
    On Error GoTo errH
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    str盘点单NO串 = Replace(mstr盘点单号, "'", "")
    
    blnStock = CheckPartProp(lng库房id)
    
    '取得当前库存
    gstrSQL = "" & _
        "   SELECT a.库房id, b.材料id, NVL (a.批次,0) AS 批次,a.实际数量,0 盘点数量,a.实际金额, a.实际差价, a.可用数量,A.平均成本价 as 成本价,a.上次批号 AS 批号,a.上次产地 AS 产地,A.批准文号,a.效期 " & _
        "   FROM 药品库存 a, 材料特性 b " & _
        "   Where a.药品id = b.材料id " & _
        "           AND a.性质=1 " & _
        "           AND a.库房id =[1] "
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "   SELECT a.库房id, b.材料id, NVL (a.批次, 0) AS 批次, " & _
        "           -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量,0 盘点数量, " & _
        "           -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
        "           -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价,0 AS 可用数量,max(decode(a.单据,22,a.单量,23,0,a.成本价)) as 成本价,a.批号,a.产地,a.批准文号,a.效期 " & _
        "   FROM 药品收发记录 a,  材料特性 b" & _
        "   Where a.药品id = b.材料id" & _
        "           AND a.库房id + 0 =[1] " & _
        "           AND a.审核日期 >[2] " & _
        " GROUP BY a.库房id, b.材料id, a.批次,a.批号,a.产地,a.批准文号,a.效期 "
        
    str盘点单 = "" & _
            "   SELECT A.库房ID,B.材料ID,NVL(A.批次, 0) AS 批次,0 AS 实际数量,SUM(A.扣率) 盘点数量," & _
            "           0 AS 实际金额,0 AS 实际差价,0 AS 可用数量,a.单量 as 成本价,A.批号,A.产地,A.批准文号,A.效期" & _
            "   FROM 药品收发记录 A,材料特性 b " & _
            "   Where A.药品ID = B.材料ID And A.单据 = 23 AND A.库房ID + 0 =[1] " & _
            "           AND A.No in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & _
            " GROUP BY A.库房ID,B.材料ID,A.批次,A.批号,A.产地,a.批准文号,A.效期,a.单量"
    
    gstrSQL = gstrSQL & _
        "   UNION ALL" & vbCrLf & str盘点单
    
    
    '取得盘点时间那一刻的帐面数量
    gstrSQL = "" & _
        "Select 库房ID,材料ID,批次,max(a.成本价) as 成本价,max(批号) 批号,max(产地) 产地 ,max(批准文号) as 批准文号,max(效期) 效期," & _
        "       sum(nvl(可用数量,0)) 可用数量," & _
        "       sum(nvl(实际数量,0)) 帐面数量," & _
        "       sum(nvl(盘点数量,0)) 盘点数量," & _
        "       sum(nvl(实际金额,0)) 实际金额," & _
        "       sum(nvl(实际差价,0)) 实际差价" & _
        "   From (" & gstrSQL & ") a" & _
        IIf(mbln只统计盘点单卫材, _
            " where  Exists (Select 1 from 药品收发记录 T1 " & _
            "                where T1.NO in (select * from Table(Cast(f_Str2list([3]) as zlTools.t_Strlist))) and T1.单据=23 and a.材料ID=T1.药品id+0 ) ", _
            "") & _
        "  Group by 库房ID,材料ID,批次"
    
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.计算单位 AS 单位, nvl(a.帐面数量,0) AS 帐面数量, nvl(a.盘点数量,0) AS 盘点数量,nvl(a.可用数量,0) AS 可用数量, '1' as 比例系数," & _
             " f.售价,decode(a.成本价,null,B.成本价,A.成本价) as 成本价,b.成本价 as 最后进价,"
        Case Else
            strUnitQuantity = "b.包装单位 AS 单位, (nvl(a.帐面数量,0) / b.换算系数) AS 帐面数量, (nvl(a.盘点数量,0) / b.换算系数) AS 盘点数量,(nvl(a.可用数量,0) / b.换算系数) AS 可用数量,b.换算系数 as 比例系数," & _
             "f.售价*b.换算系数 as 售价,decode(a.成本价,null,B.成本价,A.成本价)*b.换算系数 as 成本价,b.成本价*b.换算系数 as 最后进价, "
    End Select
    
    gstrSQL = "" & _
        "   SELECT  DISTINCT b.材料id, c.编码, c.名称 AS 商品名称," & _
        "           zlSpellCode(c.名称) 名称,c.规格, a.产地,a.批准文号,e.库房货位," & _
        "           nvl(a.批次,0) 批次, a.批号, a.效期," & strUnitQuantity & _
        "           nvl(a.实际金额,0) as 实际金额 ,nvl(a.实际差价,0) as 实际差价, b.指导差价率,c.是否变价,b.库房分批,b.在用分批,decode(a.批号,null,1,0) 批号编辑,decode(a.产地,null,1,0) 产地编辑 " & _
        "   From (" & gstrSQL & ") A , 材料特性 b,收费项目目录 c ,材料储备限额 e, " & _
        "   (select 收费细目id,执行科室id from 收费执行科室 where 执行科室id=[1]) G, " & _
        "        (SELECT 收费细目id, 现价 as 售价 From 收费价目  WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL))" & _
        GetPriceClassString("") & ") f " & _
        "   Where a.材料id = b.材料id and a.材料id=c.id   AND b.材料id=f.收费细目id " & _
        "         and (c.撤档时间 is null or c.撤档时间>[2]) " & _
        "           and A.库房id=E.库房id(+) and A.材料id=E.材料id(+) and b.材料id=g.收费细目id " & IIf(mbln盘无存储库房材料, "(+)", "") & _
        " ORDER BY " & IIf(strCompare = "0", "c.编码", IIf(strCompare = "1", "c.编码", IIf(strCompare = "2", "c.名称", "e.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
        
    Screen.MousePointer = 11
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "汇总卫生材料盘点记录", lng库房id, CDate(str盘点时间), str盘点单NO串)
    
    Set Get汇总记录单 = rsTemp
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub Form_Load()

    Dim strReg As String
    mintUnit = Val(zlDatabase.GetPara("盘点表单位", glngSys, mlngModule, "0"))
    mbln盘无存储库房材料 = Val(zlDatabase.GetPara("存储库房", glngSys, mlngModule, "0"))
    
    mbln分批卫材批号产地控制 = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    mblnFirst = True
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    initCard
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mBillCol.C_盘点成本金额) = IIf(mblnCostView = True, 1400, 0)
        .ColWidth(mBillCol.C_盘点成本金额差) = IIf(mblnCostView = True, 1400, 0)
        .ColWidth(mBillCol.C_成本价) = IIf(mblnCostView = True, 800, 0)
        .ColWidth(mBillCol.c_差价差) = IIf(mblnCostView = True, 900, 0)
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim lngRow As Long
    Dim strOrder As String, strCompare As String
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim strMoneyDigit As String
    '库房
    
    On Error GoTo errHandle
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
    Select Case mint编辑状态
        Case 1, 5, 6
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            '如果是全部盘为零，则检查是否存在未审核的盘点单
'            If mint编辑状态 = 6 Then
'                gstrSQL = "" & _
'                    "    Select Count(*) Records " & _
'                    "    From 药品收发记录" & _
'                    "    Where 单据<>23 And 审核人 Is NULL And 库房ID=[1]"
'
'                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在未审卫生材料单据", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
'                If Not rsTemp.EOF Then
'                    If Not IsNull(rsTemp!Records) Then
'                        If rsTemp!Records <> 0 Then
'                            mintParallelRecord = 5
'                            Exit Sub
'                        End If
'                    End If
'                End If
'            End If
            
            cmd固定列.Visible = (mint编辑状态 = 1)
        Case 2, 3, 4
            initGrid
            If mint编辑状态 <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "" & _
                    "   Select distinct b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   Where a.库房id=b.id " & _
                    "           and A.单据 = 22 and a.no=[1]"
                    
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取部门数据", mstr单据号)
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsTemp!名称
                txtStock.Tag = rsTemp!Id
                rsTemp.Close
            End If
            
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.计算单位 AS 单位, A.填写数量 AS 帐面数量,A.扣率 AS 实盘数量, A.实际数量 AS 数量差,'1' as 比例系数,a.零售价 as 售价,A.单量 as 成本价,"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.填写数量/ B.换算系数) AS 帐面数量,(A.扣率/ B.换算系数) AS 实盘数量, (A.实际数量 / B.换算系数) AS 数量差,B.换算系数 as 比例系数,a.零售价*B.换算系数 as 售价,a.单量*B.换算系数 as 成本价,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || D.编码 || ']' || D.名称) AS 材料信息," & _
                "                   zlSpellCode(D.名称) 名称,A.入出系数,D.规格,A.产地,A.批准文号,C.库房货位, A.批号,a.效期,a.批次," & strUnitQuantity & _
                "                   A.零售金额 as 金额差,A.差价 as 差价差, " & _
                "                   a.摘要,填制人,填制日期,审核人,审核日期,a.频次 as 盘点时间,A.单量,a.成本价 as 库存金额,a.成本金额 as 库存差价,b.指导差价率,d.是否变价,b.在用分批,nvl(a.发药方式,0) as 新批次,decode(E.上次批号,null,1,0) 批号编辑,decode(E.上次产地,null,1,0) 产地编辑 " & _
                "           FROM 药品收发记录 A, 材料特性 b,收费项目目录 D,材料储备限额 C,药品库存 E " & _
                "           Where A.药品id = B.材料id and a.药品id=d.id  " & _
                "                   And A.药品ID=C.材料ID(+) And A.库房ID=C.库房ID(+) AND A.记录状态 =[3]" & _
                "                   And A.药品ID=E.药品ID(+) And A.库房ID=E.库房ID(+) And nvl(A.批次,0) = nvl(E.批次(+),0) AND A.单据 =[1] AND A.No =[2]" & _
                "       ) " & _
                "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "材料信息", IIf(strCompare = "2", "名称", "库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 22, mstr单据号, mint记录状态)
            mstrTime_Start = GetBillInfo(22, mstr单据号)
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            Txt填制人 = rsTemp!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户名
            End If
            Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd hh:mm:ss")
            
            Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
            Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            txtCheckDate.Caption = rsTemp!盘点时间
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            lngRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    lngRow = lngRow + 1
                    .Rows = lngRow + 1
                    .TextMatrix(lngRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(lngRow, mBillCol.C_材料) = rsTemp!材料信息
                    .TextMatrix(lngRow, mBillCol.C_序号) = rsTemp!序号
                    .TextMatrix(lngRow, mBillCol.c_规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(lngRow, mBillCol.C_产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(lngRow, mBillCol.C_批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                    .TextMatrix(lngRow, mBillCol.C_库房货位) = IIf(IsNull(rsTemp!库房货位), "", rsTemp!库房货位)
                    .TextMatrix(lngRow, mBillCol.c_单位) = rsTemp!单位
                    .TextMatrix(lngRow, mBillCol.c_批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(lngRow, mBillCol.C_效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    
                    .TextMatrix(lngRow, mBillCol.C_实际差价) = Format(rsTemp!库存差价, mFMT.FM_金额)
                    .TextMatrix(lngRow, mBillCol.C_实际金额) = Format(rsTemp!库存金额, mFMT.FM_金额)
                    .TextMatrix(lngRow, mBillCol.C_指导差价率) = Format(rsTemp!指导差价率, mFMT.FM_金额) & "||" & rsTemp!是否变价 & "||" & rsTemp!在用分批
                    .TextMatrix(lngRow, mBillCol.C_批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                    .TextMatrix(lngRow, mBillCol.c_新批次) = rsTemp!新批次
                    .TextMatrix(lngRow, mBillCol.c_比例系数) = rsTemp!比例系数
                    
                    If Val(.TextMatrix(lngRow, mBillCol.C_批次)) <> 0 Then '分批材料
                        .TextMatrix(lngRow, mBillCol.c_批号编辑) = rsTemp!批号编辑
                        .TextMatrix(lngRow, mBillCol.c_产地编辑) = rsTemp!产地编辑
                    End If
                    
                    .TextMatrix(lngRow, mBillCol.C_帐面数量) = Format(rsTemp!帐面数量, mFMT.FM_数量)
                    .TextMatrix(lngRow, mBillCol.C_实盘数量) = Format(rsTemp!实盘数量, mFMT.FM_数量)
                    .TextMatrix(lngRow, mBillCol.C_数量差) = Format(rsTemp!数量差, mFMT.FM_数量)
                    If rsTemp!实盘数量 > rsTemp!帐面数量 Then
                        .TextMatrix(lngRow, mBillCol.C_标志) = "盈"
                    ElseIf rsTemp!实盘数量 < rsTemp!帐面数量 Then
                        .TextMatrix(lngRow, mBillCol.C_标志) = "亏"
                    Else
                        .TextMatrix(lngRow, mBillCol.C_标志) = "平"
                    End If
                    
                    If Val(.TextMatrix(lngRow, mBillCol.C_帐面数量)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_金额
                    End If
                    
                    .TextMatrix(lngRow, mBillCol.c_金额差) = Format(rsTemp!金额差, strMoneyDigit)
                    .TextMatrix(lngRow, mBillCol.c_差价差) = Format(rsTemp!差价差, strMoneyDigit)
                    
                    .TextMatrix(lngRow, mBillCol.C_售价) = Format(rsTemp!售价, mFMT.FM_零售价)
                    .TextMatrix(lngRow, mBillCol.C_成本价) = Format(zlStr.NVL(rsTemp!成本价, 0), mFMT.FM_成本价)
                    .TextMatrix(lngRow, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(lngRow, mBillCol.C_实盘数量)) * Val(.TextMatrix(lngRow, mBillCol.C_售价)), mFMT.FM_金额)
                    '保持与主界面金额差和差价差算法一致
                    dbl金额差 = Val(.TextMatrix(lngRow, mBillCol.c_金额差)) * rsTemp!入出系数 * IIf(mint记录状态 = 1, 1, IIf(mint记录状态 Mod 3 = 0, 1, -1))
                    dbl差价差 = Val(.TextMatrix(lngRow, mBillCol.c_差价差)) * rsTemp!入出系数 * IIf(mint记录状态 = 1, 1, IIf(mint记录状态 Mod 3 = 0, 1, -1))
                    '成本金额=成本价*实盘数量=(账面金额+金额差) -(账面差价+差价差) 用后者是为了控制报表与程序出的盘点单能对账
                    .TextMatrix(lngRow, mBillCol.C_盘点成本金额) = Format((zlStr.NVL(rsTemp!库存金额, 0) + dbl金额差) - (zlStr.NVL(rsTemp!库存差价, 0) + dbl差价差), mFMT.FM_金额)
                    .TextMatrix(lngRow, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(lngRow, mBillCol.c_金额差)) - Val(.TextMatrix(lngRow, mBillCol.c_差价差)), mFMT.FM_金额)
                    
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    Call 显示合计金额
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
        .Cols = mBillCol.C_Cols
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_行号) = ""
        .TextMatrix(0, mBillCol.C_材料) = "名称与编码"
        .TextMatrix(0, mBillCol.C_序号) = "序号"
        .TextMatrix(0, mBillCol.c_规格) = "规格"
        .TextMatrix(0, mBillCol.C_产地) = "产地"
        .TextMatrix(0, mBillCol.C_批准文号) = "批准文号"
        .TextMatrix(0, mBillCol.C_库房货位) = "库房货位"
        .TextMatrix(0, mBillCol.c_单位) = "单位"
        .TextMatrix(0, mBillCol.c_批号) = "批号"
        .TextMatrix(0, mBillCol.C_效期) = "失效期"
        .TextMatrix(0, mBillCol.C_批次) = "批次"
        .TextMatrix(0, mBillCol.C_可用数量) = "可用数量"
        .TextMatrix(0, mBillCol.c_比例系数) = "比例系数"
        .TextMatrix(0, mBillCol.C_指导差价率) = "指导差价率"
        .TextMatrix(0, mBillCol.C_实际差价) = "实际差价"
        .TextMatrix(0, mBillCol.C_实际金额) = "实际金额"
        .TextMatrix(0, mBillCol.C_帐面数量) = "帐面数量"
        .TextMatrix(0, mBillCol.C_实盘数量) = "实盘数量"
        .TextMatrix(0, mBillCol.C_标志) = "标志"
        .TextMatrix(0, mBillCol.C_数量差) = "数量差"
        .TextMatrix(0, mBillCol.C_成本价) = "成本价"
        .TextMatrix(0, mBillCol.C_售价) = "售价"
        .TextMatrix(0, mBillCol.c_金额差) = "金额差"
        .TextMatrix(0, mBillCol.c_差价差) = "差价差"
        .TextMatrix(0, mBillCol.C_盘点金额) = "盘点金额"
        .TextMatrix(0, mBillCol.C_盘点成本金额) = "盘点成本金额"
        .TextMatrix(0, mBillCol.C_盘点成本金额差) = "盘点成本金额差"
        .TextMatrix(0, mBillCol.c_新批次) = "新批次"
        
        .TextMatrix(0, mBillCol.c_批号编辑) = "批号编辑"
        .TextMatrix(0, mBillCol.c_产地编辑) = "产地编辑"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_行号) = 300
        .ColWidth(mBillCol.C_批次) = 0
        .ColWidth(mBillCol.C_序号) = 0
        .ColWidth(mBillCol.C_可用数量) = 0
        .ColWidth(mBillCol.c_比例系数) = 0
        .ColWidth(mBillCol.C_指导差价率) = 0
        .ColWidth(mBillCol.C_实际差价) = 0
        .ColWidth(mBillCol.C_实际金额) = 0
        .ColWidth(mBillCol.C_材料) = 2000
        .ColWidth(mBillCol.c_规格) = 900
        .ColWidth(mBillCol.C_产地) = 800
        .ColWidth(mBillCol.C_批准文号) = 1000
        .ColWidth(mBillCol.C_库房货位) = 2000
        .ColWidth(mBillCol.c_单位) = 500
        .ColWidth(mBillCol.c_批号) = 800
        .ColWidth(mBillCol.C_效期) = 1000
        .ColWidth(mBillCol.C_帐面数量) = 800
        .ColWidth(mBillCol.C_实盘数量) = 800
        .ColWidth(mBillCol.C_标志) = 500
        .ColWidth(mBillCol.C_数量差) = 800
        .ColWidth(mBillCol.C_成本价) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_售价) = 800
        .ColWidth(mBillCol.c_金额差) = 900
        .ColWidth(mBillCol.c_差价差) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_盘点金额) = 900
        .ColWidth(mBillCol.C_盘点成本金额) = IIf(mblnCostView = False, 0, 1400)
        .ColWidth(mBillCol.C_盘点成本金额差) = IIf(mblnCostView = False, 0, 1500)
        .ColWidth(mBillCol.c_新批次) = 0
        .ColWidth(mBillCol.c_批号编辑) = 0
        .ColWidth(mBillCol.c_产地编辑) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mBillCol.C_行号) = 5
        .ColData(mBillCol.c_规格) = 5
        .ColData(mBillCol.C_序号) = 5
        .ColData(mBillCol.C_产地) = 5
        .ColData(mBillCol.C_批准文号) = 5
        .ColData(mBillCol.C_库房货位) = 5
        .ColData(mBillCol.c_单位) = 5
        .ColData(mBillCol.c_批号) = 5
        .ColData(mBillCol.C_效期) = 5
        .ColData(mBillCol.C_批次) = 5
        .ColData(mBillCol.C_可用数量) = 5
        .ColData(mBillCol.c_比例系数) = 5
        .ColData(mBillCol.C_指导差价率) = 5
        .ColData(mBillCol.C_实际差价) = 5
        .ColData(mBillCol.C_实际金额) = 5
        .ColData(mBillCol.C_帐面数量) = 5
        
        .ColData(mBillCol.C_标志) = 5
        .ColData(mBillCol.C_数量差) = 5
        .ColData(mBillCol.C_成本价) = 5
        .ColData(mBillCol.C_售价) = 5
        .ColData(mBillCol.c_金额差) = 5
        .ColData(mBillCol.c_差价差) = 5
        .ColData(mBillCol.C_盘点金额) = 5
        .ColData(mBillCol.C_盘点成本金额) = 5
        .ColData(mBillCol.C_盘点成本金额差) = 5
        .ColData(mBillCol.c_新批次) = 5
        .ColData(mBillCol.c_批号编辑) = 5
        .ColData(mBillCol.c_产地编辑) = 5
                
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mBillCol.C_材料) = 1
            .ColData(mBillCol.C_实盘数量) = 4
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            .ColData(mBillCol.C_实盘数量) = 5
        ElseIf mint编辑状态 = 5 Or mint编辑状态 = 6 Then
'            .Active = False
            txt摘要.Enabled = True
            .ColData(mBillCol.C_实盘数量) = 5
        End If
        
        .ColAlignment(mBillCol.C_材料) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_规格) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_产地) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_批准文号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_单位) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_批号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_效期) = flexAlignLeftCenter
        
        .ColAlignment(mBillCol.C_帐面数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_实盘数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_标志) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_数量差) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_成本价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_金额差) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_差价差) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_盘点金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_盘点成本金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_盘点成本金额差) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_新批次) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_批号编辑) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_产地编辑) = flexAlignRightCenter
        
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
    
    With lblCheckSum
        lblCheckCostSum.Left = .Left + .Width + 100
        lblCheckCostSum.Top = .Top
    End With
    
    If mblnCostView = False Then
        lblCheckCostSum.Visible = False
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
    
    With cmd固定列
        .Left = CmdSave.Left - .Width - 150
        .Top = CmdSave.Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
    Dim strNo As String
    Dim str审核人 As String
    
    mblnSave = False
    SaveCheck = False
    
    str审核人 = UserInfo.用户名
    strNo = txtNO.Tag
    On Error GoTo errHandle
    
    gstrSQL = "zl_材料盘点_Verify('" & strNo & "','" & str审核人 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        
        
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function



Private Sub mnuDefault_Click()
    mshBill.MsfObj.FixedCols = 1
    mshBill.ColData(mBillCol.C_材料) = 1
    mshBill.LocateCol = mBillCol.C_材料
End Sub

Private Sub mnuFirst_Click()
    mshBill.Redraw = False
    mshBill.ColData(mBillCol.C_材料) = 5
    mshBill.MsfObj.FixedCols = 14
    mshBill.LocateCol = 17
    
    '设置对齐方式
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_规格) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_单位) = 4
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_批号) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.C_效期) = 1
    mshBill.Refresh
    mshBill.Redraw = True
End Sub

Private Sub mnuSecond_Click()
    mshBill.Redraw = False
    mshBill.ColData(mBillCol.C_材料) = 5
    mshBill.MsfObj.FixedCols = 16
    mshBill.LocateCol = 17
    
    '设置对齐方式
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_规格) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_单位) = 4
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_批号) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.C_效期) = 1
    mshBill.Refresh
    mshBill.Redraw = True
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_行号, Row)
    If mshBill.MsfObj.FixedCols > mBillCol.C_材料 Then
        mshBill.PrimaryCol = mBillCol.C_实盘数量
        mshBill.Col = mBillCol.C_实盘数量
        mshBill.PrimaryCol = mBillCol.C_材料
    End If
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mBillCol.C_行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "3456", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行卫生材料？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int点击行 As Integer
    
    On Error GoTo errHandle
    
    int点击行 = mshBill.Row
    
    If mshBill.Col = mBillCol.C_材料 Then
        Set RecReturn = Frm材料选择器.ShowMe(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, False, True, True, True, , , , , txtCheckDate.Caption, , , mbln盘无存储库房材料, mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    
                    If SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)) Then
    
                        If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                        .Row = .Row + 1
                    End If
                    
                    .Col = mBillCol.C_实盘数量
                    RecReturn.MoveNext
                Next
                
                mshBill.Row = int点击行
                
                If mstr重复卫材 <> "" Then
                    MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                    mstr重复卫材 = ""
                End If
                
    '            If RecReturn.RecordCount = 1 Then
    '                Call SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次))
    '                .Col = mBillCol.C_实盘数量
    '            End If
            End With
            RecReturn.Close
        End If
    Else
        gstrSQL = "Select rownum as id,null as 上级id,编码,名称,简码,1 as 末级 From 材料生产商 "
        Set RecReturn = zlDatabase.ShowSelect(Me, gstrSQL, 1, "材料生产商选择", True, , "选择卫生材料生产商或厂牌")
  
        If RecReturn Is Nothing Then Exit Sub
        If RecReturn.State <> 1 Then Exit Sub
        
        With RecReturn
            If CheckQualifications(mlngModule, 1, CStr(NVL(!名称))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_产地) = NVL(!名称)
        End With
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_帐面数量 Or .Col = mBillCol.C_实盘数量 Or .Col = mBillCol.C_成本价 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_帐面数量, mBillCol.C_实盘数量
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
                Case mBillCol.C_成本价
                   intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.成本价小数, g_小数位数.obj_散装小数.成本价小数)
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.零售价小数, g_小数位数.obj_散装小数.零售价小数)
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
    Dim lng批次  As Long
    Dim lng新批次  As Long
    
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
                Call 提示库存数
                If mshBill.MsfObj.FixedCols > mBillCol.C_材料 Then
                    mshBill.PrimaryCol = mBillCol.C_实盘数量
                    mshBill.Col = mBillCol.C_实盘数量
                    mshBill.PrimaryCol = mBillCol.C_材料
                End If
            Case mBillCol.c_批号
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mBillCol.C_效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .ColData(mBillCol.C_效期) = 2 Then
                    If .TextMatrix(.Row, mBillCol.c_批号) <> "" And Len(Trim(.TextMatrix(.Row, mBillCol.c_批号))) = 8 Then
                        Dim strxq As String
                        
                        If IsNumeric(.TextMatrix(.Row, mBillCol.c_批号)) Then
                            strxq = UCase(.TextMatrix(.Row, mBillCol.c_批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq <> "" Then .TextMatrix(.Row, mBillCol.C_效期) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mBillCol.C_实盘数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
        
        lng批次 = Val(.TextMatrix(.Row, mBillCol.C_批次))
        If mint编辑状态 = 1 Then
            .ColData(mBillCol.C_产地) = IIf(lng批次 = -1 Or Val(.TextMatrix(.Row, mBillCol.c_产地编辑)) = 1, 1, 5)
            .ColData(mBillCol.c_批号) = IIf(lng批次 = -1 Or Val(.TextMatrix(.Row, mBillCol.c_批号编辑)) = 1, 4, 5)
            .ColData(mBillCol.C_效期) = IIf(lng批次 = -1, 2, 5)
        End If
        
        If mint编辑状态 = 2 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Then
            lng新批次 = Val(.TextMatrix(.Row, mBillCol.c_新批次))
            .ColData(mBillCol.C_产地) = IIf(lng新批次 = 1 Or lng批次 = -1 Or Val(.TextMatrix(.Row, mBillCol.c_产地编辑)) = 1, 1, 5)
            .ColData(mBillCol.c_批号) = IIf(lng新批次 = 1 Or lng批次 = -1 Or Val(.TextMatrix(.Row, mBillCol.c_批号编辑)) = 1, 4, 5)
            .ColData(mBillCol.C_效期) = IIf(lng新批次 = 1 Or lng批次 = -1, 2, 5)
        End If
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim dbl金额差, dbl差价差 As Double
    Dim i As Integer
    Dim int点击行 As Integer
    Dim strMoneyDigit As String
    int点击行 = mshBill.Row
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
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
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, True, True, True, , , , Me.txtCheckDate.Caption, , , mbln盘无存储库房材料, mstrPrivs, , False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                        
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)) Then
                            
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
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If Not SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)) Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    
                    Call 提示库存数
                End If
            Case mBillCol.C_产地
                If strKey = "" Then Exit Sub
                If SelectAndNotAddItem(Me, mshBill, strKey, "材料生产商", "材料生产商选择器", True, True, , zl_获取站点限制(True)) = True Then
                    .Text = .TextMatrix(.Row, .Col)
                Else
                    .Text = ""
                    .Col = mBillCol.C_产地
                    Cancel = True
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
                        .Col = mBillCol.C_实盘数量
                    End If
                    
                    Cancel = True
                    Exit Sub
                End If
            Case mBillCol.C_效期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            ShowMsgBox "失效期必须为日期型！"
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        ShowMsgBox "失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！"
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
            Case mBillCol.C_实盘数量
                Dim dbl成本价 As Double
                Dim rsTemp As New ADODB.Recordset
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    ShowMsgBox "实盘数量必须输入！"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    ShowMsgBox "实盘数量必须为数字型,请重输！"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    .TextMatrix(.Row, mBillCol.C_数量差) = Format(Abs(Val(strKey) - Val(.TextMatrix(.Row, mBillCol.C_帐面数量))), mFMT.FM_数量)
                    If Val(strKey) > Val(.TextMatrix(.Row, mBillCol.C_帐面数量)) Then
                        .TextMatrix(.Row, mBillCol.C_标志) = "盈"
                    ElseIf Val(strKey) < Val(.TextMatrix(.Row, mBillCol.C_帐面数量)) Then
                        .TextMatrix(.Row, mBillCol.C_标志) = "亏"
                    Else
                        .TextMatrix(.Row, mBillCol.C_标志) = "平"
                    End If
                    
                    If Val(.TextMatrix(.Row, mBillCol.C_帐面数量)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_金额
                    End If
                    
                    '金额差=当前售价*实盘数量-实际金额
                    '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
                    .TextMatrix(.Row, mBillCol.c_金额差) = Format(Val(.TextMatrix(.Row, mBillCol.C_售价)) * Val(strKey) - Val(.TextMatrix(.Row, mBillCol.C_实际金额)), strMoneyDigit)
                    .TextMatrix(.Row, mBillCol.c_差价差) = Format(Val(strKey) * (Val(.TextMatrix(.Row, mBillCol.C_售价)) - Val(.TextMatrix(.Row, mBillCol.C_成本价))) - Val(.TextMatrix(.Row, mBillCol.C_实际差价)), strMoneyDigit)
                    
                    dbl金额差 = .TextMatrix(.Row, mBillCol.c_金额差)
                    dbl差价差 = .TextMatrix(.Row, mBillCol.c_差价差)
                    
                    If .TextMatrix(.Row, mBillCol.C_标志) = "亏" Then    '保持与库存记录中的金额差、差价差的符号一致
                        If Val(.TextMatrix(.Row, mBillCol.C_实际金额)) >= 0 Then
                            .TextMatrix(.Row, mBillCol.c_金额差) = Format(Abs(.TextMatrix(.Row, mBillCol.c_金额差)), strMoneyDigit)
                        Else
                            .TextMatrix(.Row, mBillCol.c_金额差) = Format(Abs(.TextMatrix(.Row, mBillCol.c_金额差)) * -1, strMoneyDigit)
                        End If
                        If Val(.TextMatrix(.Row, mBillCol.C_实际差价)) >= 0 Then
                            .TextMatrix(.Row, mBillCol.c_差价差) = Format(Abs(.TextMatrix(.Row, mBillCol.c_差价差)), strMoneyDigit)
                        Else
                            .TextMatrix(.Row, mBillCol.c_差价差) = Format(Abs(.TextMatrix(.Row, mBillCol.c_差价差)) * -1, strMoneyDigit)
                        End If
                    End If
                    .TextMatrix(.Row, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_售价)) * Val(strKey), mFMT.FM_金额)
                    .TextMatrix(.Row, mBillCol.C_盘点成本金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_实际金额)) + dbl金额差 - (Val(.TextMatrix(.Row, mBillCol.C_实际差价)) + dbl差价差), mFMT.FM_金额)
                    .TextMatrix(.Row, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(.Row, mBillCol.c_金额差)) - Val(.TextMatrix(.Row, mBillCol.c_差价差)), mFMT.FM_金额)
                    
                End If
                Call 显示合计金额
        End Select
    End With
    Exit Sub
errHandle:
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

Private Function ValidData() As Boolean
    ValidData = False
    Dim lngLop As Long
    Dim lng效期 As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "单据号超长,最多能输入" & CInt(txtNO.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNO.MaxLength & "个字符!"
            txtNO.SetFocus
            Exit Function
        End If
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                ShowMsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!"
                txt摘要.SetFocus
                Exit Function
            End If
        
            For lngLop = 1 To .Rows - 1
                If Trim(.TextMatrix(lngLop, mBillCol.C_材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(lngLop, mBillCol.C_实盘数量))) = "" Then
                        ShowMsgBox "第" & lngLop & "行卫生材料的实盘数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_实盘数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngLop, mBillCol.C_实盘数量)) > 9999999999# Then
                        ShowMsgBox "第" & lngLop & "行卫生材料的实盘数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_实盘数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(lngLop, mBillCol.c_金额差)) > 9999999999999# Then
                        ShowMsgBox "第" & lngLop & "行卫生材料的金额差大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_实盘数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngLop, mBillCol.C_数量差)) > 9999999999999# Then
                        ShowMsgBox "第" & lngLop & "行卫生材料的数量差大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_实盘数量
                        Exit Function
                    End If
                
                    If Val(.TextMatrix(lngLop, mBillCol.C_批次)) = -1 Or Val(.TextMatrix(lngLop, mBillCol.c_新批次)) = 1 Then '分批材料必须录入分批信息
                        If LenB(StrConv(Trim(Trim(.TextMatrix(lngLop, mBillCol.c_批号))), vbFromUnicode)) > mintBatchNoLen Then
                            ShowMsgBox "第" & lngLop & "行卫生材料的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!"
                            .SetFocus
                            .Row = lngLop
                            .MsfObj.TopRow = lngLop
                            .Col = mBillCol.c_批号
                            Exit Function
                        End If
                        
                        '判断是否为效期卫生材料
                        gstrSQL = "Select Nvl(最大效期,0) 效期 From 材料特性 Where 材料ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否为效期卫生材料", Val(.TextMatrix(lngLop, 0)))
                        
                        lng效期 = rsTemp!效期
                        If lng效期 <> 0 Then
                            If Trim(.TextMatrix(lngLop, mBillCol.c_批号)) = "" Or Trim(.TextMatrix(lngLop, mBillCol.C_效期)) = "" Then
                                ShowMsgBox "第" & lngLop & "行的卫生材料是效期材料,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                If .TextMatrix(lngLop, mBillCol.c_批号) = "" Then
                                    .Col = mBillCol.c_批号
                                Else
                                    .Col = mBillCol.C_效期
                                End If
                                Exit Function
                            End If
                        End If
                        
                        '判断产地和批次是否为空
                        If mbln分批卫材批号产地控制 = True Then
                            If Trim(.TextMatrix(lngLop, mBillCol.C_产地)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & lngLop & "行卫生材料是分批材料，请录入产地！"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.C_产地
                                Exit Function
                            End If
                            If Trim(.TextMatrix(lngLop, mBillCol.c_批号)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & lngLop & "行卫生材料是分批材料，请录入批号！"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.c_批号
                                Exit Function
                            End If
                        End If
                        
                    End If
                    
                    If Val(.TextMatrix(lngLop, mBillCol.C_批次)) > 0 Then '已有批次
                        '判断产地和批次是否为空
                        If mbln分批卫材批号产地控制 = True Then
                            If Trim(.TextMatrix(lngLop, mBillCol.C_产地)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & lngLop & "行卫生材料是分批材料，请录入产地！"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.C_产地
                                Exit Function
                            End If
                            If Trim(.TextMatrix(lngLop, mBillCol.c_批号)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & lngLop & "行卫生材料是分批材料，请录入批号！"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.c_批号
                                Exit Function
                            End If
                        End If
                    End If
                    
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
    Dim lng入出类别ID As Long
    Dim int入出系数 As Integer
    Dim lng入库类别ID As Integer
    Dim lng出库类别ID As Integer
    
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房id As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次ID As Long
    Dim str产地 As String
    Dim dat效期 As String
    Dim dbl帐面数量 As Double
    Dim dbl实盘数量 As Double
    Dim dbl数量差 As Double
    Dim dbl售价 As Double
    Dim dbl成本价  As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim dat填制日期 As String
    Dim str盘点时间 As String
    Dim dbl库存金额 As Double
    Dim dbl库存差价 As Double
    Dim rsTemp As New Recordset
    Dim lngRow As Long
    Dim strArr As Variant
    Dim i As Long
    Dim cllSQL As Collection
    Dim int新批次 As Integer
    Dim n As Long
    
    On Error GoTo errHandle
    SaveCard = False
    '在外面设置入出类别ID，主要是所有材料都要用他
    gstrSQL = "" & _
        "   SELECT b.系数,b.id AS 类别id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID " & _
        "       AND a.单据 =[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 37)
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫生材料盘点管理的入出类别，请在入出分类中设置!"
        Exit Function
    End If
    
    lng入库类别ID = 0
    lng出库类别ID = 0
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!系数 = 1 Then
            lng入库类别ID = rsTemp!类别ID
        Else
            lng出库类别ID = rsTemp!类别ID
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If lng入库类别ID = 0 Then
        ShowMsgBox "没有设置卫生材料盘点管理的入库类别，请在入出分类中设置!"
        Exit Function
    End If
    If lng出库类别ID = 0 Then
        ShowMsgBox "没有设置卫生材料盘点管理的出库类别，请在入出分类中设置!"
        Exit Function
    End If
    
    Set cllSQL = New Collection
    With mshBill
        lng库房id = txtStock.Tag
        
        chrNo = Trim(txtNO)
        If mint编辑状态 = 1 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Then 'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(75, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(75, lng库房id)
            If IsNull(chrNo) Then Exit Function
        End If
        
        txtNO.Tag = chrNo
        
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        dat填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str盘点时间 = txtCheckDate.Caption
        
        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_材料盘点_Delete('" & mstr单据号 & "')"
            AddArray cllSQL, gstrSQL
        End If
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            lngRow = recSort!行号
'        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                lng材料ID = .TextMatrix(lngRow, 0)
                str产地 = .TextMatrix(lngRow, mBillCol.C_产地)
                str批号 = .TextMatrix(lngRow, mBillCol.c_批号)
                lng批次ID = IIf(.TextMatrix(lngRow, mBillCol.C_批次) = "", 0, .TextMatrix(lngRow, mBillCol.C_批次))
    
                int新批次 = 0
                If Val(.TextMatrix(lngRow, mBillCol.C_批次)) = -1 Or Val(.TextMatrix(lngRow, mBillCol.c_新批次)) = 1 Then
                    int新批次 = 1
                End If
                
                dat效期 = IIf(.TextMatrix(lngRow, mBillCol.C_效期) = "", "", .TextMatrix(lngRow, mBillCol.C_效期))
                dat效期 = IIf(.TextMatrix(lngRow, mBillCol.C_效期) = "", "", .TextMatrix(lngRow, mBillCol.C_效期))
                
                dbl帐面数量 = Round(Val(.TextMatrix(lngRow, mBillCol.C_帐面数量)) * Val(.TextMatrix(lngRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl实盘数量 = Round(Val(.TextMatrix(lngRow, mBillCol.C_实盘数量)) * Val(.TextMatrix(lngRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl数量差 = Round(Val(.TextMatrix(lngRow, mBillCol.C_数量差)) * Val(.TextMatrix(lngRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl成本价 = Round(.TextMatrix(lngRow, mBillCol.C_成本价) / Val(.TextMatrix(lngRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.成本价小数)
                dbl售价 = Round(Val(.TextMatrix(lngRow, mBillCol.C_售价)) / Val(.TextMatrix(lngRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.零售价小数)
                
                If dbl实盘数量 = 0 Then
                    dbl金额差 = Round(Val(.TextMatrix(lngRow, mBillCol.c_金额差)), g_小数位数.obj_最大小数.金额小数)
                    dbl差价差 = Round(Val(.TextMatrix(lngRow, mBillCol.c_差价差)), g_小数位数.obj_最大小数.金额小数)
                    dbl库存金额 = Round(Val(.TextMatrix(lngRow, mBillCol.C_实际金额)), g_小数位数.obj_最大小数.金额小数)
                    dbl库存差价 = Round(Val(.TextMatrix(lngRow, mBillCol.C_实际差价)), g_小数位数.obj_最大小数.金额小数)
                Else
                    dbl金额差 = Round(Val(.TextMatrix(lngRow, mBillCol.c_金额差)), g_小数位数.obj_最大小数.金额小数)
                    dbl差价差 = Round(Val(.TextMatrix(lngRow, mBillCol.c_差价差)), g_小数位数.obj_最大小数.金额小数)
                    dbl库存金额 = Round(Val(.TextMatrix(lngRow, mBillCol.C_实际金额)), g_小数位数.obj_最大小数.金额小数)
                    dbl库存差价 = Round(Val(.TextMatrix(lngRow, mBillCol.C_实际差价)), g_小数位数.obj_最大小数.金额小数)
                End If
                If dbl帐面数量 <= dbl实盘数量 Then
                    lng入出类别ID = lng入库类别ID
                    int入出系数 = 1
                Else
                    lng入出类别ID = lng出库类别ID
                    int入出系数 = -1
                End If
                 
                lng序号 = lngRow
                'zl_材料盘点_INSERT
                '    No_In         In 药品收发记录.NO%Type,
                '    序号_In       In 药品收发记录.序号%Type,
                '    库房id_In     In 药品收发记录.库房id%Type,
                '    批次_In       In 药品收发记录.批次%Type,
                '    入出类别id_In In 药品收发记录.入出类别id%Type,
                '    入出系数_In   In 药品收发记录.入出系数%Type,
                '    材料id_In     In 药品收发记录.药品id%Type,
                '    帐面数量_In   In 药品收发记录.填写数量%Type,
                '    实盘数量_In   In 药品收发记录.扣率%Type,
                '    数量差_In     In 药品收发记录.实际数量%Type,
                '    成本价_In     In 药品收发记录.单量%Type,
                '    售价_In       In 药品收发记录.零售价%Type,
                '    金额差_In     In 药品收发记录.零售金额%Type,
                '    差价差_In     In 药品收发记录.差价%Type,
                '    填制人_In     In 药品收发记录.填制人%Type,
                '    填制日期_In   In 药品收发记录.填制日期%Type,
                '    摘要_In       In 药品收发记录.摘要%Type := Null,
                '    产地_In       In 药品收发记录.产地%Type := Null,
                '    批号_In       In 药品收发记录.批号%Type := Null,
                '    效期_In       In 药品收发记录.效期%Type := Null,
                '    盘点时间_In   In 药品收发记录.频次%Type := Null,
                '    库存金额_In   In 药品收发记录.成本价%Type := Null,
                '    库存差价_In   In 药品收发记录.成本金额%Type := Null
                '    新批次_In     In Number := 0
                gstrSQL = "zl_材料盘点_INSERT('" & _
                    chrNo & "'," & _
                    lng序号 & "," & _
                    lng库房id & "," & _
                    lng批次ID & "," & _
                    lng入出类别ID & "," & _
                    int入出系数 & "," & _
                    lng材料ID & "," & _
                    dbl帐面数量 & "," & _
                    dbl实盘数量 & "," & _
                    dbl数量差 & "," & _
                    dbl成本价 & "," & _
                    dbl售价 & "," & _
                    dbl金额差 & "," & _
                    dbl差价差 & ",'" & _
                    str填制人 & "',to_date('" & _
                    dat填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                    str摘要 & "','" & _
                    str产地 & "','" & _
                    str批号 & "'," & _
                    IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    str盘点时间 & "'," & _
                    dbl库存金额 & "," & _
                    dbl库存差价 & "," & _
                    int新批次 & ")"
                AddArray cllSQL, gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        If mint编辑状态 = 5 Then
            '刘兴宏:20060801
            '删除或更改盘存过程中的盘点记录单
            strArr = Split(mstr盘点单号, ",")
            
            For i = 0 To UBound(strArr)
                
                If mbln删除盘点单 Then
                    'Zl_材料盘点记录单_DELETE:
                    '   NO_IN
                    gstrSQL = "Zl_材料盘点记录单_DELETE(" & strArr(i) & ")"
                Else
                    'Zl_材料盘点记录单_Update:
                    '   NO_IN
                    gstrSQL = "Zl_材料盘点记录单_Update(" & strArr(i) & ")"
                End If
                AddArray cllSQL, gstrSQL
            Next
        End If
        
    End With
        
    '执行相关SQL
    Call ExecuteProcedureArrAy(cllSQL, mstrCaption)
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
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
    Dim dbl成本盘点金额 As Double
    Dim lngLop As Long
    
    dbl金额差 = 0
    dbl盘点金额 = 0
    dbl成本盘点金额 = 0
    
    With mshBill
        For lngLop = 1 To .Rows - 1
            If .TextMatrix(lngLop, 0) <> "" Then
                
                dbl金额差 = dbl金额差 + Val(.TextMatrix(lngLop, mBillCol.c_金额差)) * IIf(.TextMatrix(lngLop, mBillCol.C_标志) = "亏", -1, 1)
                dbl盘点金额 = dbl盘点金额 + Val(.TextMatrix(lngLop, mBillCol.C_盘点金额))
                dbl成本盘点金额 = dbl成本盘点金额 + Val(.TextMatrix(lngLop, mBillCol.C_盘点成本金额))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "金额差合计：" & Format(dbl金额差, mFMT.FM_金额)
    lblPurchasePrice.Width = Pic单据.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200
    
    lblCheckSum.Caption = "盘点金额合计：" & Format(dbl盘点金额, mFMT.FM_金额)
    lblCheckSum.Width = Pic单据.TextWidth(lblCheckSum.Caption)
    
    lblCheckCostSum.Top = lblCheckSum.Top
    lblCheckCostSum.Left = lblCheckSum.Left + lblCheckSum.Width + 200
    lblCheckCostSum.Caption = "盘点成本金额合计：" & Format(dbl成本盘点金额, mFMT.FM_金额)
    lblCheckCostSum.Width = Pic单据.TextWidth(lblCheckCostSum.Caption)
    
End Sub

Private Sub 提示库存数()
    Dim rsTemp As New Recordset
    Dim strKc As String
    
    On Error GoTo errHandle
    '取库存
    '20060731:刘兴宏加入，主要解决盘点时间的库存
    strKc = "" & _
        "   SELECT " & _
        "           nvl(a.可用数量,0)/[5] 可用数量,nvl(a.实际数量,0)/[5] 实际数量,a.实际金额, a.实际差价" & _
        "   FROM 药品库存 a" & _
        "   Where a.药品id=[2] and nvl(a.批次,0)=[3] " & _
        "           AND a.性质=1 " & _
        "           AND a.库房id =[1] "
           
    
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
       ' gstrSQL = "" & _
            "   Select 可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as  可用数量 " & _
            "   From 药品库存 where 库房id=[1]" & _
            "       and 药品id=[2]" & _
            "       and 性质=1 " & _
            "       and  nvl(批次,0)=[3]"
        gstrSQL = strKc
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提示库存数", Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.C_批次)), CDate(txtCheckDate.Caption), Val(.TextMatrix(.Row, mBillCol.c_比例系数)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mBillCol.C_可用数量) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_可用数量) = IIf(IsNull(rsTemp.Fields(1)), 0, rsTemp.Fields(1))
        End If
        rsTemp.Close
        
        stbThis.Panels(2).Text = "该卫生材料当前库存数为[" & Format(.TextMatrix(.Row, mBillCol.C_可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.c_单位)
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

Private Function SetPhiscRows(ByVal lngId As Long, ByVal lng批次 As Long) As Boolean
'功能：根据材料ID在盘存表上显示并处理该材料的初始盘存信息
'说明：
'   1.如果是非库房分批药,且已经输入了,则提示并退出。
'   2.如果是库房分批药，则分别处理该药的未处理的各批次库存行。
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim lngRecordCount As Long
    Dim lngRow As Long
    Dim bln库房 As Boolean
    Dim dbl成本价 As Double
    Dim dbl指导差价率 As Double
    Dim lngBatch As Long
    Dim rsprice As New Recordset
    Dim lngTmp As Long
    Dim dbl金额差, dbl差价差 As Double
    Dim strMoneyDigit As String
    
    On Error GoTo errH
    
    SetPhiscRows = False
    Set rsData = GetDateStock(txtCheckDate.Caption, txtStock.Tag, 0, True, , , lngId)
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then Exit Function
    
    bln库房 = CheckPartProp(Val(txtStock.Tag))
    '新增批次药品
    If lng批次 <> -1 Then
        rsData.MoveFirst
        rsData.Find "批次=" & lng批次
        If rsData.EOF Then Exit Function
    End If
    
    With mshBill
        '检查单据是否存在对应卫材
        If lng批次 <> -1 Then
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" Then
                    If .TextMatrix(lngRow, 0) = rsData!材料ID And IIf(.TextMatrix(lngRow, mBillCol.C_批次) = "", "0", .TextMatrix(lngRow, mBillCol.C_批次)) = lng批次 Then
                        If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & .TextMatrix(lngRow, mBillCol.C_材料) & "，"  '最多记录三个重复的卫材
    '                    MsgBox "已有卫生材料【" & .TextMatrix(lngRow, mBillCol.C_材料) & "(" & lng批次 & ")】，不再添加！", vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        End If
        
        mshBill.Redraw = False
        lngRow = .Row
        .TextMatrix(lngRow, 0) = rsData!材料ID
        
        '取出该材料的成本价
        'gstrSQL = "Select Nvl(成本价,0) 成本价,nvl(指导差价率,0) From 材料特性 Where 材料ID=[1]"
        'Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出该卫生材料的成本价", Val(NVL(rsData!材料ID)))
                
        dbl成本价 = Val(zlStr.NVL(rsData!最后进价))
        dbl指导差价率 = Val(zlStr.NVL(rsData!指导差价率))
            
        .TextMatrix(lngRow, mBillCol.C_材料) = "[" & rsData!编码 & "]" & rsData!商品名称
        .TextMatrix(lngRow, mBillCol.c_规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
        .TextMatrix(lngRow, mBillCol.C_产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
        .TextMatrix(lngRow, mBillCol.C_批准文号) = IIf(IsNull(rsData!批准文号), "", rsData!批准文号)
        .TextMatrix(lngRow, mBillCol.C_库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
        .TextMatrix(lngRow, mBillCol.c_单位) = IIf(IsNull(rsData!单位), "", rsData!单位)
        .TextMatrix(lngRow, mBillCol.C_批次) = IIf(IsNull(rsData!批次), "0", rsData!批次)
        
        If Val(.TextMatrix(lngRow, mBillCol.C_批次)) <> 0 Then
            .TextMatrix(lngRow, mBillCol.c_批号编辑) = rsData!批号编辑
            .TextMatrix(lngRow, mBillCol.c_产地编辑) = rsData!产地编辑
        End If
            
        If CheckPhysicBatch(bln库房, rsData!库房分批, rsData!在用分批) And Val(.TextMatrix(lngRow, mBillCol.C_批次)) = 0 Then
            .TextMatrix(lngRow, mBillCol.C_批次) = -1
        End If
        
        If lng批次 = -1 Then
            .TextMatrix(lngRow, mBillCol.C_批次) = lng批次
            .TextMatrix(lngRow, mBillCol.c_批号) = ""
            .TextMatrix(lngRow, mBillCol.C_效期) = ""
            .TextMatrix(lngRow, mBillCol.C_帐面数量) = Format(0, mFMT.FM_数量)
            .TextMatrix(lngRow, mBillCol.C_实盘数量) = .TextMatrix(lngRow, mBillCol.C_帐面数量)
            .TextMatrix(lngRow, mBillCol.C_盘点金额) = Format(0, mFMT.FM_金额)
            .TextMatrix(lngRow, mBillCol.C_可用数量) = 0
            .TextMatrix(lngRow, mBillCol.C_实际金额) = 0
            .TextMatrix(lngRow, mBillCol.C_实际差价) = 0
            .TextMatrix(lngRow, mBillCol.C_售价) = Format(IIf(IsNull(rsData!售价), 0, rsData!售价), mFMT.FM_零售价)
            .TextMatrix(lngRow, mBillCol.C_成本价) = Format(dbl成本价, mFMT.FM_成本价)
            .ColData(mBillCol.c_批号) = 4
            .ColData(mBillCol.C_效期) = 2
        Else
            lngBatch = Val(.TextMatrix(lngRow, mBillCol.C_批次))
            .ColData(mBillCol.c_批号) = IIf(lngBatch = -1, 4, 5)
            .ColData(mBillCol.C_效期) = IIf(lngBatch = -1, 2, 5)
            
            .TextMatrix(lngRow, mBillCol.c_批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
            .TextMatrix(lngRow, mBillCol.C_效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
            .TextMatrix(lngRow, mBillCol.C_帐面数量) = Format(IIf(IsNull(rsData!帐面数量), 0, rsData!帐面数量), mFMT.FM_数量)
            .TextMatrix(lngRow, mBillCol.C_实盘数量) = .TextMatrix(lngRow, mBillCol.C_帐面数量)
            .TextMatrix(lngRow, mBillCol.C_售价) = Format(IIf(IsNull(rsData!售价), 0, rsData!售价), mFMT.FM_零售价)
            .TextMatrix(lngRow, mBillCol.C_成本价) = Format(Val(zlStr.NVL(rsData!成本价)), mFMT.FM_成本价)
            .TextMatrix(lngRow, mBillCol.C_盘点金额) = Format(Val(.TextMatrix(lngRow, mBillCol.C_实盘数量)) * Val(.TextMatrix(lngRow, mBillCol.C_售价)), mFMT.FM_金额)
            
            .TextMatrix(lngRow, mBillCol.C_可用数量) = rsData!可用数量
            .TextMatrix(lngRow, mBillCol.C_实际金额) = rsData!实际金额
            .TextMatrix(lngRow, mBillCol.C_实际差价) = rsData!实际差价
            .TextMatrix(lngRow, mBillCol.C_成本价) = Format(Val(zlStr.NVL(rsData!成本价)), mFMT.FM_成本价)
        End If
        
        .TextMatrix(lngRow, mBillCol.c_比例系数) = rsData!比例系数
        .TextMatrix(lngRow, mBillCol.C_指导差价率) = rsData!指导差价率 & "||" & rsData!是否变价 & "||" & rsData!在用分批
        
        .TextMatrix(lngRow, mBillCol.C_标志) = "平"
        .TextMatrix(lngRow, mBillCol.C_数量差) = Format("0", mFMT.FM_金额)
            
        If rsData!是否变价 = 1 Then
            .TextMatrix(lngRow, mBillCol.C_售价) = Format(Get零售价(Val(zlStr.NVL(rsData!材料ID)), Val(txtStock.Tag), Val(zlStr.NVL(rsData!批次)), rsData!比例系数), mFMT.FM_零售价)
        End If
        
        .RowData(lngRow) = IIf(IsNull(rsData!最大效期), 0, rsData!最大效期)
        
        If Val(.TextMatrix(lngRow, mBillCol.C_帐面数量)) = 0 Then
            strMoneyDigit = "#0.00000"
        Else
            strMoneyDigit = mFMT.FM_金额
        End If
                    
        '金额差=当前售价*实盘数量-实际金额
        '差价差=金额差*iif(实际金额=0,指导差价率,(实际差价/实际金额))
        .TextMatrix(lngRow, mBillCol.c_金额差) = Format(Val(.TextMatrix(lngRow, mBillCol.C_售价)) * Val(.TextMatrix(lngRow, mBillCol.C_实盘数量)) - Val(.TextMatrix(lngRow, mBillCol.C_实际金额)), strMoneyDigit)
            
        If rsData!是否变价 = 1 And Val(.TextMatrix(lngRow, mBillCol.C_帐面数量)) = 0 Then
            .TextMatrix(lngRow, mBillCol.c_差价差) = Format(Val(.TextMatrix(lngRow, mBillCol.C_数量差)) * (Val(.TextMatrix(lngRow, mBillCol.C_售价)) - dbl成本价) - Val(.TextMatrix(lngRow, mBillCol.C_实际差价)), strMoneyDigit)
        Else
            .TextMatrix(lngRow, mBillCol.c_差价差) = Format(Val(.TextMatrix(lngRow, mBillCol.C_实盘数量)) * (Val(.TextMatrix(lngRow, mBillCol.C_售价)) - Val(.TextMatrix(lngRow, mBillCol.C_成本价))) - Val(.TextMatrix(lngRow, mBillCol.C_实际差价)), strMoneyDigit)
        End If
        
        dbl金额差 = .TextMatrix(lngRow, mBillCol.c_金额差)
        dbl差价差 = .TextMatrix(lngRow, mBillCol.c_差价差)
        
        .TextMatrix(lngRow, mBillCol.C_盘点成本金额) = Format(Val(.TextMatrix(lngRow, mBillCol.C_实际金额)) + dbl金额差 - (Val(.TextMatrix(lngRow, mBillCol.C_实际差价)) + dbl差价差), mFMT.FM_金额)
        .TextMatrix(lngRow, mBillCol.C_盘点成本金额差) = Format(Val(.TextMatrix(lngRow, mBillCol.c_金额差)) - Val(.TextMatrix(lngRow, mBillCol.c_差价差)), mFMT.FM_金额)
        
        Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
        mshBill.Redraw = True
    End With
    Call 提示库存数
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'在一行中插入
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim lngOldRows As Long
    Dim lngLop As Long
    Dim lngExchange As Long
    Dim intCol As Integer
    
    With mshBill
        blnHaveData = False
        lngOldRows = .Rows - 1
        .Rows = .Rows + intRecordCount
        For lngLop = intRow + 1 To intRecordCount
            If .TextMatrix(lngLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For lngExchange = .Rows - 1 To lngOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(lngExchange, intCol) = .TextMatrix(lngExchange - intRecordCount, intCol)
                    .TextMatrix(lngExchange - intRecordCount, intCol) = ""
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
    Dim strNo As String
    strNo = txtNO.Tag
    Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1719", mint记录状态, mintUnit, 1719, "卫生材料盘点表", strNo)
End Sub

Private Function CheckPartProp(ByVal lng库房id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '检查库房属性，如果是库房，返回真
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From 部门性质说明 " & _
        "   WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) " & _
        "           AND 部门id =[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng库房id)
    
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

Private Function CheckPhysicBatch(ByVal bln库房 As Boolean, ByVal int库房分批 As Integer, ByVal int在用分批 As Integer) As Boolean
    '返回该材料是否分批的标识
    CheckPhysicBatch = (bln库房 And (int库房分批 = 1)) Or (Not bln库房 And (int在用分批 = 1))
End Function

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Call zlDatabase.OpenRecordset(rsBatchNolen, gstrSQL, "取字段长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
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
                !批次 = Val(mshBill.TextMatrix(n, mBillCol.C_批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
