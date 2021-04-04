VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherOutputCard 
   Caption         =   "卫材其他出库单"
   ClientHeight    =   6960
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmOtherOutputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   31
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   30
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   0
      Width           =   11715
      Begin VB.ComboBox cbo外销单位 
         Height          =   300
         Left            =   7890
         TabIndex        =   5
         Text            =   "cbo外销单位"
         Top             =   600
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
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
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         Caption         =   "外销合计:"
         Height          =   180
         Left            =   6600
         TabIndex        =   33
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lbl外销单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "外销单位(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6840
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   23
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   7
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料其他出库单"
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入出类别(&T)"
         Height          =   180
         Left            =   3480
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
            Picture         =   "frmOtherOutputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1000
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
            Picture         =   "frmOtherOutputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   6600
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
            Picture         =   "frmOtherOutputCard.frx":22EA
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
            Picture         =   "frmOtherOutputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":3080
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
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmOtherOutputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mblnFirst As Boolean

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mint库存检查 As Integer             '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Dim mstrPrivs As String                     '权限

'刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private Const mlngModule = 1718
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容

Private mstrLike As String
Private Const mstrCaption As String = "卫材其他出库单"
Private mstr重复卫材 As String '记录重复的卫材

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol材料 As Integer = 2
Private Const mconIntCol序号 As Integer = 3
Private Const mconIntCol规格 As Integer = 4
Private Const mconIntCol可用数量 As Integer = 5
Private Const mconIntCol指导差价率 As Integer = 6
Private Const mconIntCol实际金额 As Integer = 7
Private Const mconIntCol实际差价 As Integer = 8
Private Const mconIntCol比例系数 As Integer = 9
Private Const mconIntCol批次 As Integer = 10
Private Const mconIntCol产地 As Integer = 11
Private Const mconIntCol批准文号 As Integer = 12
Private Const mconIntCol单位 As Integer = 13
Private Const mconIntCol批号 As Integer = 14
Private Const mconIntCol效期 As Integer = 15
Private Const mconIntCol灭菌失效期 As Integer = 16
Private Const mconIntCol数量 As Integer = 17
Private Const mconIntCol冲销数量 As Integer = 18
Private Const mconIntCol采购价 As Integer = 19
Private Const mconIntCol采购金额 As Integer = 20
Private Const mconIntCol售价 As Integer = 21
Private Const mconIntCol售价金额 As Integer = 22
Private Const mconintCol差价 As Integer = 23
Private Const mconintCol外销价 As Integer = 24
Private Const mconintCol外销金额 As Integer = 25
Private Const mconintCol增值税率 As Integer = 26
Private Const mconintCol税金 As Integer = 27
Private Const mconIntColS  As Integer = 28              '总列数


'=========================================================================================


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    GetDepend = False
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID AND A.单据 = 36"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "卫材其他出库"
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫材其他出库的出库类别，请在入出分类中设置！"
        rsTemp.Close
        Exit Function
    End If
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
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
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
        
    Call GetRegInFor(g私有模块, "卫材其他出库管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
         
     
    If mint编辑状态 = 1 Then
'        If mbln单据增加 Then
'            mstr单据号 = NextNo(74)
'        End If
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
        CmdSave.Caption = "冲销(&O)"
        cmdAllCls.Visible = True
        cmdAllSel.Visible = True
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

Private Sub cboType_Click()
    Me.lbl外销单位.Visible = False
    Me.cbo外销单位.Visible = False
    
    mshBill.ColData(mconintCol外销价) = 5
    mshBill.ColWidth(mconintCol外销价) = 0
    mshBill.ColWidth(mconintCol外销金额) = 0
    mshBill.ColWidth(mconintCol增值税率) = 0
    mshBill.ColWidth(mconintCol税金) = 0
        
    If cboType.Text = "材料外销" Then
        Me.lbl外销单位.Visible = True
        Me.cbo外销单位.Visible = True
        
        mshBill.ColWidth(mconintCol外销价) = 1000
        mshBill.ColWidth(mconintCol外销金额) = 1000
        mshBill.ColWidth(mconintCol增值税率) = 1000
        mshBill.ColWidth(mconintCol税金) = 1000
        cbo外销单位.Enabled = (mint编辑状态 = 1 Or mint编辑状态 = 2)
        mshBill.ColData(mconintCol外销价) = IIf(cbo外销单位.Enabled, 4, 5)
    End If
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo外销单位_GotFocus()
    If cbo外销单位.Style = 0 Then
        Call zlControl.TxtSelAll(cbo外销单位)
    End If
End Sub

Private Sub cbo外销单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cbo外销单位.Style = 2 And cbo外销单位.ListIndex <> -1 Then
            cbo外销单位.ListIndex = -1
        End If
    End If
End Sub


Private Sub cbo外销单位_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo外销单位.Locked And cbo外销单位.Style = 2 Then
            lngIdx = cbo.MatchIndex(cbo外销单位.hwnd, KeyAscii)
            If lngIdx = -1 And cbo外销单位.ListCount > 0 Then lngIdx = 0
            cbo外销单位.ListIndex = lngIdx
        End If
    End If
End Sub


Private Sub cbo外销单位_Validate(Cancel As Boolean)
    '功能：根据输入的内容,自动匹配
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo外销单位.ListIndex <> -1 Then Exit Sub '已选中
    If cbo外销单位.Text = "" Then cbo外销单位.Tag = "": Exit Sub '无输入
    
    strInput = UCase(NeedName(cbo外销单位.Text))
    strSQL = "Select Rownum As id,编码,简码,名称 From 材料外销单位 Where Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2] Order By 编码"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo外销单位.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "外调单位", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo外销单位.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo外销单位, zlStr.Nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!名称)
        If intIdx <> -1 Then
            cbo外销单位.ListIndex = intIdx
        Else
            cbo外销单位.AddItem zlStr.Nvl(rsTmp!编码) & "-" & Chr(13) & rsTmp!名称, cbo外销单位.ListCount - 1
            cbo外销单位.ListIndex = cbo外销单位.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的外调单位。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol冲销数量) = Format(0, mFMT.FM_数量)
                .TextMatrix(intRow, mconIntCol采购金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mconIntCol售价金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mconintCol差价) = Format(0, mFMT.FM_金额)
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
                .TextMatrix(intRow, mconIntCol冲销数量) = .TextMatrix(intRow, mconIntCol数量)
                .TextMatrix(intRow, mconIntCol采购金额) = Format(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol采购价), mFMT.FM_金额)
                .TextMatrix(intRow, mconIntCol售价金额) = Format(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol售价), mFMT.FM_金额)
                .TextMatrix(intRow, mconintCol差价) = Format(.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconIntCol采购金额), mFMT.FM_金额)
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
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                ShowMsgBox "该单据已没有可以冲销的卫材，请检查！"
            Else
                '单据已被删除
                ShowMsgBox "该单据已被删除，请检查！"
            End If
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            ShowMsgBox "该单据已被其他人审核，请检查！"
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
        
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        '刘兴宏:2007/06/10:问题10813
        mstrTime_End = GetBillInfo(21, txtNO.Tag)
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
        
        If Not 检查单价(21, txtNO.Tag, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
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
'    If mbln单据增加 Then
'        mstr单据号 = NextNo(74)
'        txtNO = mstr单据号
'    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)

    txt摘要.Text = ""
    If cboType.Enabled Then cboType.SetFocus
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
            " Where a.单据 = 21 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(b.现价, " & g_小数位数.obj_散装小数.零售价小数 & ") And" & _
              "    NVL(c.是否变价, 0) = 0" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 21 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & g_小数位数.obj_散装小数.零售价小数 & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 21 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & g_小数位数.obj_散装小数.成本价小数 & ")<>round(b.平均成本价," & g_小数位数.obj_散装小数.成本价小数 & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 材料id, 序号"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取当前价格]", CStr(Me.txtNO.Text))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol数量))
        dbl成本价 = Val(mshBill.TextMatrix(lngRow, mconIntCol采购价))
        dbl零售价 = Val(mshBill.TextMatrix(lngRow, mconIntCol售价))
        dbl成本金额 = dbl成本价 * dbl数量
        dbl零售金额 = dbl零售价 * dbl数量
        dbl差价 = dbl零售金额 - dbl成本金额
'
        If lng材料ID <> 0 Then
            rsprice.Filter = "类型='售价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mFMT.FM_零售价))
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            rsprice.Filter = "类型='成本价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl成本价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mFMT.FM_金额))
                dbl成本金额 = Val(Format(dbl成本价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mconIntCol售价) = Format(dbl零售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mconIntCol采购价) = Format(dbl成本价, mFMT.FM_成本价)
                mshBill.TextMatrix(lngRow, mconIntCol采购金额) = Format(dbl成本金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mconintCol差价) = Format(dbl差价, mFMT.FM_金额)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mintUnit = Val(strReg)
    
    mblnFirst = True
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    initGrid
    
    With cboType
        .Clear
        gstrSQL = "" & _
            "   SELECT b.Id,b.名称 " & _
            "   FROM 药品单据性质 A, 药品入出类别 B " & _
            "   Where A.类别id = B.ID AND A.单据 = 36 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo外销单位
        .Clear
        gstrSQL = "Select Rownum As id,编码,简码,名称 From 材料外销单位 Order By 编码"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "读取外销单位")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & rsTemp!名称
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    Call initCard
    
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 800, 0)
        .ColWidth(mconintCol外销价) = IIf(cboType.Text = "材料外销", 1000, 0)
        .ColWidth(mconintCol外销金额) = IIf(cboType.Text = "材料外销", 1000, 0)
        .ColWidth(mconintCol增值税率) = IIf(cboType.Text = "材料外销", 1000, 0)
        .ColWidth(mconintCol税金) = IIf(cboType.Text = "材料外销", 1000, 0)
        
        .ColWidth(mconIntCol采购价) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconIntCol采购金额) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconintCol差价) = IIf(mblnCostView = True, 900, 0)
    End With
    mblnChange = False
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
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str批次 As String, strArray As String
    
    '库房
    On Error GoTo ErrHandle
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
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 6
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   Where a.库房id=b.id and A.单据 = 21 and a.no=[1]"
                
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
                    strUnitQuantity = "c.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,nvl(a.单量,0) As 外销价,'1' as 比例系数,"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.填写数量 / B.换算系数) AS 填写数量,(A.实际数量 / B.换算系数) AS 实际数量,a.成本价*B.换算系数 as 成本价,a.零售价*B.换算系数 as 零售价,nvl(a.单量,0)*B.换算系数 As 外销价,B.换算系数 as 比例系数,"
            End Select
            
            If mint编辑状态 <> 6 Then
                    gstrSQL = "" & _
                    "   Select w.*,z.可用数量/w.比例系数 可用数量,z.实际金额,z.实际差价 " & _
                    "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || c.编码 || ']' ||c.名称) AS 卫材信息," & _
                    "                   zlSpellCode(c.名称) 名称,c.规格,c.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,a.效期," & _
                                        strUnitQuantity & _
                    "                   A.成本金额,A.零售金额, A.差价, " & _
                    "                   a.摘要,填制人,a.填制日期,a.审核人,审核日期,a.库房id,a.入出类别id,c.是否变价,b.在用分批,d.名称 AS 外销单位, To_Number(Trim(To_Char(Nvl(A.频次, '0'), '999999999999.0000'))) As 增值税率 " & _
                    "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 c,材料外销单位 D " & _
                    "           Where A.药品id = B.材料id and a.药品id=c.id  " & _
                    "                   AND A.记录状态 =[3] " & _
                    "                   AND A.单据 = 21 AND A.No = [1] And A.发药窗口=D.编码(+) " & _
                    "           ) w,(   Select  药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "                   From 药品库存 where 库房id=[2]  and 性质=1)  z " & _
                    "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0) " & _
                    " ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Else
                    gstrSQL = "" & _
                    "   Select w.*,z.可用数量/w.比例系数 可用数量,z.实际金额,z.实际差价 " & _
                    "   From (  SELECT distinct a.材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息," & _
                    "                   zlSpellCode(c.名称) 名称,c.规格,c.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,a.效期," & _
                                        strUnitQuantity & _
                    "                   A.成本金额,0 零售金额,0 差价, " & _
                    "                   a.摘要,a.库房id,a.入出类别id,c.是否变价,b.在用分批,d.名称 AS 外销单位,A.增值税率 " & _
                    "           FROM (  Select min(id) as id, sum(实际数量) as 填写数量,0 实际数量,sum(成本金额) as 成本金额,药品id 材料ID,序号,产地,批准文号, 批号,效期," & _
                    "           Nvl(批次,0) 批次,扣率,成本价,零售价,摘要,库房ID,入出类别ID,单量,发药窗口, To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) As 增值税率" & _
                    "                   From 药品收发记录 x " & _
                    "                   WHERE NO=[1] AND 单据=21  " & _
                    "                   Group by 药品ID,序号,产地,批准文号,批号,效期,Nvl(批次,0),扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID,单量,发药窗口, To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) " & _
                    "                   having sum(填写数量)<>0 " & _
                    "               ) A, 材料特性 B,收费项目目录 c,材料外销单位 d " & _
                    "           Where A.材料id = B.材料id and a.材料id=c.id And A.发药窗口=d.编码(+) " & _
                    "       ) w,(Select  药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "            From 药品库存 " & _
                    "            Where 库房id=[2]  and 性质=1)  z " & _
                    "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, cboStock.ItemData(cboStock.ListIndex), mint记录状态)
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            '刘兴宏:2007/06/10:问题10813
            mstrTime_Start = GetBillInfo(21, mstr单据号)
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!入出类别ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
                
                If .Text = "材料外销" Then
                    Me.cbo外销单位.Visible = True
                    
                    '定位外销单位
                    If Not IsNull(rsTemp!外销单位) Then
                        For i = 1 To cbo外销单位.ListCount - 1
                            If Mid(cbo外销单位.List(i), InStr(1, cbo外销单位.List(i), "-") + 1) = rsTemp!外销单位 Then
                                cbo外销单位.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If
            End With
            
            Select Case mint编辑状态
            Case 2, 6
                Txt填制人 = UserInfo.用户名
                Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint编辑状态 = 6 Then
                    Txt审核人 = UserInfo.用户名
                    Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt填制人 = rsTemp!填制人
                Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
                Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = intRow + 1
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mconIntCol材料) = rsTemp!卫材信息
                    .TextMatrix(intRow, mconIntCol序号) = rsTemp!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                    .TextMatrix(intRow, mconIntCol单位) = rsTemp!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol数量) = Format(rsTemp!填写数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mconIntCol采购价) = Format(rsTemp!成本价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mconIntCol采购金额) = Format(IIf(mint编辑状态 = 6, 0, rsTemp!成本金额), mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol售价) = Format(rsTemp!零售价, mFMT.FM_零售价)
                    .TextMatrix(intRow, mconIntCol售价金额) = Format(rsTemp!零售金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mconintCol差价) = Format(rsTemp!差价, mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsTemp!比例系数
                    .TextMatrix(intRow, mconIntCol指导差价率) = rsTemp!指导差价率 & "||" & rsTemp!是否变价 & "||" & rsTemp!在用分批
                    .TextMatrix(intRow, mconIntCol可用数量) = IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量)
                    .TextMatrix(intRow, mconIntCol实际差价) = IIf(IsNull(rsTemp!实际差价), "0", rsTemp!实际差价)
                    .TextMatrix(intRow, mconIntCol实际金额) = IIf(IsNull(rsTemp!实际金额), "0", rsTemp!实际金额)
                    
                    .TextMatrix(intRow, mconintCol外销价) = Format(rsTemp!外销价, mFMT.FM_零售价)
                    .TextMatrix(intRow, mconintCol增值税率) = GetFormat(IIf(IsNull(rsTemp!增值税率), "0", rsTemp!增值税率), 2)
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconintCol外销金额) = Format(0, mFMT.FM_金额)
                        .TextMatrix(intRow, mconintCol税金) = Format(0, mFMT.FM_金额)
                    Else
                        .TextMatrix(intRow, mconintCol外销金额) = Format(rsTemp!外销价 * rsTemp!填写数量, mFMT.FM_金额)
                        .TextMatrix(intRow, mconintCol税金) = Format(rsTemp!外销价 * rsTemp!填写数量 * (Val(.TextMatrix(intRow, mconintCol增值税率)) / 100 / (1 + Val(.TextMatrix(intRow, mconintCol增值税率)) / 100)), mFMT.FM_金额)
                    End If
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsTemp!材料ID & IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str批次 = rsTemp!材料ID & IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                        If mint编辑状态 = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!填写数量), "0", rsTemp!填写数量)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!实际数量), "0", rsTemp!实际数量)
                        End If
                        mcolUsedCount.Add Array(str批次, strArray), str批次
                    End If
                    
                    rsTemp.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
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
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol材料) = "名称与编码"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = "失效期"
        .TextMatrix(0, mconIntCol灭菌失效期) = "灭菌失效期"
        
        .TextMatrix(0, mconIntCol数量) = IIf(mint编辑状态 = 6, "数量", "填写数量")
        .TextMatrix(0, mconIntCol冲销数量) = "冲销数量"
        .TextMatrix(0, mconIntCol采购价) = "成本价"
        .TextMatrix(0, mconIntCol采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconIntCol指导差价率) = "指导差价率"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol批次) = "批次"
        
        .TextMatrix(0, mconintCol外销价) = "外销价"
        .TextMatrix(0, mconintCol外销金额) = "外销金额"
        .TextMatrix(0, mconintCol增值税率) = "增值税率%"
        .TextMatrix(0, mconintCol税金) = "税金"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol材料) = 2000
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol灭菌失效期) = 1000
        .ColWidth(mconIntCol数量) = 800
        .ColWidth(mconIntCol冲销数量) = 0
        .ColWidth(mconIntCol采购价) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mconIntCol采购金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol售价) = 800
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mconIntCol可用数量) = 0
        
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        .ColWidth(mconIntCol指导差价率) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol批次) = 0
         
        .ColWidth(mconintCol外销价) = 0
        .ColWidth(mconintCol外销金额) = 0
        .ColWidth(mconintCol增值税率) = 0
        .ColWidth(mconintCol税金) = 0
        
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol灭菌失效期) = 5
        
        .ColData(mconintCol外销金额) = 5
        .ColData(mconintCol增值税率) = 5
        .ColData(mconintCol税金) = 5
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            cboType.Enabled = True
            txt摘要.Enabled = True
            cboStock.Enabled = True
            .ColData(mconIntCol材料) = 1
            .ColData(mconIntCol数量) = 4
            .ColData(mconIntCol冲销数量) = 5
            
            .ColData(mconintCol外销价) = IIf(Me.cbo外销单位.Visible, 4, 5)
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 6 Then
            cboStock.Enabled = False
            
            cboType.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol数量) = 5
            .ColData(mconIntCol冲销数量) = 4
        ElseIf mint编辑状态 = 4 Then
            cboStock.Enabled = False
            
            cboType.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol数量) = 5
            .ColData(mconIntCol冲销数量) = 5
            
        End If
        
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        
        .ColData(mconIntCol可用数量) = 5
        
        .ColData(mconIntCol实际差价) = 5
        .ColData(mconIntCol实际金额) = 5
        .ColData(mconIntCol指导差价率) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol批次) = 5
        
        .ColAlignment(mconIntCol材料) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol冲销数量) = flexAlignRightCenter
        
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        .ColAlignment(mconIntCol灭菌失效期) = flexAlignCenterCenter
        
        .PrimaryCol = mconIntCol材料
        .LocateCol = mconIntCol材料
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mconIntCol材料) = 0
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
        lblNO.Left = .Left - lblNO.Width - 100
        .Top = LblTitle.Top
        lblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cbo外销单位.Left = mshBill.Left + mshBill.Width - cbo外销单位.Width
    lbl外销单位.Left = cbo外销单位.Left - lbl外销单位.Width - 100
    
    lblType.Left = cboStock.Left + cboStock.Width + (lbl外销单位.Left - cboStock.Left - cboStock.Width - (lblType.Width + cboType.Width + 100)) / 2
    cboType.Left = lblType.Left + lblType.Width + 100
    
'    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
'    lblType.Left = cboType.Left - lblType.Width - 100
    
    
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
    End With
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
        lblOther.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With lblOther
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 3
    End With
    
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
    Dim intRow As Integer
    Dim strNo As String
    Dim lng库房ID As Long
    Dim str审核人 As String
    Dim dat审核日期 As String
    
    Dim int序号 As Integer
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim dbl数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim lng入出类别ID As Long
    Dim n As Long
    Dim arrSQL As Variant
    
    
    arrSQL = Array()
    
    mblnSave = False
    SaveCheck = False
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng入出类别ID = cboType.ItemData(cboType.ListIndex)
    str审核人 = UserInfo.用户名
    strNo = txtNO.Tag
    
    dat审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo ErrHandle
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
'                If Val(.TextMatrix(intRow, mconIntCol实际金额)) = 0 Then
'                   .TextMatrix(intRow, mconintCol差价) = Format(Val(.TextMatrix(intRow, mconIntCol售价金额)) * Split(.TextMatrix(.Row, mconIntCol指导差价率), "||")(0) / 100, mFMT.FM_金额)
'                Else
'                   .TextMatrix(intRow, mconintCol差价) = Format(Val(.TextMatrix(intRow, mconIntCol售价金额)) * (Val(.TextMatrix(intRow, mconIntCol实际差价)) / Val(.TextMatrix(intRow, mconIntCol实际金额))), mFMT.FM_金额)
'                End If
'
'                If Val(.TextMatrix(intRow, mconIntCol数量)) = 0 Then
'                    .TextMatrix(intRow, mconIntCol采购价) = 0
'                Else
'                    .TextMatrix(intRow, mconIntCol采购价) = Format((Val(.TextMatrix(intRow, mconIntCol售价金额)) - Val(.TextMatrix(intRow, mconintCol差价))) / (Val(.TextMatrix(intRow, mconIntCol数量))), mFMT.FM_成本价)
'                End If
'
'                .TextMatrix(intRow, mconIntCol采购金额) = Format(Val(.TextMatrix(intRow, mconIntCol采购价)) * Val(.TextMatrix(intRow, mconIntCol数量)), mFMT.FM_金额)
                
                lng材料ID = Val(.TextMatrix(intRow, 0))
                lng批次 = Val(.TextMatrix(intRow, mconIntCol批次))
                dbl数量 = Round(Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_散装小数.数量小数)
                dbl成本价 = Round(Val(.TextMatrix(intRow, mconIntCol采购价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_散装小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mconIntCol采购金额)), g_小数位数.obj_散装小数.金额小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mconIntCol售价金额)), g_小数位数.obj_散装小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mconintCol差价)), g_小数位数.obj_散装小数.金额小数)
                int序号 = Val(.TextMatrix(intRow, mconIntCol序号))
                         
                'zl_材料其他出库_VERIFY( /*NO_IN*/, /*库房ID_IN*/, /*药品ID_IN*/, /*批次_IN*/,
                    '/*实际数量_IN*/, /*成本价_IN*/, /*成本金额_IN*/, /*零售金额_IN*/,
                    '/*差价_IN*/, /*入出类别ID_IN*/, /*审核人_IN*/, /*审核日期_IN*/ );
                         
                gstrSQL = "zl_材料其他出库_Verify(" & _
                    int序号 & ",'" & _
                    strNo & "'," & _
                    lng库房ID & "," & _
                    lng材料ID & "," & _
                    lng批次 & "," & _
                    dbl数量 & "," & _
                    dbl成本价 & "," & _
                    dbl成本金额 & "," & _
                    dbl零售金额 & "," & _
                    dbl差价 & "," & _
                    lng入出类别ID & ",'" & _
                    str审核人 & "',to_date('" & _
                    dat审核日期 & "','yyyy-mm-dd HH24:MI:SS'),1)"
                    
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
'    If Not 检查单价(21, txtNO.Tag) Then
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
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol数量)), Val(.TextMatrix(intRow, mconIntCol冲销数量))) Then
                    ShowMsgBox "请输入合法的冲销数量（第" & intRow & "行）！"
                    Exit Function
                End If
            End If
        Next
        
        NO_IN = Trim(txtNO.Tag)
        填制人_IN = UserInfo.用户名
        填制日期_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        行次_IN = 0
        Dim bln全冲 As Boolean, dbl实际数量 As Double
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                材料ID_IN = .TextMatrix(intRow, 0)
                冲销数量_IN = Format(.TextMatrix(intRow, mconIntCol冲销数量) * .TextMatrix(intRow, mconIntCol比例系数), mFMT.FM_数量)
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                dbl实际数量 = Val(Format(Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), mFMT.FM_数量))
                bln全冲 = (冲销数量_IN = dbl实际数量)
                
                'ZL_材料其它出库_STRIKE(
                '    行次_In       In Integer,
                '    原记录状态_In In 药品收发记录.记录状态%Type,
                '    No_In         In 药品收发记录.NO%Type,
                '    序号_In       In 药品收发记录.序号%Type,
                '    材料id_In     In 药品收发记录.药品id%Type,
                '    冲销数量_In   In 药品收发记录.实际数量%Type,
                '    填制人_In     In 药品收发记录.填制人%Type,
                '    填制日期_In   In 药品收发记录.填制日期%Type,
                '    全部冲销_In   In 药品收发记录.实际数量%Type := 0 --1-全部冲销,0-部分冲销
                
                gstrSQL = "ZL_材料其他出库_STRIKE(" & _
                    行次_IN & "," & _
                    原记录状态_IN & ",'" & _
                    NO_IN & "'," & _
                    序号_IN & "," & _
                    材料ID_IN & "," & _
                    冲销数量_IN & ",'" & _
                    填制人_IN & "',to_date('" & _
                    Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & IIf(bln全冲, 1, 0) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            
            recSort.MoveNext
        Next
        
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            ShowMsgBox "没有选择一行卫材来冲销，不能冲销，请检查！"
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
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol材料) = 0 Then
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
    
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , , , , , mstrPrivs, , False)
    If RecReturn.RecordCount > 0 Then
    
        With mshBill
            mblnChange = True
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                    
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                    .Row = .Row + 1
                    
                End If
                .Col = mconIntCol数量
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int点击行
            
            If mstr重复卫材 <> "" Then
                MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                mstr重复卫材 = ""
            End If
                
'            If RecReturn.RecordCount = 1 Then
'                SetColValue .Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
'                .Col = mconIntCol数量
'            End If
        End With
        RecReturn.Close
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol数量 Or .Col = mconIntCol冲销数量 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol数量, mconIntCol冲销数量
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
            Case mconIntCol材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mconIntCol数量, mconIntCol冲销数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    
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
            
            Case mconIntCol材料
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , , , mstrPrivs, , False)
                    
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
                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
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
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
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
                    Call 提示库存数
                End If
            
            Case mconIntCol数量, mconIntCol冲销数量
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
                    If Val(strKey) = 0 And mint编辑状态 <> 3 Then
                        MsgBox "数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint编辑状态 = 6 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "冲销数量不能大于数量！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub
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
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = Format(.TextMatrix(.Row, mconIntCol售价) * strKey, mFMT.FM_金额)
                    End If
                    
                    If mint编辑状态 <> 6 Then
'                        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
'
'                        Call 验证出库差价计算(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(.TextMatrix(.Row, mconIntCol比例系数)), Val(.TextMatrix(.Row, mconIntCol实际差价)), Val(.TextMatrix(.Row, mconIntCol实际金额)), Val(Split(.TextMatrix(.Row, mconIntCol指导差价率), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mconIntCol售价金额)), dbl差价, dbl购价, dbl成本金额)
'                        .TextMatrix(.Row, mconintCol差价) = Format(dbl差价, mFMT.FM_金额)
                        .TextMatrix(.Row, mconIntCol采购价) = Format(Get成本价(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol批次))) * Val(.TextMatrix(.Row, mconIntCol比例系数)), mFMT.FM_成本价)
'                        .TextMatrix(.Row, mconIntCol采购金额) = Format(dbl成本金额, mFMT.FM_金额)
'                    Else
'                        .TextMatrix(.Row, mconIntCol采购金额) = Format(Val(.TextMatrix(.Row, mconIntCol采购价)) * strKey, mFMT.FM_金额)
'                        .TextMatrix(.Row, mconintCol差价) = Format(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol采购金额)), mFMT.FM_金额)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol采购金额) = Format(Val(.TextMatrix(.Row, mconIntCol采购价)) * strKey, mFMT.FM_金额)
                    .TextMatrix(.Row, mconintCol差价) = Format(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol采购金额)), mFMT.FM_金额)
                    
                    If .Col = mconIntCol数量 Then
                        .TextMatrix(.Row, mconIntCol冲销数量) = strKey
                    End If
                    
                    .TextMatrix(.Row, mconintCol外销金额) = Format(Val(.TextMatrix(.Row, mconintCol外销价)) * Val(strKey), mFMT.FM_金额)
                    
                    '税金=外销金额*增值税率
                    .TextMatrix(.Row, mconintCol税金) = Format(Val(.TextMatrix(.Row, mconintCol外销价)) * Val(strKey) * (Val(.TextMatrix(.Row, mconintCol增值税率)) / 100 / (1 + Val(.TextMatrix(.Row, mconintCol增值税率)) / 100)), mFMT.FM_金额)
                End If
                显示合计金额
            Case mconintCol外销价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "外销价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "对不起，外销价必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "外销价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, mFMT.FM_零售价)
                    .TextMatrix(.Row, .Col) = .Text
                    
                    '重算外调金额
                    .TextMatrix(.Row, mconintCol外销金额) = Format(Val(.TextMatrix(.Row, mconintCol外销价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mFMT.FM_金额)
                    
                    '重算税金
                    .TextMatrix(.Row, mconintCol税金) = Format(Val(.TextMatrix(.Row, mconintCol外销价)) * Val(.TextMatrix(.Row, mconIntCol数量)) * (Val(.TextMatrix(.Row, mconintCol增值税率)) / 100 / (1 + Val(.TextMatrix(.Row, mconintCol增值税率)) / 100)), mFMT.FM_金额)
                End If
        End Select
    End With
End Sub

'从材料特性中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
        ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
        ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
        ByVal str效期 As String, ByVal str灭菌失效期 As String, ByVal num可用数量 As Double, ByVal num实际金额 As Double, _
        ByVal num实际差价 As Double, ByVal num指导差价率 As Double, _
        ByVal num比例系数 As Double, ByVal lng批次 As Long, _
        ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    SetColValue = False
    If Format(str灭菌失效期, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str灭菌失效期) <> "" Then
       If MsgBox("材料【" & str材料 & "(" & lng批次 & ")】已经过了灭菌失效期,是否还要领用！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
            Exit Function
       End If
    End If
    
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mconIntCol批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
                
        If int是否变价 = 1 Then
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "           and 药品id=[2]" & _
                "           and 性质=1 and 实际数量>0 and " & _
                "           nvl(批次,0)=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
                        
            If rsTemp.EOF Then
                MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                Exit Function
            End If
            
            If lng批次 = 0 Then
                dblPrice = rsTemp!平均零售价
            Else
                dblPrice = rsTemp!分批售价
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mconIntCol行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mconIntCol材料) = str材料
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mconIntCol售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mconIntCol可用数量) = Format(num可用数量, mFMT.FM_数量)
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntCol指导差价率) = num指导差价率 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        If int是否变价 = 1 Then .TextMatrix(intRow, mconIntCol售价) = Format(dblPrice, mFMT.FM_零售价)
        Call CheckLapse(str效期)
        
        '外销价默认为采购价=结算价/扣率
        gstrSQL = "Select A.指导批发价, A.增值税率, Nvl(B.采购价,0) As 采购价 " & _
            " From 材料特性 A, " & _
            " (Select 药品id, 上次采购价 / Nvl(上次扣率, 100) * 100 As 采购价 " & _
            " From 药品库存 " & _
            " Where 性质 = 1 And 库房id + 0 = [1] And 药品id = [2] And Nvl(批次, 0) = [3]) B " & _
            " Where A.材料id = B.药品id(+) And A.材料id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取药品外销信息", Val(cboStock.ItemData(cboStock.ListIndex)), lng材料ID, lng批次)
        
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, mconintCol增值税率) = zlStr.FormatEx(rsTemp!增值税率, 2)
            
            If rsTemp!采购价 > 0 Then
                .TextMatrix(intRow, mconintCol外销价) = Format(rsTemp!采购价 * num比例系数, mFMT.FM_零售价)
            Else
                .TextMatrix(intRow, mconintCol外销价) = Format(rsTemp!指导批发价 * num比例系数, mFMT.FM_零售价)
            End If
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
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, ";") <> 0 Then
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
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                ShowMsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!"
                txt摘要.SetFocus
                Exit Function
            End If
            If InStr(1, txt摘要.Text, ";") <> 0 Then
                ShowMsgBox "在摘要中不能输入分号!"
                txt摘要.SetFocus
                Exit Function
            End If
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        ShowMsgBox "第" & intLop & "行卫材的数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol冲销数量))) = "" And mint编辑状态 = 6 Then
                        ShowMsgBox "第" & intLop & "行卫材的数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol冲销数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的填写数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol冲销数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的实际数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol冲销数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol采购金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol数量) = 4, mconIntCol数量, mconIntCol冲销数量)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol数量) = 4, mconIntCol数量, mconIntCol冲销数量)
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


Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim lng入出类别ID As Long
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lngTypeID As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次 As Long
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str灭菌效期 As String
    Dim arrSQL As Variant
    Dim intRow As Integer
    
    Dim dblOutPrice As Double   '外调价
    Dim strOutUnit As String    '外调单位
    Dim dbl增值税率 As Double
    Dim n As Long
    
    SaveCard = False
    arrSQL = Array()
    
    
    '在外面设置入出类别ID，主要是所有卫材都要用他
    
    
    With mshBill
        chrNo = Trim(txtNO)
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint编辑状态 = 1 Then   'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(74, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(74, lng库房ID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        lng入出类别ID = cboType.ItemData(cboType.ListIndex)
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str审核人 = Txt审核人
        
        If cboType.Text = "材料外销" Then
            strOutUnit = Mid(cbo外销单位.Text, 1, InStr(1, cbo外销单位.Text, "-") - 1)
        Else
            strOutUnit = ""
        End If
        
        On Error GoTo ErrHandle
        If mint编辑状态 = 2 Or bln强制保存 = True Then       '修改
            gstrSQL = "zl_材料其他出库_Delete('" & mstr单据号 & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0" & ";" & vbCrLf & gstrSQL
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                lng批次 = .TextMatrix(intRow, mconIntCol批次)
                str效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                dbl数量 = Round(Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl成本价 = Round(Val(.TextMatrix(intRow, mconIntCol采购价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mconIntCol采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零售价 = Round(Val(.TextMatrix(intRow, mconIntCol售价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.零售价小数)
                str灭菌效期 = IIf(.TextMatrix(intRow, mconIntCol灭菌失效期) = "", "", .TextMatrix(intRow, mconIntCol灭菌失效期))
                
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mconIntCol售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mconintCol差价)), g_小数位数.obj_最大小数.金额小数)
                lng序号 = intRow
                
                If cboType.Text = "材料外销" Then
                    dblOutPrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol外销价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.零售价小数)
                End If
                
                dbl增值税率 = Val(.TextMatrix(intRow, mconintCol增值税率))
                
                'zl_材料其他出库_INSERT( /*入出类别ID_IN*/, /*NO_IN*/, /*序号_IN*/,
                    '/*库房ID_IN*/, /*材料ID_IN*/, /*批次_IN*/, /*填写数量_IN*/,
                    '/*成本价_IN*/, /*成本金额_IN*/, /*零售价_IN*/, /*零售金额_IN*/,
                    '/*差价_IN*/, /*填制人_IN*/, /*填制日期_IN*/, /*产地_IN*/,
                    '/*批号_IN*/, /*效期_IN*/灭菌效期/, /*摘要_IN*/ );
                
                gstrSQL = "zl_材料其他出库_INSERT(" & _
                    lng入出类别ID & ",'" & _
                    chrNo & "'," & _
                    lng序号 & "," & _
                    lng库房ID & "," & lng材料ID & "," & lng批次 & "," & dbl数量 & "," & _
                    dbl成本价 & "," & dbl成本金额 & "," & dbl零售价 & "," & dbl零售金额 & "," & _
                    dbl差价 & ",'" & str填制人 & "',to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" & str产地 & "','" & _
                    str批号 & "'," & _
                    IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    str摘要 & "'," & _
                    dblOutPrice & ",'" & _
                    strOutUnit & "'," & _
                    dbl增值税率 & ",1)"
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL

            End If
            
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
        If Not 检查单价(21, txtNO.Tag) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    
    SaveCard = True
    Exit Function
ErrHandle:
    
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog

End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double, Cur外销金额 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            Cur外销金额 = Cur外销金额 + Val(.TextMatrix(intLop, mconintCol外销金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
    lblOther.Caption = "外销合计：" & Format(Cur外销金额, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mconIntCol材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        gstrSQL = "" & _
            "   Select 可用数量/" & .TextMatrix(.Row, mconIntCol比例系数) & " as  可用数量 " & _
            "   From 药品库存 " & _
            "   Where 库房id=[1]" & _
            "           and 药品id=[2]" & _
            "           and 性质=1 and " & _
            "           nvl(批次,0)=[3]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        stbThis.Panels(2).Text = "该卫材当前库存数为[" & Format(.TextMatrix(.Row, mconIntCol可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mconIntCol单位)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboType_LostFocus()
    If cboType.Text = "" Then
        cboType.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub cboType_Validate(Cancel As Boolean)
    If cboType.Text = "" Then
        cboType.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
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

Private Sub txt摘要_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt摘要, KeyAscii, m文本式
    If KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If

End Sub

Private Sub txt摘要_LostFocus()
    ImeLanguage False
End Sub

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim numUsedCount As Double, dbltotal As Double
    Dim vardrug As Variant, intLop As Integer
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False

    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        dblUsableQuantity = Format(.TextMatrix(intRow, mconIntCol可用数量), mFMT.FM_数量)
        
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

'打印单据
Private Sub printbill()
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1718", mint记录状态, mintUnit, 1718, "卫材其他出库", strNo
End Sub

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
                !序号 = IIf(Val(mshBill.TextMatrix(n, mconIntCol序号)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub

