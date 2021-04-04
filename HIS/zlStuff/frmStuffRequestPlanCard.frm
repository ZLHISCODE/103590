VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffRequestPlanCard 
   Caption         =   "卫材申购单编辑"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmStuffRequestPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
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
      Left            =   8820
      TabIndex        =   9
      Top             =   5085
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10140
      TabIndex        =   10
      Top             =   5085
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   -15
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   45
      Width           =   11715
      Begin VB.ComboBox cbo类型 
         Height          =   300
         Left            =   9510
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1710
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   570
         Width           =   2055
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
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
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         ItemData        =   "frmStuffRequestPlanCard.frx":014A
         Left            =   4890
         List            =   "frmStuffRequestPlanCard.frx":014C
         TabIndex        =   3
         Text            =   "cboEnterStock"
         Top             =   570
         Width           =   2115
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "被申购库房(&I)"
         Height          =   180
         Left            =   3510
         TabIndex        =   2
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申购部门(&S)"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   630
         Width           =   990
      End
      Begin VB.Label txt计划类型 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   27
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额合计："
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   1005
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
         Caption         =   "卫材申购单"
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
      Begin VB.Label Lbl计划类型 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划类型:"
         Height          =   180
         Left            =   8550
         TabIndex        =   4
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制人"
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
         Caption         =   "编制日期"
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
            Picture         =   "frmStuffRequestPlanCard.frx":014E
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0368
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0582
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":079C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":09B6
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0BD0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0DEA
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1004
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
            Picture         =   "frmStuffRequestPlanCard.frx":121E
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1438
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1652
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":186C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1A86
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1CA0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1EBA
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":20D4
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
            Picture         =   "frmStuffRequestPlanCard.frx":22EE
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
            Picture         =   "frmStuffRequestPlanCard.frx":2B82
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffRequestPlanCard.frx":3084
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf供应商选择 
      Height          =   2565
      Left            =   5850
      TabIndex        =   29
      Top             =   1890
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
Attribute VB_Name = "frmStuffRequestPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private mintErrMsg As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintUnit As Integer            '0-散装单位,1-包装单位
Private mbln下限 As Boolean                 '仅提取低于储备下限的药品
Private mint上限 As Integer
Private mint下限 As Integer

Private mlng计划ID As Long
Private mlng库房id As Long
Private mint计划类型 As Integer
Private mint编制方法 As Integer
Private mstr供货商ID As String      '以id分隔
Private mbln中标单位 As Boolean '包含中标供货商,要与mstr供货单位一起启作用.
Private mstr期间  As String                  '月以六位表示,季以五位表示,年以四位表示
Dim mstrPrivs As String                     '权限
Private Const mlngModule = 1725
Private mstrLike As String
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mblnProvider As Boolean                 '查看上次供应商相关信息 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材申购单编辑"
Private mstr重复卫材 As String '记录重复的卫材

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Enum mHeadCol
    序号 = 1
    材料 = 2
    规格 = 3
    产地 = 4
    单位 = 5
    比例系数 = 6
    中标材料 = 7
    请购数量 = 8
    计划数量 = 9
    单价 = 10
    金额 = 11
    上次供应商 = 12
End Enum

Private Const mconIntColS  As Integer = 13     '总列数

'=========================================================================================

Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, _
        ByVal int编辑状态 As Integer, ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
    '----------------------------------------------------------------------------------------------------------------
    '功能:申购计划编辑入口
    '参数:frmMain-调用的父窗口
    '     str单据号-单据号
    '     int编辑状态-1.新增；2、修改；3、验收；4、查看；5
    '     strPrivs-权限串
    '     blnSuccess-编辑成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------------
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mintErrMsg = 1
    mstrPrivs = strPrivs

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = frmMain
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "查看供应商")
    
    If Not GetDepend(mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)) Then Exit Sub
    
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
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
End Sub

Private Sub cboEnterStock_Change()
    mblnChange = True
End Sub

Private Sub cboEnterStock_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, str站点限制 As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    If cboEnterStock.ItemData(cboEnterStock.ListIndex) = -1 And Visible Then
        strSQL = "" & _
            "   SELECT DISTINCT a.id,a.简码,a.编码||'-'||a.名称  as 名称" & _
            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "   Where c.工作性质 = b.名称 " & _
            IIf(str站点限制 <> "", " And (a.站点 = [1] or a.站点 is null) ", "") & _
            "     And b.编码 In('V','K') " & _
            "     AND a.id = c.部门id " & _
            "     AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
        vRect = zlControl.GetControlRect(cboEnterStock.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "被申购库房", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboEnterStock.Height, blnCancel, False, True, str站点限制)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboEnterStock, rsTmp!Id)
            If intIdx <> -1 Then
                cboEnterStock.ListIndex = intIdx
'            Else
'                cboEnterStock.AddItem rsTmp!名称, cboEnterStock.ListCount - 1
'                cboEnterStock.ItemData(cboEnterStock.NewIndex) = rsTmp!Id
'                cboEnterStock.ListIndex = cboEnterStock.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有被申购库房数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If

            intIdx = cbo.FindIndex(cboEnterStock, cboEnterStock.Tag)
            Call cbo.SetIndex(cboEnterStock.hwnd, intIdx)
        End If
    Else
        cboEnterStock.Tag = cboEnterStock.Text
    End If
End Sub

Private Sub cboEnterStock_GotFocus()
    Call zlControl.TxtSelAll(cboEnterStock)
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboEnterStock.Style = 2 And cboEnterStock.ListIndex <> -1 Then
            cboEnterStock.ListIndex = -1
        End If
    End If
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cboEnterStock.Locked And cboEnterStock.Style = 2 Then
            lngIdx = cbo.MatchIndex(cboEnterStock.hwnd, KeyAscii)
            If lngIdx = -1 And cboEnterStock.ListCount > 0 Then lngIdx = 0
            cboEnterStock.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String, str站点限制 As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboEnterStock.ListIndex <> -1 Then Exit Sub '已选中
    If cboEnterStock.Text = "" Then cboEnterStock.Tag = "": Exit Sub '无输入
    
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    strInput = UCase(NeedName(cboEnterStock.Text))
    strSQL = " SELECT DISTINCT a.id,a.简码,a.编码||'-'||a.名称  as 名称" & _
            "  FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "  Where c.工作性质 = b.名称 " & IIf(str站点限制 <> "", " And (a.站点 = [3] or a.站点 is null) ", "") & _
            "    And b.编码 In('V','K') " & _
            "    AND a.id = c.部门id " & _
            "    AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
            "    And (Upper(a.编码) Like [1] Or Upper(a.名称) Like [2] Or Upper(a.简码) Like [2]) " & _
            "  Order by 名称 "
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cboEnterStock.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "被申购库房", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboEnterStock.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str站点限制)
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cboEnterStock, rsTmp!Id)
        If intIdx <> -1 Then
            cboEnterStock.ListIndex = intIdx
        Else
            cboEnterStock.AddItem rsTmp!名称, cboEnterStock.ListCount - 1
            cboEnterStock.ItemData(cboEnterStock.NewIndex) = rsTmp!Id
            cboEnterStock.ListIndex = cboEnterStock.NewIndex
        End If
        mlng库房id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的被申购库房。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    strInput = UCase(NeedName(cboStock.Text))
    
    If cboStock.ItemData(cboStock.ListIndex) = -1 And Visible Then
        strSQL = "" & _
            "   SELECT DISTINCT a.id,a.简码,a.编码||'-'||a.名称 as 名称" & _
            "   FROM 部门表 a  " & _
            "   where (a.撤档时间 is null or TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01') And (a.站点=[2] or a.站点 is null) " & _
            IIf(InStr(1, mstrPrivs, ";所有部门;") > 0, "", " and  id in (Select 部门id from 部门人员 where 人员id =[1])") & _
            "   Order by 简码"
        vRect = zlControl.GetControlRect(cboStock.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "申购部门", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboStock.Height, blnCancel, False, True, strInput, gstrNodeNo)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboStock, rsTmp!Id)
            If intIdx <> -1 Then
                cboStock.ListIndex = intIdx
            Else
                cboStock.AddItem rsTmp!名称, cboStock.ListCount - 1
                cboStock.ItemData(cboStock.NewIndex) = rsTmp!Id
                cboStock.ListIndex = cboStock.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有申购部门数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If

            intIdx = cbo.FindIndex(cboStock, cboStock.Tag)
            Call cbo.SetIndex(cboStock.hwnd, intIdx)
        End If
    Else
        cboStock.Tag = cboStock.Text
        '刷新cboEnterStock
        SetEnterStock cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_GotFocus()
    Call zlControl.TxtSelAll(cboStock)
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboStock.Style = 2 And cboStock.ListIndex <> -1 Then
            cboStock.ListIndex = -1
        End If
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cboStock.Locked And cboStock.Style = 2 Then
            lngIdx = cbo.MatchIndex(cboStock.hwnd, KeyAscii)
            If lngIdx = -1 And cboStock.ListCount > 0 Then lngIdx = 0
            cboStock.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboStock.ListIndex <> -1 Then Exit Sub '已选中
    If cboStock.Text = "" Then cboStock.Tag = "": Exit Sub '无输入
    
    strInput = UCase(NeedName(cboStock.Text))
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id,a.简码,a.编码||'-'||a.名称 as 名称" & _
        "   FROM 部门表 a  " & _
        "   where (a.撤档时间 is null or TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01') And (a.站点=[3] or a.站点 is null) " & _
        IIf(InStr(1, mstrPrivs, ";所有部门;") > 0, "", " and  id in (Select 部门id from 部门人员 where 人员id =[1])") & _
        " And (Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2]) " & _
        " Order by 简码"
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cboStock.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "申购部门", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboStock.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", gstrNodeNo)
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cboStock, rsTmp!Id)
        If intIdx <> -1 Then
            cboStock.ListIndex = intIdx
        Else
            cboStock.AddItem rsTmp!名称, cboStock.ListCount - 1
            cboStock.ItemData(cboStock.NewIndex) = rsTmp!Id
            cboStock.ListIndex = cboStock.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的操作员。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo类型_Change()
    mblnChange = True
End Sub

Private Sub cbo类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
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
        FindData mshBill, mHeadCol.材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindData mshBill, mHeadCol.材料, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf供应商选择.Visible Then
            Msf供应商选择.ZOrder 1
            Msf供应商选择.Visible = False
            Exit Sub
        End If
        Call cmdCancel_Click
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
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1725", 0, mintUnit, 1725, "卫材申购单", txtNO.Tag)
        '退出
        Unload Me
        Exit Sub
    End If

    If mint编辑状态 = 3 Then        '审核
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "单据编号=" & txtNO.Tag, "单位=" & mintUnit, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "单据编号=" & txtNO.Tag, "单位=" & mintUnit, 2
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
    txt摘要.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
End Sub

Private Sub Form_Activate()
    Dim intMonth As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mint编辑状态 = 1 Then
        If cboEnterStock.Enabled Then cboEnterStock.SetFocus
        If cboStock.Enabled Then cboStock.SetFocus
    Else
'        mblnChange = False
        Select Case mintErrMsg
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
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub
 

Private Sub Form_Load()
    mFMT.FM_金额 = GetDigit

    Me.cboStock.Enabled = True
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    initCard
    RestoreWinState Me, App.ProductName, mstrCaption
    With mshBill
        .ColWidth(mHeadCol.单价) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mHeadCol.金额) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mHeadCol.上次供应商) = IIf(mblnProvider = True, 1000, 0)
    End With
End Sub
Private Sub init类型()
    With cbo类型 '
        .Clear
        .AddItem "月度计划"
        .AddItem "季度计划"
        .AddItem "年度计划"
        .ListIndex = 0
    End With
End Sub
Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strStock As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str单位 As String
    Dim strOrder As String, strCompare As String
    Dim blnNO库房 As Boolean
    
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    
    Call init类型
    
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            If InStr(1, gstrPrivs, ";所有部门;") > 0 Then
                For i = 1 To .ListCount - 1
                    If .List(i) <> "所有部门" Then
                        cboStock.AddItem .List(i)
                        cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                    End If
                Next
                mintcboIndex = .ListIndex - 1
                cboStock.ListIndex = .ListIndex - 1
            Else
                For i = 0 To .ListCount - 1
                    cboStock.AddItem .List(i)
                    cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                Next
                mintcboIndex = .ListIndex
                cboStock.ListIndex = .ListIndex
            End If
       End With
    End If
   
'   strStock = " And b.编码 In('V','K','W','12') "
'
'   gstrSQL = "SELECT DISTINCT a.id,a.编码||'-'||a.名称  as 名称 " & _
'             "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
'             "Where c.工作性质 = b.名称 " & _
'             IIf(str站点限制 <> "", " And a.站点 = [2] ", "") & _
'             strStock & _
'             "  AND a.id = c.部门id " & _
'             "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' "
'
'    Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, str站点限制)
'
'    With cboEnterStock
'        .Clear
'        Do While Not rsInitCard.EOF
'            .AddItem NVL(rsInitCard!名称)
'            .ItemData(.NewIndex) = Val(NVL(rsInitCard!Id))
'            rsInitCard.MoveNext
'        Loop
'    End With
    '初始化cboEnterStock控件
    SetEnterStock mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)

    '库房
    Select Case mint编辑状态
        Case 1
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            strUnit = "包装单位"
            Select Case mintUnit
            Case 0
                str单位 = ",j.计算单位 单位,1 比例系数"
            Case Else
                str单位 = ",m.包装单位 单位,m.换算系数 比例系数"
            End Select
            
            initGrid
            
            gstrSQL = "" & _
                "   Select a.库房id,a.部门ID,b.编码||'-'||b.名称 as 库房,c.编码||'-'||c.名称 as 部门" & _
                "   From  材料采购计划 a,部门表 b,部门表 C " & _
                "   where a.库房id=b.id(+) and a.部门id=c.id(+) and  a.单据=1 and  a.NO=[1] and rownum=1 "
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
            If rsInitCard.EOF Then
                mintErrMsg = 2
                Exit Sub
            End If
            mlng库房id = Val(zlStr.Nvl(rsInitCard!部门ID))
            With cboStock
                blnNO库房 = True
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = mlng库房id Then
                        blnNO库房 = False
                        .ListIndex = i: Exit For
                    End If
                Next
                If blnNO库房 Then
                    If mlng库房id <> 0 Then
                        .AddItem zlStr.Nvl(rsInitCard!部门)
                        .ListIndex = .NewIndex
                    Else
                        .ListIndex = 0
                    End If
                End If
            End With
            
            
            mlng库房id = Val(zlStr.Nvl(rsInitCard!库房ID))
            With cboEnterStock
                blnNO库房 = True
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = mlng库房id Then
                        blnNO库房 = False
                        .ListIndex = i: Exit For
                    End If
                Next
                If blnNO库房 Then
                    If mlng库房id <> 0 Then
                        .AddItem zlStr.Nvl(rsInitCard!库房)
                        .ListIndex = .NewIndex
                    Else
                        .ListIndex = 0
                    End If
                End If
            End With
            
            gstrSQL = "" & _
                "   SELECT a.id,nvl(a.库房id,0) as 库房id,nvl(c.名称,'全院') AS 库房,a.no, a.计划类型,a.期间, a.编制方法, a.编制人," & _
                "           TO_CHAR (a.编制日期, 'yyyy-mm-dd HH24:MI:SS') AS 编制日期, a.审核人," & _
                "           TO_CHAR (a.审核日期, 'yyyy-mm-dd HH24:MI:SS') AS 审核日期,a.编制说明," & _
                "           b.序号,b.材料id 药品id,m.招标材料,J.编码,J.名称 通用名称, J.规格" & str单位 & _
                "          ,b.请购数量,b.计划数量, b.单价, b.金额, b.上次供应商,b.上次生产商 " & _
                "   FROM 材料采购计划 a, 材料计划内容 b,部门表 c,材料特性 M,收费项目目录 J " & _
                "   Where a.id = b.计划id and nvl(a.库房id,0)=c.id(+) " & _
                "          and b.材料id=m.材料id and m.材料id=J.id and nvl(a.单据,0)=1 AND a.no = [1]" & _
                "   Order by " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "编码", "通用名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
            
            If rsInitCard.EOF Then
                mintErrMsg = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt填制人 = rsInitCard!编制人
            If mint编辑状态 = 2 Then
                Txt填制人 = gstrUserName
            End If
            Txt填制日期 = Format(rsInitCard!编制日期, "yyyy-mm-dd hh:mm:ss")

            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsInitCard!编制说明), "", rsInitCard!编制说明)
            mint计划类型 = rsInitCard!计划类型
            mint编制方法 = rsInitCard!编制方法
            mlng库房id = rsInitCard!库房ID
            mlng计划ID = rsInitCard!Id
            
            mstr期间 = zlStr.Nvl(rsInitCard!期间)
            If mint计划类型 >= 1 And mint计划类型 <= 3 Then
                cbo类型.ListIndex = mint计划类型 - 1
            End If
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintErrMsg = 3
                Exit Sub
            End If

            With mshBill
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!药品id
                    .TextMatrix(intRow, mHeadCol.材料) = "[" & rsInitCard!编码 & "]" & rsInitCard!通用名称
                    .TextMatrix(intRow, mHeadCol.规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mHeadCol.上次供应商) = IIf(IsNull(rsInitCard!上次供应商), "", rsInitCard!上次供应商)
                    .TextMatrix(intRow, mHeadCol.产地) = IIf(IsNull(rsInitCard!上次生产商), "", rsInitCard!上次生产商)
                    .TextMatrix(intRow, mHeadCol.单位) = rsInitCard!单位
                    .TextMatrix(intRow, mHeadCol.比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mHeadCol.中标材料) = IIf(Val(zlStr.Nvl(rsInitCard!招标材料)) = 1, "√", "")
                    .TextMatrix(intRow, mHeadCol.请购数量) = IIf(Format(Val(zlStr.Nvl(rsInitCard!请购数量)), mFMT.FM_数量) = 0, "", Format(rsInitCard!请购数量 / rsInitCard!比例系数, mFMT.FM_数量))
                    If mint编辑状态 = 3 Then
                        .TextMatrix(intRow, mHeadCol.计划数量) = .TextMatrix(intRow, mHeadCol.请购数量)
                    Else
                        .TextMatrix(intRow, mHeadCol.计划数量) = IIf(Format(Val(zlStr.Nvl(rsInitCard!计划数量)), mFMT.FM_数量) = 0, "", Format(rsInitCard!计划数量 / rsInitCard!比例系数, mFMT.FM_数量))
                    End If
                    .TextMatrix(intRow, mHeadCol.单价) = Format(Val(zlStr.Nvl(rsInitCard!单价)) * rsInitCard!比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mHeadCol.金额) = IIf(Format(Val(zlStr.Nvl(rsInitCard!金额)), mFMT.FM_金额) = 0, "", Format(Val(zlStr.Nvl(rsInitCard!金额)), mFMT.FM_金额))
                    If intRow = .Rows - 1 Then .Rows = .Rows + 1
                    rsInitCard.MoveNext
                Next
            End With
            rsInitCard.Close
    End Select
    Call SetEdit
    Call RefreshRowNO(mshBill, mHeadCol.序号, 1)
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'初始化编辑控件
Private Sub initGrid()
    Dim intCol As Integer

    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mHeadCol.序号) = "序号"
        .TextMatrix(0, mHeadCol.材料) = "材料名称与编码"
        .TextMatrix(0, mHeadCol.规格) = "规格"
        .TextMatrix(0, mHeadCol.产地) = "产地"
        .TextMatrix(0, mHeadCol.单位) = "单位"
        .TextMatrix(0, mHeadCol.比例系数) = "比例系数"
        .TextMatrix(0, mHeadCol.中标材料) = "中标材料"
        .TextMatrix(0, mHeadCol.请购数量) = "请购数量"
        .TextMatrix(0, mHeadCol.计划数量) = "审批数量"
        .TextMatrix(0, mHeadCol.单价) = "成本价"
        .TextMatrix(0, mHeadCol.金额) = "成本金额"
        .TextMatrix(0, mHeadCol.上次供应商) = "上次供应商"
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mHeadCol.序号) = "1"

        .ColWidth(mHeadCol.序号) = 500
        .ColWidth(mHeadCol.材料) = 2000
        .ColWidth(mHeadCol.规格) = 900
        .ColWidth(mHeadCol.产地) = 800
        .ColWidth(mHeadCol.单位) = 500
        .ColWidth(mHeadCol.中标材料) = 800
        .ColWidth(mHeadCol.请购数量) = 1000
        .ColWidth(mHeadCol.计划数量) = 1000
        .ColWidth(mHeadCol.单价) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mHeadCol.金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mHeadCol.上次供应商) = IIf(mblnProvider = False, 0, 1000)
        .ColWidth(mHeadCol.比例系数) = 0
        .ColWidth(0) = 0

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next

        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mHeadCol.材料) = 1
            .ColData(mHeadCol.请购数量) = 4
            .ColData(mHeadCol.单价) = 4
            .ColData(mHeadCol.产地) = 4
            .ColData(mHeadCol.上次供应商) = 1
        ElseIf mint编辑状态 = 3 Then
            txt摘要.Enabled = False
            .ColData(mHeadCol.计划数量) = 4
        ElseIf mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            .ColData(mHeadCol.计划数量) = 0
        End If

        .ColAlignment(mHeadCol.材料) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.规格) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.产地) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.单位) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.请购数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.计划数量) = flexAlignRightCenter
        .ColAlignment(mHeadCol.单价) = flexAlignRightCenter
        .ColAlignment(mHeadCol.金额) = flexAlignRightCenter
        .ColAlignment(mHeadCol.上次供应商) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.中标材料) = 4
        If mint编辑状态 = 3 Then
            .PrimaryCol = mHeadCol.材料
            .LocateCol = mHeadCol.计划数量
        Else
            .PrimaryCol = mHeadCol.材料
            .LocateCol = mHeadCol.材料
        End If
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mHeadCol.材料) = 0
    End With

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
    
    cbo类型.Left = mshBill.Left + mshBill.Width - cbo类型.Width
    Lbl计划类型.Left = cbo类型.Left - Lbl计划类型.Width - 50


    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 50

    LblEnterStock.Left = cboStock.Left + cboStock.Width + cboStock.Width * 0.3
    cboEnterStock.Left = LblEnterStock.Left + LblEnterStock.Width + 50

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
    Dim str审核人 As String, intRow As Integer, lng序号 As Long, 材料ID_IN As Long
    Dim dbl单价_IN As Double, dbl金额_IN As Double, dbl请购数量_IN As Double, dbl计划数量_IN As Double
    Dim 上次供应商_IN As String, 上次生产商_IN As String
    Dim cllProc As New Collection
    mblnSave = False
    SaveCheck = False

    str审核人 = gstrUserName
    'Zl_材料计划管理_Delete
    gstrSQL = "Zl_材料计划管理_Delete("
    '  Id_In       In 材料采购计划.ID%Type,
    gstrSQL = gstrSQL & "" & mlng计划ID & ","
    '  删除明细_In Integer:=0
    '  --1-只删除明细,否则全删除
    gstrSQL = gstrSQL & "1)"
    cllProc.Add gstrSQL
    '插入明细
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng序号 = .TextMatrix(intRow, mHeadCol.序号)
                材料ID_IN = .TextMatrix(intRow, 0)
                dbl单价_IN = Round(Val(.TextMatrix(intRow, mHeadCol.单价)) / Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.成本价小数)
                dbl金额_IN = Round(Val(.TextMatrix(intRow, mHeadCol.金额)), g_小数位数.obj_散装小数.金额小数)
                dbl请购数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.请购数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                dbl计划数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.计划数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_散装小数.数量小数)
                上次供应商_IN = .TextMatrix(intRow, mHeadCol.上次供应商)
                上次生产商_IN = .TextMatrix(intRow, mHeadCol.产地)
                'Zl_材料计划管理次表_Insert
                gstrSQL = "Zl_材料计划管理次表_Insert("
                '  计划id_In     In 材料计划内容.计划id%Type,
                gstrSQL = gstrSQL & "" & mlng计划ID & ","
                '  材料id_In     In 材料计划内容.材料id%Type,
                gstrSQL = gstrSQL & "" & 材料ID_IN & ","
                '  序号_In       In 材料计划内容.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  请购数量_In   In 材料计划内容.请购数量%Type,
                gstrSQL = gstrSQL & "" & dbl请购数量_IN & ","
                '  计划数量_IN   In 材料计划内容.计划数量%Type,
                gstrSQL = gstrSQL & "" & dbl计划数量_IN & ","
                '  单价_IN       In 材料计划内容.单价%Type,
                gstrSQL = gstrSQL & "" & dbl单价_IN & ","
                '  金额_IN       In 材料计划内容.金额%Type,
                gstrSQL = gstrSQL & "" & dbl金额_IN & ","
                '  前期数量_In   In 材料计划内容.前期数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  上期数量_In   In 材料计划内容.上期数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  库存数量_In   In 材料计划内容.库存数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  上次供应商_In In 材料计划内容.上次供应商%Type := Null,
                gstrSQL = gstrSQL & "'" & 上次供应商_IN & "',"
                '  上次生产商_In In 材料计划内容.上次生产商%Type := Null
                gstrSQL = gstrSQL & "'" & 上次生产商_IN & "')"
                cllProc.Add gstrSQL
            End If
        Next
    End With
    'zl_材料计划管理_VERIFY( /*ID_IN*/, /*审核人_IN*/ );
    gstrSQL = "zl_材料计划管理_VERIFY('" & mlng计划ID & "','" & str审核人 & "')"
    cllProc.Add gstrSQL
    
    err = 0: On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllProc, mstrCaption
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

Private Sub Msf供应商选择_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
        .TextMatrix(.Row, mHeadCol.上次供应商) = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
    End With
    Msf供应商选择.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub

Private Sub Msf供应商选择_GotFocus()
    If Msf供应商选择.Rows - 1 = 1 Then Call Msf供应商选择_DblClick
End Sub

Private Sub Msf供应商选择_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf供应商选择_DblClick
    End If
End Sub

Private Sub Msf供应商选择_LostFocus()
    Msf供应商选择.ZOrder 1
    Msf供应商选择.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mHeadCol.序号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mHeadCol.序号, mshBill.Row)
    Call 显示合计金额
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mHeadCol.材料) = 0 Then
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
            If MsgBox("你确实要删除该行卫生材料吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    On Error GoTo ErrHandle
    If mshBill.Col = mHeadCol.材料 Then
        
        If cboEnterStock.ListIndex = -1 Then
            MsgBox "请选择被申购的库房！", vbInformation + vbOKOnly, gstrSysName
            cboEnterStock.SetFocus
            Exit Sub
        End If
        
        Set RecReturn = Frm材料选择器.ShowMe(Me, 1, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), , , , , , , , , , , , mlngModule, , mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            RecReturn.MoveFirst
            
            If mintUnit = 0 Then
                strUnit = "散装单位"
            Else
                strUnit = "包装单位"
            End If
            
            For i = 1 To RecReturn.RecordCount
                If SetStuffRows(RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                            Switch(strUnit = "散装单位", zlStr.Nvl(RecReturn!散装单位), strUnit = "包装单位", zlStr.Nvl(RecReturn!包装单位)), Val(zlStr.Nvl(RecReturn!指导批发价)), _
                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", Val(zlStr.Nvl(RecReturn!换算系数)))) Then
                    
                    If mshBill.Row = mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1 '只有当前行是最后一行时才新增行
                    mshBill.Row = mshBill.Row + 1
                End If
            
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int点击行
            
            If mstr重复卫材 <> "" Then
                MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                mstr重复卫材 = ""
            End If
            
'            If RecReturn.RecordCount = 1 Then
'                If mintUnit = 0 Then
'                    strUnit = "散装单位"
'                Else
'                    strUnit = "包装单位"
'                End If
'                SetStuffRows RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                            Switch(strUnit = "散装单位", zlStr.Nvl(RecReturn!散装单位), strUnit = "包装单位", zlStr.Nvl(RecReturn!包装单位)), Val(zlStr.Nvl(RecReturn!指导批发价)), _
'                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", Val(zlStr.Nvl(RecReturn!换算系数)))
'            End If
            RecReturn.Close
        End If
    Else
        '药品供应商的选择
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100
        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select ID,编码,名称,简码 From 供应商 " & _
                  "Where 末级=1 And (substr(类型,5,1)=1 And (站点=[1] or 站点 is null) Or Nvl(末级,0)=0) " & _
                  "    And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) Order By 编码 "
        Set RecReturn = zlDatabase.OpenSQLRecord(gstrSQL, "读取卫生材料供应商", gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            MsgBox "请先初始化卫生材料供应商！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With Msf供应商选择
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf供应商选择
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If

        Select Case .Col
            Case mHeadCol.材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
            Case mHeadCol.产地
                .TxtCheck = False
                .MaxLength = 40
            Case mHeadCol.上次供应商
                .MaxLength = 40
                .TxtCheck = False
            Case mHeadCol.计划数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.请购数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.单价
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"

        End Select

    End With
End Sub
Private Function Get产地(Optional strKey As String = "") As Boolean
    '功能:获取人员信息
    Dim rsTemp  As ADODB.Recordset
    Dim blnCancel  As Boolean
    Dim strSearch As String
    Dim vRect As RECT
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
     
     If strKey <> "" Then
        strSearch = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select 编码 as id ,a.编码 ,a.名称 ,a.简码,a.生产企业许可证 " & _
            "   From 材料生产商 a " & _
            "   Where (编码 like [1] or 简码 like [1] or 名称 like [1])  " & _
            "   order by 编码"
         vRect = zlControl.GetControlRect(mshBill.TxtHwnd)
         Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "产地选择", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, mshBill.RowHeight(mshBill.Row) - 50, blnCancel, False, False, strSearch)
     Else
        gstrSQL = "" & _
            "   Select 编码 as id,a.编码 ,a.名称 ,a.简码,a.生产企业许可证" & _
            "   From 材料生产商 a " & _
            "   order by 编码"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "产地选择器", True, "", "请选择相关的产地", True, False, , , , , blnCancel)
    End If
    If rsTemp Is Nothing Then Exit Function
    If blnCancel = True Then Exit Function
    
    With mshBill '
        .TextMatrix(.Row, mHeadCol.产地) = zlStr.Nvl(rsTemp!名称)
        .Text = .TextMatrix(.Row, mHeadCol.产地)
    End With
    Get产地 = True
End Function

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If .Col = mHeadCol.材料 Then
            .Text = Trim(.Text)
        Else
            .Text = Trim(.Text)
        End If
        strKey = .Text

        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col

            Case mHeadCol.材料
                If strKey <> "" Then
                    If cboEnterStock.ListIndex = -1 Then
                        MsgBox "请选择被申购的库房！", vbInformation + vbOKOnly, gstrSysName
                        cboEnterStock.SetFocus
                        Exit Sub
                    End If
        
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

                    Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , mlngModule, , mstrPrivs, , False)
                    
                    If rsTemp.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If mintUnit = 0 Then
                        strUnit = "散装单位"
                    Else
                        strUnit = "包装单位"
                    End If
                    
                    rsTemp.MoveFirst
                    For i = 1 To rsTemp.RecordCount
                        If SetStuffRows(rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
                            IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                            Switch(strUnit = "散装单位", zlStr.Nvl(rsTemp!散装单位), strUnit = "包装单位", zlStr.Nvl(rsTemp!包装单位)), Val(zlStr.Nvl(rsTemp!指导批发价)), _
                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", Val(zlStr.Nvl(rsTemp!换算系数)))) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        rsTemp.MoveNext
                    Next
                    
                    mshBill.Row = int点击行
                    
                    If mstr重复卫材 <> "" Then
                        MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                        mstr重复卫材 = ""
                    End If
                    
'                    If rsTemp.RecordCount = 1 Then
'                        If mintUnit = 0 Then
'                            strUnit = "散装单位"
'                        Else
'                            strUnit = "包装单位"
'                        End If
'                        If SetStuffRows(rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
'                            IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
'                            Switch(strUnit = "散装单位", zlStr.NVL(rsTemp!散装单位), strUnit = "包装单位", zlStr.NVL(rsTemp!包装单位)), Val(zlStr.NVL(rsTemp!指导批发价)), _
'                            Switch(strUnit = "散装单位", 1, strUnit = "包装单位", Val(zlStr.NVL(rsTemp!换算系数)))) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'
'                        Cancel = True
'                    End If
                End If
            Case mHeadCol.计划数量
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "审批数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "审批数量必须在(0~99999999)内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    Cancel = True
                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.单价) <> "" Then
                        .TextMatrix(.Row, mHeadCol.金额) = Format(.TextMatrix(.Row, mHeadCol.单价) * strKey, mFMT.FM_金额)
                    End If
                End If
                Call 显示合计金额
            Case mHeadCol.请购数量
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "请购数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "请购数量必须在(0~99999999)内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    Cancel = True
                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.单价) <> "" Then
                        .TextMatrix(.Row, mHeadCol.金额) = Format(.TextMatrix(.Row, mHeadCol.单价) * strKey, mFMT.FM_金额)
                    End If
                End If
                Call 显示合计金额
            Case mHeadCol.单价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "成本价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "成本价必须在(0~99999999)内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.单价) = " "
                        .Text = " "
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.单价) = strKey
                End If
                .TextMatrix(.Row, mHeadCol.金额) = Format(Val(.TextMatrix(.Row, mHeadCol.单价)) * Val(.TextMatrix(.Row, mHeadCol.请购数量)), mFMT.FM_金额)
                Call 显示合计金额
                
            Case mHeadCol.产地
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.产地) = ""
                    End If
                    .Col = mHeadCol.请购数量
                    Cancel = True
                    Exit Sub
                Else
                    If strKey <> "" Then
                        If Get产地(strKey) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                OS.OpenIme False
            Case mHeadCol.上次供应商
'                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.上次供应商) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.上次供应商) = strKey
                Else
                    If .TxtVisible = False Then Exit Sub
                    If StrIsValid(strKey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = UCase(strKey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100
            
                    Set rsTemp = New ADODB.Recordset
                    gstrSQL = "" & _
                        "   Select ID,编码,名称,简码 " & _
                        "   From 供应商 " & _
                        "   Where 末级=1 And (substr(类型,5,1)=1 And (站点=[2] or 站点 is null) Or Nvl(末级,0)=0) " & _
                        "       And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                        "       And (upper(编码) Like [1] Or Upper(名称) Like [1] Or Upper(简码) Like [1])" & _
                        "   Order By 编码 "
                    
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取卫生材料供应商", strKey & "%", gstrNodeNo)
                    
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "没有找到符合条件的供应商！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!名称
                        Exit Sub
                    End If
                    
                    With Msf供应商选择
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf供应商选择
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
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

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If cboEnterStock.ListIndex = -1 Then
                ShowMsgBox "被申购库房不能为空！"
                cboEnterStock.SetFocus
                Exit Function
            End If

            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > 40 Then
                MsgBox "摘要超长,最多能输入20个汉字或40个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If

            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mHeadCol.材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mHeadCol.计划数量))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mHeadCol.计划数量)) Then
                            MsgBox "第" & intLop & "行卫生材料的审批数量不为数字型，请检查！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.计划数量
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mHeadCol.计划数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的审批数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.计划数量
                        Exit Function
                    End If
                    If mint编辑状态 <> 3 Then
                        If Trim(Trim(.TextMatrix(intLop, mHeadCol.请购数量))) <> "" Then
                            If Not IsNumeric(.TextMatrix(intLop, mHeadCol.请购数量)) Then
                                MsgBox "第" & intLop & "行卫生材料的请购数量不为数字型，请检查！", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mHeadCol.请购数量
                                Exit Function
                            End If
                        End If
                        
                        If Val(.TextMatrix(intLop, mHeadCol.请购数量)) > 9999999999# Then
                            MsgBox "第" & intLop & "行卫生材料的请购数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.请购数量
                            Exit Function
                        End If
                    End If
                    If Val(.TextMatrix(intLop, mHeadCol.单价)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的成本价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.单价
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mHeadCol.金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If mint编辑状态 = 3 Then
                            .Col = mHeadCol.计划数量
                        Else
                            .Col = mHeadCol.请购数量
                        End If
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
    Dim lng序号 As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim 计划类型_IN As Integer
    Dim 期间_IN As String
    Dim 库房ID_IN As Long
    Dim 编制方法_IN As Integer
    Dim 编制人_IN As String
    Dim 编制日期_IN As String
    Dim 编制说明_IN As String

    Dim 材料ID_IN As Long
    Dim dbl计划数量_IN As Double
    Dim dbl单价_IN As Double
    Dim dbl金额_IN As Double
    Dim dbl请购数量_IN As Double
    Dim 上期数量_IN As Double
    Dim 库存数量_IN As Double
    Dim 上次供应商_IN As String
    Dim 上次生产商_IN As String, intMonth As Integer
    Dim lng部门ID As Long
    Dim intRow As Integer
    Dim cllTemp As New Collection
    SaveCard = False
    Select Case cbo类型.ListIndex + 1
        Case 1       '月计划
            mstr期间 = Format(DateAdd("m", 1, sys.Currentdate), "yyyyMM")
        Case 2       '季计划
            intMonth = Month(DateAdd("Q", 1, sys.Currentdate))
            mstr期间 = Format(DateAdd("Q", 1, sys.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
        Case Else    '年计划
            mstr期间 = Format(DateAdd("yyyy", 1, sys.Currentdate), "yyyy")
    End Select
            
    With mshBill
        ID_IN = sys.NextId("材料采购计划")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = sys.GetNextNo(85, mlng库房id)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        计划类型_IN = cbo类型.ListIndex + 1
        编制方法_IN = mint编制方法
        If cboEnterStock.ListIndex < 0 Then
            库房ID_IN = 0
        Else
            库房ID_IN = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        End If
        lng部门ID = cboStock.ItemData(cboStock.ListIndex)
        编制人_IN = gstrUserName
        编制日期_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        编制说明_IN = Trim(txt摘要.Text)
        期间_IN = mstr期间

        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_材料计划管理_DELETE('" & mlng计划ID & "')"
            cllTemp.Add gstrSQL
        End If
        'Zl_材料计划管理主表_Insert
        gstrSQL = "Zl_材料计划管理主表_Insert("
        '  Id_In       In 材料采购计划.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  单据_In     In 材料采购计划.单据%Type,
        gstrSQL = gstrSQL & "" & 1 & ","
        '  No_In       In 材料采购计划.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  计划类型_In In 材料采购计划.计划类型%Type,
        gstrSQL = gstrSQL & "" & 计划类型_IN & ","
        '  期间_In     In 材料采购计划.期间%Type,
        gstrSQL = gstrSQL & "'" & 期间_IN & "',"
        '  库房id_In   In 材料采购计划.库房id%Type,
        gstrSQL = gstrSQL & "" & IIf(库房ID_IN = 0, "NULL", 库房ID_IN) & ","
        '  部门id_In   In 材料采购计划.部门id%Type,
        gstrSQL = gstrSQL & "" & lng部门ID & ","
        '  编制方法_In In 材料采购计划.编制方法%Type,
        gstrSQL = gstrSQL & "" & 编制方法_IN & ","
        '  编制人_In   In 材料采购计划.编制人%Type,
        gstrSQL = gstrSQL & "'" & 编制人_IN & "',"
        '  编制日期_In In 材料采购计划.编制日期%Type,
        gstrSQL = gstrSQL & "to_date('" & 编制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  编制说明_In In 材料采购计划.编制说明%Type := Null
        gstrSQL = gstrSQL & "'" & 编制说明_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng序号 = .TextMatrix(intRow, mHeadCol.序号)
                材料ID_IN = .TextMatrix(intRow, 0)
                dbl单价_IN = Round(Val(.TextMatrix(intRow, mHeadCol.单价)) / Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.成本价小数)
                dbl金额_IN = Round(Val(.TextMatrix(intRow, mHeadCol.金额)), g_小数位数.obj_最大小数.金额小数)
                dbl请购数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.请购数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl计划数量_IN = Round(Val(.TextMatrix(intRow, mHeadCol.计划数量)) * Val(.TextMatrix(intRow, mHeadCol.比例系数)), g_小数位数.obj_最大小数.数量小数)
                上次供应商_IN = .TextMatrix(intRow, mHeadCol.上次供应商)
                上次生产商_IN = .TextMatrix(intRow, mHeadCol.产地)
                'Zl_材料计划管理次表_Insert
                gstrSQL = "Zl_材料计划管理次表_Insert("
                '  计划id_In     In 材料计划内容.计划id%Type,
                gstrSQL = gstrSQL & "" & ID_IN & ","
                '  材料id_In     In 材料计划内容.材料id%Type,
                gstrSQL = gstrSQL & "" & 材料ID_IN & ","
                '  序号_In       In 材料计划内容.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  请购数量_In   In 材料计划内容.请购数量%Type,
                gstrSQL = gstrSQL & "" & dbl请购数量_IN & ","
                '  计划数量_IN   In 材料计划内容.计划数量%Type,
                gstrSQL = gstrSQL & "" & dbl计划数量_IN & ","
                '  单价_IN       In 材料计划内容.单价%Type,
                gstrSQL = gstrSQL & "" & dbl单价_IN & ","
                '  金额_IN       In 材料计划内容.金额%Type,
                gstrSQL = gstrSQL & "" & dbl金额_IN & ","
                '  前期数量_In   In 材料计划内容.前期数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  上期数量_In   In 材料计划内容.上期数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  库存数量_In   In 材料计划内容.库存数量%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  上次供应商_In In 材料计划内容.上次供应商%Type := Null,
                gstrSQL = gstrSQL & "'" & 上次供应商_IN & "',"
                '  上次生产商_In In 材料计划内容.上次生产商%Type := Null
                gstrSQL = gstrSQL & "'" & 上次生产商_IN & "')"
                cllTemp.Add gstrSQL
            End If
        Next
    End With
    On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim Dbl金额 As Double
    Dim intLop As Integer

    Dbl金额 = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl金额 = Dbl金额 + Val(.TextMatrix(intLop, mHeadCol.金额))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "金额合计：" & Format(Dbl金额, mFMT.FM_金额)
End Sub


Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme (True)
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
    OS.OpenIme False
End Sub

Private Function SetStuffRows(ByVal lng材料ID As Long, ByVal str材料 As String, _
        ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, _
        ByVal dbl指导批发价 As Double, ByVal dbl比例系数 As Double) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng批次 As Long
    Dim dbl库存数量 As Double
    Dim dbl成本价 As Double

    On Error GoTo errH
    SetStuffRows = False

    With mshBill
        intRow = .Row
        For intCount = 1 To .Rows - 1
            If intCount <> intRow And .TextMatrix(intCount, 0) <> "" Then
                If .TextMatrix(intCount, 0) = lng材料ID Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'MsgBox "卫生材料【" & str材料 & "】已有了，不能再输！", vbOKOnly + vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next

        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
    End With
    If cboEnterStock.ListIndex < 0 Then
        mlng库房id = 0
    Else
        mlng库房id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
    With mshBill
        .TextMatrix(.Row, mHeadCol.序号) = .Row
        .TextMatrix(.Row, mHeadCol.产地) = str产地
        .TextMatrix(.Row, 0) = lng材料ID
        .TextMatrix(.Row, mHeadCol.比例系数) = dbl比例系数
        
        '取平均成本价（如果没有设置，则取指导批发价）
        gstrSQL = "Select 成本价,指导批发价,招标材料 From  材料特性 Where 材料ID=[1]"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取成本价", lng材料ID)
        
        dbl成本价 = zlStr.Nvl(rsData!成本价, 0)
        If dbl成本价 = 0 Then dbl成本价 = zlStr.Nvl(rsData!指导批发价, 0)
        .TextMatrix(.Row, mHeadCol.中标材料) = IIf(Val(zlStr.Nvl(rsData!招标材料)) = "1", "√", "")
        
        gstrSQL = "Select a.上次产地, b.名称 As 供应商 From 材料特性 A, 供应商 B Where a.上次供应商id = b.Id And a.材料id = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取上次供应商及产地信息", lng材料ID)
            
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.上次供应商) = IIf(IsNull(rsData!供应商), "", rsData!供应商)
            .TextMatrix(.Row, mHeadCol.产地) = IIf(IsNull(rsData!上次产地), str产地, rsData!上次产地)
        End If
        .TextMatrix(.Row, mHeadCol.材料) = str材料
        .TextMatrix(.Row, mHeadCol.规格) = str规格
        .TextMatrix(.Row, mHeadCol.单位) = str单位
        .TextMatrix(.Row, mHeadCol.单价) = Format(dbl成本价 * dbl比例系数, mFMT.FM_成本价)
        
    End With
    rsData.Close
    SetStuffRows = True
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

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

'按编码，名称，别名查找某一列
Private Function FindData(ByVal mshBill As BillEdit, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    Dim strKey As String
    FindData = True
    
    On Error GoTo ErrHandle
    With mshBill
        If .Rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(strCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .MsfObj.TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.编码 " & _
                  " FROM (SELECT DISTINCT A.收费细目id " & _
                  "       FROM 收费项目别名 A" & _
                  "       Where A.简码 LIKE [1]) a, 收费项目目录 B " & _
                  " Where a.收费细目id = b.ID And (b.站点=[2] or b.站点 is null) "
        
        strKey = IIf(gstrMatchMethod = "0", "%", "") & str比较值 & "%"
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strKey, gstrNodeNo)
                  
        If rsCode.EOF Then
            FindData = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .MsfObj.TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindData = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            If mint编辑状态 <> 3 Then
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 0
                Next
            End If
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            cbo类型.Enabled = -False
        Else
            cboStock.Enabled = True
            cboEnterStock.Enabled = True
            cbo类型.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub

Private Function GetDepend(ByVal lngStockID As Long) As Boolean
    Dim rsSQL As ADODB.Recordset
    Dim str站点限制 As String
    
    On Error GoTo ErrHandle
    str站点限制 = GetDeptStationNode(lngStockID)
    gstrSQL = "SELECT DISTINCT a.id,a.编码||'-'||a.名称  as 名称 " & _
              "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
              "Where c.工作性质 = b.名称 " & _
              IIf(str站点限制 <> "", " And (a.站点 = [2] or a.站点 is null) ", "") & _
              "  And b.编码 In('V','K') " & _
              "  AND a.id = c.部门id " & _
              "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order by 名称 "
    Set rsSQL = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, str站点限制)
    
    If rsSQL.EOF Then
        MsgBox "没有任何库房允许被申购，请在部门管理中设置对应部门的工作性质为[卫材库]或[制剂室]！", vbInformation, gstrSysName
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

Private Sub SetEnterStock(ByVal lngStockID As Long)
    Dim rsSQL As ADODB.Recordset
    Dim str站点限制 As String
    
    On Error GoTo ErrHandle
    str站点限制 = GetDeptStationNode(lngStockID)
    gstrSQL = "SELECT DISTINCT a.id,a.编码||'-'||a.名称  as 名称 " & _
              "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
              "Where c.工作性质 = b.名称 " & _
              IIf(str站点限制 <> "", " And (a.站点 = [2] or a.站点 is null) ", "") & _
              "  And b.编码 In('V','K') " & _
              "  AND a.id = c.部门id " & _
              "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order by 名称 "
    Set rsSQL = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, str站点限制)
    
    With cboEnterStock
        .Clear
        Do While Not rsSQL.EOF
            .AddItem zlStr.Nvl(rsSQL!名称)
            .ItemData(.NewIndex) = Val(zlStr.Nvl(rsSQL!Id))
            rsSQL.MoveNext
        Loop
        rsSQL.Close
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub
