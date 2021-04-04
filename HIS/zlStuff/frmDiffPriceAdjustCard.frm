VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
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
      Left            =   255
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
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
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9975
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
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
         Left            =   960
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   26
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "调整额合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   24
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   1920
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "库存差价合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   20
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   3
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
      Caption         =   "材料"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDiffPriceAdjustCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln修改批发价 As Boolean           '允许修改批发价
Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位

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

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Dim mstrPrivs As String                     '权限

Private mint库存检查 As Integer             '表示卫生材料出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
'刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private Const mlngModule = 1715
Private Const mstrCaption As String = "库存差价调整单"
Private mstr重复卫材 As String '记录重复的卫材

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol材料 As Integer = 2
Private Const mconIntCol规格 As Integer = 3
Private Const mconIntCol批次 As Integer = 4
Private Const mconIntCol实际数量 As Integer = 5
Private Const mconIntCol比例系数 As Integer = 6
Private Const mconIntCol产地 As Integer = 7
Private Const mconIntCol单位 As Integer = 8
Private Const mconIntCol批号 As Integer = 9
Private Const mconIntCol效期 As Integer = 10
Private Const mconIntCol一次性材料 As Integer = 11
Private Const mconIntCol灭菌效期 As Integer = 12
Private Const mconIntCol灭菌日期    As Integer = 13
Private Const mconIntCol灭菌失效期 As Integer = 14
Private Const mconIntCol库存金额 As Integer = 15
Private Const mconintCol库存差价 As Integer = 16
Private Const mconintcol成本价 As Integer = 17
Private Const mconintcol新成本价 As Integer = 18
Private Const mconintCol调整额 As Integer = 19
Private Const mconIntColS  As Integer = 20              '总列数
'=========================================================================================


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset

    GetDepend = False
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID AND A.单据 = 33 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "库存差价调整"
    
    If rsTemp.EOF Then
        MsgBox "没有设置卫生材料库存差价调整的入出类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
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
        Optional int记录状态 As Integer = 1, Optional ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:显示或编辑单据,是唯一入口
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
    mblnFirst = True
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub

    mbln修改批发价 = IIf(Val(zlDatabase.GetPara("修改采购限价", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
   
    
    Call GetRegInFor(g私有模块, "库存差价调整管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    
    If mint编辑状态 = 1 Then
'        If mbln单据增加 Then
'            mstr单据号 = NextNo(71)
'        End If
        mblnEdit = True
        txtNO.Locked = True
        txtNO.TabStop = True

        txtNO = mstr单据号
        txtNO.Tag = mstr单据号
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
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
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
    If mint编辑状态 = 1 Then
        mshBill.ClearBill
        
        Dim str分类ID As String, lng库房ID As Long, int差价波动率 As Integer
        If frmDiffPriceAdjustCondition.GetCondition(mfrmMain, str分类ID, lng库房ID, int差价波动率) = True Then
        
            Screen.MousePointer = 11
            SearchData str分类ID, lng库房ID, int差价波动率
            Screen.MousePointer = 0
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
        mstrTime_End = GetBillInfo(18, txtNO.Tag)
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
        
'    If mbln单据增加 Then
'        mstr单据号 = NextNo(71)
'        txtNO = mstr单据号
'    End If
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    txt摘要.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
    
End Sub

Private Sub Form_Load()
   Dim strReg As String

    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
         
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    initCard
    RestoreWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
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
            txtStock = .List(.ListIndex)
            txtStock.Tag = .ItemData(.ListIndex)
            
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
                gstrSQL = "" & _
                "   Select distinct b.id,b.名称 " & _
                "   From 药品收发记录 a,部门表 b  " & _
                "   Where a.库房id=b.id and A.单据 =18 and  a.no=[1]"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                
                
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
                    strUnitQuantity = "c.计算单位 AS 单位, A.填写数量 as 可用数量,'1' as 比例系数,"
                Case Else
                    strUnitQuantity = "b.包装单位 AS 单位,(A.填写数量 / B.换算系数) AS 可用数量,B.换算系数 as 比例系数,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息, c.规格," & _
                "                   A.产地, A.批号,a.效期,a.灭菌日期,a.灭菌效期 as 灭菌失效期,b.一次性材料,b.灭菌效期,a.批次," & _
                "                   zlSpellCode(c.名称) 名称," & strUnitQuantity & _
                "                   A.成本价 as 库存差价,nvl(a.零售价,0) as 库存金额,A.差价 as 调整额,(nvl(a.零售价,0)-nvl(a.成本价,0))/a.填写数量 as 成本价,a.单量 as 新成本价, " & _
                "                   a.摘要,填制人,填制日期,审核人,审核日期,a.库房id " & _
                "           FROM 药品收发记录 A, 材料特性  b,收费项目目录 c" & _
                "           Where A.药品id = B.材料id and a.药品id=c.id " & _
                "                   AND A.记录状态 =[2]" & _
                "                   AND A.单据 =18 AND A.No = [1]" & _
                "           ) " & _
                " ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mint记录状态)
            
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            '刘兴宏:2007/06/10:问题10813
            mstrTime_Start = GetBillInfo(18, mstr单据号)
            
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
            
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = rsTemp.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mconIntCol材料) = rsTemp!卫材信息
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsTemp!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                   
                    .TextMatrix(intRow, mconIntCol一次性材料) = zlStr.Nvl(rsTemp!一次性材料)
                    .TextMatrix(intRow, mconIntCol灭菌效期) = zlStr.Nvl(rsTemp!灭菌效期)
                    .TextMatrix(intRow, mconIntCol灭菌日期) = IIf(IsNull(rsTemp!灭菌日期), "", Format(rsTemp!灭菌日期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol灭菌失效期) = IIf(IsNull(rsTemp!灭菌失效期), "", Format(rsTemp!灭菌失效期, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mconIntCol库存金额) = Format(rsTemp!库存金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mconintCol库存差价) = Format(IIf(IsNull(rsTemp!库存差价), 0, rsTemp!库存差价), mFMT.FM_金额)
                    .TextMatrix(intRow, mconintCol调整额) = Format(rsTemp!调整额, mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                    .TextMatrix(intRow, mconIntCol实际数量) = Format(IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), mFMT.FM_数量)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsTemp!比例系数
                    .TextMatrix(intRow, mconintcol成本价) = Format(rsTemp!成本价 * rsTemp!比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mconintcol新成本价) = Format(rsTemp!新成本价 * rsTemp!比例系数, mFMT.FM_成本价)
                    
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    mint库存检查 = Get出库检查(Val(txtStock.Tag))
    Exit Sub
ErrHandle:
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
        .TextMatrix(0, mconIntCol材料) = "名称与编码"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = "失效期"
        
        .TextMatrix(0, mconIntCol一次性材料) = "一次性材料"
        .TextMatrix(0, mconIntCol灭菌效期) = "灭菌效期"
        .TextMatrix(0, mconIntCol灭菌失效期) = "灭菌失效期"
        .TextMatrix(0, mconIntCol灭菌日期) = "灭菌日期"
         
        .TextMatrix(0, mconintCol库存差价) = "库存差价"
        .TextMatrix(0, mconIntCol库存金额) = "库存金额"
        .TextMatrix(0, mconintCol调整额) = "调整额"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol实际数量) = "库存数量"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconintcol成本价) = "成本价"
        .TextMatrix(0, mconintcol新成本价) = "新成本价"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol实际数量) = 0
        .ColWidth(mconIntCol比例系数) = 0
        
        .ColWidth(mconIntCol材料) = 2500
        .ColWidth(mconIntCol规格) = 1000
        .ColWidth(mconIntCol产地) = 1000
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol批号) = 1000
        .ColWidth(mconIntCol效期) = 1000
        
        .ColWidth(mconIntCol一次性材料) = 0
        .ColWidth(mconIntCol灭菌效期) = 0
        .ColWidth(mconIntCol灭菌失效期) = 1000
        .ColWidth(mconIntCol灭菌日期) = 1000
        
        .ColWidth(mconIntCol库存金额) = 1200
        .ColWidth(mconintCol库存差价) = 1200
        .ColWidth(mconintcol成本价) = 1200
        .ColWidth(mconintcol新成本价) = 1200
        .ColWidth(mconintCol调整额) = 1200
        
        
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
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        
        .ColData(mconIntCol一次性材料) = 5
        .ColData(mconIntCol灭菌效期) = 5
        .ColData(mconIntCol灭菌失效期) = 5
        .ColData(mconIntCol灭菌日期) = 2
          
        .ColData(mconintCol库存差价) = 5
        .ColData(mconIntCol库存金额) = 5
        .ColData(mconIntCol批次) = 5
        .ColData(mconIntCol实际数量) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconintcol成本价) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            
            
            .ColData(mconIntCol材料) = 1
            .ColData(mconintCol调整额) = 4
            .ColData(mconintcol新成本价) = 4
            
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            
            txt摘要.Enabled = False
            
            .ColData(mconintCol调整额) = 5
            .ColData(mconintcol新成本价) = 5
        End If
        
        .ColAlignment(mconIntCol材料) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol一次性材料) = flexAlignLeftCenter
        .ColAlignment(mconIntCol灭菌效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol灭菌失效期) = flexAlignCenterCenter
        .ColAlignment(mconIntCol灭菌日期) = flexAlignCenterCenter
        
        .ColAlignment(mconIntCol库存金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol实际数量) = flexAlignRightCenter
        .ColAlignment(mconintCol库存差价) = flexAlignRightCenter
        
        .ColAlignment(mconintCol调整额) = flexAlignRightCenter
        .ColAlignment(mconintcol新成本价) = flexAlignRightCenter
        .ColAlignment(mconintcol成本价) = flexAlignRightCenter
        
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
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
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
    Dim strNo As String
    Dim str审核人 As String
    
    mblnSave = False
    SaveCheck = False
    str审核人 = UserInfo.用户名
    strNo = txtNO.Tag
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_材料库存差价调整_Verify('" & strNo & "','" & str审核人 & "')"
    
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

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
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
    
    int点击行 = mshBill.Row
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 2, txtStock.Tag, , txtStock.Tag, False, , , , , , , , , , , , mstrPrivs & ";查看成本价;", , False)
    If RecReturn.RecordCount > 0 Then
        With mshBill
            Dim strUnit As String
            Dim intUnit As Integer
            
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                    IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
                     IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额)) Then
                    
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                    .Row = .Row + 1
                End If
                
                .Col = mconintcol新成本价
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int点击行
            
            If mstr重复卫材 <> "" Then
                MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                mstr重复卫材 = ""
            End If
            
'            If RecReturn.RecordCount = 1 Then
'
'                SetColValue .Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                    IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                    IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
'                     IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
'                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额)
'                 .Col = mconintcol新成本价
'
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
        If .Col = mconintcol新成本价 Or .Col = mconintCol调整额 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            
            Select Case .Col
                Case mconintcol新成本价
                   intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.成本价小数, g_小数位数.obj_散装小数.成本价小数)
                Case mconintCol调整额
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
        
        Select Case .Col
            Case mconIntCol材料
                .TxtCheck = False
                .MaxLength = 40
                '只在卫材列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mconintCol调整额
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890-"
          Case mconIntCol灭菌日期
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mconintcol新成本价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim dbl实际数量 As Double
    Dim dblMoney As Double
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
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, , txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, , , , , , , , , , , mstrPrivs, , False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                                IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
                                IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额)) Then
                                
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
'                                IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                                IIf(IsNull(RecReturn!批次), "0", RecReturn!批次), _
'                                IIf(IsNull(RecReturn!实际数量), "0", RecReturn!实际数量), _
'                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call 提示库存数
                End If
           
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
                        If MsgBox("该卫生材料已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
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
            Case mconintcol新成本价
               If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "成本价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "成本价必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "成本价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, 4)
                    .TextMatrix(.Row, .Col) = .Text
                End If
      
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_成本价)
                    .Text = strKey
                    .TextMatrix(.Row, mconintcol新成本价) = .Text
                    
                    '重算差价调整额(调整额＝库存金额－可用数量*成本价-库存差价)
                    dbl实际数量 = Val(.TextMatrix(.Row, mconIntCol实际数量))
                    dblMoney = Val(.TextMatrix(.Row, mconIntCol库存金额)) - dbl实际数量 * Val(strKey) - Val(.TextMatrix(.Row, mconintCol库存差价))
                       .TextMatrix(.Row, mconintCol调整额) = Format(dblMoney, mFMT.FM_金额)
                End If
            Case mconintCol调整额
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "调整额必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "调整额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "调整额不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strKey)) < 0.01 Then
                        MsgBox "调整额的必须大于0.01,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "调整额必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_金额)
                    .Text = strKey
                    
                    '重算成本价(成本价=(库存金额-库存差价-调整额)/实际数量)
                    dbl实际数量 = Val(.TextMatrix(.Row, mconIntCol实际数量))
                    If dbl实际数量 <> 0 Then
                        dblMoney = Val(.TextMatrix(.Row, mconIntCol库存金额)) - Val(.TextMatrix(.Row, mconintCol库存差价)) - Val(strKey)
                        dblMoney = dblMoney / dbl实际数量
                        .TextMatrix(.Row, mconintcol新成本价) = Format(dblMoney, mFMT.FM_成本价)
                    End If
                
                End If
                Call 显示合计金额
        End Select
    End With
End Sub

'从材料特性中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
    ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal str批号 As String, ByVal str效期 As String, _
    ByVal num库存差价 As Double, ByVal lng批次 As Long, ByVal num可用数量 As Double, _
    ByVal num比例系数 As Double, ByVal num库存金额 As Double) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    SetColValue = False
    gstrSQL = "Select 一次性材料,灭菌效期 from 材料特性 where 材料id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    
    
    With mshBill
        
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mconIntCol批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，不再添加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mconIntCol行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mconIntCol材料) = str材料
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol单位) = str单位
        
        .TextMatrix(intRow, mconIntCol一次性材料) = zlStr.Nvl(rsTemp!一次性材料)
        .TextMatrix(intRow, mconIntCol灭菌效期) = zlStr.Nvl(rsTemp!灭菌效期)
        
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol实际数量) = Format(num可用数量 / IIf(num比例系数 = 0, 1, num比例系数), mFMT.FM_数量)
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        .TextMatrix(intRow, mconIntCol库存金额) = Format(num库存金额, mFMT.FM_金额)
        .TextMatrix(intRow, mconintCol库存差价) = Format(num库存差价, mFMT.FM_金额)
        .TextMatrix(intRow, mconintcol成本价) = Format(Get成本价(lng材料ID, txtStock.Tag, lng批次) * num比例系数, mFMT.FM_成本价)
        
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
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "单据号超长,最多能输入" & CInt(txtNO.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNO.MaxLength & "个字符!"
            txtNO.SetFocus
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
                    If Trim(Trim(.TextMatrix(intLop, mconintCol调整额))) = "" Then
                        MsgBox "第" & intLop & "行卫生材料的调整额为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol调整额
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol调整额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫生材料的调整额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
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
    Dim lng入出类别ID As Long
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次ID As Long
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl可用数量 As Double
    Dim dbl库存差价 As Double
    Dim dbl库存金额 As Double
    Dim dbl调整额 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim dat填制日期 As String
    Dim rs入出类别 As New Recordset
    Dim str灭菌日期 As String
    Dim str灭菌效期 As String
    Dim dbl新成本价 As Double
      
    Dim intRow As Integer
    
    On Error GoTo ErrHandle
    SaveCard = False
    
    '在外面设置入出类别ID，主要是所有材料都要用他
    
    gstrSQL = "SELECT B.Id " _
        & " FROM 药品单据性质 A, 药品入出类别 B " _
        & "Where A.类别id = B.ID " _
      & "AND A.单据 = 33 "
    
    zlDatabase.OpenRecordset rs入出类别, gstrSQL, mstrCaption
    
    If rs入出类别.EOF Then
        MsgBox "没有设置卫生材料库存差价调整的入出类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
        rs入出类别.Close
        Exit Function
    End If
    lng入出类别ID = rs入出类别.Fields(0)
    rs入出类别.Close
    
    With mshBill
        chrNo = Trim(txtNO)
        lng库房ID = txtStock.Tag
        
        If mint编辑状态 = 1 Then   'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(71, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(71, lng库房ID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        dat填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        gcnOracle.BeginTrans
        If mint编辑状态 = 2 Then        '修改
            
            gstrSQL = "zl_材料库存差价调整_Delete('" & mstr单据号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        End If
            
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = Val(.TextMatrix(intRow, 0))
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                lng批次ID = Val(.TextMatrix(intRow, mconIntCol批次))
                str效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                str灭菌日期 = IIf(.TextMatrix(intRow, mconIntCol灭菌日期) = "", "", .TextMatrix(intRow, mconIntCol灭菌日期))
                str灭菌效期 = IIf(.TextMatrix(intRow, mconIntCol灭菌失效期) = "", "", .TextMatrix(intRow, mconIntCol灭菌失效期))
                
                dbl可用数量 = Round(Val(.TextMatrix(intRow, mconIntCol实际数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl库存金额 = Round(Val(.TextMatrix(intRow, mconIntCol库存金额)), g_小数位数.obj_最大小数.金额小数)
                dbl库存差价 = Round(Val(.TextMatrix(intRow, mconintCol库存差价)), g_小数位数.obj_最大小数.金额小数)
                dbl调整额 = Round(Val(.TextMatrix(intRow, mconintCol调整额)), g_小数位数.obj_最大小数.金额小数)
                dbl新成本价 = Round(Val(.TextMatrix(intRow, mconintcol新成本价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), g_小数位数.obj_最大小数.成本价小数)
                lng序号 = intRow
                
                'zl_材料库存差价调整_INSERT( /*入出类别ID_IN*/, /*NO_IN*/, /*序号_IN*/,
                    '/*库房ID_IN*/, /*材料ID_IN*/, /*批次_IN*/, /*可用数量_IN*/,
                    '/*库存差价_IN*/, /*调整额_IN*/, /*填制人_IN*/, /*填制日期_IN*/,
                    '/*产地_IN*/, /*批号_IN*/, /*效期_IN*/*灭菌日期_IN*/,/*灭菌效期_IN*//, /*摘要_IN*/ );
                    
                gstrSQL = "zl_材料库存差价调整_INSERT(" & lng入出类别ID & ",'" & chrNo & "'," & lng序号 & "," & _
                     lng库房ID & "," & lng材料ID & "," & lng批次ID & "," & dbl可用数量 & "," & _
                     dbl库存金额 & "," & dbl库存差价 & "," & dbl调整额 & ",'" & str填制人 & "',to_date('" & dat填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                     str产地 & "','" & str批号 & "'," & _
                     IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                     str摘要 & "'," & dbl新成本价 & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
        Next
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
    Dim dbl库存差价 As Double
    Dim dbl调整额 As Double
    Dim dbl库存金额 As Double
    
    Dim intLop As Integer
    
    dbl库存差价 = 0
    dbl调整额 = 0
    
    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl库存差价 = dbl库存差价 + Val(.TextMatrix(intLop, mconintCol库存差价))
                dbl库存金额 = dbl库存金额 + Val(.TextMatrix(intLop, mconIntCol库存金额))
                dbl调整额 = dbl调整额 + Val(.TextMatrix(intLop, mconintCol调整额))
            End If
        Next
    End With
    
    
    lblPurchasePrice.Caption = "库存金额合计：" & Format(dbl库存金额, mFMT.FM_金额)
    lblSalePrice.Caption = "库存差价合计：" & Format(dbl库存差价, mFMT.FM_金额)
    lblDifference.Caption = "调整额合计：" & Format(dbl调整额, mFMT.FM_金额)
    
End Sub

Private Sub 提示库存数()
    
    If mint编辑状态 = 4 Then Exit Sub
    With mshBill
        If .TextMatrix(.Row, mconIntCol材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        stbThis.Panels(2).Text = "该卫生材料当前库存数为[" & .TextMatrix(.Row, mconIntCol实际数量) & "]" & .TextMatrix(.Row, mconIntCol单位)
    End With
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
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1715", mint记录状态, mintUnit, 1715, "卫生材料差价调整单", strNo
    
End Sub


Private Sub SearchData(ByVal str分类ID, ByVal lng库房ID As Long, _
    ByVal intRate As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:根据相关条件，获取相关数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

    
    Dim rsTemp As New Recordset  '卫生材料库存记录集
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim intRecordCount As Integer
    
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    On Error GoTo ErrHandle:
    
    
    '设置界面显示内容
    
    stbThis.Panels(2).Text = "现在对" & txtStock & "的卫生材料进行自动差价计算"
        
    
    '构造卫生材料查询条件(材料特性)
    strPhysic = " And (c.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or c.撤档时间 is NULL)"
    
    If str分类ID <> "" Then
            strPhysic = strPhysic & " And d.分类ID IN(" & str分类ID & ")"
    End If
    
    DoEvents
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.计算单位 AS 单位, nvl(b.可用数量,0) AS 可用数量, '1' as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价,b.平均成本价) 成本价,"
                
        Case Else
            strUnitQuantity = "a.包装单位 AS 单位,(nvl(b.可用数量,0)/a.换算系数) AS 可用数量,a.换算系数 as 比例系数,decode(nvl(b.平均成本价,0),0,a.成本价,b.平均成本价) 成本价,"
    End Select
    
    gstrSQL = "" & _
        "   SELECT distinct  b.药品id 材料id,c.编码,c.名称 AS 商品名称, " & _
        "           c.规格, decode(b.上次产地,NULL,c.产地,b.上次产地) AS 产地,b.批次,b.上次批号 as 批号, b.效期,b.灭菌效期 as 灭菌失效期,a.灭菌效期,a.一次性材料," & _
        "           add_months(b.灭菌效期,-a.灭菌效期) as 灭菌日期," & _
        "           B.实际金额, B.实际差价, " & strUnitQuantity & _
        "           DECODE (SIGN (B.实际差价/B.实际金额*100-(A.指导差价率+" & intRate & ")),1,-(实际差价-B.实际金额*A.指导差价率/100)," & _
        "           DECODE (SIGN(B.实际差价/B.实际金额*100-(A.指导差价率-" & intRate & ")),-1,B.实际金额*A.指导差价率/100-实际差价)) AS 差价调整额 " & _
        "   FROM 材料特性 A,收费项目目录 c,诊疗项目目录 d, (Select 库房id, 药品id, 批次, 效期, 性质,实际数量 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号,平均成本价 From 药品库存 Where 性质=1 and  Nvl(实际金额,0)<>0) B " & _
        "   Where A.材料id=c.id and a.诊疗id=d.id  and A.材料ID = B.药品ID and B.性质(+)=1 AND B.库房id =[1]" & _
        "           AND ((B.批次>0 AND B.实际数量>0)  OR NVL(B.批次,0)=0) " & _
        "           AND ( NVL(B.实际金额,0)<>0 " & _
        "           AND (B.实际差价/Nvl(B.实际金额,1)*100>(A.指导差价率+" & intRate & ") OR B.实际差价/Nvl(B.实际金额,1)*100<A.指导差价率-" & intRate & _
        "               )) " & _
                strPhysic & _
        "   Order by c.编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "正在计算卫生材料库存数据", lng库房ID)
    
    
    intRecordCount = rsTemp.RecordCount
    
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    If intRecordCount = 0 Then
        MsgBox "未能正确读取卫生材料库存数据,请重试或手工输入卫生材料！", vbInformation, gstrSysName: Exit Sub
    End If
    
    DoEvents:
    mshBill.Redraw = False
    
    rsTemp.MoveFirst
    i = 1
    With mshBill
        Do While Not rsTemp.EOF
           If i > 1 Then .Rows = .Rows + 1
           .TextMatrix(i, 0) = rsTemp!材料ID
           .TextMatrix(i, mconIntCol材料) = "[" & rsTemp!编码 & "]" & rsTemp!商品名称
           .TextMatrix(i, mconIntCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
           .TextMatrix(i, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
           .TextMatrix(i, mconIntCol单位) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
           .TextMatrix(i, mconIntCol批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
           .TextMatrix(i, mconIntCol批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
           .TextMatrix(i, mconIntCol效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd"))
            
            .TextMatrix(i, mconIntCol一次性材料) = zlStr.Nvl(rsTemp!一次性材料)
            .TextMatrix(i, mconIntCol灭菌效期) = zlStr.Nvl(rsTemp!灭菌效期)
            .TextMatrix(i, mconIntCol灭菌日期) = IIf(IsNull(rsTemp!灭菌日期), "", Format(rsTemp!灭菌日期, "yyyy-mm-dd"))
            .TextMatrix(i, mconIntCol灭菌失效期) = IIf(IsNull(rsTemp!灭菌失效期), "", Format(rsTemp!灭菌失效期, "yyyy-mm-dd"))
            
           .TextMatrix(i, mconIntCol实际数量) = rsTemp!可用数量
           .TextMatrix(i, mconIntCol库存金额) = Format(rsTemp!实际金额, mFMT.FM_金额)
           .TextMatrix(i, mconintCol库存差价) = Format(rsTemp!实际差价, mFMT.FM_金额)
           .TextMatrix(i, mconintCol调整额) = Format(rsTemp!差价调整额, mFMT.FM_金额)
           .TextMatrix(i, mconintcol新成本价) = ""
           .TextMatrix(i, mconIntCol比例系数) = rsTemp!比例系数
           .TextMatrix(i, mconintcol成本价) = Format(rsTemp!成本价 * rsTemp!比例系数, mFMT.FM_成本价)
    
            Call ShowPercent(i / intRecordCount)
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = True
    End With
    rsTemp.Close
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1
    mshBill.Col = mconintCol调整额
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mshBill.Redraw = True
    Call SaveErrLog
    
End Sub

Private Sub ShowPercent(sngPercent As Single)
'功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub


