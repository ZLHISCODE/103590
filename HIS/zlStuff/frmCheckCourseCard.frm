VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCheckCourseCard 
   Caption         =   "材料盘点记录单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCheckCourseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin MSMask.MaskEdBox TxtCheckDate 
      Height          =   315
      Left            =   9510
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   25
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   23
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   20
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   21
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   7
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
         TabIndex        =   11
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "盘点金额合计："
         Height          =   180
         Left            =   1920
         TabIndex        =   9
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
         TabIndex        =   5
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "盘点成本金额合计："
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   17
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
         TabIndex        =   15
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   13
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
         TabIndex        =   2
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
         TabIndex        =   10
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "材料盘点记录单"
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   12
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
         TabIndex        =   18
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
            Picture         =   "frmCheckCourseCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1000
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
            Picture         =   "frmCheckCourseCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
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
            Picture         =   "frmCheckCourseCard.frx":22EA
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
            Picture         =   "frmCheckCourseCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCourseCard.frx":3080
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
Attribute VB_Name = "frmCheckCourseCard"
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
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private mintDefault As Integer              '缺省单位
Private mint库存检查 As Integer             '表示卫生材料出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Dim mstrPrivs As String                     '权限
Private mbln盘无存储库房材料 As Boolean
Private Const mstrCaption As String = "卫材盘点记录单"
Private mstr重复卫材 As String '记录重复的卫材
Private mbln分批卫材批号产地控制 As Boolean  '是否检查分批卫材批号产地是否录入


'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Const mlngModule = 1719

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位

'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol材料 As Integer = 2
Private Const mconIntCol序号 As Integer = 3
Private Const mconIntCol规格 As Integer = 4
Private Const mconIntCol批次 As Integer = 5
Private Const mconIntCol可用数量 As Integer = 6
Private Const mconIntCol比例系数 As Integer = 7
Private Const mconIntCol指导差价率 As Integer = 8
Private Const mconIntCol实际差价 As Integer = 9
Private Const mconIntCol实际金额 As Integer = 10
Private Const mconIntCol产地 As Integer = 11
Private Const mconIntCol库房货位 As Integer = 12
Private Const mconIntCol单位 As Integer = 13
Private Const mconIntCol批号 As Integer = 14
Private Const mconIntCol效期 As Integer = 15
Private Const mconIntCol灭菌效期 As Integer = 16
Private Const mconintCol帐面数量 As Integer = 17
Private Const mconintCol实盘数量 As Integer = 18
Private Const mconintCol标志 As Integer = 19
Private Const mconintCol数量差 As Integer = 20
Private Const mconIntCol成本价 As Integer = 21
Private Const mconIntCol成本金额 As Integer = 22
Private Const mconIntCol售价 As Integer = 23
Private Const mconIntCol售价金额 As Integer = 24
Private Const mconintCol金额差 As Integer = 25
Private Const mconintCol差价差 As Integer = 26
Private Const mconintCol盘点金额 As Integer = 27
Private Const mconintCol批号编辑 As Integer = 28
Private Const mconintCol产地编辑 As Integer = 29
Private Const mconIntColS  As Integer = 30             '总列数
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
        "           AND A.单据 = 39  and b.系数=1 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "卫生材料盘点管理"
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫生材料盘点记录单的入库类别，请在入出分类中设置！"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
                Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:显示或编辑单据
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

    Call GetRegInFor(g私有模块, "卫材盘点管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)

 

    
    If mint编辑状态 = 1 Then
'        If mbln单据增加 Then
'            mstr单据号 = NextNo(76)
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
        CmdSave.Visible = False
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
        FindRownew mshBill, mconIntCol材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
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
'        mstr单据号 = NextNo(76)
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

Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mint编辑状态 = 1 Then
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Else
'        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '正常
            Case 2
                '单据已被删除
                ShowMsgBox "该单据已被删除，请检查！"
                Unload Me
                Exit Sub
            Case 3
                '修改的单据已被审核
                ShowMsgBox "该单据已被其他人审核，请检查！"
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

Private Function GetDateStock(str盘存时间 As String, lng库房id As Long, str材料条件 As String, Optional blnZero As Boolean = False, Optional ByVal bln新增 As Boolean = False, Optional lng材料ID As Long = 0) As ADODB.Recordset
    '功能：获取指定条件材料在指定时间点的库存及相关信息
    '参数：str盘存时间=要求以YYYY-MM-DD HH24:MI:SS为格式的时间字符串
    '      str材料条件=" And B.材料ID=... And ..."
    '      blnZero=是否读取库存数结果为0的材料,缺省否.当强行输入该材料时,才设为是。
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    
    On Error GoTo errH
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From 部门性质说明 " & _
        "   WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) " & _
        "       AND 部门id =[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng库房id)
    
    If rsTemp.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If lng材料ID <> 0 Then
        str材料条件 = " And B.材料ID=[3]"
    End If
    '取得当前库存
    gstrSQL = "" & _
        "   SELECT a.库房id, b.材料id, NVL (批次, 0) AS 批次, a.实际数量,a.实际金额, a.实际差价, a.可用数量,a.上次批号 AS 批号,a.上次产地 AS 产地,a.效期,a.平均成本价 " & _
        "   FROM (Select 库房id,药品id,批次,实际数量,实际金额,实际差价,可用数量,上次批号,上次产地,效期,平均成本价 From 药品库存 Where 性质=1 And 库房ID+0=[1]) a, 材料特性 b,(Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 Where 库房ID+0=[1] )e " & _
        "   Where a.药品id(+) = b.材料id " & _
        "           and b.材料id=e.材料id(+) " & str材料条件
    '取得盘点时间后的净发生额
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "       SELECT a.库房id, b.材料id, NVL (a.批次, 0) AS 批次, " & _
        "               -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量, " & _
        "               -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
        "               -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价,0 AS 可用数量,a.批号,a.产地,a.效期,null 平均成本价 " & _
        "       FROM 药品收发记录 a, 材料特性 b,(Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 Where 库房ID+0=[1] )e " & _
        "       Where a.药品id = b.材料id " & _
        "               and b.材料id=e.材料id(+) " & _
        "               AND a.库房id + 0 =[1]" & _
        "               AND a.审核日期 > [2] " & str材料条件 & _
        "       GROUP BY a.库房id, b.材料id, a.批次,a.批号,a.产地,a.效期 "
    
    '取得盘点时间那一刻的帐面数量
    gstrSQL = "" & _
        "   SELECT 库房id, 材料id, 批次, SUM (实际数量) AS 帐面数量," & _
        "           SUM (实际金额) AS 实际金额, SUM (实际差价) AS 实际差价, " & _
        "           SUM(可用数量) As 可用数量,max(批号) as 批号,max(产地) as 产地 ,max(效期) as 效期, Max(平均成本价) As 平均成本价 " & _
        "   FROM ( " & gstrSQL & ") " & _
        "   GROUP BY 库房id, 材料id, 批次,平均成本价 "
    
    '(nvl(a.帐面数量,0) / b.住院包装) AS 住院帐面数量,(nvl(a.可用数量,0) / b.住院包装) AS 住院可用数量,
    If mintUnit = 0 Then
        strUnitQuantity = "c.计算单位 as 单位,'1' as 售价系数,f.售价 售价,"
    Else
        strUnitQuantity = "b.包装单位 as 单位,b.换算系数,f.售价*b.换算系数 as  售价,"
    End If
    
    gstrSQL = "" & _
        "   SELECT DISTINCT b.材料id, c.编码, c.名称 AS 商品名称,b.换算系数," & _
        "           zlSpellCode(c.名称) 名称,nvl(b.最大效期,0) 最大效期,c.规格,Decode(a.产地, Null, decode(b.上次产地,null,c.产地,b.上次产地), a.产地) As 产地,e.库房货位,a.批次, a.批号, a.效期," & strUnitQuantity & _
        "           nvl(a.实际金额,0) as 实际金额 ,nvl(a.实际差价,0) as 实际差价, b.指导差价率,c.是否变价,b.库房分批,b.在用分批,decode(a.平均成本价,null,b.成本价,a.平均成本价) 成本价,decode(a.批号,null,1,0) 批号编辑,decode(a.产地,null,1,0) 产地编辑 " & _
        "   From (" & gstrSQL & ") A ,材料特性 b,收费项目目录 c,(Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 Where 库房ID+0=[1] )e, " & _
        "        (  SELECT 收费细目id, 现价 as 售价 From 收费价目 WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL))" & _
        GetPriceClassString("") & ") f " & _
        "   Where " & IIf(blnZero = False, "a.材料id = b.材料id and a.材料id=c.id ", " b.材料id = a.材料id(+)and b.材料id=c.id ") & _
        "           AND b.材料id=f.收费细目id " & _
        "           and b.材料id=e.材料id(+) " & _
            IIf(blnZero, IIf(blnStock, " And (Nvl(b.库房分批,0)=1 ", " And (Nvl(b.在用分批,0)=1 "), "") & _
            IIf(blnZero, IIf(blnStock, " or Nvl(b.库房分批,0)=0)", " Or Nvl(b.在用分批,0)=0)"), "") & _
            IIf(blnZero = False, " AND (a.帐面数量<>0 or nvl(a.实际金额,0)<>0 or nvl(a.实际差价,0)<>0)", str材料条件) & _
        "   ORDER BY " & IIf(strCompare = "0", "c.编码", IIf(strCompare = "1", "c.编码", IIf(strCompare = "2", "c.名称", "e.库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
    
    Screen.MousePointer = 11
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng库房id, CDate(str盘存时间), lng材料ID)
     
    Set GetDateStock = rsTemp
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
    
     
    mintUnit = IIf(Val(zlDatabase.GetPara("记录单单位", glngSys, mlngModule, "0")) = 1, 1, 0)
    mbln分批卫材批号产地控制 = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    mblnFirst = True
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    
    mintBatchNoLen = GetBatchNoLen()
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    Call initCard
    RestoreWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    '库房
    
    On Error GoTo errHandle
    mbln盘无存储库房材料 = Val(zlDatabase.GetPara("存储库房", glngSys, mlngModule, "0"))
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
            TxtCheckDate.Text = Txt填制日期.Caption
            txtStock = mfrmMain.cboStock.Text
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            initGrid
        Case 2, 3, 4
            initGrid
            If mint编辑状态 <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "" & _
                    "   Select distinct b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   where a.库房id=b.id  and A.单据 = 23 and a.no=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsTemp!名称
                txtStock.Tag = rsTemp!Id
                rsTemp.Close
            End If
            
            '(Nvl(A.填写数量,0)/ B.门诊包装) AS 门诊帐面数量,(A.扣率/ B.门诊包装) AS 门诊实盘数量, (Nvl(A.实际数量,0) / B.门诊包装) AS 门诊数量差,
            Select Case mintUnit
            Case 0
                    strUnitQuantity = "A.扣率 实盘数量,A.实际数量,A.填写数量 帐面数量,A.实际数量 数量差,d.计算单位 AS 单位,'1' as 换算系数,a.零售价 as 售价售价,"
            Case Else
                    strUnitQuantity = "A.扣率/b.换算系数 as 实盘数量,A.填写数量/b.换算系数 帐面数量,A.实际数量/b.换算系数 数量差,B.包装单位 AS 单位,b.换算系数,a.零售价*b.换算系数 as 售价售价,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || d.编码 || ']' || d.名称) AS 材料信息," & _
                "                   zlSpellCode(d.名称) 名称,Nvl(B.最大效期,0) 最大效期,d.规格,A.产地,C.库房货位, A.批号,a.效期,a.批次," & strUnitQuantity & _
                "                   A.零售金额 as 金额差,A.差价 as 差价差, " & _
                "                   a.摘要,填制人,填制日期,审核人,审核日期,a.频次 as 盘点时间,a.成本价 as 库存金额,a.成本金额 as 库存差价,b.指导差价率,D.是否变价,b.在用分批,A.零售价,A.单量 As 成本价,decode(E.上次批号,null,1,0) 批号编辑,decode(E.上次产地,null,1,0) 产地编辑 " & _
                "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 D,材料储备限额 C,药品库存 E " & _
                "           Where A.药品id = B.材料id and a.药品id=D.id " & _
                "                   And A.药品ID=C.材料ID(+) And A.库房ID=C.库房ID(+) AND A.记录状态 =[2]" & _
                "                   And A.药品ID=E.药品ID(+) And A.库房ID=E.库房ID(+) And nvl(A.批次,0) = nvl(E.批次(+),0) AND A.单据 =23 AND A.No = [1] " & _
                "   ) " & _
                "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "材料信息", IIf(strCompare = "2", "名称", "库房货位"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mint记录状态)
            
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt填制人 = rsTemp!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户名
            End If
            Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd HH:mm:ss")
            
            Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
            Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd HH:mm:ss"))
            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            TxtCheckDate.Text = rsTemp!盘点时间
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = intRow + 1
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mconIntCol材料) = rsTemp!材料信息
                    .TextMatrix(intRow, mconIntCol序号) = rsTemp!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsTemp!库房货位), "", rsTemp!库房货位)
                    .TextMatrix(intRow, mconIntCol单位) = zlStr.NVL(rsTemp!单位)
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol指导差价率) = Format(rsTemp!指导差价率, mFMT.FM_金额) & "||" & rsTemp!是否变价 & "||" & rsTemp!在用分批
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = zlStr.NVL(rsTemp!换算系数)
                    .TextMatrix(intRow, mconintCol实盘数量) = Format(rsTemp.Fields("实盘数量").Value, mFMT.FM_数量)
                    
                    If Val(.TextMatrix(intRow, mconIntCol批次)) <> 0 Then '分批材料
                        .TextMatrix(intRow, mconintCol批号编辑) = rsTemp!批号编辑
                        .TextMatrix(intRow, mconintCol产地编辑) = rsTemp!产地编辑
                    End If
                    
                    .TextMatrix(intRow, mconIntCol成本价) = Format(rsTemp!成本价 * IIf(mintUnit = 0, 1, Val(.TextMatrix(intRow, mconIntCol比例系数))), mFMT.FM_成本价)
                    .TextMatrix(intRow, mconIntCol售价) = Format(rsTemp!零售价 * IIf(mintUnit = 0, 1, Val(.TextMatrix(intRow, mconIntCol比例系数))), mFMT.FM_零售价)
                    .TextMatrix(intRow, mconIntCol成本金额) = Format(Val(.TextMatrix(intRow, mconIntCol成本价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mFMT.FM_金额)
                    .TextMatrix(intRow, mconIntCol售价金额) = Format(Val(.TextMatrix(intRow, mconIntCol售价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mFMT.FM_金额)
                    
                    .RowData(intRow) = IIf(IsNull(rsTemp!最大效期), 0, rsTemp!最大效期)
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    mint库存检查 = Get出库检查(Val(txtStock.Tag))
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
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol材料) = "名称与编码"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol库房货位) = "库房货位"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = "失效期"
        .TextMatrix(0, mconIntCol灭菌效期) = "灭菌效期"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol指导差价率) = "指导差价率"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconintCol帐面数量) = "帐面数量"
        .TextMatrix(0, mconintCol实盘数量) = "实盘数量"
        .TextMatrix(0, mconintCol标志) = "标志"
        .TextMatrix(0, mconintCol数量差) = "数量差"
        .TextMatrix(0, mconIntCol成本价) = "成本价"
        .TextMatrix(0, mconIntCol成本金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol金额差) = "金额差"
        .TextMatrix(0, mconintCol差价差) = "差价差"
        .TextMatrix(0, mconintCol盘点金额) = "盘点金额"
        .TextMatrix(0, mconintCol批号编辑) = "批号编辑"
        .TextMatrix(0, mconintCol产地编辑) = "产地编辑"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol可用数量) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol指导差价率) = 0
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        
        .ColWidth(mconIntCol材料) = 2000
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol库房货位) = 2000
        .ColWidth(mconIntCol单位) = 0
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconintCol帐面数量) = 0
        .ColWidth(mconintCol实盘数量) = 1000
        
        .ColWidth(mconintCol标志) = 0
        .ColWidth(mconintCol数量差) = 0
        .ColWidth(mconIntCol成本价) = 800
        .ColWidth(mconIntCol成本金额) = 1200
        .ColWidth(mconIntCol售价) = 800
        .ColWidth(mconIntCol售价金额) = 1200
        .ColWidth(mconintCol金额差) = 0
        .ColWidth(mconintCol差价差) = 0
        .ColWidth(mconintCol盘点金额) = 0
        .ColWidth(mconIntCol灭菌效期) = 0
        
        .ColWidth(mconintCol批号编辑) = 0
        .ColWidth(mconintCol产地编辑) = 0
        
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
        .ColData(mconIntCol库房货位) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批次) = 5
        .ColData(mconIntCol可用数量) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol指导差价率) = 5
        .ColData(mconIntCol实际差价) = 5
        .ColData(mconIntCol实际金额) = 5
        .ColData(mconintCol帐面数量) = 5
        .ColData(mconintCol标志) = 5
        .ColData(mconintCol数量差) = 5
        .ColData(mconIntCol成本价) = 5
        .ColData(mconIntCol成本金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol金额差) = 5
        .ColData(mconintCol差价差) = 5
        .ColData(mconintCol盘点金额) = 5
        .ColData(mconIntCol灭菌效期) = 5
        
        .ColData(mconintCol批号编辑) = 5
        .ColData(mconintCol产地编辑) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mconIntCol材料) = 1
            .ColData(mconintCol实盘数量) = 4
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            
            .ColData(mconintCol实盘数量) = 5
        End If
        
        .ColAlignment(mconIntCol材料) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconintCol帐面数量) = flexAlignRightCenter
        .ColAlignment(mconintCol标志) = flexAlignCenterCenter
        .ColAlignment(mconintCol数量差) = flexAlignRightCenter
        .ColAlignment(mconIntCol成本价) = flexAlignRightCenter
        .ColAlignment(mconIntCol成本金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol金额差) = flexAlignRightCenter
        .ColAlignment(mconintCol差价差) = flexAlignRightCenter
        .ColAlignment(mconintCol盘点金额) = flexAlignRightCenter
        
        .ColAlignment(mconintCol批号编辑) = flexAlignRightCenter
        .ColAlignment(mconintCol产地编辑) = flexAlignRightCenter
        
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
    
    TxtCheckDate.Left = mshBill.Left + mshBill.Width - TxtCheckDate.Width
    lblCheckDate.Left = TxtCheckDate.Left - lblCheckDate.Width - 100
    
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
    
    On Error GoTo errHandle
    
    int点击行 = mshBill.Row
    
    If mshBill.Col = mconIntCol材料 Then
        If Not IsDate(TxtCheckDate) Then
            MsgBox "盘点时间不对,请输入!", vbInformation + vbDefaultButton1, gstrSysName
            If TxtCheckDate.Enabled Then TxtCheckDate.SetFocus
            Exit Sub
        End If
        Set RecReturn = Frm材料选择器.ShowMe(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, False, True, True, True, , , , , TxtCheckDate.Text, , , mbln盘无存储库房材料, mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    If SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)) Then
                        If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                        .Row = .Row + 1
                    End If
                    
                    RecReturn.MoveNext
                Next
                
                mshBill.Row = int点击行
                
                If mstr重复卫材 <> "" Then
                    MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                    mstr重复卫材 = ""
                End If
            
    '            If RecReturn.RecordCount = 1 Then
    '                Call SetPhiscRows(RecReturn!材料ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次))
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
            mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = NVL(!名称)
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
        If .Col = mconintCol实盘数量 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconintCol实盘数量
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
        
        If .Col = mconIntCol成本价 Then
            strKey = .Text
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            
            If InStr("0123456789.", Chr(KeyAscii)) > 0 Or Chr(KeyAscii) = vbBack Or Chr(KeyAscii) = vbCr Then '控制允许输入的值
                KeyAscii = KeyAscii
            Else
                KeyAscii = 0
            End If
            
            
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim lng批次 As Long
    
    With mshBill
        If .Active = False Then Exit Sub
        If mint编辑状态 = 4 Then Exit Sub
        If .Row <> .LastRow Then
            
        End If
        
        Select Case .Col
            Case mconIntCol材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mconIntCol批号
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mconIntCol效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .ColData(mconIntCol效期) = 2 Then
                    If .TextMatrix(.Row, mconIntCol批号) <> "" And Len(Trim(.TextMatrix(.Row, mconIntCol批号))) = 8 Then
                        Dim strxq As String
                        
                        If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq <> "" Then .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mconintCol实盘数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mconIntCol成本价
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                       .ColData(mconIntCol成本价) = IIf(Val(.TextMatrix(.Row, mconIntCol批次)) = -1, 4, 5)
                End If
        End Select
        
        lng批次 = Val(.TextMatrix(.Row, mconIntCol批次))
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            .ColData(mconIntCol产地) = IIf(lng批次 = -1 Or Val(.TextMatrix(.Row, mconintCol产地编辑)) = 1, 1, 5)
            .ColData(mconIntCol批号) = IIf(lng批次 = -1 Or Val(.TextMatrix(.Row, mconintCol批号编辑)) = 1, 4, 5)
            .ColData(mconIntCol效期) = IIf(lng批次 = -1, 2, 5)
        End If
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    On Error GoTo errHandle
    
    int点击行 = mshBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
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
                    If Not IsDate(TxtCheckDate) Then
                        MsgBox "盘点时间不对,请输入!", vbInformation + vbDefaultButton1, gstrSysName
                        If TxtCheckDate.Enabled Then TxtCheckDate.SetFocus
                        Exit Sub
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, True, True, True, , , , TxtCheckDate.Text, , , mbln盘无存储库房材料, mstrPrivs, , False)
                    
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
            Case mconIntCol产地
                If strKey = "" Then Exit Sub
                If SelectAndNotAddItem(Me, mshBill, strKey, "材料生产商", "材料生产商选择器", True, True, , zl_获取站点限制(True)) = True Then
                    .Text = .TextMatrix(.Row, .Col)
                Else
                    .Text = ""
                    .Col = mconIntCol产地
                    Cancel = True
                End If

            Case mconIntCol批号
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol批号) = ""
                    End If
                    If .ColData(mconIntCol效期) = 2 Then
                        .Col = mconIntCol效期
                    Else
                        .Col = mconintCol实盘数量
                    End If
                    
                    Cancel = True
                    Exit Sub
                End If
            Case mconIntCol效期
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
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol效期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mconintCol实盘数量
                If strKey <> "" Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        ShowMsgBox "实盘数量必须为数字型,请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                Else
                    .Text = IIf(.TextMatrix(.Row, .Col) = "", " ", .TextMatrix(.Row, .Col))
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If strKey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                End If
                
                '显示合计数量
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If .Col = mconintCol实盘数量 Then
                    strKey = Val(.Text) * Val(.TextMatrix(.Row, mconIntCol比例系数))
                Else
                    strKey = Val(.TextMatrix(.Row, mconintCol实盘数量)) * Val(.TextMatrix(.Row, mconIntCol比例系数))
                End If
                
                .TextMatrix(.Row, mconIntCol成本金额) = Format(Val(.TextMatrix(.Row, mconIntCol成本价)) * Val(.Text), mFMT.FM_金额)
                .TextMatrix(.Row, mconIntCol售价金额) = Format(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.Text), mFMT.FM_金额)
                
                Call 显示合计金额
            Case mconIntCol成本价
                If strKey <> "" Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        ShowMsgBox "成本价必须为数字型,请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_成本价)
                    .Text = strKey
                    
                    .TextMatrix(.Row, mconIntCol成本金额) = Format(Val(.TextMatrix(.Row, mconintCol实盘数量)) * Val(.Text), mFMT.FM_金额)
                End If
                
                
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

Private Sub TxtCheckDate_GotFocus()
    With TxtCheckDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheckDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub TxtCheckDate_Validate(Cancel As Boolean)
    
    If Not IsDate(TxtCheckDate.Text) Then
        ShowMsgBox "错误的时间格式!"
        Cancel = True
        Exit Sub
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
    Dim lng效期 As Long
    Dim rsTemp As New ADODB.Recordset
    
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
                ShowMsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!"
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol材料)) <> "" Then
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        ShowMsgBox "第" & intLop & "行卫生材料的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol批号
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintCol实盘数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol实盘数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol批次)) = -1 Then  '分批卫材检查产地和批号
                        
                        '判断是否为效期卫生材料
                        gstrSQL = "Select Nvl(最大效期,0) 效期 From 材料特性 Where 材料ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否为效期卫生材料", Val(.TextMatrix(intLop, 0)))
                        
                        lng效期 = rsTemp!效期
                        If lng效期 <> 0 Then
                            If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Or Trim(.TextMatrix(intLop, mconIntCol效期)) = "" Then
                                ShowMsgBox "第" & intLop & "行的卫生材料是效期材料,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mconIntCol批号) = "" Then
                                    .Col = mconIntCol批号
                                Else
                                    .Col = mconIntCol效期
                                End If
                                Exit Function
                            End If
                        End If
                        
                        If mbln分批卫材批号产地控制 = True Then
                            If Trim(.TextMatrix(intLop, mconIntCol产地)) = "" Then '产地必须输入
                                ShowMsgBox "第" & intLop & "行卫生材料是分批材料，请录入产地！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol产地
                                Exit Function
                            End If
                            
                            If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & intLop & "行卫生材料是分批材料，请录入批号！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol批号
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol批次)) > 0 Then '已有批次
                        If mbln分批卫材批号产地控制 = True Then
                            If Trim(.TextMatrix(intLop, mconIntCol产地)) = "" Then '产地必须输入
                                ShowMsgBox "第" & intLop & "行卫生材料是分批材料，请录入产地！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol产地
                                Exit Function
                            End If
                            
                            If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Then  '产地必须输入
                                ShowMsgBox "第" & intLop & "行卫生材料是分批材料，请录入批号！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol批号
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
    Dim dbl成本价 As Double
    Dim dbl售价 As Double
    Dim dbl金额差 As Double
    Dim dbl差价差 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str盘点时间 As String
    Dim dbl库存金额 As Double
    Dim dbl库存差价 As Double
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    On Error GoTo errHandle
    SaveCard = False
    '在外面设置入出类别ID，主要是所有卫生材料都要用他
    gstrSQL = "" & _
        "   SELECT b.系数,b.id AS 类别id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID " & _
        "           AND a.单据 = 39 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
    
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫生材料盘点管理的入出类别，请在入出分类中设置!"
        Exit Function
    End If
    
    lng入库类别ID = 0
    lng出库类别ID = 0
    
    If rsTemp!系数 = 1 Then lng入库类别ID = rsTemp!类别ID
    rsTemp.Close
    
    If lng入库类别ID = 0 Then
        ShowMsgBox "没有设置卫生材料盘点记录单的入库类别，请在入出分类中设置!"
        Exit Function
    End If
    
    With mshBill
        lng库房id = txtStock.Tag
        
        chrNo = Trim(txtNO.Text)
        If mint编辑状态 = 1 Then
            If chrNo <> "" Then
                If CheckNOExists(76, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(76, lng库房id)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
        str盘点时间 = TxtCheckDate.Text
        
        gcnOracle.BeginTrans
        If mint编辑状态 = 2 Then        '修改
            gstrSQL = "zl_材料盘点记录单_Delete('" & mstr单据号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        End If
            
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                lng批次ID = IIf(.TextMatrix(intRow, mconIntCol批次) = "", 0, .TextMatrix(intRow, mconIntCol批次))
                dat效期 = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                
                dbl帐面数量 = Round(Val(.TextMatrix(intRow, mconintCol帐面数量)) * IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol比例系数)), 1), g_小数位数.obj_最大小数.数量小数)
                dbl实盘数量 = Round(Val(.TextMatrix(intRow, mconintCol实盘数量)) * IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol比例系数)), 1), g_小数位数.obj_最大小数.数量小数)
                
                dbl数量差 = 0
                dbl成本价 = Round(Val(.TextMatrix(intRow, mconIntCol成本价)) / IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol比例系数)), 1), g_小数位数.obj_最大小数.成本价小数)
                dbl售价 = Round(Val(.TextMatrix(intRow, mconIntCol售价)) / IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol比例系数)), 1), g_小数位数.obj_最大小数.零售价小数)
                
                If Split(.TextMatrix(intRow, mconIntCol指导差价率), "||")(1) = 1 Then '时价
                    dbl售价 = Get零售价(lng材料ID, lng库房id, lng批次ID, 1) '取售价单位保存
                End If
                
                dbl金额差 = Round(Val(.TextMatrix(intRow, mconintCol金额差)), g_小数位数.obj_最大小数.金额小数)
                dbl差价差 = Round(Val(.TextMatrix(intRow, mconintCol差价差)), g_小数位数.obj_最大小数.金额小数)
                dbl库存金额 = Round(Val(.TextMatrix(intRow, mconIntCol实际金额)), g_小数位数.obj_最大小数.金额小数)
                dbl库存差价 = Round(Val(.TextMatrix(intRow, mconIntCol实际差价)), g_小数位数.obj_最大小数.金额小数)
                
                If dbl帐面数量 <= dbl实盘数量 Then
                    lng入出类别ID = lng入库类别ID
                    int入出系数 = 1
                Else
                    lng入出类别ID = lng出库类别ID
                    int入出系数 = -1
                End If
                 
                lng序号 = intRow
                
                'zl_材料盘点记录单_INSERT( /*NO_IN*/, /*序号_IN*/, /*库房ID_IN*/, /*批次_IN*/,
                    '/*入出类别ID_IN*/, /*入出系数_IN*/, /*材料ID_IN*/, /*帐面数量_IN*/,
                    '/*实盘数量_IN*/, /*数量差_IN*/, /*售价_IN*/, /*金额差_IN*/, /*差价差_IN*/,
                    '/*填制人_IN*/, /*填制日期_IN*/, /*摘要_IN*/, /*产地_IN*/, /*批号_IN*/,
                    '/*效期_IN*/, /*盘点时间_IN*/ );
                
                gstrSQL = "zl_材料盘点记录单_INSERT('" & _
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
                    str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                    str摘要 & "','" & _
                    str产地 & "','" & _
                    str批号 & "'," & _
                    IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    str盘点时间 & "'," & _
                    dbl库存金额 & "," & _
                    dbl库存差价 & ")"
                
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
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim dbl成本金额 As Double
    Dim dbl盘点金额 As Double
    Dim intLop As Integer

    dbl成本金额 = 0
    dbl盘点金额 = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl成本金额 = dbl成本金额 + Val(.TextMatrix(intLop, mconIntCol成本金额))
                dbl盘点金额 = dbl盘点金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "盘点成本金额合计：" & Format(dbl成本金额, mFMT.FM_金额)
    lblPurchasePrice.Width = Pic单据.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200

    lblCheckSum.Caption = "盘点金额合计：" & Format(dbl盘点金额, mFMT.FM_金额)
    lblCheckSum.Width = Pic单据.TextWidth(lblCheckSum.Caption)
'
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
        "   Where a.药品id=[1] and nvl(a.批次,0)=[2] " & _
        "           AND a.性质=1 " & _
        "           AND a.库房id =[3] "
        
    gstrSQL = strKc
    With mshBill
        If .TextMatrix(.Row, mconIntCol材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
'        gstrSQL = "" & _
            "   Select 可用数量/" & IIf(mintUnit = 0, 1, Val(.TextMatrix(.Row, mconIntCol比例系数))) & " as  可用数量 " & _
            "   From 药品库存 " & _
            "   where 库房id=[3]" & _
            "           and 药品id=[1]" & _
            "           and 性质=1 and " & _
            "           nvl(批次,0)=[2]"
        
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(txtStock.Tag), CDate(TxtCheckDate.Text), IIf(mintUnit = 0, 1, Val(.TextMatrix(.Row, mconIntCol比例系数))))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        stbThis.Panels(2).Text = "该卫生材料当前库存数为[" & Format(.TextMatrix(.Row, mconIntCol可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mconIntCol单位)
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
    '功能：根据卫生材料ID在盘存表上显示并处理该卫生材料的初始盘存信息
    '说明：
    '   1.如果是非库房分批药,且已经输入了,则提示并退出。
    '   2.如果是库房分批药，则分别处理该药的未处理的各批次库存行。
    Dim i As Integer, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim rsprice As New Recordset
    
    On Error GoTo errH
    
    SetPhiscRows = False
    Set rsData = GetDateStock(TxtCheckDate.Text, txtStock.Tag, "", True, True, lngId)
    intRecordCount = rsData.RecordCount
    If intRecordCount = 0 Then Exit Function
    '新增批次卫生材料
    If lng批次 <> -1 Then
        rsData.MoveFirst
        rsData.Find "批次=" & lng批次
        If rsData.EOF Then Exit Function
    End If
    
    With mshBill
        If lng批次 <> -1 Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    If Val(.TextMatrix(intRow, 0)) = lngId And IIf(.TextMatrix(intRow, mconIntCol批次) = "", "0", .TextMatrix(intRow, mconIntCol批次)) = lng批次 Then
                        If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & .TextMatrix(intRow, mconIntCol材料) & "，"  '最多记录三个重复的卫材
                        'ShowMsgBox "已有卫生材料【" & .TextMatrix(intRow, mconIntCol材料) & "(" & lng批次 & ")】，不再添加！"
                        Exit Function
                    End If
                End If
            Next
        End If
        
        mshBill.Redraw = False
        intRow = .Row
        intCurrentRow = .Row
        .TextMatrix(intRow, 0) = rsData!材料ID
        .TextMatrix(intRow, mconIntCol材料) = "[" & rsData!编码 & "]" & rsData!商品名称
        .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsData!规格), "", rsData!规格)
        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsData!产地), "", rsData!产地)
        .TextMatrix(intRow, mconIntCol库房货位) = IIf(IsNull(rsData!库房货位), "", rsData!库房货位)
        .TextMatrix(intRow, mconIntCol单位) = zlStr.NVL(rsData!单位)
        
        If lng批次 = -1 Then
            .TextMatrix(intRow, mconIntCol批次) = lng批次
            .TextMatrix(intRow, mconIntCol批号) = ""
            .TextMatrix(intRow, mconIntCol效期) = ""
        Else
            .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsData!批次), "0", rsData!批次)
            .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsData!批号), "", rsData!批号)
            .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsData!效期), "", Format(rsData!效期, "yyyy-MM-dd"))
        End If
        
        If Val(.TextMatrix(intRow, mconIntCol批次)) <> 0 Then
            .TextMatrix(intRow, mconintCol批号编辑) = rsData!批号编辑
            .TextMatrix(intRow, mconintCol产地编辑) = rsData!产地编辑
        End If
        
        .TextMatrix(intRow, mconIntCol比例系数) = IIf(mintUnit = 0, 1, zlStr.NVL(rsData!换算系数)) ' 获取比例系数(rsData)
        .TextMatrix(intRow, mconIntCol指导差价率) = rsData!指导差价率 & "||" & rsData!是否变价 & "||" & rsData!在用分批
        
        .TextMatrix(intRow, mconIntCol成本价) = Format(rsData!成本价 * Val(.TextMatrix(intRow, mconIntCol比例系数)), mFMT.FM_成本价)
        .TextMatrix(intRow, mconIntCol成本金额) = Format(Val(.TextMatrix(intRow, mconIntCol成本价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mFMT.FM_金额)
        
        If rsData!是否变价 = 1 Then
'            gstrSQL = "" & _
'                "   Select 实际金额/实际数量*" & IIf(mintUnit = 0, "1", zlStr.NVL(rsData!换算系数)) & " as  售价 " & _
'                "   From 药品库存 " & _
'                "   Where 库房id=[1] " & _
'                "           and 药品id=[2]" & _
'                "  and 性质=1 and 实际数量>0 and " & _
'                "  nvl(批次,0)=[3]"
'
'            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(txtStock.Tag), Val(zlStr.NVL(rsData!材料ID)), Val(zlStr.NVL(rsData!批次)))
'
'            If rsprice.EOF Then
'                .TextMatrix(intRow, mconIntCol售价) = Format(IIf(IsNull(rsData.Fields("售价").Value), 0, rsData.Fields("售价").Value), mFMT.FM_零售价)
'            Else
'                .TextMatrix(intRow, mconIntCol售价) = Format(rsprice.Fields(0), mFMT.FM_零售价)
'            End If
            
            .TextMatrix(intRow, mconIntCol售价) = Format(Get零售价(lngId, Val(txtStock.Tag), lng批次, Val(.TextMatrix(intRow, mconIntCol比例系数))), mFMT.FM_零售价)
            
        Else '定价
            gstrSQL = "SELECT  现价 as 售价 From 收费价目 WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL))" & _
                    GetPriceClassString("") & " And 收费细目id = [1] "
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.NVL(rsData!材料ID)))
            
            .TextMatrix(intRow, mconIntCol售价) = Format(rsprice.Fields(0) * Val(.TextMatrix(intRow, mconIntCol比例系数)), mFMT.FM_零售价)
        End If
        
        .TextMatrix(intRow, mconIntCol售价金额) = Format(Val(.TextMatrix(intRow, mconIntCol售价)) * Val(.TextMatrix(intRow, mconintCol实盘数量)), mFMT.FM_金额)
        
        .RowData(intRow) = IIf(IsNull(rsData!最大效期), 0, rsData!最大效期)
        rsData.MoveNext
        
        Call RefreshRowNO(mshBill, mconIntCol行号, 1)
        .Col = IIf(lng批次 = -1, mconIntCol产地, mconintCol实盘数量)
        mshBill.Redraw = True
    End With
    Call 提示库存数
    
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'在一行中插入
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim intOldRows As Integer
    Dim intLop As Integer
    Dim intExchange As Integer
    Dim intCol As Integer
    
    With mshBill
        blnHaveData = False
        intOldRows = .Rows - 1
        .Rows = .Rows + intRecordCount
        For intLop = intRow + 1 To intRecordCount
            If .TextMatrix(intLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For intExchange = .Rows - 1 To intOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(intExchange, intCol) = .TextMatrix(intExchange - intRecordCount, intCol)
                    .TextMatrix(intExchange - intRecordCount, intCol) = ""
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
'    Dim StrNo As String
'    StrNo = txtNO.Tag
'    Call FrmBillPrint.ShowME(Me, glngSys, "zl1_bill_1719", mint记录状态, mintUnit, 1719, "卫生材料盘点记录单", StrNo)
End Sub

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Call zlDatabase.OpenRecordset(rsBatchNolen, gstrSQL, "取字段长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


