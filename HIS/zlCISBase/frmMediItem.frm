VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品品种编辑"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmMediItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk辅助用药 
      Caption         =   "辅助用药"
      Height          =   210
      Left            =   6750
      TabIndex        =   30
      Top             =   4560
      Width           =   1050
   End
   Begin VB.ComboBox cbo抗生素 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox chk抗生素 
      Caption         =   "抗菌药物(&Q)"
      Height          =   270
      Left            =   5265
      TabIndex        =   32
      Top             =   5175
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6870
      TabIndex        =   38
      Top             =   6225
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmMediItem.frx":058A
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6225
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存退出(&O)"
      Height          =   350
      Left            =   5385
      TabIndex        =   35
      Top             =   6225
      Width           =   1215
   End
   Begin VB.CheckBox chk原料药 
      Caption         =   "原料药(&M)"
      Height          =   210
      Left            =   5265
      TabIndex        =   22
      Top             =   3690
      Width           =   1155
   End
   Begin VB.CheckBox chk新药 
      Caption         =   "新药(&W)"
      Height          =   210
      Left            =   6750
      TabIndex        =   26
      Top             =   3390
      Width           =   1155
   End
   Begin VB.CheckBox chk皮试 
      Caption         =   "皮试(&Y)"
      Height          =   210
      Left            =   6750
      TabIndex        =   27
      Top             =   3690
      Width           =   1155
   End
   Begin VB.CheckBox chk急救药 
      Caption         =   "急救药(&J)"
      Height          =   210
      Left            =   5265
      TabIndex        =   21
      Top             =   3390
      Width           =   1155
   End
   Begin VB.TextBox txt英文 
      Height          =   300
      Left            =   1230
      TabIndex        =   6
      Top             =   1575
      Width           =   3675
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Left            =   3135
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1200
      Width           =   1170
   End
   Begin VB.ComboBox cbo药品类型 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1575
      Width           =   1455
   End
   Begin VB.ComboBox cbo医保职务 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2295
      Width           =   1455
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   120
      Width           =   3360
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1215
      Width           =   1170
   End
   Begin VB.ComboBox cbo剂型 
      Height          =   300
      Left            =   3450
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1935
      Width           =   1470
   End
   Begin VB.ComboBox cbo毒理 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cbo货源 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   855
      Width           =   1455
   End
   Begin VB.ComboBox cbo价值 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   495
      Width           =   1455
   End
   Begin VB.ComboBox cbo梯次 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1215
      Width           =   1455
   End
   Begin VB.TextBox txt处方限量 
      Height          =   300
      Left            =   6510
      MaxLength       =   16
      TabIndex        =   19
      Text            =   "0"
      Top             =   2655
      Width           =   1455
   End
   Begin VB.ComboBox cbo单位 
      Height          =   300
      Left            =   1230
      TabIndex        =   7
      Top             =   1950
      Width           =   1155
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1230
      MaxLength       =   40
      TabIndex        =   3
      Top             =   855
      Width           =   3675
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1230
      MaxLength       =   13
      TabIndex        =   2
      Top             =   495
      Width           =   1935
   End
   Begin VB.ComboBox cbo处方职务 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1935
      Width           =   1455
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -285
      TabIndex        =   40
      Top             =   5895
      Width           =   8490
   End
   Begin VB.TextBox txt参考 
      Height          =   300
      Left            =   1230
      TabIndex        =   9
      Top             =   2295
      Width           =   3135
   End
   Begin VB.CommandButton cmd参考 
      Caption         =   "…"
      Height          =   285
      Left            =   4350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2295
      Width           =   285
   End
   Begin VB.CheckBox chk品种医嘱 
      Caption         =   "药品按品种下长期医嘱"
      Height          =   210
      Left            =   5265
      TabIndex        =   31
      Top             =   4890
      Width           =   2115
   End
   Begin VB.ComboBox cbo适用性别 
      Height          =   300
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3015
      Width           =   1455
   End
   Begin VB.CommandButton cmdDel参考 
      Height          =   285
      Left            =   4650
      Picture         =   "frmMediItem.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "…"
      Height          =   285
      Left            =   4545
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   135
      Width           =   285
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "保存后新增品种(&A)"
      Height          =   350
      Left            =   1500
      TabIndex        =   37
      Top             =   6225
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "保存后新增规格(&B)"
      Height          =   350
      Left            =   3450
      TabIndex        =   36
      Top             =   6225
      Width           =   1695
   End
   Begin VB.TextBox txtAtccode 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6510
      MaxLength       =   50
      TabIndex        =   34
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CheckBox chkTumour 
      Caption         =   "肿瘤药(&T)"
      Height          =   210
      Left            =   5265
      TabIndex        =   23
      Top             =   3990
      Width           =   1155
   End
   Begin VB.CheckBox chkSolvent 
      Caption         =   "溶媒(&M)"
      Height          =   210
      Left            =   6750
      TabIndex        =   28
      Top             =   3990
      Width           =   1155
   End
   Begin VB.CheckBox chk原研药 
      Caption         =   "原研药(&P)"
      Height          =   210
      Left            =   5265
      TabIndex        =   24
      Top             =   4275
      Width           =   1155
   End
   Begin VB.CheckBox chk专利药 
      Caption         =   "专利药"
      Height          =   210
      Left            =   6750
      TabIndex        =   29
      Top             =   4275
      Width           =   1110
   End
   Begin VB.CheckBox chk单独定价 
      Caption         =   "单独定价"
      Height          =   210
      Left            =   5265
      TabIndex        =   25
      Top             =   4575
      Width           =   1140
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4305
      Top             =   6555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":0A97
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":1031
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":15CB
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediItem.frx":1B65
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   660
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6495
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ZL9BillEdit.BillEdit msf别名 
      Height          =   2805
      Left            =   135
      TabIndex        =   11
      Top             =   3015
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4948
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "注：该品种建立于2003-09-01"
      Height          =   180
      Left            =   135
      TabIndex        =   62
      Top             =   5970
      Width           =   2340
   End
   Begin VB.Label lbl别名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "其他别名"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   61
      Top             =   2715
      Width           =   720
   End
   Begin VB.Label lbl英文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   60
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label Lbl药品类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品类型(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   59
      Top             =   1635
      Width           =   990
   End
   Begin VB.Label Lbl医保职务 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保职务(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   58
      Top             =   2370
      Width           =   990
   End
   Begin VB.Label Lbl处方职务 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "处方职务(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   57
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   56
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称简码(&S)              (拼音)               (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   55
      Top             =   1275
      Width           =   4680
   End
   Begin VB.Label Lbl剂型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "剂型(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2760
      TabIndex        =   54
      Top             =   1995
      Width           =   630
   End
   Begin VB.Label Lbl毒理 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "毒理分类(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   53
      Top             =   195
      Width           =   990
   End
   Begin VB.Label Lbl货源 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "货源情况(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   52
      Top             =   915
      Width           =   990
   End
   Begin VB.Label Lbl价值 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "价值分类(&V)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   51
      Top             =   555
      Width           =   990
   End
   Begin VB.Label Lbl梯次 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "用药梯次(&G)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   50
      Top             =   1275
      Width           =   990
   End
   Begin VB.Label Lbl处方限量 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "处方限量(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   49
      Top             =   2715
      Width           =   990
   End
   Begin VB.Label Lbl单位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "剂量单位(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   48
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "通用名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   47
      Top             =   915
      Width           =   990
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   46
      Top             =   555
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "参考项目(&F)"
      Height          =   255
      Left            =   165
      TabIndex        =   45
      Top             =   2355
      Width           =   1095
   End
   Begin VB.Label lbl适用性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用性别(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5265
      TabIndex        =   44
      Top             =   3075
      Width           =   990
   End
   Begin VB.Label lblAtccode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ATCCODE(&H)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5310
      TabIndex        =   43
      Top             =   5580
      Width           =   900
   End
End
Attribute VB_Name = "frmMediItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前材质：由Me.tag存放，分别为1-西成药，2-中成药，由上级程序传入
'   2、编辑状态：由Me.cmdCancel.Tag存放，分别为"增加"、"修改"、"查阅"，由上级程序传入
'---------------------------------------------------
Public lng分类id As Long        '被编辑的分类ID，上级程序传递进入
Public lng药名id As Long        '被编辑的药名ID，修改、查阅时由上级程序传递进入
Public strPrivs As String       '当前用户对本程序的权限，由上级别程序传递进入
Public lng抗生素 As Long         '当前的抗生素级别
Private mint编码规则 As Integer     '药品品种编码产生规则
Private mblnOK As Boolean       '记录确定按钮是否被点击了
Private mblnCancel As Boolean   '记录取消按钮是否被点击了

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Dim mstrMatch As String, strRefer As String '参考名称
Private mstrID As String, mstr名称 As String, mstr编码 As String
Private mblnLoad As Boolean      '记录窗体加载的次数

Private mlng编码长度 As Long
Private mlng简码长度 As Long
Private mint名称长度 As Integer
Private mint英文长度 As Integer
Private mstr所有记录 As String
Private mbln自管药 As Boolean

Public Sub ShowMe(ByVal bln自管药 As Boolean, ByVal frmPar As Form)
    mbln自管药 = bln自管药
    Me.Show vbModal, frmPar
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    
    gstrSql = "Select 编码 From 收费项目目录 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng编码长度 = rsTmp.Fields("编码").DefinedSize
        
    txt编码.MaxLength = mlng编码长度
    
    gstrSql = "Select 名称,简码 From 诊疗项目别名 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng简码长度 = rsTmp.Fields("简码").DefinedSize
    mint英文长度 = rsTmp.Fields("名称").DefinedSize
    
    txt拼音.MaxLength = mlng简码长度
    txt五笔.MaxLength = mlng简码长度
    txt英文.MaxLength = mint英文长度
    
    gstrSql = "Select 名称 From 诊疗项目目录 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mint名称长度 = rsTmp.Fields("名称").DefinedSize
    txt名称.MaxLength = mint名称长度
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cbo处方职务_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo单位_GotFocus()
    Me.cbo单位.SelStart = 0: Me.cbo单位.SelLength = 100
End Sub

Private Sub cbo单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Or (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Then KeyAscii = 0
    
End Sub

Private Sub cbo毒理_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo货源_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo剂型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo价值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkTumour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chkSolvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chk辅助用药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo梯次_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo抗生素_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo药品类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo适用性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cbo医保职务_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk急救药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk抗生素_Click()
    If Me.chk抗生素.Value = 1 Then
        Me.cbo抗生素.Enabled = True
        txtAtccode.Enabled = True
    Else
        Me.cbo抗生素.Enabled = False
        txtAtccode.Enabled = False
        txtAtccode.Text = ""
    End If
End Sub

Private Sub chk抗生素_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk皮试_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk品种医嘱_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk新药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk原料药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk原研药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk专利药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk单独定价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Dim strTemp As String
    Dim str别名 As String
    Dim i As Integer
    
    With msf别名
        For i = 1 To .Rows - 1
            str别名 = str别名 & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    strTemp = txt分类.Text & "|" & txt编码.Text & "|" & txt名称.Text & "|" & txt拼音.Text & "|" & txt五笔.Text & "|" & txt英文.Text & "|" & cbo单位.Text & "|" & _
                cbo剂型.Text & "|" & txt参考.Text & "|" & cbo毒理.Text & "|" & cbo价值.Text & "|" & cbo货源.Text & "|" & cbo梯次.Text & "|" & cbo药品类型.Text & "|" & _
                cbo处方职务.Text & "|" & cbo医保职务.Text & "|" & txt处方限量.Text & "|" & cbo适用性别.Text & "|" & chk急救药.Value & "|" & chk新药.Value & "|" & chk原料药.Value & "|" & _
                chk原研药.Value & "|" & chk专利药.Value & "|" & chk单独定价.Value & "|" & chk辅助用药.Value & "|" & chk皮试.Value & "|" & chk品种医嘱.Value & "|" & chk抗生素.Value & "|" & cbo抗生素.Text & "|" & str别名 & "|" & txtAtccode.Text
        If strTemp <> mstr所有记录 Then
        mblnCancel = True
        If MsgBox("有数据被修改了确定退出？", vbYesNo, gstrSysName) = vbYes Then
            gblnCancel = True
            Unload Me
        Else
            mblnCancel = False
        End If
    Else
        gblnCancel = True
        Unload Me
    End If
    Exit Sub
End Sub

Private Sub cmdDel参考_Click()
    Me.txt参考.Text = ""
    Me.txt参考.Tag = ""
    strRefer = ""
    Me.txt参考.SetFocus
End Sub

Private Sub cmdOK_Click()

    '编辑数据检查
    mblnOK = True
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入药品编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > mlng编码长度 Then
        MsgBox "药品编码的长度超长（最多" & mlng编码长度 & "个字符）！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Sub
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入通用名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > mint名称长度 Then
        MsgBox "通用名称长度超长（最多" & mint名称长度 & "个字符或" & Int(mint名称长度 / 2) & "个汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: Exit Sub
    End If
    If Trim(Me.cbo单位.Text) = "" Then
        MsgBox "请输入剂量单位！", vbInformation, gstrSysName
        Me.cbo单位.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.cbo单位.Text), vbFromUnicode)) > 10 Then
        MsgBox "剂量单位的长度超长（最多10个字符或5个汉字）！", vbInformation, gstrSysName
        Me.cbo单位.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtAtccode.Text), vbFromUnicode)) > 50 Then
        MsgBox "ATCCODE的长度超长（最多50个字符或25个汉字）！", vbInformation, gstrSysName
        Me.txtAtccode.SetFocus: Exit Sub
    End If
    
    '别名检查
    strTemp = ";" & Trim(Me.txt名称.Text) & ";" & Trim(Me.txt英文.Text)
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "别名存在重复（包括通用名称和英文名称）！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    
    '数据保存
    If Me.cmdCancel.Tag = "增加" Then
        lng药名id = Sys.NextId("诊疗项目目录")
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text), lng药名id) = True Then Exit Sub
    End If
    gstrSql = Me.txt分类.Tag & "," & lng药名id & ",'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt名称.Text) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.txt拼音.Text)) & "','" & MoveSpecialChar(Trim(Me.txt五笔.Text)) & "','" & MoveSpecialChar(Trim(Me.txt英文.Text)) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.cbo单位.Text)) & "','" & Mid(Me.cbo剂型.Text, InStr(1, Me.cbo剂型.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo毒理.Text, InStr(1, Me.cbo毒理.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo价值.Text, InStr(1, Me.cbo价值.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo货源.Text, InStr(1, Me.cbo货源.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo梯次.Text, InStr(1, Me.cbo梯次.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Left(Me.cbo药品类型.Text, 1) & ",'" & Left(Me.cbo处方职务.Text, 1) & Left(Me.cbo医保职务.Text, 1) & "'"
    gstrSql = gstrSql & "," & Val(Trim(Me.txt处方限量.Text))
    gstrSql = gstrSql & "," & Me.chk急救药.Value & "," & Me.chk新药.Value & "," & Me.chk原料药.Value & "," & Me.chk皮试.Value & "," & IIf(Me.chk抗生素.Value = 0, Me.chk抗生素.Value, Me.cbo抗生素.ListIndex + 1)
    gstrSql = gstrSql & "," & ZVal(Me.txt参考.Tag) & "," & Me.chk品种医嘱.Value & "," & Left(Me.cbo适用性别.Text, 1)
    strTemp = ""
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    '检查别名串长度
    If LenB(strTemp) > 4000 Then
        msf别名.SetFocus
        MsgBox "别名字符串太长，请减少别名个数或者别名长度。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    err = 0: On Error GoTo ErrHand
    If Me.cmdCancel.Tag = "增加" Then
        gstrSql = "zl_成药品种_INSERT('" & IIf(Me.Tag = 1, "5", "6") & "'," & gstrSql & ",'" & strTemp & "'," & IIf(mbln自管药 = True, "1", "Null") & "," & IIf(Trim(txtAtccode.Text) = "", "NULL,", "'" & txtAtccode.Text & "',") & Me.chkTumour.Value & "," & Me.chkSolvent.Value & "," & Me.chk原研药.Value & "," & Me.chk专利药.Value & "," & Me.chk单独定价.Value & "," & Me.chk辅助用药.Value & ")"
    Else
        gstrSql = "zl_成药品种_UPDATE(" & gstrSql & ",'" & strTemp & "'," & IIf(mbln自管药 = True, "1", "Null") & "," & IIf(Trim(txtAtccode.Text) = "", "NULL,", "'" & txtAtccode.Text & "',") & Me.chkTumour.Value & "," & Me.chkSolvent.Value & "," & Me.chk原研药.Value & "," & Me.chk专利药.Value & "," & Me.chk单独定价.Value & "," & Me.chk辅助用药.Value & ")"
    End If
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    
    If Me.cmdCancel.Tag = "增加" Then
        'Val(zldatabase.GetPara("品种增加模式", glngSys, 1023, 0))
        Select Case ActiveControl
        Case cmdSaveAddItem  '品种连续增加
            Call frmMediLists.zlRefRecords(lng药名id)
            lng药名id = 0
            Call Form_Activate
            Me.txt名称.SetFocus
            mblnOK = False
        Case cmdSaveAddSpec  '品种增加后增加规格
            Unload Me
            mblnOK = False
            Call frmMediLists.zlRefRecords(lng药名id)
            With frmMediSpec
                .mlng分类id = lng分类id
                .stbSpec.Tag = "增加"
                .lng药名id = lng药名id
                .lng药品ID = 0
                .strPrivs = Me.strPrivs
                .Show 1, frmMediLists
            End With
        Case Else
            Unload Me
        End Select
    Else
        Unload Me
    End If
    
    If lng抗生素 <> 0 And mblnOK = True Then Call frmMediLists.ZlRefBut(3)
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    
    Call cmdOK_Click
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmd参考_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo errHandle
    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng分类id)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 编码"
    Else
        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
            " Where A.ID=B.参考目录ID And A.类型=" & iAttr & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.简码) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = " Select DISTINCT A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 0, "参考", , , , , True)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmd分类_Click()
    Dim blnRe As Boolean
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 = " & Me.Tag & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    blnRe = frmTreeSel.ShowTree(gstrSql, mstrID, mstr名称, mstr编码, "", "药品分类", "所有分类", False)
    If blnRe Then
        txt分类.Text = "[" & mstr编码 & "]" & mstr名称
        txt分类.Tag = mstrID
        Me.txt分类.SetFocus
        lng分类id = mstrID
    End If
    mblnLoad = True
End Sub

Private Sub Command2_Click()
    Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    Dim strCode As String
    Dim str别名 As String
    Dim i As Integer
    
    gblnCancel = False
    If cmdCancel.Tag <> "增加" Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    
    If mblnLoad = False Then
        If Me.Tag = "1" Then
            Me.Caption = "西成药品种" & Me.cmdCancel.Tag
        Else
            Me.Caption = "中成药品种" & Me.cmdCancel.Tag
        End If
        
        '基础数据检测
        If Me.cbo剂型.ListCount = 0 Then MsgBox "未设置药品剂型，请在字典管理中进行设置", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo毒理.ListCount = 0 Then MsgBox "无毒理分类数据，请联系系统管理员", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo价值.ListCount = 0 Then MsgBox "无价值分类数据，请联系系统管理员", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo货源.ListCount = 0 Then MsgBox "无货源分类数据，请联系系统管理员", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo梯次.ListCount = 0 Then MsgBox "无用药梯次数据，请联系系统管理员", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo抗生素.ListCount <> 0 Then
            Me.cbo抗生素.ListIndex = IIf(lng抗生素 <> 0, lng抗生素 - 1, 0)
            If lng抗生素 <> 0 Then
                Me.chk抗生素.Value = 1
                Me.cbo抗生素.Enabled = True
                txtAtccode.Enabled = True
            End If
        End If
        
        '新增或修改时，根据权限限制
        If InStr(1, strPrivs, "医保用药目录") = 0 Then
            Me.cbo医保职务.Enabled = False
        End If
        
        '新增或修改时，装入本类型的分类
        err = 0: On Error GoTo ErrHand
        
        '分类选择树装入
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = [1] and id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, lng分类id)
        
        With rsTemp
            Me.txt分类.Text = "[" & !编码 & "]" & !名称
            Me.txt分类.Tag = lng分类id
        End With
        
        '根据编辑状态，调整各数据显示
        If Me.cmdCancel.Tag = "增加" Then
            lng药名id = 0
    
            If mint编码规则 = 0 Then
'                gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
'                        " From 诊疗项目目录" & _
'                        " Where 类别 = [1]"

                gstrSql = "Select Nvl(Max(编码), '0000000') As 编码" & vbNewLine & _
                                "From (Select 编码 From 诊疗项目目录 Where 类别 = [1] Order By Length(编码) Desc, 编码 Desc, 建档时间 Desc)" & vbNewLine & _
                                "Where Rownum = 1 "
                                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = 1, "5", "6"))
                
'                Me.txt编码.Text = Right(String(13, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
                Me.txt编码.Text = zlCommFun.IncStr(rsTemp!编码)

            Else
                strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
'                gstrSql = "select nvl(max(编码),'') as 编码" & _
'                        " From 诊疗项目目录" & _
'                        " Where 类别 = [1] and 编码 like [2] and length(编码)>=[3]"

                gstrSql = "Select Nvl(Max(编码), '') As 编码" & vbNewLine & _
                                "From (Select 编码" & vbNewLine & _
                                "       From 诊疗项目目录" & vbNewLine & _
                                "       Where 类别 = [1] And 编码 Like [2] And Length(编码) >=[3] " & vbNewLine & _
                                "       Order By Length(编码) Desc, 编码 Desc, 建档时间 Desc)" & vbNewLine & _
                                "Where Rownum = 1"

                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = 1, "5", "6"), IIf(Me.Tag = 1, "5", "6") & strTemp & "%", Len("*" & strTemp & "**"))
                
                err = 0: On Error Resume Next
    '            Me.txt编码.Text = IIf(Me.Tag = 1, "5", "6") & strTemp & Right(String(13, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码) - 1 - Len(strTemp))
                
                strTemp = IIf(Me.Tag = 1, "5", "6") & strTemp
                If Nvl(rsTemp!编码) = "" Then
                    Me.txt编码.Text = strTemp & "01"
                Else
                    strCode = rsTemp!编码
                    strCode = Mid(strCode, Len(strTemp) + 1)
                    strCode = zlCommFun.IncStr(strCode)
                    Me.txt编码.Text = strTemp & strCode
                End If
            End If
    
            Me.txt名称.Text = "": Me.txt英文.Text = ""
            Me.lblNote.Visible = False
            Me.txt参考 = "": Me.txt参考.Tag = "": strRefer = ""
        Else
            '基本信息项目
            gstrSql = "select I.分类ID,I.编码,I.名称,I.计算单位,T.药品剂型," & _
                    "        T.毒理分类,T.货源情况,T.价值分类,T.用药梯次," & _
                    "        nvl(T.药品类型,0) as 药品类型,nvl(T.处方职务,'00') as 处方职务,nvl(T.处方限量,0) as 处方限量," & _
                    "        nvl(T.急救药否,0) as 急救药否,nvl(T.是否原料,0) as 是否原料,nvl(t.抗生素,0) as 抗生素,nvl(T.是否新药,0) as 是否新药,nvl(T.是否皮试,0) as 是否皮试,nvl(T.是否原研药,0) as 是否原研药,nvl(T.是否专利药,0) as 是否专利药,nvl(T.是否单独定价,0) as 是否单独定价,Nvl(t.是否辅助用药, 0) As 是否辅助用药," & _
                    "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,B.名称 as 参考名称,I.参考目录id,nvl(T.品种医嘱,0) as 品种医嘱,Nvl(I.适用性别,0) AS 适用性别,t.ATCCODE,nvl(T.是否肿瘤药,0) as 肿瘤药,T.溶媒" & _
                    " from 诊疗项目目录 I,药品特性 T,诊疗参考目录 B" & _
                    " where I.ID=T.药名ID and I.ID=[1] and I.参考目录id=B.id(+) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
            
            With rsTemp
                If Not .EOF Then
                    Me.lblNote.Caption = "注：该药品建立于" & Format(!建档时间, "YYYY-MM-DD")
                    If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                        Me.lblNote.Caption = Me.lblNote.Caption & "，于" & Format(!撤档时间, "YYYY-MM-DD") & "停用。"
                    End If
                    Me.txt编码.Text = !编码
                    Me.txt名称.Text = !名称
                    Me.cbo单位.Text = IIf(IsNull(!计算单位), "", !计算单位)
                    Me.txt参考.Text = Nvl(!参考名称)
                    Me.txt参考.Tag = Nvl(!参考目录ID)
                    strRefer = Me.txt参考.Text
                    For intCount = 0 To Me.cbo剂型.ListCount - 1
                        If Mid(Me.cbo剂型.List(intCount), InStr(1, Me.cbo剂型.List(intCount), "-") + 1) = IIf(IsNull(!药品剂型), "", !药品剂型) Then
                            Me.cbo剂型.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo毒理.ListCount - 1
                        If Mid(Me.cbo毒理.List(intCount), InStr(1, Me.cbo毒理.List(intCount), "-") + 1) = IIf(IsNull(!毒理分类), "", !毒理分类) Then
                            Me.cbo毒理.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo价值.ListCount - 1
                        If Mid(Me.cbo价值.List(intCount), InStr(1, Me.cbo价值.List(intCount), "-") + 1) = IIf(IsNull(!价值分类), "", !价值分类) Then
                            Me.cbo价值.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo货源.ListCount - 1
                        If Mid(Me.cbo货源.List(intCount), InStr(1, Me.cbo货源.List(intCount), "-") + 1) = IIf(IsNull(!货源情况), "", !货源情况) Then
                            Me.cbo货源.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo梯次.ListCount - 1
                        If Mid(Me.cbo梯次.List(intCount), InStr(1, Me.cbo梯次.List(intCount), "-") + 1) = IIf(IsNull(!用药梯次), "", !用药梯次) Then
                            Me.cbo梯次.ListIndex = intCount: Exit For
                        End If
                    Next
                    Me.cbo药品类型.ListIndex = !药品类型
                    Me.cbo处方职务.ListIndex = IIf(CInt(Left(Format(!处方职务, "00"), 1)) <> 9, CInt(Left(Format(!处方职务, "00"), 1)), Me.cbo处方职务.ListCount - 1)
                    Me.cbo医保职务.ListIndex = IIf(CInt(Right(Format(!处方职务, "00"), 1)) <> 9, CInt(Right(Format(!处方职务, "00"), 1)), Me.cbo医保职务.ListCount - 1)
                    Me.txt处方限量.Text = !处方限量
                    Me.txtAtccode.Text = IIf(IsNull(!ATCCODE), "", !ATCCODE)
                    Me.chk急救药.Value = IIf(!急救药否 = 0, 0, 1)
                    Me.chk原料药.Value = IIf(!是否原料 = 0, 0, 1)
                    Me.chk新药.Value = IIf(!是否新药 = 0, 0, 1)
                    Me.chk皮试.Value = IIf(!是否皮试 = 0, 0, 1)
                    Me.chkTumour.Value = IIf(!肿瘤药 = 0, 0, 1)
                    Me.chk原研药.Value = IIf(!是否原研药 = 0, 0, 1)
                    Me.chk专利药.Value = IIf(!是否专利药 = 0, 0, 1)
                    Me.chk单独定价.Value = IIf(!是否单独定价 = 0, 0, 1)
                    Me.chk辅助用药.Value = IIf(!是否辅助用药 = 0, 0, 1)
                    
                    '刘兴宏:2008/03/17加入抗生素,主要应用于院感系统的抗菌药物的监测:12753
                    Me.chk抗生素.Value = IIf(!抗生素 <> 0, 1, 0)
                    If !抗生素 <> 0 Then
                        Me.cbo抗生素.ListIndex = !抗生素 - 1
                    End If
                    Me.chk品种医嘱.Value = IIf(!品种医嘱 = 0, 0, 1)
                    Me.cbo适用性别.ListIndex = !适用性别
                    Me.chkSolvent.Value = IIf(Nvl(!溶媒, 0) = 0, 0, 1)
                End If
            End With
            
            '正名简码与英文名
            gstrSql = "select 名称,性质,简码,码类 from 诊疗项目别名 where 性质 in (1,2) and 诊疗项目ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
            
            With rsTemp
                Do While Not .EOF
                    If !性质 = 1 And !码类 = 1 Then Me.txt拼音.Text = !简码
                    If !性质 = 1 And !码类 = 2 Then Me.txt五笔.Text = !简码
                    If !性质 = 2 Then Me.txt英文.Text = !名称
                    .MoveNext
                Loop
            End With
                
            '其他别名
            gstrSql = "select N.名称,P.简码 as 拼音,W.简码 as 五笔" & _
                    " from (select distinct 名称 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9) N," & _
                    "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9 and 码类=1) P," & _
                    "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9 and 码类=2) W" & _
                    " where N.名称=P.名称(+) and N.名称=W.名称(+)"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药名id)
            
            With rsTemp
                Do While Not .EOF
                    If Me.msf别名.Rows - 1 < .AbsolutePosition Then Me.msf别名.Rows = Me.msf别名.Rows + 1
                    Me.msf别名.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                    Me.msf别名.TextMatrix(.AbsolutePosition, 1) = !名称
                    Me.msf别名.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!拼音), "", !拼音)
                    Me.msf别名.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!五笔), "", !五笔)
                    .MoveNext
                Loop
            End With
        End If
        
        If Me.cmdCancel.Tag = "查阅" Then
            '查阅时调整控件的编辑状态
            Me.cmdOK.Visible = False
            Me.cmdCancel.Caption = "关闭(&C)"
            Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
            Me.txt编码.Enabled = False
            Me.txt名称.Enabled = False
            Me.txt拼音.Enabled = False: Me.txt五笔.Enabled = False
            Me.txt英文.Enabled = False
            Me.cbo单位.Enabled = False: Me.cbo剂型.Enabled = False
            Me.cbo毒理.Enabled = False: Me.cbo价值.Enabled = False: Me.cbo货源.Enabled = False: Me.cbo梯次.Enabled = False
            Me.cbo药品类型.Enabled = False: Me.cbo处方职务.Enabled = False: Me.cbo医保职务.Enabled = False: Me.txt处方限量.Enabled = False: txtAtccode.Enabled = False
            Me.chk急救药.Enabled = False: Me.chk原料药.Enabled = False: Me.chk皮试.Enabled = False: Me.chk新药.Enabled = False
            Me.chk品种医嘱.Enabled = False
            Me.chk抗生素.Enabled = False
            Me.msf别名.Active = False
            Me.txt参考.Enabled = False
            Me.cmd参考.Enabled = False
            Me.cmdDel参考.Enabled = False
            Me.cbo适用性别.Enabled = False
            Me.chk原研药.Enabled = False
            Me.chk专利药.Enabled = False
            Me.chk单独定价.Enabled = False
            Me.chk辅助用药.Enabled = False
        End If
    End If
    mstr所有记录 = ""
    str别名 = ""
    With msf别名
        For i = 1 To .Rows - 1
            str别名 = str别名 & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    mstr所有记录 = txt分类.Text & "|" & txt编码.Text & "|" & txt名称.Text & "|" & txt拼音.Text & "|" & txt五笔.Text & "|" & txt英文.Text & "|" & cbo单位.Text & "|" & _
                cbo剂型.Text & "|" & txt参考.Text & "|" & cbo毒理.Text & "|" & cbo价值.Text & "|" & cbo货源.Text & "|" & cbo梯次.Text & "|" & cbo药品类型.Text & "|" & _
                cbo处方职务.Text & "|" & cbo医保职务.Text & "|" & txt处方限量.Text & "|" & cbo适用性别.Text & "|" & chk急救药.Value & "|" & chk新药.Value & "|" & chk原料药.Value & "|" & _
                chk原研药.Value & "|" & chk专利药.Value & "|" & chk单独定价.Value & "|" & chk辅助用药.Value & "|" & chk皮试.Value & "|" & chk品种医嘱.Value & "|" & chk抗生素.Value & "|" & cbo抗生素.Text & "|" & str别名 & "|" & txtAtccode.Text
    If txt名称.Enabled = True Then
        txt名称.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '取药品品种编码的产生规则
    mint编码规则 = Val(zlDatabase.GetPara(87, glngSys))
    
    Call GetDefineSize
    
    '-------------下拉选择数据装载-----------------------
    On Error GoTo errHandle
    With rsTemp
        gstrSql = "select 编码||'-'||名称 from 药品剂型 order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo剂型.Clear
        Do While Not rsTemp.EOF
            Me.cbo剂型.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo剂型.ListCount > 0 Then Me.cbo剂型.ListIndex = 0
    
        gstrSql = "select distinct 计算单位 from 诊疗项目目录 where 类别 in ('5','6') and 计算单位 is not null"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Do While Not rsTemp.EOF
            Me.cbo单位.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        
        gstrSql = "select 编码||'-'||名称 from 药品毒理分类 order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo毒理.Clear
        Do While Not rsTemp.EOF
            Me.cbo毒理.AddItem rsTemp.Fields(0).Value
            If InStr(1, rsTemp.Fields(0).Value, "普通") > 0 Then
                Me.cbo毒理.ListIndex = Me.cbo毒理.NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If Me.cbo毒理.ListIndex = -1 And Me.cbo毒理.ListCount > 0 Then Me.cbo毒理.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品价值分类 order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo价值.Clear
        Do While Not rsTemp.EOF
            Me.cbo价值.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo价值.ListCount > 0 Then Me.cbo价值.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品货源情况 order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo货源.Clear
        Do While Not rsTemp.EOF
            Me.cbo货源.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo货源.ListCount > 0 Then Me.cbo货源.ListIndex = 0
    
        gstrSql = "select 编码||'-'||名称 from 药品用药梯次 order by 编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo梯次.Clear
        Do While Not rsTemp.EOF
            Me.cbo梯次.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo梯次.ListCount > 0 Then Me.cbo梯次.ListIndex = 0
    End With
    
    aryTemp = Split("0-未设定;1-处方药;2-甲类非处方药;3-乙类非处方药;4-非处方药;5-其它用药", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo药品类型.AddItem aryTemp(intCount)
    Next
    Me.cbo药品类型.ListIndex = 0
    
    aryTemp = Split("0-不限;1-正高;2-副高;3-中级;4-助理/师级;5-员/士;9-待聘", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo处方职务.AddItem aryTemp(intCount): Me.cbo医保职务.AddItem aryTemp(intCount)
    Next
    Me.cbo处方职务.ListIndex = 0: Me.cbo医保职务.ListIndex = 0
    
    aryTemp = Split("0-无性别区分;1-男性;2-女性", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo适用性别.AddItem aryTemp(intCount)
    Next
    Me.cbo适用性别.ListIndex = 0
    
    aryTemp = Split("1-非限制使用;2-限制使用;3-特殊使用", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo抗生素.AddItem aryTemp(intCount)
    Next
    Me.cbo抗生素.ListIndex = 0
    
    '初始化设置表格编辑
    With Me.msf别名
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "药品名称": .TextMatrix(0, 2) = "拼音码": .TextMatrix(0, 3) = "五笔码"
        .colData(0) = 5: .colData(1) = 4: .colData(2) = 4: .colData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 2500: .ColWidth(2) = 950: .ColWidth(3) = 950
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    strRefer = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim str别名 As String
    Dim i As Integer
    
    If mblnOK = False And mblnCancel = False Then
        With msf别名
            For i = 1 To .Rows - 1
                str别名 = str别名 & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
            Next
        End With
        strTemp = txt分类.Text & "|" & txt编码.Text & "|" & txt名称.Text & "|" & txt拼音.Text & "|" & txt五笔.Text & "|" & txt英文.Text & "|" & cbo单位.Text & "|" & _
                    cbo剂型.Text & "|" & txt参考.Text & "|" & cbo毒理.Text & "|" & cbo价值.Text & "|" & cbo货源.Text & "|" & cbo梯次.Text & "|" & cbo药品类型.Text & "|" & _
                    cbo处方职务.Text & "|" & cbo医保职务.Text & "|" & txt处方限量.Text & "|" & cbo适用性别.Text & "|" & chk急救药.Value & "|" & chk新药.Value & "|" & chk原料药.Value & "|" & _
                    chk原研药.Value & "|" & chk专利药.Value & "|" & chk单独定价.Value & "|" & chk辅助用药.Value & "|" & chk皮试.Value & "|" & chk品种医嘱.Value & "|" & chk抗生素.Value & "|" & cbo抗生素.Text & "|" & str别名 & "|" & txtAtccode.Text
        If strTemp <> mstr所有记录 Then
            If MsgBox("有数据被修改了确定退出？", vbYesNo, gstrSysName) = vbYes Then
                mblnLoad = False
                mblnOK = False
                mblnCancel = False
                If mblnOK = True Then
                    gblnCancel = True
                End If
            Else
                Cancel = 1
            End If
        Else
            mblnLoad = False
            mblnOK = False
            mblnCancel = False
            If mblnOK = True Then
                gblnCancel = True
            End If
        End If
    End If
    mblnLoad = False
    mblnOK = False
    mblnCancel = False
End Sub

Private Sub msf别名_AfterAddRow(Row As Long)
    With Me.msf别名
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_AfterDeleteRow()
    With Me.msf别名
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_EditKeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf别名_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf别名
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Call OS.PressKey(vbKeyTab): Exit Sub
            strTemp = Trim(.Text)
            If strTemp <> "" Then
                .TextMatrix(.Row, 1) = strTemp
                .TextMatrix(.Row, 2) = zlStr.GetCodeByORCL(strTemp, False, mlng简码长度)
                .TextMatrix(.Row, 3) = zlStr.GetCodeByORCL(strTemp, True, mlng简码长度)
            Else
                Call OS.PressKey(vbKeyTab): Exit Sub
            End If
        End If
    End With
End Sub

Private Sub msf别名_KeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtAtccode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Asc("-")
        If InStr(1, txt编码.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt参考_GotFocus()
    Me.txt参考.SelStart = 0: Me.txt参考.SelLength = 100
End Sub


Private Sub txt参考_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt参考 <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt参考))
            If rsTmp Is Nothing Then
                Me.txt参考 = strRefer
                Me.SetFocus
                Exit Sub
            Else
                Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
            End If
        End If
        Call OS.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt参考_LostFocus()
    If Me.txt参考 <> strRefer Then
        Me.txt参考 = strRefer
    End If
End Sub


Private Sub txt处方限量_GotFocus()
    Me.txt处方限量.SelStart = 0: Me.txt处方限量.SelLength = 100
End Sub

Private Sub txt处方限量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    Dim strTmp As String
    '重新检查名称，并去 掉特殊字符
    strTmp = MoveSpecialChar(txt名称.Text)
    If txt名称.Text <> strTmp Then
        txt名称.Text = strTmp
    End If
    Me.txt拼音.Text = zlStr.GetCodeByORCL(strTmp, False, mlng简码长度)
    Me.txt五笔.Text = zlStr.GetCodeByORCL(strTmp, True, mlng简码长度)
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("？")
        Case Asc("%")
            KeyAscii = Asc("％")
        Case Asc("_")
            KeyAscii = Asc("＿")
    End Select
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    Else
        If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Me.txt拼音.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, False, mlng简码长度)
        Me.txt五笔.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, True, mlng简码长度)
    End If
    
        
End Sub

Private Sub txt名称_LostFocus()
'    Me.txt拼音.Text = zlGetSymbol(Me.txt名称.Text, 0, mlng简码长度)
'    Me.txt五笔.Text = zlGetSymbol(Me.txt名称.Text, 1, mlng简码长度)
    Call OS.OpenIme(False)
End Sub

Private Sub txt拼音_GotFocus()
    Me.txt拼音.SelStart = 0: Me.txt拼音.SelLength = 100
End Sub

Private Sub txt拼音_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt五笔_GotFocus()
    Me.txt五笔.SelStart = 0: Me.txt五笔.SelLength = 100
End Sub

Private Sub txt五笔_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt英文_GotFocus()
    Me.txt英文.SelStart = 0: Me.txt英文.SelLength = 100
End Sub

Private Sub txt英文_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("？")
        Case Asc("%")
            KeyAscii = Asc("％")
        Case Asc("_")
            KeyAscii = Asc("＿")
        Case vbKeyReturn
            Call OS.PressKey(vbKeyTab)
    End Select
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
