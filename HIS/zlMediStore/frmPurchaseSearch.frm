VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmPurchaseSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   5010
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmPurchaseSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1560
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   4815
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmPurchaseSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmPurchaseSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   5505
         Begin MSComctlLib.ListView lvw剂型 
            Height          =   2835
            Left            =   1200
            TabIndex        =   39
            Top             =   3360
            Visible         =   0   'False
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   5001
            View            =   1
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "imgsDrug"
            SmallIcons      =   "imgsDrug"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "名称"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TreeView tvw类别 
            Height          =   4245
            Left            =   120
            TabIndex        =   38
            Top             =   3240
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   7488
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imgsDrug"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.CheckBox chk发票日期 
            Caption         =   "发票审核日期"
            Height          =   405
            Left            =   600
            TabIndex        =   55
            Top             =   2340
            Width           =   1035
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品"
            Height          =   300
            Left            =   600
            TabIndex        =   54
            Top             =   1140
            Width           =   990
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1140
            Width           =   3255
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   52
            Top             =   1140
            Width           =   255
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   51
            Top             =   750
            Width           =   3255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   50
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Cmd生产商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   49
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox txt生产商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   48
            Top             =   1920
            Width           =   3255
         End
         Begin VB.CheckBox Chk生产商 
            Caption         =   "生产商"
            Height          =   300
            Left            =   600
            TabIndex        =   47
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   46
            Top             =   1560
            Width           =   255
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1530
            Width           =   3255
         End
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商"
            Height          =   300
            Left            =   600
            TabIndex        =   44
            Top             =   1530
            Width           =   1110
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   43
            Top             =   750
            Width           =   255
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "药品剂型"
            Height          =   300
            Left            =   600
            TabIndex        =   42
            Top             =   750
            Width           =   1095
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   41
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox chkClass 
            Caption         =   "药品分类"
            Height          =   300
            Left            =   600
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   9
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   10
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt开始发票号 
            Height          =   300
            Left            =   1530
            TabIndex        =   11
            Top             =   3330
            Width           =   1365
         End
         Begin VB.TextBox Txt结束发票号 
            Height          =   300
            Left            =   3780
            TabIndex        =   12
            Top             =   3330
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpStart发票 
            Height          =   315
            Left            =   1650
            TabIndex        =   35
            Top             =   2340
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpEnd发票 
            Height          =   315
            Left            =   3600
            TabIndex        =   36
            Top             =   2340
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   3360
            TabIndex        =   37
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   975
            TabIndex        =   28
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   3120
            TabIndex        =   27
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl发票号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发票号"
            Height          =   180
            Left            =   975
            TabIndex        =   26
            Top             =   3390
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   25
            Top             =   3390
            Width           =   180
         End
      End
      Begin VB.Frame fra范围 
         Height          =   4050
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "未审核退库"
            Height          =   180
            Left            =   720
            TabIndex        =   57
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "已审核退库"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2400
            TabIndex        =   56
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CheckBox chkAcc 
            Caption         =   "未财务审核"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   34
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CheckBox chk无发票 
            Caption         =   "无发票"
            Height          =   180
            Left            =   2400
            TabIndex        =   33
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CheckBox chk有发票 
            Caption         =   "有发票"
            Height          =   180
            Left            =   720
            TabIndex        =   32
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CheckBox chk未标记 
            Caption         =   "未做付款标记"
            Height          =   255
            Left            =   2400
            TabIndex        =   31
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CheckBox chk已标记 
            Caption         =   "已做付款标记"
            Height          =   255
            Left            =   720
            TabIndex        =   30
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CheckBox chkAccStrike 
            Caption         =   "已财务审核"
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   29
            Top             =   2640
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246284291
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246284291
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   246349827
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   23
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   22
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   21
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   20
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   19
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   18
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   14
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   13
      Top             =   435
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   6480
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private BlnAdvance As Boolean '是否展开
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '父窗体
Private mstrSelectTag As String     '当前选择的对象
Public lng药品id As Long
Private mstrMatch As String '匹配方式 0-双向匹配 1-从左向右单向匹配

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    str填制人 As String
    str审核人 As String
    lng生产商 As Long
    str产地 As String
    str发票号开始 As String
    str发票号结束 As String
    int填制审核一并查询 As Integer
    int未标记 As Integer
    int已标记 As Integer
    int有发票 As Integer
    int无发票 As Integer
    lng药品分类 As Long
    str剂型 As String
    date发票审核日期开始 As Date
    date发票审核日期结束 As Date
End Type

Private SQLCondition As Type_SQLCondition

Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO开始 As String, _
        ByRef strNO结束 As String, _
        ByRef date填制时间开始 As Date, _
        ByRef date填制时间结束 As Date, _
        ByRef date审核时间开始 As Date, _
        ByRef date审核时间结束 As Date, _
        ByRef lng药品 As Long, _
        ByRef str填制人 As String, _
        ByRef str审核人 As String, _
        ByRef lng生产商 As Long, _
        ByRef str产地 As String, _
        ByRef str发票号开始 As String, _
        ByRef str发票号结束 As String, _
        ByRef lng药品分类 As Long, _
        ByRef str剂型 As String, _
        ByRef date发票审核日期开始 As Date, _
        ByRef date发票审核日期结束 As Date, _
        ByRef intNo标记 As Integer, _
        ByRef intYes标记 As Integer, _
        ByRef intNo发票 As Integer, _
        ByRef intYes发票 As Integer, _
        Optional ByRef intTmp As Integer = 0) As String
    mstrFind = ""
    mstrSelectTag = ""
    Set mfrmMain = FrmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
    strNO开始 = SQLCondition.strNO开始
    strNO结束 = SQLCondition.strNO结束
    date填制时间开始 = SQLCondition.date填制时间开始
    date填制时间结束 = SQLCondition.date填制时间结束
    date审核时间开始 = SQLCondition.date审核时间开始
    date审核时间结束 = SQLCondition.date审核时间结束
    lng药品 = SQLCondition.lng药品
    str审核人 = SQLCondition.str审核人
    str填制人 = SQLCondition.str填制人
    lng生产商 = SQLCondition.lng生产商
    str产地 = SQLCondition.str产地
    str发票号开始 = SQLCondition.str发票号开始
    str发票号结束 = SQLCondition.str发票号结束
    lng药品分类 = SQLCondition.lng药品分类
    str剂型 = SQLCondition.str剂型
    date发票审核日期开始 = SQLCondition.date发票审核日期开始
    date发票审核日期结束 = SQLCondition.date发票审核日期结束
    intNo标记 = SQLCondition.int未标记
    intYes标记 = SQLCondition.int已标记
    intNo发票 = SQLCondition.int无发票
    intYes发票 = SQLCondition.int有发票
    intTmp = SQLCondition.int填制审核一并查询
End Function


Private Sub chkClass_Click()
    If chkClass.Value = 1 Then
        txtClass.Enabled = True
        cmdClass.Enabled = True
    Else
        txtClass.Enabled = False
        cmdClass.Enabled = False
    End If
End Sub

Private Sub chkJiXin_Click()
    If chkJiXin.Value = 1 Then
        txtJiXing.Enabled = True
        cmdJiXin.Enabled = True
    Else
        txtJiXing.Enabled = False
        cmdJiXin.Enabled = False
    End If
End Sub

Private Sub chkStrike_Click()
    chkAccStrike.Enabled = IIf(chkStrike.Value = 1, True, False)
End Sub
Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        cmd确定.SetFocus
    End If
    
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        cmd确定.SetFocus
    End If
End Sub

Private Sub chk发票日期_Click()
    If chk发票日期.Value = 1 Then
        dtpStart发票.Enabled = True
        dtpEnd发票.Enabled = True
    Else
        dtpStart发票.Enabled = False
        dtpEnd发票.Enabled = False
    End If
End Sub

Private Sub Chk供应商_Click()
    txt供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    Cmd供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    
End Sub

Private Sub Chk供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk供应商.Value = 1 Then
        txt供应商.SetFocus
    Else
        Chk生产商.SetFocus
    End If
End Sub


Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk审核.Value = 0 Then
            cmd确定.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub Chk生产商_Click()
    Me.txt生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
    Cmd生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
End Sub

Private Sub Chk生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        
        Else
            Txt填制人.SetFocus
        End If
    End If
End Sub

Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    chkNOVerifyBack.Enabled = IIf(chk填制.Value = 1, True, False)
    If chk填制.Value = 0 Then chkNOVerifyBack.Value = 0
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    chk已标记.Enabled = IIf(chk审核.Value = 1, True, False)
    chk未标记.Enabled = IIf(chk审核.Value = 1, True, False)
    chkAcc.Enabled = IIf(chk审核.Value = 1, True, False)
    chkYesVerifyBack.Enabled = IIf(chk审核.Value = 1, True, False)
    If chk审核.Value = 0 Then chkYesVerifyBack.Value = 0
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = IIf(Chk药品.Value = 1, True, False)
End Sub

Private Sub Chk药品_GotFocus()
    sstFilter.Tab = 1
    Chk药品.SetFocus
End Sub

Private Sub Chk药品_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk药品.Value = 1 Then
        Txt药品.SetFocus
    ElseIf Chk供应商.Visible = True Then
        Chk供应商.SetFocus
    End If
End Sub



Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng库房id As Long
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    tvw类别.Left = txtClass.Left
    tvw类别.Top = txtClass.Top + txtClass.Height
    tvw类别.Visible = True
    tvw类别.SetFocus
        
    gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
              "Where Instr([1], 编码, 1) > 0 " & _
              "Order by 编码 "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw类别
        .Nodes.Clear
'        Set nodTmp = .Nodes.Add(, , "Root", "所有", 2, 2)
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!名称, rsTmp!名称, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!编码
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    gstrSQL = "Select ID, 上级ID, 名称, 1 as 末级, decode(类型,1,'西成药',2,'中成药','中草药') as 材质, 类型 " & _
                  "From 诊疗分类目录 " & _
                  "Where 类型 in (1,2,3) " & _
                  "Start With 上级ID IS NULL Connect By Prior ID=上级ID Order by level,ID "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "提取药品用途分类")
    
    With rsTmp
        If .EOF Then
            Exit Sub
        End If
        
        '将药品用途分类数据装入
        Do While Not .EOF
            Int末级 = IIf(!末级 = 1, 3, 2)
            If IsNull(!上级ID) Then
                Set nodTmp = tvw类别.Nodes.Add("Root" & !材质, 4, "K_" & !id, !名称, Int末级, Int末级)
            Else
                Set nodTmp = tvw类别.Nodes.Add("K_" & !上级ID, 4, "K_" & !id, !名称, Int末级, Int末级)
            End If
            nodTmp.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With

    With tvw类别
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Int末级 = 1
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Int末级 = 2
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Int末级 = 3
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Int末级 = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdJiXin_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lng库房id As Long
    
    lvw剂型.Left = txtJiXing.Left
    lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
    lvw剂型.Visible = True
    lvw剂型.SetFocus
    
    On Error GoTo errHandle
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng库房id <> 0 Then
        '提取该库房现有剂型，供用户选择
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                  "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                  "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] " & _
                  "Order by J.名称 "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房id)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型 order by 名称 "
        Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型")
    End If
    
    With rsTmp
        lvw剂型.ListItems.Clear
        Do While Not .EOF
            lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd供应商_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt供应商.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "产地", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt供应商.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt供应商.SetFocus
    txt供应商.Tag = rsProvider!id
    txt供应商.Text = rsProvider!名称
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    Dim 未审核子条件 As String
    Dim 已审核子条件 As String
    
    '初始准备
    intNO = 21
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '检查数据
    If chkClass.Value = 1 Then
        If txtClass.Tag = "" Then
            MsgBox "请选择要查询的分类信息！", vbInformation, gstrSysName
            Me.txtClass.SetFocus
            Exit Sub
        End If
    End If
    If chkJiXin.Value = 1 Then
        If txtJiXing.Tag = "" Then
            MsgBox "请选择要查询的剂型信息！", vbInformation, gstrSysName
            Me.txtJiXing.SetFocus
            Exit Sub
        End If
    End If
    If Chk药品.Value = 1 Then
        If Txt药品.Tag = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.Txt药品.SetFocus
            Exit Sub
        End If
    End If
    If Chk供应商.Value = 1 Then
        If txt供应商.Tag = 0 Then
            MsgBox "请选择需查询的药品供应商信息！", vbInformation, gstrSysName
            Me.txt供应商.SetFocus
            Exit Sub
        End If
    End If
    If Chk生产商.Value = 1 Then
        If txt生产商.Tag = 0 Then
            MsgBox "请选择需查询的药品生产商信息！", vbInformation, gstrSysName
            Me.txt生产商.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '基本查询条件
    Dim i As Integer
    
    '未审核子条件
    If chkNOVerifyBack = 0 Then '不勾选只显示入库的；勾选就退库也显示
       未审核子条件 = " and nvl(a.发药方式,0)=0 "
    End If
    '已审核子条件
    If chkStrike.Value = 1 Then
        If chkAccStrike.Value = 0 And chkAcc.Value = 1 Then '未财务审核
            已审核子条件 = " And Nvl(A.费用ID,0)<>1 "
        ElseIf chkAccStrike.Value = 1 And chkAcc.Value = 0 Then  '已财务审核
            已审核子条件 = " And Nvl(A.费用ID,0)<>0  "
        End If
    Else
        If chkAcc.Value = 1 Then    '未财务审核
            已审核子条件 = " And Nvl(A.费用ID,0)=0 "
        End If
        已审核子条件 = 已审核子条件 & " And a.记录状态 =1 "
    End If
    If chk已标记.Value = 1 And chk未标记.Value = 0 Then
        已审核子条件 = 已审核子条件 & " And d.付款标志 =1"
        SQLCondition.int未标记 = 0
        SQLCondition.int已标记 = 1
    ElseIf chk未标记.Value = 1 And chk已标记.Value = 0 Then
        已审核子条件 = 已审核子条件 & " And d.付款标志 <>1"
        SQLCondition.int未标记 = 1
        SQLCondition.int已标记 = 0
    End If
    If chkYesVerifyBack.Value = 0 Then
        已审核子条件 = 已审核子条件 & " and nvl(a.发药方式,0)=0 "
    End If
    
    SQLCondition.int填制审核一并查询 = 0
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        SQLCondition.int填制审核一并查询 = 1

        mstrFind = "  and ((A.填制日期 between [3] and [4] and A.审核日期 is null " & 未审核子条件 & ") or (a.审核日期 between [5] and [6] " & 已审核子条件 & "))"
        
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        mstrFind = " And A.审核日期 Between [5] And [6] " & 已审核子条件
            
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [3] And [4]) and A.审核日期 is null " & 未审核子条件
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    If chk有发票.Value = 1 And chk无发票.Value = 0 Then
        mstrFind = mstrFind & " And d.发票号 is not null"
        SQLCondition.int有发票 = 1
        SQLCondition.int无发票 = 0
    ElseIf chk无发票.Value = 1 And chk有发票.Value = 0 Then
        mstrFind = mstrFind & " And d.发票号 is null"
        SQLCondition.int有发票 = 0
        SQLCondition.int无发票 = 1
    End If
        
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房id)
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房id)
    End If
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2]"
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
    
    SQLCondition.strNO开始 = Me.txt开始No
    SQLCondition.strNO结束 = Me.txt结束NO
    SQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    
    
    '扩展查询条件
    SQLCondition.lng药品分类 = 0
    SQLCondition.str剂型 = ""
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk药品.Value = 1 Then
        lng药品id = Txt药品.Tag
        mstrFind = mstrFind & " And A.药品ID + 0 =[7] "
    End If
    If Chk供应商.Value = 1 Then mstrFind = mstrFind & " And A.供药单位ID + 0 =[11] "
    If Chk生产商.Value = 1 Then mstrFind = mstrFind & " And A.产地=[12] "
    
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like [10] "
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like [9] "
    If Me.Txt开始发票号 <> "" And Me.Txt结束发票号 <> "" Then mstrFind = mstrFind & " And d.发票号 >= [13] And d.发票号 <=[14] "
    If Me.Txt开始发票号 <> "" And Me.Txt结束发票号 = "" Then mstrFind = mstrFind & " And d.发票号 >= [13] "
    If Me.Txt开始发票号 = "" And Me.Txt结束发票号 <> "" Then mstrFind = mstrFind & " And d.发票号 <= [14] "
        
    If chkClass.Value = 1 Then
        SQLCondition.lng药品分类 = Val(txtClass.Tag)
    End If
        
    If chkJiXin.Value = 1 Then
        SQLCondition.str剂型 = txtJiXing.Tag
    End If
    If chk发票日期.Value = 1 Then
        SQLCondition.date发票审核日期开始 = CDate(Format(dtpStart发票.Value, "yyyy-mm-dd") & " 00:00:00")
        SQLCondition.date发票审核日期结束 = CDate(Format(dtpEnd发票.Value, "yyyy-mm-dd") & " 23:59:59")
        mstrFind = mstrFind + " and d.审核日期 between [19] and [20]"
    End If
    
    SQLCondition.lng药品 = Val(Txt药品.Tag)
    SQLCondition.lng生产商 = txt供应商.Tag
    SQLCondition.str产地 = txt生产商
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
    SQLCondition.str发票号开始 = Me.Txt开始发票号
    SQLCondition.str发票号结束 = Me.Txt结束发票号
    
    Unload Me
End Sub

Private Sub Cmd生产商_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt生产商.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码 as id ,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码 "
'    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "药品生产商", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
    
    If blnCancel = True Then txt生产商.SetFocus: Exit Sub '打开选择器时，点Esc不做以下处理
    
    If rsProvider.State = 0 Then Exit Sub
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    txt生产商.SetFocus
    txt生产商.Tag = 1
    txt生产商.Text = rsProvider!名称

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "药品外购入库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)

'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品ID
        
    If Chk供应商.Visible = True Then
        Chk供应商.SetFocus
    End If
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp结束时间(Index).SetFocus
End Sub


Private Sub Form_Activate()
    SQLCondition.int未标记 = 0
    SQLCondition.int已标记 = 0
    SQLCondition.int无发票 = 0
    SQLCondition.int有发票 = 0
    
    If gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
        chk已标记.Visible = True
        chk未标记.Visible = True
        chk有发票.Top = chk已标记.Height + chk已标记.Top + 70
        chk无发票.Top = chk有发票.Top
    Else
        chk已标记.Visible = False
        chk未标记.Visible = False
        chk有发票.Top = chk已标记.Top
        chk无发票.Top = chk有发票.Top
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    
    Me.dtp结束时间(0) = Sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.txt供应商.Tag = 0
    Me.Txt药品.Tag = 0
    Me.txt生产商.Tag = 0
    lng药品id = 0
    
    sstFilter.Tab = 0
    BlnAdvance = False
    chk已标记.Enabled = False
    chk未标记.Enabled = False
    mstrMatch = IIf(zlDataBase.GetPara("输入匹配", , , 0) = "0", "%", "")
    
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    dtpStart发票.Value = DateAdd("m", -1, CDate(StrToday))
    dtpEnd发票.Value = CDate(StrToday)
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    gstrSQL = "Select id,上级ID,编码,简码,末级,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And 名称 is Not NULL " & _
              "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is NULL Connect by prior id=上级id"
    Set rsCompete = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-供应商", gstrNodeNo)
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "药品供应商信息不全，请在供药单位管理中设置药品供应商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null "
    Set rsCompete = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-药品生产商", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "药品生产商信息不全,请在字典管理中设置药品生产商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            
            Case "Booker"
                Txt填制人.SetFocus
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
            Case "Verify"
                Txt审核人.SetFocus
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
        End Select
        Cancel = True
    End If
    Call ReleaseSelectorRS
End Sub

Private Sub lvw剂型_DblClick()
    Dim i As Integer
    Dim strName As String
    
    With lvw剂型
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked = True Then
                strName = strName & .ListItems(i).Text & ","
            End If
        Next
        lvw剂型.Visible = False
        txtJiXing.Tag = strName
        txtJiXing.Text = strName
    End With
End Sub

Private Sub lvw剂型_LostFocus()
    lvw剂型.Visible = False
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    Txt开始发票号.SetFocus
                
            End Select
            .Visible = False
            
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
    
End Sub

Private Sub tvw类别_DblClick()
    With tvw类别
        If .SelectedItem.Text <> "" Then
            If .SelectedItem.Key Like "Root*" Then Exit Sub
            txtClass.Tag = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "_") + 1)
            txtClass.Text = .SelectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub tvw类别_LostFocus()
    tvw类别.Visible = False
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = 100
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng库房id As Long
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(Trim(txtClass.Text))
        If strTemp <> "" Then
            tvw类别.Left = txtClass.Left
            tvw类别.Top = txtClass.Top + txtClass.Height
            tvw类别.Visible = True
            tvw类别.SetFocus
            
            gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
                      "Where Instr([1], 编码, 1) > 0 " & _
                      "Order by 编码 "
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
            With tvw类别
                .Nodes.Clear
                Do While Not rsTmp.EOF
                    Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!名称, rsTmp!名称, 2, 2)
                    nodTmp.Tag = "Root" & rsTmp!编码
                    rsTmp.MoveNext
                Loop
                rsTmp.Close
            End With
            
            gstrSQL = "Select ID, 上级id, 名称, 1 As 末级, 材质, 类型" & _
                        " From (Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 材质, 类型" & _
                               " From 诊疗分类目录" & _
                               " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                     " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1])" & _
                               " Start With 上级id Is Null" & _
                               " Connect By Prior ID = 上级id" & _
                               " Union " & _
                               " Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 材质, 类型" & _
                               " From 诊疗分类目录" & _
                               " Where ID In (Select 上级id" & _
                                            " From 诊疗分类目录" & _
                                            " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                                  " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1])))" & _
                        " Start With 上级id Is Null" & _
                        " Connect By Prior ID = 上级id" & _
                        " Order By Level, ID"
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "查询品种", "%" & strTemp & mstrMatch)
            
            With rsTmp
                If .EOF Then
                    Exit Sub
                End If
                
                '将药品用途分类数据装入
                Do While Not .EOF
                    Int末级 = IIf(!末级 = 1, 3, 2)
                    If IsNull(!上级ID) Then
                        Set nodTmp = tvw类别.Nodes.Add("Root" & !材质, 4, "K_" & !id, !名称, Int末级, Int末级)
                    Else
                        Set nodTmp = tvw类别.Nodes.Add("K_" & !上级ID, 4, "K_" & !id, !名称, Int末级, Int末级)
                    End If
                    nodTmp.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
                    .MoveNext
                Loop
            End With
        
            With tvw类别
                .Nodes(1).Selected = True
                If .Nodes(1).Children <> 0 Then
                    Int末级 = 1
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(2).Children <> 0 Then
                    Int末级 = 2
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(3).Children <> 0 Then
                    Int末级 = 3
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                Else
                    Int末级 = 0
                    .Nodes(1).Selected = True
                    .SelectedItem.Selected = True
                End If
                If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        txtClass.Tag = 0
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtJiXing_GotFocus()
    txtJiXing.SelStart = 0
    txtJiXing.SelLength = 100
End Sub

Private Sub txtJiXing_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim lng库房id As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind = "" Then Exit Sub
        
        lvw剂型.Left = txtJiXing.Left
        lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
        lvw剂型.Visible = True
        lvw剂型.SetFocus
        
        On Error GoTo errHandle
        lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If lng库房id <> 0 Then
            '提取该库房现有剂型，供用户选择
            gstrSQL = "Select Distinct J.编码,J.名称 " & _
                      "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                      "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] and (j.编码 like [2] or j.名称 like [2] or j.简码 like [2]) " & _
                      "Order by J.名称 "
            Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房id, "%" & strFind & mstrMatch)
        Else
            gstrSQL = "Select 编码,名称 From 药品剂型 where 编码 like [1] or 名称 like [1] or 简码 like [1] order by 名称 "
            Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型", "%" & strFind & mstrMatch)
        End If
        
        With rsTmp
            lvw剂型.ListItems.Clear
            Do While Not .EOF
                lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
                .MoveNext
            Loop
        End With
    ElseIf KeyCode = vbKeyDelete Then
        txtJiXing.Tag = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt供应商_GotFocus()
'    Tvw.Visible = False
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        vRect = zlControl.GetControlRect(txt供应商.hWnd)

        gstrSQL = "Select id,编码,简码,名称 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (编码 like [1] or 简码 like [1] or 名称 like [1]) "
'        Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[药品供应商]", IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%", gstrNodeNo)
        Set RecTmp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%", , gstrNodeNo)
        
        If blnCancel Then txt供应商.SetFocus: Exit Sub
        
        If RecTmp.State = 0 Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            KeyCode = 0
            txt供应商.Tag = 0
            txt供应商.SelStart = 0
            txt供应商.SelLength = Len(txt供应商.Text)
            Exit Sub
        End If
        
        txt供应商 = RecTmp!名称
        txt供应商.Tag = RecTmp!id
        
    End If
    
    If Chk生产商.Value = 1 Then
        txt生产商.SetFocus
    Else
        Chk生产商.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = 21
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房id)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt结束发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.cmd确定.SetFocus
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = 21
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房id)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt开始发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt结束发票号.SetFocus
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Txt开始发票号.SetFocus
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            Txt开始发票号.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)
        
        With rstemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rstemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt审核人.Top - Txt审核人.Height - 50
                    .Width = Me.ScaleWidth - sstFilter.Left - fra附加条件.Left - Txt审核人.Left - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt审核人 = IIf(IsNull(!姓名), "", !姓名)
                Txt开始发票号.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt生产商 = "" Then Exit Sub
        If Trim(txt生产商) = "" Then Exit Sub
        txt生产商 = UCase(txt生产商)
        vRect = zlControl.GetControlRect(txt生产商.hWnd)
    
        Dim rstemp As New ADODB.Recordset

        gstrSQL = "Select 编码 as id,简码,名称 From 药品生产商 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) Order By 编码"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[药品生产商]", _
'                        IIf(gstrMatchMethod = "0", "%", "") & Me.txt生产商 & "%", _
'                        Me.txt生产商 & "%", gstrNodeNo)
        Set rstemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & Me.txt生产商 & "%", Me.txt生产商 & "%", gstrNodeNo)
        
        If blnCancel Then txt生产商.SetFocus: Exit Sub
        
        If rstemp.State = 0 Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            KeyCode = 0
            txt生产商.Tag = 0
            txt生产商.SelStart = 0
            txt生产商.SelLength = Len(txt生产商.Text)
            Exit Sub
        End If
        
        txt生产商 = IIf(IsNull(rstemp!名称), "", rstemp!名称)
        
        txt生产商.Tag = 1
        Txt填制人.SetFocus

    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As New ADODB.Recordset
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)
        
        With rstemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rstemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top + Txt填制人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt填制人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt填制人.Top - Txt填制人.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt填制人 = IIf(IsNull(!姓名), "", !姓名)
                Me.Txt审核人.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra附加条件.Left + Txt药品.Left
    sngTop = Me.Top + sstFilter.Top + fra附加条件.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "药品外购入库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品ID
    
    If Chk供应商.Visible = True Then
        If Chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

