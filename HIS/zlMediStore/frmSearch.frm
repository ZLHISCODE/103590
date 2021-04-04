VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1935
      Left            =   840
      TabIndex        =   74
      Top             =   5520
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
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
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   31
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6240
      TabIndex        =   32
      Top             =   960
      Width           =   1100
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   5175
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   34
         Top             =   600
         Width           =   5505
         Begin MSComctlLib.ListView lvw剂型 
            Height          =   2835
            Left            =   1200
            TabIndex        =   46
            Top             =   3960
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
            Left            =   -240
            TabIndex        =   38
            Top             =   3960
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
         Begin VB.ComboBox cbo编制方法 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   750
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.ComboBox Cbo计划类型 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.CheckBox chk编制方法 
            Caption         =   "编制方法"
            Height          =   300
            Left            =   600
            TabIndex        =   41
            Top             =   720
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CheckBox Chk计划类型 
            Caption         =   "计划类型"
            Height          =   300
            Left            =   600
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox txt复核人 
            Height          =   300
            Left            =   1515
            MaxLength       =   8
            TabIndex        =   73
            Top             =   3690
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.ComboBox Cbo库房 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   51
            Text            =   "Cbo库房"
            Top             =   1530
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1530
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CheckBox Chk移入库房 
            Caption         =   "移入库房"
            Height          =   300
            Left            =   600
            TabIndex        =   50
            Top             =   1530
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   44
            Top             =   750
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox Txt结束发票号 
            Height          =   300
            Left            =   3780
            TabIndex        =   71
            Top             =   3330
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Txt开始发票号 
            Height          =   300
            Left            =   1530
            TabIndex        =   69
            Top             =   3330
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   67
            Top             =   2940
            Width           =   1365
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   65
            Top             =   2940
            Width           =   1365
         End
         Begin VB.CheckBox chkClass 
            Caption         =   "药品分类"
            Height          =   300
            Left            =   600
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "药品剂型"
            Height          =   300
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   45
            Top             =   750
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商"
            Height          =   300
            Left            =   600
            TabIndex        =   52
            Top             =   1530
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   54
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Chk生产商 
            Caption         =   "生产商"
            Height          =   300
            Left            =   600
            TabIndex        =   55
            Top             =   1920
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txt生产商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   56
            Top             =   1920
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CommandButton Cmd生产商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   57
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            TabIndex        =   49
            Top             =   1140
            Width           =   255
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   48
            Top             =   1140
            Width           =   3255
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品"
            Height          =   300
            Left            =   600
            TabIndex        =   47
            Top             =   1140
            Width           =   990
         End
         Begin VB.CheckBox chk发票日期 
            Caption         =   "发票审核日期"
            Height          =   405
            Left            =   600
            TabIndex        =   60
            Top             =   2340
            Visible         =   0   'False
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpStart发票 
            Height          =   315
            Left            =   1650
            TabIndex        =   61
            Top             =   2340
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   141819907
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpEnd发票 
            Height          =   315
            Left            =   3600
            TabIndex        =   63
            Top             =   2340
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   129499139
            CurrentDate     =   36263
         End
         Begin VB.ComboBox Cbo类别 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1920
            Visible         =   0   'False
            Width           =   3550
         End
         Begin VB.CheckBox Chk类别 
            Caption         =   "类别"
            Height          =   300
            Left            =   600
            TabIndex        =   58
            Top             =   1920
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl复核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "复核人"
            Height          =   180
            Left            =   930
            TabIndex        =   72
            Top             =   3750
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   5
            Left            =   3240
            TabIndex        =   70
            Top             =   3390
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Lbl发票号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发票号"
            Height          =   180
            Left            =   975
            TabIndex        =   68
            Top             =   3390
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   3120
            TabIndex        =   66
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   975
            TabIndex        =   64
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   3360
            TabIndex        =   62
            Top             =   2400
            Visible         =   0   'False
            Width           =   180
         End
      End
      Begin VB.Frame fra范围 
         Height          =   4170
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkAccStrike 
            Caption         =   "已财务审核"
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   20
            Top             =   2760
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chk已标记 
            Caption         =   "已做付款标记"
            Height          =   255
            Left            =   720
            TabIndex        =   25
            Top             =   3113
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk未标记 
            Caption         =   "未做付款标记"
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   3113
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk有发票 
            Caption         =   "有发票"
            Height          =   255
            Left            =   720
            TabIndex        =   27
            Top             =   3487
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chk无发票 
            Caption         =   "无发票"
            Height          =   255
            Left            =   2400
            TabIndex        =   28
            Top             =   3487
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkAcc 
            Caption         =   "未财务审核"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   19
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "已审核退库"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2400
            TabIndex        =   30
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "未审核退库"
            Height          =   180
            Left            =   720
            TabIndex        =   29
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   17
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   420
            Left            =   480
            TabIndex        =   4
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   3
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   2
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "未审核冲销"
            Height          =   300
            Left            =   720
            TabIndex        =   10
            Top             =   1400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "已审核冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   16
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   9
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   312
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   312
            Index           =   1
            Left            =   3588
            TabIndex        =   15
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   2
            Left            =   3600
            TabIndex        =   24
            Top             =   2835
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   22
            Top             =   2835
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   143720451
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk复核 
            Caption         =   "已复核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   18
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chk已打印 
            Caption         =   "已打印单据"
            Height          =   255
            Left            =   720
            TabIndex        =   75
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chk未打印 
            Caption         =   "未打印单据"
            Height          =   255
            Left            =   2400
            TabIndex        =   76
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   7
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   5
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   2640
            TabIndex        =   8
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
            TabIndex        =   12
            Top             =   2028
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   14
            Top             =   2034
            Width           =   180
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   1
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   2
            Left            =   3360
            TabIndex        =   23
            Top             =   2895
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "复核日期"
            Height          =   180
            Index           =   2
            Left            =   900
            TabIndex        =   21
            Top             =   2895
            Visible         =   0   'False
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   6480
      Top             =   1560
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
            Picture         =   "frmSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long                 '单据类型
Private mfrmMain As Form                 '父窗体
Private mint冲销申请 As Integer          '0-不需要申请;1-需要申请
Private mblnAdvance As Boolean           '是否展开过附加条件项
Private mstrMatch As String              '匹配方式 0-双向匹配 1-从左向右单向匹配
Private mstrSelectTag As String          '当前选择的对象
Private mblnStock As Boolean             '当前操作员是否是药库人员，仅对领用单据有效
Private mint入出类型 As Integer
Private mblnCancel As Boolean            '点击取消

Private Const mint填制 As Integer = 0
Private Const mint审核 As Integer = 1
Private Const mint复核 As Integer = 2
Private Const mintNo As Integer = 3
Private Const mint发票日期 As Integer = 4
Private Const mint发票号 As Integer = 5

Private Const mintTab范围 As Integer = 0

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    date复核时间开始 As Date
    date复核时间结束 As Date
    lng药品 As Long
    lng库房 As Long
    str填制人 As String
    str审核人 As String
    str复核人 As String
    lng计划类型 As Long
    lng编制方法 As Long
    lng供应商 As Long
    lng生产商 As Long
    str产地 As String
    lng入出类别 As Long
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
    int未打印 As Integer
    int已打印 As Integer
End Type

Private SQLCondition As Type_SQLCondition

Private Type Type_TemporaryInquiries
    int未审核单据 As Integer 'int未处理单据 As Integer
    int未审核冲销 As Integer
    int已审核单据 As Integer 'int已处理单据 As Integer
    int已审核冲销 As Integer
    int已复核单据 As Integer
    int包含冲销 As Integer
    
    int未财务审核 As Integer
    int已财务审核 As Integer
    
    int未审核退库 As Integer
    int已审核退库 As Integer
End Type

Private TemporaryInquiries As Type_TemporaryInquiries   '临时查询保留。用于恢复上次设置的过滤条件。（主界面关闭后过滤条件将重置）

'通用过滤窗口用到的集合Key值常量
Private Const mstrNO开始Key As String = "NO开始"
Private Const mstrNO结束Key As String = "NO结束"
Private Const mstr填制时间开始Key As String = "填制时间开始"
Private Const mstr填制时间结束Key As String = "填制时间结束"
Private Const mstr审核时间开始Key As String = "审核时间开始"
Private Const mstr审核时间结束Key As String = "审核时间结束"
Private Const mstr复核时间开始Key As String = "复核时间开始"
Private Const mstr复核时间结束Key As String = "复核时间结束"
Private Const mstr药品IDKey As String = "药品ID"
Private Const mstr供应商Key As String = "供应商"
Private Const mstr填制人Key As String = "填制人"
Private Const mstr审核人Key As String = "审核人"
Private Const mstr复核人Key As String = "复核人"
Private Const mstr计划类型Key As String = "计划类型"
Private Const mstr编制方法Key As String = "编制方法"
Private Const mstr库房IDKey As String = "库房ID"
Private Const mstr未审核单据Key As String = "未审核单据" '未处理单据
Private Const mstr已审核单据Key As String = "已审核单据" '已处理单据
Private Const mstr未审核冲销Key As String = "未审核冲销"
Private Const mstr已审核冲销Key As String = "已审核冲销"
Private Const mstr已复核单据Key As String = "已复核单据"
Private Const mstr产地Key As String = "产地"             '生产商
Private Const mstr入出类别Key As String = "入出类别"
Private Const mstr包含冲销Key As String = "包含冲销"
Private Const mstr发票号开始Key As String = "发票号开始"
Private Const mstr发票号结束Key As String = "发票号结束"
Private Const mstr药品分类Key As String = "药品分类"
Private Const mstr剂型Key As String = "剂型"
Private Const mstr发票审核日期开始Key As String = "发票审核日期开始"
Private Const mstr发票审核日期结束Key As String = "发票审核日期结束"
Private Const mstr无标记Key As String = "无标记"
Private Const mstr有标记Key As String = "有标记"
Private Const mstr无发票Key As String = "无发票"
Private Const mstr有发票Key As String = "有发票"
Private Const mstr填制审核一并查询Key As String = "填制审核一并查询"
Private Const mstr未财务审核Key As String = "未财务审核"
Private Const mstr已财务审核Key As String = "已财务审核"
Private Const mstr未审核退库Key As String = "未审核退库"
Private Const mstr已审核退库Key As String = "已审核退库"
Private Const mstr未打印Key As String = "未打印"
Private Const mstr已打印Key As String = "已打印"

Public Property Get getKey_NO开始() As String
    getKey_NO开始 = mstrNO开始Key
End Property

Public Property Get getKey_NO结束() As String
    getKey_NO结束 = mstrNO结束Key
End Property

Public Property Get getKey_填制时间开始() As String
    getKey_填制时间开始 = mstr填制时间开始Key
End Property

Public Property Get getKey_填制时间结束() As String
    getKey_填制时间结束 = mstr填制时间结束Key
End Property

Public Property Get getKey_审核时间开始() As String
    getKey_审核时间开始 = mstr审核时间开始Key
End Property

Public Property Get getKey_审核时间结束() As String
    getKey_审核时间结束 = mstr审核时间结束Key
End Property

Public Property Get getKey_复核时间开始() As String
    getKey_复核时间开始 = mstr复核时间开始Key
End Property

Public Property Get getKey_复核时间结束() As String
    getKey_复核时间结束 = mstr复核时间结束Key
End Property

Public Property Get getKey_药品ID() As String
    getKey_药品ID = mstr药品IDKey
End Property

Public Property Get getKey_供应商() As String
    getKey_供应商 = mstr供应商Key
End Property

Public Property Get getKey_填制人() As String
    getKey_填制人 = mstr填制人Key
End Property

Public Property Get getKey_审核人() As String
    getKey_审核人 = mstr审核人Key
End Property

Public Property Get getKey_复核人() As String
    getKey_复核人 = mstr复核人Key
End Property

Public Property Get getKey_计划类型() As String
    getKey_计划类型 = mstr计划类型Key
End Property

Public Property Get getKey_编制方法() As String
    getKey_编制方法 = mstr编制方法Key
End Property

Public Property Get getKey_库房ID() As String
    getKey_库房ID = mstr库房IDKey
End Property

Public Property Get getKey_未审核单据() As String
    getKey_未审核单据 = mstr未审核单据Key
End Property

Public Property Get getKey_已审核单据() As String
    getKey_已审核单据 = mstr已审核单据Key
End Property

Public Property Get getKey_未审核冲销() As String
    getKey_未审核冲销 = mstr未审核冲销Key
End Property

Public Property Get getKey_已审核冲销() As String
    getKey_已审核冲销 = mstr已审核冲销Key
End Property

Public Property Get getKey_已复核单据() As String
    getKey_已复核单据 = mstr已复核单据Key
End Property

Public Property Get getKey_产地() As String
    getKey_产地 = mstr产地Key
End Property

Public Property Get getKey_入出类别() As String
    getKey_入出类别 = mstr入出类别Key
End Property

Public Property Get getKey_包含冲销() As String
    getKey_包含冲销 = mstr包含冲销Key
End Property

Public Property Get getKey_发票号开始() As String
    getKey_发票号开始 = mstr发票号开始Key
End Property

Public Property Get getKey_发票号结束() As String
    getKey_发票号结束 = mstr发票号结束Key
End Property

Public Property Get getKey_药品分类() As String
    getKey_药品分类 = mstr药品分类Key
End Property

Public Property Get getKey_剂型() As String
    getKey_剂型 = mstr剂型Key
End Property

Public Property Get getKey_发票审核日期开始() As String
    getKey_发票审核日期开始 = mstr发票审核日期开始Key
End Property

Public Property Get getKey_发票审核日期结束() As String
    getKey_发票审核日期结束 = mstr发票审核日期结束Key
End Property

Public Property Get getKey_无标记() As String
    getKey_无标记 = mstr无标记Key
End Property

Public Property Get getKey_有标记() As String
    getKey_有标记 = mstr有标记Key
End Property

Public Property Get getKey_无发票() As String
    getKey_无发票 = mstr无发票Key
End Property

Public Property Get getKey_有发票() As String
    getKey_有发票 = mstr有发票Key
End Property

Public Property Get getKey_填制审核一并查询() As String
    getKey_填制审核一并查询 = mstr填制审核一并查询Key
End Property

Public Property Get getKey_未财务审核() As String
    getKey_未财务审核 = mstr未财务审核Key
End Property

Public Property Get getKey_已财务审核() As String
    getKey_已财务审核 = mstr已财务审核Key
End Property

Public Property Get getKey_未审核退库() As String
    getKey_未审核退库 = mstr未审核退库Key
End Property

Public Property Get getKey_已审核退库() As String
    getKey_已审核退库 = mstr已审核退库Key
End Property

Public Property Get getKey_未打印() As String
    getKey_未打印 = mstr未打印Key
End Property

Public Property Get getKey_已打印() As String
    getKey_已打印 = mstr已打印Key
End Property

Public Property Get In_入出类型() As Integer
    In_入出类型 = mint入出类型
End Property

Public Property Let In_入出类型(ByVal vNewValue As Integer)
    mint入出类型 = vNewValue
End Property

Private Sub cbo库房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    '获取可操作的库房
    Select Case mlngMode
        Case 模块号.药品移库
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.药品领用
            str工作性质 = "O"
        Case 模块号.其他出库
            Exit Sub
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo库房.ListCount = 0 Then Exit Sub
    
    If cbo库房.ListIndex >= 0 Then
        If Val(cbo库房.Tag) = cbo库房.ItemData(cbo库房.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cbo库房, Trim(cbo库房.Text), str工作性质) = False Then
        Exit Sub
    End If
    If cbo库房.ListIndex >= 0 Then
        cbo库房.Tag = cbo库房.ItemData(cbo库房.ListIndex)
    End If
End Sub

Private Sub chk未打印_Click()
    If chk未打印.Value = 1 Then chk已打印.Value = 0
End Sub

Private Sub chk已打印_Click()
    If chk已打印.Value = 1 Then chk未打印.Value = 0
End Sub

Private Sub chkAcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkAccStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkClass_Click()
    If chkClass.Value = 1 Then
        txtClass.Enabled = True
        cmdClass.Enabled = True
    Else
        txtClass.Enabled = False
        cmdClass.Enabled = False
    End If
End Sub

Private Sub chkClass_GotFocus()
    If sstFilter.Tab = 0 Then sstFilter.Tab = 1
End Sub

Private Sub chkClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
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

Private Sub chkJiXin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkNoStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkNOVerifyBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkYesStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chkYesVerifyBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk编制方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
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

Private Sub chk发票日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk复核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chk复核.Value = 1 Then
            SendKeys vbTab
        Else
            cmd确定.SetFocus
        End If
    End If
End Sub

Private Sub Chk供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk计划类型_Click()
    Cbo计划类型.Enabled = IIf(Chk计划类型.Value = 1, True, False)
End Sub

Private Sub chk编制方法_Click()
    cbo编制方法.Enabled = IIf(chk编制方法.Value = 1, True, False)
End Sub

Private Sub Chk计划类型_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk计划类型.SetFocus
    End If
End Sub

Private Sub Chk计划类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk生产商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk填制_GotFocus()
    If sstFilter.Tab = 1 Then sstFilter.Tab = 0
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk未标记_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk无发票_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = Txt药品.Enabled
End Sub

Private Sub Chk供应商_Click()
    txt供应商.Enabled = IIf(chk供应商.Value = 1, True, False)
    cmd供应商.Enabled = txt供应商.Enabled
End Sub

Private Sub Chk生产商_Click()
    Me.txt生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
    Cmd生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
End Sub

Private Sub Chk类别_Click()
    Cbo类别.Enabled = IIf(Chk类别.Value = 1, True, False)
End Sub

Private Sub chkStrike_Click()
    chkAccStrike.Enabled = IIf(chkStrike.Value = 1, True, False)
End Sub

Private Sub chk复核_Click()
    DTP开始时间(mint复核).Enabled = IIf(chk复核.Value = 1, True, False)
    DTP结束时间(mint复核).Enabled = DTP开始时间(mint复核).Enabled
End Sub

Private Sub chk审核_Click()
    DTP开始时间(mint审核).Enabled = IIf(chk审核.Value = 1, True, False)
    DTP结束时间(mint审核).Enabled = IIf(chk审核.Value = 1, True, False)
    
    Select Case mlngMode
        Case 模块号.外购入库
            chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
            chk已标记.Enabled = IIf(chk审核.Value = 1, True, False)
            chk未标记.Enabled = IIf(chk审核.Value = 1, True, False)
            chkAcc.Enabled = IIf(chk审核.Value = 1, True, False)
            chkYesVerifyBack.Enabled = IIf(chk审核.Value = 1, True, False)
            If chk审核.Value = 0 Then chkYesVerifyBack.Value = 0
        Case 模块号.自制入库, 模块号.其他入库, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
            If 模块号.药品移库 = mlngMode Then chkYesStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    End Select
End Sub

Private Sub chk填制_Click()
    DTP开始时间(mint填制).Enabled = IIf(chk填制.Value = 1, True, False)
    DTP结束时间(mint填制).Enabled = IIf(chk填制.Value = 1, True, False)

    Select Case mlngMode
        Case 模块号.外购入库
            chkNOVerifyBack.Enabled = IIf(chk填制.Value = 1, True, False)
            If chk填制.Value = 0 Then chkNOVerifyBack.Value = 0
            chkNoStrike.Enabled = IIf(chk填制.Value = 1, True, False)
        Case 模块号.药品移库 ', 模块号.药品领用, 模块号.其他出库
            chkNoStrike.Enabled = IIf(chk填制.Value = 1, True, False)
    End Select
End Sub

Private Sub chk未标记_Click()
    If chk未标记.Value = 1 Then
        chk已标记.Value = 0
    End If
End Sub

Private Sub chk无发票_Click()
    If chk无发票.Value = 1 Then
        chk有发票.Value = 0
    End If
End Sub

Private Sub Chk药品_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk药品.SetFocus
    End If
End Sub

Private Sub Chk药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk移入库房_click()
    cbo库房.Enabled = IIf(Chk移入库房.Value = 1, True, False)
End Sub

Private Sub Chk移入库房_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk已标记_Click()
    If chk已标记.Value = 1 Then
        chk未标记.Value = 0
    End If
End Sub

Private Sub chk已标记_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk有发票_Click()
    If chk有发票.Value = 1 Then
        chk无发票.Value = 0
    End If
End Sub

Private Sub chk有发票_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    tvw类别.Left = txtClass.Left
    tvw类别.Top = txtClass.Top + txtClass.Height
    tvw类别.Visible = True
    tvw类别.SetFocus
        
    gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
              "Where Instr([1], 编码, 1) > 0 " & _
              "Order by 编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw类别
        .Nodes.Clear
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
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品用途分类")
    
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
    Dim lng库房ID As Long
    
    lvw剂型.Left = txtJiXing.Left
    lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
    lvw剂型.Visible = True
    lvw剂型.SetFocus
    
    On Error GoTo errHandle
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng库房ID <> 0 Then
        '提取该库房现有剂型，供用户选择
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                  "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                  "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] " & _
                  "Order by J.名称 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型 order by 名称 "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型")
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
              " Where (站点 = [1] Or 站点 is Null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              " Start with 上级ID is null and (站点 = [1] Or 站点 is Null) " & _
              " connect by prior ID =上级ID and (站点 = [1] Or 站点 is Null) order by level,ID"
    
    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "产地", True, "", "", False, False, _
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
    
    If mlngMode = 模块号.质量管理 Then
        Txt填制人.SetFocus
    ElseIf mlngMode = 模块号.外购入库 Then
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        Else
            Chk生产商.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim lng库房ID As Long
    Dim intNO As Integer
    
    '该过程各模块都有的检查
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        If mlngMode = 模块号.质量管理 Then
            MsgBox "对不起，必须选择一个登记日期或者处理日期!", vbInformation, gstrSysName
            chk填制.SetFocus
            Exit Sub
        ElseIf mlngMode = 模块号.药品计划 Then
            If chk复核.Value = 0 Then
                MsgBox "对不起，必须选择一个填制日期或者审核日期或者复核日期!", vbInformation, gstrSysName
                chk填制.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
            chk填制.SetFocus
            Exit Sub
        End If
    End If
    
    If mlngMode <> 模块号.质量管理 Then
        intNO = Switch(mlngMode = 模块号.外购入库, 21, mlngMode = 模块号.其他入库, 24, mlngMode = 模块号.自制入库, 22, _
                        mlngMode = 模块号.差价调整, 25, mlngMode = 模块号.药品移库, 26, mlngMode = 模块号.药品领用, 27, _
                        mlngMode = 模块号.其他出库, 28, mlngMode = 模块号.药品计划, 32)
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    End If
    
    '===================范围选项卡==================
    Select Case mlngMode
        Case 模块号.质量管理
            If Chk药品.Value = 1 Then
                If Txt药品.Tag = 0 Then
                    MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
                    Me.Txt药品.SetFocus
                    Exit Sub
                End If
            End If
            If chk供应商.Value = 1 Then
                If txt供应商.Tag = 0 Then
                    MsgBox "请选择需查询的药品供应商信息！", vbInformation, gstrSysName
                    Me.txt供应商.SetFocus
                    Exit Sub
                End If
            End If
        Case 模块号.药品计划
            If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
                txt开始No.Text = zlCommFun.GetFullNo(txt开始No.Text, intNO, lng库房ID)
            End If
            If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
                txt结束No.Text = zlCommFun.GetFullNo(txt结束No.Text, intNO, lng库房ID)
            End If
            
            SQLCondition.strNO开始 = Me.txt开始No
            SQLCondition.strNO结束 = Me.txt结束No
            SQLCondition.date复核时间开始 = CDate(Format(DTP开始时间(mint复核), "yyyy-mm-dd") & " 00:00:00")
            SQLCondition.date复核时间结束 = CDate(Format(DTP结束时间(mint复核), "yyyy-mm-dd") & " 23:59:59")
            TemporaryInquiries.int已复核单据 = chk复核.Value
            
        Case 模块号.其他入库, 模块号.自制入库
            '检查数据
            If Chk药品.Value = 1 Then
                If Txt药品.Tag = 0 Then
                    MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
                    Me.Txt药品.SetFocus
                    Exit Sub
                End If
            End If
            
            If mlngMode = 模块号.其他入库 Then
                If Chk生产商.Value = 1 Then
                    If txt生产商.Tag = 0 Then
                        MsgBox "请选择需查询的药品生产商信息！", vbInformation, gstrSysName
                        Me.txt生产商.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
                txt开始No.Text = zlCommFun.GetFullNo(txt开始No.Text, intNO, lng库房ID)
            End If
            If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
                txt结束No.Text = zlCommFun.GetFullNo(txt结束No.Text, intNO, lng库房ID)
            End If
            
            SQLCondition.strNO开始 = Me.txt开始No
            SQLCondition.strNO结束 = Me.txt结束No
            TemporaryInquiries.int包含冲销 = chkStrike.Value
            
        Case 模块号.外购入库
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
            If chk供应商.Value = 1 Then
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
            
            If chk已标记.Value = 1 And chk未标记.Value = 0 Then
                SQLCondition.int未标记 = 0
                SQLCondition.int已标记 = 1
            ElseIf chk未标记.Value = 1 And chk已标记.Value = 0 Then
                SQLCondition.int未标记 = 1
                SQLCondition.int已标记 = 0
            End If
            
            SQLCondition.int填制审核一并查询 = 0
            If chk填制.Value = 1 And chk审核.Value = 1 Then SQLCondition.int填制审核一并查询 = 1
            
            If chk有发票.Value = 1 And chk无发票.Value = 0 Then
                SQLCondition.int有发票 = 1
                SQLCondition.int无发票 = 0
            ElseIf chk无发票.Value = 1 And chk有发票.Value = 0 Then
                SQLCondition.int有发票 = 0
                SQLCondition.int无发票 = 1
            End If
                
            If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
                txt开始No.Text = zlCommFun.GetFullNo(txt开始No.Text, intNO, lng库房ID)
            End If
            If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
                txt结束No.Text = zlCommFun.GetFullNo(txt结束No.Text, intNO, lng库房ID)
            End If
            
            SQLCondition.strNO开始 = Me.txt开始No
            SQLCondition.strNO结束 = Me.txt结束No
            TemporaryInquiries.int包含冲销 = chkStrike.Value
            TemporaryInquiries.int未财务审核 = chkAcc.Value
            TemporaryInquiries.int已财务审核 = chkAccStrike.Value
            TemporaryInquiries.int未审核退库 = chkNOVerifyBack.Value
            TemporaryInquiries.int已审核退库 = chkYesVerifyBack.Value
            
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
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
            
            '基本查询条件
            SQLCondition.int填制审核一并查询 = 0
            If chk填制.Value = 1 And chk审核.Value = 1 Then SQLCondition.int填制审核一并查询 = 1
            
            If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
                txt开始No.Text = zlCommFun.GetFullNo(txt开始No.Text, intNO, lng库房ID)
            End If
            If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
                txt结束No.Text = zlCommFun.GetFullNo(txt结束No.Text, intNO, lng库房ID)
            End If
            
            SQLCondition.strNO开始 = Me.txt开始No
            SQLCondition.strNO结束 = Me.txt结束No
            TemporaryInquiries.int未审核冲销 = chkNoStrike.Value
            TemporaryInquiries.int已审核冲销 = chkYesStrike.Value
            TemporaryInquiries.int包含冲销 = chkStrike.Value
            
            If mlngMode = 模块号.药品移库 Then
                If chk已打印.Value = 1 And chk未打印.Value = 0 Then
                    SQLCondition.int未打印 = 0
                    SQLCondition.int已打印 = 1
                ElseIf chk未打印.Value = 1 And chk已打印.Value = 0 Then
                    SQLCondition.int未打印 = 1
                    SQLCondition.int已打印 = 0
                End If
            End If
    End Select
    
    '该过程范围选项卡各模块都有的语句
    SQLCondition.date填制时间开始 = CDate(Format(DTP开始时间(mint填制), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(DTP结束时间(mint填制), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(DTP开始时间(mint审核), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(DTP结束时间(mint审核), "yyyy-mm-dd") & " 23:59:59")
    TemporaryInquiries.int未审核单据 = chk填制.Value
    TemporaryInquiries.int已审核单据 = chk审核.Value
    
    '==================附加条件选项卡====================
    '扩展查询条件
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    Select Case mlngMode
        Case 模块号.质量管理
            SQLCondition.lng供应商 = IIf(chk供应商.Value = 1, Val(txt供应商.Tag), 0)
            
        Case 模块号.药品计划
            SQLCondition.str复核人 = IIf(Me.txt复核人 = "", "", Me.txt复核人 & "%")
            SQLCondition.lng计划类型 = IIf(Chk计划类型.Value = 1, Cbo计划类型.ListIndex + 1, 0)
            SQLCondition.lng编制方法 = IIf(chk编制方法.Value = 1, cbo编制方法.ListIndex + 1, 0)
            
        Case 模块号.其他入库
            SQLCondition.str产地 = IIf(Chk生产商.Value = 1, txt生产商, "")
            SQLCondition.lng入出类别 = IIf(Chk类别.Value = 1, Cbo类别.ItemData(Cbo类别.ListIndex), 0)
            
        Case 模块号.外购入库
            SQLCondition.lng药品分类 = 0
            SQLCondition.str剂型 = ""
            
            SQLCondition.lng药品分类 = IIf(chkClass.Value = 1, Val(txtClass.Tag), 0)
            SQLCondition.str剂型 = IIf(chkJiXin.Value = 1, txtJiXing.Tag, "")
            If chk发票日期.Value = 1 Then
                SQLCondition.date发票审核日期开始 = CDate(Format(dtpStart发票.Value, "yyyy-mm-dd") & " 00:00:00")
                SQLCondition.date发票审核日期结束 = CDate(Format(dtpEnd发票.Value, "yyyy-mm-dd") & " 23:59:59")
            End If
            
            SQLCondition.lng生产商 = IIf(chk供应商.Value = 1, Val(txt供应商.Tag), 0)
            SQLCondition.str产地 = IIf(Chk生产商.Value = 1, txt生产商, "")
            SQLCondition.str发票号开始 = Me.txt开始发票号
            SQLCondition.str发票号结束 = Me.txt结束发票号
            
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            SQLCondition.lng药品分类 = 0
            SQLCondition.str剂型 = ""
            
            SQLCondition.lng药品分类 = IIf(chkClass.Value = 1, Val(txtClass.Tag), 0)
            SQLCondition.str剂型 = IIf(chkJiXin.Value = 1, txtJiXing.Tag, "")
            If cbo库房.Visible Then SQLCondition.lng库房 = IIf(Chk移入库房.Value = 1, cbo库房.ItemData(cbo库房.ListIndex), 0)
            
    End Select
    '该过程附加条件选项卡各模块都有的语句
    SQLCondition.lng药品 = IIf(Chk药品.Value = 1, Val(Txt药品.Tag), 0)
    SQLCondition.str审核人 = IIf(Me.Txt审核人 = "", "", Me.Txt审核人 & "%")
    SQLCondition.str填制人 = IIf(Me.Txt填制人 = "", "", Me.Txt填制人 & "%")
    
    Unload Me
End Sub

Private Sub Cmd生产商_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt生产商.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码 as id ,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码 "
    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
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
    
    If mlngMode = 模块号.其他入库 Then
        If Chk类别.Visible = True Then
            If Chk类别.Value = 1 Then
                Cbo类别.SetFocus
            Else
                Chk类别.SetFocus
            End If
        End If
    Else '外购
        chk发票日期.SetFocus
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    Dim strModeName As String
    
    
    strModeName = Switch(mlngMode = 模块号.外购入库, "药品外购入库管理", mlngMode = 模块号.自制入库, "药品自制入库管理", mlngMode = _
        模块号.其他入库, "药品其他入库管理", mlngMode = 模块号.差价调整, "药品移库管理", mlngMode = 模块号.药品移库, "药品移库管理", mlngMode = _
        模块号.药品领用, "药品移库管理", mlngMode = 模块号.其他出库, "药品移库管理", mlngMode = 模块号.药品计划, "药品计划管理", mlngMode = _
        模块号.质量管理, "药品质量管理")
    
    Select Case mlngMode
        Case 模块号.外购入库, 模块号.自制入库, 模块号.其他入库, 模块号.药品计划
            Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库, 模块号.质量管理
            Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
    End Select
    
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.质量管理 Then
        If chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            chk供应商.SetFocus
        End If
    ElseIf mlngMode = 模块号.其他入库 Then
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        Else
            Chk生产商.SetFocus
        End If
    ElseIf mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
        If Chk移入库房.Value = 1 Then
            cbo库房.SetFocus
        Else
            Chk移入库房.SetFocus
        End If
    ElseIf mlngMode = 模块号.药品计划 Or mlngMode = 模块号.自制入库 Then
        Txt填制人.SetFocus
    End If
End Sub

Private Sub dtpEnd发票_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub dtpStart发票_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.dtpEnd发票.SetFocus
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.DTP结束时间(Index).SetFocus
End Sub

Private Sub Form_Activate()
    If mlngMode = 模块号.外购入库 Then
        SQLCondition.int未标记 = 0
        SQLCondition.int已标记 = 0
        SQLCondition.int无发票 = 0
        SQLCondition.int有发票 = 0
    ElseIf mlngMode = 模块号.药品移库 Then
        SQLCondition.int未打印 = 0
        SQLCondition.int已打印 = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    LoadSstFilter范围
    LoadSstFilter附加条件
    LoadData '初始化
    
End Sub

Private Sub LoadData()
    Dim StrToday As String
    '功能：加载数据
    
    '恢复上一次的设置
    '本过程各模块都存在的语句
    Me.DTP结束时间(mint填制) = SQLCondition.date填制时间结束
    Me.DTP结束时间(mint审核) = SQLCondition.date审核时间结束
    Me.DTP开始时间(mint填制) = SQLCondition.date填制时间开始
    Me.DTP开始时间(mint审核) = SQLCondition.date审核时间开始
    Me.chk填制.Value = TemporaryInquiries.int未审核单据
    Me.chk审核.Value = TemporaryInquiries.int已审核单据
    sstFilter.Tab = 0
    mblnAdvance = False
    
    Select Case mlngMode
        Case 模块号.质量管理
            Txt药品.Tag = 0
            txt供应商.Tag = 0
        Case 模块号.药品计划
            Me.DTP结束时间(mint复核) = SQLCondition.date复核时间结束
            Me.DTP开始时间(mint复核) = SQLCondition.date复核时间开始
            
            Me.chk复核.Value = TemporaryInquiries.int已复核单据
            
            SQLCondition.lng药品 = 0
        Case 模块号.其他入库, 模块号.自制入库
            Me.chkStrike.Value = TemporaryInquiries.int包含冲销
            
            Me.Txt药品.Tag = 0
            If mlngMode = 模块号.其他入库 Then Me.txt生产商.Tag = 0
            
        Case 模块号.外购入库
            chkStrike.Value = TemporaryInquiries.int包含冲销
            chkAcc.Value = TemporaryInquiries.int未财务审核
            chkAccStrike.Value = TemporaryInquiries.int已财务审核
            chk已标记.Value = SQLCondition.int已标记
            chk未标记.Value = SQLCondition.int未标记
            chk有发票.Value = SQLCondition.int有发票
            chk无发票.Value = SQLCondition.int无发票
            chkNOVerifyBack.Value = TemporaryInquiries.int未审核退库
            chkYesVerifyBack.Value = TemporaryInquiries.int已审核退库
            
            Me.txt供应商.Tag = 0
            Me.Txt药品.Tag = 0
            Me.txt生产商.Tag = 0
            
            chk已标记.Enabled = IIf(TemporaryInquiries.int已审核单据 = 1, True, False)
            chk未标记.Enabled = IIf(TemporaryInquiries.int已审核单据 = 1, True, False)
            mstrMatch = IIf(zlDatabase.GetPara("输入匹配", , , 0) = "0", "%", "")
            
            StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
            dtpStart发票.Value = DateAdd("m", -1, CDate(StrToday))
            dtpEnd发票.Value = CDate(StrToday)
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            chkNoStrike.Value = TemporaryInquiries.int未审核冲销
            chkYesStrike.Value = TemporaryInquiries.int已审核冲销
            chkStrike.Value = TemporaryInquiries.int包含冲销
            
            mblnStock = Check是否是药库人员
            mstrMatch = IIf(zlDatabase.GetPara("输入匹配", , , 0) = "0", "%", "")
            
            Me.Txt药品.Tag = 0
            If mlngMode = 模块号.药品移库 Then
                mint冲销申请 = Val(zlDatabase.GetPara("冲销申请", glngSys, 模块号.药品移库))
                If mint入出类型 = -1 Then
                    Chk移入库房.Caption = "移入库房"
                Else
                    Chk移入库房.Caption = "移出库房"
                End If
                If mint冲销申请 = 0 Then    '不需要申请
                    chkStrike.Visible = True
                    chkNoStrike.Visible = False
                    chkYesStrike.Visible = False
                Else
                    chkStrike.Visible = False
                    chkNoStrike.Visible = True
                    chkYesStrike.Visible = True
                End If
                
                chk已打印.Value = SQLCondition.int已打印
                chk未打印.Value = SQLCondition.int未打印
            ElseIf mlngMode = 模块号.药品领用 Then
                    Chk移入库房.Caption = "领用部门"
            ElseIf mlngMode = 模块号.其他出库 Then
                    Chk移入库房.Caption = "入出类别"
            End If
    End Select
End Sub

Private Sub LoadSstFilter范围()
    '功能：设置范围选项卡下面的控件显示、位置及大小等
    
    '默然窗体容器大小
    frmSearch.Height = 4510: sstFilter.Height = 3850: fra范围.Height = 2850: fra附加条件.Height = 2850
    Select Case mlngMode
        Case 模块号.外购入库
            '外购独有的显示
            chkAcc.Visible = True
            chkAccStrike.Visible = True
            If gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                chk已标记.Visible = True
                chk未标记.Visible = True
            Else
                chk有发票.Top = chk已标记.Top
                chk无发票.Top = chk有发票.Top
            End If
            chk有发票.Visible = True
            chk无发票.Visible = True
            chkNOVerifyBack.Visible = True
            chkYesVerifyBack.Visible = True
            '窗体高度5500、选项页高度4810、fra范围高度4050
            frmSearch.Height = 5800: sstFilter.Height = 5110: fra范围.Height = 4150: fra附加条件.Height = 4150
            '设置取消按钮Cancel
'            cmd取消.Cancel = False
            
        Case 模块号.药品移库
            '移库申请冲销可见
            mint冲销申请 = Val(zlDatabase.GetPara("冲销申请", glngSys, 模块号.药品移库))
            chkNoStrike.Visible = True
            chkYesStrike.Visible = True
            chkStrike.Visible = False
            chk未打印.Visible = True
            chk已打印.Visible = True
            
            frmSearch.Height = 4810: sstFilter.Height = 4150: fra范围.Height = 3150: fra附加条件.Height = 3150
         Case 模块号.药品计划
            '计划的复核
            chk复核.Visible = True
            lbl时间(mint复核).Visible = True
            DTP开始时间(mint复核).Visible = True
            lbl至(mint复核).Visible = True
            DTP结束时间(mint复核).Visible = True
            chkStrike.Visible = False
            '窗体高度5500、选项页高度4810、fra范围高度4050
            frmSearch.Height = 5150: sstFilter.Height = 4450: fra范围.Height = 3450: fra附加条件.Height = 3450
        Case 模块号.质量管理
            '质量管理过滤无No和冲销条件
            lblNO.Visible = False
            txt开始No.Visible = False
            lbl至(mintNo).Visible = False
            txt结束No.Visible = False
            chkStrike.Visible = False
            '改变Caption
            chk填制.Caption = "未处理单据"
            lbl时间(mint填制).Caption = "登记日期"
            chk审核.Caption = "已处理单据"
            lbl时间(mint审核).Caption = "处理日期"
            'No条件隐藏，改变显示控件的top
            chk填制.Top = chk填制.Top - 240
            lbl时间(mint填制).Top = lbl时间(mint填制).Top - 240
            DTP开始时间(mint填制).Top = DTP开始时间(mint填制).Top - 240
            lbl至(mint填制).Top = lbl至(mint填制).Top - 240
            DTP结束时间(mint填制).Top = DTP结束时间(mint填制).Top - 240
            chk审核.Top = chk审核.Top - 240
            lbl时间(mint审核).Top = lbl时间(mint审核).Top - 240
            DTP开始时间(mint审核).Top = DTP开始时间(mint审核).Top - 240
            lbl至(mint审核).Top = lbl至(mint审核).Top - 240
            DTP结束时间(mint审核).Top = DTP结束时间(mint审核).Top - 240
            '窗体高度5500、选项页高度4810、fra范围高度4050
            frmSearch.Height = 4250: sstFilter.Height = 3550: fra范围.Height = 2550: fra附加条件.Height = 2250
    End Select
End Sub

Private Sub LoadSstFilter附加条件()
    '功能：设置附加条件选项卡下面的控件显示、位置及大小等
    
    Select Case mlngMode
        Case 模块号.外购入库
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            chk供应商.Visible = True: txt供应商.Visible = True: cmd供应商.Visible = True
            Chk生产商.Visible = True: txt生产商.Visible = True: Cmd生产商.Visible = True
            chk发票日期.Visible = True: dtpStart发票.Visible = True: lbl至(mint发票日期).Visible = True: dtpEnd发票.Visible = True
            Lbl发票号.Visible = True: txt开始发票号.Visible = True: lbl至(mint发票号).Visible = True: txt结束发票号.Visible = True
        Case 模块号.自制入库
            Chk药品.Top = 480: Txt药品.Top = 480: Cmd药品.Top = 480
            Lbl填制人.Top = 1200: Txt填制人.Top = 1140: Lbl填制人.Left = 930:  Txt填制人.Left = 1650
            Lbl审核人.Top = 1800: Txt审核人.Top = 1740: Lbl审核人.Left = Lbl填制人.Left: Txt审核人.Left = Txt填制人.Left
         Case 模块号.其他入库
            Chk生产商.Visible = True: txt生产商.Visible = True: Cmd生产商.Visible = True
            Chk类别.Visible = True: Cbo类别.Visible = True
            Chk药品.Top = 360: Txt药品.Top = 360: Cmd药品.Top = 360
            Chk生产商.Top = 950: txt生产商.Top = 950: Cmd生产商.Top = 950
            Chk类别.Top = 1540: Cbo类别.Top = 1540
            Lbl填制人.Top = 2190: Lbl填制人.Left = Lbl填制人.Left - 100: Txt填制人.Top = 2150: Txt填制人.Left = Cbo类别.Left
            Lbl审核人.Top = 2190: Txt审核人.Top = 2150
        Case 模块号.药品移库
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk移入库房.Visible = True: cbo库房.Visible = True
            Lbl填制人.Top = 2000: Txt填制人.Top = 1940: Lbl填制人.Left = 930:  Txt填制人.Left = 1650
            Lbl审核人.Top = 2400: Txt审核人.Top = 2340: Lbl审核人.Left = Lbl填制人.Left: Txt审核人.Left = Txt填制人.Left
        Case 模块号.药品领用
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk移入库房.Visible = True: cbo库房.Visible = True
            Lbl填制人.Top = 2000: Txt填制人.Top = 1940: Lbl填制人.Left = 930:  Txt填制人.Left = 1650
            Lbl审核人.Top = 2400: Txt审核人.Top = 2340: Lbl审核人.Left = Lbl填制人.Left: Txt审核人.Left = Txt填制人.Left
            Chk移入库房.Caption = "领用部门"
        Case 模块号.其他出库
            chkClass.Visible = True: txtClass.Visible = True: cmdClass.Visible = True
            chkJiXin.Visible = True: txtJiXing.Visible = True: cmdJiXin.Visible = True
            Chk移入库房.Visible = True: cbo库房.Visible = True
            Lbl填制人.Top = 2000: Txt填制人.Top = 1940: Lbl填制人.Left = 930:  Txt填制人.Left = 1650
            Lbl审核人.Top = 2400: Txt审核人.Top = 2340: Lbl审核人.Left = Lbl填制人.Left: Txt审核人.Left = Txt填制人.Left
            Chk移入库房.Caption = "入出类别"
        Case 模块号.质量管理
            chk供应商.Visible = True: txt供应商.Visible = True: cmd供应商.Visible = True
            chk供应商.Caption = "供药单位": Lbl填制人.Caption = "登记人": Lbl审核人.Caption = "处理人"
            Chk药品.Top = 360: Txt药品.Top = 360: Cmd药品.Top = 360
            chk供应商.Top = 750: txt供应商.Top = 750: cmd供应商.Top = 750
            Lbl填制人.Top = 1400: Txt填制人.Top = 1340: Lbl填制人.Left = 1050:  Txt填制人.Left = 1650
            Lbl审核人.Top = 1800: Txt审核人.Top = 1740: Lbl审核人.Left = Lbl填制人.Left: Txt审核人.Left = Txt填制人.Left
        Case 模块号.药品计划
            Chk计划类型.Visible = True: Cbo计划类型.Visible = True
            chk编制方法.Visible = True: cbo编制方法.Visible = True
            lbl复核人.Visible = True: txt复核人.Visible = True
            Lbl填制人.Top = 1700: Txt填制人.Top = 1640: Lbl填制人.Left = 870:  Txt填制人.Left = 1650
            Lbl审核人.Top = 2100: Txt审核人.Top = 2040: Lbl审核人.Left = Lbl填制人.Left:   Txt审核人.Left = Txt填制人.Left
            lbl复核人.Top = 2500: txt复核人.Top = 2440: lbl复核人.Left = Lbl填制人.Left: txt复核人.Left = Txt填制人.Left
    End Select
End Sub

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByRef colParameter As Collection) As Boolean
    Dim lngloop As Long
    
    GetSearch = False
    mblnCancel = False
    mlngMode = lngMode
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.自制入库 Then mstrSelectTag = ""
    Set mfrmMain = FrmMain
    
    getParameterValue colParameter '记录集合传过来的参数值
    If mlngMode = 模块号.其他入库 Or mlngMode = 模块号.外购入库 Or mlngMode = 模块号.自制入库 Then If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    
    If mblnCancel = True Then Exit Function '点击取消不用将选择的条件记录到集合
    setParameterValue colParameter '将选择的条件记录到集合中传回去
    
    GetSearch = True
End Function

Private Sub setParameterValue(ByRef colParameter As Collection)
    '功能：将窗体选择的条件记录到集合中传回到主界面,且卸载相关的模块变量
    
    '本过程各模块都存在的语句
    
    CollectionModify colParameter, TemporaryInquiries.int未审核单据, frmSearch.getKey_未审核单据: TemporaryInquiries.int未审核单据 = 0
    CollectionModify colParameter, SQLCondition.date填制时间开始, frmSearch.getKey_填制时间开始: SQLCondition.date填制时间开始 = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.date填制时间结束, frmSearch.getKey_填制时间结束: SQLCondition.date填制时间结束 = CDate("00:00:00")
    CollectionModify colParameter, TemporaryInquiries.int已审核单据, frmSearch.getKey_已审核单据: TemporaryInquiries.int已审核单据 = 0
    CollectionModify colParameter, SQLCondition.date审核时间开始, frmSearch.getKey_审核时间开始: SQLCondition.date审核时间开始 = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.date审核时间结束, frmSearch.getKey_审核时间结束: SQLCondition.date审核时间结束 = CDate("00:00:00")
    CollectionModify colParameter, SQLCondition.lng药品, frmSearch.getKey_药品ID: SQLCondition.lng药品 = 0
    CollectionModify colParameter, SQLCondition.str填制人, frmSearch.getKey_填制人: SQLCondition.str填制人 = ""
    CollectionModify colParameter, SQLCondition.str审核人, frmSearch.getKey_审核人: SQLCondition.str审核人 = ""
    
    Select Case mlngMode
        Case 模块号.质量管理
            CollectionModify colParameter, SQLCondition.lng供应商, frmSearch.getKey_供应商: SQLCondition.lng供应商 = 0
            
        Case 模块号.药品计划
            CollectionModify colParameter, SQLCondition.strNO开始, frmSearch.getKey_NO开始: SQLCondition.strNO开始 = ""
            CollectionModify colParameter, SQLCondition.strNO结束, frmSearch.getKey_NO结束: SQLCondition.strNO结束 = ""
            CollectionModify colParameter, TemporaryInquiries.int已复核单据, frmSearch.getKey_已复核单据: TemporaryInquiries.int已复核单据 = 0
            CollectionModify colParameter, SQLCondition.date复核时间开始, frmSearch.getKey_复核时间开始: SQLCondition.date复核时间开始 = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.date复核时间结束, frmSearch.getKey_复核时间结束: SQLCondition.date复核时间结束 = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.lng计划类型, frmSearch.getKey_计划类型: SQLCondition.lng计划类型 = 0
            CollectionModify colParameter, SQLCondition.lng编制方法, frmSearch.getKey_编制方法: SQLCondition.lng编制方法 = 0
            CollectionModify colParameter, SQLCondition.str复核人, frmSearch.getKey_复核人: SQLCondition.str复核人 = ""
            
        Case 模块号.其他入库, 模块号.自制入库
            CollectionModify colParameter, SQLCondition.strNO开始, frmSearch.getKey_NO开始: SQLCondition.strNO开始 = ""
            CollectionModify colParameter, SQLCondition.strNO结束, frmSearch.getKey_NO结束: SQLCondition.strNO结束 = ""
            CollectionModify colParameter, TemporaryInquiries.int包含冲销, frmSearch.getKey_包含冲销: TemporaryInquiries.int包含冲销 = 0
            If mlngMode = 模块号.其他入库 Then
                CollectionModify colParameter, SQLCondition.str产地, frmSearch.getKey_产地: SQLCondition.str产地 = ""
                CollectionModify colParameter, SQLCondition.lng入出类别, frmSearch.getKey_入出类别: SQLCondition.lng入出类别 = 0
            End If
            
        Case 模块号.外购入库
            CollectionModify colParameter, SQLCondition.strNO开始, frmSearch.getKey_NO开始: SQLCondition.strNO开始 = ""
            CollectionModify colParameter, SQLCondition.strNO结束, frmSearch.getKey_NO结束: SQLCondition.strNO结束 = ""
            CollectionModify colParameter, TemporaryInquiries.int包含冲销, frmSearch.getKey_包含冲销: TemporaryInquiries.int包含冲销 = 0
            CollectionModify colParameter, TemporaryInquiries.int未财务审核, frmSearch.getKey_未财务审核: TemporaryInquiries.int未财务审核 = 0
            CollectionModify colParameter, TemporaryInquiries.int已财务审核, frmSearch.getKey_已财务审核: TemporaryInquiries.int已财务审核 = 0
            CollectionModify colParameter, SQLCondition.int未标记, frmSearch.getKey_无标记: SQLCondition.int未标记 = 0
            CollectionModify colParameter, SQLCondition.int已标记, frmSearch.getKey_有标记: SQLCondition.int已标记 = 0
            CollectionModify colParameter, SQLCondition.int无发票, frmSearch.getKey_无发票: SQLCondition.int无发票 = 0
            CollectionModify colParameter, SQLCondition.int有发票, frmSearch.getKey_有发票: SQLCondition.int有发票 = 0
            CollectionModify colParameter, TemporaryInquiries.int未审核退库, frmSearch.getKey_未审核退库: TemporaryInquiries.int未审核退库 = 0
            CollectionModify colParameter, TemporaryInquiries.int已审核退库, frmSearch.getKey_已审核退库: TemporaryInquiries.int已审核退库 = 0
            CollectionModify colParameter, SQLCondition.int填制审核一并查询, frmSearch.getKey_填制审核一并查询: SQLCondition.int填制审核一并查询 = 0
            CollectionModify colParameter, SQLCondition.lng药品分类, frmSearch.getKey_药品分类: SQLCondition.lng药品分类 = 0
            CollectionModify colParameter, SQLCondition.str剂型, frmSearch.getKey_剂型: SQLCondition.str剂型 = ""
            CollectionModify colParameter, SQLCondition.lng生产商, frmSearch.getKey_供应商: SQLCondition.lng生产商 = 0
            CollectionModify colParameter, SQLCondition.str产地, frmSearch.getKey_产地: SQLCondition.str产地 = ""
            CollectionModify colParameter, SQLCondition.date发票审核日期开始, frmSearch.getKey_发票审核日期开始: SQLCondition.date发票审核日期开始 = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.date发票审核日期结束, frmSearch.getKey_发票审核日期结束: SQLCondition.date发票审核日期结束 = CDate("00:00:00")
            CollectionModify colParameter, SQLCondition.str发票号开始, frmSearch.getKey_发票号开始: SQLCondition.str发票号开始 = ""
            CollectionModify colParameter, SQLCondition.str发票号结束, frmSearch.getKey_发票号结束: SQLCondition.str发票号结束 = ""
            
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            CollectionModify colParameter, SQLCondition.strNO开始, frmSearch.getKey_NO开始: SQLCondition.strNO开始 = ""
            CollectionModify colParameter, SQLCondition.strNO结束, frmSearch.getKey_NO结束: SQLCondition.strNO结束 = ""
            CollectionModify colParameter, TemporaryInquiries.int未审核冲销, frmSearch.getKey_未审核冲销: TemporaryInquiries.int未审核冲销 = 0
            CollectionModify colParameter, TemporaryInquiries.int已审核冲销, frmSearch.getKey_已审核冲销: TemporaryInquiries.int已审核冲销 = 0
            CollectionModify colParameter, TemporaryInquiries.int包含冲销, frmSearch.getKey_包含冲销: TemporaryInquiries.int包含冲销 = 0
            CollectionModify colParameter, SQLCondition.int填制审核一并查询, frmSearch.getKey_填制审核一并查询: SQLCondition.int填制审核一并查询 = 0
            CollectionModify colParameter, SQLCondition.lng药品分类, frmSearch.getKey_药品分类: SQLCondition.lng药品分类 = 0
            CollectionModify colParameter, SQLCondition.str剂型, frmSearch.getKey_剂型: SQLCondition.str剂型 = ""
            CollectionModify colParameter, SQLCondition.lng库房, frmSearch.getKey_库房ID: SQLCondition.lng库房 = 0
            
            If mlngMode = 模块号.药品移库 Then
                CollectionModify colParameter, SQLCondition.int未打印, frmSearch.getKey_未打印: SQLCondition.int未打印 = 0
                CollectionModify colParameter, SQLCondition.int已打印, frmSearch.getKey_已打印: SQLCondition.int已打印 = 0
            End If
    End Select
End Sub

Private Sub CollectionModify(ByRef colParameter As Collection, ByVal varConditionn As Variant, ByVal strConditionnKey As String)
    '功能：集合修改指定key值的value
    colParameter.Remove strConditionnKey
    colParameter.Add varConditionn, strConditionnKey
End Sub

Private Sub getParameterValue(ByVal colParameter As Collection)
    '功能：将主窗体传过来的参数赋值给该窗体对应的变量用于数据初始化
    
    '临时查询初始化
    '本过程各模块都存在的语句
    
    SQLCondition.date填制时间开始 = colParameter(frmSearch.getKey_填制时间开始)
    SQLCondition.date填制时间结束 = colParameter(frmSearch.getKey_填制时间结束)
    SQLCondition.date审核时间开始 = colParameter(frmSearch.getKey_审核时间开始)
    SQLCondition.date审核时间结束 = colParameter(frmSearch.getKey_审核时间结束)
    TemporaryInquiries.int未审核单据 = colParameter(frmSearch.getKey_未审核单据)
    TemporaryInquiries.int已审核单据 = colParameter(frmSearch.getKey_已审核单据)
    
    Select Case mlngMode
        Case 模块号.药品计划
            SQLCondition.date复核时间开始 = colParameter(frmSearch.getKey_复核时间开始)
            SQLCondition.date复核时间结束 = colParameter(frmSearch.getKey_复核时间结束)
            TemporaryInquiries.int已复核单据 = colParameter(frmSearch.getKey_已复核单据)
            
        Case 模块号.其他入库, 模块号.自制入库
            TemporaryInquiries.int包含冲销 = colParameter(frmSearch.getKey_包含冲销)
            
        Case 模块号.外购入库
            TemporaryInquiries.int包含冲销 = colParameter(frmSearch.getKey_包含冲销)
            TemporaryInquiries.int未财务审核 = colParameter(frmSearch.getKey_未财务审核)
            TemporaryInquiries.int已财务审核 = colParameter(frmSearch.getKey_已财务审核)
            SQLCondition.int已标记 = colParameter(frmSearch.getKey_有标记)
            SQLCondition.int未标记 = colParameter(frmSearch.getKey_无标记)
            SQLCondition.int无发票 = colParameter(frmSearch.getKey_无发票)
            SQLCondition.int有发票 = colParameter(frmSearch.getKey_有发票)
            TemporaryInquiries.int未审核退库 = colParameter(frmSearch.getKey_未审核退库)
            TemporaryInquiries.int已审核退库 = colParameter(frmSearch.getKey_已审核退库)
            
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            TemporaryInquiries.int未审核冲销 = colParameter(frmSearch.getKey_未审核冲销)
            TemporaryInquiries.int已审核冲销 = colParameter(frmSearch.getKey_已审核冲销)
            TemporaryInquiries.int包含冲销 = colParameter(frmSearch.getKey_包含冲销)
            If mlngMode = 模块号.药品移库 Then
                SQLCondition.int已打印 = colParameter(frmSearch.getKey_已打印)
                SQLCondition.int未打印 = colParameter(frmSearch.getKey_未打印)
            End If
    End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If tvw类别.Visible = True Then
        tvw类别.Visible = False
        txtClass.SetFocus
        Cancel = True
        Exit Sub
    End If
    If lvw剂型.Visible = True Then
        lvw剂型.Visible = False
        txtJiXing.SetFocus
        Cancel = True
        Exit Sub
    End If
        
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Maker"
                txt生产商.SetFocus
                txt生产商.SelStart = 0
                txt生产商.SelLength = Len(txt生产商.Text)
            Case "Booker"
                Txt填制人.SetFocus
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
            Case "Verify"
                Txt审核人.SetFocus
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
            Case "Checker"
                txt复核人.SetFocus
                txt复核人.SelStart = 0
                txt复核人.SelLength = Len(txt复核人.Text)
        End Select
        Cancel = True
        Exit Sub
    End If
    
    If Not mfrmMain Is Nothing Then
        Set mfrmMain = Nothing
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
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
        If Chk药品.Value = 1 Then
            Txt药品.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
End Sub

Private Sub lvw剂型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvw剂型_DblClick
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
                Case "Maker"
                    txt生产商.Text = .TextMatrix(.Row, 1)
                    txt生产商.Tag = 1
                    Chk类别.SetFocus
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    cmd确定.SetFocus
                    If mlngMode = 模块号.药品计划 Then txt复核人.SetFocus
                    If mlngMode = 模块号.外购入库 Then txt开始发票号.SetFocus
                Case "Checker"
                    txt复核人 = .TextMatrix(.Row, 2)
                    cmd确定.SetFocus
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
    Dim rsDepartment As New Recordset
    Dim strStock As String
    Dim str站点限制 As String
    
    On Error GoTo errHandle
    str站点限制 = GetDeptStationNode(mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
            If cbo库房.ListCount < 1 Then
                Select Case mlngMode
                    Case 模块号.药品计划
                        If Cbo计划类型.ListCount < 1 Then
                            With Cbo计划类型
                                .Clear
                                .AddItem "月度计划", 0
                                .AddItem "季度计划", 1
                                .AddItem "年度计划", 2
                                .AddItem "周计划", 3
                                .ListIndex = 0
                            End With
                            
                            With cbo编制方法
                                .Clear
                                .AddItem "往年同期线形参照法", 0
                                .AddItem "临近期间平均参照法", 1
                                .AddItem "药品储备定额参照法", 2
                                .AddItem "药品日销售量参照法", 3
                                .ListIndex = 0
                            End With
                        End If
                    
                    Case 模块号.药品移库
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                                & "Where " & IIf(str站点限制 <> "", " (a.站点 = [3] or a.站点 is null) AND ", "") & " c.工作性质 = b.名称 " _
                                & "  AND Instr([1],b.编码,1) > 0 " _
                                & "  AND a.id = c.部门id " _
                                & "  AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 模块号.药品领用
                        strStock = "O"
                        gstrSQL = " Select C.ID " & _
                            " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                            " Where " & IIf(str站点限制 <> "", " (c.站点 = [3] or c.站点 is null) AND ", "") & " A.工作性质=B.名称 And A.部门ID=C.ID " & _
                            "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                            "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])"
                        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                            & "Where " & IIf(str站点限制 <> "", " (a.站点 = [3] or a.站点 is null) AND ", "") & " c.工作性质 = b.名称 " _
                            & "  AND Instr([1],b.编码,1) > 0 " _
                            & "  AND a.id = c.部门id " _
                            & "  AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')" _
                            & IIf(mblnStock, "", " And a.ID IN (Select Distinct 领用部门ID From 药品领用控制 Where 领用部门ID IN (" & gstrSQL & "))")
                    Case 模块号.其他出库
                       gstrSQL = "SELECT b.Id,b.名称 " _
                               & "FROM 药品单据性质 A, 药品入出类别 B " _
                               & "Where A.类别id = B.ID AND A.单据 = 11 "
                    Case 模块号.差价调整, 模块号.药品盘点
                        If Chk移入库房.Visible = True Then
                            Chk移入库房.Visible = False
                            cbo库房.Visible = False
                        End If
                        Exit Sub
                End Select
                
                If mlngMode = 模块号.差价调整 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Or mlngMode = 模块号.药品盘点 Then
                    Set rsDepartment = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock, UserInfo.用户ID, gstrNodeNo)
                
                    With cbo库房
                        Do While Not rsDepartment.EOF
                            .AddItem rsDepartment.Fields(1)
                            .ItemData(.NewIndex) = rsDepartment.Fields(0)
                            rsDepartment.MoveNext
                        Loop
                        If .ListCount > 0 Then .ListIndex = 0
                    End With
                    rsDepartment.Close
                End If
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
        If chkJiXin.Value = 1 Then
            txtJiXing.SetFocus
        Else
            chkJiXin.SetFocus
        End If
    End If
End Sub

Private Sub tvw类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then tvw类别_DblClick
End Sub

Private Sub tvw类别_LostFocus()
    tvw类别.Visible = False
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
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
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
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
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & strTemp & mstrMatch)
            
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
        Else
            If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
                If chkJiXin.Value = 1 Then
                    txtJiXing.SetFocus
                Else
                    chkJiXin.SetFocus
                End If
            End If
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

Private Sub txtJiXing_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim lng库房ID As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind <> "" Then
            lvw剂型.Left = txtJiXing.Left
            lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
            lvw剂型.Visible = True
            lvw剂型.SetFocus
            
            On Error GoTo errHandle
            lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            If lng库房ID <> 0 Then
                '提取该库房现有剂型，供用户选择
                gstrSQL = "Select Distinct J.编码,J.名称 " & _
                          "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                          "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] and (j.编码 like [2] or j.名称 like [2] or j.简码 like [2]) " & _
                          "Order by J.名称 "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID, "%" & strFind & mstrMatch)
            Else
                gstrSQL = "Select 编码,名称 From 药品剂型 where 编码 like [1] or 名称 like [1] or 简码 like [1] order by 名称 "
                Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型", "%" & strFind & mstrMatch)
            End If
            
            With rsTmp
                If .RecordCount = 0 Then
                    lvw剂型.Visible = False
                    MsgBox "输入值无效！", vbInformation, gstrSysName
                     txtJiXing.SetFocus: Exit Sub
                End If
                lvw剂型.ListItems.Clear
                Do While Not .EOF
                    lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
                    .MoveNext
                Loop
            End With
        Else
            If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
                If Chk药品.Value = 1 Then
                    Txt药品.SetFocus
                Else
                    Chk药品.SetFocus
                End If
            End If
        End If
        
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

Private Sub txt复核人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(txt复核人.Text) = "" Then
            If mlngMode = 模块号.外购入库 Then
                txt开始发票号.SetFocus
            Else
                Me.cmd确定.SetFocus
            End If
            Exit Sub
        End If
        txt复核人.Text = UCase(txt复核人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.txt复核人 & "%", _
                        Me.txt复核人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                txt复核人.SelStart = 0
                txt复核人.SelLength = Len(txt复核人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Checker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Left = sstFilter.Left + fra附加条件.Left + txt复核人.Left
                    .Height = txt复核人.Top - sstFilter.Top - fra附加条件.Top - 50
                    .Top = sstFilter.Top + fra附加条件.Top + txt复核人.Top - .Height
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
                txt复核人 = IIf(IsNull(!姓名), "", !姓名)
                If mlngMode = 模块号.外购入库 Then
                    txt开始发票号.SetFocus
                Else
                    Me.cmd确定.SetFocus
                End If
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

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        vRect = zlControl.GetControlRect(txt供应商.hWnd)
        
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [3] Or 站点 is Null) " & _
                  "  And 末级=1 And substr(类型,1,1)=1 " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] Or zlSpellCode(名称) Like [2] Or zlWbCode(名称) Like [2])" & _
                  "  Start with 上级ID is null and (站点 = [3] Or 站点 is Null) connect by prior ID =上级ID and (站点 = [3] Or 站点 is Null) "
        Set RecTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%", txt供应商 & "%", gstrNodeNo)
        
        
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
    
    If mlngMode = 模块号.质量管理 Then
        Txt填制人.SetFocus
    ElseIf mlngMode = 模块号.外购入库 Then
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        Else
            Chk生产商.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt结束发票号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmd确定.SetFocus
End Sub

Private Sub txt开始No_GotFocus()
    If sstFilter.Tab = 1 Then sstFilter.Tab = 0
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer
    
    '初始准备
    intNO = Switch(mlngMode = 模块号.外购入库, 21, mlngMode = 模块号.自制入库, 22, mlngMode = _
        模块号.其他入库, 24, mlngMode = 模块号.差价调整, 25, mlngMode = 模块号.药品移库, 26, mlngMode = _
        模块号.药品领用, 27, mlngMode = 模块号.其他出库, 28, mlngMode = 模块号.药品计划, 32)
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNo(txt开始No.Text, intNO, lng库房ID)
        End If
        txt结束No.SetFocus
    End If
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer
    
    '初始准备
    intNO = Switch(mlngMode = 模块号.外购入库, 21, mlngMode = 模块号.自制入库, 22, mlngMode = _
        模块号.其他入库, 24, mlngMode = 模块号.差价调整, 25, mlngMode = 模块号.药品移库, 26, mlngMode = _
        模块号.药品领用, 27, mlngMode = 模块号.其他出库, 28, mlngMode = 模块号.药品计划, 32)
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt结束No) < 8 And Len(txt结束No) > 0 Then
            txt结束No.Text = zlCommFun.GetFullNo(txt结束No.Text, intNO, lng库房ID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = 100
End Sub

Private Sub txtJiXing_GotFocus()
    txtJiXing.SelStart = 0
    txtJiXing.SelLength = 100
End Sub

Private Sub Txt开始发票号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            If mlngMode = 模块号.外购入库 Then
                txt开始发票号.SetFocus
            ElseIf mlngMode = 模块号.药品计划 Then
                Me.txt复核人.SetFocus
            Else
                Me.cmd确定.SetFocus
            End If
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt审核人.Top - Txt审核人.Height - 50
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
                If mlngMode = 模块号.外购入库 Then
                    txt开始发票号.SetFocus
                ElseIf mlngMode = 模块号.药品计划 Then
                    Me.txt复核人.SetFocus
                Else
                    Me.cmd确定.SetFocus
            End If
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



Private Sub txt生产商_GotFocus()
    txt生产商.SelStart = 0
    txt供应商.SelLength = Len(txt供应商.Text)
End Sub

Private Sub txt生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt生产商.hWnd)
    
    On Error GoTo errHandle
    
    If KeyCode = vbKeyReturn Then
        If Trim(txt生产商) <> "" Then
            txt生产商 = UCase(txt生产商)
            
            gstrSQL = "Select 编码 as id,简码,名称 From 药品生产商 " & _
                      "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) " & _
              "Order By 编码"
    
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品生产商", False, "", "", False, False, _
                    True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & Me.txt生产商 & "%", Me.txt生产商 & "%", gstrNodeNo)
            
            If blnCancel Then txt生产商.SetFocus: Exit Sub
            
            If rsTemp.State = 0 Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                txt生产商.Tag = 0
                txt生产商.SelStart = 0
                txt生产商.SelLength = Len(txt生产商.Text)
                Exit Sub
            End If
            
            txt生产商 = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
            txt生产商.Tag = 1
        End If
        
        If mlngMode = 模块号.其他入库 Then
            If Chk类别.Visible = True Then
                If Chk类别.Value = 1 Then
                    Cbo类别.SetFocus
                Else
                    Chk类别.SetFocus
                End If
            End If
        Else '外购
            chk发票日期.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
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
    Dim strModeName As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) <> "" Then
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
        
        strModeName = Switch(mlngMode = 模块号.外购入库, "药品外购入库管理", mlngMode = 模块号.自制入库, "药品自制入库管理", mlngMode = _
            模块号.其他入库, "药品其他入库管理", mlngMode = 模块号.差价调整, "药品移库管理", mlngMode = 模块号.药品移库, "药品移库管理", mlngMode = _
            模块号.药品领用, "药品移库管理", mlngMode = 模块号.其他出库, "药品移库管理", mlngMode = 模块号.药品计划, "药品计划管理", mlngMode = _
            模块号.质量管理, "药品质量管理")
        
        Select Case mlngMode
            Case 模块号.外购入库, 模块号.自制入库, 模块号.其他入库, 模块号.药品计划
                Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
            Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库, 模块号.质量管理
                Call SetSelectorRS(1, strModeName, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
        End Select
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then Exit Sub
        If gint药品名称显示 = 1 Then
            Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        Txt药品.Tag = RecReturn!药品id
    End If
    
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.质量管理 Then
        If chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            chk供应商.SetFocus
        End If
    ElseIf mlngMode = 模块号.其他入库 Then
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        Else
            Chk生产商.SetFocus
        End If
    ElseIf mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Then
        If Chk移入库房.Value = 1 Then
            cbo库房.SetFocus
        Else
            Chk移入库房.SetFocus
        End If
    ElseIf mlngMode = 模块号.药品计划 Or mlngMode = 模块号.自制入库 Then
        Txt填制人.SetFocus
    End If
    
End Sub

Private Sub Txt药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Cbo库房_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0

    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Cbo库房_Validate(Cancel As Boolean)
    If cbo库房.ListCount > 0 Then
        If cbo库房.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Txt填制人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Txt药品_GotFocus()
    Txt药品.SelStart = 0
    Txt药品.SelLength = Len(Txt药品.Text)
End Sub

Private Sub txt供应商_GotFocus()
    txt供应商.SelStart = 0
    txt供应商.SelLength = Len(txt供应商.Text)
End Sub

Private Sub Txt填制人_GotFocus()
    Txt填制人.SelStart = 0
    Txt填制人.SelLength = Len(Txt填制人.Text)
End Sub

Private Sub Txt审核人_GotFocus()
    Txt审核人.SelStart = 0
    Txt审核人.SelLength = Len(Txt审核人.Text)
End Sub

Private Sub txt复核人_GotFocus()
    txt复核人.SelStart = 0
    txt复核人.SelLength = Len(txt复核人.Text)
End Sub

Private Sub Cbo计划类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cbo编制方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Function Check是否是药库人员() As Boolean
    Dim rsDepend As ADODB.Recordset
    
    On Error GoTo errHandle
    '判断是不是药库人员使用本模块
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [2] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr('HIJKLMN', b.编码, 1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " _
            & "  And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1]) "
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.用户ID, gstrNodeNo)
                  
    Check是否是药库人员 = (rsDepend.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    
    If mlngMode = 模块号.外购入库 Then
        gstrSQL = "Select id,上级ID,编码,简码,末级,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And 名称 is Not NULL " & _
              "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is NULL Connect by prior id=上级id"
        Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-供应商", gstrNodeNo)
        With rsCompete
            If .EOF Then
                .Close
                MsgBox "药品供应商信息不全，请在供药单位管理中设置药品供应商信息！", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null"
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "药品生产商", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "药品生产商信息不全,请在字典管理中设置药品生产商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If mlngMode = 模块号.其他入库 Then
        gstrSQL = "SELECT B.Id,b.名称 " _
                & "FROM 药品单据性质 A, 药品入出类别 B " _
                & "Where A.类别id = B.ID AND A.单据 = 4 "
        Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsCompete
            If .EOF Then
                MsgBox "药品其他入库没有设置相应的入出类别，请检查药品入出分类！", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveFirst
            Do While Not .EOF
                Cbo类别.AddItem .Fields(1)
                Cbo类别.ItemData(Cbo类别.NewIndex) = .Fields(0)
                .MoveNext
            Loop
            Cbo类别.ListIndex = 0
            .Close
        End With
    End If
    
    CheckCompete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

