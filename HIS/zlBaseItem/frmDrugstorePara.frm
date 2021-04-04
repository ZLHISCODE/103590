VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugstorePara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药店运行参数"
   ClientHeight    =   5265
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7950
   Icon            =   "frmDrugstorePara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   3
      Left            =   210
      TabIndex        =   56
      Top             =   510
      Width           =   7425
      Begin VB.CommandButton cmdOperate 
         Caption         =   "增加(&A)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   0
         Left            =   6240
         TabIndex        =   63
         Top             =   510
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "修改(&M)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   1
         Left            =   6240
         TabIndex        =   62
         Top             =   990
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "删除(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   2
         Left            =   6240
         TabIndex        =   61
         Top             =   1470
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "清除(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   3
         Left            =   6240
         TabIndex        =   57
         Top             =   1950
         Width           =   1100
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   6600
         Top             =   2880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugstorePara.frx":000C
               Key             =   "Limit"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   3465
         Index           =   1
         Left            =   300
         TabIndex        =   64
         Top             =   480
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "操作人"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "单据类型"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "历史天数"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "允许修改他人单据"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "金额上限"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "按操作员对不同单据的操作权限，针对单据的历史天数和最初操作人进行限制"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   65
         Top             =   180
         Width           =   6120
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   4
      Left            =   150
      TabIndex        =   58
      Top             =   480
      Width           =   7500
      Begin ZL9BillEdit.BillEdit bill 
         Height          =   3585
         Index           =   0
         Left            =   210
         TabIndex        =   60
         Top             =   330
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   6324
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   3
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "控制药品在不同库房间的流通方向"
         Height          =   180
         Index           =   23
         Left            =   240
         TabIndex        =   59
         Top             =   60
         Width           =   2700
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   5
      Left            =   150
      TabIndex        =   67
      Top             =   405
      Width           =   7530
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf库房单位 
         Height          =   3900
         Left            =   180
         TabIndex        =   68
         Top             =   105
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   6879
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483631
         AllowBigSelection=   0   'False
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "药品库房|售价单位|门诊单位|住院单位|药库单位"
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   1
      Left            =   210
      TabIndex        =   34
      Top             =   510
      Width           =   7425
      Begin VB.Frame fra 
         Caption         =   "密文显示"
         Height          =   1095
         Index           =   11
         Left            =   4020
         TabIndex        =   39
         Top             =   180
         Width           =   2595
         Begin VB.CheckBox chk 
            Caption         =   "会员卡号码密文显示"
            Height          =   285
            Index           =   14
            Left            =   420
            TabIndex        =   40
            ToolTipText     =   "表示各个输入就诊卡号码处是否为密文显示"
            Top             =   450
            Width           =   1920
         End
      End
      Begin VB.Frame fra 
         Caption         =   "收据行次"
         Height          =   1095
         Index           =   10
         Left            =   300
         TabIndex        =   35
         Top             =   150
         Width           =   2685
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "3"
            Top             =   480
            Width           =   435
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   2
            Left            =   2010
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   3
            BuddyControl    =   "txtUD(2)"
            BuddyDispid     =   196618
            BuddyIndex      =   2
            OrigLeft        =   1965
            OrigTop         =   390
            OrigRight       =   2205
            OrigBottom      =   690
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "收费收据行次"
            Height          =   180
            Index           =   18
            Left            =   420
            TabIndex        =   36
            Top             =   540
            Width           =   1080
         End
      End
      Begin VB.Frame fra 
         Height          =   75
         Index           =   9
         Left            =   1230
         TabIndex        =   42
         Top             =   1620
         Width           =   5415
      End
      Begin VB.CheckBox chk 
         Caption         =   "票号严格控制"
         Height          =   285
         Index           =   13
         Left            =   5295
         TabIndex        =   47
         ToolTipText     =   "表示各个输入就诊卡号码处是否为密文显示"
         Top             =   3045
         Width           =   1380
      End
      Begin VB.TextBox txtUD 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "7"
         Top             =   3060
         Width           =   390
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   300
         Index           =   4
         Left            =   4395
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3060
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   7
         BuddyControl    =   "txtUD(4)"
         BuddyDispid     =   196618
         BuddyIndex      =   4
         OrigLeft        =   3795
         OrigTop         =   3630
         OrigRight       =   4035
         OrigBottom      =   3915
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1125
         Index           =   0
         Left            =   300
         TabIndex        =   43
         Top             =   1800
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1984
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "票据类型"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "号码长度"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "票号严格控制"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "号码长度"
         Height          =   180
         Index           =   19
         Left            =   3195
         TabIndex        =   44
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据汇总"
         Height          =   180
         Index           =   9
         Left            =   300
         TabIndex        =   41
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   510
      Width           =   7425
      Begin VB.Frame fra会员卡 
         Caption         =   "会员卡价格"
         Height          =   1155
         Left            =   3135
         TabIndex        =   27
         Top             =   2640
         Width           =   4275
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Height          =   255
            Left            =   2175
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chk变价 
            Caption         =   "价格可以变动"
            Height          =   285
            Left            =   2820
            TabIndex        =   30
            Top             =   330
            Width           =   1395
         End
         Begin VB.TextBox txt价格 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1200
            TabIndex        =   29
            Top             =   300
            Width           =   1245
         End
         Begin VB.TextBox txt收入项目 
            Height          =   300
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label lbl收入项目 
            AutoSize        =   -1  'True
            Caption         =   "所属收入项目"
            Height          =   180
            Left            =   90
            TabIndex        =   31
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label lbl价格 
            AutoSize        =   -1  'True
            Caption         =   "当前价格"
            Height          =   180
            Left            =   420
            TabIndex        =   28
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "杂项"
         Height          =   3750
         Index           =   4
         Left            =   210
         TabIndex        =   2
         Top             =   60
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "时价药品以加成率入库"
            Height          =   285
            Index           =   21
            Left            =   120
            TabIndex        =   66
            Top             =   3360
            Width           =   2160
         End
         Begin VB.CheckBox chk 
            Caption         =   "收费完成后是否自动发药"
            Height          =   285
            Index           =   17
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   2370
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "未配药处方发药"
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   8
            Top             =   2670
            Width           =   1680
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1260
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "指定药店时限定药品的库存"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   2355
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "指导批发价定价单位"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "分币处理"
            Height          =   180
            Index           =   15
            Left            =   120
            TabIndex        =   5
            Top             =   1020
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "药品出库"
         Height          =   1215
         Index           =   8
         Left            =   4950
         TabIndex        =   14
         Top             =   75
         Width           =   2475
         Begin VB.OptionButton opt 
            Caption         =   "检查，不足禁止"
            Height          =   195
            Index           =   6
            Left            =   375
            TabIndex        =   17
            Top             =   915
            Width           =   1560
         End
         Begin VB.OptionButton opt 
            Caption         =   "检查，不足提醒"
            Height          =   195
            Index           =   5
            Left            =   375
            TabIndex        =   16
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton opt 
            Caption         =   "不进行库存检查"
            Height          =   195
            Index           =   4
            Left            =   375
            TabIndex        =   15
            Top             =   315
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.Frame fra 
         Caption         =   "收费时会员输入"
         Height          =   1215
         Index           =   6
         Left            =   3135
         TabIndex        =   10
         Top             =   75
         Width           =   1650
         Begin VB.CheckBox chk 
            Caption         =   "姓名"
            Height          =   210
            Index           =   7
            Left            =   315
            TabIndex        =   11
            Top             =   330
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "会员ID"
            Height          =   225
            Index           =   8
            Left            =   315
            TabIndex        =   12
            Top             =   615
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox chk 
            Caption         =   "刷会员卡"
            Height          =   210
            Index           =   9
            Left            =   315
            TabIndex        =   13
            Top             =   930
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VB.Frame fra 
         Caption         =   "对外上下班时间"
         Height          =   1125
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   1410
         Width           =   4305
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   825
            TabIndex        =   20
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.3541666667
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   22
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.5
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   2
            Left            =   825
            TabIndex        =   24
            Top             =   675
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.5625
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   26
            Top             =   675
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.75
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "上午"
            Height          =   180
            Index           =   2
            Left            =   330
            TabIndex        =   19
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   1785
            TabIndex        =   21
            Top             =   345
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "下午"
            Height          =   180
            Index           =   3
            Left            =   330
            TabIndex        =   23
            Top             =   735
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   5
            Left            =   1785
            TabIndex        =   25
            Top             =   750
            Width           =   180
         End
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   2
      Left            =   210
      TabIndex        =   48
      Top             =   510
      Width           =   7425
      Begin VB.ListBox lst 
         Height          =   3420
         Index           =   1
         Left            =   2430
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   390
         Width           =   1935
      End
      Begin VB.ListBox lst 
         Height          =   3420
         Index           =   0
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "公费病人适用费用类型"
         Height          =   180
         Index           =   21
         Left            =   2430
         TabIndex        =   51
         Top             =   150
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医保病人适用费用类型"
         Height          =   180
         Index           =   20
         Left            =   270
         TabIndex        =   49
         Top             =   150
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   300
      TabIndex        =   55
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6510
      TabIndex        =   54
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5190
      TabIndex        =   53
      Top             =   4785
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4530
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   7990
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      TabMinWidth     =   2117
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "票据管理"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "权限"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "单据操作"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品流向"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品库房单位"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmDrugstorePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum const数
    ud_历史天数 = 0
    ud_收费收据 = 2
    ud_号码长度 = 4
End Enum

Private Enum constChk
    chk_限定药品的库存 = 2
    chk_病人姓名 = 7
    chk_病人ID = 8
    chk_刷就诊卡 = 9
    chk_票号控制 = 13
    chk_密文显示 = 14
    chk_未配药处方发药 = 16
    chk_收费同时发药 = 17
    chk_时价药品入库 = 21
End Enum

Private Enum const日期
    dtp_上午上班 = 0
    dtp_上午下班 = 1
    dtp_下午上班 = 2
    dtp_下午下班 = 3
End Enum

Private Enum constCmb
    cmb_分币处理 = 1
    cmb_定价单位 = 3
End Enum

Private Enum constBill
    bill_药品流向 = 0
End Enum

Private Enum constLvw
    lvw_票据 = 0
    lvw_单据 = 1
End Enum

Private Enum constListBox
    lst_医保病人 = 0
    lst_公费病人 = 1
End Enum

Private Enum constOpt
    opt_不进行库存检查 = 4
    opt_不足提醒 = 5
    opt_不足禁止 = 6
End Enum

'变量声明
Dim mblnChange As Boolean     '是否改变了
Dim mblnInit As Boolean       '是否初始化失败
Dim mblnLoad As Boolean
Dim mintColumn As Integer '

'用于会员卡价格设置而特别增加的变量
Dim mlng会员卡ID  As Long
Dim mstr会员卡编码  As String
Dim mlng价目ID As Long

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim str姓名 As String, str人员ID As String, str单据 As String
    Dim lng单据 As Long, lng天数 As Long, bln修改他人 As Boolean
    Dim dbl金额上限 As Double
    Dim lst As ListItem
    
    
    Select Case Index
        Case 0 '新增
            If frmBillPrivilege.编辑权限(str姓名, str人员ID, str单据, lng单据, lng天数, bln修改他人, dbl金额上限, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_单据).ListItems
                If lst.Tag = str人员ID And lst.ListSubItems(1).Tag = lng单据 Then
                    MsgBox "本次新增的操作限制已经存在。", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        Case 1 '修改
            If lvw(lvw_单据).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_单据).SelectedItem
                str姓名 = .Text
                str单据 = .SubItems(1)
                lng天数 = Val(.SubItems(2))
                bln修改他人 = (.SubItems(3) = "是")
                dbl金额上限 = Val(.SubItems(4))
                str人员ID = .Tag
                lng单据 = .ListSubItems(1).Tag
            End With
            If frmBillPrivilege.编辑权限(str姓名, str人员ID, str单据, lng单据, lng天数, bln修改他人, dbl金额上限, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_单据).ListItems
                If Not lst Is lvw(lvw_单据).SelectedItem Then
                    If lst.Tag = str人员ID And lst.ListSubItems(1).Tag = lng单据 Then
                        MsgBox "本次改变的操作限制已经存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 '删除
            If lvw(lvw_单据).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_单据).SelectedItem
                If MsgBox("你确实要删除“" & .Text & "”对“" & .SubItems(1) & "”的操作限制？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                lvw(lvw_单据).ListItems.Remove .Index
            End With
        Case 3 '清除
            If MsgBox("你确实要删除所有的操作限制？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            lvw(lvw_单据).ListItems.Clear
    End Select
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_单据).ListItems.Add(, , str姓名, , "Limit")
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_单据).SelectedItem
            lst.Text = str姓名
        End If
        lst.SubItems(1) = str单据
        lst.SubItems(2) = lng天数
        lst.SubItems(3) = IIF(bln修改他人 = True, "是", "否")
        lst.SubItems(4) = Format(dbl金额上限, "0.00")
        lst.Tag = str人员ID
        lst.ListSubItems(1).Tag = lng单据
    End If
    mblnChange = True
End Sub

Private Sub cmdSelect_Click()
'选择收入项目
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim strTemp As String
    Dim strID As String
    Dim lngRow As Long
    
    strTemp = txt收入项目.Text
    strID = txt收入项目.Tag
    strSQL = "select ID,上级ID,名称,末级  from 收入项目 where " & Where撤档时间() & _
        "  start with 上级ID is null  connect by prior ID =上级ID"
    blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, strTemp, "收入项目")
    If blnRe Then
        On Error Resume Next
        txt收入项目.Tag = strID
        txt收入项目.Text = strTemp
        mblnChange = True
    End If
End Sub

Private Sub Form_Activate()
    If mblnLoad = False Then Exit Sub
    '以下部分只运行一次
    mblnLoad = False
    If mblnInit = False Then Unload Me
    Call tabMain_Click
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    mblnLoad = True
    '进行初始化
    Call InitEnv
    Call LoadPara
    Call Load会员卡
    Call Load药品流向
    Call Load药品库房单位
    
    RestoreFlexState Bill(bill_药品流向), App.ProductName & "\" & Me.Name & bill_药品流向
    '初始化成功
    mblnChange = False
    mblnInit = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitEnv()
    '初始化窗口，这是的是不需要读数据库的
    cmb(cmb_分币处理).AddItem "0-不处理"
    cmb(cmb_分币处理).AddItem "1-四舍五入"
    cmb(cmb_分币处理).AddItem "2-补整收取"
    cmb(cmb_分币处理).AddItem "3-舍分收取"
    cmb(cmb_分币处理).ListIndex = 0
    
    cmb(cmb_定价单位).AddItem "0-售价单位"
    cmb(cmb_定价单位).AddItem "1-采购单位"
    cmb(cmb_定价单位).ListIndex = 0
    
    lvw(lvw_票据).ListItems.Add , "C1", "收费收据"
    lvw(lvw_票据).ListItems.Add , "C5", "会员卡"
    
    With Bill(bill_药品流向)
        .Cols = 4 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "所在库房"
        .TextMatrix(0, 1) = "对方库房"
        .TextMatrix(0, 2) = "对方库房ID"
        .TextMatrix(0, 3) = "流向"
        .ColWidth(0) = 1700
        .ColWidth(1) = 1700
        .ColWidth(2) = 0
        .ColWidth(3) = 3600
        .ColData(0) = 3
        .ColData(1) = 3
        .ColData(2) = 5
        .ColData(3) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    '库房单位
    msf库房单位.AllowUserResizing = flexResizeNone
    msf库房单位.Cols = 3
    msf库房单位.FormatString = "药品库房|售价单位|药库单位"
    msf库房单位.ColWidth(1) = 900
    msf库房单位.ColWidth(2) = 900
    msf库房单位.ColAlignment(1) = 4
    msf库房单位.ColAlignment(2) = 4
    msf库房单位.ColWidth(0) = msf库房单位.Width - 900 * 2 - 27 * Screen.TwipsPerPixelX
End Sub

Private Sub LoadPara()
'系统参数表
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    '首先对费用类型进行初始化
    Call Load费用类型
    
    On Error GoTo ErrHandle
    gstrSQL = "select 参数号,参数值 from Zlparameters Where 系统 = " & glngSys & " And Nvl(私有, 0) = 0 And 模块 Is Null Order By 参数号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    Do Until rsTemp.EOF
        Select Case rsTemp("参数号")
            Case 1 '上午上下班时间
                i = InStr(UCase(rsTemp("参数值")), "AND")
                strTemp = Mid(rsTemp("参数值"), 1, i - 2)
                dtp(dtp_上午上班).Value = CDate(strTemp)
                strTemp = Mid(rsTemp("参数值"), i + 4)
                dtp(dtp_上午下班).Value = CDate(strTemp)
            Case 2 '下午上下班时间
                i = InStr(UCase(rsTemp("参数值")), "AND")
                strTemp = Mid(rsTemp("参数值"), 1, i - 2)
                dtp(dtp_下午上班).Value = CDate(strTemp)
                strTemp = Mid(rsTemp("参数值"), i + 4)
                dtp(dtp_下午下班).Value = CDate(strTemp)
            Case 4 '收费收据总行次
                If Not IsNull(rsTemp("参数值")) Then
                    ud(ud_收费收据).Value = rsTemp("参数值")
                End If
            Case 8 '未配药处方发药
                chk(chk_未配药处方发药) = IIF(rsTemp("参数值") <> 0, 1, 0)
            Case 9 '药品出库库存检查
                '该组第一个控件的Index值是4
                opt(CInt(IIF(IsNull(rsTemp("参数值")), "0", rsTemp("参数值"))) + 4).Value = True
            Case 12 '就诊卡号密文显示
                chk(chk_密文显示).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            Case 14 '收费分币处理
                cmb(cmb_分币处理).ListIndex = IIF(IsNull(rsTemp("参数值")), 0, rsTemp("参数值"))
            Case 17 '病人输入方式，分别为姓名、就诊卡、挂号单、病人ID
                strTemp = IIF(IsNull(rsTemp("参数值")), "1111", rsTemp("参数值"))
                chk(chk_病人姓名).Value = Val(Mid(strTemp, 1, 1))
                chk(chk_刷就诊卡).Value = Val(Mid(strTemp, 2, 1))
                chk(chk_病人ID).Value = Val(Mid(strTemp, 4, 1))
            Case 18 '指定药房时限制库存
                chk(chk_限定药品的库存).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            Case 45 '收费同时发药
                chk(chk_收费同时发药).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            Case 54 '时价药品以加价率入库
                chk(chk_时价药品入库).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            Case 20 '表示各种票据的号码长度，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
                strTemp = IIF(IsNull(rsTemp("参数值")), "77777", rsTemp("参数值"))
                lvw(lvw_票据).ListItems("C1").SubItems(1) = IIF(CLng(Mid(strTemp, 1, 1)) = 0, 10, CLng(Mid(strTemp, 1, 1)))
                lvw(lvw_票据).ListItems("C5").SubItems(1) = IIF(CLng(Mid(strTemp, 5, 1)) = 0, 10, CLng(Mid(strTemp, 5, 1)))
            Case 22 '日报时间限制
'                chk(chk_日报时间).Value = IIf(rsTemp("参数值") <> 0, 1, 0)
            Case 24 '表示是否严格控制管理对票据的使用，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
                strTemp = IIF(IsNull(rsTemp("参数值")), "11111", rsTemp("参数值"))
                lvw(lvw_票据).ListItems("C1").SubItems(2) = IIF(Mid(strTemp, 1, 1) = "1", "√", "")
                lvw(lvw_票据).ListItems("C5").SubItems(2) = IIF(Mid(strTemp, 5, 1) = "1", "√", "")
            Case 29 '指导批发价定价单位
                cmb(cmb_定价单位).ListIndex = IIF(rsTemp("参数值") = "1", 1, 0)
            Case 41 '医保病人适用费用类型
                SetListByText lst(lst_医保病人), Replace(IIF(IsNull(rsTemp("参数值")), "", rsTemp("参数值")), "|", ",")
            Case 42 '公费病人适用费用类型
                SetListByText lst(lst_公费病人), Replace(IIF(IsNull(rsTemp("参数值")), "", rsTemp("参数值")), "|", ",")
        End Select
        rsTemp.MoveNext
    Loop
    '显示当前票据的情况
    lvw(lvw_票据).ListItems("C1").Selected = True
    lvw_ItemClick lvw_票据, lvw(lvw_票据).SelectedItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load单据操作()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str单据 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.人员ID,B.姓名,A.单据,A.时间限制,A.他人单据,A.金额上限 from 单据操作控制 A,人员表 B where A.人员ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    lvw(lvw_单据).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_单据).ListItems.Add(, , rsTemp("姓名"), , "Limit")
        
        str单据 = Switch(rsTemp("单据") = 2, "收费单", rsTemp("单据") = 8, "会员卡")
        lst.SubItems(1) = str单据
        lst.SubItems(2) = rsTemp("时间限制")
        lst.SubItems(3) = IIF(rsTemp("他人单据") = 1, "是", "否")
        lst.SubItems(4) = IIF(IsNull(rsTemp("金额上限")), "", Format(rsTemp("他人单据"), "0.00"))
        lst.Tag = rsTemp("人员ID")
        lst.ListSubItems(1).Tag = rsTemp("单据")
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load费用类型()
'功能：初始化费用类型
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select 编码,名称 From 费用类型 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    lst(lst_医保病人).Clear
    lst(lst_公费病人).Clear
    Do Until rsTemp.EOF
        lst(lst_医保病人).AddItem rsTemp("编码") & "." & rsTemp("名称")
        lst(lst_公费病人).AddItem rsTemp("编码") & "." & rsTemp("名称")
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load药品流向()
'功能:装入药品流向数据
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_药品流向)
        '首向装入库房
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('中药库','西药库','成药库','制剂室','中药房','西药房','成药房') " & _
                   " and  b.部门ID=a.ID and " & Where撤档时间("A") & " order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("编码") & "-" & rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("ID")
            
            rsTemp.MoveNext
        Loop
        
        '装入流向控制数据
        gstrSQL = "select A.所在库房ID,A.对方库房ID,A.流向" & _
                ",B.编码 as 所在编码,B.名称 as 所在名称,C.编码 as 对方编码,C.名称 as 对方名称 " & _
                " from 药品流向控制 A,部门表 B,部门表 C " & _
                " where A.所在库房ID= B.ID and A.对方库房ID=C.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("所在库房ID")
            .TextMatrix(lngRow, 0) = rsTemp("所在编码") & "-" & rsTemp("所在名称")
            .TextMatrix(lngRow, 1) = rsTemp("对方编码") & "-" & rsTemp("对方名称")
            .TextMatrix(lngRow, 2) = rsTemp("对方库房ID")
            .TextMatrix(lngRow, 3) = Switch(rsTemp("流向") = 1, "1-所在库房可流向对方库房", _
                                            rsTemp("流向") = 2, "2-对方库房可流向所在库房", _
                                                          True, "3-两库房间可双向流通")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load药品库房单位()
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long, lngTemp As Long, lng单位 As Long, i As Long, lngMaxRow As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean
    
    '输出库房单位
    On Error GoTo ErrHandle
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.编码,'') 编码,nvl(b.名称,'') 名称,a.服务对象,a.工作性质" & vbCrLf & _
            "          FROM 部门性质说明 A, 部门表 B" & vbCrLf & _
            " WHERE B.ID=A.部门ID AND A.工作性质 IN ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房')  "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        msf库房单位.Rows = 2
        lngTemp = 0
        lngMaxRow = rsTmp.RecordCount
        For lngRow = 1 To lngMaxRow
            rsTmp.Filter = "id=" & rsTmp!ID
            strobjTemp = "": strWorkTemp = ""
            blnHave = False
            For i = 0 To msf库房单位.Rows - 1
                If msf库房单位.RowData(i) = rsTmp!ID Then
                    blnHave = True
                    Exit For
                End If
            Next
            If blnHave = False Then
                For i = 1 To rsTmp.RecordCount
                    strobjTemp = strobjTemp & rsTmp!服务对象
                    strWorkTemp = strWorkTemp & rsTmp!工作性质
                    rsTmp.MoveNext
                Next
                '1-售;2-门;3-住;4-库
                If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
                    '住院单位
                    lng单位 = 1
                ElseIf InStr(strobjTemp, "1") <> 0 Then
                    '门诊单位
                    lng单位 = 1
                ElseIf InStr(strWorkTemp, "药库") <> 0 Then
                    '药库单位
                    lng单位 = 2
                Else
                    '售价单位：主要是制剂室
                    lng单位 = 1
                End If
                If lngTemp > 0 Then
                    msf库房单位.AddItem ""
                End If
                rsTmp.MoveFirst
                msf库房单位.TextMatrix(msf库房单位.Rows - 1, 0) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                msf库房单位.TextMatrix(msf库房单位.Rows - 1, 1) = ""
                msf库房单位.TextMatrix(msf库房单位.Rows - 1, 2) = ""
                msf库房单位.TextMatrix(msf库房单位.Rows - 1, lng单位) = "√"
                msf库房单位.RowData(msf库房单位.Rows - 1) = rsTmp!ID
                lngTemp = lngTemp + 1
            End If
            rsTmp.Filter = ""
            rsTmp.MoveFirst
            rsTmp.Move lngRow, adBookmarkFirst
        Next
        gstrSQL = "select * from 药品库房单位"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            lngMaxRow = rsTmp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To msf库房单位.Rows - 1
                    If rsTmp!库房id = msf库房单位.RowData(i) Then
                        msf库房单位.TextMatrix(i, 1) = ""
                        msf库房单位.TextMatrix(i, 2) = ""
                        msf库房单位.TextMatrix(i, IIF(rsTmp!性质 = 1, 1, 2)) = "√"
                        Exit For
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
    Else
        msf库房单位.Rows = 2
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load会员卡()
'功能：得到会员卡价格
    Dim rsTemp As New ADODB.Recordset
    
    mlng会员卡ID = 0
    mlng价目ID = 0
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select ID,编码,是否变价 From 收费细目 where 末级=1 and 类别='Z' and 名称='会员卡'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    If rsTemp.RecordCount > 0 Then
        mlng会员卡ID = rsTemp("ID")
        mstr会员卡编码 = rsTemp("编码")
        chk变价.Value = IIF(rsTemp("是否变价") = 1, 1, 0)
        
        '获得价格信息
        rsTemp.Close
        
        gstrSQL = "Select A.ID,A.收入项目ID,A.现价,B.名称 From 收费价目 A,收入项目 B " & _
                  "where A.收入项目ID=B.ID and A.终止日期=to_date('3000-01-01','yyyy-MM-dd') and A.收费细目ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng会员卡ID)
                
        If rsTemp.RecordCount > 0 Then
            mlng价目ID = rsTemp("ID")
            txt价格.Text = Format(rsTemp("现价"), "###########0.000;-##########0.000;0.000;0.000")
            txt收入项目.Text = rsTemp("名称")
            txt收入项目.Tag = rsTemp("收入项目ID")
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存列宽
    SaveFlexState Bill(bill_药品流向), App.ProductName & "\" & Me.Name & bill_药品流向
    
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If Save数据() = False Then Exit Sub
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    
    If IsNumeric(txt价格.Text) = False Then
        MsgBox "请设置正确的会员卡价格。", vbInformation, gstrSysName
        Call ShowTab(1)
        txt价格.SetFocus
        Exit Function
    End If
    
    If Val(txt价格.Text) < 0 Or Val(txt价格.Text) > 10000 Then
        MsgBox "会员卡价格不合理。", vbInformation, gstrSysName
        Call ShowTab(1)
        txt价格.SetFocus
        Exit Function
    End If
    
    If txt收入项目.Tag = "" Then
        MsgBox "请为会员卡选择收入项目。", vbInformation, gstrSysName
        Call ShowTab(1)
        txt收入项目.SetFocus
        Exit Function
    End If
    
    With Bill(bill_药品流向)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(5)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(5)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(5)
                    Exit Function
                End If
            Next
        Next
    End With
    
    IsValid = True
End Function

Private Function Save数据() As Boolean
    On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    Call SavePara
    Call Save药品流向
    Call Save库房单位
    
    If Save会员卡 = False Then
        '由于该过程的SQL语句比较多，所以单独的错误处理
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    '保存完毕，事务提交
    gcnOracle.CommitTrans
    Save数据 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Save库房单位()
    '保存库房单位设置
    Dim i As Long
    Dim lngTmp As Long
    
    On Error GoTo ErrHandle
    If msf库房单位.Rows > 1 And Trim(msf库房单位.TextMatrix(1, 0)) <> "" Then
        gstrSQL = ""
        For i = 1 To msf库房单位.Rows - 1
            gstrSQL = gstrSQL & msf库房单位.RowData(i) & ","
            lngTmp = 1
            Select Case True
                Case msf库房单位.TextMatrix(i, 1) = "√"
                    lngTmp = 1
                Case msf库房单位.TextMatrix(i, 2) = "√"
                    lngTmp = 4
            End Select
            gstrSQL = gstrSQL & lngTmp & ","
        Next
        gstrSQL = "ZL_药品库房单位_INSERT('" & gstrSQL & "')"
        Call gcnOracle.Execute(gstrSQL)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SavePara()
    Dim strTemp As String
    Dim lngTemp As Long
    
    '逐个对参数进行保存
    On Error GoTo ErrHandle
    strTemp = "1," & Format(dtp(dtp_上午上班).Value, "HH:mm") & " AND " & Format(dtp(dtp_上午下班).Value, "HH:mm") & ","
    strTemp = strTemp & "2," & Format(dtp(dtp_下午上班).Value, "HH:mm") & " AND " & Format(dtp(dtp_下午下班).Value, "HH:mm") & ","
    strTemp = strTemp & "4," & ud(ud_收费收据).Value & ","
    strTemp = strTemp & "8," & chk(chk_未配药处方发药).Value & ","
    If opt(opt_不进行库存检查).Value = True Then
        lngTemp = 0
    ElseIf opt(opt_不足提醒).Value = True Then
        lngTemp = 1
    Else
        lngTemp = 2
    End If
    strTemp = strTemp & "9," & lngTemp & ","
    strTemp = strTemp & "12," & chk(chk_密文显示).Value & ","
    strTemp = strTemp & "14," & cmb(cmb_分币处理).ListIndex & ","
    strTemp = strTemp & "17," & chk(chk_病人姓名).Value & chk(chk_刷就诊卡).Value & "0" & chk(chk_病人ID).Value & ","
    strTemp = strTemp & "18," & chk(chk_限定药品的库存).Value & ","
    strTemp = strTemp & "45," & chk(chk_收费同时发药).Value & ","
    strTemp = strTemp & "54," & chk(chk_时价药品入库).Value & ","
    strTemp = strTemp & "20,"
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C1").SubItems(1) = 10, "0", lvw(lvw_票据).ListItems("C1").SubItems(1))
    strTemp = strTemp & "777" '中间的三种票据没使用
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C5").SubItems(1) = 10, "0", lvw(lvw_票据).ListItems("C5").SubItems(1)) & ","
    strTemp = strTemp & "24,"
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C1").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & "000" '中间的三种票据没使用
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C5").SubItems(2) = "√", "1", "0") & ","
    strTemp = strTemp & "29," & cmb(cmb_定价单位).ListIndex & ","
    '注意返回值是以,分隔，且外面有引号。保存时要转变一下
    strTemp = strTemp & "41," & Replace(Replace(GetTextFromList(lst(lst_医保病人)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "42," & Replace(Replace(GetTextFromList(lst(lst_公费病人)), "'", ""), ",", "|") & ","
    
    gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save单据操作()
    Dim lst As ListItem
    
    '首先删除以前的所有单据操作
    On Error GoTo ErrHandle
    gstrSQL = "zl_单据操作控制_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '再增加新的
    For Each lst In lvw(lvw_单据).ListItems
        gstrSQL = "zl_单据操作控制_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                    "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "是", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
    
Private Sub Save药品流向()
    Dim strTemp As String
    Dim lngRow As Long
    Dim str流向 As String
    
    On Error GoTo ErrHandle
    With Bill(bill_药品流向)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                str流向 = Left(.TextMatrix(lngRow, 3), 1)
                If str流向 = "" Then str流向 = "3"
                
                strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str流向 & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_药品流向控制_Modify('" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Save会员卡() As Boolean
'功能：保存对会员卡的设置
    Dim lng细目ID As Long
    Dim lng价目ID As Long
    Dim str编码 As String
    Dim oldlng上级 As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mlng会员卡ID = 0 Then
        '新增收费细目
        lng细目ID = zlDatabase.GetNextId("收费细目")
        str编码 = GetMaxLocalCode("", "收费细目", " and 类别='Z' ")
        
        gstrSQL = "zl_收费细目_insert(" & lng细目ID & ",'Z','" & str编码 & "','','','会员卡','HYK',1" & _
            ",'','','张','',0," & chk变价.Value & ",0,null,null,0,'')"
    Else
        '修改收费细目
        lng细目ID = mlng会员卡ID
        gstrSQL = "select * from 收费细目 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng细目ID)
        
        If rsTmp.RecordCount < 1 Then
            MsgBox "会员卡项目不存在！", vbInformation, gstrSysName
            Exit Function
        End If
        oldlng上级 = zlCommFun.Nvl(rsTmp!上级id, 0)
        gstrSQL = "zl_收费细目_update(" & lng细目ID & ",'" & mstr会员卡编码 & "','','','会员卡','HYK'" & _
            IIF(oldlng上级 = 0, ", Null", "," & oldlng上级) & ",'','','张','',0," & chk变价.Value & ",0,null,0,'')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If mlng价目ID = 0 Then
        '新增价格
        lng价目ID = zlDatabase.GetNextId("收费价目")
        gstrSQL = "zl_收费价目_insert(" & _
           lng价目ID & ",null," & lng细目ID & "," & txt收入项目.Tag & ",0," & txt价格.Text & _
           ",0,0,''," & lng价目ID & ",'" & gstrUserName & "',sysdate)"
    Else
        '修改价格
        gstrSQL = "zl_收费价目_update(" & lng细目ID & "," & txt收入项目.Tag & ",0," & txt价格.Text & _
             ",0,0,''," & mlng价目ID & ",'" & gstrUserName & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '设置特定收费项目
    gstrSQL = "zl_收费特定项目_Modify('就诊卡," & lng细目ID & ",')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save会员卡 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    If Index = chk_票号控制 Then
        lvw(lvw_票据).SelectedItem.SubItems(2) = IIF(chk(Index).Value = 1, "√", "")
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    mblnChange = True
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtp_Change(Index As Integer)
    Dim intNext As Integer

    mblnChange = True
    If Index < dtp_下午下班 Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).Value
        If dtp(intNext).Value < dtp(intNext).MinDate Then
            dtp(intNext).Value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     If Index = lvw_单据 Then
        If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
            lvw(lvw_单据).SortOrder = IIF(lvw(lvw_单据).SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            mintColumn = ColumnHeader.Index - 1
            lvw(lvw_单据).SortKey = mintColumn
            lvw(lvw_单据).SortOrder = lvwAscending
        End If
     End If
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Index = lvw_单据 Then
        Call cmdOperate_Click(1)
    End If
End Sub

Private Sub lvw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    mblnChange = True
    
    Dim itemTemp As MSComctlLib.ListItem
    For Each itemTemp In lvw(Index).ListItems
        If Not itemTemp Is Item Then
            itemTemp.Checked = False
        End If
    Next
End Sub

Private Sub lvw_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim lng原值 As Long
    
    If Index = lvw_票据 Then
        lng原值 = Val(Item.SubItems(1))
        
        If Item.Text = "就诊卡" Then
            ud(ud_号码长度).Max = 8
        Else
            ud(ud_号码长度).Max = 10
        End If
        '设置最大值时，可能已经更改了列表中的值
        ud(ud_号码长度).Value = lng原值
        chk(chk_票号控制).Value = IIF(Item.SubItems(2) = "√", 1, 0)
    End If
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_票据 Then
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    Else
        If KeyAscii = vbKeyReturn Then cmdOperate_Click (1)
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt价格_GotFocus()
    zlControl.TxtSelAll txt价格
End Sub

Private Sub txt价格_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt收入项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    Else
        If KeyAscii = Asc("*") Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub chk变价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub ud_Change(Index As Integer)
    mblnChange = True
    '动态改变票号长度
    If Index = ud_号码长度 Then
        lvw(lvw_票据).SelectedItem.SubItems(1) = ud(ud_号码长度).Value
    End If
End Sub

Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If Index <> bill_药品流向 Then Exit Sub
    
    With Bill(bill_药品流向)
        If ListIndex < 0 Then Exit Sub
        If .Col = 0 Then
            .RowData(.Row) = .ItemData(ListIndex)
        Else
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
        End If
        .TextMatrix(.Row, .Col) = .CboText
        
        If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
    End With
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If Index = bill_药品流向 And .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                .RowData(.Row) = .ItemData(.ListIndex)
            End If
            
            If Index = bill_药品流向 Then
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
            End If
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'处理最后一列的变化
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_药品流向 Then
        If .MouseCol <> .Cols - 1 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "1"
                .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
            Case "2"
                .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
            Case Else
                .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
        End Select
    End If
    mblnChange = True
End With
    
End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
With Bill(Index)
    If Index = bill_药品流向 Then
        If .Col = 3 Then
            Select Case KeyAscii
                Case Asc(" ")
                    '切换计算标志
                    Select Case Left(.TextMatrix(.Row, .Col), 1)
                        Case "1"
                            .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                        Case "2"
                            .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                        Case Else
                            .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                    End Select
                    mblnChange = True
                Case vbKey1
                    .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                    mblnChange = True
                Case vbKey2
                    .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                    mblnChange = True
                Case vbKey3
                    .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                    mblnChange = True
            End Select
        End If
    End If
End With

End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    
    For i = fraMain.LBound To fraMain.UBound
        fraMain(i).Visible = False
    Next
    
    i = tabMain.SelectedItem.Index - 1
    fraMain(i).Visible = True
    Select Case tabMain.SelectedItem.Index
        Case 1 '常规
            cmb(cmb_定价单位).SetFocus
        Case 2 '票据管理
            txtUD(ud_收费收据).SetFocus
        Case 3 '权限
            lst(lst_医保病人).SetFocus
        Case 4 '单据
            lvw(lvw_单据).SetFocus
        Case 5 '药品流向
            Bill(bill_药品流向).SetFocus
        Case 6  '药品库房单位
            msf库房单位.SetFocus
    End Select
End Sub

Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub

Private Function NumIsValid(ByVal strNumber As String) As Boolean
'功能:分析输入内容是否是一个有效的数字
'参数:strNumber  输入内容
'返回值:有效返回True,否则为False
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "请输入一个数值。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "这个数太大了。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) < 0 Then
        MsgBox "不能为负数。", vbInformation, gstrSysName
        Exit Function
    End If
    NumIsValid = True
End Function


